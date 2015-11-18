<%
'-- NOTA: SE REQUIERE HABER CARGADO functions.asp

Dim DALC_DataAccess, DALC_sqlString, DALC_value


'-- Guardar registro de los comandos SQL ejecutados
sub SQLExecute(sqlString)
	
	'error handling
	On Error Resume Next
	
	oCNDB.Execute sqlString
	
	If Err.Number <> 0 then
		SendCrashMail "Error al ejecutar comando SQL", "Error al ejecutar comando SQL:<br><strong>"&_
		Err.Description & "</strong><br><br>" & sqlString
	end if
	
	call SaveSQLLog(sqlString) 
		
	'end error handling
	On Error GoTo 0
	
end sub


'-- execute a SQL command with parameters and return recordset with SQL resultset
function SQLExecuteCommand(ByVal sqlString, ByVal parametersArr)
	
	'error handling
	On Error Resume Next
	
	Dim SQLEC_cmd, SQLEC_paramStr
	SQLEC_paramStr = ""
	
	Set SQLEC_cmd = server.CreateObject("ADODB.Command")
	SQLEC_cmd.ActiveConnection	= oCNDB
	SQLEC_cmd.CommandText		= sqlString
	SQLEC_cmd.CommandType		= 4
	
	'--parametros
	if UBound(parametersArr, 1) >= 1 Then
		for arrayIndex = 0 to UBound(parametersArr, 2)
			'temporal, quien invoca debe enviar NULL, no "NULL", y debemos verificar que funcione
			if parametersArr(0, arrayIndex) <> "" AND parametersArr(1, arrayIndex) <> "NULL" AND NOT IsNull(parametersArr(1, arrayIndex)) Then
				SQLEC_cmd.parameters(parametersArr(0, arrayIndex)) = parametersArr(1, arrayIndex)
				
				SQLEC_paramStr = SQLEC_paramStr & parametersArr(0, arrayIndex) & " = " & parametersArr(1, arrayIndex) & ", "
			end if
		next
	end if
	
	set SQLExecuteCommand = SQLEC_cmd.Execute
	
	SQLEC_paramStr = RemoveLastSeparator(SQLEC_paramStr, ", ")
	
	If Err.Number <> 0 then
		SendCrashMail "Error al ejecutar comando SQL con parametros", "Error al ejecutar comando SQL con parametros:<br><strong>"&_
		Err.Description & "</strong><br><br>" & sqlString & " " & SQLEC_paramStr
	end if
	
	call SaveSQLLog(sqlString & " " & SQLEC_paramStr) 
	
	'end error handling
	On Error GoTo 0
	
end function


'-- save SQL query into system log in database
sub SaveSQLLog(byVal sqlString)
	Dim scriptName, remoteHost, userAgent, idUser, userType, urlStr
	
	idUser     = ConvertToSQLStringValue(session("SYS_id_user"))
	userType   = ConvertToSQLStringValue(session("SYS_user_type"))
	scriptName = ConvertToSQLStringValue(Request.ServerVariables("SCRIPT_NAME"))
	urlStr	   = ConvertToSQLStringValue(Request.ServerVariables("URL"))
	remoteHost = ConvertToSQLStringValue( Request.ServerVariables("REMOTE_HOST") & " ("& Request.ServerVariables("REMOTE_ADDR") &")" )
	userAgent  = ConvertToSQLStringValue(Request.ServerVariables("HTTP_USER_AGENT")) 
	
	oCNDB.EXECUTE "EXEC SaveSQLLog @command = "& ConvertToSQLStringValue(sqlString) &", @id_user = "& idUser &", @script_name = "& scriptName &", @url = "& urlStr &", @remote_host = "& remoteHost &", @user_agent = "& userAgent
	
end sub


'------------------------------------
'-- Funciones para validar y adecuar valores a ser enviados como parameros a comandos SQL
'------------------------------------
function ConvertToSQLNumericParam(valueToConvert, defaultValue)
	if NOT isNumeric(valueToConvert) OR valueToConvert = "" Then
		ConvertToSQLNumericParam = defaultValue
	else
		ConvertToSQLNumericParam = valueToConvert
	end if
end function


function ConvertToSQLStringParam(textToConvert)
	Dim CSQLSP_cleanText
	CSQLSP_cleanText = removeInvalidSQLChars(textToConvert)
	
	if CSQLSP_cleanText <> "" then
		ConvertToSQLStringParam = CSQLSP_cleanText
	else
		ConvertToSQLStringParam = Null
	end if
end function


function ConvertToSQLDateParam(dateParam)
	if not IsNull(dateParam) and dateParam <> "" and IsDate(dateParam) then
		if hour(dateParam) > 0 AND minute(dateParam) > 0 Then
			ConvertToSQLDateParam = year(dateParam) & "-" & month(dateParam) & "-" & day(dateParam) & " " & hour(dateParam) & ":" & minute(dateParam)
		else
			ConvertToSQLDateParam = year(dateParam) & "-" & month(dateParam) & "-" & day(dateParam)
		end if
	else
		ConvertToSQLDateParam = null
	end if
end function

'------------------------------------
'-- END Funciones para validar y adecuar valores a ser enviados como parameros a comandos SQL
'------------------------------------

function CountDBRecords(objRecordset)
	Dim recordCount
	recordCount = 0
	
	While NOT objRecordset.EOF
		recordCount = recordCount + 1
		objRecordset.MoveNext
	Wend
	
	if NOT objRecordset.BOF Then objRecordset.movefirst
	
	CountDBRecords = recordCount
	
end function



'-- close an ADODB recordset
sub SafeCloseRecordSet(paramRS)
	if NOT IsNull(paramRS) Then
		if paramRS.State <> 0 Then
			paramRS.close()
		end if
	end if
end sub


function ValueExists(tableName, columnName, keyValue, exludeDeleted)
	set DALC_listObject = Server.CreateObject("ADODB.Recordset")
	ssqlTemp = "SELECT "&columnName&" FROM "&tableName& " WHERE "&columnName& " = " & DALC_ConvertToSQLStringValue(keyValue)
	if exludeDeleted then
		ssqlTemp = ssqlTemp & " AND logicaldeletion IS NULL "
	end if
	
	DALC_listObject.Open ssqlTemp, oCNDB, 3, 3
	ValueExists = (NOT DALC_listObject.EOF)
	
	DALC_listObject.close
end function


function GetCatalogList(acceso_db, ssqlTemp, idKey, nameKey, idSelected, defaultValueLabel, defaultValue)
	Set DALC_listObject = Server.CreateObject("ADODB.Recordset")
	DALC_listObject.Open ssqlTemp, acceso_db
	
	Dim GCL_returnValue
	if IsNull(defaultValue) Then
		GCL_returnValue = ""
	else
		GCL_returnValue = "<option value="""&defaultValue&""" selected>"&defaultValueLabel&"</option>"
	end if
	
	WHILE NOT DALC_listObject.EOF
		selectedStr = ""
		
		if NOT IsNull(idSelected) Then
			if Cstr(idSelected) = Cstr(DALC_listObject(idKey)) Then
				selectedStr = "selected"
			end if
		end if
		
		GCL_returnValue = GCL_returnValue & "<option value="""&DALC_listObject(idKey)&""" "&_
		selectedStr &">"&DALC_listObject(nameKey)&"</option>"
		
		DALC_listObject.movenext
	WEND
	
	DALC_listObject.close
	set DALC_listObject = nothing
	GetCatalogList = GCL_returnValue
end function

sub PrintGeneralCatalog(acceso_db, tableName, getSeleted)
	Set DALC_listObject = Server.CreateObject("ADODB.Recordset")
	ssqlTemp = "SELECT * FROM "&tableName & " ORDER BY " & tableName
	DALC_listObject.Open ssqlTemp, acceso_db
	
	WHILE NOT DALC_listObject.EOF
		response.Write("<option value="""&DALC_listObject("id_"&tableName)&""">"&DALC_listObject(tableName)&"</option>")
		DALC_listObject.movenext
	WEND
	
	DALC_listObject.close
	set DALC_listObject = nothing
end sub


function PrintGeneralCatalog4(acceso_db, ssqlTemp, idKey, nameKey, idSelected, idMore, textMore)
	Dim catalogString, catalogCounter
	catalogString = ""
	catalogCounter = 0
	
	Set DALC_listObject = Server.CreateObject("ADODB.Recordset")
	DALC_listObject.Open ssqlTemp, acceso_db
	
	if NOT DALC_listObject.EOF AND textMore <> "" Then
		catalogString = catalogString & "<option value="""& idMore &""">"& textMore &"</option>"
	end if
	
	WHILE NOT DALC_listObject.EOF
		selectedStr = ""
		
		if NOT IsNull(idSelected) Then
			if Cstr(idSelected) = Cstr(DALC_listObject(idKey)) Then
				selectedStr = "selected"
			end if
		end if
		
		catalogString = catalogString & "<option value="""&DALC_listObject(idKey)&""" "&_
		selectedStr &">"&DALC_listObject(nameKey)&"</option>"
		
		DALC_listObject.movenext
		catalogCounter = catalogCounter + 1
	WEND
	
	if catalogCounter > 10 Then
		catalogString = catalogString & "<option value="""& idMore &""">"& textMore &"</option>"
	end if
	
	DALC_listObject.close
	set DALC_listObject = nothing
	
	PrintGeneralCatalog4 = catalogString
end function


sub PrintGeneralCatalog3(acceso_db, ssqlTemp, idKey, nameKey, idSelected)
	Set DALC_listObject = Server.CreateObject("ADODB.Recordset")
	DALC_listObject.Open ssqlTemp, acceso_db
	
	WHILE NOT DALC_listObject.EOF
		selectedStr = ""
		
		if NOT IsNull(idSelected) Then
			if Cstr(idSelected) = Cstr(DALC_listObject(idKey)) Then
				selectedStr = "selected"
			end if
		end if
		
		response.Write("<option value="""&DALC_listObject(idKey)&""" "&_
		selectedStr &">"&DALC_listObject(nameKey)&"</option>")
		
		DALC_listObject.movenext
	WEND
	
	DALC_listObject.close
	set DALC_listObject = nothing
end sub

sub PrintGeneralCatalog2(tableName, selectedValue)
	'Set DALC_listObject = Server.CreateObject("ADODB.Recordset")
	'ssqlTemp = "SELECT * FROM "&tableName&" ORDER BY " & tableName
	'DALC_listObject.Open ssqlTemp, oCNDB
	
	'TEMPORAL, cambiar
	set DALC_listObject = Server.CreateObject("ADODB.Recordset")
	DALC_listObject.ActiveConnection = general_mdb
	DALC_listObject.Source = "SELECT * FROM "&tableName&" ORDER BY " & tableName
	DALC_listObject.CursorType = 0
	DALC_listObject.CursorLocation = 2
	DALC_listObject.LockType = 3
	DALC_listObject.Open()
	
	Dim DALC_selectedStr
	Dim itemValue
	
	WHILE NOT DALC_listObject.EOF
		DALC_selectedStr = ""
		itemValue = Cint(DALC_listObject("id_"&tableName))
		if Cint(selectedValue) = itemValue then
			DALC_selectedStr = "selected"
		end if
		
		response.Write("<option value="""&itemValue&""" "&DALC_selectedStr&">"&DALC_listObject(tableName)&"</option>")
		DALC_listObject.movenext
	WEND
	
	DALC_listObject.close
	set DALC_listObject = nothing
	set DALC_selectedStr = nothing
end sub


sub PrintCatalog(catalogName, selectedValue, paramObj)
	Set DALC_listObject = Server.CreateObject("ADODB.Recordset")

	select case catalogName
		case "Strategy":
			response.write "<option value=""-1"""
			if selectedValue = "" Then 
				response.Write(" selected")
			end if
			response.write ">-todas-</option>"

			DALC_sqlString = "select id_strategy AS id, name from strategy where logicaldeletion IS NULL ORDER BY name"
			DALC_listObject.Open DALC_sqlString, oCNDB, 3, 3
			
			WHILE NOT DALC_listObject.EOF
				DALC_value = Cdbl(DALC_listObject("id"))

				response.Write("<option value=""" & DALC_value & """")
				if DALC_value = selectedValue Then
					response.write " selected"
				end if
				response.Write(">" & removeInvalidHTMLChars(DALC_listObject("name")) &"</option>")
				DALC_listObject.movenext()
			WEND
			
			DALC_listObject.close
			
		
		
			
	end select
	set DALC_listObject = nothing

end sub


function PrintOptionsCatalog(storeProcedure, keyValue, keyName, matchKey, matchValue, insertDefault, defaultValue, defaultName)
	DALC_sqlString = "exec " & storeProcedure
	PrintOptionsCatalog = ""
	
	if insertDefault then
		DALC_string = " selected"
		if NOT IsNull(matchKey) then
			DALC_string = ""
		end if
		
		response.write("<option value="""& defaultValue &""""& DALC_string &">"& defaultName &"</option>")
	end if
		
	set DALC_DataAccess = server.CreateObject("ADODB.Recordset")
	DALC_DataAccess.Open DALC_sqlString, oCNDB, 3, 3
	
	while NOT DALC_DataAccess.EOF
	
		Response.Write("<option value=""" & DALC_DataAccess(keyValue) & """")
		
		if matchKey <> "" Then
			if not IsNull(DALC_DataAccess(matchKey)) and not IsNull(matchValue) then
				if CStr(DALC_DataAccess(matchKey)) = CStr(matchValue) then 
					Response.Write(" selected")
					PrintOptionsCatalog = DALC_DataAccess(keyValue)
				end if
			end if
		end if
		
		Response.Write(">" & DALC_DataAccess(keyName) & "</option>")
		DALC_DataAccess.MoveNext()
	wend
	DALC_DataAccess.Close()
end function


'Get mail recordset
function GetMailRecordSet(DALC_mailID)
	Dim mailRS
	set mailRS = Server.CreateObject("ADODB.Recordset")
	SQLString = "SELECT * FROM mail WHERE id_mail = " & DALC_mailID
	mailRS.Open SQLString, oCNDB, 3, 3
	
	set GetMailRecordSet = mailRS
end function


'Obtener un valor de base de datos
function GetDBValue(tableName, columnKey, indexColumn, indexValue)
	Set DALC_DataAccess = Server.CreateObject("ADODB.Recordset")
	sqlString = "SELECT "&columnKey&" FROM "&tableName&" WHERE "&indexColumn&" = "& indexValue
	DALC_DataAccess.Open sqlString, oCNDB, 3, 3
	
	if DALC_DataAccess.EOF then
		GetDBValue = -1
	else
		if IsNumeric(DALC_DataAccess(columnKey)) then
			GetDBValue = Cdbl(DALC_DataAccess(columnKey))
		else
			'if DALC_DataAccess(columnKey) <> null then
			'	GetDBValue = Cstr(DALC_DataAccess(columnKey))
			'else
				GetDBValue = DALC_DataAccess(columnKey)
			'end if
		end if
	end if
	DALC_DataAccess.close
	
end function

'Obtener un valor de base de datos, especificando la condicion
function GetDBValue2(tableName, columnKey, whereClause)
	Set DALC_DataAccess = Server.CreateObject("ADODB.Recordset")
	sqlString = "SELECT "&columnKey&" FROM "&tableName&" WHERE "&whereClause
	DALC_DataAccess.Open sqlString, oCNDB, 3, 3
	
	if DALC_DataAccess.EOF then
		GetDBValue2 = -1
	else
		if IsNumeric(DALC_DataAccess(columnKey)) then
			GetDBValue2 = Cdbl(DALC_DataAccess(columnKey))
		else
			GetDBValue2 = DALC_DataAccess(columnKey)
		end if
	end if
	DALC_DataAccess.close
	
end function

'Obtener un valor de base de datos
function GetDBValue3(tableName, columnKey, querySelect, queryFilter)
	Set DALC_DataAccess = Server.CreateObject("ADODB.Recordset")
	sqlString = "SELECT "&querySelect&" FROM "&tableName&" WHERE "&queryFilter
	DALC_DataAccess.Open sqlString, oCNDB, 3, 3

	if DALC_DataAccess.EOF then
		GetDBValue3 = -1
	else
		if IsNumeric(DALC_DataAccess(columnKey)) then
			GetDBValue3 = Cdbl(DALC_DataAccess(columnKey))
		else
			GetDBValue3 = DALC_DataAccess(columnKey)
		end if
	end if
	DALC_DataAccess.close
	
end function

function GetDBValue4(tableName, columnKey, compareValue, columnID, itemID)
	Set DALC_DataAccess = Server.CreateObject("ADODB.Recordset")
	if itemID <> "" then
		sqlString = "SELECT "&columnKey&", "&columnID&" FROM "&tableName&" WHERE "&columnID&" <> "& itemID &" AND logicaldeletion IS NULL"
	else
		sqlString = "SELECT "&columnKey&", "&columnID&" FROM "&tableName&" WHERE logicaldeletion IS NULL"
	end if
	DALC_DataAccess.Open sqlString, oCNDB, 3, 3
	Dim readValue
	if DALC_DataAccess.EOF then
		GetDBValue4 = -1
	else
		do while Not DALC_DataAccess.EOF
			readValue = normalizeString(DALC_DataAccess(columnKey))
			if readValue = compareValue then
				GetDBValue4 = DALC_DataAccess(columnID)
				Exit do
			else
				GetDBValue4 = -1
			end if
		DALC_DataAccess.movenext
		loop
	end if
	DALC_DataAccess.close
	
	set readValue = Nothing
end function


function getInsertedID
	Set DALC_insertedITEM = Server.CreateObject("ADODB.Recordset")
	DALC_insertedITEM.Open "SELECT @@IDENTITY AS itemID", oCNDB
	
	if NOT DALC_insertedITEM.EOF then
		if IsNumeric(DALC_insertedITEM("itemID")) then
			getInsertedID = Cint(DALC_insertedITEM("itemID"))
		else
			getInsertedID = DALC_insertedITEM("itemID")
		end if
	else
		getInsertedID = -1
	end if
end function

'Obtener maximo ID de una tabla
function GetMaxID(tableName, columnKey)
	GetMaxID = -1
	
	set DALC_AccessObj = server.CreateObject("ADODB.RecordSet")
	DALC_AccessObj.open "select MAX("&columnKey&") AS "&columnKey&" from " & tableName, oCNDB
	
	if NOT DALC_AccessObj.EOF then
		GetMaxID = Cint(DALC_AccessObj(columnKey))
	end if
	
	DALC_AccessObj.close
	set DALC_AccessObj = nothing
end function

function getInsertedIDStrong(tableName, columnKey)
	'Obtener ID recien insertado
	getInsertedIDStrong = getInsertedID
	
	'Corregir posible falla en GetInsertedID
	if NOT IsNumeric(getInsertedIDStrong) Then
		getInsertedIDStrong = GetMaxID(tableName, columnKey)
	end if
end function

function getConfigurationValue (keyValue, columnKey)
	Set DALC_recordSet = Server.CreateObject("ADODB.Recordset")
	DALC_recordSet.Open "SELECT * FROM configuration WHERE keyname = '"& keyValue &"'", oCNDB
	
	if NOT DALC_recordSet.EOF then
		getConfigurationValue = DALC_recordSet(columnKey)
	else
		getConfigurationValue = ""
	end if
end function






'-- VARIABLE HANDLING
function DALC_ConvertToSQLStringValue(textToConvert)
	Dim cleanText
	cleanText = DALC_removeInvalidSQLChars(textToConvert)
	
	if cleanText <> "" then
		DALC_ConvertToSQLStringValue = "'"&cleanText&"'"
	else
		DALC_ConvertToSQLStringValue = "NULL"
	end if
end function


function DALC_removeInvalidSQLChars(textToClean)
	if textToClean <> "" then
		DALC_removeInvalidSQLChars = trim(replace(textToClean,"'","''"))
	else
		DALC_removeInvalidSQLChars = textToClean
	end if
	
end function


set DALC_DataAccess = Nothing
set DALC_sqlString = Nothing
set DALC_value = nothing
%>