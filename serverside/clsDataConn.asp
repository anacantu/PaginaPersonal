<%
' ------------------------------------------------------------------------------
'	Author:		Carlos Trevino
'	Email:		hobbes313@gmail.com
'	Date:		May 22, 2009
' ------------------------------------------------------------------------------
'
'	NOTE: /serverside/function.asp must be loaded before
'
' Demo:
'	set DataConn = New clsDataConn
'	call DataConn.ExecuteCommand(ProcName, ParamArray)
'
'	DataConn.oCNDB
'		DataBase Connection Object (ADODB.Connection)
'
'
'	ADODB.Connection Object state values
'	0	The object is closed
'	1	The object is open
'	2	The object is connecting
'	4	The object is executing a command
'	8	The rows of the object are being retrieved
'
'	ADODB.Connection Object mode values
'
'	0	Permissions have not been set or cannot be determined.
'	1	Read-only.
'	2	Write-only.
'	3	Read/write.
'	4	Prevents others from opening a connection with read permissions.
'	8	Prevents others from opening a connection with write permissions.
'	12	Prevents others from opening a connection.
'	16	Allows others to open a connection with any permissions.
'

Class clsDataConn
' ------------------------------------------------------------------------------
	
	Private p_connString
	Private p_oCNDB
	Private p_RS
	Private p_Command
	
	'-- error flags
	Private p_CommandError
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		'-- get DB connection string
		p_connString = Application("connString")
		
		'-- initialize and open DB connection object
		set p_oCNDB		= Server.CreateObject("ADODB.Connection")
			
		'-- initialize DB Command object
		set p_Command	= Server.CreateObject("ADODB.Command")
		
		'-- initialize DB RecordSet Object
		set p_RS		= Server.CreateObject("ADODB.RecordSet")
		
		'-- error flags
		p_CommandError = false
		
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub Class_Terminate()
		'-- close and destroy DB connection
		call CloseObject(p_oCNDB)
		call CloseObject(p_RS)
		call CloseObject(p_Command)
		
		set p_oCNDB		= Nothing
		set p_RS 		= Nothing
		set p_Command	= Nothing
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub OpenReadOnlyConn()
		
		'-- open DB connection
		if p_oCNDB.State = 0 Then
			p_oCNDB.Mode	= 1 ' Read only
			p_oCNDB.Open(p_connString)
		end if
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub OpenReadWriteConn()
		
		'-- open DB connection
		if p_oCNDB.State = 0 Then
			p_oCNDB.Mode	= 3 ' Read and write
			p_oCNDB.Open(p_connString)
		end if
		
	End Sub
	
	Private Sub SaveSQLLog(byVal sqlString)
		
		Dim scriptName, remoteHost, userAgent, idUser, userType, urlStr
		
		idUser     = ConvertToSQLStringValue(session("SYS_id_user"))
		userType   = ConvertToSQLStringValue(session("SYS_user_type"))
		scriptName = ConvertToSQLStringValue(Request.ServerVariables("SCRIPT_NAME"))
		urlStr	   = ConvertToSQLStringValue(Request.ServerVariables("URL"))
		remoteHost = ConvertToSQLStringValue(Request.ServerVariables("REMOTE_HOST") & " ("& Request.ServerVariables("REMOTE_ADDR") &")")
		userAgent  = ConvertToSQLStringValue(Request.ServerVariables("HTTP_USER_AGENT")) 
		
		p_oCNDB.EXECUTE "EXEC SaveSQLLog @command = '"& sqlString &"', @id_user = "& idUser &", @script_name = "& scriptName &", @url = "& urlStr &", @remote_host = "& remoteHost &", @user_agent = "& userAgent
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub CloseObject(ADOObject)	
		if ADOObject.State > 0 Then ADOObject.Close()
	End Sub
	
' ------------------------------------------------------------------------------
' - PUBLIC PROPERTIES
' ------------------------------------------------------------------------------
	Public Property Get R_DBConn()
		call OpenReadOnlyConn()
		R_DBConn	 = p_oCNDB
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ResultSet()
		ResultSet	= p_RS
	End Property
' ------------------------------------------------------------------------------	
	Public Property Get CommandError()
		CommandError= p_CommandError
	End Property
	
' ------------------------------------------------------------------------------
' - PUBLIC METHODS
' ------------------------------------------------------------------------------
	Public Function ResultSetValue(columnKey)
		ResultSetValue = ""
		if p_RS.State > 0 Then
			if NOT p_RS.EOF Then ResultSetValue = p_RS(columnKey)
		end if
	End Function
' ------------------------------------------------------------------------------	
	Public Sub ExecuteCommand(ByVal SPname, ByVal parametersArr)
		
		CloseObject(p_RS)
		OpenReadWriteConn()
		
		'error handling
		On Error Resume Next
		
		Dim parameterStr
		parameterStr = ""
		
		p_Command.ActiveConnection	= p_oCNDB
		p_Command.CommandText		= SPname
		p_Command.CommandType		= 4 '--Store.Procedure Call.Type
		
		'--parametros
		if UBound(parametersArr, 1) >= 1 Then
			for arrayIndex = 0 to UBound(parametersArr, 2)
				p_Command.parameters(parametersArr(0, arrayIndex)) = parametersArr(1, arrayIndex)
				
				'-- concatenate parameter name and value
				parameterStr = parameterStr & parametersArr(0, arrayIndex) & "= " & SafeCstr(parametersArr(1, arrayIndex)) & ", "
				
			next
		end if
		
		set p_RS = p_Command.Execute
		
		'-- close command object
		'CloseObject(p_Command)
		
		'-- save query log
		parameterStr = RemoveLastSeparator(parameterStr, ",")
		SaveSQLLog("exec " & SPname & " " & parameterStr)
		
		p_CommandError = Err.Number <> 0
		
		If p_CommandError then
			SendCrashMail "Error al ejecutar comando SQL (clsDtaConn.ExecuteCommand)", "Error al ejecutar comando SQL:<br><strong>"&_
			Err.Description & "</strong><br><br>" & SPname & "<br>" &_
			parameterStr
			
			Err.Clear
		end if
		
		'end error handling
		On Error GoTo 0
		
	end sub
' ------------------------------------------------------------------------------
	'-- close an ADODB recordset
	Public Sub CloseResultSet()
		call CloseObject(p_RS)
	end sub
' ------------------------------------------------------------------------------
	'-- close an ADODB command
	Public Sub CloseCommand()
		call CloseObject(p_Command)
	end sub
' ------------------------------------------------------------------------------
	'-- close an ADODB connection
	Public Sub CloseConn()
		call CloseObject(p_oCNDB)
	end sub
' ------------------------------------------------------------------------------
End Class
' ------------------------------------------------------------------------------
%>