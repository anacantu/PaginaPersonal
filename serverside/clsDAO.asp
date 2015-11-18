<%
' ------------------------------------------------------------------------------
'	Author:		Israel Morales @ Pulso Vital Consulting Group
'	URL:		http://www.pulsovital.com
'	Date:		Apr 09, 2009
' ------------------------------------------------------------------------------
'
'	Data Access Object : Objeto para interaccion con base de datos
'

Class clsDAO
' ------------------------------------------------------------------------------
	
	Private dbcommand
	Private dbconnection
	Private strParameters
	
	'Class constructor
	Private Sub Class_Initialize()
		
		'-- database connection object
		Set dbconnection = Server.CreateObject("ADODB.Connection")
		dbconnection.ConnectionTimeout = 600 ' 10 min
		dbconnection.CommandTimeout = 300 ' 5 min
		dbconnection.Open Application("connString")
		
		'-- database command object
		Set dbcommand = server.CreateObject("ADODB.Command")
		dbcommand.CommandType = 4
		dbcommand.ActiveConnection = dbconnection
		
		strParameters = ""
	
	End Sub
	
	'Class destructor
	Private Sub Class_Terminate()
		if dbcommand.state <> 0 Then dbcommand.close
		set dbcommand = Nothing
		
		if dbconnection.state <> 0 Then dbconnection.close
		set dbconnection = nothing
		
	End Sub
	
	
	'-- Properties to insert command's parameters before execution
	Public Property Let CommandType(ByVal typeDbCommand)
	    dbcommand.CommandType = typeDbCommand
	End Property
	
	Public Sub SetCommand(ByVal sqlString)
	    dbcommand.CommandText = sqlString
	End Sub
	
		
	Public Sub SetCommandConexion(ByVal sqlString, ByVal dbconnection )
	     dbcommand.ActiveConnection = dbconnection
         dbcommand.CommandText = sqlString
	    End Sub
	
	Public sub SetCommandText(ByVal sqlString)
	    dbcommand.CommandText = sqlString
	End sub
	
	Public Sub AddParameter(ByVal name, ByVal tipo, ByVal direction,ByVal size, ByVal value)
	    dbcommand.Parameters.Append dbcommand.CreateParameter(name, tipo, direction, size, value)
	    strParameters = strParameters & name & " = " & value & " type = " & tipo & ", "
	End Sub
	
	Public Sub AddParameterReturn(ByVal name, ByVal tipo, ByVal direction, ByVal size)
	    dbcommand.Parameters.Append dbcommand.CreateParameter(name, tipo, direction, size)
	    strParameters = strParameters & name & " out, " & " type = " & tipo 
	End Sub
	
	Public Sub ParametersClear()
		strParameters=""
		While dbcommand.Parameters.Count  
			dbcommand.Parameters.Delete  0
		Wend
	End Sub
	
	Public Function GetParameterReturn(ByVal name)
	    Set GetParameterReturn = dbcommand.Parameters(name)
	End Function
	
	Public Function GetReturnValue(ByVal name)
	    Set GetReturnValue = dbcommand(name)
	End Function
	
	'Ejecutar comando SQL
	Public Function ExecuteCommand()
		SaveSQLLog
		dbcommand.Prepared = True
		Set ExecuteCommand = dbcommand.Execute
	End Function
	
	'Save SQL command to Database Log table
	Private Sub SaveSQLLog()
		
		'start error handling
		On Error Resume Next

		'Los excel no tienen DefaultDatabase
		If dbcommand.ActiveConnection.DefaultDatabase = "" then		
		   Exit Sub
		End If
		
		Dim dbSQLLog
		Set dbSQLLog = server.CreateObject("ADODB.Command")
		dbSQLLog.ActiveConnection = dbcommand.ActiveConnection
		
		Dim scriptName, remoteHost, userAgent, idUser, urlStr,script
		
		idUser = "NULL"
		
		If Not TypeName(Session(SESSION_USER)) = "Empty" Then
			On Error Resume Next
				idUser = personItem.IDUser
				
			If Err.Number <> 0 Then
				DProfile = "NULL" 
				idUser = "NULL"
			End If
			
			On Error GoTo 0
		end if 
		
		scriptName	= ConvertToSQLStringValue(left(Request.ServerVariables("SCRIPT_NAME"), 100))
		urlStr		= ConvertToSQLStringValue(left(Request.ServerVariables("URL"), 100))
		remoteHost	= ConvertToSQLStringValue(left(Request.ServerVariables("REMOTE_HOST") & " ("& Request.ServerVariables("REMOTE_ADDR") &")", 100))
		userAgent	= ConvertToSQLStringValue(left(Request.ServerVariables("HTTP_USER_AGENT"), 200))
		
		script		=  dbcommand.CommandText
		if strParameters <> "" then script = script & " Parametros:" & strParameters
		script =  left(dbcommand.CommandText, 2000)
		
		dbSQLLog.CommandText =  "EXEC SaveSQLLog @command = '"& script &"', @id_user = "& idUser &", @script_name = "& scriptName &", @url = "& urlStr &", @remote_host = "& remoteHost &", @user_agent = "& userAgent
		dbSQLLog.Prepared = True
		dbSQLLog.Execute
		
		if dbSQLLog.state <> 0 Then dbSQLLog.close
		set dbSQLLog = nothing
		
		Err.Clear
		
		'end error handling
		On Error GoTo 0
		
	End Sub
	
	
	
	' ------------------------------------------------------------------------------
	' -- Validation of values
	
	Public Function DBString(ByVal value)
		If IsNull(value) Then	 		
			DBString = ""
			Exit Function
		End If
		DBString = trim(value)
	End Function
	
	Public Function DBStringNull(ByVal value)
		If IsNull(value) Then	 		
			DBStringNull = null
			Exit Function
		End If
		
		If  trim(value) = "" then
			DBStringNull = null
			Exit Function
		End If 
		
		DBStringNull = trim(value)
	End Function
	
	
	Public Function DBBool(ByVal value)
		If IsNull(value) Then	 		
			DBBool = 0
			Exit Function
		End If
		
		'start error handling
		On Error Resume Next
		DBBool = cbool(value)
			
		If Err.Number <> 0 Then
			DBBool = 0
			Err.Clear
		End If
		
		'end error handling
		On Error GoTo 0
		
	End Function
	
	Public Function DBBoolNull(ByVal value)
		If IsNull(value) Then	 		
			DBString = null
			Exit Function
		End If
	
		On Error Resume Next
		DBString = cbool(value)
			
		If Err.Number <> 0 Then
			DBString = null
			Err.Clear
		End If
		
		'end error handling
		On Error GoTo 0
		
	End Function
	
	Public Function DBInteger(ByVal value)
		If IsNull(value) Then	 		
			DBInteger = -1
			Exit Function
		End If
		
		If IsNumeric(value) Then
		DBInteger = cint(value)
	Else
		DBInteger = -1
	End IF
		
	End Function
	
	Public Function DBIntegerNull(ByVal value)
		If IsNull(value) Then	 		
			DBIntegerNull = null
			Exit Function
		End If
		
		If  trim(value) = "" then
			DBIntegerNull = null
			Exit Function
		End If 
		
		If IsNumeric(value) Then
			DBIntegerNull = cint(value)
		Else
			DBIntegerNull = null
		End IF
			
	End Function
	
	
	'------------------------------------------------------
	'-- Crash reports
	Private Sub DAO_SendCrashMail(ByVal subjectParam, ByVal bodyParam)
		
		Dim DAO_CrashMailer_iMsg
		Dim DAO_CrashMailer_iConf
		Dim DAO_CrashMailer_Flds
		
		Dim DAO_loggedUserName, DAO_loggedUser_Name, DAO_IDUser
		If Not TypeName(Session(SESSION_USER)) = "Empty" Then
			
			On Error Resume Next
				DAO_loggedUserName = personItem.UserName
				DAO_loggedUser_Name = personItem.CompleteName
				DAO_IDUser = personItem.IDUser
				
			If Err.Number <> 0 Then
				DAO_loggedUserName = ""
				DAO_loggedUser_Name = ""
				DAO_IDUser = ""
			End If
			
			On Error GoTo 0
			
			
		else
			DAO_loggedUserName = ""
			DAO_loggedUser_Name = ""
			DAO_IDUser = ""
		end if
		
		set DAO_CrashMailer_iMsg = CreateObject("CDO.Message")
		set DAO_CrashMailer_iConf = CreateObject("CDO.Configuration")
		
		Set DAO_CrashMailer_Flds = DAO_CrashMailer_iConf.Fields
		
		With DAO_CrashMailer_Flds
		    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = GlobalMailSMTP
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = GlobalSMTPPort
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' 1 - pickup.   2 - port 
			if GlobalMailUser <> "" AND GlobalMailPassword <> "" Then
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1  
				.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = GlobalMailUser  
				.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = GlobalMailPassword
			end if
			.Update
		End With
		
		' apply the settings to the message
		With DAO_CrashMailer_iMsg
		    Set .Configuration 	= DAO_CrashMailer_iConf
		    .To 				= GlobalCrashMsgRecipient
		    if GlobalCrashMsgRecipient <> ccMailAddress Then
		    	.Cc				= ccMailAddress
		    end if
		    
		    .From 		= GlobalName & "<" & GlobalMailSender & ">"
		    .Subject 	= SystemName & ": " & subjectParam
		    
		    '--TEMPORAL, revisar y actualizar
		    .HTMLBody 	= "Mensaje automático de los sistemas de Pulso Vital Consulting Group<br><br>"&_
							"Sistema: " & SystemName & "<br>" &_
		 					"fecha: " & now & "<br>" &_
		 					"servidor: " & urlServer & "<br><br>" &_
							"username: " & DAO_loggedUserName & "<br>" &_
							"nombre: " & DAO_loggedUser_Name & "<br>" &_
							"id_user: " & DAO_IDUser & "<br>" &_
							"URL: " & Request.ServerVariables("SCRIPT_NAME") & "<br>" &_
							"QueryString: " & Request.QueryString() & "<br>" &_
							"<br>-Detalles-<br>" & bodyParam
		    .Send
		End With
		
		' cleanup of variables
		Set DAO_CrashMailer_iMsg = Nothing
		Set DAO_CrashMailer_iConf = Nothing
		Set DAO_CrashMailer_Flds = Nothing
		
	end sub
	
End Class

 %>