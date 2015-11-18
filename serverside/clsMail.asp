<%
' ------------------------------------------------------------------------------
'	Author:		Carlos Trevino @ STILO
'	Email:		hobbes313@gmail.com
'	URL:		http://www.stilo.com.mx
'	Date:		Jan 10, 2010
' ------------------------------------------------------------------------------
'
'	NOTE: 	data_storage/database.asp 
'			serverside/functions.asp	MUST BE LOADED
'
' Demo:
'	Set mailItem = new clsMail(IDMail)
'		Initializes object 
'
'	mailItem.Subject
'		mail's subject
'

Class clsMail
' ------------------------------------------------------------------------------
	
	Private p_idMail
	Private p_senderName
	Private p_senderEmail
	Private p_subject
	Private p_message
	Private p_description
	
	Private DBConn
	
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
	End Sub
' ------------------------------------------------------------------------------
	Public Sub LoadData(paramIDMail)
		
		if NOT IsNumeric(paramIDMail) Then exit sub
		
		p_idMail = paramIDMail
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		SQLString = "EXEC clsMail @idMail=" & paramIDMail
		DBConn.Open SQLString, oCNDB, 3, 3
		
		if NOT DBConn.EOF then
			
			p_idMail		= Cdbl(DBConn("id_mail"))
			p_senderName 	= DBConn("from_name")
			p_senderEmail 	= DBConn("from_email")
			p_subject 		= DBConn("subject")
			p_message 		= DBConn("message")
			p_description 	= DBConn("description")
		else
			p_idMail		= -1
		end if
		
		DBConn.close
		
	End Sub	
' ------------------------------------------------------------------------------
	Private Sub Class_Terminate()
		
		Set DBConn = Nothing
		
	End Sub

' ------------------------------------------------------------------------------
	Public Property Get IDMail()
		IDMail	 = p_idMail
	End Property
' ------------------------------------------------------------------------------
	Public Property Get SenderName()
		SenderName	 = p_senderName
	End Property
' ------------------------------------------------------------------------------
	Public Property Get SenderMail()
		SenderMail = p_senderEmail
	End Property
' ------------------------------------------------------------------------------
	Public Property Get From()
		From = ""
		
		if p_senderName <> "" and p_senderEmail <> "" then
			From = """" & p_senderName & """ <" & p_senderEmail & ">"
		else
			if p_senderName <> "" then From = p_senderName
			if p_senderEmail <> "" then From = p_senderEmail
		end if
		
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Subject()
		Subject	 = p_subject
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Message()
		Message	 = p_message
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Ending()
		Ending	 = GlobalMailEnding
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Description()
		Description	 = p_description
	End Property	

' ------------------------------------------------------------------------------
End Class
' ------------------------------------------------------------------------------
%>