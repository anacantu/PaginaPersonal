<!--#include virtual="/serverside/database.asp" -->
<!--#include virtual="/serverside/functions.asp" -->
<!--#include virtual="/serverside/DALC.asp" -->
<!--#include virtual="/serverside/clsMail.asp" -->
<%
'-- RemoteProcedureCalls --
'-- method: SendMail

'Definicion de variables
Dim SM_methodName, SM_recipientMail, SM_sharedEmail, SM_emaiID, mailInfo, mailMessage, SM_dataObj, SM_IDEvaluation
Dim SM_employeeName, SM_employeeUser, SM_employeePass, SM_mailItem

SM_sharedEmail = false
SM_methodName = RequestServerValue("method", "")

set SM_mailItem = new clsMail

Select case SM_methodName
	
	case "SendUserPassword":
		
		SM_employeeName		= session("name") 
		SM_recipientMail	= session("email")
		SM_employeePass		= session("password") 
		
		SM_mailItem.LoadData(email_IDSendUserPassword)
		
		if SM_mailItem.IDMail <> -1 Then
				
			mailMessage = SM_mailItem.Message & SM_mailItem.Ending
			mailMessage = SafeReplace(mailMessage,"[link]","http://"&urlServer&clientSystemURL)
			mailMessage = SafeReplace(mailMessage,"[user_name]", SM_employeeName)
			mailMessage = SafeReplace(mailMessage,"[user_username]", SM_recipientMail )
			mailMessage = SafeReplace(mailMessage,"[user_password]", SM_employeePass)
			
			//enviar correo	
			sendMail SM_mailItem.From, SM_mailItem.Subject, SM_recipientMail, mailMessage
			
		end if

    case "PasswordRecoveryPortafolios":

		SM_employeeName		= session("name") 
		SM_recipientMail	= session("email")
		SM_employeePass		= session("password") 
        SM_cuenta		    = session("cuenta") 
		
		SM_mailItem.LoadData(1)
		
		if SM_mailItem.IDMail <> -1 Then
				
			mailMessage = SM_mailItem.Message & SM_mailItem.Ending
			mailMessage = SafeReplace(mailMessage,"[link]","http://www.stilo.com.mx/portafolio/login.asp")
			mailMessage = SafeReplace(mailMessage,"[user_name]", SM_employeeName)
			mailMessage = SafeReplace(mailMessage,"[user_username]", SM_recipientMail )
			mailMessage = SafeReplace(mailMessage,"[user_password]", SM_employeePass)
            mailMessage = SafeReplace(mailMessage,"[cuenta]", SM_cuenta)
			
			//enviar correo	
			//sendMail SM_mailItem.From, SM_mailItem.Subject, SM_recipientMail, mailMessage
            sendMailMessageAdvanced SM_mailItem.From, SM_mailItem.Subject, SM_recipientMail, mailMessage, "", "", "NO"
			
		end if
		
end select


' -- Limpiar memoria
set SM_methodName = Nothing
set SM_IDRecord = nothing
set IDString = Nothing
set SM_recipientMail = nothing
set SM_sharedEmail = Nothing
set SM_emailID = Nothing
set mailInfo = nothing 
set mailMessage = nothing
set SM_dataObj = Nothing
set SM_employeeName = Nothing
set SM_employeeUser = Nothing
set SM_employeePass = Nothing
set SM_mailItem = Nothing

%>
<!--#include virtual="/serverside/functions_end.asp" -->
<!--#include virtual="/serverside/database_close.asp" -->