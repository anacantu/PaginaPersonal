<%
'set normal script timeout
Server.ScriptTimeOut = 90

'constantes generales
const SystemName		= "STILO version 1.0"
const siteTitle			= "Stilo - Interiorismo, Arquitectura y Dise&ntilde;o"
const siteCompanyName	= "STILO"

Dim urlServer
urlServer				= Request.ServerVariables("server_name") & "/"
const clientSystemURL	= "clientes/"
const mailSender		= "Contacto STILO <hola@stilo.com.mx>"

'datos de contacto
const publicContact	= "hola@stilo.com.mx"
const adminContact	= "Soporte Técnico Stilo <a href=""mailto:soporte@stilo.com.mx"">soporte@stilo.com.mx</a>"

'dimensiones foto thumbnail
const GlobalThumbWidth		= 160
const GlobalThumbHeight		= 100

'constantes para el envio de mails
const GlobalName			= "Revista STILO"
const GlobalMailSender 		= "soporte@stilo.com.mx"
const GlobalMailUser 		= "soporte@stilo.com.mx"
const GlobalMailNewsletterSender = "newsletter@stilo.com.mx"
const GlobalMailPassword 	= "so459st2il9"
const GlobalMailReplyTo		= ""
const GlobalMailSMTP 		= "stilo.com.mx"
const GlobalCrashMsgRecipient	= "hobbes313@gmail.com"
const GlobalSupportMail		= "soporte@stilo.com.mx"
const GlobalSMTPPort 		= 25
const GlobalMailEnding		= "<br><br>Revista STILO<br>http://www.stilo.com.mx<br>hola@stilo.com.mx"

'Estatus registro cliente
const csIncomplete		= 1
const csMailUnconfirmed	= 8
const csUnauthorized	= 9
const csAuthorized		= 10

'Cantidad maxima de categorias por negocio
const GlobalMaxBusinessCat	= 3

'Ubicacion inicial mapas de google
const googleXInit		= "25.672628"
const googleYInit		= "-100.315132"
const googleZInit		= "13"

'Localization
const currencySymbol		= "$"
const percentageSymbol		= "%"
const symbolDecimalSepartor	= "."
const currencyGroupSymbol	= ","

const facebookURL			= "http://www.facebook.com/pages/Stilo-Magazine/117905476992"
const twitterURL			= "http://twitter.com/stilomag"
const vimeoURL              = "http://www.vimeo.com/stilomagazine"
const youtubeURL            = "http://www.youtube.com/user/stilomag"
const rssURL                = "http://blog.stilo.com.mx/?feed=rss2"

const adminLogin			= "/admin/default.asp?reason=timeout"
const clientLogin			= "/clientes/default.asp?reason=timeout"

'Tiempo maximo permitido para los scripts, en segundos
const NormalServerScriptTimeOut = 90

'Tiempo maximo para procesos largos, en segundos
const MaxServerScriptTimeOut = 600 ' 10 minutos

'correo al cual enviar una copia de los mails de competencias (y recordatorios)
const ccMailAddress		= "hobbes313@gmail.com"

'-- Seguridad, variables de identificacion de usuario
Dim sessionUserType, sessionUsername, sessionUser_Name, sessionIDUser, sessionName, sessionEmail, sessionAdminMail
sessionIDUser		= session("SYS_id_user")
sessionUserType		= session("SYS_user_type")
sessionUsername		= session("STILO_username")
sessionUser_Name	= session("STILO_user_name")
sessionName			= session("STILO_name")
sessionEmail	 	= session("STILO_email")
sessionAdminMail	= session("STILO_adminMail")

Dim loadActions, fileName
fileName = GetFileName("", false)

'Mail constants
const email_IDSendUserPassword	= 1


'Get month name, in spanish
function getMonth(monthNumber)
	select case Cstr(Cint(monthNumber))
		Case "1": mes = "enero"
		Case "2": mes = "febrero"
		Case "3": mes = "marzo"
		Case "4": mes = "abril"
		Case "5": mes = "mayo"
		Case "6": mes = "junio"
		Case "7": mes = "julio"
		Case "8": mes = "agosto"
		Case "9": mes = "septiembre"
		Case "10": mes = "octubre"
		Case "11": mes = "noviembre"
		Case "12": mes = "diciembre"
	End Select
	
	getMonth = mes

end function


function GenerateThumbPictureProj(fileName)
	
	Dim GPT_picture, originalWidthPic
	Set GPT_picture = Server.CreateObject("Persits.Jpeg")

	GPT_picture.Open Server.MapPath(Application("ProjectPicPath") & fileName)
	
	aspectRatio =  GPT_picture.OriginalWidth / GPT_picture.OriginalHeight
	originalWidthPic = GPT_picture.OriginalWidth
	
	if aspectRatio > GlobalThumbWidth / GlobalThumbHeight then
		newWidth	= aspectRatio * GlobalThumbHeight
		newHeight	= GlobalThumbHeight
	else
		newWidth	= GlobalThumbWidth
		newHeight	= GlobalThumbWidth / aspectRatio
	end if
	
	GPT_picture.Width	= newWidth
	GPT_picture.Height	= newHeight
	
	if newWidth > GlobalThumbWidth Then
		GPT_picture.Crop (newWidth-GlobalThumbWidth)/2, 0, (newWidth+GlobalThumbWidth)/2, GlobalThumbHeight
	else
		GPT_picture.Crop 0, (newHeight-GlobalThumbHeight)/2, GlobalThumbWidth, (newHeight+GlobalThumbHeight)/2
	end if
	
	GPT_picture.Save Server.MapPath(Application("ProjectPicPath")) & "\" & ThumbVersion(fileName)
	set GPT_picture = nothing
	
    'GENERATE THREE DIFFERENT SIZE VERSION OF THE IMAGE:
        Set oFS = Server.CreateObject("Scripting.FileSystemObject") 

        oFS.CopyFile Server.MapPath(Application("ProjectPicPath")) & "\" & fileName, Server.MapPath(Application("ProjectPicPath")) & "\" & sizeVersion(fileName, "small")
        oFS.CopyFile Server.MapPath(Application("ProjectPicPath")) & "\" & fileName, Server.MapPath(Application("ProjectPicPath")) & "\" & sizeVersion(fileName, "medium")
        oFS.CopyFile Server.MapPath(Application("ProjectPicPath")) & "\" & fileName, Server.MapPath(Application("ProjectPicPath")) & "\" & sizeVersion(fileName, "large")
		
        call ResizeImage(Server.MapPath(Application("ProjectPicPath")) & "\" & sizeVersion(fileName, "small"), 107, 85)
        call ResizeImage(Server.MapPath(Application("ProjectPicPath")) & "\" & sizeVersion(fileName, "medium"), 235, 170)
        call ResizeImage(Server.MapPath(Application("ProjectPicPath")) & "\" & sizeVersion(fileName, "large"), 490, 290)

		if originalWidthPic > 750 then
			oFS.CopyFile Server.MapPath(Application("ProjectPicPath")) & "\" & fileName, Server.MapPath(Application("ProjectPicPath")) & "\" & sizeVersion(fileName, "fullSize")
			call ResizeImage(Server.MapPath(Application("ProjectPicPath")) & "\" & sizeVersion(fileName, "fullSize"), 750, 500)
		end if
					
        Set oFS = nothing
    'END FILE GENERATION

end function

function GenerateThumbPicture(fileName)
	
	Dim GPT_picture, originalWidthPic
	Set GPT_picture = Server.CreateObject("Persits.Jpeg")
	GPT_picture.Open Server.MapPath(Application("ArticlePicPath") & fileName)
	
	aspectRatio =  GPT_picture.OriginalWidth / GPT_picture.OriginalHeight
	originalWidthPic = GPT_picture.OriginalWidth
	
	if aspectRatio > GlobalThumbWidth / GlobalThumbHeight then
		newWidth	= aspectRatio * GlobalThumbHeight
		newHeight	= GlobalThumbHeight
	else
		newWidth	= GlobalThumbWidth
		newHeight	= GlobalThumbWidth / aspectRatio
	end if
	
	GPT_picture.Width	= newWidth
	GPT_picture.Height	= newHeight
	
	if newWidth > GlobalThumbWidth Then
		GPT_picture.Crop (newWidth-GlobalThumbWidth)/2, 0, (newWidth+GlobalThumbWidth)/2, GlobalThumbHeight
	else
		GPT_picture.Crop 0, (newHeight-GlobalThumbHeight)/2, GlobalThumbWidth, (newHeight+GlobalThumbHeight)/2
	end if
	
	GPT_picture.Save Server.MapPath(Application("ArticlePicPath")) & "\" & ThumbVersion(fileName)
	set GPT_picture = nothing
	
    'GENERATE THREE DIFFERENT SIZE VERSION OF THE IMAGE:
        Set oFS = Server.CreateObject("Scripting.FileSystemObject") 

        oFS.CopyFile Server.MapPath(Application("ArticlePicPath")) & "\" & fileName, Server.MapPath(Application("ArticlePicPath")) & "\" & sizeVersion(fileName, "small")
        oFS.CopyFile Server.MapPath(Application("ArticlePicPath")) & "\" & fileName, Server.MapPath(Application("ArticlePicPath")) & "\" & sizeVersion(fileName, "medium")
        oFS.CopyFile Server.MapPath(Application("ArticlePicPath")) & "\" & fileName, Server.MapPath(Application("ArticlePicPath")) & "\" & sizeVersion(fileName, "large")	
		
        call ResizeImage(Server.MapPath(Application("ArticlePicPath")) & "\" & sizeVersion(fileName, "small"), 107, 85)
        call ResizeImage(Server.MapPath(Application("ArticlePicPath")) & "\" & sizeVersion(fileName, "medium"), 235, 170)
        call ResizeImage(Server.MapPath(Application("ArticlePicPath")) & "\" & sizeVersion(fileName, "large"), 490, 290)
		
		if originalWidthPic > 750 then
			oFS.CopyFile Server.MapPath(Application("ArticlePicPath")) & "\" & fileName, Server.MapPath(Application("ArticlePicPath")) & "\" & sizeVersion(fileName, "fullSize")
			call ResizeImage(Server.MapPath(Application("ArticlePicPath")) & "\" & sizeVersion(fileName, "fullSize"), 750, 500)
		end if
		
        Set oFS = nothing
    'END FILE GENERATION

end function

sub ResizeImage(ByVal fileName, ByVal setWidth, ByVal setHeight)
	Dim GPT_picture
	Set GPT_picture = Server.CreateObject("Persits.Jpeg")
	GPT_picture.Open fileName
	
	aspectRatio =  GPT_picture.OriginalWidth / GPT_picture.OriginalHeight
	
	if aspectRatio > setWidth / setHeight then
		newWidth	= aspectRatio * setHeight
		newHeight	= setHeight
	else
		newWidth	= setWidth
		newHeight	= setWidth / aspectRatio
	end if
	
	GPT_picture.Width	= newWidth
	GPT_picture.Height	= newHeight
	
	if newWidth > setWidth Then
		GPT_picture.Crop (newWidth-setWidth)/2, 0, (newWidth+setWidth)/2, setHeight
	else
		GPT_picture.Crop 0, (newHeight-setHeight)/2, setWidth, (newHeight+setHeight)/2
	end if
	
	GPT_picture.Save fileName
	set GPT_picture = nothing
end sub

sub ResizeVideoImage(ByVal fileName, ByVal setWidth, ByVal setHeight)
	Dim GPT_picture
	Set GPT_picture = Server.CreateObject("Persits.Jpeg")
	GPT_picture.Open fileName

	aspectRatio =  GPT_picture.OriginalWidth / GPT_picture.OriginalHeight
	
	if aspectRatio > setWidth / setHeight then
		newWidth	= aspectRatio * setHeight
		newHeight	= setHeight
	else
		newWidth	= setWidth
		newHeight	= setWidth / aspectRatio
	end if
	
	GPT_picture.Width	= newWidth
	GPT_picture.Height	= newHeight
	
	if newWidth > setWidth Then
		GPT_picture.Crop (newWidth-setWidth)/2, 0, (newWidth+setWidth)/2, setHeight
	else
		GPT_picture.Crop 0, (newHeight-setHeight)/2, setWidth, (newHeight+setHeight)/2
	end if
	
    GPT_picture.Canvas.DrawPNG 86, 57, Server.MapPath("\images\play01.png")

	GPT_picture.Save fileName
	set GPT_picture = nothing
end sub


'Get a string representation of date
function formatDateString(dateParam, idFormat)
	formatDateString = dateParam
	
	if not IsNull(dateParam) and dateParam <> "" and IsDate(dateParam) then
		select case idFormat
		
			case 1:
				formatDateString = day(dateParam) & " " & left(getMonth(month(dateParam)),3) & " " & year(dateParam)
			case 2:
				formatDateString = left(getMonth(month(dateParam)),3) & " " & year(dateParam)
			case 3:
				formatDateString = day(dateParam) & "." & month(dateParam) & "." & year(dateParam)
		end select
	end if
end function

'Function to return one of two values, depending on boolean condition
Function CaseIf(boolCondition, valueYes, valueNo)
	if boolCondition then
		CaseIf = valueYes
	else
		CaseIf = valueNo
	end if
end Function

' Obtener nombre del archivo cargado actual
Function GetFileName(fileURL, includeURLVars)
	Dim lsPath, arPath, lclFileName
	lsPath = fileURL
	
	' Obtain the virtual file path
	if fileURL = "" Then lsPath = Request.ServerVariables("SCRIPT_NAME")
	
	arPath 		= Split(lsPath, "/")
	lclFileName	= arPath(UBound(arPath,1))
	
	' Remove vars in querystring
	if NOT includeURLVars Then
		arPath 		= Split(lclFileName, "?")
		lclFileName	= arPath(0)
	end if
	
	GetFileName = lclFileName
End Function

function RequestServerValue(valueKey, defaultValue)
	RequestServerValue = defaultValue
	
	if Request.QueryString(valueKey) <> "" then
		RequestServerValue = Request.QueryString(valueKey)
	else
		if Request.Form(valueKey) <> "" then
			RequestServerValue = Request.Form(valueKey)
		end if
	end if
	
end function

function GetRandomNumber()
	Randomize
	
	GetRandomNumber = Int((1000)*Rnd)
end function


Function RemoveHTML(ByVal strText)
	Dim RegEx

	Set RegEx = New RegExp

	RegEx.Pattern = "<[^>]*>"
	RegEx.Global = True

	RemoveHTML = RegEx.Replace(strText, "")
End Function

'replace invalid chars (' for '') in string
function removeInvalidSQLChars(textToClean)
	if textToClean <> "" then
		removeInvalidSQLChars = trim(replace(textToClean,"'","''"))
	else
		removeInvalidSQLChars = textToClean
	end if
	
end function


'If string param is empty, send NULL
function ConvertToSQLStringValue(textToConvert)
	Dim cleanText
	cleanText = removeInvalidSQLChars(textToConvert)
	
	if cleanText <> "" then
		ConvertToSQLStringValue = "'"&cleanText&"'"
	else
		ConvertToSQLStringValue = "NULL"
	end if
end function

'Convert stirng value to SQL LIKE string value
function ConvertToSQLLIKEStringValue(textToConvert)
	Dim cleanText
	cleanText = removeInvalidSQLChars(textToConvert)
	
	if cleanText <> "" then
		ConvertToSQLLIKEStringValue = "'%"&cleanText&"%'"
	else
		ConvertToSQLLIKEStringValue = "NULL"
	end if
end function

'Convert param to boolean corrected value for SQL usage
function ConvertToSQLBooleanValue(valueToConvert)
	if IsNull(valueToConvert) Then
		ConvertToSQLBooleanValue = "0"
	else
		if replace(valueToConvert," ","") = "" or valueToConvert = "0" or valueToConvert = "null" or valueToConvert = "false" or valueToConvert = false or UCASE(valueToConvert) = "NO" Then
			ConvertToSQLBooleanValue = "0"
		else
			ConvertToSQLBooleanValue = "1"
		end if
	end if
end function

'Convert param to boolean corrected value for SQL usage
function ConvertToSQLNumericValue(valueToConvert, allowNulls)
	if IsNull(valueToConvert) OR valueToConvert = "" Then
		ConvertToSQLNumericValue = "NULL"
	else
		if isNumeric(valueToConvert) Then
			ConvertToSQLNumericValue = valueToConvert
		else
			noGroupSeparator = safeReplace(valueToConvert, currencyGroupSymbol, "")
			if isNumeric(safeReplace(noGroupSeparator, symbolDecimalSepartor, "")) Then
				ConvertToSQLNumericValue = noGroupSeparator
			else
				ConvertToSQLNumericValue = "NULL"
			end if
		end if
	end if
	
	if NOT allowNulls AND ConvertToSQLNumericValue = "NULL" Then
		ConvertToSQLNumericValue = "0"
	end if	
end function

'Convert param to boolean corrected value for SQL usage
function ConvertToSQLNumericValue2(valueToConvert, allowNulls, defaultValue)
	if IsNull(valueToConvert) OR valueToConvert = "" Then
		ConvertToSQLNumericValue2 = "NULL"
	else
		if isNumeric(valueToConvert) Then
			ConvertToSQLNumericValue2 = valueToConvert
		else
			noGroupSeparator = safeReplace(valueToConvert, currencyGroupSymbol, "")
			if isNumeric(safeReplace(noGroupSeparator, symbolDecimalSepartor, "")) Then
				ConvertToSQLNumericValue2 = noGroupSeparator
			else
				ConvertToSQLNumericValue2 = "NULL"
			end if
		end if
	end if
	
	if NOT allowNulls AND ConvertToSQLNumericValue2 = "NULL" Then
		ConvertToSQLNumericValue2 = defaultValue
	end if	
end function

'-- convierte una fecha en formato ingles a formato AAAA-MM-DD para SQL
function ConvertToSQLDateFormat(dateParam)
	if not IsNull(dateParam) and dateParam <> "" and IsDate(dateParam) then
		if hour(dateParam) > 0 AND minute(dateParam) > 0 Then
			ConvertToSQLDateFormat = "'" & year(dateParam) & "-" & month(dateParam) & "-" & day(dateParam) & " " & hour(dateParam) & ":" & minute(dateParam) & "'"
		else
			ConvertToSQLDateFormat = "'" & year(dateParam) & "-" & month(dateParam) & "-" & day(dateParam) & "'"
		end if
		
	else
		ConvertToSQLDateFormat = "NULL"
	end if
end function 

function ConvertDateFormat(dateParam, sourceFormat, destFormat)
	if not IsNull(dateParam) and dateParam <> "" Then
		
		select case destFormat
			case "english":
				select case sourceFormat
					case "file": ConvertDateFormat = Mid(dateParam,5,2) & "/" & right(dateParam,2) & "/" & left(dateParam,4)
				end select
			
		end select
		
	else
		ConvertDateFormat = dateParam
	end if
end function


'------------------------------------
'-- Funciones para validar y adecuar valores a ser enviados como parameros a comandos SQL
'------------------------------------
function ConvertToSQLStringParam(textToConvert)
	Dim CSQLSP_cleanText
	CSQLSP_cleanText = removeInvalidSQLChars(textToConvert)
	
	if CSQLSP_cleanText <> "" then
		ConvertToSQLStringParam = CSQLSP_cleanText
	else
		ConvertToSQLStringParam = Null
	end if
end function


'------------------------------------
'-- END Funciones para validar y adecuar valores a ser enviados como parameros a comandos SQL
'------------------------------------


'-- add SQLstring part only if value changed
function ConcatenateSQLString(byRef SQLString, oldValue, newValue, columnKey, addComma)
	ConcatenateSQLString = addComma
	
	if IsNumeric(newValue) AND IsNumeric(oldValue) Then
		newValue = cdbl(newValue)
		oldValue = cdbl(oldValue)
	end if
	
	if newValue <> oldValue then
		
		if addComma Then SQLString = SQLString & ", "
		SQLString = SQLString & columnKey & " = " & newValue
		ConcatenateSQLString = true
	end if
end function


function FixedDigitNumber(inputNumber, digitSize)
	FixedDigitNumber = String(digitSize - Len(Cstr(inputNumber)), "0") & Cstr(inputNumber)
end function

'get today datein format dd/mm/yyyy
function getTodayString()
	todayDate = Cstr(year(date))
	if CInt(month(date)) < 10 then
		todayDate = "0" & month(date) & "/" & todayDate
	else 
		todayDate = month(date) & "/" & todayDate
	end if
	if CInt(day(date)) < 10 then ' temp codecoloring >
		todayDate = "0" & day(date) & "/" & todayDate
	else 
		todayDate = day(date) & "/" & todayDate
	end if
	getTodayString = todayDate
end function


'get DateFormat in spanish - day/month/year
function spanishDateFormat(dateParam)
	if dateParam <> "" then
		returnDate = Cstr(year(dateParam))
		if CInt(month(dateParam)) < 10 then
			returnDate = "0" & month(dateParam) & "/" & returnDate
		else 
			returnDate = month(dateParam) & "/" & returnDate
		end if
		if CInt(day(dateParam)) < 10 then' temp codecoloring >
			returnDate = "0" & day(dateParam) & "/" & returnDate
		else 
			returnDate = day(dateParam) & "/" & returnDate
		end if
		spanishDateFormat = returnDate
		'spanishDateFormat = day(dateParam) & "/" & month(dateParam) & "/" & year(dateParam)
	else
		spanishDateFormat = ""
	end if
end function

'get DateFormat in english - month/day/year, considering parameter comes in spanish - day/month/year
function englishDateFormat(dateParam)
	if IsNULL(dateParam) OR len(dateParam) = 0 then
		englishDateFormat = ""
	else
		tempdate = split(dateParam,"/")
		englishDateFormat = tempdate(1) & "/" & tempdate(0) & "/" & tempdate(2)
	end if
end function

'SQL date format yyyy-mm-dd
function sqlDateFormat(dateParam)
	sqlDateFormat = year(dateParam) & "-" & month(dateParam) & "-" & day(dateParam)
end function 

function numericDateFormat(dateParam)
	if not IsNull(dateParam) and dateParam <> "" and IsDate(dateParam) then
		numericDateFormat = year(dateParam) & FixedDigitNumber(month(dateParam), 2) & FixedDigitNumber(day(dateParam), 2)
	else
		numericDateFormat = ""
	end if
end function 


'-- Verifies if a column exists in the given recordSet
function ColumnExists(lclRecordSet, nameToCheck)
	ColumnExists = false

	For Each oField In lclRecordSet.Fields
		If LCASE(oField.Name) = LCASE(nameToCheck) Then
			ColumnExists = True
			Exit For
		End If
	Next

end function


Sub HandleError
' Send email notifying the webmaster of the site about the error

' Write the error message in a application error log file

' Display friendly error message to the user

' Stop the execution of the ASP page

End Sub

sub SendCrashMail(subjectParam, bodyParam)
	
	Dim CrashMailer_iMsg
	Dim CrashMailer_iConf
	Dim CrashMailer_Flds
	
	set CrashMailer_iMsg = CreateObject("CDO.Message")
	set CrashMailer_iConf = CreateObject("CDO.Configuration")
	
	Set CrashMailer_Flds = CrashMailer_iConf.Fields
	
	' set the CDOSYS configuration fields to use port 25 on the SMTP server
	With CrashMailer_Flds
	    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = GlobalMailSMTP
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = GlobalSMTPPort
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' 1 - pickup.   2 - port 
		if GlobalMailUser <> "" Then
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1  
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = GlobalMailUser  
			.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = GlobalMailPassword
		end if
		.Update
	End With
	
	' apply the settings to the message
	With CrashMailer_iMsg
	    Set .Configuration = CrashMailer_iConf
	    .To 		= GlobalCrashMsgRecipient
	    if GlobalCrashMsgRecipient <> ccMailAddress Then
	    	.Cc			= ccMailAddress
	    end if
	    .From 		= GlobalName & "<" & GlobalMailSender & ">"
	    .Subject 	= SystemName & ": " & subjectParam
	    .HTMLBody 	= "Mensaje automático del sitio web de STILO<br><br>"&_
						"Sistema: " & SystemName & "<br>" &_
	 					"fecha: " & now & "<br>" &_
	 					"servidor: " & urlServer & "<br><br>" &_
						"tipo de usuario: " & sessionUserType & "<br>" &_
						"username: " & sessionUsername & "<br>" &_
						"nombre: " & sessionUser_Name & "<br>" &_
						"id_user: " & sessionIDUser & "<br>" &_
						"URL: " & Request.ServerVariables("SCRIPT_NAME") & "<br>" &_
						"QueryString: " & Request.QueryString() & "<br>" &_
						"<br>-Detalles-<br>" & bodyParam
	    .Send
	End With
	
	' cleanup of variables
	Set CrashMailer_iMsg = Nothing
	Set CrashMailer_iConf = Nothing
	Set CrashMailer_Flds = Nothing
	
end sub

'--deprecated, use sendMailMessageAdvanced
function sendMail(mailSender, mailSubject, mailRecipient, mailBody)
	sendMailMessageAdvanced mailSender, mailSubject, mailRecipient, mailBody, "", "", ""
end function

'--deprecated, use sendMailMessageAdvanced
sub sendMailMessage(mailSender, mailSubject, mailRecipient, mailBody, attachmentFile, attachmentFileName)
	sendMailMessageAdvanced mailSender, mailSubject, mailRecipient, mailBody, attachmentFile, attachmentFileName, ""
end sub

sub sendMailMessageAdvanced(mailSender, mailSubject, mailRecipient, mailBody, attachmentFile, attachmentFileName, ccAddress)
	
	if mailRecipient <> "" then
		if isEmailValid(mailRecipient) Then
			
			if len(mailSender) = 0 then
				mailSender = """" & GlobalName & """ <" & GlobalMailSender & ">"
			end if
			
			' send by connecting to port 25 of the SMTP server
			Dim mail_iMsg
			Dim mail_iConf
			Dim mail_Flds
			Dim mail_CC
			
			set mail_iMsg = CreateObject("CDO.Message")
			set mail_iConf = CreateObject("CDO.Configuration")
			
			Set mail_Flds = mail_iConf.Fields
			
			' set the CDOSYS configuration fields 
			With mail_Flds
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = GlobalMailSMTP
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = GlobalSMTPPort
				.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' 1 - pickup.   2 - port 
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1  
        		.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = GlobalMailUser  
        		.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = GlobalMailPassword
				.Update
			End With
			
			mail_CC = ccMailAddress
			if ccAddress <> "" and ccAddress <> "NO" Then
				if isEmailValid(ccAddress) Then
					mail_CC = ccAddress & "," & mail_CC
				else
				
					'--direccion de correo CC incorrecta, mandar mensaje de error
					paramStr = "Sender: " & replace( replace(mailSender,"<","&lt;"), ">", "&gt;") &_
					"<br>Subject: " & mailSubject &_
					"<br>Recipient: " & mailRecipient &_
					"<br>Carbon Copy: " & ccAddress &_
					"<br>Body: " & mailBody
						
					'Enviar reporte de correo invalido
					SendCrashMail "Cuenta de correo incorrecta", "Se ha intentado enviar un correo (CC) a un email invalido.<br>email: &lt;"&_
					ccAddress &"&gt;<br><br>" & paramStr
				end if
			end if
			
			'error handling
			On Error Resume Next
			
			' apply the settings to the message
            if ccAddress = "NO" then
			    With mail_iMsg
				    Set .Configuration = mail_iConf
				    .To		= mailRecipient
				    .From		= mailSender
				    .Subject	= mailSubject
				    .HTMLBody	= mailBody

				    if attachmentFile <> "" then
					    .AddAttachment attachmentFile	
					
					    if attachmentFileName <> "" then				
						    .Attachments(1).Fields.Item("urn:schemas:mailheader:content-disposition") = "attachment;filename=" & attachmentFileName
						    .Attachments(1).Fields.Update
					    end if
				    end if
				
				    .Send
			    End With
			
            else

			    With mail_iMsg
				    Set .Configuration = mail_iConf
				    .To		= mailRecipient
				    .Cc		= mail_CC
				    .From		= mailSender
				    .Subject	= mailSubject
				    .HTMLBody	= mailBody

				    if attachmentFile <> "" then
					    .AddAttachment attachmentFile	
					
					    if attachmentFileName <> "" then				
						    .Attachments(1).Fields.Item("urn:schemas:mailheader:content-disposition") = "attachment;filename=" & attachmentFileName
						    .Attachments(1).Fields.Update
					    end if
				    end if
				
				    .Send
			    End With
			
            end if				

			If Err.Number <> 0 then
				
				paramStr = "Sender: " & replace( replace(mailSender,"<","&lt;"), ">", "&gt;") &_
				"<br>Subject: " & mailSubject &_
				"<br>Recipient: " & mailRecipient &_
				"<br>Body: " & mailBody
			
				SendCrashMail "Error al enviar correo", "Error al enviar correo electronico:<br><strong>"&_
				Err.Description & "</strong><br><br>" & paramStr
			end if
			
			'end error handling
			On Error GoTo 0
			
			' cleanup of variables
			Set mail_iMsg = Nothing
			Set mail_iConf = Nothing
			Set mail_Flds = Nothing
			set mail_CC = Nothing
			
		else
		
			paramStr = "Sender: " & replace( replace(mailSender,"<","&lt;"), ">", "&gt;") &_
			"<br>Subject: " & mailSubject &_
			"<br>Recipient: " & mailRecipient &_
			"<br>Body: " & mailBody
				
			'Enviar reporte de correo invalido
			SendCrashMail "Cuenta de correo incorrecta", "Se ha intentado enviar un correo a un email invalido.<br>email: &lt;"&_
			mailRecipient &"&gt;<br><br>" & paramStr
		
		end if
	end if
	
end sub



'Verificar si un correo es valido
Function isEmailValid(email) 
        Set regEx = New RegExp 
        regEx.Pattern = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w{2,}$" 
        isEmailValid = regEx.Test(trim(email)) 
End Function 

'remover caracteres invalidos
function removeInvalidXMLChars(ByVal InString)
	
	if InString <> "" then
		InString = replace(InString,"á","a")
		InString = replace(InString,"é","e")
		InString = replace(InString,"í","i")
		InString = replace(InString,"ó","o")
		InString = replace(InString,"ú","u")
        InString = replace(InString,"²","2")
        InString = replace(InString,"ñ","&ntilde;")
	
		InString = trim(replace(InString,"&","&amp;"))
		removeInvalidXMLChars = replace(InString,"<","&lt;")
	else
		removeInvalidXMLChars = InString
	end if
	
end function

function normalizeString(ByVal InString)

	if InString <> "" then
		InString = replace(InString," ","")
		InString = replace(InString,"'","")
		InString = replace(InString,"á","a")
		InString = replace(InString,"é","e")
		InString = replace(InString,"í","i")
		InString = replace(InString,"ó","o")
		InString = replace(InString,"ú","u")
		
		normalizeString = safeTrim(LCase(InString))
	else
		normalizeString = InString
	end if
end function
		

'remover caracteres invalidos
function removeInvalidHTMLChars(InString)
	
	if InString <> "" then
		InString = trim(InString)
		removeInvalidHTMLChars = replace(InString,"<","&lt;")
	else
		removeInvalidHTMLChars = InString
	end if
end function

function RenderHTML(text)
	Dim RHTML
	RHTML = trim(text)
	if RHTML <> "" then
		RHTML=replace(RHTML,vbCrLf,"<BR>")
		RHTML=replace(RHTML,vbCr,"<BR>")
		RHTML=replace(RHTML,vbLf,"<BR>")
	end if
	RenderHTML= RHTML
	
end function

function renderLineFeed(text)
	text = trim(text)
	
	if text <> "" then
		text=replace(text,vbCr,"")
		text=replace(text,vbLf,"")
		text=replace(text,vbCrLf,"")
		text=replace(text,"<BR>",vbCrLf)
		text=replace(text,"<br>",vbCrLf)
	end if
	renderLineFeed = text
end function

Function renderTAGValue (byVal stringVal)
	if isEmpty(stringVal) or isNull(stringVal) then
		renderTAGValue = ""
	else
		renderTAGValue = Replace(Replace(Server.HTMLEncode(stringVal),"""", "&dblquote;" ), "'", "&#39;")
	end if
End Function

Function renderJSValue (s)
	if isEmpty(s) or isNull(s) then
		renderJSValue = ""
	else
		renderJSValue = renderTAGValue( Replace( Replace( Replace( s, "\", "\\"), """", "\""" ) , "'", "\'" ))
	end if
End Function

function RemoveLastSeparator(listString, separator)
	if len(listString) = 0 Then
		RemoveLastSeparator = ""
	else
		if Right(listString,len(separator)) = separator Then
			RemoveLastSeparator = left(listString,len(listString) - len(separator))
		else
			RemoveLastSeparator = listString
		end if
	end if
end function

function RemoveFirstSeparator(listString, separator)
	if len(listString) = 0 Then
		RemoveFirstSeparator = ""
	else
		if Left(listString,len(separator)) = separator Then
			RemoveFirstSeparator = right(listString,len(listString) - len(separator))
		else
			RemoveFirstSeparator = listString
		end if
	end if
end function

'-- Remove first and last separator
function TrimSeparator(listString, separator)
	if len(listString) = 0 Then
		TrimSeparator = ""
	else
		TrimSeparator = RemoveFirstSeparator(RemoveLastSeparator(listString, separator), separator)
	end if
end function

'Replace que acepta nulos
function SafeReplace(inputString, textToFind, replaceValue)
	if len(replaceValue) > 0 then
		SafeReplace = replace(inputString, textToFind, replaceValue)
	else
		SafeReplace = replace(inputString, textToFind, "")
	end if
end function

function safeCint(paramToConvert)
	if paramToConvert = "" Then
		safeCint = ""
	else
		if IsNumeric(paramToConvert) Then
			safeCint = Cint(paramToConvert)
		else
			safeCint = ""
		end if
	end if
end function

function safeCdbl(paramToConvert)
	'error handling
	On Error Resume Next
	safeCdbl = Cdbl(paramToConvert)
	
	If Err.Number <> 0 then safeCdbl = ""
	
	'end error handling
	On Error GoTo 0
	
end function

function SafeCbool(paramToConvert)
	if IsNull(paramToConvert) Then
		SafeCbool = false
	else
		SafeCbool = Cbool(paramToConvert)
	end if
end function


function SafeCstr(paramToConvert)
	if IsNull(paramToConvert) Then
		SafeCstr = ""
	else
		SafeCstr = Cstr(paramToConvert)
	end if
end function

function safeTrim(paramToConvert)
	if paramToConvert = "" Then
		safeTrim = ""
	else
		safeTrim = trim(paramToConvert)
	end if
end function

'-- round numeric data
function safeRound(paramToConvert, decimals)
	
	Dim SR_decimals, SR_paramToConvert
	
	if decimals = "" Then
		SR_decimals = 0
	else
		if IsNull(paramToConvert) Then
			SR_decimals = 0
		else
			if IsNumeric(paramToConvert) Then
				SR_decimals = safeCint(paramToConvert)
			else
				SR_decimals = 0
			end if
		end if
	end if
	
	SR_paramToConvert = safeCdbl(paramToConvert)
	
	if SR_paramToConvert = "" Then 
		safeRound = ""
	else
		if IsNull(SR_paramToConvert) Then
			safeRound = ""
		else
			if IsNumeric(SR_paramToConvert) Then
				safeRound = FormatNumber(SR_paramToConvert,decimals)
			else
				safeRound = ""
			end if
		end if
	end if
	
	set SR_decimals = nothing
	set SR_paramToConvert = Nothing
	
end function


function SafeHTMLValue(paramToConvert)
	Dim SHTML_response
	
	if paramToConvert = "" OR IsNull(paramToConvert) Then
		SafeHTMLValue = ""
	else
		SHTML_response = trim(paramToConvert)
		SHTML_response = replace(SHTML_response,"""","&quot;")
		SHTML_response = replace(SHTML_response,"<","&lt;")
		SHTML_response = replace(SHTML_response,">","&gt;")
		
		SafeHTMLValue = SHTML_response
	end if
	
	set SHTML_response = Nothing
end function


function renderStarList(starCount)
	Dim lclStarList
	
	if IsNull(starCount) OR starCount = "" Then
		lclStarList = ""
	else
		if isNumeric(starCount) Then
			for i=1 to 5
				if i <= starCount Then
					lclStarList = lclStarList & "<img src=""/images/iconos/star.png"" width=""11"" height=""11"">"
				else
					lclStarList = lclStarList & "<img src=""/images/iconos/star_off.png"" width=""11"" height=""11"">"
				end if
			next
		else
			lclStarList = ""
		end if
	end if
	
	renderStarList = lclStarList
end function


'-----------------------------
' FILE FUNCTIONS
'-----------------------------

' Makes sure that given file name does not contain path info
Function SecureFileName(name)
	SecureFileName = replace(name,"/","?")
	SecureFileName = replace(SecureFileName,"\","?") '"
End Function

' Adds given type of the slash to the end of the path if required
Function FixPath(path, slash)
	If Right(path, 1) <> slash Then
		FixPath = path & slash
	Else
		FixPath = path
	End If
End Function

' Converts the given path to physical path
Function RealizePath(thePath)
	Dim path
	path = replace(thePath,"/","\") '"
	If left(path,1) = "\" Then '"
		on error resume next
		RealizePath = FixPath(server.MapPath(path),"\") '"
		If err.Number<>0 Then RealizePath = thePath
	Else
		If InStr(1,path, ":", 1) <> 0 Then
			RealizePath = FixPath(path,"\") '"
		Else
			RealizePath = thePath & "?"
		End If
	End If
End Function

' Formats given size in bytes,KB,MB and GB '"
Function FormatSize (givenSize)
	If (givenSize < 1024) Then
		FormatSize = givenSize & " B"
	ElseIf (givenSize < 1024*1024) Then
		FormatSize = FormatNumber(givenSize/1024,2) & " KB"
	ElseIf (givenSize < 1024*1024*1024) Then
		FormatSize = FormatNumber(givenSize/(1024*1024),2) & " MB"
	Else
		FormatSize = FormatNumber(givenSize/(1024*1024*1024),2) & " GB"
	End If
End Function

function GetFileExtension(ByVal fileName)
	if fileName <> "" Then
		fileName = replace(fileName, "'", "")
			
		Dim GFEArray
		GFEArray = split(fileName,".")
		
		GetFileExtension = GFEArray(UBound(GFEArray))
	else
		GetFileExtension = ""
	end if
end function

function GetFileIcon(ByVal fileExtension)

	select case fileExtension
    	case "bmp", "doc", "jpg", "pdf", "ppt", "txt", "xls", "zip":
    		GetFileIcon = fileExtension & ".png"
    	
    	case "png", "gif":
    		GetFileIcon = "jpg.png"
    	
    	case else:
    		GetFileIcon = "txt.png"
    	
    end select
end function


function sizeVersion(ByVal fileName, ByVal sizeImg)
	Dim TV_fileExtension
	TV_fileExtension = GetFileExtension(fileName)
	
	sizeVersion = left(fileName, len(fileName) - 1 - len(TV_fileExtension)) & "_" & sizeImg & "." & TV_fileExtension
	
end function


function ThumbVersion(ByVal fileName)
	Dim TV_fileExtension
	TV_fileExtension = GetFileExtension(fileName)
	
	ThumbVersion = left(fileName, len(fileName) - 1 - len(TV_fileExtension)) & "_thumb." & TV_fileExtension
	
end function

function IsImage(fileName)
	
	if fileName = "" then
		IsImage = false
	else
		fileExtension = right(fileName, 3)
		
		select case fileExtension
	    	case "bmp", "jpg", "png", "gif":
	    		IsImage = true
	    	
	    	case else:
	    		IsImage = false
	    end select
	    
	   end if
end function

'-- XML response for RPC
sub WriteXMLList(RPC_sqlString)
	response.ContentType = "text/xml"
		
	Response.Write "<?xml version=""1.0"" encoding=""iso-8859-1""?>"
	Response.Write "<items>"
	
	Set RPC_dbObject = Server.CreateObject("ADODB.Recordset")
	RPC_dbObject.Open RPC_sqlString, oCNDB
	    			
	while NOT RPC_dbObject.EOF
	    Response.Write "<item id="""&RPC_dbObject("id_item")&""">" & removeInvalidXMLChars(RPC_dbObject("name")) & "</item>"
	    RPC_dbObject.movenext()
	wend
	RPC_dbObject.close()
	set RPC_dbObject = nothing
	
	Response.Write "</items>"
end sub

sub WriteSteppedListResponse(RPC_sqlString, columnsArray, startIndex, bulkSize)
	response.ContentType = "text/html; charset=iso-8859-1"
	Set WSLT_rs = Server.CreateObject("ADODB.Recordset")
	
	'tope de registros
	RPC_sqlString = replace(RPC_sqlString, "/*TopSelect*/", "TOP " & Cstr( (startIndex + 1)*bulkSize + 1 ))
	
	'abrir recordset
	WSLT_rs.Open RPC_sqlString, oCNDB, 3, 3
	
	if WSLT_rs.EOF Then
		response.write "0"
	else
	
		Dim WSLT_startIndex, WSLT_bulkSize
		WSLT_startIndex = startIndex
		WSLT_bulkSize	= bulkSize
		
		'recorrer registros ya enviados
		While (NOT WSLT_rs.EOF) AND (WSLT_startIndex > 0)
			WSLT_rs.movenext()
			
			WSLT_startIndex = WSLT_startIndex -1
		Wend
		
		cols = split(columnsArray,",")
		
		while (NOT WSLT_rs.EOF) AND (WSLT_bulkSize > 0)
			strResponse =""
			
			for each item in cols
				item = trim(item)
				
				//DEBUG LINE: encontrar columna faltante
				//response.write item & ","
				
				RPC_ItemValue = WSLT_rs(item)
				
				if RPC_ItemValue <> "" Then
					RPC_ItemValue = replace(RPC_ItemValue, "|", "-")
				end if
				
				strResponse = strResponse & RPC_ItemValue & "|"
			next
			
			' -- remover ultimo separador
			if right(strResponse,1) = "|" Then
				strResponse = left(strResponse, len(strResponse)-1)
			end if
			
			' -- imprimir con salto de linea
			response.write strResponse & "~"
			
			WSLT_rs.movenext()
			WSLT_bulkSize = WSLT_bulkSize - 1
		wend
		
		'Si aun quedan registros, agregar bandera con datos de control
		if NOT WSLT_rs.EOF Then
			response.write "*more*|"& startIndex + bulkSize &"|"& startIndex + 2*bulkSize - 1
		end if
	end if
	
	WSLT_rs.close()
	set WSLT_rs = nothing
	
end sub


sub WriteXMLStructuredData(RPC_sqlString, columnsArray)
	
	response.ContentType = "text/xml"
	Response.Write "<?xml version=""1.0"" encoding=""iso-8859-1""?>"
	Response.Write "<items>"

	Set WXMLSD_rs = Server.CreateObject("ADODB.Recordset")
	
	cols = split(columnsArray,",")
	
	'abrir recordset
	WXMLSD_rs.Open RPC_sqlString, oCNDB, 3, 3
	
	while (NOT WXMLSD_rs.EOF)
		
		response.write "<item>"
		for each item in cols
			response.write "<"& item &">" & removeInvalidXMLChars(WXMLSD_rs(item)) & "</"& item &">"
		next
		response.write "</item>"
		
		WXMLSD_rs.movenext()
	wend
		
	Response.Write "</items>"
	
	WXMLSD_rs.close()
	set WXMLSD_rs = nothing
	
end sub


function GetCorrectURL(inputURL)
	
	if IsNull(inputURL) OR inputURL = "" Then
		GetCorrectURL = ""
	else
		Dim leftURL
		leftURL = left(inputURL,5)
		
		if leftURL = "ftp:/" OR leftURL = "http:" then
			GetCorrectURL = inputURL
		else
			GetCorrectURL = "http://" & inputURL
		end if
	end if
	
end function


sub PrintInfoMessage (messageString)
	
	response.write "<div class=""msgInfo"">"& messageString &"</div>"
	
end sub

sub PrintBackButtonNavigation(defaultURL, aprovedURL, rejectedURL)
	Dim PBBN_referer, PBBN_badURL
	PBBN_referer = Request.ServerVariables("HTTP_REFERER")
	
	if PBBN_referer = "" OR NOT IsArray(rejectedURL) Then
		response.write("javascript:history.back(-1)")
	else
		PBBN_referer = GetFileName(PBBN_referer, false)
		PBBN_badURL = false
		
		for i=0 to UBound(rejectedURL)
			if LCase(rejectedURL(i)) = LCase(PBBN_referer) Then PBBN_badURL = true
		next
		
		if PBBN_badURL Then
			response.write("Navigate('"& defaultURL &"')")
		else
			response.write("javascript:history.back(-1)")
		end if
		
	end if
end sub

'-- Context Menu
' displayType (absolute / relative)
sub PrintContextMenu (displayType, menuTitle, optionsList)
	Dim PCM_index
	PCM_index = 1
	
	if menuTitle = "" Then menuTitle = "OPCIONES"
	
	response.write vbcrlf & "<!-- context menu -->"
	response.write "<div  id=""cm_holder"" style=""position:"&displayType&"; text-align:left; top:0; left:0"">"
	
	'-- header
	response.write "<div id=""cm_header_container"" style=""display:none""><div id=""cm_header""> "&menuTitle&" </div></div>"
	
	'-- options list
	response.write "<div id=""cm_menu"" onclick=""cm_bShow=true"" style=""display:none"">"
	for listIndex = 0 To uBound(optionsList)
		if optionsList(listIndex) = "*s*" Then
			'-- separator
			response.write "<div class=""cm_menu_separator""></div>"
		else
			response.write "<span id=""cm_menu_"& PCM_index &""">"& optionsList(listIndex) &"</span>"
			PCM_index = PCM_index + 1
		end if
	Next
	
	response.write "</div></div>" & vbcrlf
	
end sub

sub CloseAndRedirect(redirectURL)
	if Isobject(oCNDB) Then
		'--close connections
		if oCNDB.State > 0 Then oCNDB.Close
		set oCNDB = Nothing
	end if
	
	'-- redirect
	if redirectURL = "" Then
		if sessionUserType = "admin" Then 
			response.redirect(adminLogin)
		else
			response.redirect(clientLogin)
		end if
	else
		response.redirect(redirectURL)
	end if
end sub

function GetBusinessPicturesCount(IDBusiness)
	Dim FSO, folderPath, picCount
	picCount = 0
	Set FSO = server.CreateObject ("Scripting.FileSystemObject")
	folderPath = RealizePath(Application("ClientLogoPath"))
	
	for i = 1 to 5
		if FSO.FileExists(folderPath & IDBusiness & "_" & i & ".jpg") Then
			picCount = picCount + 1
		end if
	Next
	
	Set FSO = Nothing
	Set folderPath = nothing
	
end function								

function GetBusinessLogo(IDBusiness, smallSize)
	'-- FALTA HACER BUSQUEDA DE OTRAS EXTENSIONES
	if NOT smallSize Then
		GetBusinessLogo = "../files/images/logos/" & IDBusiness & ".gif"
	else
		GetBusinessLogo = "../files/images/logos/" & IDBusiness & "_small.gif"
	end if
	
end function


'--RTE - Rich Text Editor

'Funcion para filtrar el texto a desplegar en el control RTE
function RTESafe(strText)
	'returns safe code for preloading in the RTE
	dim tmpString
	
	tmpString = trim(strText)
	
	'convert all types of single quotes
	tmpString = replace(tmpString, chr(145), chr(39))
	tmpString = replace(tmpString, chr(146), chr(39))
	tmpString = replace(tmpString, "'", "&#39;")
	
	'convert all types of double quotes
	tmpString = replace(tmpString, chr(147), chr(34))
	tmpString = replace(tmpString, chr(148), chr(34))
'	tmpString = replace(tmpString, """", "\""")
	
	'replace carriage returns & line feeds
	tmpString = replace(tmpString, chr(10), " ")
	tmpString = replace(tmpString, chr(13), " ")
	
	RTESafe = tmpString
end function


Function ConvUCase(strName)

   Dim name, int, myArray, elm

   IF (InStr(strName," ")) Then

       myArray = Split(strName, " ")
       int = 0
       For Each elm in myArray

           myArray(int) = UCase(Left(myArray(int),1)) & Right(myArray(int), Len(myArray(int))-1)

            If (InStr(strName,".")) Then

                myArray(int) = UCase(Left(myArray(int),InStr(myArray(int),".")+1)) &  Right(myArray(int),(Len(myArray(int))) - (InStr(myArray(int),".") +1))

            End If
            int = int + 1
       Next

       name = Join(myArray)

   Else
        If (InStr(strName,".")) Then

            name = UCase(Left(strName,1)) & Right(strName,Len(strName)-1)
            name = UCase(Left(strName,InStr(strName,".")+1)) & Right(strName,(Len(strName)) - (InStr(strName,".")+1))
        Else
            name = UCase(Left(strName,1)) & Right(strName,Len(strName)-1)
        End If
   End IF

   convUCase = name

End Function



Function encrypt(text)
   textEncrypted = ""
   For i = 1 to Len(text)
      j = Mid(text, i, 1)
      k = Asc(j) 
      if k >= 97 and k =< 109 then
         k = k + 13 
      elseif k >= 110 and k =< 122 then
         k = k - 13 
      elseif k >= 65 and k =< 77 then
         k = k + 13 
      elseif k >= 78 and k =< 90 then
         k = k - 13
      end if

   textEncrypted = textEncrypted & Chr(k)

   Next

encrypt = textEncrypted

End Function


function create_links(strText)
    strText = " " & strText
    strText = ereg_replace(strText, "(^|[\n ])([\w]+?://[^ ,""\s<]*)", "$1<a href=""$2"" target=""_blank"">$2</a>")
    strText = ereg_replace(strText, "(^|[\n ])((www|ftp)\.[^ ,""\s<]*)", "$1<a href=""http://$2"" target=""_blank"">$2</a>")
    strText = ereg_replace(strText, "(^|[\n ])([a-z0-9&\-_.]+?)@([\w\-]+\.([\w\-\.]+\.)*[\w]+)", "$1<a href=""mailto:$2@$3"">$2@$3</a>")
    strText = right(strText, len(strText)-1)
    create_links = strText
end function

function ereg_replace(strOriginalString, strPattern, strReplacement)
    ' Function replaces pattern with replacement
    dim objRegExp : set objRegExp = new RegExp
    objRegExp.Pattern = strPattern
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    ereg_replace = objRegExp.replace(strOriginalString, strReplacement)
    set objRegExp = nothing
end function


%>