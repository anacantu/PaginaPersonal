<!--#include virtual="/serverside/database.asp" -->
<!--#include virtual="/serverside/functions.asp" -->
<!--#include virtual="/serverside/DALC.asp" -->
<%
'-- RemoteProcedureCalls --
'Metodos invocados por codigo tipo AJAX

Dim strResponse, RPC_IDItem, textAnswer, RPC_dbObject, RPC_Value, SQLString, rs, RPC_pageSize, RPC_startIndex

select case request("method")

    '-- Mensajes nuevos plataforma Stilo
    case "mensajesPlataforma":
        RPC_Value = ConvertToSQLNumericValue(request.Form("idUsuario"), false)
        RPC_Value2 = ConvertToSQLStringValue(request.Form("tipoUsuario"))

        if RPC_Value <> "NULL" then

			Set rs = Server.CreateObject("ADODB.Recordset")
			SQLString = "PS_GetMensajesPortafolio @idUsuario=" & RPC_Value & ", @tipoUsuario=" & RPC_Value2
			rs.Open SQLString, oCNDB, 3, 3

        	    While NOT rs.EOF
                        response.Write(rs("Mensaje") & "|" & rs("Notificaciones"))
                    rs.moveNext()
                Wend

            if NOT rs.BOF Then rs.movefirst 
            rs.close

        end if


    case "envioNotificaciones":

        Dim proceso
        RPC_Value = request("id")
        RPC_Value2 = request("action")

			Set rs = Server.CreateObject("ADODB.Recordset")
			SQLString = "GetEventos_Envio " & RPC_Value & ", '" & RPC_Value2 & "'"
			rs.Open SQLString, oCNDB, 3, 3

                if rs.EOF then

                    proceso = "fin"
                else
                    
                    proceso = "enviando"
        	        While NOT rs.EOF

                            vCorreo = rs("email")
                            vNombre = rs("name")
                            vGenero = rs("genero")
                            vSubject = rs("Subject") 
                            vBodyEmail = rs("textBodyEmailClean")
                            vArchivoAdjunto = rs("archivoAdjunto")

                        rs.moveNext()
                    Wend
                end if

            if NOT rs.BOF Then rs.movefirst 
            rs.close

    Dim mailSenderSusc
    mailSenderSusc = "Eventos Stilo <eventos@stilo.com.mx>"
    
    ' send by connecting to port 25 of the SMTP server
    Dim mail_iMsg
    Dim mail_iConf
    Dim mail_Flds
    Dim mail_CC
    Dim mailBcc    
			
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
			
            mailSubject = vSubject

            if RPC_Value2 = "test" then
                vNombre = "Fernando"
                vCorreo = "rubio@stilo.com.mx"                
                mailRecipient = "rubio@stilo.com.mx"
                mailBcc = "hobbes313@gmail.com"		
                vGenero = "m"		
            end if

            if RPC_Value2 = "procesando" then
                mailRecipient = vCorreo
            end if

            mailBody = "<br /><br /><div style=""width:786px;text-align:center""><font size=""1"">Si no puedes visualizar correctamente este mensaje haz <a href=""http://www.stilo.com.mx/rsvp/evento.asp?id=" & RPC_Value & "&email=" & vCorreo & """ target=""_blank"">clic aqu&iacute;</a>.</font></div>"
            
			if RPC_Value = 1 or RPC_Value = 8 then
						
				    if (LCase(vGenero) = "m" or LCase(vGenero) = "f") and vNombre <> "" and len(vNombre) > 2 then
					    if LCase(vGenero) = "m" then
						    mailBody = mailBody & "Estimado " & vNombre
						    mailBody = mailBody & "<br /><br />"
					    end if

					    if LCase(vGenero) = "f" then
						    mailBody = mailBody & "Estimada " & vNombre
						    mailBody = mailBody & "<br /><br />"
					    end if                
				    else
					    mailBody = mailBody & "&iexcl;Buen d&iacute;a!"
					    mailBody = mailBody & "<br /><br />"
				    end if					
			else

						if RPC_Value = 2 then
							mailBody = mailBody & "<div style=""width:786px;text-align:center; margin-top:22px"">Estimado Acad&eacute;mico</div>"
							mailBody = mailBody & "<br />"			
						else
							mailBody = mailBody & "<div style=""width:786px;text-align:center; margin-top:22px"">&nbsp;</div>"
							mailBody = mailBody & "<br />"			
						end if	
			end if
			
			
            mailBody = mailBody & vBodyEmail
            mailBody = mailBody & "<br /><br />Para confirmar su asistencia haga clic en la siguiente liga:<br /> <a href=""http://www.stilo.com.mx/rsvp/default.asp?id=" & RPC_Value & "&email=" & vCorreo & """ target=""_blank"">http://www.stilo.com.mx/rsvp/default.asp?id=" & RPC_Value & "&email=" & vCorreo & "</a><br /><br />"
            mailBody = mailBody & "<br /><center><img src=""http://www.stilo.com.mx/files/pictures/" & vArchivoAdjunto & """ alt="""" /></center><br /><br />" & _                 
                "<center><font size=""1"">Este mensaje fue enviado a " & vCorreo & ". Si no eres el usuario o <br />si deseas ser borrado de nuestro listado de env&iacute;os s&oacute;lo haz <a href=""http://www.stilo.com.mx/rsvp/cancelarSuscripcion.asp?id=" & RPC_Value & "&email=" & vCorreo & """ target=""_blank"">click aqu&iacute;</a>.</font></center>"

			'error handling
			On Error Resume Next
			

            if proceso = "enviando" then
			    ' apply the settings to the message
			    With mail_iMsg
				    Set .Configuration = mail_iConf
                    .To     = mailRecipient
				    .Bcc	= mailBcc
				    .From		= mailSenderSusc
				    .Subject	= mailSubject
				    .HTMLBody	= mailBody
                    '.CreateMHTMLBody mailBody
								
				    .Send
			    End With
			
			    If Err.Number <> 0 then
				
				    paramStr = "Sender: " & replace( replace(mailSenderSusc,"<","&lt;"), ">", "&gt;") &_
				    "<br>Subject: " & mailSubject &_
				    "<br>Recipient: " & mailRecipient &_
				    "<br>Body: " & mailBody
			
				    SendCrashMail "Error al enviar correo", "Error al enviar correo electronico:<br><strong>"&_
				    Err.Description & "</strong><br><br>" & paramStr

			    end if
			
                
                SQLExecute "EXEC UpdateEventos_Envio " & RPC_Value & ", '" & vCorreo & "'"

			    'end error handling
			    response.Write("ok")
            end if

            if proceso = "fin" then
                response.write("endprocess")
            end if
			
			' cleanup of variables
			Set mail_iMsg = Nothing
			Set mail_iConf = Nothing
			Set mail_Flds = Nothing
			set mail_CC = Nothing



    case "ordenamientoFotos":
        RPC_Value = ConvertToSQLNumericValue(request.Form("idArticle"), false)
        RPC_Value2 = ConvertToSQLStringValue(request.Form("ordenamiento"))

        if RPC_Value <> "NULL" then

			Set rs = Server.CreateObject("ADODB.Recordset")
			SQLString = "UpdateArticlePictureOrden @id_article=" & RPC_Value & ", @orden=" & RPC_Value2

			rs.Open SQLString, oCNDB, 3, 3

        	    While NOT rs.EOF
                        response.Write(rs("Mensaje"))
                    rs.moveNext()
                Wend

            if NOT rs.BOF Then rs.movefirst 
            rs.close

        end if
	
	case "GetListCities":
		RPC_IDItem = ConvertToSQLNumericValue2(RequestServerValue("id",-1), false, -1)
		WriteXMLStructuredData "exec GetListCities " & RPC_IDItem, "id_city,name"	

	case "GetListCitiesArch":
		RPC_IDItem = ConvertToSQLNumericValue2(RequestServerValue("id",-1), false, -1)
		WriteXMLStructuredData "exec GetListCitiesArch " & RPC_IDItem, "id_city,name"

    case "GetListStatesArchitect_Proj":
		RPC_IDItem = ConvertToSQLNumericValue2(RequestServerValue("id",-1), false, -1)
		WriteXMLStructuredData "exec GetListStatesArchitect_Proj " & RPC_IDItem, "id_state,name"

	case "GetListCitiesArch_Proj":
		RPC_IDItem = ConvertToSQLNumericValue2(RequestServerValue("id",0), false, 0)
		WriteXMLStructuredData "exec GetListCitiesArch_Proj " & RPC_IDItem, "id_city,name"

	case "GetListAreaTypes":
		RPC_IDItem = ConvertToSQLNumericValue2(RequestServerValue("id",-1), false, -1)
		WriteXMLStructuredData "exec GetListAreaTypes " & RPC_IDItem, "id_areas_type,areaType"

    case "GetListAreaTypes_Proj":
		RPC_IDItem = ConvertToSQLNumericValue2(RequestServerValue("id",-1), false, -1)
		WriteXMLStructuredData "exec GetListAreaTypes_Proj " & RPC_IDItem, "id_areas_type,areaType"

	case "GetBusinessList":
		RPC_IDItem = ConvertToSQLNumericValue2(RequestServerValue("id",-1), false, -1)
		WriteXMLStructuredData "exec GetBusinessList @idClient=" & RPC_IDItem, "id_business,name"
	
	case "GetBusinessLineTypeList"
		WriteXMLStructuredData "exec GetBusinessLineTypeList " & RPC_IDItem, "id_businesslinetype,name"
	
	case "GetBusinessLineList":
		RPC_IDItem = ConvertToSQLNumericValue2(RequestServerValue("id",-1), false, -1)
		WriteXMLStructuredData "exec GetBusinessLineList @idBusinesslinetype=" & RPC_IDItem, "id_businessline,name"
	
	case "CheckNewEmail":
		RPC_Value = ConvertToSQLStringValue(RequestServerValue("value",""))
		if RPC_Value <> "" Then
			RPC_IDItem = GetDBValue2("member_detail", "id_member", "email = " & RPC_Value & " AND magazine = 1 ")
			if RPC_IDItem = -1 Then response.write "ok" else response.write RPC_IDItem end if
		else
			response.write "nok"
		end if

	case "CheckNewArqEmail":
		RPC_Value = ConvertToSQLStringValue(RequestServerValue("value",""))
		if RPC_Value <> "" Then
			RPC_IDItem = GetDBValue2("architect", "id_architect", "email = " & RPC_Value)
			if RPC_IDItem = -1 Then response.write "ok" else response.write RPC_IDItem end if
		else
			response.write "nok"
		end if
	
    case "Codigo":
		RPC_Value = ConvertToSQLStringValue(RequestServerValue("value",""))        
		if RPC_Value <> "" Then
			RPC_IDItem = GetDBValue2("stilo_codigoPromocional", "id_codigo", "codigoPromocional = " & RPC_Value & "  and fechaAsignacion is null and id_member is null")
			if RPC_IDItem = -1 Then response.write "ok" else response.write RPC_IDItem end if
		else
			response.write "ok"
		end if                   


	case "CheckNewArqAccount":
		RPC_Value = ConvertToSQLStringValue(RequestServerValue("value",""))
		if RPC_Value <> "" Then
			RPC_IDItem = GetDBValue2("architect", "id_architect", "cuenta = " & RPC_Value)
			if RPC_IDItem = -1 Then response.write "ok" else response.write RPC_IDItem end if
		else
			response.write "nok"
		end if

	case "CheckNewMemberEmail":
		RPC_Value = ConvertToSQLStringValue(RequestServerValue("value",""))
		if RPC_Value <> "" Then
			RPC_IDItem = GetDBValue2("stilo_member", "id_member", "email = " & RPC_Value)
			if RPC_IDItem = -1 Then response.write "ok" else response.write RPC_IDItem end if
		else
			response.write "nok"
		end if
	
	case "CheckNewMemberAccount":
		RPC_Value = ConvertToSQLStringValue(RequestServerValue("value",""))
		if RPC_Value <> "" Then
			RPC_IDItem = GetDBValue2("stilo_member", "id_member", "cuenta = " & RPC_Value)
			if RPC_IDItem = -1 Then response.write "ok" else response.write RPC_IDItem end if
		else
			response.write "nok"
		end if

	case "CheckNewEmailNewsletter":
		RPC_Value = ConvertToSQLStringValue(RequestServerValue("value",""))
		if RPC_Value <> "" Then
			RPC_IDItem = GetDBValue2("member_detail", "id_member", "email = " & RPC_Value & " AND newsletter = 1 ")
			if RPC_IDItem = -1 Then response.write "ok" else response.write RPC_IDItem end if
		else
			response.write "nok"
		end if
	
	case "UpdateBusinessStars":
		RPC_IDItem = RequestServerValue("id",-1)
		RPC_Value = RequestServerValue("value",-1)
		SQLExecute "EXEC UpdateBusinessStars @idBusiness=" & RPC_IDItem & ", @starCount=" & RPC_Value

	case "UpdateFacebookLikes":
		RPC_IDItem = RequestServerValue("id",-1)
		RPC_Value = RequestServerValue("value",-1)
		SQLExecute "EXEC UpdateProjectLikes @id_project=" & RPC_IDItem & ", @count=" & RPC_Value
		
	case "DeleteObject":
		RPC_IDItem = ConvertToSQLNumericValue(request.querystring("value"), false)
		
		select case request("object")
			
			case "menu":
				SQLExecute "DeleteMenuItem " & RPC_IDItem
			
			case "menutop":
				SQLExecute "UPDATE menu_top SET item_order = item_order - 1 WHERE logicaldeletion IS NULL "&_
				" AND item_order > " & GetDBValue("menu_top", "item_order", "id_menu_top", RPC_IDItem)
				SQLExecute "UPDATE menu_top SET logicaldeletion = getdate() WHERE id_menu_top = " & RPC_IDItem
			
			case "module":
				SQLExecute "UPDATE module SET logicaldeletion = getdate() WHERE id_module = " & RPC_IDItem
			
			case "article":
				SQLExecute "UPDATE article SET logicaldeletion = getdate() WHERE id_article = " & RPC_IDItem

			case "project":
				SQLExecute "UPDATE project SET logicaldeletion = getdate() WHERE id_project = " & RPC_IDItem

			case "projectAutorization":
				SQLExecute "UPDATE project SET datePublished = getdate(), published=1 WHERE id_project = " & RPC_IDItem

			case "architect":
				SQLExecute "UPDATE architect SET logicaldeletion = getdate() WHERE id_architect = " & RPC_IDItem

			case "staff":
				SQLExecute "UPDATE staff SET logicaldeletion = getdate() WHERE id_staff = " & RPC_IDItem

			case "architectAutorization":
				SQLExecute "UPDATE architect SET activationDate = getdate() WHERE id_architect = " & RPC_IDItem
			
			case "client":
				SQLExecute "UPDATE branch	SET logicaldeletion = getdate() WHERE id_client = " & RPC_IDItem
				SQLExecute "UPDATE business SET logicaldeletion = getdate() WHERE id_client = " & RPC_IDItem
				SQLExecute "UPDATE client   SET logicaldeletion = getdate() WHERE id_client = " & RPC_IDItem
				
			case "business":
				SQLExecute "UPDATE business	SET logicaldeletion = getdate() WHERE id_business = " & RPC_IDItem
				SQLExecute "UPDATE branch	SET logicaldeletion = getdate() WHERE id_business = " & RPC_IDItem
			
			case "branch":
				SQLExecute "UPDATE branch SET logicaldeletion = getdate() WHERE id_branch = " & RPC_IDItem
			
			case "sliderItem":
				SQLExecute "UPDATE slider SET logicaldeletion = getdate() WHERE id_slider = " & RPC_IDItem
				
				'-- reorder slider items
				SQLExecute "UPDATE slider SET item_order = item_order - 1 WHERE logicaldeletion IS NULL "&_
				" AND item_order >= (SELECT item_order FROM slider WHERE id_slider = "&RPC_IDItem&") "

			case "newsletter":
				SQLExecute "UPDATE newsletter SET logicaldeletion = getdate() WHERE id_newsletter = " & RPC_IDItem

			case "distributionList":
				SQLExecute "UPDATE distributionList SET logicaldeletion = getdate() WHERE id_distributionlist = " & RPC_IDItem

				
		end select
		
		response.Write(RPC_IDItem)
	
	'-- save click on banner
	case "SaveBannerClick":
		call SQLExecute("SaveBannerClick @idBanner = " & RequestServerValue("id", "-1"))

	case "PasswordRecoveryPortafolios":
		RPC_Value = ConvertToSQLStringValue(request.QueryString("cuenta"))
		
		if RPC_Value <> "NULL" then
		
			Set RPC_dbObject = Server.CreateObject("ADODB.Recordset")
			SQLString = "GetUserPortafoliosAccess @cuenta=" & RPC_Value
			RPC_dbObject.Open SQLString, oCNDB, 3, 3
		
			select case CountDBRecords(RPC_dbObject)
				case 1:
					if RPC_dbObject("password") = "" Then
						'no cuenta con acceso
						Response.Write "El email ingresado no tiene acceso al sistema."
					else
						'correcto, enviar correo
						session("id_user")	= RPC_dbObject("id_user")
						session("name")		= RPC_dbObject("name")
                        session("cuenta")	= RPC_dbObject("cuenta")
						session("email")	= RPC_dbObject("email")
						session("password")	= RPC_dbObject("password")
						
						Server.Execute "RPC_SendMail.asp"
						
						Session.Contents.Remove("id_user")
						Session.Contents.Remove("name")
						Session.Contents.Remove("email")
                        Session.Contents.Remove("cuenta")
						Session.Contents.Remove("password")
						
						Response.Write "Tu cuenta y password han sido enviados al email registrado."
					end if
					
				case 0:
					'no encontrado
					Response.Write "El email ingresado no esta registrado, verifique e intente nuevamente."
				case else
					'mail compartido
					Response.Write "Cuenta duplicada"
			end select
		
			RPC_dbObject.close()
		
		end if

	
	case "SendUserPassword":
		RPC_Value = ConvertToSQLStringValue(request.QueryString("email"))
		
		if RPC_Value <> "NULL" then
		
			Set RPC_dbObject = Server.CreateObject("ADODB.Recordset")
			SQLString = "GetUserSystemAccess @email=" & RPC_Value
			RPC_dbObject.Open SQLString, oCNDB, 3, 3
		
			select case CountDBRecords(RPC_dbObject)
				case 1:
					if RPC_dbObject("password") = "" Then
						'no cuenta con acceso
						Response.Write "noAccess"
					else
						'correcto, enviar correo
						session("id_user")	= RPC_dbObject("id_user")
						session("name")		= RPC_dbObject("name")
						session("email")	= RPC_dbObject("email")
						session("password")	= RPC_dbObject("password")
						
						Server.Execute "RPC_SendMail.asp"
						
						Session.Contents.Remove("id_user")
						Session.Contents.Remove("name")
						Session.Contents.Remove("email")
						Session.Contents.Remove("password")
						
						Response.Write "ok"
					end if
					
				case 0:
					'no encontrado
					Response.Write "notFound"
				case else
					'mail compartido
					Response.Write "multipleMail"
			end select
		
			RPC_dbObject.close()
		
		end if
	
	' -- DEFAULT
	case else
		response.ContentType = "text/html; charset=iso-8859-1"
		response.write "Servicio activo"
		response.end

end select



' ---- Metodos auxiliares

sub WriteItemData(RPC_sqlString, columnsArray)
	response.ContentType = "text/html; charset=iso-8859-1"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open RPC_sqlString, oCNDB, 3, 3
	
	if rs.recordCount = 0 Then
		response.write "0"
	else
	
		cols = split(columnsArray,",")
		
		if NOT rs.EOF Then
			strResponse =""
			
			for each item in cols
				RPC_ItemValue = rs(item)
				
				if RPC_ItemValue <> "" Then
					RPC_ItemValue = replace(RPC_ItemValue, "|", "-")
				end if
				
				strResponse = strResponse & RPC_ItemValue & "|"
			next
			
			' -- remover ultimo separador
			if right(strResponse,1) = "|" Then
				strResponse = left(strResponse, len(strResponse)-1)
			end if
			
			' -- imprimir
			response.write strResponse
			
		end if
	end if
	
	rs.close()
	set rs = nothing
end sub
	
	
sub WriteListResponse(RPC_sqlString, columnsArray, extraString)
	response.ContentType = "text/html; charset=iso-8859-1"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open RPC_sqlString, oCNDB, 3, 3
	
	if rs.recordCount = 0 Then
		response.write "0"
	else
	
		cols = split(columnsArray,",")
		
		while NOT rs.EOF
			strResponse =""
			
			for each item in cols
				RPC_ItemValue = rs(item)
				
				if RPC_ItemValue <> "" Then
					RPC_ItemValue = replace(RPC_ItemValue, "|", "-")
				end if
				
				strResponse = strResponse & RPC_ItemValue & "|"
			next
			
			' -- remover ultimo separador
			if right(strResponse,1) = "|" Then
				strResponse = left(strResponse, len(strResponse)-1)
			end if
			
			' -- valor extra
			if extraString <> "" Then
				strResponse = strResponse & "|" & extraString
			end if
			
			' -- imprimir con salto de linea
			response.write strResponse & "~"
			
			rs.movenext()
		wend
	end if
	
	rs.close()
	set rs = nothing
end sub


function GetEvaluatorTypeListQueryString()
	Dim GETLQS_sqlString
	GETLQS_sqlString = "EXEC GetEvaluatorTypeList "
	GETLQS_sqlString	= GETLQS_sqlString & " @idCompany = " & ConvertToSQLNumericValue2(request.QueryString("company"), false, -1)
	GETLQS_sqlString	= GETLQS_sqlString & ", @idDepartment = " & ConvertToSQLNumericValue2(request.QueryString("department"), false, -1)
	GETLQS_sqlString	= GETLQS_sqlString & ", @idArea = " & ConvertToSQLNumericValue2(request.QueryString("area"), false, -1)
	GETLQS_sqlString	= GETLQS_sqlString & ", @idLevel = " & ConvertToSQLNumericValue2(request.QueryString("level"), false, -1)
	GETLQS_sqlString	= GETLQS_sqlString & ", @idLocation = " & ConvertToSQLNumericValue2(request.QueryString("location"), false, -1)
	
	'parametros evaluacion global
	if request.queryString("datePeriod") <> "" Then
		tempValue = split(request.queryString("datePeriod"), "|")
		GETLQS_sqlString		= GETLQS_sqlString & ", @dateStart = " & ConvertToSQLDateFormat(tempValue(0))
		GETLQS_sqlString		= GETLQS_sqlString & ", @dateEnd = " & ConvertToSQLDateFormat(tempValue(1))
	end if
	
	GetEvaluatorTypeListQueryString = GETLQS_sqlString

end function

sub ValidateUniqueValue(tableName, columnIndex, columnKey)
	itemValue 	= ConvertToSQLStringValue(request.queryString("value"))
	itemID		= ConvertToSQLNumericValue2(request.queryString("id"), false, -1)
	companyID	= ConvertToSQLNumericValue2(request.queryString("id_company"), false, -1)

	if request.queryString("type") = "competence" then
	itemValue = normalizeString(itemValue)
	end if
	
	if itemValue <> "NULL" then
	select case request.queryString("type")
		case "competence":
			RPC_itemID = GetDBValue4(tableName, columnKey, itemValue, columnIndex, itemID)
		
		case "position":
			RPC_itemID = GetDBValue2(tableName & " P INNER JOIN area A ON A.id_area = P.id_area", columnIndex, " P.logicaldeletion IS NULL AND P."&columnKey&" = "&itemValue&" AND P."&columnIndex&" <> " & itemID & " AND A.id_company = " & companyID)
		
		case "positionCode":
			RPC_itemID = GetDBValue2(tableName & " P INNER JOIN area A ON A.id_area = P.id_area", columnIndex, " P.logicaldeletion IS NULL AND P."&columnKey&" = "&itemValue&" AND P."&columnIndex&" <> " & itemID & " AND A.id_company = " & companyID)
		
		case "number":
			RPC_itemID = GetDBValue2(tableName, columnIndex, " logicaldeletion IS NULL AND "&columnKey&" = "&itemValue&" AND "&columnIndex&" <> " & itemID & " AND id_company = " & companyID)
	
	end select
		
		if RPC_itemID = -1 then
			Response.Write "OK"
		else
			Response.Write RPC_itemID
		end if
	else
		Response.Write "NOK"
	end if
end sub


	
set strResponse = nothing
set RPC_dbObject = nothing
set RPC_IDItem 	= nothing
set RPC_Value 	= nothing
set SQLString = nothing
set rs = nothing
set RPC_pageSize = nothing
set RPC_startIndex = nothing

%>
<!--#include virtual="/serverside/functions_end.asp" -->
<!--#include virtual="/serverside/database_close.asp" -->