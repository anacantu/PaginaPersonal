<%

            dim mailSenderSusc
            mailSenderSusc = "Suscripciones Stilo <hola@stilo.com.mx>"

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
			
			'mail_CC = ccMailAddress
            mailSubject = "Nueva suscripción a la Revista Stilo."
            mailRecipient = Replace(email, "'", "")
            mailBcc = "rubio@stilo.com.mx;zey@stilo.com.mx"
            mailBody = "Hola <strong>" & Replace(name, "'", "") & "</strong><br /><br />Felicidades, te has registrado con &eacute;xito a la suscripci&oacute;n de la revista de Stilo.<br /><br />Para poder recibir los 5 ejemplares gratis, debes realizar el pago de: $ " & formatnumber(paqueteria, 2)
                    if paqueteria = 300 or paqueteria = 400 then
                            mailBody = mailBody & " (Servicio Postal Mexicano). <br /><br />"
                    end if


                    if paqueteria = 450 or paqueteria = 550 then
                            mailBody = mailBody & " (Paquetería privada). <br /><br />"
                    end if
                            mailBody = mailBody & "Haz tu dep&oacute;sito bancario en la siguiente cuenta:<br />" & _
                            "<strong>Banco: </strong>Banorte<br />" & _
                            "<strong>Beneficiario: </strong>Zeyttel Sainz Kim<br />" & _
                            "<strong>No. cuenta: </strong>0838442209<br />" & _
                            "<strong>CLABE: </strong>072 580 008384422090<br />" & _
                            "<br />Una vez realizado el pago, env&iacute;anos tu comprobante al correo: hola@stilo.com.mx" & _
                            "<br /><br />" & _ 
                            "<strong>Te confirmamos la direcci&oacute;n de env&iacute;o:</strong><br />" & _ 
                            "<strong>Atenci&oacute;n a: </strong>" & Replace(name, "'", "") & " " & Replace(lastname1, "'", "")  & " " & Replace(lastname2, "'", "") & _ 
                            "<br />" & Replace(street , "'", "") & " # " & Replace(number, "'", "") & _
                            "<br />" & Replace(district , "'", "") & _
                            "<br />" & ciudadDesc & ", " & estadoDesc & _
                            "<br />" & Replace(postCode, "'", "") & " M&eacute;xico" & _
                            "<br />Tel&eacute;fono: " & Replace(phone , "'", "") & _
                            "<br /><br /><strong>&iexcl;Gracias!<br /><br />Stilo Action Team.</strong>"
            
			'error handling
			On Error Resume Next
			
			' apply the settings to the message
			With mail_iMsg
				Set .Configuration = mail_iConf
                .To     = mailRecipient
				.Bcc	= mailBcc
				.Cc		= mail_CC
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
            %>
                <script language="javascript" type="text/javascript">
                    alert('Ocurrio un error al intentar enviar los correos, notificar al administrador.');
                </script>                
            <%

			end if
			
			'end error handling
			On Error GoTo 0
			
			' cleanup of variables
			Set mail_iMsg = Nothing
			Set mail_iConf = Nothing
			Set mail_Flds = Nothing
			set mail_CC = Nothing

%>