<%
  recaptcha_challenge_field  = Request("recaptcha_challenge_field")
  recaptcha_response_field   = Request("recaptcha_response_field")
  recaptcha_public_key       = "6Le1y8cSAAAAAEZwhJfNGYx8bxwYdS7ScdZKIXM-" ' your public key
  recaptcha_private_key      = "6Le1y8cSAAAAAHyGmxjfdhfnxd77EpYfXobieJg8" ' your private key

  ' returns the HTML for the widget
  function recaptcha_challenge_writer()
  recaptcha_challenge_writer = _
  "<script type=""text/javascript"">" & _
  "var RecaptchaOptions = {" & _
  "   theme : 'clean'," & _
  "   lang : 'es'," & _
  "   tabindex : 0" & _
  "};" & _
  "</script>" & _
  "<script type=""text/javascript"" src=""http://www.google.com/recaptcha/api/challenge?k=" & recaptcha_public_key & """></script>" & _
  "<noscript>" & _
    "<iframe src=""http://www.google.com/recaptcha/api/noscript?k=" & recaptcha_public_key & """ frameborder=""1""></iframe><>" & _
      "<textarea name=""recaptcha_challenge_field"" rows=""3"" cols=""40""></textarea>" & _
      "<input type=""hidden"" name=""recaptcha_response_field""value=""manual_challenge"">" & _
  "</noscript>"

  end function

  ' returns "" if correct, otherwise it returns the error response
  function recaptcha_confirm(rechallenge,reresponse)
  Dim VarString
  VarString = _
          "privatekey=" & recaptcha_private_key & _
          "&remoteip=" & Request.ServerVariables("REMOTE_ADDR") & _
          "&challenge=" & rechallenge & _
          "&response=" & reresponse

  Dim objXmlHttp
  Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
  objXmlHttp.open "POST", "http://www.google.com/recaptcha/api/verify", False
  objXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  objXmlHttp.send VarString
  Dim ResponseString
  ResponseString = split(objXmlHttp.responseText, vblf)
  Set objXmlHttp = Nothing

  if ResponseString(0) = "true" then
    'They answered correctly
     recaptcha_confirm = ""
  else
    'They answered incorrectly
     recaptcha_confirm = ResponseString(1)
  end if

  end function

  server_response = ""
  newCaptcha = True
  if (recaptcha_challenge_field <> "" or recaptcha_response_field <> "") then
    server_response = recaptcha_confirm(recaptcha_challenge_field, recaptcha_response_field)
    newCaptcha = False
  end if

  %>
