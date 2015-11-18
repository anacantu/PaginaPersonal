<%@ Language=VBScript %>

<%
Response.addHeader "pragma", "no-cache"
Response.CacheControl = "Private"
Response.Expires = -1
Response.Buffer = "false"
 
Dim user, noaccess, validate
user = Cstr(request.cookies("validUserStilo"))
AppIDUser = user
noaccess = Cstr(request.cookies("noaccessStilo"))

response.write("x: " & user)

if user = "" or noaccess="1" then
    response.Redirect("/portafolio/login.asp")
end if

%>