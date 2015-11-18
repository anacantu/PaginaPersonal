<%
	Dim oCNDB
	Set oCNDB = Server.CreateObject("ADODB.Connection")
	oCNDB.Open Application("connString")
	
%>