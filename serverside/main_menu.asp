<ul class="sidemenu" id="mainmenu">
	<li class="clearItem">&nbsp;</li>
<%
'id="current"

Dim mainMenuRSet, linkAndTxt
set mainMenuRSet = Server.CreateObject("ADODB.RecordSet")
mainMenuRSet.open "exec GetMainMenuList", oCNDB, 3, 3

While NOT mainMenuRSet.EOF
	subMenuStr = ""
	
	if IsNumeric(mainMenuRSet("id_parent")) Then
		subMenuStr = "class=""submenu"" style=""display:none"" idparent="""& mainMenuRSet("id_parent") &""""
	end if
	
	if safeCint(mainMenuRSet("child_count")) > 0 Then
		linkAndTxt = "<a onclick=""OpenSubMenu('"&mainMenuRSet("id_menu")&"')"">"& mainMenuRSet("name") &"</a>"
	else
		if mainMenuRSet("id_module") <> "" Then
			linkAndTxt = "<a href=""/module.asp?id="& mainMenuRSet("id_module") &""">"& mainMenuRSet("name") &"</a>"
		else
			linkAndTxt = "<a href="""& mainMenuRSet("linkName") &""">"& mainMenuRSet("name") &"</a>"
		end if
	end if
	
	response.write "<li "&subMenuStr&">"&linkAndTxt&"</li>" & vbcrlf
	
	mainMenuRSet.movenext()
Wend
mainMenuRSet.close()
set mainMenuRSet = Nothing
%>
								<li class="clearItem">&nbsp;</li>
							</ul>