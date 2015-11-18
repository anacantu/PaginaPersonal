<%
' ------------------------------------------------------------------------------
'	Author:		Carlos Trevino @ Pulso Vital Consulting Group
'	Email:		hobbes313@hotmail.com
'	URL:		
'	Date:		Jul 12, 2009
' ------------------------------------------------------------------------------
'
'	NOTE: 	data_storage/database.asp 
'			serverside/functions.asp	MUST BE LOADED
'
' Demo:
'	Set articleItem = new clsArticle
'	call articleItem.LoadData(IDArticle)
'
'

Class clsModule
' ------------------------------------------------------------------------------
	
	Private p_id_module
	Private p_name
	Private p_dateUp
	Private p_active
	
	Private p_id_menu_parent
	Private p_menu_parent_name
	Private p_id_menu
	Private p_menu_name
			
	Private DBConn
	
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
	End Sub
' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
		if NOT IsNumeric(paramID) OR paramID = "" Then exit sub
		
		p_id_module = paramID
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		SQLString = "clsModuleQuery @idModule = " & p_id_module
		DBConn.Open SQLString, oCNDB, 3, 3
		
		if NOT DBConn.EOF then
			
			p_id_module		= safeCdbl(DBConn("id_module"))
			p_name			= DBConn("name")
			p_dateUp		= DBConn("dateUp")
			p_active		= CBool(DBConn("active"))
			
		else
			p_id_module	= -1
		end if
		
		DBConn.close
		
		'-- Get menu item names
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		SQLString = "GetModuleMenuData @idModule = " & p_id_module
		DBConn.Open SQLString, oCNDB, 3, 3
		if NOT DBConn.EOF then
			p_id_menu_parent	= safeCdbl(DBConn("id_menu_parent"))
			p_menu_parent_name	= DBConn("name_parent")
			p_id_menu			= safeCdbl(DBConn("id_menu"))
			p_menu_name			= DBConn("name")
		end if
		
		DBConn.close
		
		
	End Sub	
' ------------------------------------------------------------------------------
	Public Sub PrintRelatedArticlesCover(ByVal avoidArticleID)
		
		Dim articleItem, p_counter
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		p_sqlString = "SELECT TOP 2 "&_
		" id_article, id_module, moduleName, title, dateUp, subtitle, convert(varchar(1000),content) as contenido, datePublished, published "&_
		" FROM article_view "&_
		" WHERE id_module = " & p_id_module &_
		" /*filter*/ "&_
		" ORDER BY NEWID()"
		
		if avoidArticleID <> "" Then
			p_sqlString = replace(p_sqlString, "/*filter*/", "AND id_article <> " & safeCstr(avoidArticleID))
		end if
		
		DBConn.Open p_sqlString, oCNDB, 3, 3
		p_counter = 0
		
		While NOT DBConn.EOF
			p_counter = p_counter + 1
			
			set articleItem = New clsArticle
			call articleItem.LoadRecordData(DBConn, true, false)
			
			response.write "<div class=""article-ad"
			if p_counter=2 then response.write "-right"
			response.write """>"
			
			response.write "<a href=""/article.asp?id=" & articleItem.IDArticle & """>"&_
					" <img src="""& articleItem.MainThumbPictureFull &""" width=""160"" height=""100"">"&_
					" <h3 class=""aad-title"">"& articleItem.Title &"</h3>"&_
					" <span class=""aad-date"">"& formatDateString(articleItem.EnglishDatePublished,3) &"</span>"&_
					" <span class=""aad-text"">"& articleItem.CoverText &"</span>"&_
					" </a> </div>"
			
			DBConn.movenext
			
		Wend
		
		DBConn.close
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub Class_Terminate()
		
		Set DBConn = Nothing
		
	End Sub
' ------------------------------------------------------------------------------
	Public Property Get IDModule()
		IDModule = p_id_module
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Name()
		Name	 = p_name
	End Property
' ------------------------------------------------------------------------------
	Public Property Get DateUp()
		DateUp	 = p_dateUp
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Active()
		Active	 = p_active
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDMenu()
		Active	 = p_id_menu
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Menu()
		Menu	 = p_menu_name
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDParentMenu()
		IDParentMenu	 = p_id_menu_parent
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ParentMenu()
		ParentMenu	 = p_menu_parent_name
	End Property
' ------------------------------------------------------------------------------
	Public Property Get MenuTrace()
		if ParentMenu = "" AND Menu = "" Then
			MenuTrace = Name
		else
			if ParentMenu = "" Then
				MenuTrace = Menu
			else
				MenuTrace = ParentMenu & " / " & Menu
			end if
		end if
	End Property
' ------------------------------------------------------------------------------




End Class
' ------------------------------------------------------------------------------
%>