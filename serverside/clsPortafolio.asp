<%

'
'	NOTE: 	data_storage/database.asp 
'			serverside/functions.asp	MUST BE LOADED
'
' Demo:
'	Set portafolioItem = new clsPortafolio
'	call portafolioItem.LoadData(IDArchitect)
'

Class clsPortafolio
' ------------------------------------------------------------------------------
	
	Private p_id
	Private p_business
	Private p_name
	Private p_lastname1
	Private p_lastname2
	Private p_email
	Private p_website
	Private p_idstate
    Private p_idcity
	Private	p_dateUp
    Private p_tipoDespacho
    Private p_arquitectos
    Private p_city
    Private p_state
    Private p_estudiante
    Private p_seudonimo
    Private p_idcountry
    Private p_country

	Private DBConn
	
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
	End Sub
' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
		if NOT IsNumeric(paramID) OR paramID = "" Then exit sub
		
		p_id = paramID
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		SQLString = "clsPortafolioQuery @idArquitect = " & p_id
		DBConn.Open SQLString, oCNDB, 3, 3
		
		if NOT DBConn.EOF then
			
		    p_id            = safeCdbl(DBConn("id_architect"))
		    p_business      = DBConn("business")
		    p_name          = DBConn("name")
		    p_lastname1     = DBConn("lastname1")
		    p_lastname2     = DBConn("lastname2")
		    p_email         = DBConn("email")
		    p_website       = DBConn("website")
		    p_idstate       = safeCdbl(DBConn("id_state"))
            p_idcity        = safeCdbl(DBConn("id_city"))
			p_dateUp		= DBConn("dateUp")
            p_tipoDespacho  = DBConn("tipoDespacho")
            p_arquitectos   = DBConn("arquitectos")
            p_city          = DBConn("city")
            p_state         = DBConn("state")
            p_estudiante    = DBConn("flagEstudiante")
            p_seudonimo     = DBConn("seudonimo")
            p_idcountry     = DBConn("id_country")
            p_country       = DBConn("country")
			
		else
			p_id_module	= -1
		end if
		
		DBConn.close
		
	End Sub	
' ------------------------------------------------------------------------------
	Public Sub PrintRelatedProjectCover(ByVal avoidArticleID)
		
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
	Public Property Get ID()
		ID = p_id
	End Property
' ------------------------------------------------------------------------------
	Public Property Get business()
		business	 = p_business
	End Property
' ------------------------------------------------------------------------------
	Public Property Get estudiante()
		estudiante	 = p_estudiante
	End Property
' ------------------------------------------------------------------------------
	Public Property Get seudonimo()
		seudonimo	 = p_seudonimo
	End Property
' ------------------------------------------------------------------------------
	Public Property Get tipoDespacho()
		tipoDespacho	 = p_tipoDespacho
	End Property    
' ------------------------------------------------------------------------------
	Public Property Get name()
		name	 = p_name
	End Property
' ------------------------------------------------------------------------------
	Public Property Get lastName1()
		lastName1	 = p_lastname1
	End Property
' ------------------------------------------------------------------------------
	Public Property Get lastName2()
		lastName2	 = p_lastname2
	End Property
' ------------------------------------------------------------------------------
	Public Property Get email()
		email	 = p_email
	End Property
' ------------------------------------------------------------------------------
	Public Property Get webSite()
		webSite	 = p_website
	End Property
' ------------------------------------------------------------------------------
	Public Property Get idState()
		idState	 = p_idstate
	End Property
' ------------------------------------------------------------------------------
	Public Property Get idCity()
		idCity	 = p_idcity
	End Property
' ------------------------------------------------------------------------------
	Public Property Get city()
		city	 = p_city
	End Property
' ------------------------------------------------------------------------------
	Public Property Get state()
		state	 = p_state
	End Property
' ------------------------------------------------------------------------------
	Public Property Get country()
		country	 = p_country
	End Property
' ------------------------------------------------------------------------------
	Public Property Get id_country()
		id_country	 = p_idcountry
	End Property
' ------------------------------------------------------------------------------
	Public Property Get DateUP()
		DateUP	 = p_dateUp
	End Property
' ------------------------------------------------------------------------------
	Public Property Get arquitectos()
		arquitectos	 = p_arquitectos
	End Property
' ------------------------------------------------------------------------------



End Class
' ------------------------------------------------------------------------------
%>