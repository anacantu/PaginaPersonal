<%
'	NOTE: 	data_storage/database.asp 
'			serverside/functions.asp	MUST BE LOADED
'
' Demo:
'	Set projectItem = new clsProject
'	call projectItem.LoadData(IDProject)
'
'

Class clsProject
' ------------------------------------------------------------------------------

	Private p_id_project
	Private p_id_architect
    Private p_business
    Private p_tipoDespacho
	Private p_title
	Private p_location
    Private p_id_estate
    Private p_id_city
    Private p_area
    Private p_tipoProyecto
	Private p_mtsTerreno
	Private p_mtsConstruccion
	Private p_disenoArquitectonico
	Private p_arquitecto
	Private p_colaboradores
	Private p_content
	Private p_content_noHTML
	Private p_datePublished
	Private p_views
	Private p_tags
	Private p_grade
    Private p_estudiante
    Private p_seudonimo
    Private p_id_country
    Private p_country
    Private p_estate
    Private p_city
	
	'-- pictures
	Private p_project_pictures
	Private p_thumb_picture
	Private p_main_picture
	
	'-- video
	Private p_project_videos
	
	Private DBConn
	Private ContributorsObj
	
	Private p_contributors_text
	Private p_sqlString
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
		p_thumb_picture = null
		p_main_picture = null
		p_contributors_text = ""
		set p_project_pictures	= new clsProjectPictures
		set p_project_videos	= new clsProjectVideos
		
	End Sub
' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
		if NOT IsNumeric(paramID) Then exit sub
		
		p_id_project = paramID
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		p_sqlString = "clsProjectQuery @idProject= " & p_id_project
		DBConn.Open p_sqlString, oCNDB, 3, 3
		
		if NOT DBConn.EOF then
        			
	        p_id_project                = safeCdbl(DBConn("id_project"))
	        p_id_architect              = safeCdbl(DBConn("id_architect"))
            p_business                  = DBConn("business")
            p_tipoDespacho              = DBConn("tipoDespacho")
	        p_title                     = DBConn("title")
	        'p_location                 = DBConn("location")
            if rtrim(DBConn("city"))  = "" or IsNull(DBConn("city")) then
                p_location              = DBConn("estate") & ", " & DBConn("country")
            else
	            p_location              = DBConn("city") & ", " & DBConn("estate") & ". " & DBConn("country")
            end if
            p_id_estate                 = DBConn("id_estate")
            p_id_city                   = DBConn("id_city")
            p_area                      = DBConn("areas")
            p_tipoProyecto              = DBConn("areas_type")
	        p_mtsTerreno                = DBConn("mtsTerreno")
	        p_mtsConstruccion           = DBConn("mtsConstruccion")
	        p_disenoArquitectonico      = DBConn("disenoArquitectonico")
	        p_arquitecto                = DBConn("arquitecto")
	        p_colaboradores             = DBConn("colaboradores")
	        p_content                   = DBConn("content")
	        p_content_noHTML            = RemoveHTML(p_content)
	        p_datePublished             = SpanishDateFormat(DBConn("datePublished"))
	        p_views                     = safeCdbl(DBConn("views"))
	        p_tags                      = DBConn("tags")
	        p_grade                     = safeCint(DBConn("grade"))
            p_estudiante                = DBConn("flagEstudiante")
            p_seudonimo                 = DBConn("seudonimo")
            p_id_country                = DBConn("id_country")
            p_country                   = DBConn("country")
            p_estate                    = DBConn("estate")
            p_city                      = DBConn("city")
			
			'--Load Project Pictures
			p_project_pictures.LoadData(p_id_project)
			
			'--Load Project Videos
			p_project_videos.LoadData(p_id_project)					
			
		else
			p_id_project	= -1
		end if
		
		DBConn.close
		
	End Sub	
' ------------------------------------------------------------------------------
	Public Sub LoadDataMaster(paramID)
		
		if NOT IsNumeric(paramID) Then exit sub
		
		p_id_project = paramID
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		p_sqlString = "clsProjectQueryMaster @idProject= " & p_id_project
		DBConn.Open p_sqlString, oCNDB, 3, 3
		
		if NOT DBConn.EOF then
        			
	        p_id_project                = safeCdbl(DBConn("id_project"))
	        p_id_architect              = safeCdbl(DBConn("id_architect"))
            p_business                  = DBConn("business")
            p_tipoDespacho              = DBConn("tipoDespacho")
	        p_title                     = DBConn("title")
	        'p_location                  = DBConn("location")
	        p_location                  = DBConn("city") & ", " & DBConn("estate")
            p_id_estate                 = DBConn("id_estate")
            p_id_city                   = DBConn("id_city")
            p_area                      = DBConn("areas")
            p_tipoProyecto              = DBConn("areas_type")
	        p_mtsTerreno                = DBConn("mtsTerreno")
	        p_mtsConstruccion           = DBConn("mtsConstruccion")
	        p_disenoArquitectonico      = DBConn("disenoArquitectonico")
	        p_arquitecto                = DBConn("arquitecto")
	        p_colaboradores             = DBConn("colaboradores")
	        p_content                   = DBConn("content")
	        p_content_noHTML            = RemoveHTML(p_content)
	        p_datePublished             = SpanishDateFormat(DBConn("datePublished"))
	        p_views                     = safeCdbl(DBConn("views"))
	        p_tags                      = DBConn("tags")
	        p_grade                     = safeCint(DBConn("grade"))
            p_estudiante                = DBConn("flagEstudiante")
            p_seudonimo                 = DBConn("seudonimo")
            p_id_country                = DBConn("id_country")
            p_country                   = DBConn("country")
            p_estate                    = DBConn("estate")
            p_city                      = DBConn("city")
			
			'--Load Project Pictures
			p_project_pictures.LoadData(p_id_project)
			
			'--Load Project Videos
			p_project_videos.LoadData(p_id_project)					
			
		else
			p_id_project	= -1
		end if
		
		DBConn.close
		
	End Sub	
' ------------------------------------------------------------------------------
	Public Sub PrintRelatedProjectsCover(ByVal avoidProjectID)
		
		Dim projectItem, p_counter
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		p_sqlString = "SELECT TOP 6 "&_
		"id_project ,p.id_architect ,business ,case tipoDespacho"&_
        "	when 1 then 'Arquitectura'      "&_
		"	when 2 then 'Diseño Industrial'"&_
		"	when 4 then 'Interiorismo'"&_
		"	when 3 then 'Arquitectura y Diseño Industrial'"&_
		"	when 5 then 'Arquitectura e Interiorismo'"&_
		"	when 6 then 'Diseño Industrial e Interiorismo'"&_
		"	when 7 then 'Arquitectura, Diseño Industrial e Interiorismo'"&_
		" end as tipoDespacho"&_
		",title,location,mtsTerreno,mtsConstruccion,disenoArquitectonico,arquitecto,colaboradores,content"&_
		",datePublished,views,tags,grade"&_
        ",p.id_country, cn.name as country, p.estate, p.city"&_
        ",p.id_estate,e.name as estate,p.id_city,c.name as city,p.areas"&_
		", area = substring( ( SELECT ',' + aa.Area "&_
		"		FROM architect_areas aa "&_
		"		WHERE id_area IN ( SELECT convert(int,Value) FROM dbo.Split(Areas,'|') ) FOR XML path(''), elements "&_
		"		),2,500) "&_
		",p.areas_type "&_
		", areaType = substring( ( SELECT ',' + t.areaType "&_
		"		FROM architect_areaTypes t "&_
		"		WHERE id_areas_type IN ( SELECT convert(int,Value) FROM dbo.Split(areas_type,'|') ) FOR XML path(''), elements "&_
		"		),2,500) "&_
        ",a.flagEstudiante, a.seudonimo "&_
	    " FROM project p "&_
		"	inner join architect a on (p.id_architect = a.id_architect)"&_
        "	inner join country cn on (p.id_country = cn.id_country)"&_
		"	inner join [state] e on (p.id_estate = e.id_state)"&_
		"	inner join city c on (p.id_city = c.id_city)"&_
	    " WHERE p.id_architect = " & p_id_architect &_
	    " AND	p.logicaldeletion is null"&_
		" /*filter*/ "&_
		" ORDER BY NEWID()"
		
		if avoidProjectID <> "" Then
			p_sqlString = replace(p_sqlString, "/*filter*/", "AND id_project <> " & safeCstr(avoidProjectID))
		end if
		

		DBConn.Open p_sqlString, oCNDB, 3, 3
		p_counter = 0
		
		While NOT DBConn.EOF
			p_counter = p_counter + 1
			
			set projectItem = New clsProject
			call projectItem.LoadRecordData(DBConn, true, false)
						
            %>
     		<div class="AreaArt_pic6_<%=p_counter%>">
                <a href="/project.asp?id=<%= projectItem.IDProject %>">
					<img src="../<%= Replace(projectItem.MainThumbPictureFull, "_thumb", "_small") %>" width="100%" height="87" alt="" />
				</a>
            </div>
			<div class="AreaArt_pic6_<%=p_counter%>_title">
                <a style="font-size: 12px; font-family: Helvetica, Arial;" href="/project.asp?id=<%= projectItem.IDProject %>">
                    <h3 class="aad-title"><%= projectItem.Title %></h3>
				</a>
            </div>	            
        <%


			DBConn.movenext
			
		Wend
		response.Write "<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />"
		DBConn.close
		
	End Sub

' ------------------------------------------------------------------------------
	Public Sub LoadRecordData(lclRecordSet, loadPictures, loadVideos)
		
		call CleanValues()
		
		if isNull(lclRecordSet) Then exit sub
		if lclRecordSet.State = 0 Then exit sub
		if lclRecordSet.EOF Then exit sub

	        p_id_project                = safeCdbl(lclRecordSet("id_project"))
	        p_id_architect              = safeCdbl(lclRecordSet("id_architect"))
            p_business                  = lclRecordSet("business")
            p_tipoDespacho              = lclRecordSet("tipoDespacho")
	        p_title                     = lclRecordSet("title")
	        'p_location                  = lclRecordSet("location")
	        p_location                  = lclRecordSet("city") & ", " & lclRecordSet("estate")
            p_id_estate                 = lclRecordSet("id_estate")
            p_id_city                   = lclRecordSet("id_city")
            p_area                      = lclRecordSet("areas")
            p_tipoProyecto              = lclRecordSet("areas_Type")
	        p_mtsTerreno                = lclRecordSet("mtsTerreno")
	        p_mtsConstruccion           = lclRecordSet("mtsConstruccion")
	        p_disenoArquitectonico      = lclRecordSet("disenoArquitectonico")
	        p_arquitecto                = lclRecordSet("arquitecto")
	        p_colaboradores             = lclRecordSet("colaboradores")
	        p_content                   = lclRecordSet("content")
	        p_content_noHTML            = RemoveHTML(p_content)
	        p_datePublished             = SpanishDateFormat(lclRecordSet("datePublished"))
	        p_views                     = safeCdbl(lclRecordSet("views"))
	        p_tags                      = lclRecordSet("tags")
	        p_grade                     = safeCint(lclRecordSet("grade"))
            p_estudiante                = lclRecordSet("flagEstudiante")
            p_seudonimo                 = lclRecordSet("seudonimo")
            p_id_country                = lclRecordSet("id_country")
            p_country                   = lclRecordSet("country")
            p_estate                    = lclRecordSet("estate")
            p_city                      = lclRecordSet("city")

		if loadPictures Then
			'--Load Project Pictures
			p_project_pictures.LoadData(p_id_project)
		end if
		
		if loadVideos Then
			'--Load Project Videos
			p_project_videos.LoadData(p_id_project)
		end if
		
	End Sub	
' ------------------------------------------------------------------------------
	Private Sub CleanValues()
	        p_id_project                = -1
	        p_id_architect              = -1
            p_business                  = NULL
            p_tipoDespacho              = NULL
	        p_title                     = NULL
	        p_location                  = NULL
            p_id_estate                 = -1
            p_id_city                   = -1
            p_area                      = NULL
            p_tipoProyecto              = NULL
	        p_mtsTerreno                = NULL
	        p_mtsConstruccion           = NULL
	        p_disenoArquitectonico      = NULL
	        p_arquitecto                = NULL
	        p_colaboradores             = NULL
	        p_content                   = NULL
	        p_content_noHTML            = NULL
	        p_datePublished             = NULL
	        p_views                     = NULL
	        p_tags                      = NULL
	        p_grade                     = NULL
            p_estudiante                = NULL
            p_seudonimo                 = NULL
            p_id_country                = -1
            p_country                   = NULL
            p_estate                    = NULL
            p_city                      = NULL

	end sub	
' ------------------------------------------------------------------------------
	Private Sub Class_Terminate()
		
		Set DBConn = Nothing
		
	End Sub
' ------------------------------------------------------------------------------
	Public Property Get IDProject()
		IDProject = p_id_project
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDArchitect()
		IDArchitect = p_id_architect
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Business()
		Business = p_business
	End Property
' ------------------------------------------------------------------------------
	Public Property Get estudiante()
		estudiante = p_estudiante
	End Property
' ------------------------------------------------------------------------------
	Public Property Get seudonimo()
		seudonimo = p_seudonimo
	End Property
' ------------------------------------------------------------------------------
	Public Property Get TipoDespacho()
		TipoDespacho = p_tipoDespacho
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Title()
		Title = p_title
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Location()
		Location = p_location
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IdEstate()
		IdEstate = p_id_estate
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IdCountry()
		IdCountry = p_id_country
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Country()
		Country = p_country
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Estate()
		Estate = p_estate
	End Property
' ------------------------------------------------------------------------------
	Public Property Get City()
		City = p_city
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IdCity()
		IdCity = p_id_city
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Area()
		Area = p_area
	End Property
' ------------------------------------------------------------------------------
	Public Property Get TipoProyecto()
		TipoProyecto = p_tipoProyecto
	End Property
' ------------------------------------------------------------------------------
	Public Property Get MtsTerreno()
		MtsTerreno = p_mtsTerreno
	End Property
' ------------------------------------------------------------------------------
	Public Property Get MtsConstruccion()
		MtsConstruccion = p_mtsConstruccion
	End Property
' ------------------------------------------------------------------------------
	Public Property Get DisenoArquitectonico()
		DisenoArquitectonico = p_disenoArquitectonico
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Arquitecto()
		Arquitecto = p_arquitecto
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Colaboradores()
		Colaboradores = p_colaboradores
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Content()
		Content = p_content
	End Property
' ------------------------------------------------------------------------------
	Public Property Get CoverText()
		CoverText = trim(left(p_content_noHTML,85)) & "..."
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ShortCoverText()
		ShortCoverText = trim(left(p_content_noHTML,60)) & "..."
	End Property
' ------------------------------------------------------------------------------
	Public Property Get LongCoverText()
		LongCoverText = trim(left(p_content_noHTML,200)) & "..."
	End Property
' ------------------------------------------------------------------------------
	Public Property Get DatePublished()
		DatePublished = p_datePublished
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Views()
		Views = p_views
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Tags()
		Tags = p_tags
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Grade()
		Grade = p_grade
	End Property
' ------------------------------------------------------------------------------
	Public Property Get PicturesPath()
		PicturesPath	 = replace(Application("ProjectPicPath"), "\", "/")
	End Property
' ------------------------------------------------------------------------------
	Public Property Get VideoThumbsPath()
		VideosThumbPath	 = replace(Application("VideoProjThumbsPath"), "\", "/")
	End Property

' ------------------------------------------------------------------------------
' -  Pictures
' ------------------------------------------------------------------------------
	Public Property Get Pictures()
		set Pictures = p_project_pictures
	End Property

	Public Property Get MainThumbPictureFull()
		MainThumbPictureFull = Pictures.MainThumbPictureFull
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ThumbPicture()
		if IsNull(p_thumb_picture) Then
			p_thumb_picture = SearchPicture("thumb_" & IDProject)
		end if
		
		ThumbPicture = p_thumb_picture		
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ThumbPictureLink()
		ThumbPictureLink = replace(Application("ProjectPicPath"), "\", "/") & ThumbPicture
	End Property
' ------------------------------------------------------------------------------
	Public Property Get MainPicture()
		if IsNull(p_main_picture) Then
			p_main_picture = SearchPicture("main_" & IDProject)
		end if
		
		MainPicture = p_main_picture		
	End Property
' ------------------------------------------------------------------------------
	Public Property Get MainPictureLink()
		MainPictureLink = replace(Application("ProjectPicPath"), "\", "/") & MainPicture
	End Property
' ------------------------------------------------------------------------------


	Private function SearchPicture(ByVal picFileName)
		Dim FSO, foderPath, lclPicture
		Set FSO = server.CreateObject ("Scripting.FileSystemObject")
		foderPath = RealizePath(Application("ProjectPicPath")) & "\"
		
		lclPicture = ""
		
		if lclPicture = "" AND FSO.FileExists(foderPath & picFileName & ".jpg") Then lclPicture = picFileName & ".jpg"
		if lclPicture = "" AND FSO.FileExists(foderPath & picFileName & ".png") Then lclPicture = picFileName & ".png"
		if lclPicture = "" AND FSO.FileExists(foderPath & picFileName & ".gif") Then lclPicture = picFileName & ".gif"
		if lclPicture = "" AND FSO.FileExists(foderPath & picFileName & ".bmp") Then lclPicture = picFileName & ".bmp"
		
		SearchPicture = lclPicture
		
		Set FSO = Nothing
		Set foderPath = Nothing
	end function

	Public Sub SaveView()
		oCNDB.EXECUTE "IncrementProjectViews @id_project = " & p_id_project
	End Sub
	

' ------------------------------------------------------------------------------
' -  Videos
' ------------------------------------------------------------------------------
	Public Property Get Videos()
		set Videos = p_project_videos
	End Property


End Class
' ------------------------------------------------------------------------------







Class clsProjectPictures
	
	Private p_arr_pictures
	
	Private p_top_count
	Private p_id_main_picture
	Private p_pictures_count
	
	Private p_dbConnection
	Private p_sqlString
	Private p_fso
	Private p_folderPath
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
		Set p_dbConnection	= Server.CreateObject("ADODB.Recordset")
		Set p_fso			= server.CreateObject ("Scripting.FileSystemObject")
		p_folderPath		= RealizePath(Application("ProjectPicPath"))
		
		Redim p_arr_pictures(0)
		
		p_top_count = -1
		p_id_main_picture = -1
		p_pictures_count = -1
		
	End Sub
	
	Private Sub Class_Terminate()
	
		set p_dbConnection	= nothing
		Set p_fso			= nothing
		Set p_folderPath	= nothing
		
	End Sub

' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
		if NOT IsNumeric(paramID) Then exit sub
		
		p_sqlString = "GetProjectPictures @idProject = " & paramID
		p_dbConnection.Open p_sqlString, oCNDB, 3, 3
		
		p_top_count = 0
		p_pictures_count = 0
		
		While NOT p_dbConnection.EOF
			
			
			if cBool(p_dbConnection("is_main_pic")) Then p_id_main_picture = safeCdbl(p_dbConnection("id_picture"))
			
			Redim Preserve p_arr_pictures(p_pictures_count)
			set p_arr_pictures(p_pictures_count) = new clsPictureP
			
			call p_arr_pictures(p_pictures_count).LoadRecordData(p_dbConnection, p_dbConnection("id_project"))
			
			if p_arr_pictures(p_pictures_count).IsTop then p_top_count = p_top_count + 1
			
			p_dbConnection.movenext()
			p_pictures_count = p_pictures_count + 1
		Wend
		
		p_dbConnection.close
		
	End Sub	
' ------------------------------------------------------------------------------
	Public Property Get Count()
		Count = p_pictures_count
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Item(index)
		set Item = p_arr_pictures(index)
	End Property
' ------------------------------------------------------------------------------
	Public Property Get TopCount()
		TopCount = p_top_count
	End Property
' ------------------------------------------------------------------------------
	Public Property Get MainThumbPictureFull()
		if p_pictures_count > 0 Then
			MainThumbPictureFull = Application("ProjectPicPath") & GetThumbVersion(p_arr_pictures(0).FileName)
		else
			MainThumbPictureFull = Application("ProjectPicPath") & GetThumbVersion("")
		end if
	End Property
	
    '--Get large file version
	Private function GetLargeVersion(ByVal fileName)
		Dim GTV_fileExtension, GTV_fileName
		
		if fileName = "" Then
			GetThumbVersion = "noimage.png"
			exit function
		end if
		
		GTV_fileExtension = GetFileExtension(fileName)
		
		GTV_fileName = left(fileName, len(fileName) - 1 - len(GTV_fileExtension)) & "_large." & GTV_fileExtension
		
		if p_fso.FileExists(p_folderPath & GTV_fileName) Then
			GetThumbVersion = GTV_fileName
		else
			GetThumbVersion = "noimage.png"
		end if
		
	end function


	'--Get thumbnail file version
	Private function GetThumbVersion(ByVal fileName)
		Dim GTV_fileExtension, GTV_fileName
		
		if fileName = "" Then
			GetThumbVersion = "noimage.png"
			exit function
		end if
		
		GTV_fileExtension = GetFileExtension(fileName)
		
		GTV_fileName = left(fileName, len(fileName) - 1 - len(GTV_fileExtension)) & "_thumb." & GTV_fileExtension
		
		if p_fso.FileExists(p_folderPath & GTV_fileName) Then
			GetThumbVersion = GTV_fileName
		else
			GetThumbVersion = "noimage.png"
		end if
		
	end function
	
End Class




Class clsProjectVideos
	
	Private p_arr_videos
	
	Private p_videos_count
	
	Private p_dbConnection
	Private p_sqlString
	Private p_fso
	Private p_folderPath
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
		Set p_dbConnection	= Server.CreateObject("ADODB.Recordset")
		Set p_fso			= server.CreateObject ("Scripting.FileSystemObject")
		p_folderPath		= RealizePath(Application("ProjectPicPath"))
		
		Redim p_arr_videos(0)
		
		p_videos_count = -1
		
	End Sub
	
	Private Sub Class_Terminate()
	
		set p_dbConnection	= nothing
		Set p_fso			= nothing
		Set p_folderPath	= nothing
		
	End Sub

' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
		if NOT IsNumeric(paramID) Then exit sub
		
		p_sqlString = "GetProjectVideos @idProject = " & paramID
		p_dbConnection.Open p_sqlString, oCNDB, 3, 3
		
		p_videos_count = 0
		
		While NOT p_dbConnection.EOF
			
			
			Redim Preserve p_arr_videos(p_videos_count)
			set p_arr_videos(p_videos_count) = new clsVideoP
			
			call p_arr_videos(p_videos_count).LoadRecordData(p_dbConnection, p_dbConnection("id_project"))
			
			p_dbConnection.movenext()
			p_videos_count = p_videos_count + 1
		Wend
		
		p_dbConnection.close
		
	End Sub	
' ------------------------------------------------------------------------------
	Public Property Get Count()
		Count = p_videos_count
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Item(index)
		set Item = p_arr_videos(index)
	End Property
	
End Class




Class clsPictureP

	Private p_id_picture
	Private p_id_project
	Private p_filename
	Private p_dateUp
	Private p_author
	Private p_tags
	Private p_footer
	Private p_is_main
	Private p_is_top
	Private p_width
	Private p_height
	
	Private p_dbConnection
	Private p_sqlString
	Private p_fso
	Private p_folderPath
	Private p_linkPath
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
		Set p_dbConnection	= Server.CreateObject("ADODB.Recordset")
		Set p_fso			= server.CreateObject ("Scripting.FileSystemObject")
		p_folderPath		= RealizePath(Application("ProjectPicPath"))
		p_linkPath			= replace(Application("ProjectPicPath"), "\", "/")
		
		p_id_picture		= -1
		p_id_project		= -1
	End Sub
	
	Private Sub Class_Terminate()
	
		set p_dbConnection	= nothing
		Set p_fso			= nothing
		Set p_folderPath	= nothing
		
	End Sub

' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
	End Sub
' ------------------------------------------------------------------------------
	
	Public Sub LoadRecordData(lclRecordSet, IDProject)
		
		call CleanValues()
		
		if isNull(lclRecordSet) Then exit sub
		if lclRecordSet.State = 0 Then exit sub
		if lclRecordSet.EOF Then exit sub
		
		Dim picData
		
		if NOT IsNull(IDProject) Then p_id_project	= safeCdbl(IDProject)
		
		p_id_picture	= safeCdbl(lclRecordSet("id_picture"))
		p_filename		= lclRecordSet("filename")
		p_dateUp		= lclRecordSet("dateup")
		p_author		= lclRecordSet("author")
		p_tags			= lclRecordSet("tags")
		p_footer		= lclRecordSet("footer")
		p_is_main		= cBool(lclRecordSet("is_main_pic"))
		
		picData			= GetPictureData(p_filename)
		p_width			= picData(0)
		p_height		= picData(1)
		
		p_is_top		= 1 'p_width >= 550 AND p_width <= 580
		
		set picData = nothing
	End Sub		
' ------------------------------------------------------------------------------
	Private Sub CleanValues()
		p_id_picture	= -1
		p_id_project	= -1
		p_filename		= null
		p_dateUp		= null
		p_author		= null
		p_tags			= null
		p_footer		= null
		p_is_main		= null
		p_is_top		= null
		p_width			= null
		p_height		= null
	end sub
	
' ------------------------------------------------------------------------------
	Public Property Get FileName()
		FileName = p_filename
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Author()
		Author = p_author
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Footer()
		Footer = p_footer
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Link()
		Link = p_linkPath & p_filename
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Width()
		Width = p_width
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Height()
		Height = p_height
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IsTop()
		IsTop = p_is_top
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ThumbLink()
		ThumbLink = p_linkPath & GetThumbVersion(p_filename)
	End Property
	

' -- Get picture dimensions ----------------------------------------------------
	Private Function GetPictureData(ByVal fileName)
		
		Dim paramArr(1)
		
		if p_fso.FileExists(p_folderPath & fileName) AND IsImage(fileName) then

			if InStr(lcase(fileName), ".png") then
				paramArr(0) = "90%"
				paramArr(1) = "90%"		

				GetPictureData = paramArr				
			else
				Dim imgObj
                
				On Error Resume Next
				set imgObj = loadpicture(p_folderPath & fileName)

                if Err.Number <> 0 Then
				    paramArr(0) = "1000"
				    paramArr(1) = "1000"		
    				GetPictureData = paramArr
                                       
                else
				    paramArr(0) = round(imgObj.width / 26.4583)
				    paramArr(1) = round(imgObj.height / 26.4583)
				    GetPictureData = paramArr
    			    set imgObj = nothing
				end if

			end if
		else
			'-- no picture
				paramArr(0) = "100%"
				paramArr(1) = "100%"	

			
			GetPictureData = paramArr
		end if
		
	end function
	
	'--Get thumbnail file version
	Private function GetThumbVersion(ByVal fileName)

		Dim GTV_fileExtension, GTV_fileName
		GTV_fileExtension = GetFileExtension(fileName)
		
		if fileName = "" OR GTV_fileExtension = fileName Then
			GetThumbVersion = "noimage.png"
		else
			GTV_fileName = left(fileName, len(fileName) - 1 - len(GTV_fileExtension)) & "_thumb." & GTV_fileExtension
			
			if p_fso.FileExists(p_folderPath & GTV_fileName) Then
				GetThumbVersion = GTV_fileName
			else
				GetThumbVersion = "noimage.png"
			end if
		end if
		
	end function


End Class




Class clsVideoP

	Private p_id_video
	Private p_id_project
	Private p_url
	Private p_dateUp
	Private p_author
	Private p_tags
	Private p_footer
	
	Private p_dbConnection
	Private p_sqlString
	Private p_fso
	Private p_thumbFolderPath
	Private p_linkPath
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
		Set p_dbConnection	= Server.CreateObject("ADODB.Recordset")
		Set p_fso			= server.CreateObject ("Scripting.FileSystemObject")
		p_thumbFolderPath	= RealizePath(Application("VideoProjThumbsPath"))
		p_linkPath			= replace(Application("VideoProjThumbsPath"), "\", "/")
		
		p_id_video			= -1
		p_id_project		= -1
	End Sub
	
	Private Sub Class_Terminate()
	
		set p_dbConnection	= nothing
		Set p_fso			= nothing
		Set p_thumbFolderPath	= nothing
		
	End Sub

' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
	End Sub
' ------------------------------------------------------------------------------
	
	Public Sub LoadRecordData(lclRecordSet, IDProject)
		
		call CleanValues()
		
		if isNull(lclRecordSet) Then exit sub
		if lclRecordSet.State = 0 Then exit sub
		if lclRecordSet.EOF Then exit sub
		
		if NOT IsNull(IDProject) Then p_id_project	= safeCdbl(IDProject)
		
		p_id_video		= safeCdbl(lclRecordSet("id_video"))
		p_url			= lclRecordSet("url")
		p_dateUp		= lclRecordSet("dateup")
		p_author		= lclRecordSet("author")
		p_tags			= lclRecordSet("tags")
		p_footer		= lclRecordSet("footer")
		
	End Sub		
' ------------------------------------------------------------------------------
	Private Sub CleanValues()
		p_id_video		= -1
		p_id_project	= -1
		p_url			= null
		p_dateUp		= null
		p_author		= null
		p_tags			= null
		p_footer		= null
	end sub
	
' ------------------------------------------------------------------------------
	Public Property Get IDVideo()
		IDVideo = p_id_video
	End Property
' ------------------------------------------------------------------------------
	Public Property Get URL()
		URL = p_url
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Author()
		Author = p_author
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Footer()
		Footer = p_footer
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ThumbLink()
		ThumbLink = p_linkPath & GetThumbVersion()
	End Property
	
	'--Get thumbnail file version
	Private function GetThumbVersion()
		Dim GTV_fileName
		GTV_fileName = "vid_" & p_id_video & ".jpg"
		
		if p_fso.FileExists(p_thumbFolderPath & GTV_fileName) Then
			GetThumbVersion = GTV_fileName
		else
			GetThumbVersion = "noimage.png"
		end if
		
	end function
	
End Class


%>