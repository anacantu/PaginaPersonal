<%
' ------------------------------------------------------------------------------
'	Author:		Carlos Trevino @ Pulso Vital Consulting Group
'	Email:		hobbes313@hotmail.com
'	URL:		
'	Date:		Jul 09, 2009
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

Class clsArticle
' ------------------------------------------------------------------------------
	
	Private p_id_article
	Private p_id_module
	Private p_moduleName
	Private p_title
	Private p_dateUp
	Private p_subtitle
	Private p_views
	Private p_datePublished
	Private p_EnglishdatePublished
	Private p_id_layout
	Private p_grade
    Private p_content
	Private p_content_noHTML
	Private p_tags
    Private p_published
	
	'-- pictures
	Private p_article_pictures
	Private p_thumb_picture
	Private p_main_picture
	
	'-- video
	Private p_article_videos
	
	Private DBConn
	Private ContributorsObj
	
	Private p_contributors_text
	Private p_sqlString
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
		p_thumb_picture = null
		p_main_picture = null
		p_contributors_text = ""
		set p_article_pictures	= new clsArticlePictures
		set p_article_videos	= new clsArticleVideos
		
	End Sub
' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
		if NOT IsNumeric(paramID) Then exit sub
		
		p_id_article = paramID
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		p_sqlString = "clsArticleQuery @idArticle = " & p_id_article
		DBConn.Open p_sqlString, oCNDB, 3, 3
		
		if NOT DBConn.EOF then
        			
			p_id_article	= safeCdbl(DBConn("id_article"))
			p_id_module		= safeCdbl(DBConn("id_module"))
			p_moduleName	= DBConn("moduleName")
			p_title			= DBConn("title")
			p_dateUp		= DBConn("dateUp")
			p_subtitle		= DBConn("subtitle")
			p_views			= safeCdbl(DBConn("views"))
			p_content		= DBConn("contenido") & DBConn("contenido2")
			p_content_noHTML= RemoveHTML(p_content)
			p_datePublished	= SpanishDateFormat(DBConn("datePublished"))
			p_EnglishdatePublished	= DBConn("datePublished")
			p_id_layout		= safeCdbl(DBConn("id_layout"))
			p_grade			= safeCint(DBConn("grade"))
			p_tags			= DBConn("tags")
            p_published		= safeCint(DBConn("published"))
			
			'--Load Article Pictures
			p_article_pictures.LoadData(p_id_article)
			
			'--Load Article Videos
			p_article_videos.LoadData(p_id_article)
			
			'--Load Article Contributors
			Set ContributorsObj = Server.CreateObject("ADODB.Recordset")
			p_sqlString = "SELECT name FROM contributor WHERE id_contributor IN ( SELECT id_contributor FROM article_contributor WHERE id_article = "&p_id_article&" )"
			ContributorsObj.Open p_sqlString, oCNDB, 3, 3
			
			While NOT ContributorsObj.EOF
				p_contributors_text = p_contributors_text & ContributorsObj("name") & ", "
				ContributorsObj.movenext()
			Wend
			p_contributors_text = RemoveLastSeparator(p_contributors_text, ", ")
				
			ContributorsObj.close()
			set ContributorsObj = Nothing
			
		else
			p_id_article	= -1
		end if
		
		DBConn.close
		
	End Sub	
' ------------------------------------------------------------------------------
	
	Public Sub LoadRecordData(lclRecordSet, loadPictures, loadVideos)
		
		call CleanValues()
		
		if isNull(lclRecordSet) Then exit sub
		if lclRecordSet.State = 0 Then exit sub
		if lclRecordSet.EOF Then exit sub
		
		p_id_article	= safeCdbl(lclRecordSet("id_article"))
		p_id_module		= safeCdbl(lclRecordSet("id_module"))
		p_moduleName	= lclRecordSet("moduleName")
		p_title			= lclRecordSet("title")
		p_subtitle		= lclRecordSet("subtitle")
		p_content		= lclRecordSet("contenido")
		p_content_noHTML= RemoveHTML(p_content)
		p_datePublished	= SpanishDateFormat(lclRecordSet("datePublished"))
		p_EnglishdatePublished	= lclRecordSet("datePublished")
		p_published		= safeCint(lclRecordSet("published"))

		if loadPictures Then
			'--Load Article Pictures
			p_article_pictures.LoadData(p_id_article)
		end if
		
		if loadVideos Then
			'--Load Article Videos
			p_article_videos.LoadData(p_id_article)
		end if
		
	End Sub	
' ------------------------------------------------------------------------------
	Public Sub LoadRandomCover(ByVal avoidModuleID)
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		p_sqlString = "SELECT TOP 1 "&_
		" id_article, id_module, moduleName, title, dateUp, subtitle, convert(varchar(1000),content) as contenido, datePublished, published "&_
		" FROM article_view "&_
		" /*filter*/ "&_
		" ORDER BY NEWID()"
		
		if avoidModuleID <> "" Then
			p_sqlString = replace(p_sqlString, "/*filter*/", "WHERE id_module <> " & safeCstr(avoidModuleID))
		end if
		
		DBConn.Open p_sqlString, oCNDB, 3, 3
		
		if NOT DBConn.EOF then
			
			p_id_article	= safeCdbl(DBConn("id_article"))
			p_id_module		= safeCdbl(DBConn("id_module"))
			p_moduleName	= DBConn("moduleName")
			p_title			= DBConn("title")
			p_subtitle		= DBConn("subtitle")
			p_content		= DBConn("contenido")
			p_content_noHTML= RemoveHTML(p_content)
			p_datePublished	= SpanishDateFormat(DBConn("datePublished"))
			p_EnglishdatePublished	= DBConn("datePublished")
            p_published = safeCint(DBConn("published"))
			
			'--Load Article Pictures
			p_article_pictures.LoadData(p_id_article)
			
		else
			p_id_article	= -1
		end if
		
		DBConn.close
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub CleanValues()
		p_id_article	= -1
		p_id_module		= -1
		p_moduleName	= null
		p_title			= null
		p_subtitle		= null
		p_content		= null
		p_content_noHTML= null
		p_datePublished	= null
		p_EnglishdatePublished	= null
        p_published = null
	end sub	
' ------------------------------------------------------------------------------
	Private Sub Class_Terminate()
		
		Set DBConn = Nothing
		
	End Sub
' ------------------------------------------------------------------------------
	Public Property Get IDArticle()
		IDArticle = p_id_article
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDModule()
		IDModule = p_id_module
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ModuleName()
		ModuleName	 = p_moduleName
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Title()
		Title	 = p_title
	End Property	
' ------------------------------------------------------------------------------
	Public Property Get DateUp()
		DateUp	 = p_dateUp
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Subtitle()
		Subtitle	 = p_subtitle
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Views()
		Views	 = p_views
	End Property
' ------------------------------------------------------------------------------
	Public Property Get DatePublished()
		DatePublished	 = p_datePublished
	End Property
' ------------------------------------------------------------------------------
	Public Property Get EnglishDatePublished()
		EnglishDatePublished	 = p_EnglishdatePublished
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ReadableDatePublished()
		ReadableDatePublished	 = formatDateString(p_EnglishdatePublished, 1)
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDLayout()
		IDLayout	 = p_id_layout
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Grade()
		Grade	 = p_grade
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Content()
		Content	 = p_content
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
	Public Property Get Tags()
		Tags	 = p_tags
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Published()
		Published	 = p_published
	End Property
' ------------------------------------------------------------------------------
	Public Property Get PicturesPath()
		PicturesPath	 = replace(Application("ArticlePicPath"), "\", "/")
	End Property
' ------------------------------------------------------------------------------
	Public Property Get VideoThumbsPath()
		VideosThumbPath	 = replace(Application("VideoThumbsPath"), "\", "/")
	End Property

' ------------------------------------------------------------------------------
' -  Contributors
' ------------------------------------------------------------------------------
	Public Property Get ContributorsText()
		ContributorsText = p_contributors_text
	End Property



' ------------------------------------------------------------------------------
' -  Pictures
' ------------------------------------------------------------------------------
	Public Property Get Pictures()
		set Pictures = p_article_pictures
	End Property

	Public Property Get MainThumbPictureFull()
		MainThumbPictureFull = Pictures.MainThumbPictureFull
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ThumbPicture()
		if IsNull(p_thumb_picture) Then
			p_thumb_picture = SearchPicture("thumb_" & IDArticle)
		end if
		
		ThumbPicture = p_thumb_picture		
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ThumbPictureLink()
		ThumbPictureLink = replace(Application("ArticlePicPath"), "\", "/") & ThumbPicture
	End Property
' ------------------------------------------------------------------------------
	Public Property Get MainPicture()
		if IsNull(p_main_picture) Then
			p_main_picture = SearchPicture("main_" & IDArticle)
		end if
		
		MainPicture = p_main_picture		
	End Property
' ------------------------------------------------------------------------------
	Public Property Get MainPictureLink()
		MainPictureLink = replace(Application("ArticlePicPath"), "\", "/") & MainPicture
	End Property
' ------------------------------------------------------------------------------


	Private function SearchPicture(ByVal picFileName)
		Dim FSO, foderPath, lclPicture
		Set FSO = server.CreateObject ("Scripting.FileSystemObject")
		foderPath = RealizePath(Application("ArticlePicPath")) & "\"
		
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
		oCNDB.EXECUTE "IncrementArticleViews @id_article = " & p_id_article
	End Sub
	

' ------------------------------------------------------------------------------
' -  Videos
' ------------------------------------------------------------------------------
	Public Property Get Videos()
		set Videos = p_article_videos
	End Property

	
End Class
' ------------------------------------------------------------------------------







Class clsArticlePictures
	
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
		p_folderPath		= RealizePath(Application("ArticlePicPath"))
		
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
		
		p_sqlString = "GetArticlePictures @idArticle = " & paramID
		p_dbConnection.Open p_sqlString, oCNDB, 3, 3
		
		p_top_count = 0
		p_pictures_count = 0
		
		While NOT p_dbConnection.EOF
			
			
			if cBool(p_dbConnection("is_main_pic")) Then p_id_main_picture = safeCdbl(p_dbConnection("id_picture"))
			
			Redim Preserve p_arr_pictures(p_pictures_count)
			set p_arr_pictures(p_pictures_count) = new clsPicture
			
			call p_arr_pictures(p_pictures_count).LoadRecordData(p_dbConnection, p_dbConnection("id_article"))
			
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
			MainThumbPictureFull = Application("ArticlePicPath") & GetThumbVersion(p_arr_pictures(0).FileName)
		else
			MainThumbPictureFull = Application("ArticlePicPath") & GetThumbVersion("")
		end if
	End Property
	
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




Class clsArticleVideos
	
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
		p_folderPath		= RealizePath(Application("ArticlePicPath"))
		
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
		
		p_sqlString = "GetArticleVideos @idArticle = " & paramID
		p_dbConnection.Open p_sqlString, oCNDB, 3, 3
		
		p_videos_count = 0
		
		While NOT p_dbConnection.EOF
			
			
			Redim Preserve p_arr_videos(p_videos_count)
			set p_arr_videos(p_videos_count) = new clsVideo
			
			call p_arr_videos(p_videos_count).LoadRecordData(p_dbConnection, p_dbConnection("id_article"))
			
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




Class clsPicture

	Private p_id_picture
	Private p_id_article
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
		p_folderPath		= RealizePath(Application("ArticlePicPath"))
		p_linkPath			= replace(Application("ArticlePicPath"), "\", "/")
		
		p_id_picture		= -1
		p_id_article		= -1
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
	
	Public Sub LoadRecordData(lclRecordSet, IDArticle)
		
		call CleanValues()
		
		if isNull(lclRecordSet) Then exit sub
		if lclRecordSet.State = 0 Then exit sub
		if lclRecordSet.EOF Then exit sub
		
		Dim picData
		
		if NOT IsNull(IDArticle) Then p_id_article	= safeCdbl(IDArticle)
		
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
		p_id_article	= -1
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

				set imgObj = loadpicture(p_folderPath & fileName)
				
				paramArr(0) = round(imgObj.width / 26.4583)
				paramArr(1) = round(imgObj.height / 26.4583)
				
				GetPictureData = paramArr
				set imgObj = nothing
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




Class clsVideo

	Private p_id_video
	Private p_id_article
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
		p_thumbFolderPath	= RealizePath(Application("VideoThumbsPath"))
		p_linkPath			= replace(Application("VideoThumbsPath"), "\", "/")
		
		p_id_video			= -1
		p_id_article		= -1
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
	
	Public Sub LoadRecordData(lclRecordSet, IDArticle)
		
		call CleanValues()
		
		if isNull(lclRecordSet) Then exit sub
		if lclRecordSet.State = 0 Then exit sub
		if lclRecordSet.EOF Then exit sub
		
		if NOT IsNull(IDArticle) Then p_id_article	= safeCdbl(IDArticle)
		
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
		p_id_article	= -1
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