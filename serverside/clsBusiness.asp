<%
' ------------------------------------------------------------------------------
'	Author:		Carlos Trevino @ STILO
'	Email:		hobbes313@hotmail.com
'	URL:		
'	Date:		Ene 05, 2010
' ------------------------------------------------------------------------------
'
'	NOTE: 	data_storage/database.asp 
'			serverside/functions.asp	MUST BE LOADED
'
' Demo:
'	Set businessItem = new clsBusiness
'	call businessItem.LoadData(IDBusiness)
'
'

Class clsBusiness
' ------------------------------------------------------------------------------
	
	Private p_id_business
	Private p_id_client
	Private p_name
	Private p_website
	Private p_dateUp
	Private p_email
	Private p_productservice
	Private p_labels
	Private p_branch
	Private p_description
	Private p_logo
	Private p_image1
	Private p_image2
	Private p_image3
	Private p_image4
	Private p_image5
	Private p_promotion
	Private p_promotion_startdate
	Private p_promotion_enddate	
	Private p_priority
	Private p_voteCount
	Private p_starCount
	
	Private DBConn
	
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
	End Sub
' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
		if NOT IsNumeric(paramID) Then exit sub
		
		p_id_business = paramID
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		SQLString = "clsBusinessQuery @idBusiness = " & p_id_business
		DBConn.Open SQLString, oCNDB, 3, 3
		
		if NOT DBConn.EOF then
			
			p_id_business	   	= safeCdbl(DBConn("id_business"))
			p_id_client		   	= safeCdbl(DBConn("id_client"))
			p_name			   	= DBConn("name")
			p_website		   	= DBConn("website")
			p_dateUp		   	= DBConn("dateUp")
			p_email			   	= DBConn("email")
			p_productservice   	= DBConn("productservice")
			p_labels		   	= DBConn("labels")
			p_branch		   	= DBConn("branch")
			p_description	   	= DBConn("description")
			p_logo			   	= DBConn("logo")
			p_image1		   	= DBConn("image1")
			p_image2		   	= DBConn("image2")
			p_image3		   	= DBConn("image3")
			p_image4		   	= DBConn("image4")
			p_image5		   	= DBConn("image5")
			p_promotion		   	= DBConn("promotion")
			p_promotion_startdate	= DBConn("promotion_startdate")			
			p_promotion_enddate	= DBConn("promotion_enddate")
			p_priority		   	= DBConn("priority")
			p_voteCount		   	= DBConn("voteCount")
			p_starCount		   	= DBConn("starCount")
			
		else
			p_id_business	= -1
		end if
		
		DBConn.close
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub Class_Terminate()
		
		Set DBConn = Nothing
		
	End Sub
' ------------------------------------------------------------------------------
	Public Property Get IDBusiness()
		IDBusiness = p_id_business
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDClient()
		IDClient = p_id_client
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Name()
		Name = p_name
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Website()
		Website = p_website
	End Property
' ------------------------------------------------------------------------------
	Public Property Get DateUp()
		DateUp = spanishDateFormat(p_dateUp)
	End Property
' ------------------------------------------------------------------------------
	Public Property Get EnglishDateUp()
		DateUp = p_dateUp
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Email()
		Email = p_email
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ProductService()
		ProductService = p_productservice
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Labels()
		Labels = p_labels
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Branch()
		Branch = p_branch
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Description()
		Description = p_description
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Image1()
		Image1 = p_image1
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Image2()
		Image2 = p_image2
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Image3()
		Image3 = p_image3
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Image4()
		Image4 = p_image4
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Image5()
		Image5 = p_image5
	End Property	
' ------------------------------------------------------------------------------
	Public Property Get Promotion()
		Promotion = p_promotion
	End Property	
' ------------------------------------------------------------------------------
	Public Property Get PromotionStartDate()
		PromotionStartDate = spanishDateFormat(p_promotion_startdate)
	End Property
' ------------------------------------------------------------------------------
	Public Property Get EnglishPromotionStartDate()
		EnglishPromotionStartDate = p_promotion_startdate
	End Property
' ------------------------------------------------------------------------------
	Public Property Get PromotionEndDate()
		PromotionEndDate = spanishDateFormat(p_promotion_enddate)
	End Property
' ------------------------------------------------------------------------------
	Public Property Get EnglishPromotionEndDate()
		EnglishPromotionEndDate = p_promotion_enddate
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Priority()
		Priority = p_priority
	End Property	
' ------------------------------------------------------------------------------
	Public Property Get VoteCount()
		VoteCount = p_voteCount
	End Property	
' ------------------------------------------------------------------------------
	Public Property Get StarCount()
		StarCount = p_starCount
	End Property		


End Class
' ------------------------------------------------------------------------------
%>