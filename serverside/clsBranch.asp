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
'	Set branchItem = new clsBranch
'	call branchItem.LoadData(IDBranch)
'
'

Class clsBranch
' ------------------------------------------------------------------------------
	
	Private p_id_branch
	Private p_id_business
	Private p_businessName
	Private p_id_client
	Private p_clientName
	Private p_name
	Private p_id_zone
	Private p_id_country
	Private p_id_state
	Private p_id_city
	Private p_street
	Private p_number1
	Private p_local
	Private p_zip_code
	Private p_colonia

	
	Private p_areacode1
	Private p_phone1
	Private p_areacode2
	Private p_phone2
	Private p_areacode3
	Private p_phone3
	Private p_areacode4
	Private p_phone4
	Private p_areacode5
	Private p_phone5
	
	Private p_default_view
	Private p_img
	Private p_opening_hours	
	Private p_googlex
	Private p_googley
	Private p_zoom
	
	Private DBConn
	
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
	End Sub
' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
		if NOT IsNumeric(paramID) Then exit sub
		
		p_id_branch = paramID
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		SQLString = "clsBranchQuery @idBranch = " & p_id_branch
		
		DBConn.Open SQLString, oCNDB, 3, 3
		
		if NOT DBConn.EOF then
			
			p_id_branch			= safeCdbl(DBConn("id_branch"))
			p_id_business		= safeCdbl(DBConn("id_business"))
			p_businessName		= safeCdbl(DBConn("businessName"))
			p_id_client			= safeCdbl(DBConn("id_client"))
			p_clientName		= safeCdbl(DBConn("clientName"))
			p_name				= DBConn("name")
			p_id_zone			= safeCdbl(DBConn("id_zone"))
			p_id_country		= safeCdbl(DBConn("id_country"))
			p_id_state			= safeCdbl(DBConn("id_state"))
			p_id_city			= safeCdbl(DBConn("id_city"))
			p_street			= DBConn("street")
			p_number1			= DBConn("number1")
			p_local				= DBConn("local")
			p_zip_code			= DBConn("zip_code")
			p_colonia			= DBConn("colonia")
			
			p_areacode1			= DBConn("areacode1")
			p_phone1			= DBConn("phone1")
			p_areacode2			= DBConn("areacode2")
			p_phone2			= DBConn("phone2")
			p_areacode3			= DBConn("areacode3")
			p_phone3			= DBConn("phone3")
			p_areacode4			= DBConn("areacode4")
			p_phone4			= DBConn("phone4")
			p_areacode5			= DBConn("areacode5")
			p_phone5			= DBConn("phone5")
			
			p_default_view		= DBConn("default_view")
			p_img				= DBConn("img")
			p_opening_hours		= DBConn("opening_hours")
			p_googlex			= DBConn("googlex")
			p_googley			= DBConn("googley")
			p_zoom				= DBConn("zoom")
			
		else
			p_id_branch	= -1
		end if
		
		DBConn.close
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub Class_Terminate()
		
		Set DBConn = Nothing
		
	End Sub
' ------------------------------------------------------------------------------
	Public Property Get IDBranch()
		IDBranch = p_id_branch
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDBusiness()
		IDBusiness = p_id_business
	End Property
' ------------------------------------------------------------------------------
	Public Property Get BusinessName()
		BusinessName = p_businessName
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDClient()
		IDClient = p_id_client
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ClientName()
		ClientName = p_clientName
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Name()
		Name = p_name
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDZone()
		IDZone = p_id_zone
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDCountry()
		IDCountry = p_id_country
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDState()
		IDState = p_id_state
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDCity()
		IDCity = p_id_city
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Street()
		Street = p_street
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Number()
		Number = p_number1
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Local()
		Local = p_local
	End Property
' ------------------------------------------------------------------------------
	Public Property Get ZipCode()
		ZipCode = p_zip_code
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Colonia()
		Colonia = p_colonia
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Areacode1()
		Areacode1 = p_areacode1
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Phone1()
		Phone1 = p_phone1
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Areacode2()
		Areacode2 = p_areacode2
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Phone2()
		Phone2 = p_phone2
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Areacode3()
		Areacode3 = p_areacode3
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Phone3()
		Phone3 = p_phone3
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Areacode4()
		Areacode4 = p_areacode4
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Phone4()
		Phone4 = p_phone4
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Areacode5()
		Areacode5 = p_areacode5
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Phone5()
		Phone5 = p_phone5
	End Property
' ------------------------------------------------------------------------------
	Public Property Get DefaultView()
		DefaultView = p_default_view
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Img()
		Img = p_img
	End Property
' ------------------------------------------------------------------------------
	Public Property Get OpeningHours()
		OpeningHours = p_opening_hours
	End Property
' ------------------------------------------------------------------------------
	Public Property Get GoogleX()
		GoogleX = p_googlex
	End Property
' ------------------------------------------------------------------------------
	Public Property Get GoogleY()
		GoogleY = p_googley
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Zoom()
		Zoom = p_zoom
	End Property



End Class
' ------------------------------------------------------------------------------
%>