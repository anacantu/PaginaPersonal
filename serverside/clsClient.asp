<%
' ------------------------------------------------------------------------------
'	Author:		Carlos Trevino @ STILO
'	Email:		hobbes313@hotmail.com
'	URL:		
'	Date:		Ene 02, 2010
' ------------------------------------------------------------------------------
'
'	NOTE: 	data_storage/database.asp 
'			serverside/functions.asp	MUST BE LOADED
'
' Demo:
'	Set clientItem = new clsClient
'	call clientItem.LoadData(IDClient)
'
'

Class clsClient
' ------------------------------------------------------------------------------
	
	Private p_id_client
	Private p_dateUp
	Private p_lastaccess
	Private p_id_client_status 'p_intregistry
	
	Private p_name	
	Private p_lastname_1
	Private p_lastname_2
	Private p_birthdate
	Private p_email
	Private p_gender
	Private p_id_marital
	Private p_id_occupation
	Private p_id_country
	Private p_id_state
	Private p_id_city
	Private p_id_zone
	Private p_phone
	Private p_phone_areacode
	Private p_cellphone
	
	Private p_pr_cellphone
	Private p_pr_email
	Private p_pr_name
	Private p_pr_phone
	Private p_pr_phone_areacode
	
	Private p_sales_cellphone
	Private p_sales_email
	Private p_sales_name
	Private p_sales_phone
	Private p_sales_phone_areacode
	
	Private p_username
	Private p_password
	
	Private DBConn
	
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
	End Sub
' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
		if NOT IsNumeric(paramID) Then exit sub
		
		p_id_client = paramID
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		SQLString = "clsClientQuery @idClient = " & p_id_client
		DBConn.Open SQLString, oCNDB, 3, 3
		
		if NOT DBConn.EOF then
			
			p_id_client			= safeCdbl(DBConn("id_client"))
			p_dateUp			= DBConn("dateUp")
			p_lastaccess		= DBConn("lastaccess")
			p_id_client_status	= safeCdbl(DBConn("id_client_status"))
			
			p_name				= DBConn("name")
			p_lastname_1		= DBConn("lastname_1")
			p_lastname_2		= DBConn("lastname_2")
			
			p_birthdate			= DBConn("birthdate")
			p_email				= DBConn("email")
			p_gender			= DBConn("gender")
			p_id_marital		= safeCdbl(DBConn("id_marital"))
			p_id_occupation		= safeCdbl(DBConn("id_occupation"))
			p_id_country		= safeCdbl(DBConn("id_country"))
			p_id_state			= safeCdbl(DBConn("id_state"))
			p_id_city			= safeCdbl(DBConn("id_city"))
			p_id_zone			= safeCdbl(DBConn("id_zone"))
			p_phone				= DBConn("phone")
			p_phone_areacode	= DBConn("phone_areacode")
			p_cellphone			= DBConn("cellphone")
			
			p_pr_cellphone		= DBConn("pr_cellphone")
			p_pr_email			= DBConn("pr_email")
			p_pr_name			= DBConn("pr_name")
			p_pr_phone			= DBConn("pr_phone")
			p_pr_phone_areacode	= DBConn("pr_phone_areacode")
			
			p_sales_cellphone	= DBConn("sales_cellphone")
			p_sales_email		= DBConn("sales_email")
			p_sales_name		= DBConn("sales_name")
			p_sales_phone		= DBConn("sales_phone")
			p_sales_phone_areacode	= DBConn("sales_phone_areacode")
			
			p_username			= DBConn("username")
			p_password			= DBConn("password")
			
		else
			p_id_client	= -1
		end if
		
		DBConn.close
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub Class_Terminate()
		
		Set DBConn = Nothing
		
	End Sub
' ------------------------------------------------------------------------------
	Public Property Get IDClient()
		IDClient = p_id_client
	End Property
' ------------------------------------------------------------------------------
	Public Property Get DateUp()
		DateUp = p_dateUp
	End Property
' ------------------------------------------------------------------------------
	Public Property Get LastAccess()
		LastAccess = p_lastaccess
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDClientStatus()
		IDClientStatus = p_id_client_status
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Name()
		Name	 = p_name
	End Property
' ------------------------------------------------------------------------------
	Public Property Get LastName1()
		LastName1	 = p_lastname_1
	End Property
' ------------------------------------------------------------------------------
	Public Property Get LastName2()
		LastName2	 = p_lastname_2
	End Property
' ------------------------------------------------------------------------------
	Public Property Get CompleteName()
		CompleteName	 = p_name & " " & p_lastname_ & " " & p_lastname_2
	End Property
' ------------------------------------------------------------------------------
	Public Property Get BirthDate()
		BirthDate	 = spanishDateFormat(p_birthdate)
	End Property
' ------------------------------------------------------------------------------
	Public Property Get EnglishBirthDate()
		EnglishBirthDate	 = p_birthdate
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Email()
		Email	 = p_email
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Gender()
		Gender	 = p_gender
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDMarital()
		IDMarital	 = p_id_marital
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDOccupation()
		IDOccupation	 = p_id_occupation
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDCountry()
		IDCountry	 = p_id_country
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDState()
		IDState	 = p_id_state
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDCity()
		IDCity	 = p_id_city
	End Property
' ------------------------------------------------------------------------------
	Public Property Get IDZone()
		IDZone	 = p_id_zone
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Phone()
		Phone	 = p_phone
	End Property
' ------------------------------------------------------------------------------
	Public Property Get PhoneAreacode()
		PhoneAreacode	 = p_phone_areacode
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Cellphone()
		Cellphone	 = p_cellphone
	End Property
' ------------------------------------------------------------------------------
	Public Property Get PRCellphone()
		PRCellphone	 = p_pr_cellphone
	End Property
' ------------------------------------------------------------------------------
	Public Property Get PREmail()
		PREmail	 = p_pr_email
	End Property
' ------------------------------------------------------------------------------
	Public Property Get PRName()
		PRName	 = p_pr_name
	End Property
' ------------------------------------------------------------------------------
	Public Property Get PRPhone()
		PRPhone	 = p_pr_phone
	End Property
' ------------------------------------------------------------------------------
	Public Property Get PRPhoneAreacode()
		PRPhoneAreacode	 = p_pr_phone_areacode
	End Property
' ------------------------------------------------------------------------------
	Public Property Get SalesCellphone()
		SalesCellphone	 = p_sales_cellphone
	End Property
' ------------------------------------------------------------------------------
	Public Property Get SalesEmail()
		SalesEmail	 = p_sales_email
	End Property
' ------------------------------------------------------------------------------
	Public Property Get SalesName()
		SalesName	 = p_sales_name
	End Property
' ------------------------------------------------------------------------------
	Public Property Get SalesPhone()
		SalesPhone	 = p_sales_phone
	End Property
' ------------------------------------------------------------------------------
	Public Property Get SalesPhoneAreacode()
		SalesPhoneAreacode	 = p_sales_phone_areacode
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Username()
		Username	 = p_username
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Password()
		Password	 = p_password
	End Property
' ------------------------------------------------------------------------------


End Class
' ------------------------------------------------------------------------------
%>