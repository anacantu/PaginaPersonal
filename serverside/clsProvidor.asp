<%
' ------------------------------------------------------------------------------
'	Author:		Carlos Trevino @ Pulso Vital Consulting Group
'	Email:		hobbes313@hotmail.com
'	URL:		
'	Date:		Sep 13, 2009
' ------------------------------------------------------------------------------
'
'	NOTE: 	data_storage/database.asp 
'			serverside/functions.asp	MUST BE LOADED
'
' Demo:
'	Set providorItem = new clsProvidor
'	call providorItem.LoadData(IDProvidor)
'
'

Class clsProvidor
' ------------------------------------------------------------------------------
	
	Private DBConn
	
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
	End Sub
' ------------------------------------------------------------------------------
	Public Sub LoadData(paramID)
		
		
		
		
	End Sub	
' -- Print list of article's related providors ---------------------------------
	Public Sub PrintRelatedProvidors(ByVal ArticleID)
		
		Set DBConn = Server.CreateObject("ADODB.Recordset")
		p_sqlString = "GetArticleRelatedProvidors @idArticle=" & ArticleID
		DBConn.Open p_sqlString, oCNDB, 3, 3
		
		if NOT DBConn.EOF Then
			'-- print header
			response.write "<div style=""width:160px; margin-bottom:20px""><h3 class=""bottom-line""><strong>Proveedores Relacionados</strong></h3>"
			
			While NOT DBConn.EOF
				'-- print providor's link
				response.write "<a class=""providor-link"" rev=""width: 750px; height: 530px; scrolling: no; "" title="""& DBConn("name") &""" rel=""lyteframe"" href=""/guide/business4.asp?buss="& DBConn("id_business") &""">"& DBConn("name") &"</a>"
				
				DBConn.movenext
				
			Wend
			'-- print footer
			response.write "<span class=""aad-more"" style=""margin:0; float:right""><a class=""headerColor"" href=""/guide/search.asp"">+ proveedores</a></span><div class=""clear""></div></div>"
			
		end if
		DBConn.close
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub Class_Terminate()
		
		Set DBConn = Nothing
		
	End Sub



End Class
' ------------------------------------------------------------------------------
%>