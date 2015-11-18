<!-- !Ads -->
				<div id="ads_container">
					<img src="/images/ads/providors_header.png" width="120" height="27">
					<%
					'--banner registro proveedores
					if fileName = "search.asp" OR fileName = "results.asp" Then
					%>
					<a href="/register.asp"><img src="/images/ads/bannerregistro.gif" width="120" height="258"></a>
					<%
					else
					%>
					<!-- suscripcion usuarios -->
					<a href="/suscription.asp"><img src="/images/ads/banner_suscribete.jpg" width="120" height="299"></a>
					<%
					end if
					%>
					<img src="/images/ads/providors_header.png" width="120" height="27">
					<%
					
					set adsObj = server.createObject("ADODB.RecordSet")
					sqlString = "EXEC GetPublicBanners"
					adsObj.Open sqlString, oCNDB, 3, 3
					
					While NOT adsObj.EOF
						
						if adsObj("imagefile") <> "" Then
							
							if adsObj("url") <> "" Then response.write "<a href="""& adsObj("url") &""" target=""_blank"" onclick=""SaveBannerClick('"& adsObj("id_banner") &"')"">"
							response.write "<img src=""/images/ads/"& adsObj("imagefile") &""" width=""120"" height=""120"">" & vbcrlf
							if adsObj("url") <> "" Then response.write "</a>"
							
						end if
						
						adsObj.movenext
					Wend
					
					adsObj.close
					set adsObj = Nothing
					%>
				</div>
				<!-- End Ads -->