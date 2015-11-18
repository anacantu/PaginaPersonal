<!-- @start Areas Wrapper -->
		<div id="areas_wrapper">
			<div id="areas">
				
				<div id="footer-columns">
		
					<div class="col3">
						<ul>
							<li><a href="/">Inicio</a></li>
							
							<%
							Dim footerMenuRSet, footerItemCounter
							set footerMenuRSet = Server.CreateObject("ADODB.RecordSet")
							footerMenuRSet.open "exec GetTopMenuList", oCNDB, 3, 3
							
							footerItemCounter = 1
							
							While NOT footerMenuRSet.EOF
								response.write "<li><a href="""& footerMenuRSet("linkName") &""">"& footerMenuRSet("name") &"</a></li>" & vbcrlf
								footerMenuRSet.movenext()
								footerItemCounter = footerItemCounter + 1
								
								if footerItemCounter = 4 AND NOT footerMenuRSet.EOF Then
									footerItemCounter = 0
									response.write "</ul></div><div class=""col3""><ul>"
								end if
								
							Wend
							footerMenuRSet.close()
							set footerMenuRSet = Nothing
							%>
						</ul>
					</div>
					
					
				<!-- footer-columns ends -->
				</div>
			</div>
		</div>
    	<!-- @end Areas Wrapper -->
    	
    	<!-- @start Footer -->
		<div id="footer_wrapper">
   			<div id="footer">
    		
	   			<!-- @start Copyright -->
	   			<div id="copyright" class="copyright">
	   				Derechos Reservados &copy; 2009 <strong>Stilo</strong>
	   			</div>
	   			<!-- @end Copyright -->
	    			
	   			<!-- @start Footer Menu -->
	   			<ul id="footer_menu">
	   				<li><a href="/">Inicio</a></li>
	   				<li><a href="/legal.asp">Aviso legal</a></li>
	   				<li><a href="/privacidad.asp">Pol&iacute;ticas de privacidad</a></li>
	   				<li class="last-child"><a href="/contacto.asp">Contacto</a></li>
	   			</ul>
	   			<!-- @end Footer Menu -->
	   				
	   			<div id="main_footer">
	   				<a href="http://<%= urlServer %>" title="Stilo"><span>Stilo</span></a>
	   			</div>
	   				
	   		</div>
		</div>
		<!-- @end Footer -->