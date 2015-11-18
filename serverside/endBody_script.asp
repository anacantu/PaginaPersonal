<div id="iCoverLayer" style="position:absolute; visibility:hidden; left:0px; top:0px; filter: alpha(opacity=60, style=0); background-color:#000000; width:100%; height:0px; opacity: .5; -moz-opacity: .5; -khtml-opacity: .5"></div>
<div id="iMessageLayer" style="position:absolute; display:none; left:320px; top:660px;">
	<table id="iMessageTable" width="350" height="80" border="1" cellpadding="0" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
	  <tr>
		<td align="center" valign="middle" bordercolor="#FF6600"><table width="250" border="0" cellpadding="0" cellspacing="0" class="negro_small">
			<tr>
			  <td width="30" height="26" align="center" valign="middle"><img src="../images/iconos/wait20trans.gif" width="20" height="20"></td>
			  <td valign="middle"><em><span id="fulScreenMessageString">Estableciendo conexión con servidor...</span></em></td>
			</tr>
		  </table></td>
	  </tr>
	</table>
</div>
<script language="javascript">
	//init context menu, if defined
	if (isDefined('cm_Init')) { eval('cm_Init()') }
	
	//init popcalendar, if defined
	if (isDefined('initCal')) { eval('initCal()') }
	
	//init sorttable, if defined
	if (isDefined('sorttable')) { eval('sorttable.init()') }
	
	//init ajax connection objetc, if defined
	if (isDefined('ajax_connectionObject')) { eval('var connectionObject = new ajax_connectionObject(\'AJAXConnectionObj\')') }
	
	//init lytebox, if defined
	if (isDefined('LyteBox')) { eval('initLytebox()') }
	
	//load actions
	<%= loadActions %>
	
	document.onclick = function () {
		
		//Calendar
		if (isDefined('Calendar_DocOnClick')) { eval('Calendar_DocOnClick()') }
		
		//Context Menu
		if (isDefined('cm_DocOnClick')) { eval('cm_DocOnClick()') }
	}
	
	//corregir altura
	var tableContainer = GetElement('table-wrap');
	var adsContainer = GetElement('ads_container');
	if( tableContainer != undefined && adsContainer != undefined )
	{
		if(adsContainer.offsetHeight > tableContainer.offsetHeight)
			tableContainer.style.height = adsContainer.offsetHeight.toString() + 'px';
	}
</script>
<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try {
var pageTracker = _gat._getTracker("UA-801426-18");
pageTracker._trackPageview();
} catch(err) {}</script>