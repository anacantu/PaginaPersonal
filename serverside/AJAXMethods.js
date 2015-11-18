<!--
<!--
/**************************************************************************************************
	Metodos de comunicacion con el servidor, programacion tipo AJAX
***************************************************************************************************/				


/**************************************************************************************************
	Predefined Variables
***************************************************************************************************/				
	var generalURL = '/serverside/RPC.asp';

	var connObj = new Array();

/**************************************************************************************************
	AJAX Server Requests
***************************************************************************************************/				
	function build_HttpRequest() {
		if (window.XMLHttpRequest) { // Mozilla, Safari,...
			var request = new XMLHttpRequest();
			if (request.overrideMimeType) { request.overrideMimeType('text/xml');}
		} else if (window.ActiveXObject) { // IE
			try {var request = new ActiveXObject("Msxml2.XMLHTTP");} 
			catch (e) {
				try {var request = new ActiveXObject("Microsoft.XMLHTTP");} catch (e) {}
			}
		}
		if (!request) { alert('Cannot create an XMLHTTP instance');return false;}
		
		
		connObj[connObj.length] = request;
		
						// ID      -     httpRequest
		return new Array(connObj.length-1, request);
	}
	
// FUNCTION: Make Request to Server passing GET variables
	function ajax_makeRequest(methodName, urlParams, startFunction, endFunction) {
		var newhttpbj = build_HttpRequest();
		var http_request = newhttpbj[1];
		var httpID		= newhttpbj[0];
		
		if (startFunction != '')
		{	
			eval(startFunction + '(http_request, methodName, '+ httpID.toString() +')');
		}
		else
		{	
			if(typeof AJAX_onStartRequest == 'function') {
				eval('AJAX_onStartRequest(http_request, methodName, '+ httpID.toString() +')');
			} 
		}
		
		http_request.onreadystatechange = function() {ajaxReadyStateChange(endFunction, methodName, httpID);}
		http_request.open('GET', generalURL + '?method=' + methodName + '&' + urlParams + '&flagID=' + getRandomNumber(), true);
		http_request.send(null);

		return httpID;
	}
	
	
// FUNCTION: Make Request to Server using POST method
	function ajax_makeRequest_Post(methodName, urlParams, startFunction, endFunction) {
		var newhttpbj = build_HttpRequest();
		var http_request = newhttpbj[1];
		var httpID		= newhttpbj[0];
				
		if (startFunction != '')
		{	
			eval(startFunction + '(http_request, methodName, '+ httpID.toString() +')');
		}
		else
		{	
			if(typeof AJAX_onStartRequest == 'function') {
				eval('AJAX_onStartRequest(http_request, methodName, '+ httpID.toString() +')');
			}
		}
		
		http_request.onreadystatechange = function() {ajaxReadyStateChange(endFunction, methodName, httpID);}
		http_request.open('POST', generalURL);
		http_request.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
		http_request.send('method=' + methodName + '&' + urlParams + '&flagID=' + getRandomNumber());
	}
	

	function ajaxReadyStateChange(ejecFunction, methodName, httpID)
	{
	// readyState values
	// 0	The request is not initialized
	// 1	The request has been set up
	// 2	The request has been sent
	// 3	The request is in process
	// 4	The request is complete
	
		var http_request = connObj[httpID];
		
		switch(http_request.readyState)
		{
			case 4:
				//verificar errores
				
				/*
				if (http_request.status != 200) 
				{
					//enviar reporte de error
					ajax_makeRequest('SendErrorReport', 'subject=' + escape('SAD: Error dentro de RPC') + '&body=' + escape('Ha ocurrido un error al intentar ejecutar el metodo: ') + methodName + '<br><br>Response Text: ' + escape(http_request.responseText),'','');
				}*/
				
				// resultado correcto
				if (ejecFunction != '')
				{	
					eval(ejecFunction + '(http_request, methodName, httpID)');
				}
				else
				{	
					if(typeof AJAX_onCompleteRequest == 'function') {
						eval('AJAX_onCompleteRequest(http_request, methodName, httpID)');
					}
				}
				
				delete(connObj[httpID]);
				
				break;
			
		}
	}

	function ajax_connectionObject(id)
	{
		this.makeRequest = ajax_makeRequest;
		this.makePostRequest = ajax_makeRequest_Post;
		this.id = id;
		//this.http_request = http_request;
	}
	
	var auxID = ''
	
	function AJAX_SaveData(tablename, IndexColumn, IndexValue, columnKey, columnType, value, waitIconDivID)
	{
		var lcl_urlParams = 'tablename='+tablename+'&indexcolumn='+IndexColumn+'&id='+IndexValue+'&columnKey='+columnKey+'&columnType='+columnType+'&value=' + value;
		
		var waitDiv = AJAX_GetElement(waitIconDivID);
		if (waitDiv != null && waitDiv != undefined)
			waitDiv.style.display = '';
		
		//respaldar ID div
		auxID = waitIconDivID;
		
		ajax_makeRequest('SaveData', lcl_urlParams, '', 'AJAX_SavedData');
	}
	
	function AJAX_SavePostData(tablename, IndexColumn, IndexValue, columnKey, columnType, value, waitIconDivID)
	{
		var lcl_urlParams = 'tablename='+tablename+'&indexcolumn='+IndexColumn+'&id='+IndexValue+'&columnKey='+columnKey+'&columnType='+columnType+'&value=' + value;
		
		var waitDiv = AJAX_GetElement(waitIconDivID);
		if (waitDiv != null && waitDiv != undefined)
			waitDiv.style.display = '';
		
		//respaldar ID div
		auxID = waitIconDivID;
		
		ajax_makeRequest_Post('SavePostData', lcl_urlParams, '', 'AJAX_SavedData');
	}
	
	function AJAX_SavedData(http_request, methodName, httpID)
	{
		if(auxID != '')
		{
			 var waitDiv = AJAX_GetElement(auxID);
			if (waitDiv != null && waitDiv != undefined)
				waitDiv.style.display = 'none';
		}
		auxID = '';
	}
	
	
	
	//Obtener todos los elementos de una forma 
	//y concatenar un string con el nombre del campo y el valor
	function build_DataString(tagsArrayString, encodeTextAreaToHTML)
	{
		var tagsArray = tagsArrayString.split("|");
		var getstr = '';
		
		for (thisIndex=0; thisIndex<tagsArray.length; thisIndex++) {
			var thisElement = document.getElementById(tagsArray[thisIndex]);
			
			switch(thisElement.tagName)
			{
				case "INPUT":
					if (thisElement.type == "text" || thisElement.type == "hidden") {
					   getstr += thisElement.name + "=" + escape(thisElement.value) + "&";
					}
					if (thisElement.type == "checkbox") {
					   if (thisElement.checked) {
						  getstr += thisElement.name + "=" + escape(thisElement.value) + "&";
					   } else {
						  getstr += thisElement.name + "=&";
					   }
					}
					if (thisElement.type == "radio") {
					   if (thisElement.checked) {
						  getstr += thisElement.name + "=" + escape(thisElement.value) + "&";
					   }
					}
					break;
				case "SELECT":
					var sel = thisElement;
					getstr += sel.name + "=" + sel.options[sel.selectedIndex].value + "&";
					break; 
				case "TEXTAREA":
					if (encodeTextAreaToHTML)
						getstr += thisElement.name + "=" + AJAX_encodeToHTML(escape(thisElement.value)) + "&";
					else
						getstr += thisElement.name + "=" + escape(thisElement.value) + "&";
					break;
			 }	
		}
		return getstr;
	}

	//Obtener un numero al azar entre cero y cinco mil
	function getRandomNumber()
	{
		var ranNum= Math.floor(Math.random()*5000);
		return ranNum;
	}
	
	//Cambiar saltos de linea por <br>
	function AJAX_encodeToHTML(texte)
	{
		var Resultat = '';
		for (encodeCounter=0;encodeCounter<texte.length;encodeCounter++)
		{
			numer=texte.charCodeAt(encodeCounter);
			if((numer==13)&&(texte.charCodeAt(encodeCounter+1)==10))
			{
				encodeCounter++;
				Resultat += '<br>';
			}
			else
				Resultat += String.fromCharCode(numer);			
		}
		return Resultat;
	}
	
	//Get Element (Object) found by its ID
	function AJAX_GetElement(tagId) 
	{
		var lclObj = document.getElementById(tagId);
		if(lclObj && lclObj.length && lclObj[0].id==tagId)
			lclObj=lclObj[0];
		return lclObj;
	}
	
	
-->