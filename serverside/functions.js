<!--

//Detect IE
var IEBrowser = (navigator.appVersion.indexOf("MSIE")!=-1);


function SocialShare(siteKey)
{
	var u = location.href;
	var t = document.title;
	
	switch(siteKey)
	{
		case 'facebook': window.open('http://www.facebook.com/sharer.php?u='+encodeURIComponent(u)+'&t='+encodeURIComponent(t),'sharer','toolbar=0,status=0,width=626,height=436');
			break;
		
	}
	return false;
}

//-- picture navigation
function PictureNavigation(galleryID, itemIndex)
{
	var linkTag = GetElement(galleryID + '_link');
	var picTag = GetElement(galleryID);
	var indexTag = GetElement(galleryID + '_index');
	var newSrc = '';
	var oldBtn = GetElement(galleryID + '_sel_' + indexTag.value.toString());
	
	var imgPreloader = new Image();
	imgPreloader.onload = function() {
		picTag.src = imgPreloader.src;
		picTag.width = imgPreloader.width;
		picTag.height = imgPreloader.height;
		
		imgPreloader.onload = function() {};
	}
	
	switch(itemIndex)
	{
		case -1:
			if(indexTag.value > 1)
			{
				indexTag.value = indexTag.value*1 - 1;
				newSrc = GetTagValue(galleryID + '_' + indexTag.value);
				imgPreloader.src = newSrc;
				linkTag.href = newSrc;
				
				oldBtn.src = '/images/pic_selector_off.png';
				GetElement(galleryID + '_sel_' + indexTag.value.toString()).src = '/images/pic_selector_on.png';
			}
			break;
			
		case 0:
			var nextIndex = indexTag.value*1 + 1;
			if(GetTagValue(galleryID + '_' + nextIndex.toString()) != null)
			{
				indexTag.value = nextIndex;
				newSrc = GetTagValue(galleryID + '_' + indexTag.value);
				imgPreloader.src = newSrc;
				linkTag.href = newSrc;
				
				oldBtn.src = '/images/pic_selector_off.png';
				GetElement(galleryID + '_sel_' + indexTag.value.toString()).src = '/images/pic_selector_on.png';
			}
			break;
			
		default:
			indexTag.value = itemIndex;
			newSrc = GetTagValue(galleryID + '_' + indexTag.value);
			imgPreloader.src = newSrc;
			linkTag.href = newSrc;
			
			oldBtn.src = '/images/pic_selector_off.png';
			GetElement(galleryID + '_sel_' + indexTag.value.toString()).src = '/images/pic_selector_on.png';
			
	}
}


function UpdateSelItemList(itemListTagID, checkboxTag)
{
	var listStr;
	if (checkboxTag.checked)
		listStr = GetTagValue(itemListTagID) + checkboxTag.value + ','
	else
		listStr = ReplaceSubString(GetTagValue(itemListTagID),',' + checkboxTag.value + ',',',')
	
	SetTagValue(itemListTagID, listStr);
}

	
//menu functions
function OpenSubMenu(idParent)
{
	var menuItem = GetElement('mainmenu');
	for (loopIndex = 0; loopIndex < menuItem.childNodes.length; loopIndex++)
	{
		var menuChildItem = menuItem.childNodes[loopIndex];
		
		//jucu.coma = true;
		
		if(menuChildItem.tagName == 'LI' && menuChildItem.getAttribute('idparent') == idParent)
		{
			menuChildItem.style.display = (menuChildItem.style.display == '') ? 'none' : '';
		}
	}
}

//Verificar si una direccion de correo esta escrita correctamente
function isEmailValid(email) 
{ 
    return /^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w{2,}$/.test(email); 
} 

//validar si un texto es un valor numerico
function IsNumeric(sText)
{
	var ValidChars = "0123456789.,$";
	var IsNumber=true;
	var Char;
	
	for (i = 0; i < sText.length && IsNumber == true; i++) 
	  { 
	  Char = sText.charAt(i); 
	  if (ValidChars.indexOf(Char) == -1) 
	     {
	     IsNumber = false;
	     }
	  }
	return IsNumber;
   
}

function isDefined(variable)
{
	return eval('(typeof('+variable+') != "undefined" && typeof('+variable+') != "unknown");');
}

function typeOf(value) {
    var s = typeof value;
    if (s === 'object') {
        if (value) {
            if (typeof value.length === 'number' &&
                    !(value.propertyIsEnumerable('length')) &&
                    typeof value.splice === 'function') {
                s = 'array';
            }
        } else {
            s = 'null';
        }
    }
    return s;
}

//inputParam in format mm/dd/yyyy
function isDateValid(inputParam)
{
	if(inputParam != '') 
	{
		inputParam = replaceAll(inputParam,'-','/'); 
		
		// regular expression to match required date format 
		re = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/; 
	 
		if(regs = inputParam.match(re)) 
		{ 
			var dteDate;
			//javascript months start at 0 (0-11 instead of 1-12)
			dteDate = new Date(regs[3],regs[1]-1, regs[2]);
			
			return ((regs[2]==dteDate.getDate()) && (regs[1]-1==dteDate.getMonth()) && (regs[3]==dteDate.getFullYear()));
		}
		else
			return false;
	} 
	else 
		return false;	
}

function RemoveinvalidCharsInJs(inputStr)
{
	var returnValue = RemoveSubString(inputStr, '\'');
	returnValue = RemoveSubString(returnValue, '"');
	return returnValue;
}

function RemoveLastSeparator(inputString, separatorChar)
{
	if(inputString.length == 0)
		return '';
	
	if(inputString.substring(inputString.length-1, inputString.length) == separatorChar)
		return inputString.substring(0, inputString.length-1);
	
	return inputString.substring(0, inputString.length);
}

//character count for a textarea field
function TextArea_CharacterCounter(charLimitLength, textAreaTag, displayTagID)
{
	var displayTag = GetElement(displayTagID);
	var charsLeft = charLimitLength - textAreaTag.value.length;
	
	if (displayTag == null || displayTag == undefined)
		return;
		
	if (textAreaTag.value.length <= charLimitLength)
	{
		displayTag.innerHTML = charsLeft;
		displayTag.style.color = '#000000';
	}
	else
	{
		displayTag.innerHTML = '<strong>'+ charsLeft.toString() +'</strong>';
		displayTag.style.color = '#FF0000';
	}
}

//character count for a textarea field and stop writing
function TextArea_CharacterCounter2(charLimitLength, textAreaTag, displayTagID)
{
	var displayTag = GetElement(displayTagID);
	var charsLeft = charLimitLength - textAreaTag.value.length;
	
	if (displayTag == null || displayTag == undefined)
		return;
		
	if (textAreaTag.value.length <= charLimitLength)
	{
		displayTag.innerHTML = charsLeft;
		displayTag.style.color = '#000000';
	}
	else
	{
		displayTag.innerHTML = '<strong>0</strong>';
		displayTag.style.color = '#FF0000';
        textAreaTag.value = textAreaTag.value.substring(0, charLimitLength);
	}
}

function CleanHTMLToClient(textToClean, ControlToStore)
{
	Resultat=""
	Resultat = textToClean.replace(/<br>/gi, String.fromCharCode(10));
	if(ControlToStore!=null && ControlToStore.value!=null)
		ControlToStore.value = Resultat;
}

function htmlEncode(s){
	return s.replace(/&(?!\w+([;\s]|$))/g, "&amp;")
.replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\"/g, "&quot;");
}

function JSEncode(s){
	return s.replace(/\"/g, "&quot;").replace(/\'/g, "\\'");
}

function EncodeHTMLToServer(textToEncode, ControlToStore)
{
	var Resultat="";
	
	for (i=0;i<textToEncode.length;i++)
	{
		numer=textToEncode.charCodeAt(i);
		if((numer==13)&&(textToEncode.charCodeAt(i+1)==10))
			{
				i++;
				Resultat += '<br>';
			}
		else
			Resultat += String.fromCharCode(numer);			
	}
	if(ControlToStore!=null && ControlToStore.value!=null)
		ControlToStore.value = Resultat;
}

function renderHTML(textToEncode)
{
	var Resultat="";
	var numer = 0;
	
	for (i=0;i<textToEncode.length;i++)
	{
		numer=textToEncode.charCodeAt(i);
		if((numer==13)&&(textToEncode.charCodeAt(i+1)==10))
			{
				i++;
				Resultat += '<br>';
			}
		else
			Resultat += String.fromCharCode(numer);			
	}
	return Resultat;
}

//barra grafica de avance
function GetPercentageBar(percentageParam, advanceParam, maxLength)
{
	var boxWidth = Math.round(percentageParam * maxLength);
	var boxPixFile = 'pixel_red';

	if (percentageParam >= 0.20)
		if (percentageParam >= 0.40)
			if (percentageParam >= 0.60)
				if (percentageParam >= 0.80)
					boxPixFile = 'pixel_blue';
				else
					boxPixFile = 'pixel_green';
			else
				boxPixFile = 'pixel_yellow';
		else
			boxPixFile = 'pixel_orange';
	
	return '<img src=\'../images/' + boxPixFile + '.gif\' style=\'border:1px solid #333333\' height=\'10\' width=\''+ boxWidth +'\' title=\''+ advanceParam +'\'>';
}

function GetSpanTitle(inputStr, titleParam)
{
	if (inputStr == undefined || inputStr == null) return inputStr;
	
	var titleValue = (titleParam == undefined) ? inputStr : titleParam;
	
	if (inputStr.length == 0)
		return inputStr;
	else
		return '<span title="'+ htmlEncode(titleValue.toString()) +'">' + htmlEncode(inputStr.toString()) + '</span>';
}

function AddComboItem(comboListObject, newItemString, newItemValue)
{
	//Agregar elementos intermedios
	var lclObjOpt 	= window.document.createElement("OPTION");
	lclObjOpt.text 	= newItemString;
	lclObjOpt.value = newItemValue;
	try {
      comboListObject.add(lclObjOpt, null);
    }
    catch(ex) {
      comboListObject.add(lclObjOpt);
    }
}

function openModal(){
if (window.showModalDialog('mensajes/modalControl.asp','dialogHeight: 160px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;'))
	document.location.href = 'controlModule.asp';
}

function OpenModalWindow(windowURL, width, height)
{
	window.showModalDialog(  windowURL, "",   "dialogHeight:"+height+"px;dialogWidth:"+width+"px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
}

function OpenWindow(target){ 
popupWin = window.open(target, "Impresiones", "toolbar=yes,directories=no,status=no,menubar=yes,width=850,height=450,left=30,top=30"); 
popupWin.focus(); 
}

function OpenWindow2(target){ 
popupWin = window.open(target, "Impresiones", "toolbar=yes,directories=no,status=no,scrollbars=yes,menubar=yes,width=750,height=450,left=30,top=30"); 
popupWin.focus(); 
}

function ClearCombo(combo)
{
	while (combo.options.length > 0)
		combo.remove(0);
}
		
function FillCombo(combo, response, defaultString, selected)
{
	ClearCombo(combo);		
	var itemArray = response.split('~');
	
	AddComboItem(combo,defaultString,"");
	
	for (thisIndex=0; thisIndex<itemArray.length; thisIndex++)
	{
		var itemValues=itemArray[thisIndex].split('|');
		
		if(itemValues.length>1)
		{
			var item=AddComboItem(combo,itemValues[1],itemValues[0]);
			if(itemValues[1]==selected)
				item.selected=true;
		}
	}
}

//cargar datos de un XML a combo
function FillComboXML(combo, responseObj, defaultString, defaultValue, selectedItem)
{
	ClearCombo(combo);		
	
	if (defaultString != null)
		AddComboItem(combo, defaultString, defaultValue);
	
	if (responseObj != null)
	{
		var xmlDocument = responseObj.responseXML;
		var parenElement = xmlDocument.documentElement;
		if (parenElement == null)
			parenElement = xmlDocument;
			
		for (loopIndex = 0; loopIndex < parenElement.childNodes.length; loopIndex++)
		{
			var itemID = parenElement.childNodes[loopIndex].firstChild.firstChild.nodeValue.toString();
			var itemName = parenElement.childNodes[loopIndex].lastChild.firstChild.nodeValue.toString()
			
			var item = AddComboItem(combo, itemName, itemID);
			if(itemID == selectedItem.toString())
				item.selected=true;
			
		}
	}
}


//cargar datos de un XML a combo. XML con estructura especifica
function FillComboXMLStructured(combo, responseObj, defaultString, defaultValue, selectedItem, catalogName)
{
	ClearCombo(combo);		
	
	if (defaultString != null)
		AddComboItem(combo, defaultString, defaultValue);
	
	if (responseObj != null)
	{
		var xmlDocument = responseObj.responseXML;
		var parenElement = xmlDocument.documentElement;
		if (parenElement == null)
			parenElement = xmlDocument;
			
		for (loopIndex = 0; loopIndex < parenElement.childNodes.length; loopIndex++)
		{
			switch(catalogName)
			{
				case 'evaluationsList':
				
					var id = GetNodeValue(parenElement.childNodes[loopIndex].childNodes[0]);
					
					var dateStart = String(GetNodeValue(parenElement.childNodes[loopIndex].childNodes[1]));
					dateStart = SpanishDateFormat(dateStart);
					
					var dateEnd = String(GetNodeValue(parenElement.childNodes[loopIndex].childNodes[2]));
					dateEnd = SpanishDateFormat(dateEnd);
					
					var item = AddComboItem(combo, dateStart + ' - ' + dateEnd, id);
					
					if (id == selectedItem)
						item.selected=true;
					break;
					
				default:
					var item = AddComboItem(combo, parenElement.childNodes[loopIndex].childNodes[0].nodeValue, parenElement.childNodes[loopIndex].getAttribute('id'));
					if(parenElement.childNodes[loopIndex].getAttribute('id') == selectedItem)
						item.selected=true;
			}
		}
	}
}

function IsValueInArray(arrayObj, value)
{
	if(arrayObj==null) return false;
	if(arrayObj=='undefined') return false;
	
	for(var i=arrayObj.length-1; i>=0; i--)
	{
		if (arrayObj[i] == value)
		{
			return true;
		}
	}
	return false;
}

function ComboBox_GetTextByValue(tagId, value)
{
	if(tagId==null) return '';
	if(tagId=='undefined') return '';
	var comboObj = GetElement(tagId);
	
	for(var i=comboObj.options.length-1; i>=0; i--)
	{
		if (comboObj.options[i].value == value)
		{
			return comboObj.options[i].text;
		}
	}
	return '';
}

function ComboBox_SearchItemByText(tagId, textValue)
{
	if(tagId==null) return -1;
	if(tagId=='undefined') return -1;
	var comboObj = GetElement(tagId);
	
	for(var i=comboObj.options.length-1; i>=0; i--)
	{
		if (comboObj.options[i].text == textValue)
		{
			return i;
		}
	}
	return -1;
}

function ComboBox_SelectByValue(tagId, value)
{
	if(tagId==null) return;
	if(tagId=='undefined') return;
	var comboObj = GetElement(tagId);
	
	for(var i=comboObj.options.length-1; i>=0; i--)
	{
		if (comboObj.options[i].value == value)
		{
			comboObj.selectedIndex = i;
			return;
		}
	}
	return;
}

//remover un renglon de un tableBody
function DeleteTableRow(tBodyID, rowID)
{
	var tbody = GetElement(tBodyID);
	
	if (tbody != undefined)
	{
		for (var rowIndex=0; rowIndex<tbody.rows.length; rowIndex++)
		{
			if(tbody.rows[rowIndex].id == rowID)
			{
				tbody.deleteRow(rowIndex);
				break;
			}
		}
	}
	
	return tbody.rows.length;
}

//-- Mostrar todos los elementos del tipo indicado
function showAllByTag(tagName,dispType) {
	var elements = document.getElementsByTagName(tagName);
	var i = 0;
	if (dispType == "") {
	        dispType = 'inline';
	}
	while (i < elements.length) {
	        elements[i].style.display = dispType;
	        i++;
	        }
}

//-- Ocultar todos los elementos del tipo indicado
function hideAllByTag(tagName) {
	var elements = document.getElementsByTagName(tagName);
	var i = 0;
	while (i < elements.length) {
	        elements[i].style.display = "none";
	        i++;
	        }
}

//Objeto para leer evento
function eventBuilder(e){
	var ev= (window.event)? window.event: e;
	if(!ev || !ev.type) return false;
	var ME= ev;
	
	if(ME.type.indexOf('key')!= -1){
		if(iz('ie') || ME.type.indexOf('keypress')!= -1){
			ME.key= (ev.keyCode)? ev.keyCode: ((ev.charCode)? ev.charCode: ev.which);
		}
		else ME.key= ev.charCode;
		if(ME.key) ME.letter= String.fromCharCode(ME.key);
	}
	return ME;
}
//... se utiliza asi:
//function handleKey(e){
//   var c = eventBuilder(e).key;
    // do something with c;
//}



//Get Element (Object) found by its ID
function GetElement(tagId) 
{
	var lclObj = document.getElementById(tagId);
	if(lclObj && lclObj.length && lclObj[0].id==tagId)
		lclObj=lclObj[0];
	return lclObj;
}



function getElementsByClassName(classname, node) {
	if(!node) node = document.getElementsByTagName("body")[0];
	var a = [];
	
	var re = new RegExp('\\b' + classname + '\\b');
	var els = node.getElementsByTagName("*");
	for(var i=0,j=els.length; i<j; i++)
		if(re.test(els[i].className))a.push(els[i]);
			return a;
}

//change display of an HTML element
function HideElement(elementID)
{
	if( isDefined(elementID) )
		GetElement(elementID).style.display = 'none';
}

//change display of an HTML element
function ShowElement(elementID)
{
	if( isDefined(elementID) )
		GetElement(elementID).style.display = '';
}

function GetNodeValue(nodeObject, nullReplaceStr)
{
	var returnValue = null;
	
	if(typeof nodeObject == 'object') {
		if (nodeObject.firstChild != null)
			returnValue =  nodeObject.firstChild.nodeValue;
		else
			returnValue = null;
	}
	
	if(nullReplaceStr != null && returnValue == null)
		returnValue = nullReplaceStr;
	
	return returnValue;
}

//obtener etiqueta de la opcion seleccionada en una lista
function GetSelectedLabel(elementID)
{
	var lclObj = GetElement(elementID);
	if (lclObj.selectedIndex == -1)
		return '';
	else
		return lclObj.options[lclObj.selectedIndex].text;
}

//-- Deshabilitar todos los controles en el modulo
function DisableAllControls(disableValue) {

	DisableControlsList(document.getElementsByTagName('input'), disableValue);
	DisableControlsList(document.getElementsByTagName('select'), disableValue);
}

//-- Deshabilitar/habilitar el control indicado
function DisableControlsList(controlsList, disableValue)
{
	var i = 0;
	while (i < controlsList.length) 
	{
		controlsList[i].disabled = disableValue;
	    i++;
	}
}

//Get form element value
function GetTagValue(elementID)  //obj,use_default,delimiter
{
	var use_default = false;
	var delimiter = ',';
	
	if (elementID == '')
		return null;
	
	var lclObj = GetElement(elementID);
	
	if(null==lclObj)
		return null;
		
	switch(lclObj.type)
	{
		case 'radio': 
		case 'checkbox': 
			return(((use_default)?lclObj.defaultChecked:lclObj.checked)?lclObj.value:null);
		
		case 'text': 
		case 'hidden': 
		case 'textarea': 
		case 'file':
			return(use_default)?lclObj.defaultValue:lclObj.value;
		
		case 'password': 
			return((use_default)?null:lclObj.value);
		case 'select-one':
			if(lclObj.options==null)
			{
				return null;
			}
			if(use_default)
			{
				var o=lclObj.options;
				for(var i=0;i<o.length;i++)
				{
					if(o[i].defaultSelected)
					{
						return o[i].value;
					}
				}
				return o[0].value;
			}
			if(lclObj.selectedIndex<0)
			{
				return null;
			}
			return (lclObj.options.length>0)?lclObj.options[lclObj.selectedIndex].value:null;
			
		case 'select-multiple':
			if(lclObj.options==null)
			{
				return null;
			}
			var values=new Array();
			for(var i=0;i<lclObj.options.length;i++)
			{
				if((use_default&&lclObj.options[i].defaultSelected)||(!use_default&&lclObj.options[i].selected))
				{
					values[values.length]=lclObj.options[i].value;
				}
			}
			
			return(values.length==0)?null:CommifyArray(values,delimiter);
	}
			
	//alert("FATAL ERROR: Field type "+lclObj.type+" is not supported for this function");
	return null;
}

//Set form element value
function SetTagValue(elementID, newValue)  //obj,use_default,delimiter
{
	if (elementID == '')
		return;
	
	var lclObj = GetElement(elementID);
	
	if(null==lclObj)
		return;
		
	switch(lclObj.type)
	{
		//case 'radio': 
		case 'checkbox': 
			lclObj.checked = (newValue == true || newValue == 'True' || newValue == '1');

		case 'text': 
		case 'hidden': 
		case 'textarea': 
			lclObj.value = newValue;
		
		case 'select-one':
			if(lclObj.options==null)
				return;
			
			var o=lclObj.options;
			for(var i=0;i<o.length;i++)
			{
				if(o[i].value == newValue)
				{
					lclObj.selectedIndex = i;
					break;
				}
			}
			
		//case 'select-multiple':
	}
		
}


function CleanFormTags(tagsArrayString)
{
	var tagsArray = tagsArrayString.split("|");
	
	for (thisIndex=0; thisIndex<tagsArray.length; thisIndex++) {
		ClearTagValue(tagsArray[thisIndex]);
	}
}

function ClearTagValue(elementID)
{
	if (elementID == '')
	return;
	
	var lclObj = GetElement(elementID);
	
	if(null==lclObj || undefined==lclObj)
		return;
		
	switch(lclObj.type)
	{
		case 'radio': 
		case 'checkbox': 
			lclObj.checked = false;
		
		case 'text': 
		case 'hidden': 
		case 'textarea': 
		case 'file':
			lclObj.value = '';
		
		case 'password': 
			lclObj.value = '';
			
		case 'select-multiple':
		case 'select-one':
			lclObj.selectedIndex = 0;
	}
}


function switchRowColor(objectID, highlightcolor)
{
		var object = document.getElementById(objectID);
		switchRowColor2(object, highlightcolor);
}

function switchRowColor2(objectRef, highlightcolor)
{
		var object = objectRef;
		if (object.tagName == "TD")
			var row = object.parentElement;
		else
			var row =object;

		if (row.highlighted == null || !row.highlighted)
		{
			row.highlighted = true;
			row.originalBackgroundColor=row.style.backgroundColor;
			row.style.backgroundColor=highlightcolor;
		}
		else
		{
			row.highlighted = false;
			row.style.backgroundColor=row.originalBackgroundColor;
		}
}

function HighlightRow()
{
	this.originalBackgroundColor = this.style.backgroundColor;
	this.style.backgroundColor = this.bgcolor2;
}
function UnHighlightRow()
{
	this.style.backgroundColor = this.originalBackgroundColor;
}

function highlightRowColor(objectID, highlightcolor, on)
{
		var object = document.getElementById(objectID);
		if (object.tagName == "TD")
			var row = object.parentElement;
		else
			var row =object;

		if (!row.highlighted)
		{
			if(on)
				switchRowColor(objectID, highlightcolor);
		}
		else
		{
			if(!on)
				switchRowColor(objectID, highlightcolor);
		}
}

//Hides and appears description and editable elements, to activate EDIT pages.
function switchEditElements(startEdit){	
	var elements = document.getElementsByTagName('span')
	if (startEdit)
	{
		for (i=0; i<elements.length; i++){
			if(elements[i].id.substr(0,5)=='show_')
				elements[i].style.display = 'none';
			if(elements[i].id.substr(0,5)=='edit_')
				elements[i].style.display = '';
		}
	}
	else
	{
		for (i=0; i<elements.length; i++){
			if(elements[i].id.substr(0,5)=='edit_')
				elements[i].style.display = 'none';
			if(elements[i].id.substr(0,5)=='show_')
				elements[i].style.display = '';
		}
	}
}

function ValidateControlLength(controlID, maxLengthParam)
{
	var controlTag = GetElement(controlID);
	var maxLength = (maxLengthParam==undefined || maxLengthParam==null) ? controlTag.getAttribute('maxlength')*1 : maxLengthParam;
	
	if (isNaN(maxLength))
		return false;
	else
		return (controlTag.value.length <= maxLength);
}

function validateRequiredFields(formID)
{
	var inputObjList = getElementsByClassName('inputRed', GetElement(formID));
	for (var i=0; i<inputObjList.length; i++) 
		UpdateFieldStatus(inputObjList[i], '');
}
		
function UpdateFieldStatus(formInput, emptyValue)
{
	inputValue = trim(GetTagValue(formInput.id));
	if( inputValue == emptyValue )
		formInput.className = ReplaceSubString(formInput.className, 'inputGreen', 'inputRed');
	else
		formInput.className = ReplaceSubString(formInput.className, 'inputRed', 'inputGreen');
		
	if (formInput.previousSibling != undefined && formInput.previousSibling.className == 'inputRedMark')
	{
		if( inputValue == emptyValue )
			formInput.previousSibling.src = '/images/pixel_red.gif';
		else
			formInput.previousSibling.src = '/images/pixel_green.gif';
		
		formInput.previousSibling.className = 'inputRedMark';
	}
}


//function onKeyPress(DataType, MinValue, MaxValue){
function onKeyPress(event,DataType){
	//TAB and ENTER keys accepted
	if (event.keyCode==9||event.keyCode==13)
		return;
	// . accepted
	if (DataType==3&&event.keyCode==46)
		return;
	// - accepted
	if (DataType==3&&event.keyCode==45)
		return;
	// _ accepted
	if (DataType==6&&event.keyCode==46)
		return;
	//digits accepted
	if (event.keyCode<48 || event.keyCode>57)
		event.returnValue = false;
}

function onKeyDown(event, element){
	//switch enter for tab
	if (event.keyCode==13)
		event.keyCode=9;
}

function RestrictUserInput(elem) {
    if (/[^\d]/g.test(elem.value))
       elem.value = elem.value.replace(/[^\d]/g, '');
}


function trim(stringToTrim)
{
	if (stringToTrim == undefined || stringToTrim == null)
		return null;
	
	return stringToTrim.replace(/^\s+/,'');
} 


//get DateFormat in different formats
function formatDateString(dateParam, idFormat)
{
	if(dateParam==null || dateParam==undefined) {return '';}
	
	if (dateParam == '')
	{
		return (dateParam);
	}
	else
	{
		var paramArray = dateParam.split("/");
		var month = paramArray[1];
		var day = paramArray[0];
		var year = paramArray[2];
		
		//completar con 0s
		if(day.length == 1)
			day = '0' + day.toString();
		
		if(month.length == 1)
			month = '0' + month.toString();
		
		switch(idFormat)
		{
			case 3: //YYYYMMDD
				return (year + month + day);
				break;
		}
	}
}

//get DateFormat in english - month/day/year, considering parameter comes in spanish - day/month/year
function EnglishDateFormat(dateParam)
{
	if(dateParam==null || dateParam==undefined) {return '';}
	
	if (dateParam == '')
	{
		return (dateParam);
	}
	else
	{
		var paramArray = dateParam.split("/");
		var month = paramArray[1];
		var day = paramArray[0];
		var year = paramArray[2];
		
		//completar con 0s
		if(day.length == 1)
			day = '0' + day.toString();
		
		if(month.length == 1)
			month = '0' + month.toString();
		
		return (month + "/" + day + "/" + year);
	}
}

//get DateFormat in spanish - day/month/year, considering parameter comes in english
function SpanishDateFormat(dateParam)
{
	if(dateParam==null || dateParam==undefined) {return '';}
		
	//remover hora
	var noHourArray	= dateParam.split(" ");
	
	if (noHourArray[0] == '')
	{
		return (dateParam);
	}
	else
	{
		var paramArray = noHourArray[0].split("/");
		var month = paramArray[0];
		var day = paramArray[1];
		var year = paramArray[2];
		
		//completar con 0s
		if(day*1 < 10)
			day = '0' + day.toString();
		
		if(month*1 < 10)
			month = '0' + month.toString();
		
		return (day + "/" + month + "/" + year);
	}
}

function GetDateObject(dateParam, formatDate)
{
	var newDateObj 	= new Date();
	var paramArray 	= dateParam.split("/");
	
	newDateObj.setFullYear(paramArray[2]*1, paramArray[0]*1-1, paramArray[1]*1);
	
	if (formatDate != undefined)
	{
		switch(formatDate)
		{
			case 1:
				//year - month - day
				return newDateObj.getFullYear().toString() + '/' + newDateObj.getDate().getMonth() + '/' + newDateObj.getDate().toString();
				break;
			default:
				return newDateObj;
				break;
		}
	}
	return newDateObj;
}

//input date as yyyy/mm/dd, and number of days to add
//output yyyy/mm/dd
function AddDays(strDate,iDays)
{
strDate = Date.parse(strDate);
strDate = parseInt(strDate, 10);
strDate = strDate + iDays*(24*60*60*1000);
strDate = new Date(strDate);
returnValue = strDate.getFullYear() + "/" + strDate.getMonth()*1+1 + "/" + strDate.getDate();

return returnValue;
}

function DateAdd(interval,n,dt)
{
	if(!interval||!n||!dt) return;	
	
	var s=1,m=1,h=1,dd=1,i=interval;
	
	if(i=='month'||i=='year')
	{
		dt=new Date(dt);
		if(i=='month')
		{
			newMonth = dt.getMonth() + n;
			if (newMonth > 11)
				newMonth = newMonth - 12;
			dt.setMonth(dt.getMonth()+n);
			if(dt.getMonth() != newMonth)
			{
				//overshot month due to date, so go to last day of previous month 
				dt.setDate(0);
			}
		}
		if(i=='year') dt.setFullYear(dt.getFullYear()+n);		
	}
	else if (i=='second'||i=='minute'||i=='hour'||i=='day'){
		dt=Date.parse(dt);
		if(isNaN(dt)) return;
		if(i=='second') s=n;
		if(i=='minute'){s=60;m=n}
		if(i=='hour'){s=60;m=60;h=n};
		if(i=='day'){s=60;m=60;h=24;dd=n};
		dt+=((((1000*s)*m)*h)*dd);
		dt=new Date(dt);
	}
	return dt;
}

//-------------- new Validation Functions

//Validates a value against a regular expression and limit values
function ValueValidator(stringRegExp, inputValue, selectionRange, OldValue, DecimalSeparator, MinValue, MaxValue){
	var validation = true;
	var objRegExp = new RegExp(stringRegExp.toString());
	var endIndex = selectionRange.text.length;
	while(selectionRange.expand("character")){}
	var startIndex = OldValue.length - selectionRange.text.length;
	endIndex = endIndex + startIndex;
	var newValueCandidate = OldValue.substring(0,startIndex) + inputValue + OldValue.substring(endIndex,OldValue.length);
	if(!objRegExp.test(newValueCandidate))
		return false;
	else
	{
		//Validate for MinValue and MaxValue
		if (MinValue!=null || MaxValue!=null)
		{
			var CompareNum = parseFloat(newValueCandidate.replace(DecimalSeparator,'.'));
			if (MinValue!=null && CompareNum < MinValue) validation = false;
			if (MaxValue!=null && CompareNum > MaxValue) validation = false;
			return validation;
		}
		else
			return true;
	}
}

//Validates Currency Values
function ValidateCurrency(InputObj, minvalue, maxvalue){
	var key = window.event.keyCode;
	var selectionRange = document.selection.createRange ();
	var TextBoxValue = InputObj.value;
	
	var DecimalSeparator = '.';
	var DecimalDigits = 2;
	var MinValue = minvalue;
	var MaxValue = maxvalue;
	var NegativeSign = "";
	if (MinValue < 0) var NegativeSign = "-?";
	
	var stringRegExp = "(^"+NegativeSign+"\\d*\\"+DecimalSeparator+"?\\d{0,"+DecimalDigits+"}$)";
	if (!ValueValidator(stringRegExp, String.fromCharCode(key), selectionRange, TextBoxValue, DecimalSeparator, MinValue, MaxValue))
		event.keyCode = 0;	
}

//Validates Clipboard against a regular expression to allow or deny pasting it in the TextBox
function ValidateClipboard(InputObj, minvalue, maxvalue, DataType){
//	var Clipboard = igtbl_trim(window.clipboardData.getData("Text"));
	var ClipBoard = window.clipboardData.getData("Text");
	if (!event || Clipboard.length==0 || Clipboard == undefined) return;
	
	var selectionRange = document.selection.createRange ();
	var TextBoxValue = this.value;
		
	var MinValue = minvalue;
	var MaxValue = maxvalue;
		var NegativeSign = "";
	// Allow for negative sign to be approved by Regular Expression
	if (MinValue < 0) var NegativeSign ="-?";
	
	switch(DataType){
		case 3: //Currency
			var stringRegExp = "(^-?\\d*\\.?\\d{0,2}$)";
			if (!ValueValidator(stringRegExp, Clipboard, selectionRange, TextBoxValue, Culture.CurrencyDecimalSeparator, MinValue, MaxValue))
				event.returnValue = false;
			break;
	}	
	return;
}

function Format(num,decimaldigits,symbol,dot,groupseparator,groupdigits, percentFormat){
		var lclDecimalDigits = 1;
		for (var i=0; i<decimaldigits; i++)
			lclDecimalDigits = lclDecimalDigits * 10;
		var RegularExp = new RegExp("/\\" + symbol + "|\\" + groupseparator + "/g");
		num = num.toString().replace(RegularExp,'');
		num = num.replace(dot,'.');	//replace decimal separator with '.' for Math class properties
		if(isNaN(num)) return "0";
		sign = (num == (num = Math.abs(num)));
		num = Math.floor(num*lclDecimalDigits+0.50000000001);
		cents = num%lclDecimalDigits;
		centsStr = cents.toString();
		num = Math.floor(num/lclDecimalDigits).toString();
		for (var i=0; i<decimaldigits-centsStr.length; i++)
			centsStr = "0" + centsStr;
		cents = (decimaldigits > 0) ? centsStr : "";
		for (var i = 0; i < Math.floor((num.length-(1+i))/groupdigits); i++)
			num = num.substring(0,num.length-((groupdigits+1)*i+groupdigits))+groupseparator+num.substring(num.length-((groupdigits+1)*i+groupdigits));
		if (percentFormat)
			if (cents != "")
				return (((sign)?'':'-') + num + dot + cents + symbol);
			else
				return (((sign)?'':'-') + num + symbol);
		else
			if (cents != "")
				return (((sign)?'':'-') + symbol + num + dot + cents);
			else
				return (((sign)?'':'-') + symbol + num);
}

function FormatCurrency(InputObj){
	InputObj.value = Format(InputObj.value, 2, '$', '.' , ',', 3, false);
}

function CleanFormat(InputObj,DataType){
	if (InputObj.value.substring(0,1)=='$')
		InputObj.value = InputObj.value.substring(1,InputObj.value.length);
	InputObj.value = InputObj.value.replace(',','');
}
//------------Functions

function Navigate(Module){
eval("document.location='" + Module + "'");
}

function ParentNavigate(Module){
eval("window.parent.location='" + Module + "'");
}

// Array Remove - By John Resig (MIT Licensed)
// http://ejohn.org/blog/javascript-array-remove/
Array.prototype.remove = function(from, to) {
  var rest = this.slice((to || from) + 1 || this.length);
  this.length = from < 0 ? this.length + from : from;
  return this.push.apply(this, rest);
};

function replaceAll(oldStr,findStr,repStr) {
  var srchNdx = 0;  // srchNdx will keep track of where in the whole line
                    // of oldStr are we searching.
  var newStr = "";  // newStr will hold the altered version of oldStr.
  while (oldStr.indexOf(findStr,srchNdx) != -1)  
                    // As long as there are strings to replace, this loop
                    // will run. 
  {
    newStr += oldStr.substring(srchNdx,oldStr.indexOf(findStr,srchNdx));
                    // Put it all the unaltered text from one findStr to
                    // the next findStr into newStr.
    newStr += repStr;
                    // Instead of putting the old string, put in the
                    // new string instead. 
    srchNdx = (oldStr.indexOf(findStr,srchNdx) + findStr.length);
                    // Now jump to the next chunk of text till the next findStr.           
  }
  newStr += oldStr.substring(srchNdx,oldStr.length);
                    // Put whatever's left into newStr.             
  return newStr;
}

//format a String to be used as Regular Expression
function FormatRegExpString(convertString)
{
	convertString = convertString.replace(/\\/g,'\\\\');
	convertString = convertString.replace(/\//g,'\\/');
	convertString = convertString.replace(/\*/g,'\\*');
	convertString = convertString.replace(/\$/g,'\\$');
	convertString = convertString.replace(/\+/g,'\\+');
	convertString = convertString.replace(/\?/g,'\\?');
	convertString = convertString.replace(/\./g,'\\.');
	convertString = convertString.replace(/\^/g,'\\^');
	convertString = convertString.replace(/\[/g,'\\[');
	convertString = convertString.replace(/\]/g,'\\]');
	convertString = convertString.replace(/\(/g,'\\(');
	convertString = convertString.replace(/\)/g,'\\)');
	convertString = convertString.replace(/\{/g,'\\{');
	convertString = convertString.replace(/\}/g,'\\}');
	convertString = convertString.replace(/\|/g,'\\|');				
	
	return convertString;
	
}   

//Remove substring
function RemoveSubString(thisString,thisSubString){
	eval('var re = /' + FormatRegExpString(thisSubString) + '/g');
	return thisString.replace(re, "");
}



//Replace substring
function ReplaceSubString(thisString,thisSubString,newSubString){
	eval('var re = /' + FormatRegExpString(thisSubString) + '/g');
	return thisString.replace(re, newSubString);
}

//Obtener un numero al azar entre cero y cinco mil
function get_random()
{
    var ranNum= Math.floor(Math.random()*5000);
    return ranNum;
}

//obtener la posicion de un objeto
function GetObjectPosition(lclObj) {
	var curleft = curtop = 0;
	if (lclObj.offsetParent) {
		curleft = lclObj.offsetLeft;
		curtop = lclObj.offsetTop;
		while (lclObj = lclObj.offsetParent) {
			curleft += lclObj.offsetLeft
			curtop += lclObj.offsetTop
		}
	}
	return [curleft,curtop];
}

function GetWindowClientWidth() {
	return f_filterResults (
		window.innerWidth ? window.innerWidth : 0,
		document.documentElement ? document.documentElement.clientWidth : 0,
		document.body ? document.body.clientWidth : 0
	);
}

function GetWindowClientHeight() {
	return f_filterResults (
		window.innerHeight ? window.innerHeight : 0,
		document.documentElement ? document.documentElement.clientHeight : 0,
		document.body ? document.body.clientHeight : 0
	);
}

function GetWindowScrollLeft() {
	return f_filterResults (
		window.pageXOffset ? window.pageXOffset : 0,
		document.documentElement ? document.documentElement.scrollLeft : 0,
		document.body ? document.body.scrollLeft : 0
	);
}

function GetWindowScrollTop() {
	return f_filterResults (
		window.pageYOffset ? window.pageYOffset : 0,
		document.documentElement ? document.documentElement.scrollTop : 0,
		document.body ? document.body.scrollTop : 0
	);
}

function f_filterResults(n_win, n_docel, n_body) {
	var n_result = n_win ? n_win : 0;
	if (n_docel && (!n_result || (n_result > n_docel)))
		n_result = n_docel;
	return n_body && (!n_result || (n_result > n_body)) ? n_body : n_result;
}

//--------------- COVER LAYERS

function ShowConnectionMessage(showValue, useAllScreen)
{
	if (useAllScreen)
	{
		ShowCoverLayer(showValue);
		ShowMessage(showValue);
	}
	else
	{
		ShowConfirmMessage(showValue, 'Estableciendo conexión con servidor...');
	}
}

function ShowTopMessage(showValue, messageStr)
{
	var messageDiv = GetElement('topMessageDIV');
	var messageTag = GetElement('topMessageTAG');

	if (messageDiv != undefined && messageTag != undefined)
	{
		messageTag = messageStr;
		
		if(showValue)
			slidedown(messageDiv);
		else
			slideup(messageDiv);
	}
}

function SetTopMessage(setConfirmIcon, autoClose, messageStr)
{
	var messageDiv = GetElement('topMessageDIV');
	var messageTag = GetElement('topMessageTAG');

	if (messageDiv != undefined && messageTag != undefined)
	{
		//if(setConfirmIcon)
		// -- falta: cambiar icono a confirmacion
		
		messageTag = messageStr;
		
		//Cerrar mensaje luego de 5 segundos
		if(autoClose)
			setTimeout('ShowTopMessage(false, null)',5000);
	}
}

function ShowFullScreenMessage(showValue, messageStr)
{
	ShowCoverLayer(showValue);
	if (messageStr != '')
		document.getElementById("fulScreenMessageString").innerHTML = messageStr;
		
	ShowMessage(showValue);
	
	//regresar valor original
	if (!showValue)
		document.getElementById("fulScreenMessageString").innerHTML = 'Estableciendo conexión con servidor...';
}

// Cambiar mensaje de la ventana de conexion a servidor
function SetConnectionMessage(messageString)
{
	GetElement('fulScreenMessageString').innerHTML = messageString;
}

function ShowCoverLayer(showValue){
	docu	= document.body;
	coscura	= document.getElementById('iCoverLayer');
	coscura.style.height=docu.scrollHeight+'px';
	if(showValue) 
		coscura.style.visibility = "visible";
	else 
		coscura.style.visibility = "hidden";
}

function ShowMessage(showValue) {
	var tableObj = document.getElementById("iMessageTable");
	var objLayer = document.getElementById("iMessageLayer");
	
	if(showValue) 
		objLayer.style.display = '';
	else 
		objLayer.style.display = 'none';
	
	var left = ((document.body.clientWidth - tableObj.width) / 2) + GetWindowScrollLeft();
	var top  = ((document.body.clientHeight - tableObj.height) / 2) + GetWindowScrollTop();
		
	objLayer.style.left = parseInt(left);
	objLayer.style.top 	= parseInt(top);
}

//TEMPORAL
function ShowLocalMessage(messageDIVid, showValue) {
	var objLayer = GetElement(messageDIVid);
	
	if(showValue) 
		objLayer.style.display = '';
	else 
		objLayer.style.display = "none";	
		
	var left = ((document.body.clientWidth - objLayer.offsetWidth) / 2) + GetWindowScrollLeft();
	var top  = ((document.body.clientHeight - objLayer.offsetHeight) / 2) + GetWindowScrollTop();
		
	objLayer.style.left = parseInt(left);
	objLayer.style.top 	= parseInt(top);
}

function ShowConfirmMessage(showValue, messageStr)
{
	if (messageStr != '')
		document.getElementById("TopMessageString").innerHTML = messageStr;
	
	if(showValue) 
		document.getElementById("TopMessageTable").style.display = '';
	else 
		document.getElementById("TopMessageTable").style.display = 'none';
}

function ShowConnectionTag(idTag)
{
	var tagObject = GetElement(idTag);
	if (tagObject != undefined)
		tagObject.style.display = '';
}

function CloseConnectionTag(idTag, timeInterval)
{
	var tagObject = GetElement(idTag);
	if (tagObject != undefined)
	{
		if(timeInterval == 0)
			tagObject.style.display = 'none';
		else
			setTimeout("CloseConnectionTag('"+ idTag +"', 0)", timeInterval);	
	}
}

function SetConnectionTag(idTag, messageType, messageStr)
{
	var tagObject = GetElement(idTag);
	if (tagObject != undefined)
	{
		switch(messageType)
		{
			case 'confirm':
				if (messageStr == '')
					messageStr = 'Comando ejecutado satisfactoriamente';
					
				tagObject.innerHTML = '<img src="../images/iconos/checked_green.gif" width="14" height="12" hspace="5">' + messageStr;
				break;
				
			case 'wait':
				if (messageStr == '')
					messageStr = 'Conectando a servidor...';
					
				tagObject.innerHTML = '<img src="../images/iconos/antenna.gif" width="14" height="12" hspace="5">' + messageStr;
				break;
			
			case 'stop':
				if (messageStr == '')
					messageStr = 'Atención';
					
				tagObject.innerHTML = '<img src="../images/iconos/stop.gif" width="14" height="12" hspace="5">' + messageStr;
				break;
			
			case 'alert':
				if (messageStr == '')
					messageStr = 'Atención';
					
				tagObject.innerHTML = '<img src="../images/iconos/alert2.gif" width="16" height="12" hspace="5">' + messageStr;
				break;
		}
	}
}


/* hides TAG objects under specified floating DIV */
function hideUnderneathElements( elmID, overDiv )
{
	var obj, objLeft, objTop, objParent, objLeft, objTop, objParent, objHeight, objWidth;

    for( i = 0; i < document.all.tags( elmID ).length; i++ )
    {
      obj = document.all.tags( elmID )[i];
      if( !obj || !obj.offsetParent || obj.parentControl == overDiv.id)
      {
        continue;
      }
  
      // Find the element's offsetTop and offsetLeft relative to the BODY tag.
      objLeft   = obj.offsetLeft;
      objTop    = obj.offsetTop;
      objParent = obj.offsetParent;
      
      while( objParent.tagName.toUpperCase() != "BODY" )
      {
        objLeft  += objParent.offsetLeft;
        objTop   += objParent.offsetTop;
        objParent = objParent.offsetParent;
      }
  
      objHeight = obj.offsetHeight;
      objWidth = obj.offsetWidth;
  
      if(( overDiv.offsetLeft + overDiv.offsetWidth ) <= objLeft );
      else if(( overDiv.offsetTop + overDiv.offsetHeight ) <= objTop );
      else if( overDiv.offsetTop >= ( objTop + objHeight ));
      else if( overDiv.offsetLeft >= ( objLeft + objWidth ));
      else
      {
        obj.style.visibility = "hidden";
      }
    }
}

/*
* unhides <select> and <applet> objects (for IE only)
*/
function showHiddenElementByTag( elmID )
{
	var obj;
	
    for( i = 0; i < document.all.tags( elmID ).length; i++ )
    {
      obj = document.all.tags( elmID )[i];
      
      if( !obj || !obj.offsetParent )
      {
        continue;
      }
    
      obj.style.visibility = "";
    }
}


function ChangeMsgDisplay(idActive,idCorrect, idError, visActive, visCorrect, visError)
{
	var activeMsgID 	= (idActive == '') ? 'connectionMsg' : idActive;
	var correctMsgID 	= (idActive == '') ? 'connectionMsgEnd' : idCorrect;
	var errorMsgID 		= (idActive == '') ? 'connectionMsgEndError' : idError;
	
	GetElement(activeMsgID).style.display 	= (visActive == 1) ? '' : 'none';
	GetElement(correctMsgID).style.display 	= (visCorrect == 1) ? '' : 'none';
	GetElement(errorMsgID).style.display 	= (visError == 1) ? '' : 'none';
	
}


function alpha(e) {
var k;
document.all ? k = e.keyCode : k = e.which;
return ((k > 64 && k < 91) || (k > 96 && k < 123) || k == 8  || k == 9 || k==32);
}

function userinput(e) {
var k;
document.all ? k = e.keyCode : k = e.which;
//alert(k);
return ((k > 64 && k < 91) || (k > 96 && k < 123) || k == 8  || k == 9 || (k>47 && k<58) || k==95 || k==46);

}



function Calendar_get_daysofmonth(monthNo, p_year) {
 

   DOMonth = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
   lDOMonth = new Array(31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
 
	
 
	if ((p_year % 4) == 0) {
		if ((p_year % 100) == 0 && (p_year % 400) != 0)
			return DOMonth[monthNo];
	
		return lDOMonth[monthNo];
	} else
		return DOMonth[monthNo];
}


function filldays(cmbday,cmbmonth,cmbyear)
  {
 	
	
 var x;
 var i;
 
 
if (cmbmonth.selectedIndex > 0 && cmbyear.selectedIndex>0)
{
	cmbday.disabled=false;
	//form.dateday.options.lenght= 0;
	cmbday.options.length=0;

	
	var dropdownIndex = cmbmonth.selectedIndex;
    var dropdownValue = cmbmonth[dropdownIndex].value;
	
	var dropdownIndex1 = cmbyear.selectedIndex;
    var dropdownValue1 = cmbyear[dropdownIndex1].value;
	
	
	//alert(dropdownValue)
	//var item=0;
	var days = Calendar_get_daysofmonth(dropdownValue-1,dropdownValue1);
	 
	
	var opt1 = document.createElement("option");
	cmbday.options.add(opt1);
	opt1.text = "dia";
	opt1.value = 0;
	


	for ( i = 0; i < days ; i++) {
	
	var item = i+1;
	var opt = document.createElement("option");
	cmbday.options.add(opt);
	opt.text = i+1;
	opt.value = i+1;
	
	}

}
else
{
cmbday.disabled=true;
}
  }


function cbeGeturlArguments() { 
var idx = location.href.indexOf('?'); 
var params = new Array(); 
if (idx != -1) { 
var pairs = location.href.substring(idx+1, location.href.length).split('&'); 
for (var i=0; i<pairs.length; i++) { 
nameVal = pairs[i].split('='); 
params[i] = nameVal[1]; 
params[nameVal[0]] = nameVal[1]; 
} 
} 
return params; 
} 



function setCookie(cookieName,cookieValue,expires,path,domain,secure) { 
cookieValue = encrypt(cookieValue);
document.cookie=
escape(cookieName)+'='+escape(cookieValue) 
+(expires?'; EXPIRES='+expires:'') 
+(path?'; PATH='+path:'') 
+(domain?'; DOMAIN='+domain:'') 
+(secure?'; SECURE':''); 
} 


function getCookie(cookieName) { 
var cookieValue=null; 
var posName=document.cookie.indexOf(escape(cookieName)+'='); 
if (posName!=-1) { 
var posValue=posName+(escape(cookieName)+'=').length; 
var endPos=document.cookie.indexOf(';',posValue); 
if (endPos!=-1) cookieValue=unescape(document.cookie.substring(posValue,endPos)); 
else cookieValue=unescape(document.cookie.substring(posValue)); 
} 
return cookieValue; 
}


function encrypt(text) 
{
    var textEncrypted = "";

    for (i = 1 ; i < (text.length + 1); i++) 
    {
        k = text.charCodeAt(i-1);        
        if (k >= 97 && k <= 109) 
        {
            k = k + 13;
        } 
        else
            if (k >= 110 && k <= 122) 
            {
                k = k - 13;
            } 
            else
                if (k >= 65 && k <= 77) 
                {
                    k = k + 13;
                } 
                else
                    if (k >= 78 && k <= 90) 
                    {
                        k = k - 13;
                    }

        textEncrypted = textEncrypted + String.fromCharCode(k);
    }

    return textEncrypted;
}


function killCookies()
{
    setCookie('validUserPortafolio', '', '1/1/1970', '/');
    setCookie('tPortafolio', '', '1/1/1970', '/');
    setCookie('manPortafolio', '0', '1/1/1970', '/');

    var now = new Date();
    now.setMinutes(now.getMinutes() + 30);
    setCookie('noaccessPortafolio', '1', now.toUTCString(), '/');
}

function fnLogOut()
{
    setCookie('validUserPortafolio', '', '1/1/1970', '/');
    setCookie('tPortafolio', '', '1/1/1970', '/');
    setCookie('manPortafolio', '0', '1/1/1970', '/');

    var now = new Date();
    now.setMinutes(now.getMinutes() + 30);
    setCookie('noaccessPortafolio', '1', now.toUTCString(), '/');

    //location.reload(true);
    window.location.href = "http://www.stilo.com.mx/portafolio/login.asp";
}


-->