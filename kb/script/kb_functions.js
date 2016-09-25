var g_FILTER_PAGE = 0
var g_FILTER_SORT = 1
var g_FILTER_CATEGORY = 2
var g_FILTER_SOFTWARE = 3
var g_FILTER_AUTHOR = 4

PreFetchImages();

/*------------------------------------------------------------------------
	Name: 		PreFetchImages()
	Purpose: 	cache mouse-over images
Modifications:
	Date:		Name:	Description:
	1/1/03		JEA		Creation
------------------------------------------------------------------------*/
function PreFetchImages() {
	var sImageName;
	var sImagePath;
	var re = /\.gif/gi;
	var aImages = new Array(document.images.length);
	for (var x = 0; x < document.images.length; x++) {
		sImagePath = document.images[x].src;
		sImageName = sImagePath.substring(sImagePath.lastIndexOf("/") + 1, sImagePath.length);
		if (sImageName.substr(0,4) == "btn_" && sImageName.indexOf("_on.gif") == -1) {
			aImages[x] = new Image();
			aImages[x].src = sImagePath.replace(re, "_on.gif");
		}
	}
}

/*------------------------------------------------------------------------
	Name: 		GetClientTime()
	Purpose: 	get SQL allowed format of client's time
	Return:		string (date)
Modifications:
	Date:		Name:	Description:
	12/30/02	JEA		Creation
	1/31/03		JEA		Accomodate Mozilla bug in .getYear()
------------------------------------------------------------------------*/
function GetClientTime() {
	var oDate = new Date();
	var lHour = oDate.getHours();
	var lMonth = oDate.getMonth() + 1;			// getMonth is 0-based
	var lYear = oDate.getYear() + "";
	var sAMPM = (lHour >= 12) ? " PM" : " AM";
	var sTime = (lHour > 12) ? lHour - 12 : lHour;
	sTime += ":" + PadNumber(oDate.getMinutes(), 2) + ":" + PadNumber(oDate.getSeconds(), 2) + sAMPM;
	lYear = lYear.substr(lYear.length - 2, 2);	// mozilla was returning "103" for the year 2003
	return lMonth + "/" + oDate.getDate() + "/" + lYear + " " + sTime;
}

/*------------------------------------------------------------------------
	Name: 		PadNumber()
	Purpose: 	pad number with leading zeros to match given length
Modifications:
	Date:		Name:	Description:
	12/30/02	JEA		Creation
------------------------------------------------------------------------*/
function PadNumber(v_lNumber, v_lLength) {
	v_lNumber += "";
	var lShort = v_lLength - v_lNumber.length;
	if (lShort > 0) { for (x = 0; x < lShort; x++) { v_lNumber = "0" + v_lNumber; } }
	return v_lNumber;
}

/*------------------------------------------------------------------------
	Name: 		ClearSelection()
	Purpose: 	clear selections from option list
Modifications:
	Date:		Name:	Description:
	1/5/03		JEA		Creation
------------------------------------------------------------------------*/
function ClearSelection(v_sFieldName) {
	var oField = eval("m_oForm." + v_sFieldName);
	for (var x = 0; x < oField.options.length; x++) {
		oField.options[x].selected = false;
	}
}

/*------------------------------------------------------------------------
	Name: 		SelectCategories()
	Purpose: 	pre-select categories in list
Modifications:
	Date:		Name:	Description:
	1/5/03		JEA		Creation
------------------------------------------------------------------------*/
function SelectCategories() {
	var sCats = m_oForm.fldCategoryList.value;
	if (sCats != "") {
		var aCats = sCats.split(",");
		var oCatList = m_oForm.fldCategories;
		for (var x = 0; x < aCats.length; x++) {
			for (var y = 0; y < oCatList.options.length; y++) {
				if (oCatList.options[y].value == aCats[x]) {
					oCatList.options[y].selected = true;
				}
			}
		}
	}
}

function newFilter(r_oField, v_sItemType, v_lFilter, v_aFilter) {
	v_aFilter[v_lFilter] = r_oField.options[r_oField.selectedIndex].value;
	location.href = "kb_" + v_sItemType + ".asp?sort=" + v_aFilter[g_FILTER_SORT]
		+ "&cat=" + v_aFilter[g_FILTER_CATEGORY]
		+ "&sw=" + v_aFilter[g_FILTER_SOFTWARE]
		+ "&author=" + v_aFilter[g_FILTER_AUTHOR];
}