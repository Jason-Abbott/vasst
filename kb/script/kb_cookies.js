// read a particular cookie (jea:7/31/00)
// improved regexp (jea:7/13/01)
// returns string --------------------------------------------------------
function getCookie(sName) {
	//re = new RegExp(sName + "=(\\S+)[\\;\\b]");
	re = new RegExp(sName + "=([^;]+)");
	if (re.test(unescape(document.cookie))) {
		var aMatch = re.exec(unescape(document.cookie));
		return(aMatch[1]);
	}
	return "";
}
// save a cookie (jea:8/14/00)
// finally added expiration code (jea:7/13/01)
// updates collection ----------------------------------------------------
function setCookie(sName, sValue, sExpire) {
	if (sExpire != "") {
		var oDate = new Date(sExpire);
		document.cookie = sName + "=" + escape(sValue) + "; expires=" + oDate.toGMTString();
	} else {
		// write cookie without expiration
		document.cookie = sName + "=" + escape(sValue);
	}
}
// erase pharmacy cookies (jea:8/22/00)
// 	VBS not available in .js so can't use constants
// updates collection ----------------------------------------------------
function delRxCookies() {
	var aCookies = ["WID","WFROM","PT","PN","PATIENTID","PAY","STORE","RID","OID","GID","FLOW"];
	for (var x in aCookies) {
		delCookie(aCookies[x]);
	}
}
// erase cookie (jea:8/22/00)
// updates collection ----------------------------------------------------
function delCookie(sName) {
    document.cookie = sName + "=null; expires=Thu, 01-Jan-80 00:00:01 GMT";
}
// read query string value (jea:8/29/00)
// returns string --------------------------------------------------------
function getQS(sName) {
// [&\\b]
	re = new RegExp("[&?]" + sName + "=(\\w+)\\b");
	if (re.test(unescape(location.search))) {
		var aMatch = re.exec(unescape(location.search));
		return(aMatch[1]);
	}
	return "";
}
// erase store cookies (jea:2/5/01)
// updated (jea:5/20/02)
// updates collection ----------------------------------------------------
function delStoreCookies() {
	var aCookies = ["DA","SD","ST","tempSD","tempDA"];
	var re;
	for (var sName in aCookies) {
		re = new RegExp(sName + "=([^;]+)");
		if (re.test(unescape(document.cookie))) { delCookie(aCookies[sName]); }
	}
}