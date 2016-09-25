/*------------------------------------------------------------------------
	Name: 		field()
	Purpose: 	field object constructor
	Return:		collection object
Modifications:
	Date:		Name:	Description:
	7/24/00		JEA		Creation
------------------------------------------------------------------------*/
function field(desc,type,req) {
	this.desc = desc;
	this.type = type;
	this.req = req;
}

/*------------------------------------------------------------------------
	Name: 		IsValid()
	Purpose: 	master validation function
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	7/27/00		JEA		Creation
	7/28/04		JEA		Special URL handling
------------------------------------------------------------------------*/
function IsValid(v_sForm, r_oFields) {
	var oForm = eval("document." + v_sForm);
	var oErrors = new Array;
	var bNeedToCheck;
	var lCount = 0;
	
	for (var i in r_oFields) {
		// pass every field to the appropriate validation function
		var oField = eval("oForm." + i);
		
		if (r_oFields[i].type == "URL" && oField.value == "http://") {
			bNeedToCheck = false;
		} else {
			bNeedToCheck = (r_oFields[i].req || oField.value.length > 0);
		}
		
		if (bNeedToCheck) {
			if (!(eval("is" + r_oFields[i].type + "(oField)"))) {
				oErrors[lCount] = r_oFields[i].desc;
				lCount++;
			}
		}
	}
	if (lCount > 0) {
		// must be some errors
		var sMessage = "Please enter valid\n";
		for (i = 0; i < lCount; i++) {
			sMessage += "- " + oErrors[i] + "\n";
		}
		alert(sMessage);
		return false;
	} else {
		return true;
	}
}

/*------------------------------------------------------------------------
	Name: 		isSelect()
	Purpose: 	make sure they made some selection
				this assumes that layout options, like lines, are <= 0
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	7/27/00		JEA		Creation
	8/6/02		JEA		Check field type
	1/8/03		JEA		allow multiple selects
------------------------------------------------------------------------*/
function isSelect(r_oField) {
	if (r_oField.type == "select-one") {
		var re = /[a-zA-Z:]/;	// \D gives bad result
		var val = r_oField.options[r_oField.selectedIndex].value;
		// true if option value > 0 or non-numeric
		return (val > 0 || re.test(val)) ? true : false;
	} else if (r_oField.type == "select-multiple") {
		for (var x = 0; x < r_oField.options.length; x++) {
			if (r_oField.options[x].selected) { return true; }
		}
	}
	return false;				// maybe default true if not select?
}

/*------------------------------------------------------------------------
	Name: 		isRadio()
	Purpose: 	make sure one radio button was checked
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	7/27/00		JEA		Creation
------------------------------------------------------------------------*/
function isRadio(r_oField) {
	for (var i = 0; i < r_oField.length; i++) {
		// cycle through each item in the radio collection
		if (r_oField[i].checked) { return true; }
	}
	// if we made it here then no radio is checked
	return false;
}

/*------------------------------------------------------------------------
	Name: 		isBirthDate()
	Purpose: 	allow only future expiration dates
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	10/5/00		JEA		Creation
	8/6/02		JEA		Use GetDate()
------------------------------------------------------------------------*/
function isCCExpire(r_oField) {
	var oDate = GetDate(r_oField);
	if (oDate) {
		var oThisDate = new Date();
		if (oDate >= oThisDate) { return true; }
	}
	return false;
}

/*------------------------------------------------------------------------
	Name: 		isBirthDate()
	Purpose: 	validate birthdate
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	8/23/00		JEA		Creation
	8/6/02		JEA		Use GetDate()
------------------------------------------------------------------------*/
function isBirthDate(r_oField) {
	var oDate = GetDate(r_oField);
	if (oDate) {
		var oThisDate = new Date();
		var lThisYear = cleanYear(oThisDate.getYear());
		var lMaxAge = 120;		// not many people older than that
		// can't have birthday in future or more than MaxAge
		if (oDate < oThisDate && lYear > (lThisYear - lMaxAge)) { return true; }
	}
	return false;
}

/*------------------------------------------------------------------------
	Name: 		isDate()
	Purpose: 	make sure that a date has been entered
				allows dashes or slashes, m/d/yy or mm/dd/yyyy
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	7/27/00		JEA		Creation
	8/6/02		JEA		Use GetDate()
------------------------------------------------------------------------*/
function isDate(r_oField) {
	var oDate = GetDate(r_oField)
	if (oDate) { return true; }
	return false;
}

/*------------------------------------------------------------------------
	Name: 		cleanYear()
	Purpose: 	make years four digits; assume century break on xx10
	Return:		int
Modifications:
	Date:		Name:	Description:
	8/22/00		JEA		Creation
------------------------------------------------------------------------*/
function cleanYear(v_sYear) {
	if (v_sYear.length == 2) { v_sYear = (v_sYear > 10 ? "19" : "20") + v_sYear; }
	return v_sYear;
}

/*------------------------------------------------------------------------
	Name: 		GetDate()
	Purpose: 	return valid date object
	Return:		object/boolean
Modifications:
	Date:		Name:	Description:
	8/6/02		JEA		Creation
------------------------------------------------------------------------*/
function GetDate(r_oField) {
	var re = /^((\d{1,2})[\/-\\](\d{1,2})[\/-\\](\d{2,4}))$/;
	var YEAR = 4; var MONTH = 2; var DAY = 3;
	if (re.test(r_oField.value)) {
		// format is right--now check each date value
		var arMatch = re.exec(r_oField.value);
		var lMonth = arMatch[MONTH];
		var lDay = arMatch[DAY];
		var lYear = cleanYear(arMatch[YEAR]);
		
		if (lMonth <= 12 && lMonth >= 1 && lDay <= 31 && lDay >= 1 && lYear <= 2010 && lYear >= 1850) {
			// -1 on month seems necessary for js date glitch
			return new Date(lYear, arMatch[MONTH] - 1, arMatch[DAY]);
		}
	} 
	// invalid date format
	return false;
}

/*------------------------------------------------------------------------
	Name: 		isMoney()
	Purpose: 	validate monetary fields
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	9/14/01		JEA		Creation
------------------------------------------------------------------------*/
function isMoney(r_oField) {
	var lMoney = r_oField.value;
	if (parseFloat(lMoney) != lMoney * 1) {
		return false;	// non-numeric values in field
	}
	var lCents = lMoney * 100;
	if (Math.abs(lCents - Math.floor(lCents)) > 0) {
		return false;	// fractional pennies not allowed
	}
	return true;
}

/*------------------------------------------------------------------------
	Name: 		isCCN()
	Purpose: 	do Mod10 check on CCN
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	10/12/00	JEA		Creation
------------------------------------------------------------------------*/
function isCCN(r_oField) {
	var sCCN = toNumeric(r_oField.value)
	
	// temp validation to check out with test CCN
	var re = /^41{14,15}$/;
	if (re.test(sCCN)) { return true; }
	// end temp stuff ---------------------------
	
	var re = /^\d{15,16}$/;
	// fail if wrong length
	if (!(re.test(sCCN))) {	return false; }

	var lLength = sCCN.length;
  	var bEven = lLength & 1;
	var lCheckSum = 0;

	for (var i = 0; i < lLength; i++) {
		var thisNum = parseInt(sCCN.charAt(i));
		if (!((i & 1) ^ bEven)) {
			thisNum *= 2;
			if (thisNum > 9) { thisNum -= 9; }
		}
		lCheckSum += thisNum;
	}
	// fail if non-zero Mod10
	if (lCheckSum % 10 != 0) { return false; }
	return true;
}

/*------------------------------------------------------------------------
	Name: 		isPosting()
	Purpose: 	only allow basic HTML in postings
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	12/24/02	JEA		Creation
------------------------------------------------------------------------*/
function isPosting(r_oField) {
	if (r_oField.value == "") { return false; }
	var re = /<[^abiu\/]/gi;
	return (!(re.test(r_oField.value) || (r_oField.value.indexOf("<img") != -1)));
}

/*------------------------------------------------------------------------
	Name: 		isImage()
	Purpose: 	limit input box to files with image extension
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	1/5/03		JEA		Creation
------------------------------------------------------------------------*/
function isImage(r_oField) {
	var lCursor = r_oField.value.lastIndexOf(".")
	if (lCursor == -1) { return false; }
	var sExt = r_oField.value.substring(lCursor + 1, r_oField.value.length);
	sExt = sExt.toLowerCase();
	return (sExt == "gif" || sExt == "jpg" || sExt == "jpeg");
}

/*------------------------------------------------------------------------
	Name: 		toNumeric()
	Purpose: 	converts a field to all numbers
	Return:		number
Modifications:
	Date:		Name:	Description:
	7/27/00		JEA		Creation
	7/12/02		JEA		updated to regexp
------------------------------------------------------------------------*/
function toNumeric(v_sField) { return v_sField.replace(/\D/g, ""); }

/*------------------------------------------------------------------------
	Purpose: 	basic regular expression pattern checks
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	12/23/02	JEA		Creation
------------------------------------------------------------------------*/
function TestField(r_oField, r_oRE) {	return r_oRE.test(r_oField.value); }
function isPassword(r_oField) {			return TestField(r_oField, /^.{6,20}$/); }
function isCVV(r_oField) {				return TestField(r_oField, /\d{3,4}/); }
function isEmail(r_oField) { 			return TestField(r_oField, /^([\w\._\-]+@[\w\-]+\.[\w\-]+\.*[\w\-]*.*[\w\-]*)$/); }
function isZip(r_oField) {				return TestField(r_oField, /^(\d{5})$/); }
function isZip4(r_oField) {				return TestField(r_oField, /^(\d{4})$/); }
function isNumeric(r_oField) {			return TestField(r_oField, /^(\d+)$/); }
function isURL(r_oField) {				return TestField(r_oField, /\.\w{2,3}\/?/); }
function isName(r_oField) {				return TestField(r_oField, /\w{2,}/); }

/*------------------------------------------------------------------------
	Purpose: 	customized checks
	Return:		boolean
Modifications:
	Date:		Name:	Description:
	12/23/02	JEA		Creation
------------------------------------------------------------------------*/
function isInRange(r_oField) { return (r_oField.value >= m_lRangeLow && r_oField.value <= m_lRangeHigh); }
function isPhone(r_oField) { var sPhone = toNumeric(r_oField.value); return (sPhone.length == 10); }
function isSSN(r_oField) { var sSSN = toNumeric(r_oField.value); return (sSSN.length == 9); }
function isString(r_oField) { return (r_oField.value != ""); }

/*------------------------------------------------------------------------
	Name: 		IsValidType()
	Purpose: 	check for valid upload type
Modifications:
	Date:		Name:	Description:
	7/28/04		JEA		Creation
------------------------------------------------------------------------*/
function IsValidType(v_sTypeList, v_sFileName, v_sTypeName) {
	var aTypes = v_sTypeList.split(",");
	var sExtension = v_sFileName.substr(v_sFileName.lastIndexOf(".") + 1, v_sFileName.length);
	
	for (var x = 0; x < aTypes.length; x++) {
		if (aTypes[x] == sExtension) { return true; }
	}
	alert("." + sExtension + " Files are not an allowed " + v_sTypeName + " type.\n\nPlease select a " + v_sTypeName + " file to upload.");
	return false;
}