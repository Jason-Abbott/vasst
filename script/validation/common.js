var Validation;
AddEvent(window, "global", InitValidation);

/*------------------------------------------------------------------------
	initialize form validation
	http://www.aplus.co.yu/scripts/validating-forms/

	Date:		Name:	Description:
	1/2/05		JEA		Creation
	2/23/05		JEA		Check that validators exist
------------------------------------------------------------------------*/
function InitValidation() {
	Validation = new ValidationClass();
	Validation.Fields = (typeof(Validators) == "undefined") ? new Array() : Validators();
}

/*------------------------------------------------------------------------
	master validation function

	Date:		Name:	Description:
	7/27/00		JEA		Creation
	7/28/04		JEA		Special URL handling
	1/2/05		JEA		Get element by ID and add error css class
------------------------------------------------------------------------*/
function ValidationClass() {
	var me = this;
	var _errors = new FieldErrors();
	this.Fields;			// array of validator objects
	this.Functions = [];	// array of custom validation functions

	Global.Form().onsubmit = function() { return me.Check(); };

	this.Check = function() {
		_errors.Clear();
		var error;
		
		for (var x = 0; x < this.Functions.length; x++) {
			// execute all custom validation functions
			error = this.Functions[x]();
			if (error) { _errors.Add(error); }
		}
		for (var x = 0; x < this.Fields.length; x++) {
			// perform individual field validations
			var node = this.Fields[x];
			if (!node.Valid()) { _errors.Add(node); }
			else { node.ClearError(); }
		}
		if (_errors.Exist) { alert(_errors.Display()); return false; }
		else { return true; }
	}
	
	/*------------------------------------------------------------------------
		error container object

		Date:		Name:	Description:
		1/3/05		JEA		Creation
	------------------------------------------------------------------------*/
	function FieldErrors() {
		this.Focused = false;							// has focus been set on a node
		this.Exist = false;								// do errors exist
		this.Messages = new Array;
		var _stringTitle = false;
		var _fieldTitle = false;

		this.Add = function(validator) {				// overloaded (sorta)
			var message;
			if (typeof(validator) == "string") {		// treat parameter as string
				message = " - " + validator;
				if (!_stringTitle) {
					if (this.Messages.length > 0) { this.Messages.push(""); }
					this.Messages.push("The following issues were encountered:");
					_stringTitle = true;
				}
			} else {									// treat parameter as Validator object
				message = " - " + validator.Message;
				validator.SetError();
				if (!this.Focused) { validator.Focus(); this.Focused = true; }
				if (!_fieldTitle) {
					if (this.Messages.length > 0) { this.Messages.push(""); }
					this.Messages.push("The following fields could not be validated:");
					_fieldTitle = true;
				}
			}
			this.Messages.push(message);
			this.Exist = true;
		}
		// function to display errors
		this.Display = function() {
			//var message = "Please enter a valid\n";
			var message = "";
			for (var x = 0; x < this.Messages.length; x++) {
				message += this.Messages[x] + "\n";
			}
			return message;
		}
		this.Clear = function() {
			this.Messages.length = 0;
			this.Exist = false;
			_stringTitle = false;
			_fieldTitle = false;
		}
	}
	
	/*------------------------------------------------------------------------
		field validator object

		Date:		Name:	Description:
		7/24/00		JEA		Creation
		1/3/05		JEA		Updated node names and added methods
		2/24/05		JEA		Corrected .Required logic
	------------------------------------------------------------------------*/
	this.Field = function(id, type, message, required) {
		var me = this;
		var _validation = Validation["Is" + type];
		var _node = document.getElementById(id);
		var _type = type;

		this.Message = message;
		this.Required = required;
		
		this.Valid = function() {
			var value = "";
			if (_node.type.indexOf("select-") != -1) {
				if (_node.selectedIndex != -1) {
					value = _node.options[_node.selectedIndex].value;
				}
			} else {
				value = _node.value;
			}
			if (value.length > 0) {
				if (value == this.IgnoreValue()) { return true; }
			} else {
				return !this.Required;
			}
			return _validation(_node)
		}
		// ignore default values in certain node types
		this.IgnoreValue = function() {
			switch (_type) {
				case "URL":	return "http://";
				default:	return ""
			}
		}
		// clear and set css error style
		this.ClearError = function() {
			_node.className = _node.className.replace("error", "");
		}
		this.SetError = function() {
			if (_node.className.indexOf("error") == -1) {
				_node.className += " error";
			}
		}
		this.HasValue = function() {
			return (_node.value.length > 0 && _node.value != this.IgnoreValue())
		}
		this.Match = function(text) { return (_node.id.indexOf(text) != -1); }
		this.Focus = function() { _node.focus(); }
	}

	/*------------------------------------------------------------------------
		make sure they made some selection
		this assumes that layout options, like lines, are <= 0

		Date:		Name:	Description:
		7/27/00		JEA		Creation
		8/6/02		JEA		Check node type
		1/8/03		JEA		allow multiple selects
		1/3/05		JEA		change valid value regex
	------------------------------------------------------------------------*/
	this.IsSelect = function(node) {
		if (node.type == "select-one") {
			var re = /[\w,:\.]/;
			var val = node.options[node.selectedIndex].value;
			// true if option value > 0 or non-numeric
			return (re.test(val) && val != 0);
		} else if (node.type == "select-multiple") {
			for (var x = 0; x < node.options.length; x++) {
				if (node.options[x].selected) { return true; }
			}
		}
		return false;				// maybe default true if not select?
	}

	/*------------------------------------------------------------------------
		make sure one radio button was checked

		Date:		Name:	Description:
		7/27/00		JEA		Creation
	------------------------------------------------------------------------*/
	this.IsRadio = function(node) {
		for (var i = 0; i < node.length; i++) {
			// cycle through each item in the radio collection
			if (node[i].checked) { return true; }
		}
		// if we made it here then no radio is checked
		return false;
	}

	/*------------------------------------------------------------------------
		allow only future expiration dates

		Date:		Name:	Description:
		10/5/00		JEA		Creation
		8/6/02		JEA		Use GetDate()
	------------------------------------------------------------------------*/
	this.IsCCExpire = function(node) {
		var date = me.GetDate(node);
		if (date) {
			var today = new date();
			if (date >= today) { return true; }
		}
		return false;
	}

	/*------------------------------------------------------------------------
		validate birthdate

		Date:		Name:	Description:
		8/23/00		JEA		Creation
		8/6/02		JEA		Use GetDate()
	------------------------------------------------------------------------*/
	this.IsBirthDate = function(node) {
		var date = me.GetDate(node);
		if (date) {
			var today = new Date();
			var thisYear = CleanYear(today.getYear());
			var maxAge = 120;		// not many people older than that
			// can't have birthday in future or more than MaxAge
			if (date < today && thisYear > (thisYear - maxAge)) { return true; }
		}
		return false;
	}

	/*------------------------------------------------------------------------
		make sure that a date has been entered
		allows dashes or slashes, m/d/yy or mm/dd/yyyy

		Date:		Name:	Description:
		7/27/00		JEA		Creation
		8/6/02		JEA		Use GetDate()
	------------------------------------------------------------------------*/
	this.IsDate = function(node) {
		if (me.GetDate(node)) { return true; }
		return false;
	}

	/*------------------------------------------------------------------------
		make years four digits; assume century break on xx10

		Date:		Name:	Description:
		8/22/00		JEA		Creation
	------------------------------------------------------------------------*/
	function CleanYear(year) {
		if (year.length == 2) { year = (year > 10 ? "19" : "20") + year; }
		return year;
	}

	/*------------------------------------------------------------------------
		return valid date object

		Date:		Name:	Description:
		8/6/02		JEA		Creation
	------------------------------------------------------------------------*/
	this.GetDate = function(node) {
		var re = /^((\d{1,2})[\/-\\](\d{1,2})[\/-\\](\d{2,4}))$/;
		var YEAR = 4; var MONTH = 2; var DAY = 3;
		if (re.test(node.value)) {
			// format is right--now check each date value
			var matches = re.exec(node.value);
			var month = matches[MONTH];
			var day = matches[DAY];
			var year = CleanYear(matches[YEAR]);
			
			if (month <= 12 && month >= 1 && day <= 31 && day >= 1 && year <= 2010 && year >= 1850) {
				// -1 on month seems necessary for js date glitch
				return new Date(year, matches[MONTH] - 1, matches[DAY]);
			}
		} 
		// invalid date format
		return false;
	}

	/*------------------------------------------------------------------------
		validate monetary nodes

		Date:		Name:	Description:
		9/14/01		JEA		Creation
	------------------------------------------------------------------------*/
	this.IsMoney = function(node) {
		var money = node.value;
		if (parseFloat(money) != money * 1) {
			return false;	// non-numeric values in node
		}
		var cents = money * 100;
		if (Math.abs(cents - Math.floor(cents)) > 0) {
			return false;	// fractional pennies not allowed
		}
		return true;
	}

	/*------------------------------------------------------------------------
		do Mod10 check on CCN

		Date:		Name:	Description:
		10/12/00	JEA		Creation
	------------------------------------------------------------------------*/
	this.IsCCN = function(node) {
		var sCCN = ToNumeric(node.value)
		
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
		only allow basic HTML in postings

		Date:		Name:	Description:
		12/24/02	JEA		Creation
	------------------------------------------------------------------------*/
	this.IsPosting = function(node) {
		if (node.value == "") { return false; }
		var re = /(<|&lt;)[^abiu\/]/gi;
		return (!(re.test(node.value) || (node.value.indexOf("<img") != -1)));
	}
	
	/*------------------------------------------------------------------------
		allow only plain text; check for esacped variants (&gt;, %20)

		Date:		Name:	Description:
		1/13/05		JEA		Creation	
	------------------------------------------------------------------------*/
	this.IsPlainText = function(node) {
		// e.g. precludes <div> or &#216; or %20
		var re = /<|>|\&\#?\w{1,10}\;|\%\d{2,3}/g;
		return (!(re.test(node.value)));
	}

	/*------------------------------------------------------------------------
		limit input box to files with image extension

		Date:		Name:	Description:
		1/5/03		JEA		Creation
	------------------------------------------------------------------------*/
	this.IsImage = function(node) {
		var index = node.value.lastIndexOf(".")
		if (index == -1) { return false; }
		var extension = node.value.substring(index + 1, node.value.length);
		extension = extension.toLowerCase();
		return (extension == "gif" || extension == "jpg" || extension == "jpeg");
	}

	/*------------------------------------------------------------------------
		check if URL is valid and active with xmlHttp

		Date:		Name:	Description:
		2/14/05		JEA		Creation
	------------------------------------------------------------------------*/
	this.IsActiveURL = function(node) {
		if (me.IsURL(node)) {
			var request = new ServerCall(null);
			return request.Exists(node.value);
		}
		return false;
	}

	/*------------------------------------------------------------------------
		converts a node to all numbers

		Date:		Name:	Description:
		7/27/00		JEA		Creation
		7/12/02		JEA		updated to regexp
	------------------------------------------------------------------------*/
	function ToNumeric(node) { return node.replace(/\D/g, ""); }

	/*------------------------------------------------------------------------
		basic regular expression pattern checks

		Date:		Name:	Description:
		12/23/02	JEA		Creation
		1/3/05		JEA		Added IsFile()
	------------------------------------------------------------------------*/
	function TestField(node, re) {		return re.test(node.value); }
	this.IsPassword = function(node) {	return TestField(node, /^.{6,}$/); }
	this.IsCVV = function(node) {		return TestField(node, /\d{3,4}/); }
	this.IsEmail = function(node) {		return TestField(node, /^([\w\._\-]+@[\w\-]+\.[\w\-]+\.*[\w\-]*.*[\w\-]*)$/); }
	this.IsZip = function(node) {		return TestField(node, /^(\d{5})$/); }
	this.IsZip4 = function(node) {		return TestField(node, /^(\d{4})$/); }
	this.IsNumeric = function(node) {	return TestField(node, /^(\d+)$/); }
	//http://www.sundancemediagroup.com/articles/ragged_text_in_Sony_Vegas.htm
	this.IsURL = function(node) {		return TestField(node, /\.\w{2,3}/); }
	this.IsName = function(node) {		return TestField(node, /^[a-zA-Z\s'_]{2,}$/); }
	this.IsFile = function(node) {		return TestField(node, /^(\w\:|\\)\\[^\/\:\*\?\"\<\>\|]+\.\w{3,4}$/); }

	/*------------------------------------------------------------------------
		customized checks

		Date:		Name:	Description:
		12/23/02	JEA		Creation
		2/19/05		JEA		Added NonZero()
	------------------------------------------------------------------------*/
	this.IsPhone = function(node) { var sPhone = ToNumeric(node.value); return (sPhone.length == 10); }
	this.IsSSN = function(node) { var sSSN = ToNumeric(node.value); return (sSSN.length == 9); }
	this.IsString = function(node) { return (node.value != ""); }
	this.IsNonZero = function(node) {
		if (me.IsNumeric(node) && node.value > 0) { return true; } 
		return false;
	}
}