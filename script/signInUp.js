var Login;
AddEvent(window, "global", function() { Login = new LoginClass(); } );

/*------------------------------------------------------------------------
	object to handle login validation and functions

	Date:		Name:	Description:
	1/5/05		JEA		Creation
	1/28/05		JEA		Added tip functionality
	2/26/05		JEA		Re-worked for fieldset layout
------------------------------------------------------------------------*/
function LoginClass(_form) {
	var me = this;
	// DOM element references
	var _links = DOM.GetNode("links", true);
	var _signupSection = DOM.GetNode("signup", true);
	var _signupHiddenField = DOM.GetNode("fldSignup");
	var _emailNote = DOM.GetNode("emailNote", true);
	var _emailField = DOM.GetNode("fldEmail");
	var _passwordNote = DOM.GetNode("passwordNote", true);
	var _passwordRepeat = DOM.GetNode("fldPasswordAgain", true);
	var _confirmHiddenField = DOM.GetNode("fldConfirmation");
	var _confirmNote = DOM.GetNode("confirmNote", true);
	var _signUpTip;
	var _signInTip;
	
	// interface members for broker/account.js and validation/account.js -----
	this.Status = Status.Signin;
	this.AllowServerCall = true;	// false during validation
	this.ConfirmHiddenField = function() { return _confirmHiddenField; }
	this.ConfirmNote = function() { return _confirmNote; }
	this.EmailNote = function() { return _emailNote; }
	this.EmailField = function() { return _emailField; }
	this.ShowConfirmation = function(visible) { return; }
	this.EmailChanged = function() { return true; }
	this.ShowSignup = function(visible) {
		var display = (visible) ? "block" : "none";
		me.Status = (visible) ? Status.Create : Status.Signin;
		_emailField.focus();
		_emailNote.innerHTML = "a confirmation code will be sent to this address";
		_passwordNote.innerHTML = "at least six characters&mdash;enter twice to verify";
		_signupSection.style.display = display;
		_links.style.display = (visible) ? "none" : "block";
		_emailNote.style.display = display
		_passwordNote.style.display = display
		_passwordRepeat.style.display = (visible) ? "inline" : "none";
		_signupHiddenField.value = visible;
		ShowTip();
	}
	this.FatalError = function(errors, message) {
		var response = "Not able to continue.  " + message + ".";
		for (var x = 0; x < errors.length; x++) {
			response += "\n\n" + errors[x];
		}
		alert(response);
		location.href = ".";
	}
	// end interface members -------------------------------------------------

	UpdateTime();
	LoadTips();
	ShowTip();

	this.Validation = new AccountValidation(this);
	this.Broker = new AccountBroker(this);
	
	function UpdateTime() {	DOM.GetNode("fldClientTime").value = ClientTime();	}
	
	/*------------------------------------------------------------------------
		load DOM references to tip text

		Date:		Name:	Description:
		1/28/05		JEA		Creation
	------------------------------------------------------------------------*/
	function LoadTips() {
		var tipType = DOM.GetNode("fldTipType");
		if (tipType != null) {
			_signUpTip = DOM.GetNode(tipType.value + "SignUp", true);
			_signInTip = DOM.GetNode(tipType.value + "SignIn", true);
		}
	}
	function ShowTip() {
		var tip = (me.Create) ? _signUpTip : _signInTip;
		var oldTip = (me.Create) ? _signInTip : _signUpTip;
		if (typeof(tip) == "object") { tip.style.display = "block"; }
		if (typeof(oldTip) == "object") { oldTip.style.display = "none"; }
	}

	/*------------------------------------------------------------------------
		get SQL allowed format of client's time

		Date:		Name:	Description:
		12/30/02	JEA		Creation
		1/31/03		JEA		Accomodate Mozilla bug in .getYear()
	------------------------------------------------------------------------*/
	function ClientTime() {
		var now = new Date();
		var hour = now.getHours();
		var month = now.getMonth() + 1;			// getMonth is 0-based
		var year = now.getYear() + "";
		var amPm = (hour >= 12) ? " PM" : " AM";
		var time = (hour > 12) ? hour - 12 : hour;
		time += ":" + PadNumber(now.getMinutes(), 2) + ":" + PadNumber(now.getSeconds(), 2) + amPm;
		year = year.substr(year.length - 2, 2);	// mozilla was returning "103" for the year 2003
		return month + "/" + now.getDate() + "/" + year + " " + time;
	}
	
	/*------------------------------------------------------------------------
		pad number with leading zeros to match given length

		Date:		Name:	Description:
		12/30/02	JEA		Creation
	------------------------------------------------------------------------*/
	function PadNumber(number, length) {
		number += "";
		var shortBy = length - number.length;
		if (shortBy > 0) { for (x = 0; x < shortBy; x++) { number = "0" + number; } }
		return number;
	}
}