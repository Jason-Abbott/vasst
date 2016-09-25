var Account;
AddEvent(null, "global", function() { Account = new AccountClass(); } );

/*------------------------------------------------------------------------
	object to handle account validation and functions

	Date:		Name:	Description:
	2/9/05		JEA		Creation
------------------------------------------------------------------------*/
function AccountClass() {
	var me = this;
	// DOM element references
	var _emailField = DOM.GetNode("fldEmail");
	var _emailNote = DOM.GetNode("emailNote", true);
	var _confirmHiddenField = DOM.GetNode("fldConfirmation");
	var _confirmInput = DOM.GetNode("confirm", true);
	var _confirmNote = DOM.GetNode("confirmNote", true);
	
	// interface members for broker/account.js and validation/account.js -----
	this.Status = Status.Edit;
	this.AllowServerCall = true;	// false during validation
	this.ConfirmHiddenField = function() { return _confirmHiddenField; }
	this.ConfirmNote = function() { return _confirmNote; }
	this.EmailNote = function() { return _emailNote; }
	this.EmailField = function() { return _emailField; }
	this.ShowSignup = function(visible) { return; }
	this.ShowConfirmation = function(visible) {
		_confirmInput.style.display = (visible) ? "block" : "none";
	}
	this.EmailChanged = function() {
		return (DOM.GetNode("fldOldEmail").value != _emailField.value);
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
	
	this.Validation = new AccountValidation(this);
	this.Broker = new AccountBroker(this);
}