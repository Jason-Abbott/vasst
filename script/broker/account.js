/*------------------------------------------------------------------------
	server calls

	Date:		Name:	Description:
	1/5/05		JEA		Creation
------------------------------------------------------------------------*/
function AccountBroker(account) {
	var me = this;
	var _timer;
	var _lastChecked = null;
	var _emailField = account.EmailField();
	var _emailNote = account.EmailNote();
	var _confirmNote = account.ConfirmNote();
	var _confirmHiddenField = account.ConfirmHiddenField();
	
	_emailField.onblur = CheckAddress;
	_emailField.onkeydown = function() { account.AllowServerCall = true; }
	
	/*------------------------------------------------------------------------
		do out-of-band call to send new password

		Date:		Name:	Description:
		1/9/05		JEA		Creation
	------------------------------------------------------------------------*/
	this.SendPassword = function() {
		var _tried = false;
		var _email = null;
		GetEmail();
		
		function GetEmail() {
			if (!Validation.IsEmail(_emailField)) {
				var message = "Please enter your registered e-mail address";
				if (_tried) { message = "\"" + _emailField.value + "\" is not a valid address\n" + message; }
				_emailField.value = window.prompt(message, _emailField.value);
				// let them quit if nothing entered
				if (_emailField.value != "") { _tried = true; GetEmail(); }
			}
		}
		if (_emailField.value != "") {
			var request = new ServerCall("ResetPassword");
			request.Callback = SendPasswordResult;
			request.Parameters = "address=" + _emailField.value;
			Global.ProgressBar.Start("Sending Password");
			request.Start();
		}
	}
	function SendPasswordResult(email) {
		Global.ProgressBar.Stop();
		if (email.Errors.length != 0) {
			account.FatalError(email.Errors, "We were unable to send your password");
			return;
		}
		if (email.Sent) {
			alert("A new password has been e-mailed to " + _emailField.value);
		} else if (!email.Exists) {
			if (confirm("\"" + _emailField.value + "\" is not a registered address\n"
				+ "Would you like to register it now?")) {
				
				account.ShowSignup(true);
			}
		}	
	}
	
	/*------------------------------------------------------------------------
		do out-of-band call to send confirmation e-mail

		Date:		Name:	Description:
		1/9/05		JEA		Creation
		2/28/05		JEA		Only send confirmation if address changed
	------------------------------------------------------------------------*/
	function SendConfirmation() {
		if (account.AllowServerCall && _emailField.value != _lastChecked) {
			var request = new ServerCall("SendConfirmation");
			request.Callback = SendConfirmationResult;
			request.Parameters = "address=" + _emailField.value;
			Global.ProgressBar.Start("E-mailing Confirmation");
			_lastChecked = _emailField.value;
			request.Start();
		}
	}
	function SendConfirmationResult(email) {
		Global.ProgressBar.Stop();
		
		
		if (email.Errors.length == 0 && email.Sent) {
			_emailNote.innerHTML = "a confirmation code has been sent to this address"
			_confirmHiddenField.value = email.Code;
			_confirmNote.innerHTML = "check your e-mail for the code";
			account.ShowConfirmation(true);
		} else {
			account.FatalError(email.Errors, "We were unable to send the confirmation e-mail");
			_confirmHiddenField.value = "";
			_lastChecked = null;
		}
	}
	
	/*------------------------------------------------------------------------
		do out-of-band call to see if e-mail address is available

		Date:		Name:	Description:
		1/7/05		JEA		Creation
		2/28/05		JEA		Added checks for account edit
	------------------------------------------------------------------------*/
	function CheckAddress() {
		if (account.EmailChanged()) {
			if (account.AllowServerCall && Validation.IsEmail(_emailField)) {
				var serverCall = new ServerCall("EmailCheck");
				serverCall.Callback = CheckAddressResult;
				serverCall.Parameters = "address=" + _emailField.value;
				Global.ProgressBar.Start("Validating e-mail");
				serverCall.Start();
			}
		} else {
			DOM.ClearError(_emailField);
			account.ShowConfirmation(false);
		}
	}
	function CheckAddressResult(email) {
		Global.ProgressBar.Stop();
		if (email.Errors.length != 0) {
			account.FatalError(email.Errors, "There was a failure while looking up your address");
			return;
		}
		if (email.Exists) {		// existing e-mail scenarios --------
			
			// creating account with existing e-mail
			if (account.Status == Status.Create) {
				if (confirm("\"" + _emailField.value + "\" is already in use\n"
					+ "Would you like to sign in instead?")) {
				
					account.ShowSignup(false);
					DOM.ClearError(_emailField);
				} else {
					DOM.SetError(_emailField);
				}
				return;
			}
			
			// trying to update account with e-mail that already exists
			if (account.Status == Status.Edit && account.EmailChanged()) {
				if (confirm("\"" + _emailField.value + "\" is already in use\n"
					+ "Would you like to sign in with that address?")) {
				
					location.href = "./signin.aspx??action=out"
				} else {
					DOM.SetError(_emailField);
				}
				return;
			}
			
		
		} else {				// new e-mail scenarios -------------
			
			// saving account with new e-mail, send confirmation
			if (account.Status & (Status.Create | Status.Edit)) {
				DOM.ClearError(_emailField);
				SendConfirmation(); return;
			}
			
			// signing up with non-existent e-mail
			if (account.Status == Status.Signin) {
				if (confirm("We have no record of \"" + _emailField.value + "\"\n"
					+ "Would you like to sign up instead?")) {
				
					account.ShowSignup(true);
					DOM.ClearError(_emailField);
				} else {
					DOM.SetError(_emailField);
				}
			}
		}
		
		// account disabled
		if (email.Disabled) {
			account.FatalError([], "Your account is not currently enabled");
			return;
		}
	}
}