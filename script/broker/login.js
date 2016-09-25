/*------------------------------------------------------------------------
	server calls

	Date:		Name:	Description:
	1/5/05		JEA		Creation
------------------------------------------------------------------------*/
function LoginBroker(login) {
	var me = this;
	var _timer;
	var _email;
	var _emailField = login.EmailField();
	var _emailNote = login.EmailNote();
	var _confirmNote = login.ConfirmNote();
	var _confirmHiddenField = login.ConfirmHiddenField();
	
	_emailField.onblur = CheckAddress; //this.CheckAddress;
	_emailField.onkeydown = function() { login.AllowServerCall = true; }
	
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
			login.FatalError(email.Errors, "We were unable to send your password");
			return;
		}
		if (email.Sent) {
			alert("A new password has been e-mailed to " + _emailField.value);
		} else if (!email.Exists) {
			if (confirm("\"" + _emailField.value + "\" is not a registered address\n"
				+ "Would you like to register it now?")) {
				
				login.ShowSignup(true);
			}
		}	
	}
	
	/*------------------------------------------------------------------------
		do out-of-band call to send confirmation e-mail

		Date:		Name:	Description:
		1/9/05		JEA		Creation
	------------------------------------------------------------------------*/
	function SendConfirmation() {
		if (login.AllowServerCall) {
			var request = new ServerCall("SendConfirmation");
			request.Callback = SendConfirmationResult;
			request.Parameters = "address=" + _emailField.value;
			Global.ProgressBar.Start("E-mailing Confirmation");
			request.Start();
		}
	}
	function SendConfirmationResult(email) {
		Global.ProgressBar.Stop();
		_needConfirmation = false;
		
		if (email.Errors.length == 0 && email.Sent) {
			_emailNote.innerHTML = "a confirmation code has been sent to this address"
			_confirmHiddenField.value = email.Code;
			_confirmNote.innerHTML = "check your e-mail for the code";
		} else {
			login.FatalError(email.Errors, "We were unable to send the confirmation e-mail");
			_confirmHiddenField.value = "";
		}
	}
	
	/*------------------------------------------------------------------------
		do out-of-band call to see if e-mail address is available

		Date:		Name:	Description:
		1/7/05		JEA		Creation
	------------------------------------------------------------------------*/
	function CheckAddress() {
		if (login.AllowServerCall && Validation.IsEmail(_emailField)) {
			var serverCall = new ServerCall("EmailCheck");
			serverCall.Callback = CheckAddressResult;
			serverCall.Parameters = "address=" + _emailField.value;
			Global.ProgressBar.Start("Validating e-mail");
			serverCall.Start();
		}
	}
	function CheckAddressResult(email) {
		Global.ProgressBar.Stop();
		if (email.Errors.length != 0) {
			login.FatalError(email.Errors, "There was a failure while looking up your address");
			return;
		}
		// registering with existing e-mail
		if (login.Registering && email.Exists) {
			if (confirm("\"" + _emailField.value + "\" is already in use\n"
				+ "Would you like to sign in instead?")) {
			
				login.ShowSignup(false);
				DOM.ClearError(_emailField);
			} else {
				login.Validation.Select(_emailField)
			}
			return;
		}
		// account disabled
		if (email.Disabled) {
			login.FatalError([], "Your account is not currently enabled");
			return;
		}
		// registering with new e-mail
		if (login.Registering && !email.Exists) {
			DOM.ClearError(_emailField);
			SendConfirmation(); return;
		}
		// signing up with non-existent e-mail
		if (!(login.Registering || email.Exists)) {
			if (confirm("We have no record of \"" + _emailField.value + "\"\n"
				+ "Would you like to sign up instead?")) {
			
				login.ShowSignup(true);
				DOM.ClearError(_emailField);
			} else {
				DOM.SetError(_emailField);
			}
		}
	}
}