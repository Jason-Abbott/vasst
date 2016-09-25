/*------------------------------------------------------------------------
	object to handle login validation and functions

	Date:		Name:	Description:
	1/5/05		JEA		Creation
------------------------------------------------------------------------*/
function LoginValidation(login) {
	var me = this;
	var _ignoreFields = [];
	var _confirmHiddenField = login.ConfirmHiddenField();
	
	SetupValidation();

	/*------------------------------------------------------------------------
		override standard validation with local function
		wait for validation object to complete initialization

		Date:		Name:	Description:
		1/8/05		JEA		Creation
	------------------------------------------------------------------------*/
	function SetupValidation() {
		if (typeof(Validation) == "undefined") { setTimeout(SetupValidation, 100); return; }
		
		// intercept normal validation
		Global.Form().onsubmit = function() { return me.IsValid(); };
		
		// add custom validation function
		// validate two password entries
		Validation.Functions.push( function() {
			if (login.Registering) {
				var password = DOM.GetNode("fldPassword");
				var again = DOM.GetNode("fldPasswordAgain", true);
				if (Validation.IsPassword(password)) {
					if (password.value != again.value) {
						DOM.SetError(again);
						return "The second password value does not match the first";
					}
				}
				DOM.ClearError(again);
			}
			return null;
		} );
		// validate confirmation code
		Validation.Functions.push( function() {
			if (login.Registering) {
				var confirm = DOM.GetNode("fldConfirmationCode");
				if (confirm.value != _confirmHiddenField.value) {
					DOM.SetError(confirm);
					return "The confirmation code you have entered does not match the one most recently e-mailed";
				}
				DOM.ClearError(confirm);
			}
			return null;
		} );
	}

	/*------------------------------------------------------------------------
		update validation object array (see validation.js)

		Date:		Name:	Description:
		1/7/05		JEA		Creation
	------------------------------------------------------------------------*/
	this.IsValid = function() {
		login.AllowServerCall = false;	// prevent server calls until validation finishes
		
		var fields = Validation.Fields;
		var keepFields = [];
		
		// first restore any previously ignored fields
		while (_ignoreFields.length > 0) {
			fields.push(_ignoreFields.pop())
		}
		
		// now remove non-pertinent fields
		for (var x = 0; x < fields.length; x++) {
			//if ((!login.Registering && LoginField(fields[x])) || login.Registering) {
			if (login.Registering || LoginField(fields[x])) {
				keepFields.push(fields[x]);
			} else {
				_ignoreFields.push(fields[x]);
			}
		}
		Validation.Fields = keepFields;	// overwrite original array
		
		function LoginField(field) {
			return (field.Match("fldEmail") || field.Match("fldPassword"));
		}
		return Validation.Check();
	}
}