AddEvent(window, "load", ShareValidation );

/*------------------------------------------------------------------------
	object to handle validation for first share page

	Date:		Name:	Description:
	1/12/05		JEA		Creation
	1/28/05		JEA		Add terms acceptance
------------------------------------------------------------------------*/
function ShareValidation() {
	// wait for main validation object to initialize
	if (typeof(Validation) == "undefined") { setTimeout(ShareValidation, 500); return; }
		
	// ensure only one optional value was entered
	Validation.Functions.push( function() {
		var hasValue = false;
		var file = false;
		
		for (var x = 0; x < Validation.Fields.length; x++) {
			if (Validation.Fields[x].HasValue()) {
				if (hasValue) {
					// too many values
					return "A file or the address of a web page, not both, must be entered to continue";
				}
				hasValue = true;
				if (Validation.Fields[x].Match("fldUpload")) { file = true; }
			}
		}
		if (!hasValue) {
			return "A file or the address of a web page must be entered to continue";
		} else {
			if (file) {
				// ensure that terms were accepted
				if (!DOM.GetNode("fldTerms").checked) {
					return "The Terms & Conditions for uploaded files must be accepted to continue";
				}
			}
			return null;
		}
	} );
}