var Entry;
AddEvent(null, "global", function() { Entry = new ContestEntry(); } );

function ContestEntry() {
	var _actionNode = DOM.GetNode("fldAction");
	var _edit = (_actionNode != null);
	if (!_edit) { SetupValidation(); }
	
	this.Delete = function() {
		if (confirm("Click OK to permanently delete this entry")) {
			DoAction("delete");
		}
	}
	this.Approve = function(id) { SetID(id); DoAction("approve"); }
	this.Deny = function(id) { SetID(id); DoAction("deny"); }
	
	function DoAction(action) { _actionNode.value = action; Global.Form().submit(); }
	function SetID(id) { DOM.GetNode("fldEntryID").value = id; }
	
	function SetupValidation() {
		if (typeof(Validation) == "undefined") { setTimeout(SetupValidation, 100); return; }
	
		Validation.Functions.push( function() {
			if (!DOM.GetNode("fldRules").checked) {
				return "The contest rules must be accepted to continue";
			}
			return null;
		} );
	}
}