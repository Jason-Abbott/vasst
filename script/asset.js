var Asset;
AddEvent(null, "global", function() { Asset = new AssetClass(); } );

function AssetClass() {
	var _actionNode = DOM.GetNode("fldAction");
	
	this.Delete = function() {
		if (confirm("Click OK to permanently delete this resource")) {
			DoAction("delete");
		}
	}
	this.Approve = function(id) { SetID(id); DoAction("approve"); }
	this.Deny = function(id) { SetID(id); DoAction("deny"); }
	
	function DoAction(action) { _actionNode.value = action; Global.Form().submit(); }
	function SetID(id) { DOM.GetNode("fldAssetID").value = id;	}
}