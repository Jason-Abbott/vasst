/*------------------------------------------------------------------------
	server calls for roles

	Date:		Name:	Description:
	1/29/05		JEA		Creation
	2/14/05		JEA		Reload page to good state on error
------------------------------------------------------------------------*/
function RolesBroker() {

	this.AddPermission = function(roleID, permissionID) {
		var request = new ServerCall("RolePermissionAdd");
		request.Callback = AddPermissionResult;
		request.Service = "../service.aspx?";
		request.Parameters = "rID=" + roleID +"&pID=" + permissionID;
		request.Start();
	}
	function AddPermissionResult(role) {
		if (role.Errors.length > 0) {
			alert("The permissions could not be added.\nThe screen will now be refreshed.");
			location.reload(false);
		}
		//alert(role.Errors[0] + (role.Errors.length > 1) ? "\n\n" + role.Errors[1] : "");
	}
	this.RemovePermission = function(roleID, permissionID) {
		var request = new ServerCall("RolePermissionRemove");
		request.Callback = RemovePermissionResult;
		request.Service = "../service.aspx?";
		request.Parameters = "rID=" + roleID +"&pID=" + permissionID;
		request.Start();
	}
	function RemovePermissionResult(role) {
		if (role.Errors.length > 0) {
			alert("The permissions could not be removed.\nThe screen will now be refreshed.");
			location.reload(false);
		}
		//alert(role.Errors[0] + (role.Errors.length > 1) ? "\n\n" + role.Errors[1] : "");
	}
}