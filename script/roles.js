AddEvent(window, "global", function() { BuildRoles("tree"); } );

/*------------------------------------------------------------------------
	add members to given role node

	Date:		Name:	Description:
	2/1/05		JEA		Creation
------------------------------------------------------------------------*/
function BuildRoles(rootNodeID) {
	var _onRole = null;
	var _rootNode = DOM.GetNode(rootNodeID, true);
	var _roles = new Array(_rootNode.childNodes.length);
	var _permissionBox = DOM.GetNode("permissionsBox", true);
	var _permission = new DragDrop(_permissionBox);
	
	InitializeChildNodes();

	function InitializeChildNodes() {
		var node, id;
		_permissionBox.Under = DOM.Under;
		for (var x = 0; x < _roles.length; x++) {
			// roles are enumerated so ids are sequential
			node = _rootNode.childNodes[x]; id = node.id.split(":")[1];
			_roles[id] = RoleNode(node, _permissionBox);
		}
	}
	_permission.OnDrop = function(node) {
		if (_onRole != null) {
			var roleNode = _roles[_onRole];
			if (!roleNode.Add(node)) { alert("The role already has that permission"); }
			roleNode.style.backgroundColor = "";
		}
	}
	_permission.OnDrag = function(node) {
		_onRole = null;
		for (var x = 0; x < _roles.length; x++) {
			if (_roles[x].Under(node)) {
				var id = node.id.split(":")[1];
				if (_roles[x].Permissions.Contains(id)) {
					_roles[x].style.backgroundColor = "#700";
				} else {
					_roles[x].style.backgroundColor = "#070";
				}
				_onRole = x;
			} else {
				_roles[x].style.backgroundColor = "";
			}
		}
	}
}

/*------------------------------------------------------------------------
	add members to given role node

	Date:		Name:	Description:
	2/1/05		JEA		Creation
------------------------------------------------------------------------*/
function RoleNode(node, permissionBox) {
	node.Under = DOM.Under;
	node.AddChild = AddChild;
	node.Permissions = new PermissionCollection(node);
	node.RoleID = node.id.split(":")[1];
	node.Add = function(dragNode) {
		var id = dragNode.id.split(":")[1];
		if (!this.Permissions.Contains(id)) {
			this.Permissions.Add(id);
			var newNode = dragNode.cloneNode(true);
			with (newNode) {
				className = "add";
				// IE needs these styles set explicitly
				style.left = "auto"; style.top = "auto"; style.position = "relative";
			}
			// add to nodes showing inheritance
			for (var x = 0; x < this.Inherits.length; x++) {
				this.Inherits[x].AddChild(newNode.cloneNode(true), id);
			}
			// add to main role node and attach event for dragging
			SetupDragDrop(this.AddChild(newNode, id), false);
			// oob
			var oob = new RolesBroker();
			oob.AddPermission(this.RoleID, id);
			return true;
		}
		return false;
	}
	
	InitializeChildNodes(node);
	
	// add expando members to children of role node
	function InitializeChildNodes(node) {
		if (node.childNodes.length > 1) {
			// setup permission nodes
			var roots = DOM.GetElementsByClassName("permission", node);
			var id, labels, ns, permissionNode;
			
			// handle permissions being dragged off (deleted)
			SetupDragDrop(node.childNodes[1]);
			
			// build array of all permissions belonging to role
			for (var x = 0; x < roots.length; x++) {
				for (var y = 0; y < roots[x].childNodes.length; y++) {
					permissionNode = roots[x].childNodes[y];
					ns = permissionNode.id.split("_");
					id = ns[ns.length - 1].split(":")[1];
					node.Permissions.Add(id);
				}
			}
			// make role labels clickable
			labels = DOM.GetElementsByClassName("label", node);
			for (var x = 0; x < labels.length; x++) {
				AddEvent(labels[x], "click", Toggle);
			}
		}
		// get nodes that display inheritance of this role
		var re = new RegExp("inherit:" + node.RoleID + "$", "i");
		node.Inherits = DOM.GetElementsByRegExp(re, "id");
		for (var x = 0; x < node.Inherits.length; x++) {
			node.Inherits[x].AddChild = AddChild;
			node.Inherits[x].RemoveChild = RemoveChild;
			node.Inherits[x].ParentRole = RoleParent(node.Inherits[x]);
		}
	}
	// find role node that inherits from given role node
	function RoleParent(node) {
		var parent = node;
		var x = 0;	// prevent infinite loop
		while (typeof(parent.RoleID) == "undefined" && x < 20) { parent = parent.parentNode; x++; }
		return parent;
	}
	// attach drag-drop object to node and handle events as deletions
	function SetupDragDrop(node, forChildren) {
		node.DragOff = new DragDrop(node, false, forChildren);
		node.DragOff.OnDrag = DeleteDrag;
		node.DragOff.OnDrop = DeleteDrop;
	}
	// add permission node to current node
	function AddChild(node, id) {
		var firstNode = this.childNodes[1].childNodes[0];
		var newNode = firstNode.parentNode.insertBefore(node, firstNode);
		// add permission value to any role inheriting from this one
		if (typeof(this.ParentRole) != "undefined") { this.ParentRole.Permissions.Add(id); }
		return newNode;
	}
	// remove permission node with given id from current node
	function RemoveChild(id) {
		node = DOM.GetNode("permission:" + id);
		node.parentNode.removeChild(node);
		// remove permission value from any role inheriting from this one
		this.ParentRole.Permissions.Remove(id);
	}
	// remove dropped permission node from parent role nodes
	function DeleteDrop(node) {
		if (node.Active) {
			// delete permissions with oob call
			var role = node.Original.parentNode.parentNode;
			var nodeID = node.id.split("_")
			var id = nodeID[nodeID.length - 1].split(":")[1];
			role.Permissions.Remove(id);
			node.Original.parentNode.removeChild(node.Original);
			
			// remove nodes showing inheritance
			for (var x = 0; x < role.Inherits.length; x++) {
				role.Inherits[x].RemoveChild(id);
			}
			
			// oob
			var roleID = role.id.split(":")[1];
			var oob = new RolesBroker();
			oob.RemovePermission(roleID, id);
		} else {
			node.Original.style.display = "block";
		}
	}
	// check if dragged permission node is over deletion area
	function DeleteDrag(node) {
		node.Active = false;
		if (permissionBox.Under(node)) {
			node.Active = true;
			node.className += " noDrop";
		} else {
			node.className = node.className.replace(" noDrop", "");
		}
	}
	// display or hide permissions list
	function Toggle() {
		if (window.event) {		// IE
			var node = event.srcElement;
			if (node.nodeName == "IMG") { node = node.parentNode; }
		} else {
			var node = this;
		}
		var oldState = parseInt(node.Expanded);
		if (isNaN(oldState)) { oldState = 0; }
		var newState = (oldState == 1) ? 0 : 1;
		var childNodes = node.parentNode.childNodes;
		var image = node.childNodes[0];
		var suffix = ["+.","-."];
		var display = ["none","block"];
		
		for (var x = 1; x < childNodes.length; x++) {
			childNodes[x].style.display = display[newState];
		}
		if (typeof(image.style.filter) == "string") {	// IE
			image.style.filter = image.style.filter.replace(suffix[oldState], suffix[newState]);
		} else {
			image.src = image.src.replace(suffix[oldState], suffix[newState]);	
		}
		node.Expanded = newState;
	}
	return node;
}

/*------------------------------------------------------------------------
	collection of permissions for given role

	Date:		Name:	Description:
	1/29/05		JEA		Creation
------------------------------------------------------------------------*/
function PermissionCollection(roleNode) {
	var _permissions = new Array();
	var me = this;
	
	this.Contains = function(id) {
		for (var x = 0; x < _permissions.length; x++) {
			if (_permissions[x] == id) { return true; }
		}
		return false;
	}
	this.Add = function(id) {
		if (!me.Contains(id)) { _permissions.push(id); }
	}
	this.Remove = function(id) {
		if (me.Contains(id)) {
			var permissions = new Array(_permissions.length - 1);
			for (var x = 0; x < _permissions.length; x++) {
				if (_permissions[x] != id) { permissions.push(_permissions[x]); }
			}
			_permissions = permissions;
		}
	}
	this.RoleNode = function() { return roleNode; }
}