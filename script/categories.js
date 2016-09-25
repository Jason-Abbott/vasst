var Category;
AddEvent(null, "global", function() { Category = new CategoryClass(); } );

/*------------------------------------------------------------------------
	object to manage categories

	Date:		Name:	Description:
	2/15/05		JEA		Creation
------------------------------------------------------------------------*/
function CategoryClass() {
	var me = this;
	var _changed = false;
	var _category = null;
	var _sections = new Object();
	var _entities = new Object();
	var _assets = new Object();
	var _software = new Object();
	var _service = "../service.aspx?";
	var _inputs;
	
	Initialize();
	
	// create objects containing inputs for easy access
	function Initialize() {
		var entities = document.getElementById("entities");
		var sections = document.getElementById("sections");
		var x, inputs;

		inputs = DOM.GetElementsByRegExp(/checkbox/, "type", sections);
		for (x = 0; x < inputs.length; x++) {		// sections
			inputs[x].onclick = Click;
			_sections[inputs[x].value] = inputs[x];
		}
		inputs = DOM.GetElementsByRegExp(/Asset\+Types_/, "id", entities);
		for (x = 0; x < inputs.length; x++) {		// assets
			_assets[inputs[x].value] = inputs[x];
		}
		inputs = DOM.GetElementsByRegExp(/Software\+Types_/, "id", entities);
		for (x = 0; x < inputs.length; x++) {		// software
			_software[inputs[x].value] = inputs[x];
		}
		inputs = DOM.GetElementsByRegExp(/Entity_/, "id", entities);
		for (x = 0; x < inputs.length; x++) {		// business entities
			switch (parseInt(inputs[x].value)) {
				case 1: MakeFamily(_assets, inputs[x]); break;
				case 32: MakeFamily(_software, inputs[x]); break;
				default: inputs[x].onclick = Click; break;
			}
			_entities[inputs[x].value] = inputs[x];
		}
		_inputs = [_sections, _entities, _assets, _software];
	}
	// associate parent and child input nodes
	function MakeFamily(nodeList, parentNode) {
		parentNode.childInputs = nodeList;
		parentNode.onclick = ParentClick;
		for (var id in nodeList) {
			nodeList[id].parentInput = parentNode;
			nodeList[id].onclick = ChildClick;
		}
	}
	// click events
	function Click() {
		_changed = true;
		if (_category != null) { _category.className = "pending"; }
	}
	function ParentClick() {
		Click();
		var node = (window.event) ? event.srcElement : this;
		for (var id in node.childInputs) {
			node.childInputs[id].checked = node.checked;
		}
	}
	function ChildClick() {
		Click();
		var node = (window.event) ? event.srcElement : this;
		var siblings = node.parentInput.childInputs;
		var checked = false;
		for (var id in siblings) {
			if (siblings[id].checked) { checked = true; break; }
		}
		node.parentInput.checked = checked;
	}
	// return list of object members
	function SelectionList(object) {
		var list = "";
		for (var id in object) { if (object[id].checked) { list += "," + id; } }
		return list.substr(1);
	}
	// clear all inputs to prepare for new category
	function ClearInputs() {
		for (var x = 0; x < _inputs.length; x++) {
			for (id in _inputs[x]) { _inputs[x][id].checked = false; }
		}
	}
	// make server request in response to category selection
	this.Select = function(node) {
		me.Save(node);
		_category = node;
		node.className = "active";
		var request = new ServerCall("CategoryLoad")
		request.Parameters = "id=" + node.id;
		request.Callback = Load;
		request.Service = _service;
		Global.ProgressBar.Start("Loading Category");
		request.Start();
	}
	// out-of-band response loads category details
	function Load(category) {
		Global.ProgressBar.Stop();
		if (category.Errors.length == 0) {
			ClearInputs();
			var x;
			for (x = 0; x < category.Sections.length; x++) {
				_sections[category.Sections[x]].checked = true;
			}
			for (x = 0; x < category.Entities.length; x++) {
				_entities[category.Entities[x]].checked = true;
			}
			for (x = 0; x < category.Assets.length; x++) {
				_assets[category.Assets[x]].checked = true;
			}
			for (x = 0; x < category.Software.length; x++) {
				_software[category.Software[x]].checked = true;
			}
		} else {
			alert(category.Errors[0]);
		}
	}
	// out-of-band call to save category detail if needed
	this.Save = function(newNode) {
		if (_category == null) { return; }
		if (!_changed || newNode == _category) { _category.className = ""; return; }
		
		var method = (_category.id == "new") ? "CategoryAdd" : "CategorySave";
		var request = new ServerCall(method);
		
		request.Parameters = "id=" + _category.id + 
			"&name=" + escape(_category.value) +
			"&entities=" + SelectionList(_entities) +
			"&assets=" + SelectionList(_assets) +
			"&software=" + SelectionList(_software) +
			"&sections=" + SelectionList(_sections);
		request.Callback = SaveResult;
		request.Service = _service;
		
		Global.ProgressBar.Start("Saving Category");
		request.Start();
	}
	function SaveResult(category) {
		Global.ProgressBar.Stop();
		if (category.Errors.length == 0) {
			if (category.New) {		// reload page to show new category
				location.reload(false);
			} else {
				var node = document.getElementById(category.ID);
				node.className = "saved";
				setTimeout("DOM.ClassForId('" + category.ID + "','')", 2000);
			}
		} else {
			alert(category.Errors[0] + ((category.Errors.length > 1) ? category.Errors[1] : ""));
		}
	}
	// setup node for entering new category data
	this.Add = function(node) {
		me.Save(node);
		ClearInputs();
		_category = node;
		node.value = null;
		node.className = "active";
	}
}