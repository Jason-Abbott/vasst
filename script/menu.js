var Menu = function() {
	this.On = function() {return null;}
	this.Off = function() {return null;}
}
AddEvent(window, "load", function() { Menu = new MenuClass(); } );

/*------------------------------------------------------------------------
	process menu events

	Date:		Name:	Description:
	12/5/04		JEA		Creation
------------------------------------------------------------------------*/
function MenuClass() {
	var _node;
	var _timer;

	this.On = function(node, id) {
		if (_node == node.offsetParent) {
			clearTimeout(_timer);
		} else if (_node != null) {
			// another menu showing; hide it
			Hide(_node);
		}
		if (id.length > 0) {
			// menu has children
			_node = DOM.GetNode(id, true);
			_node.parent = node;
			_node.visible = false;
			_timer = setTimeout("Menu.Show(" + (node.offsetTop + node.offsetHeight + 1) + "," + (node.offsetLeft - 1) + ")", 150);
			//_timer = setTimeout(Show, 150, (node.offsetTop + node.offsetHeight + 1), (node.offsetLeft - 1));
		} else {
			clearTimeout(_timer);
		}
		Hover(node, true);
	}
	this.Off = function(node, id) {
		if (id.length == 0) { Hover(node, false); }
		if (_node != null) {
			if (_node.visible) {
				// hide visible menu after delay
				_timer = setTimeout(Hide, 500);
			} else {
				// cancel pending display of menu
				Hover(node, false);
				clearTimeout(_timer);
			}
		}
	}
	// this has to be public so IE can eval Menu.Show()
	this.Show = function(pTop, pLeft) {
		with (_node.style) {
			display = "block";
			left = pLeft + "px";	//offsetLeft - 1 + "px";
			top = pTop + "px";  	//offsetTop + node.offsetHeight + 1 + "px";
		}
		_node.visible = true;
	}
	function Hide() {
		clearTimeout(_timer);
		if (_node != null) {
			Hover(_node.parent, false);
			_node.style.display = "none";
			_node = null;
		}
		
	}
	function Hover(node, over) {
		if (over) {
			node.className += " hover";
		} else {
			node.className = node.className.replace(" hover", "");
		}
	}
}