function DOMClass() {
	var me = this;
	var _opacity;	
	var _timer = null;
	var _node = null;
	
	// http://www.sitepoint.com/blog-post-view.php?id=211431
	this.SetOpacity = function(node, level) {
		with (node.style) {
			filter = "alpha(opacity:" + (100 * level) + ")";	// IE
			KHTMLOpacity = level;								// Konqueror, old Safari
			MozOpacity = level;									// old Mozilla
			opacity = level;									// W3C
		}
	}
	
	this.Show = function(id) { me.GetNode(id,true).style.display = "block"; }
	
	/*------------------------------------------------------------------------
		fade out given node

		Date:		Name:	Description:
		12/7/04		JEA		Creation	
	------------------------------------------------------------------------*/
	this.Fade = function(node, speed) {
		if (_timer != null) { clearInterval(_timer); }
		_opacity = 0.999;		// 1 causes FF bug
		_node = node;
		_timer = setInterval(FadeStep, speed);
	}
	function FadeStep() {
		if (_opacity > 0) {
			me.SetOpacity(_node, _opacity);
			_opacity -= 0.01;
		} else {
			_node.style.display = "none";
			_node = null;
			_opacity = 0.999;
			clearInterval(_timer);
		}
	}
	
	this.ClassForId = function(id, className) {
		document.getElementById(id).className = className;
	}
	
	/*------------------------------------------------------------------------
		handle button rollover

		Date:		Name:	Description:
		11/25/04	JEA		Creation	
	------------------------------------------------------------------------*/
	this.Button = function(img) {
		var image = (typeof(img.style.filter) == "string") ? img.style.filter : img.src;
		var re = /_on\.(png|gif|jpg)/;
		if (re.test(image)) {		// turn off
			var changeTo = ".$1";
		} else {					// turn on
			var changeTo = "_on.$1";
			re = /\.(png|gif|jpg)/;
		}
		if (typeof(img.style.filter) == "string") {
			img.style.filter = image.replace(re, changeTo);
		} else {
			img.src = image.replace(re, changeTo);
		}
	}
	/*------------------------------------------------------------------------
		attempt to infer namespace added by .NET
		if greedy then namespace is everything to last underscore

		Date:		Name:	Description:
		1/4/05		JEA		Creation
		2/9/05		JEA		Improved with regex
	------------------------------------------------------------------------*/
	this.InferNamespace = function(greedy) {
		var elements = Global.Form().elements;
		var re = (greedy) ? /^_[^_]+_\w+$/ : /^_[^_]+_[^_]+$/
		for (var x = 0; x < elements.length; x++) {
			if (re.test(elements[x].id)) {
				id = elements[x].id;
				return id.substr(0, id.lastIndexOf("_") + 1);
			}
		}
		return null;
	}
	/*------------------------------------------------------------------------
		clear selections from option list

		Date:		Name:	Description:
		1/5/03		JEA		Creation
	------------------------------------------------------------------------*/
	this.ClearSelection = function(id) {
		var node = Global.Form.elements[id];
		for (var x = 0; x < node.options.length; x++) {
			node.options[x].selected = false;
		}
	}
	
	/*------------------------------------------------------------------------
		utility functions for drag & drop

		Date:		Name:	Description:
		2/1/05		JEA		Creation
	------------------------------------------------------------------------*/
	this.Under = function(node) {
		var under = new Rectangle(this);
		var over = new Rectangle(node);
		return (
			over.Middle < under.Bottom &&
			over.Middle > under.Top &&
			over.Left < under.Right &&
			over.Right > under.Left);
	}
	function Rectangle(node) {
		var _left = 0;
		var _top = 0;
		var _width = node.offsetWidth;
		var _height = node.offsetHeight;
		
		while (node.offsetParent) {
			_left += (node.offsetLeft - node.scrollLeft);
			_top += (node.offsetTop - node.scrollTop);
			node = node.offsetParent;
		}
		this.Left = _left;
		this.Top = _top;
		this.Right = _left + _width;
		this.Bottom = _top + _height;
		this.Middle = _top + (_height / 2);
	}
	
	/*------------------------------------------------------------------------
		change style of node to indicate or clear error

		Date:		Name:	Description:
		1/7/05		JEA		Creation
	------------------------------------------------------------------------*/
	this.SetError = function(node) {
		node.className += " error";
		node.select();
	}
	this.ClearError = function(node) {
		node.className = node.className.replace(" error", "");
	}
	
	this.GetNode = function(id, exact) {
		if (exact) { return document.getElementById(id); }
		var re = new RegExp(id + "$", "i");
		var nodes = DOM.GetElementsByID(re, null);
		return (nodes.length > 0) ? nodes[0] : null;
	}
	
	this.GetElementByNsId = function(id, greedy) {
		var ns = me.InferNamespace(greedy);
		return document.getElementById(ns + id);
	}
	
	/*------------------------------------------------------------------------
		recursively find all nodes with given classname

		Date:		Name:	Description:
		1/31/05		JEA		Creation
	------------------------------------------------------------------------*/
	this.GetElementsByClassName = function(className, node) {
		if (!node) { node = document.body; }
		var re = new RegExp(className, "i");
		return ElementsByRegExp(re, node, new Array(), "className");
	}
	/*------------------------------------------------------------------------
		recursively find all nodes with partial id

		Date:		Name:	Description:
		1/31/05		JEA		Creation
	------------------------------------------------------------------------*/
	this.GetElementsByID = function(re, node) {
		if (!node) { node = document.body; }
		return ElementsByRegExp(re, node, new Array(), "id");
	}
	/*------------------------------------------------------------------------
		recursively find all nodes with attribute matching regular expression

		Date:		Name:	Description:
		1/31/05		JEA		Creation
	------------------------------------------------------------------------*/
	this.GetElementsByRegExp = function(re, attribute, node) {
		if (!node) { node = document.body; }
		return ElementsByRegExp(re, node, new Array(), attribute);
	}
	function ElementsByRegExp(re, node, nodes, attribute) {
		for (var x = 0; x < node.childNodes.length; x++) {
			if (re.test(node.childNodes[x][attribute])) {
				nodes.push(node.childNodes[x]);
			}
			nodes = ElementsByRegExp(re, node.childNodes[x], nodes, attribute)
		}
		return nodes;
	}
}