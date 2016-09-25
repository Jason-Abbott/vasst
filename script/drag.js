/*------------------------------------------------------------------------
	move a node with mouse events to enable drag-drop

	Date:		Name:	Description:
	2/1/05		JEA		Creation
------------------------------------------------------------------------*/
function DragDrop(rootNode, move, children) {
	var me = this;
	var _move = (move == null) ? false : true;			// default false
	var _children = (children == null) ? true : false;	// default true
	var _node;
	var _rootNode = rootNode;
	var _resolution = 5;
	
	this.OnDrag = new Function();
	this.OnDrop = new Function();
	this.Node = function() { return _node; }
	
	if (_children) { InitializeChildNodes(); } else { Attach(_rootNode); }
	
	function InitializeChildNodes() {
		for (var x = 0; x < _rootNode.childNodes.length; x++) {
			Attach(_rootNode.childNodes[x]);
		}
	}

	function Attach(node) {
		node.onmousedown = Start;
		//node.addEventListener("mousedown", Start, true);
		// normalize styles
		if (isNaN(parseInt(node.style.left ))) { node.style.left  = "0px"; }
		if (isNaN(parseInt(node.style.top))) { node.style.top = "0px"; }
	}
	function Start(e) {
		_node = this.cloneNode(true);
		_node.Original = this;

        var point = new Point(this);
        
        if (_move) { _node.Original.style.display = "none"; }
		e = FixEvent(e);
       
		_node.style.position = "absolute";
        _node.className = "drag"; 
        _node.style.left = point.X + "px"; //(this.offsetLeft + _rootNode.offsetLeft) + "px";
        _node.style.top = point.Y + "px"; //(this.offsetTop + _rootNode.offsetTop) + "px";
        _node.lastMouseX = e.clientX;
        _node.lastMouseY = e.clientY;

        //_rootNode.parentNode.appendChild(_node);
        document.body.appendChild(_node);
        //document.documentElement.appendChild(_node);
        
        document.onmousemove = Move;
        document.onmouseup = End;
        
        //_report.innerHTML = _rootNode.scrollTop;

        return false;
	}
	
	function Point(node) {
		var _x = 0;
		var _y = 0;
		
		while (node.offsetParent) {
			_x += (node.offsetLeft - node.scrollLeft);
			_y += (node.offsetTop - node.scrollTop);
			node = node.offsetParent;
		}
		this.X = _x;
		this.Y = _y;
	}
	
	function Move(e) {
		e = FixEvent(e);
		
		var eventY = e.clientY;
        var eventX = e.clientX;
        var styleY = parseInt(_node.style.top);
        var styleX = parseInt(_node.style.left);
		var newX = styleX + (eventX - _node.lastMouseX);
		var newY = styleY + (eventY - _node.lastMouseY);
		
        _node.style.left = newX + "px";
        _node.style.top = newY + "px";
        _node.lastMouseX = eventX;
        _node.lastMouseY = eventY;
        
        //if (eventY % _resolution == 0 && eventX % _resolution == 0) {
	        me.OnDrag(_node);
	    //}

        return false;
	}
	
	function End() {
		me.OnDrop(_node);
		document.onmousemove = null;
        document.onmouseup = null;
        _node.parentNode.removeChild(_node);
	}
	
	function FixEvent(e) {
		if (typeof(e) == 'undefined') { e = window.event; }
        if (typeof(e.layerX) == 'undefined') { e.layerX = e.offsetX; }
        if (typeof(e.layerY) == 'undefined') { e.layerY = e.offsetY; }
        return e;
	}
}