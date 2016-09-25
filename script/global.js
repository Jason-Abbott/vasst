var Message;
var Search;
var Global;
var Status = { Signin: 1, Create: 2, Edit: 4 }
var DOM;
var LoadFunctions = new Array();

/*------------------------------------------------------------------------
	function call included in template for all body tags

	Date:		Name:	Description:
	11/25/04	JEA		Creation	
------------------------------------------------------------------------*/
function BodyLoad() {
	DOM = new DOMClass();
	Global = new GlobalClass();
	WebPart = new WebPartClass();
	Search = new SearchClass();
	Message = new MessageClass();
	Message.Show();
	WebPart.LoadPreferences();
	for (var x = 0; x < LoadFunctions.length; x++) { LoadFunctions[x](); }
}

/*------------------------------------------------------------------------
	generic method to add a listener for an object event

	Date:		Name:	Description:
	1/3/05		JEA		Creation
------------------------------------------------------------------------*/
function AddEvent(node, type, fn) {
	if (type == "global") {
		// fire event after global object is created
		LoadFunctions.push(fn); return true;
	} else {
		if (node.addEventListener) { node.addEventListener(type, fn, true); return true; }
		else if (node.attachEvent) { return node.attachEvent("on" + type, fn); }
	}
	return false;
}

/*------------------------------------------------------------------------
	functionality used globally

	Date:		Name:	Description:
	1/3/05		JEA		Creation
------------------------------------------------------------------------*/
function GlobalClass() {
	var me = this;
	var _form = document.forms[0];
	
	PreFetchImages();
	FixIEfilterPath();
	
	this.Form = function() { return _form; }
	this.ProgressBar = new BarClass();

	/*------------------------------------------------------------------------
		give time for page to render then redirect, usually for download

		Date:		Name:	Description:
		1/23/05		JEA		Creation	
	------------------------------------------------------------------------*/
	this.Redirect = function(url) { 
		location.href = url;
		//setTimeout("location.href=\"" + url + "\"", 500);
	}

	/*------------------------------------------------------------------------
		cache mouse-over images

		Date:		Name:	Description:
		1/1/03		JEA		Creation
		11/28/04	JEA		Support other image formats
		1/24/05		JEA		Handle input images
	------------------------------------------------------------------------*/
	function PreFetchImages() {
		var _cached = new Array();
		var _re = /\.(png|gif|jpg)/gi;
		var _images = document.getElementsByTagName('img');
		var _input = document.getElementsByTagName('input');

		for (var x = 0; x < _images.length; x++) { CacheSrc(_images[x]); } 
		for (var x = 0; x < _input.length; x++) {
			if (_input[x].getAttribute('type') == "image") { CacheSrc(_input[x]); }
		}
		function CacheSrc(img) {
			var imagePath = null;
			if (typeof(img.style.filter) == "string") {
				var re = /src=[\'\"](.*)[\'\"],/i;
				var matches = re.exec(img.style.filter);
				if (matches != null) { imagePath = matches[1]; }
			} else {
				imagePath = img.getAttribute('src');
			}
			if (imagePath != null) {
				var imageName = imagePath.substring(imagePath.lastIndexOf("/") + 1, imagePath.length);
				if (imageName.substr(0,4) == "btn_" && imageName.indexOf("_on.") == -1) {
					_cached.push(new Image());
					_cached[_cached.length - 1].src = imagePath.replace(_re, "_on.$1");
				}
			}
		}
	}
	
	/*------------------------------------------------------------------------
		paths in IE css filter properties are page- rather than css-relative
		change pathing in stylesheet filter properties if in sub-folder

		Date:		Name:	Description:
		2/2/05		JEA		Creation
	------------------------------------------------------------------------*/
	function FixIEfilterPath() {
		var match = navigator.appVersion.match(/MSIE (\d+\.\d+)/, '');
		if (match != null && Number(match[1]) >= 5.5 && InSubFolder()) {
			var filter, re;
			for (var s = 0;	s < document.styleSheets.length; s++) {
				for (var r = 0; r < document.styleSheets[s].rules.length; r++) {
					filter = document.styleSheets[s].rules[r].style["filter"];
					re = /(\(src=['"])(\.+)/;
					if (filter && re.test(filter)) {
						match = re.exec(filter);
						if (match[2].length == 1) {
							filter = filter.replace(re, "$1$2.");
							document.styleSheets[s].rules[r].style["filter"] = filter;
						}
					}
				}
			}
		}
		function InSubFolder() { var re = /admin/; return re.test(location.href); }
	}
}

/*------------------------------------------------------------------------
	perform search

	Date:		Name:	Description:
	11/25/04	JEA		Creation	
------------------------------------------------------------------------*/
function SearchClass() {
	var me = this;
	var _node = document.getElementById("fldSearch");
	
	_node.focus();
	//_node.onkeydown = me.CheckKey;
	
	this.Execute = function() {
		var text = _node.value.replace(/\s+$/,"");
		if (text != "") {
			location.href = "search.aspx?text=" + escape(text);
		}		
	}
	this.CheckKey = function(e) {
		if ((e.which && e.which == 13) || (e.keyCode && e.keyCode == 13)) {
			e.returnValue = false; e.cancel = true; me.Execute();
		}
	}
	this.Focus = function() { _node.focus(); }
}

/*------------------------------------------------------------------------
	fade text in div tag

	Date:		Name:	Description:
	12/7/04		JEA		Creation
------------------------------------------------------------------------*/
function MessageClass() {
	var _opacity;
	var _timer;
	var _node = DOM.GetNode("message", true);
	
	this.Text = function(text) {
		if (text.length > 0) { _node.value = text; }
		return _node.value;
	}
	this.Show = function() {
		if (_node != undefined) { setTimeout(Fade, 3000); }
	}
	function Fade() { DOM.Fade(_node, 50); }
}

/*------------------------------------------------------------------------
	display progress bar

	Date:		Name:	Description:
	2/15/05		JEA		Creation
------------------------------------------------------------------------*/
function BarClass() {
	var me = this;
	var _timer;
	var _width;
	var _node = DOM.GetNode("progressBar", true);
	_node.text = DOM.GetNode("text", true);
	_node.bar = DOM.GetNode("bar", true);
	
	this.Start = function(message) {
		clearInterval(_timer);
		_width = 0;
		_node.text.innerHTML = message;
		_node.bar.style.width = "0%";
		_node.style.display = "block";
		_timer = setInterval(Grow, 300);
	}
	this.Stop = function() {
		clearInterval(_timer);
		DOM.Fade(_node, 20);
	}
	var Grow = function() {
		_width++;
		_node.bar.style.width = _width + "%";
	}
}