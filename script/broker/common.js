/*------------------------------------------------------------------------
	create an xmlHttp request
	
	http://jibbering.com/2002/4/httprequest.html
	http://georgenava.com/samples/sodascript.html
	http://xkr.us/code/javascript/XHConn/
	http://msdn.microsoft.com/msdnmag/issues/04/08/CuttingEdge/
	http://webfx.eae.net/dhtml/xmlextras/demo.html

Modifications:
	Date:		Name:	Description:
	1/5/05		JEA		Creation
	2/14/05		JEA		Add method to synchronously check URL
	2/26/05		JEA		Handle fatal errors
------------------------------------------------------------------------*/
function ServerCall(method) {
	var me = this;
	var _http = GetXmlHttp();
	var _method = method;

	this.Callback;
	this.Parameters;
	this.Service = "service.aspx?";
	
	this.Start = function() {
		if (_http) {
			var qs = "method=" + _method + "&" + me.Parameters;
			_http.open("GET", me.Service + qs, true);
			_http.onreadystatechange = function() {
				var response;
				if (_http.readyState == 4) {
					try {
						response = eval(_http.responseText);
					} catch(e) {
						var error = _http.responseText;
						error = error.substring(error.lastIndexOf("<!--") + 5, error.length - 5);
						response = { Errors:[e, error] };
					}
					me.Callback(response);
				}
			};
			_http.send(qs)
		} else {
			// unable to create object
			me.Callback( { Errors:["Unable to create server connection object"] } );
		}
	}
	
	// could do this directly but permissions disallow
	this.Exists = function(url) {
		if (_http) {
			var qs = "method=UrlCheck&url=" + escape(url);
			_http.open("GET", me.Service + qs, false);
			_http.send(null);
			try { return eval(_http.responseText); } catch(e) {  }
		} 
		// if something fails then default to true
		return true;
	}

	function GetXmlHttp() {
		var xmlHttp = false;

		try { xmlHttp = new XMLHttpRequest(); }
		catch (e1) {
			try { xmlHttp = new ActiveXObject("Msxml2.XMLHTTP"); } 
			catch (e2) {
				try { xmlHttp = new ActiveXObject("Microsoft.XMLHTTP"); }
				catch (e3) { xmlHttp = new FrameBroker(); }
			}
		}
		return xmlHttp;
	}
}

/*------------------------------------------------------------------------
	fail over to out-of-band call through hidden frame if no xmlhttp
	
	http://www.scss.com.au/family/andrew/webdesign/xmlhttprequest/
	http://www.scss.com.au/scripts/loadelement.js
	
Modifications:
	Date:		Name:	Description:
	3/5/05		JEA		Creation
------------------------------------------------------------------------*/
function FrameBroker() {
	var me = this;
	var _url;
	var _async;
	var _status = 0;
	var _method = "GET";
	
	this.onreadystatechange;
	this.responseText;
	this.readyState;
	this.status = function() { return _status; }
	
	this.open = function(method, url, async) {
		_method = method;
		_url = url;
		_async = async;
		_readyState = 1;
	};
	this.send = function(data) {
		var frame = document.createElement('IFRAME');
		with (frame.style) {
			border = 'none';
			visibility = 'hidden';
			position = 'absolute';
			top = "0";
			left = "0";
		}
		document.body.appendChild(frame);
		frame.src = _url;
		frame.onload = function() {
			me.responseText = frame.document.body.innerText.replace(/\&quot;/g, "\"");
			me.readyState = 4;
			me.onreadystatechange();
			document.body.removeChild(frame);
		};
	}
}