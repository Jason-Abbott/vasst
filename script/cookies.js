/*------------------------------------------------------------------------
	manage browser cookies
	http://www.jibbering.com/faq/faq_notes/cookies.html

	Date:		Name:	Description:
	8/14/00		JEA		Created
	1/24/05		JEA		Updated old methods
------------------------------------------------------------------------*/
function Cookie() {
	var me = this;
	var _supported = (typeof document.cookie == "string");
	var _expires = new Date("January 1, 2010 12:00:00");

	this.Get = function(name) {
		if (_supported) {
			var cookies = unescape(document.cookie);
			var re = new RegExp(name + "=([^;]+)");
			if (re.test(cookies)) { return unescape(re.exec(cookies)[1]); }
		}
		return "";
	}
	this.Set = function(name, value, expires, path, domain, secure) {
		if (_supported) {
			expires = ((expires) ? expires : _expires);
			document.cookie = name + "=" + escape(value) +
				((expires) ? ";expires=" + expires.toGMTString() : "") +
				((path) ? ";path=" + path : "") +
				((domain) ? ";domain=" + domain : "") +
				((secure) ? ";secure" : "");
			
			//alert(expires.toGMTString());
		}
	}
	this.Delete = function(name, path, domain) {
		if (me.Get(name)) {
			document.cookie = name + "=" +
			((path) ? ";path=" + path : "") +
			((domain) ? ";domain=" + domain : "") +
			";expires=Thu, 01-Jan-70 00:00:01 GMT";
		}
	}
}