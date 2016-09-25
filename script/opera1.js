function load(url, targetObj) {
	if (typeof(targetObj) == "string") {targetObj = document.getElementById(targetObj);}
	if (typeof(download) != "undefined" && typeof(download.startDownload) != "undefined") {
		loadIE(url, targetObj);
	} else if (typeof(XMLHttpRequest) != "undefined") {
        loadXML(url, targetObj);
	} else if (navigator.javaEnabled()) {
		loadJava(url, targetObj);
	} else {
		loadFrame(url, targetObj);
	}
}

function loadFrame(url, targetObj) {
	var iframe = document.createElement("IFRAME");
	iframe.style.border="none";
	iframe.style.width="0px";
	iframe.style.height="0px";
	iframe.style.visbility="hidden";
	document.body.appendChild(iframe);
	iframe.src = url;
	iframe.onload = function() {
		targetObj.innerHTML = iframe.document.body.innerHTML;
		document.body.removeChild(iframe);
	};
}

function loadJava(srcUrl, targetObj) {
	var source = "";
	var url = new java.net.URL(new java.net.URL(window.location.href),srcUrl);
	var stream = new java.io.DataInputStream(url.openStream());
	var line    = "";
	while ((line = stream.readLine()) != null) {
		source += line + "\n";
	}
	stream.close();
	targetObj.innerHTML = source;
}

function loadXML(srcUrl, targetObj) {
	var req = new XMLHttpRequest();
	req.onreadystatechange = function() {
	if (req.readyState == 4) {
		if (req.status == 200) {
		targetObj.innerHTML = req.responseText;
		} else {
		alert("no data");
		}
	}
	};
	req.open("GET", srcUrl, true);
	req.send(null);
}

function loadIE(srcUrl, targetObj) {
	var replaceContent = function(content) { targetObj.innerHTML=content; };
	download.startDownload(srcUrl, replaceContent);
}

  </script>  
  <IE:DOWNLOAD ID="download" STYLE="behavior:url(#default#download)"/>
</head>
<body>
  <div id="dl" style="border:1px solid
black;padding:10px;background:rgb(230,230,230)"></div>
  <a href="javascript:load('load.html','dl')">bla</a>
</body>
</html 