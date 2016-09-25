<%@ Control Language="VB" AutoEventWireup="false" Codebehind="ThreeColumn.ascx.vb" Inherits="AMP.Templates.ThreeColumn" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="Feature" Src="~/control/Feature.ascx"%>
<%@ Import Namespace="AMP" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/tr/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
	<meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml" />
	
	<title runat="server"><%=Say("SiteName")%><asp:PlaceHolder id="phTitle" runat="server" /></title>
	<asp:placeholder id="phScriptFiles" runat="server" />
	<asp:placeholder id="phStyleSheet" runat="server" />
</head>
<body onload="BodyLoad();">
	<div id="header">
		<div id="search">
			<fieldset>
				<input type="text" maxlength="50" id="fldSearch" onkeydown="Search.CheckKey(event)" />
				<amp:button id="smgSearch" resx="Search" runat="server" onclick="Search.Execute()" />
			</fieldset>
		</div>
		<div id="corner" class="left"></div>
		<a id="homeLink" href="<%=Global.BasePath%>/" title="<%=Say("SiteName")%> Home"><amp:image src="~/images/logo/VASST.png" transparency="true" id="smgLogo" runat="server" /></a>
	</div>

	<amp:form id="frmMain" method="post" runat="server">
		<div id="menuStrip"><div id="spot" class="left" onclick="location.href='./admin/'"></div><amp:menu runat="server" id="smgMenu" /></div>
		<div id="tagLine" class="left"><amp:image src="~/images/backgrounds/side_logo.png" transparency="true" runat="server" /></div>
		<div id="sideBar" class="right"><asp:PlaceHolder id="phWebParts" runat="server" /></div>
		<div id="body" class="middle"><asp:PlaceHolder id="phBody" runat="server" /></div>
		<div id="footer" class="footer middle" runat="server">
			<div id="footerRight">
				<a class="image" href="http://www.spreadfirefox.com/?q=user/register&amp;r=42022"><amp:image src="~/images/getfirefox.png" alt="Get Firefox" rollover="DOM.Button(this);" runat="server" transparency="true" /></a>
				<a class="image" href="http://validator.w3.org/check/referer"><amp:image src="~/images/xhtml10.png" alt="XHTML 1.0 Validation" rollover="DOM.Button(this);" runat="server" transparency="true" /></a><br/>
				<a class="privacy" href="<%=Global.BasePath%>/privacy.aspx">Privacy Statement</a>
			</div>
			&copy; Copyright <%=Date.Now.Year()%> <a href="http://www.sundancemediagroup.com"><%=Say("CompanyName")%></a>. All Rights Reserved.
		</div>
		<amp:feature id="Feature" runat="server" />
		<asp:PlaceHolder id="phMessage" runat="server" />
		<div id="progressBar" class="topRight">
			<div id="text"></div>
			<div id="outline"><div id="bar"></div></div>
		</div>
		<fieldset id="viewState"><asp:PlaceHolder id="phViewState" runat="server" /></fieldset>
	</amp:Form>

	<asp:PlaceHolder id="phScriptBlock" runat="server" />
</body>
</html>