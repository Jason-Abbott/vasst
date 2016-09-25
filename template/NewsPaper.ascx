<%@ Control Language="vb" AutoEventWireup="false" Codebehind="NewsPaper.ascx.vb" Inherits="AMP.Templates.NewsPaper" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.ParadigmEdit" %>
<%@ Import Namespace="AMP" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
	<meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml" />
	
	<title runat="server"><%=Say("CompanyName")%><asp:PlaceHolder id="phTitle" runat="server" /></title>
	<asp:PlaceHolder id="phScriptFiles" runat="server" />
	<asp:PlaceHolder id="phStyleSheet" runat="server" />
</head>
<body onload="BodyLoad();">

<amp:form id="frmMain" method="post" runat="server">
	<div id="container">
		<div id="logo">logo
			<a id="homeLink" href="<%=Global.BasePath%>/default.aspx" title="AMP Home"><amp:image src="~/images/logo/logo.png" transparency="true" id="ampLogo" runat="server" /></a>
		</div>
		<div id="body">
			<div id="headline"><asp:PlaceHolder id="phHeadline" runat="server" /></div>
			<div id="rightColumn"><asp:PlaceHolder id="phRightColumn" runat="server" /></div>
			<div id="leftColumn"><asp:PlaceHolder id="phLeftColumn" runat="server" /></div>
		</div>
		<div id="footer">
			<div id="footLinks">
				<a class="image" href="http://www.spreadfirefox.com/?q=user/register&amp;r=42022"><amp:image src="~/images/getfirefox.png" alt="Get Firefox" rollover="Button(this);" runat="server" transparency="true" id="Image1" name="Image1"/></a>
				<a class="image" href="http://validator.w3.org/check/referer"><amp:image src="~/images/xhtml10.png" alt="XHTML 1.0 Validation" rollover="Button(this);" runat="server" transparency="true" id="Image2" name="Image2"/></a><br/>
				<a class="privacy" href="<%=Global.BasePath%>/privacy.aspx">Privacy Statement</a>
			</div>
			&copy; Copyright <%=Date.Now.Year()%> <a href="http://www.paradigmedit.com"><%=Say("CompanyName")%></a>. All Rights Reserved.
		</div>
	</div>
	<asp:PlaceHolder id="phMessage" runat="server" />
	<fieldset id="viewState"><asp:PlaceHolder id="phViewState" runat="server" /></fieldset>
</amp:form>

<asp:PlaceHolder id="phScriptBlock" runat="server" />
</body>
</html>
