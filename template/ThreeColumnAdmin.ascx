<%@ Control Language="vb" AutoEventWireup="false" Codebehind="ThreeColumnAdmin.ascx.vb" Inherits="AMP.Templates.ThreeColumnAdmin" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="Statistics" Src="~/control/Statistics.ascx"%>
<%@ Register TagPrefix="amp" TagName="TaskList" Src="~/control/TaskList.ascx"%>
<%@ Import Namespace="AMP" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/tr/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
	<meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml" />
	
	<title runat="server" id="Title1"><%=Say("SiteName")%> : Administration<asp:placeholder id="phTitle" runat="server" /></title>
	<asp:placeholder id="phScriptFiles" runat="server" />
	<asp:placeholder id="phStyleSheet" runat="server" />
</head>
<body onload="BodyLoad();">

<div id="header">
	<fieldset id="search">
		<input type="text" id="fldSearch" onkeydown="Search.CheckKey(event)" />
		<amp:button id="smgSearch" resx="Search" runat="server" onclick="Search.Execute()" />
	</fieldset>
	<div id="corner" class="left"></div>
	<a id="homeLink" href="<%=Global.BasePath%>/" title="<%=Say("SiteName")%> Home"><amp:image src="~/images/logo/VASST.png" transparency="true" id="smgLogo" runat="server" /></a>
</div>

<amp:form id="frmMain" method="post" runat="server">
	<div id="menuStrip"><div id="spot" class="left"></div><amp:menu runat="server" id="smgMenu" file="~/admin/sitemap.xml" /></div>
	<div id="tagLine" class="left"><amp:image src="~/images/backgrounds/side_logo.png" transparency="true" runat="server" id="Image1"/></div>
	<div id="sideBar" class="right"><asp:placeholder id="phWebParts" runat="server" /><amp:TaskList id="ampTaskList" runat="server"/><amp:Statistics id="ampStatistics" runat="server"/></div>
	<div id="body" class="middle"><asp:placeholder id="phBody" runat="server" /></div>
	<div id="footer" class="footer middle" runat="server">
		This page is proprietary <a href="http://www.sundancemediagroup.com"><%=Say("CompanyName")%></a> content.  All access is logged.
	</div>
	
	<asp:placeholder id="phMessage" runat="server" />
	<div id="progressBar" class="topRight">
		<div id="text"></div>
		<div id="outline"><div id="bar"></div></div>
	</div>
	<fieldset id="viewState"><asp:placeholder id="phViewState" runat="server" /></fieldset>
</amp:Form>
<div id="feature" class="topRight">
<asp:placeholder id="phScriptBlock" runat="server" />
</body>
</html>