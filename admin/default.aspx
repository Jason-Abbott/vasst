<%@ Page Language="vb" AutoEventWireup="false" Codebehind="default.aspx.vb" Inherits="AMP.Pages.Administration.Home" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:panel id="pnlBody" runat="server">
	<h1>Administration</h1>
	<br/>
	<a href="categories.aspx">Categories</a><br/>
	<a href="roles.aspx">Roles</a><br/>
	<a href="contests.aspx">Contests</a><br/>
	<a href="contest-entries.aspx">Contest Entries</a><br/>
	<br/>
	<amp:button id="btnFlush" text="Flush Data" runat="server" submitform="true" />
	
</asp:panel>

<asp:panel id="pnlWindows" runat="server"/>

</body>
</html>