<%@ Page Language="vb" AutoEventWireup="false" Codebehind="users.aspx.vb" Inherits="AMP.Pages.Administration.Users" %>
<%@ Register TagPrefix="amp" TagName="ContentWebPart" Src="~/control/Content.ascx"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:panel id="pnlBody" runat="server">
	<h1>User Administration</h1>
	<a href="roles.aspx">Roles</a>
</asp:panel>

<asp:panel id="pnlWindows" runat="server">
	<amp:contentwebpart file="help/Roles.txt" title="Tips & Help" id="ampContent" runat="server" />
</asp:panel>

</body>
</html>