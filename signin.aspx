<%@ Page Language="vb" AutoEventWireup="false" Codebehind="signin.aspx.vb" Inherits="AMP.Pages.Signin" EnableViewState="true"%>
<%@ Register TagPrefix="amp" TagName="ContentWebPart" Src="~/control/Content.ascx"%>
<%@ Register TagPrefix="amp" TagName="SignInUp" Src="~/control/SignInUp.ascx"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:Panel id="pnlBody" runat="server">
	<h2 runat="server" class="form">to continue, please</h2>
	<amp:SignInUp id="ampSignInUp" runat="server"></amp:SignInUp>
	<div class="nav">
		<amp:Button id="btnLogin" runat="server" text="Sign in" submitform="true" />
	</div>
	<input id="tipType" type="hidden" value="welcome" />
</asp:Panel>

<asp:Panel id="pnlWindows" runat="server">
	<amp:contentwebpart id="ampContent" title="Welcome" runat="server" file="help/SignInOrUp.html" />
</asp:Panel>

</body>
</html>