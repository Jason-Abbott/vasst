<%@ Page Language="vb" AutoEventWireup="false" Codebehind="unsubscribe.aspx.vb" Inherits="AMP.Pages.Unsubscribe"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="NewAssets" Src="~/control/NewAssets.ascx"%>
<%@ Register TagPrefix="amp" TagName="Tours" Src="~/control/Tours.ascx"%>
<%@ Register TagPrefix="amp" TagName="Contests" Src="~/control/Contests.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:Panel id="pnlBody" runat="server">
	<h1>Cancel Mailing List Subscription</h1>

	<div id="pleaseNote">
		Please Note: Canceling e-mail for an account used to view and download free resources will cause that account to be disabled.
	</div>
	
	<div id="form">
		<label style="width: 120px;" runat="server"><%=Say("Label.EmailAddress")%>:</label>
		<amp:Validation target="tbEmail" type="Email" message="E-mail Address" required="true" runat="server" />
		<asp:TextBox ID="tbEmail" Runat="server" Width="200" />
		<amp:button id="btnUnsubscribe" submitform="true" text="Unsubscribe" runat="server" />
	</div>
	
</asp:Panel>

<asp:Panel id="pnlWindows" runat="server">
	<amp:Contests id="Contests" runat="server" />
	<amp:NewAssets count="10" id="ampNewListings" runat="server" />
	<amp:Tours NewerThan="1/1/04" id="ampTours" runat="server" />
</asp:Panel>

</body>
</html>
