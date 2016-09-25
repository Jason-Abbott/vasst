<%@ Page AutoEventWireup="false" Codebehind="default.aspx.vb" Inherits="AMP.Pages.Home" enableViewState="False" language="VB"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="NewAssets" Src="~/control/NewAssets.ascx"%>
<%@ Register TagPrefix="amp" TagName="Tours" Src="~/control/Tours.ascx"%>
<%@ Register TagPrefix="amp" TagName="Contests" Src="~/control/Contests.ascx"%>
<%@ Register TagPrefix="amp" TagName="Quote" Src="~/control/Quote.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:Panel id="pnlBody" runat="server">
	<amp:Quote runat="server" id="ampQuote" />
	<amp:Content file="welcome.html" runat="server" id="ampWelcome" />
	<amp:Content file="mission.html" runat="server" id="ampMission" />
	<amp:Content file="origin.txt" runat="server" id="ampOrigin" />
</asp:Panel>

<asp:Panel id="pnlWindows" runat="server">
	<amp:Contests id="Contests" runat="server" />
	<amp:NewAssets count="10" id="ampNewListings" runat="server" />
	<amp:Tours NewerThan="1/1/04" id="ampTours" runat="server" />
</asp:Panel>

</body>
</html>