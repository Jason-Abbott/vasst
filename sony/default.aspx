<%@ Page Language="vb" AutoEventWireup="false" Codebehind="default.aspx.vb" Inherits="AMP.Pages.Sony.Home" enableViewState="False" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="NewAssets" Src="~/control/NewAssets.ascx"%>
<%@ Register TagPrefix="amp" TagName="Tours" Src="~/control/Tours.ascx"%>
<%@ Register TagPrefix="amp" TagName="Contests" Src="~/control/Contests.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head><body>

<asp:panel id="pnlBody" runat="server">
	<div class="content">
		<h1>Sony Vegas Resource Gateway</h1>
		<amp:Image src="./images/vegas5_boxes.png" transparency="true" runat="server" style="float: right; padding-left: 8px;" />
		<p>Welcome to the ultimate Sony Vegas<sup>&reg;</sup> resource gateway. Here you will find everything for Vegas whether it's training, books, helpful veg files, or links to valuable Vegas resources.</p>
		<ul class="todo"><li>Also show Vegas-related products along the right.</li></ul>
	</div>
</asp:panel>

<asp:panel id="pnlWindows" runat="server">
	<amp:contests id="Contests" section="sony" runat="server" />
	<amp:newassets count="10" section="sony" id="smgNewListings" runat="server" />
</asp:panel>

</body>
</html>