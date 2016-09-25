<%@ Page Language="vb" AutoEventWireup="false" Codebehind="default.aspx.vb" Inherits="AMP.Pages.Ulead.Home"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="NewAssets" Src="~/control/NewAssets.ascx"%>
<%@ Register TagPrefix="amp" TagName="Tours" Src="~/control/Tours.ascx"%>
<%@ Register TagPrefix="amp" TagName="Contests" Src="~/control/Contests.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head><body>

<asp:panel id="pnlBody" runat="server">
	<div class="content">
		<amp:image src="./images/ulead_boxes.png" transparency="true" runat="server" style="float: right; padding-left: 8px;" />
		<h1>Ulead</h1>
		<p>Welcome to the Ulead <sup>&reg;</sup> resource gateway. Here you will find everything for Ulead products whether it's training, books, helpful project files, or links to valuable Ulead resources.</p>
	</div>
</asp:panel>

<asp:panel id="pnlWindows" runat="server">
	<amp:contests id="Contests" section="ulead" runat="server" />
	<amp:newassets count="10" section="ulead" id="smgNewListings" runat="server" />
</asp:panel>

</body>
</html>
