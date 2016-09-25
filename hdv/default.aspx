<%@ Page AutoEventWireup="false" Codebehind="default.aspx.vb" Inherits="AMP.Pages.HDV.Home" enableViewState="False" language="VB"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="NewAssets" Src="~/control/NewAssets.ascx"%>
<%@ Register TagPrefix="amp" TagName="Tours" Src="~/control/Tours.ascx"%>
<%@ Register TagPrefix="amp" TagName="Contests" Src="~/control/Contests.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head><body>

<asp:panel id="pnlBody" runat="server">
	<div style="float: right; text-align: right;">
		<amp:image src="./images/hvr-z1u.png" transparency="true" runat="server" /><br/>
		<amp:image src="./images/hdr-fx1.png" transparency="true" runat="server" />
	</div>
	<div class="content">
		<h1>HDV</h1>
		<ul class="todo">
			<li>Line up and group HDV resources along the right, the same way Vegas files are shown under "New Resources" and such.  This provides a consistent user experience.</li>
			<li>The area along the right can also show related products, in this case the HDV book.  When you're in the Vegas section it could show Ultimate-S and other Vegas products.</li>
		</ul>
	</div>
</asp:panel>

<asp:panel id="pnlWindows" runat="server">
	<amp:contests id="Contests" section="hdv" runat="server" />
	<amp:newassets count="10" section="hdv" id="smgNewListings" runat="server" />
</asp:panel>

</body>
</html>