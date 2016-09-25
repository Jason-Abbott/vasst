<%@ Page Language="vb" AutoEventWireup="false" Codebehind="resource.aspx.vb" Inherits="AMP.Pages.Resource"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="Contests" Src="~/control/Contests.ascx"%>
<%@ Register TagPrefix="amp" TagName="Results" Src="~/control/SearchResults.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:Panel id="pnlWindows" runat="server">
	<amp:Contests id="Contests" runat="server" />
	<amp:Results id="SearchResults" runat="server" />
</asp:Panel>

<asp:Panel id="pnlBody" runat="server">
	<div id="buttons">
		<amp:button id="btnApprove" resx="ApproveAsset" visible="false" runat="server" />
		<amp:button id="btnEdit" resx="Edit" visible="false" runat="server" />
		<amp:button id="btnDelete" resx="Delete" visible="false" runat="server" />
		<amp:Button id="btnVideo" resx="PreviewVideo" visible="false" runat="server" />
		<amp:Button id="btnView" visible="false" runat="server" />
	</div>
	<div id="meta">
		<amp:Rating id="ampRating" radius="15" runat="server" visible="false" />
		<asp:Label id="lblViews" runat="server" cssClass="views" visible="false" />
		<asp:repeater id="rptCategories" runat="server">
			<headertemplate><div class="list">Categories:<ul></headertemplate>
			<itemtemplate><li><span><%# DirectCast(Container.DataItem, AMP.Category).SearchLink %></span></li></itemtemplate>
			<footertemplate></ul></div></footertemplate>
		</asp:repeater>
		<asp:repeater id="rptPlugins" runat="server">
			<headertemplate><div class="list">Plugins:<ul></headertemplate>
			<itemtemplate><li><span><%# DirectCast(Container.DataItem, AMP.Software).NameLink %></span></li></itemtemplate>
			<footertemplate></ul></div></footertemplate>
		</asp:repeater>
	</div>
	<h2 class="asset"><asp:label id="lblType" runat="server" /></h2>
	<h1 class="asset"><asp:Label id="lblName" runat="server" /></h1>
	<asp:Label id="lblUser" runat="server" cssClass="assetAuthor" />
	<asp:Label ID="lblDate" Runat="server" CssClass="assetDate" />
	
	<asp:Label id="lblDescription" runat="server" cssClass="description" />
	
	<fieldset id="rating">
		<legend class="smaller">Feedback</legend>
		<asp:Panel id="pnlRateIt" runat="server" visible="False" cssclass="newComment">
			<span class="saveComment"><amp:Button id="btnSaveComment" resx="SaveComment" visible="true" submitform="true" runat="server" /></span>
			<div id="label">Your rating: <asp:dropdownlist id="ddlRating" runat="server" EnableViewState="True" /> stars</div>
			<amp:validation target="tbComment" type="PlainText" message="Comment (plain text only)" required="true" runat="server" /> 
			<asp:textbox id="tbComment" runat="server" textmode="MultiLine" />
		</asp:panel>
		<asp:Repeater id="rptComments" Runat="server">
			<HeaderTemplate><table class="comments" cellspacing="0" cellpadding="0"></HeaderTemplate>
			<ItemTemplate>
				<tr class="comment">
					<td class="about">
						<amp:rating rating="<%# DirectCast(Container.DataItem, AMP.Rating).Value %>" runat="server" id="Rating1" name="Rating1"/><br/>
						<%# DirectCast(Container.DataItem, AMP.Rating).Person.DetailLink %>
					</td>
					<td class="text">
						<div class="date"><%# String.Format("{0:MMMM d, yyyy}", DirectCast(Container.DataItem, AMP.Rating).Date) %></div>
						<%# DirectCast(Container.DataItem, AMP.Rating).Comment %>
					</td>
				</tr>
			</ItemTemplate>
			<FooterTemplate></table></FooterTemplate>
		</asp:Repeater>
		
		<input type="hidden" id="fldAction" runat="server" />
	</fieldset>
</asp:Panel>

</body></html>