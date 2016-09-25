<%@ Control Language="vb" AutoEventWireup="false" Codebehind="SearchResults.ascx.vb" Inherits="AMP.Controls.SearchResults" TargetSchema="http://www.w3.org/1999/xhtml" %>

<asp:Repeater id="rptResults" runat="server">
	<HeaderTemplate><div class="newAssets"></HeaderTemplate>
<ItemTemplate><div class="row">
	<div class="date"><%# String.Format("{0:M/dd/yy}", DirectCast(Container.DataItem, AMP.Asset).SubmitDate) %></div>
	<div class="name assetType<%# DirectCast(Container.DataItem, AMP.Asset).Type %>"><%# DirectCast(Container.DataItem, AMP.Asset).DetailLink %></div>
</div></ItemTemplate>
	<FooterTemplate></div></FooterTemplate>
</asp:Repeater>