<%@ Control Language="vb" AutoEventWireup="false" Codebehind="NewAssets.ascx.vb" Inherits="AMP.Controls.NewAssets" TargetSchema="http://www.w3.org/1999/xhtml" %>
<%@ OutputCache Duration="300" VaryByControl="Count" VaryByCustom="browser, section" %>
<%@ Import Namespace="AMP" %> 

<asp:Repeater id="rptAssets" runat="server">
	<HeaderTemplate><div class="newAssets"></HeaderTemplate>
<ItemTemplate><div class="row">
	<div class="date"><%# String.Format("{0:M/dd/yy}", DirectCast(Container.DataItem, AMP.Asset).VersionDate) %></div>
	<div class="name assetType<%# CType(Container.DataItem, AMP.Asset).Type %>"><%# DirectCast(Container.DataItem, AMP.Asset).DetailLink %></div>
</div></ItemTemplate>
	<FooterTemplate></div></FooterTemplate>
</asp:Repeater>