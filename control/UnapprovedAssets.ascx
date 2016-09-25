<%@ Control Language="vb" AutoEventWireup="false" Codebehind="UnapprovedAssets.ascx.vb" Inherits="AMP.Controls.UnapprovedAssets" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<%@ Import Namespace="AMP" %> 

<asp:repeater id="rptAssets" runat="server">
	<headertemplate><div class="newAssets"></headertemplate>
<itemtemplate><div class="row">
	<div class="date"><%# String.Format("{0:M/dd/yy}", DirectCast(Container.DataItem, AMP.Asset).SubmitDate) %></div>
	<div class="name assetType<%# CType(Container.DataItem, AMP.Asset).Type %>"><%# DirectCast(Container.DataItem, AMP.Asset).DetailLink %></div>
</div></itemtemplate>
	<footertemplate></div></footertemplate>
</asp:repeater>