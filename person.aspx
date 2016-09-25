<%@ Page Language="vb" AutoEventWireup="false" Codebehind="person.aspx.vb" Inherits="AMP.Pages.Person" EnableViewState="true"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="NewAssets" Src="~/control/NewAssets.ascx"%>
<%@ Register TagPrefix="amp" TagName="Contests" Src="~/control/Contests.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:Panel id="pnlWindows" runat="server">
	<amp:Contests id="Contests" runat="server" />
	<amp:NewAssets count="10" id="ampNewListings" runat="server" />
</asp:Panel>

<asp:Panel id="pnlBody" runat="server">
	<amp:Image id="imgPerson" CssClass="picture" runat="server" visible="false" /><br/>

	<div id="name">
		<h2><asp:Label ID="lblRealName" Runat="server" /></h2>
		<h1><asp:Label ID="lblDisplayName" Runat="server" /></h1>
		<asp:Label ID="lblRole" cssclass="role" Runat="server" />
	</div>
	<div id="links">
		<asp:Label ID="lblEmail" CssClass="email" Runat="server" />
		<asp:Label ID="lblWebSite" cssclass="website" Runat="server" />
	</div>
	
	<asp:Label ID="lblDescription" CssClass="description" Runat="server" />
	
	<asp:Label ID="lblResources" Runat="server" Visible="False" CssClass="resourceLabel" />
	<asp:Panel ID="pnlResources" Runat="server" Visible="False" CssClass="resources">
		<asp:Repeater ID="rptResources" Runat="server">
		<ItemTemplate><div class="resource assetType<%# DirectCast(Container.DataItem, AMP.Asset).Type %>">
		<%# DirectCast(Container.DataItem, AMP.Asset).DetailLink %>
		</div></ItemTemplate>
		</asp:Repeater>
	</asp:Panel>
	
	<div id="buttons">
		<amp:button id="btnEdit" runat="server" resx="Edit" visible="false" />
	</div>
</asp:Panel>

</body>
</html>