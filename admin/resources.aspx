<%@ Page Language="vb" AutoEventWireup="false" Codebehind="resources.aspx.vb" Inherits="AMP.Pages.Administration.Resources" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:panel id="pnlBody" runat="server">
	<h1>Unapproved Resources</h1>
	<br/>
	<asp:Repeater ID="rptUnapproved" Runat="server">
		<HeaderTemplate><table id="list" cellspacing="0" cellpadding="0"></HeaderTemplate>
		<ItemTemplate>
		<tr>
			<td class="left">
				<div class="type assetType<%# DirectCast(Container.DataItem, AMP.Asset).Type %>"><%# DirectCast(Container.DataItem, AMP.Asset).FullType %></div>
				<div class="title"><%# DirectCast(Container.DataItem, AMP.Asset).DetailLink %></div>
			</td><td class="right">
				<div class="author">By <%# DirectCast(Container.DataItem, AMP.Asset).AuthoredBy.DetailLink %></div>
				<div class="date"><%# String.Format("{0:MMMM dd, yyyy}", DirectCast(Container.DataItem, AMP.Asset).VersionDate) %></div>
				<div class="action">
					<amp:button resx="ApproveAsset" runat="server" onclick="<%# Me.ApproveLink(Container.DataItem) %>" /><br/>
					<amp:button resx="DenyAsset" runat="server" onclick="<%# Me.DenyLink(Container.DataItem) %>" />
				</div>
				<div class="description"><%# DirectCast(Container.DataItem, AMP.Asset).Description %></div>
			</td>
		</tr><tr>
			<td class="left bottom">
				<amp:Button text="<%#Say(DirectCast(Container.DataItem, AMP.Asset).ViewAction)%>" runat="server" onclick="<%# Me.ViewLink(Container.DataItem) %>" />
			</td><td class="right bottom">
				<amp:button resx="Edit" runat="server" onclick="<%# Me.EditLink(Container.DataItem) %>" />
				<amp:repeater id="rptCategories" runat="server" DataSource="<%# DirectCast(Container.DataItem, AMP.Asset).Categories %>">
					<headertemplate><div class="list"><%=Say("Label.Categories")%>: </headertemplate>
					<itemtemplate><span class="item"><%# DirectCast(Container.DataItem, AMP.Category).Name %></span></itemtemplate>
					<separatortemplate>, </separatortemplate>
					<footertemplate></div></footertemplate>
				</amp:repeater>
				<amp:repeater id="rptPlugins" runat="server" DataSource="<%# DirectCast(Container.DataItem, AMP.Asset).Plugins() %>">
					<headertemplate><div class="list"><%=Say("Label.Plugins")%>: </headertemplate>
					<itemtemplate><span class="item"><%# DirectCast(Container.DataItem, AMP.Software).FullNameLink %></span></itemtemplate>
					<separatortemplate>, </separatortemplate>
					<footertemplate></div></footertemplate>
				</amp:repeater>
			</td>	
		</tr>
		</ItemTemplate>
		<FooterTemplate></table></FooterTemplate>
	</asp:Repeater>
	
	<input type="hidden" id="fldAction" runat="server" />
	<input type="hidden" id="fldAssetID" runat="server" />
</asp:panel>

<asp:panel id="pnlWindows" runat="server" />

</body>
</html>