<%@ Page Language="vb" AutoEventWireup="false" Codebehind="search.aspx.vb" Inherits="AMP.Pages.Search" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="Contests" Src="~/control/Contests.ascx"%>
<%@ Register TagPrefix="amp" TagName="Results" Src="~/control/SearchResults.ascx"%>
<%@ OutputCache Duration="300" VaryByParam="*" VaryByCustom="browser,section,message,role" %>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:Panel id="pnlWindows" runat="server">
	<amp:Contests id="Contests" runat="server" />
	<amp:Results id="SearchResults" runat="server" />
</asp:Panel>

<asp:Panel id="pnlBody" runat="server">
	<h2><asp:Label ID="lblCount" runat="server" Visible="False" /></h2>
	<h1><asp:Label ID="lblMatches" Runat="server" Visible="False" /></h1>
	
	<asp:Panel cssclass="tip" Runat="server" ID="pnlTip" Visible="False">
		<div id="line1">Not seeing what you're after?</div>
		<div id="line2">Try the search.  It will find matches in a</div>
		<div id="line3">resource title, description, author's name and categories</div>
	</asp:Panel>

	<asp:Repeater ID="rptResults" Runat="server" Visible="False">
		<HeaderTemplate><table id="list" cellspacing="0" cellpadding="0"></HeaderTemplate>
		<ItemTemplate>
		<tr>
			<td class="left top">
				<div class="type assetType<%# DirectCast(Container.DataItem, AMP.Asset).Type %>"><%# DirectCast(Container.DataItem, AMP.Asset).FullType %></div>
				<div class="title"><%# Me.Highlight(DirectCast(Container.DataItem, AMP.Asset).DetailLink) %></div>
				<div class="views"><%# DirectCast(Container.DataItem, AMP.Asset).Views() %></div>
				<amp:Rating radius="7" runat="server" cssClass="rating" visible="true" rating="<%# DirectCast(Container.DataItem, AMP.Asset).Ratings.Average %>" /><br/>
			</td><td class="right top">
				<div class="author">By <%# Me.Highlight(DirectCast(Container.DataItem, AMP.Asset).AuthoredBy.DetailLink) %></div>
				<div class="date"><%# String.Format("{0:MMMM dd, yyyy}", DirectCast(Container.DataItem, AMP.Asset).VersionDate) %></div>
				<div class="description"><%# Me.Highlight(DirectCast(Container.DataItem, AMP.Asset).Description) %></div>
			</td>
		</tr><tr>
			<td class="left bottom">
				<amp:Button text="<%#Say(DirectCast(Container.DataItem, AMP.Asset).ViewAction)%>" runat="server" onclick="<%# Me.ViewLink(Container.DataItem) %>" />
			</td><td class="right bottom">
				<amp:button resx="Edit" runat="server" visible="<%# DirectCast(Container.DataItem, AMP.Asset).CanEdit %>" onclick="<%# Me.EditLink(Container.DataItem) %>" />
				<amp:repeater id="rptCategories" runat="server" DataSource="<%# DirectCast(Container.DataItem, AMP.Asset).Categories %>">
					<headertemplate><div class="list"><%=Say("Label.Categories")%>: </headertemplate>
					<itemtemplate><span class="item"><%# DirectCast(Container.DataItem, AMP.Category).SearchLink %></span></itemtemplate>
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
	
	<div id="buttons">
		<amp:Button ID="btnPrevious" Resx="Previous" Runat="server" Visible="False" />
		<amp:Button ID="btnNext" Resx="Next" Runat="server" Visible="False" />
	</div>
	
</asp:Panel>

</body></html>