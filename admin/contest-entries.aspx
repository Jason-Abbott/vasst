<%@ Page Language="vb" AutoEventWireup="false" Codebehind="contest-entries.aspx.vb" Inherits="AMP.Pages.Administration.ContestEntries" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:panel id="pnlBody" runat="server">
	<fieldset>
		<legend>Contest Entries</legend>

		<asp:Repeater ID="rptContests" Runat="server">
			<ItemTemplate>
			<asp:repeater id="rptUnapproved" runat="server" DataSource="<%# DirectCast(Container.DataItem, AMP.Contest).Entries.Unapproved %>">
				<HeaderTemplate><div id="entries"></HeaderTemplate>
				<itemtemplate>
					<div>
						"<%# DirectCast(Container.DataItem, AMP.ContestEntry).ViewLink %>"
						from
						<%# DirectCast(Container.DataItem, AMP.ContestEntry).Contestant.DetailLink %>
						<amp:button resx="ApproveAsset" runat="server" onclick="<%# Me.ApproveLink(Container.DataItem) %>" />
						<amp:button resx="DenyAsset" runat="server" onclick="<%# Me.DenyLink(Container.DataItem) %>" />
					</div>
				</itemtemplate>
				<FooterTemplate></div></FooterTemplate>
			</asp:repeater>
			</ItemTemplate>
		</asp:Repeater>
	
		<input type="hidden" id="fldAction" runat="server" />
		<input type="hidden" id="fldEntryID" runat="server" />
	</fieldset>
</asp:panel>

<asp:panel id="pnlWindows" runat="server"/>
</body>
</html>