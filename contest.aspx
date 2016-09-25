<%@ Page Language="vb" AutoEventWireup="false" Codebehind="contest.aspx.vb" Inherits="AMP.Pages.Contest" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="NewAssets" Src="~/control/NewAssets.ascx"%>
<%@ Register TagPrefix="amp" TagName="Tours" Src="~/control/Tours.ascx"%>
<%@ Register TagPrefix="amp" TagName="Contests" Src="~/control/Contests.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:Panel id="pnlBody" runat="server">

	<fieldset id="fsBallot" class="ballot" runat="server" visible="false">
		<legend class="smaller"><%=Say("Title.Vote")%></legend>
		<div id="ballot">
			<asp:Label ID="lblNoVote" cssclass="noVote" Runat="server" Visible="False">
				Check back soon to vote when there are at least two entries
			</asp:Label>
			<asp:Panel ID="pnlCanVote" Runat="server">
				<asp:Label ID="lblInstructions" Runat="server" />
				<amp:Ballot id="ampBallot" runat="server" />
				<amp:Button id="btnVote" runat="server" resx="Vote" submitform="true" />
			</asp:Panel>
		</div>
	</fieldset>
	
	<asp:Repeater ID="rptPrizes" Visible="false" Runat="server">
		<HeaderTemplate><dl id="prizes"></HeaderTemplate>
		<ItemTemplate>
			<dt><%#AMP.Format.RankForNumber(Container.ItemIndex + 1, True)%> Prize</dt>
			<dd><%#Container.DataItem.toString%><dd>
		</ItemTemplate>
		<FooterTemplate></dl></FooterTemplate>
	</asp:Repeater>
	
	<asp:Repeater ID="rptRules" Visible="false" Runat="server">
		<HeaderTemplate><div id="rules"><div>Contest Rules</div><ul></HeaderTemplate>
		<ItemTemplate><li><%#Container.DataItem.toString%></li></ItemTemplate>
		<FooterTemplate></ul></div></FooterTemplate>
	</asp:Repeater>
	
	<h1 class="title"><asp:Label ID="lblTitle" Runat="server" /></h1>
	<asp:Label ID="lblAbout" CssClass="about" Runat="server" Visible="False" />
	<asp:Label ID="lblDescription" CssClass="description" Runat="server" />
	
	<div id="contestants">
		<asp:label ID="lblStatus" cssclass="heading" Runat="server" />
		<asp:repeater id="rptVotes" runat="server">
			<headertemplate><table id="votes" cellpadding="0" cellspacing="0"></headertemplate>
			<itemtemplate>
				<tr>
					<td class="title"><%#DirectCast(Container.DataItem, AMP.ContestEntry).ViewLink%></td>
					<td class="graph">
						<div class="bar" style="width: <%#(Me.Entity.PercentOfHighest(DirectCast(Container.DataItem, AMP.ContestEntry).Votes) * 0.8) + 10%>%;">
						<%#Me.Entity.VotePercent(DirectCast(Container.DataItem, AMP.ContestEntry).Votes, 1) %>%
						</div>
						<%#Me.Entity.Points(DirectCast(Container.DataItem, AMP.ContestEntry).Votes)%>
					</td>
				</tr>
			</itemtemplate>
			<footertemplate></table></footertemplate>
		</asp:repeater>
	</div>
	
	<div id="buttons">
		<amp:Button id="btnEnter" runat="server" resx="EnterContest" visible="false" />
		<amp:Button id="btnEdit" runat="server" resx="Edit" visible="false" />
	</div>
</asp:Panel>

<asp:Panel id="pnlWindows" runat="server">
	<amp:Contests id="Contests" runat="server" />
	<amp:NewAssets count="10" id="smgNewListings" runat="server" />
	<amp:Tours NewerThan="1/1/04" id="smgTours" runat="server" />
</asp:Panel>

</body>
</html>