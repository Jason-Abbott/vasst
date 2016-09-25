<%@ Control Language="vb" AutoEventWireup="false" Codebehind="Contests.ascx.vb" Inherits="AMP.Controls.Contests" TargetSchema="http://www.w3.org/1999/xhtml" %>
<%@ OutputCache Duration="300" VaryByParam="none" VaryByCustom="browser, section" %>
<%@ Import Namespace="AMP" %>

<asp:Repeater id="rptContests" runat="server">
	<HeaderTemplate><ul id="contests"></HeaderTemplate>
	<ItemTemplate>
		<li><%# DirectCast(Container.DataItem, AMP.Contest).DetailLink %>
			<div class="notice"><%# Me.Subtext(Container.DataItem) %></div>
		</li>
	</ItemTemplate>
	<FooterTemplate></ul></FooterTemplate>
</asp:Repeater>