<%@ Control Language="vb" AutoEventWireup="false" Codebehind="Statistics.ascx.vb" Inherits="AMP.Controls.Statistics" TargetSchema="http://www.w3.org/1999/xhtml" %>

<div id="statistics">
	<dl>
		<dt>Users</dt><dd><asp:label id="lblUserCount" runat="server" /></dd>
		<dt>Resources</dt><dd><asp:label id="lblAssetCount" runat="server" /></dd>
		<dt>Products</dt><dd><asp:label id="lblProductCount" runat="server" /></dd>
		<dt>Started</dt><dd><asp:label id="lblAppStart" runat="server" /></dd>
	</dl>
	
	<div class="section">Server</div>
	<dl>
		<dt>OS</dt><dd><asp:label id="lblOS" runat="server" /></dd>
		<dt>Name</dt><dd><asp:label id="lblServerName" runat="server" /></dd>
		<dt>Web</td><dd><asp:label id="lblIIS" runat="server" /></dd>
		<dt>ASP.NET</td><dd><asp:label id="lblAspNetVersion" runat="server" /></dd>
		<dt>User</td><dd><asp:label id="lblServerUser" runat="server" /></dd>
	</dl>
	
	<div class="section">Data File</div>
	<div id="dataFile">
		<span style="font-size: 8pt;"><asp:hyperlink ID="lnkDataFile" Runat="server" /></span><br/>
		<asp:label id="lblDataSize" runat="server" /><br/>
		Saved <asp:label id="lblDataSaved" runat="server" />
	</div>
	
	<div class="section">Today</div>
	<dl>
		<dt>Visits</dt><dd><asp:label id="lblVisitsToday" runat="server" /></dd>
		<dt>Sales</dt><dd><asp:label id="lblSales" runat="server" /></dd>
	</dl>
</div>
