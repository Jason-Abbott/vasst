<%@ Page Language="vb" AutoEventWireup="false" Codebehind="contest-enter.aspx.vb" Inherits="AMP.Pages.EnterContest" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="ContentWebPart" Src="~/control/Content.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:Panel id="pnlBody" runat="server">
	<h2 runat="server"><%=Me.Say("Action.EnterContest").toLower%></h2>
	<h1><asp:Label ID="lblTitle" Runat="server" /></h1>
	
	<asp:Label ID="lblDescription" CssClass="description" Runat="server" />
	<asp:Label ID="lblLimit" CssClass="limit" Runat="server" />
	
	<asp:Panel ID="pnlFileUpload" visible="False" Runat="server">
		<fieldset id="upload">
			<legend class="smaller">Upload Entry</legend>
			
			<amp:upload id="ampUpload" resx="File" required="true"
				style="width: 15em;" runat="server" />
			<amp:field id="fldTitle" type="text" resx="Title" validate="PlainText"
				required="true" style="width: 15em;" maxlength="50" runat="server" />
				
			<amp:Field ID="fldRules" Resx="File" Type="checkbox" Runat="server" />
			<div id="rules">
				<asp:Repeater ID="rptRules" Runat="server">
					<HeaderTemplate><ul></HeaderTemplate>
					<ItemTemplate><li><%#Container.DataItem.toString%></li></ItemTemplate>
					<FooterTemplate></ul></FooterTemplate>
				</asp:Repeater>
			</div>
			
			<amp:Button id="btnUpload" runat="server" resx="Upload" submitform="true" />
			<amp:button id="btnCancel" runat="server" resx="Cancel" onclick="history.back()" />
		</fieldset>
	</asp:Panel>
	
	<!--
	<ul>
	<li>if media entry type
		<ul>
		<li>upload directly from here
		<li>can't be a shared resource
		</ul>
	<li>if project entry type
		<ul>
		<li>if existing assets qualify, show form to add those
		<li>if creating a new entry, show option:
			<ul>
			<li>use this entry only for this contest
			<li>if anonymous contest, second option is
				<ul><li>make this entry a shared resource when the contest is done</ul>
			<li>else
				<ul><li>make this entry a shared resource</ul>
			</ul>
		</ul>
	</ul>
	-->
</asp:Panel>

<asp:Panel id="pnlWindows" runat="server">
	<amp:contentwebpart file="help/EnterContest.txt" title="Tips & Help" id="ampContent" runat="server" />
</asp:Panel>

</body>
</html>