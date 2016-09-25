<%@ Page Language="vb" AutoEventWireup="false" Codebehind="resource-edit.aspx.vb" Inherits="AMP.Pages.ResourceEdit"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="NewAssets" Src="~/control/NewAssets.ascx"%>
<%@ Register TagPrefix="amp" TagName="Tours" Src="~/control/Tours.ascx"%>
<%@ Register TagPrefix="amp" TagName="Contests" Src="~/control/Contests.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:panel id="pnlBody" runat="server">
	<fieldset runat="server">
		<legend>Resource</legend>
	
		<div id="classification">
			<fieldset>
				<legend><%=Say("Label.Categories")%></legend>
				<amp:validation target="lbCategories" type="Select" message="Category" required="true" runat="server" />
				<amp:listbox id="lbCategories" shownote="false" showlink="false" datatextfield="Name" selectionmode="Multiple" datavaluefield="ID" rows="6" runat="server" />
			</fieldset>
			<fieldset id="fsPlugins" runat="server">
				<legend><%=Say("Label.Plugins")%></legend>
				<amp:listbox id="lbPlugins" shownote="false" showlink="false" datatextfield="FullName" selectionmode="Multiple" datavaluefield="ID" rows="5" runat="server" />
			</fieldset>
			<fieldset id="fsSections" runat="server">
				<legend><%=Say("Label.Sections")%></legend>
				<amp:validation target="ampSectionList" type="Select" message="Section" required="true" runat="server" />
				<amp:EnumList id="ampSectionList" shownote="false" showlink="false" runat="server" rows="5" />
			</fieldset>
		</div>
		
		<amp:field id="fldTitle" type="text" style="width: 40%;" resx="Title"
			validate="PlainText" maxlength="50" required="true" shownote="false" runat="server" /><br/>
			
		<asp:panel id="pnlFiles" runat="server" visible="False">
			<amp:Upload ID="fldFile" Resx="File" Runat="server" />
				<div class="note">leave blank except to upload a new version</div>
				
			<label class="required"><%=Say("Label.SoftwareSelection")%></label>
				<amp:validation target="ampSoftware" type="Select" message="Software Selection" required="true" runat="server" />
				<amp:softwarelist id="ampSoftware" runat="server" />
				
			<fieldset id="links"><legend><%=Say("Label.ProjectLinks")%></legend>
				<amp:field id="fldRenderedURL" maxlength="150" type="text" resx="RenderedUrl"
					validate="URL" value="http://" runat="server" />
				<amp:field id="fldMediaURL" maxlength="150" type="text" resx="MediaUrl"
					validate="URL" value="http://" runat="server" />
			</fieldset>
				
		</asp:panel>
		
		<asp:panel id="pnlLinks" runat="server" visible="False">
			<amp:Field Type="text" ID="fldLink" Resx="WebSite" Required="True" Validate="ActiveURL"
				value="http://" MaxLength="150" style="width: 40%;" Runat="server" />
			<label class="required"><%=Say("Label.ArticleType")%></label>
				<amp:validation target="ampTypeList" type="Select" message="Resource Type" required="true" runat="server" />
				<amp:EnumList id="ampTypeList" runat="server" />
			<amp:Field ID="fldScrape" type="checkbox" Resx="ScrapeLink" Runat="server" />
		</asp:panel>
		
		<fieldset id="description">
			<legend><%=Say("Label.Description")%></legend>
			<amp:validation target="tbDescription" type="PlainText" message="Description (plain text only)" required="true" runat="server" />
			<asp:textbox id="tbDescription" maxlength="1000" textmode="MultiLine" rows="5" runat="server" />
			<br/><span class="note" style="padding:0;"><%=Say("Note.PlainText")%></span>
		</fieldset>
	</fieldset>
	
	<div id="buttons">
		<amp:button id="btnSave" runat="server" resx="Save" submitform="true" />
		<amp:button id="btnCancel" runat="server" resx="Cancel" onclick="history.back()" />
	</div>
</asp:panel>

<asp:panel id="pnlWindows" runat="server">
	<amp:contests id="Contests" runat="server" />
	<amp:newassets count="10" id="ampNewListings" runat="server" />
	<amp:tours newerthan="1/1/04" id="ampTours" runat="server" />
</asp:panel>

</body>
</html>