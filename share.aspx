<%@ Page Language="vb" AutoEventWireup="false" Codebehind="share.aspx.vb" Inherits="AMP.Pages.Share" EnableViewState="true"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="ContentWebPart" Src="~/control/Content.ascx"%>
<%@ Register TagPrefix="amp" TagName="SignInUp" Src="~/control/SignInUp.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:Panel id="pnlBody" runat="server">
	<h2 id="h2Step" class="form" runat="server" />

<!-- Step 1: select resource -->
	
	<fieldset class="wizard" id="fsStep1" visible="false" runat="server">
		<legend class="title"><%=Say("Title.Share1")%></legend>
		<fieldset>
			<legend>A Project</legend>
			<amp:Upload ID="fldUpload" Resx="File" Runat="server" />
			<amp:Field ID="fldTerms" Resx="File" Type="checkbox" Runat="server" />
			<div id="TermsConditions"><amp:Content file="terms.html" runat="server" /></div>
		</fieldset>
	
		<div id="or">or</div>
		
		<fieldset>
			<legend>An Article</legend>
			<amp:Field type="text" resx="WebSite" id="fldWebSite" shownote="false"
				value="http://" Validate="URL" Style="width: 20em;" runat="server" />
		</fieldset>
		<input type="hidden" name="step" value="1" />
	</fieldset>

<!-- Step 2: describe resource -->

	<fieldset class="wizard" id="fsStep2" visible="false" runat="server">
		<legend class="title"><%=Say("Title.Share2")%></legend>
		
		<amp:Field ID="fldTitle" Resx="Title" Style="width: 20em;" Required="True" Type="text"
			Validate="PlainText" MaxLength="50" Runat="server" /><br/>
		
		<asp:Panel ID="pnlSoftware" Runat="server">
			<label class="required"><%=Say("Label.SoftwareSelection")%></label>
			<amp:validation target="ampSoftware" type="Select" message="Software Selection" required="true" runat="server" />
			<amp:SoftwareList id="ampSoftware" runat="server" />
		</asp:Panel>
		
		<amp:Field Type="checkbox" ID="fldActivate" Resx="Active" Visible="False" Runat="server" />
		
		<fieldset id="description">
			<legend><%=Say("Label.Description")%></legend>
			<amp:validation target="tbDescription" type="PlainText" message="Description (plain text only)" required="true" runat="server" />
			<asp:textbox id="tbDescription" maxlength="1000" textmode="MultiLine" runat="server" cssclass="description" width="95%" />
			<div class="note"><%=Say("Note.PlainText")%></div>
		</fieldset>
		
		<fieldset id="fsOptionalLinks" class="links" runat="server">
			<legend><%=Say("Label.ProjectLinks")%></legend>
			<amp:Field ID="fldRenderedURL" Type="text" Value="http://" Validate="URL"
				MaxLength="150" Resx="RenderedUrl" Runat="server" />
			<amp:Field ID="fldMediaURL" Type="text" Value="http://" Validate="URL"
				MaxLength="150" Resx="MediaUrl" Runat="server" />
		</fieldset>
	
		<input type="hidden" name="step" value="2" />
	</fieldset>

<!-- Step 3: classify resource -->

	<fieldset class="wizard" id="fsStep3" visible="false" runat="server">
		<legend class="title"><%=Say("Title.Share3")%></legend>
		
		<fieldset class="categories">
			<legend><%=Say("Label.Categories")%></legend>
			<amp:validation target="lbCategories" type="Select" message="Category (choose at least one)" required="true" runat="server" ID="Validation1" NAME="Validation1"/>
			<amp:listbox id="lbCategories" datatextfield="Name"
				selectionmode="Multiple" datavaluefield="ID" rows="5" runat="server" />
		</fieldset>
		
		<fieldset id="fsPlugins" class="plugins" visible="false" runat="server">
			<legend><%=Say("Label.Plugins")%></legend>
			<amp:listbox id="lbPlugins" datatextfield="FullName"
				selectionmode="Multiple" datavaluefield="ID" rows="5" runat="server" />
		</fieldset>
		
		<asp:Panel ID="pnlTypes" Visible="False" Runat="server">
			<label><%=Say("Label.ArticleType")%></label>
				<amp:validation target="ampTypeList" type="Select" message="Resource Type" required="true" runat="server" />
				<amp:EnumList id="ampTypeList" runat="server" />
		</asp:Panel>
		
		<fieldset id="fsSections" class="sections" visible="false" runat="server">
			<legend><%=Say("Label.Sections")%></legend>
			<amp:validation target="ampSectionList" type="Select" message="Section" required="true" runat="server" />
			<amp:EnumList id="ampSectionList" ShowNote="False" runat="server" rows="4" />
		</fieldset>

		<div id="multinote" class="note"><%=Say("Note.MultiSelectNote")%></div>
			
		<input type="hidden" name="step" value="3" />
	</fieldset>

<!-- Step 4: login or register -->
	
	<amp:SignInUp id="ampSignInUp" runat="server" visible="false" Redirect="false" />
	
	<div class="nav">
		<amp:Button id="btnPrevious" resx="Previous" runat="server" visible="false" />
		<amp:Button id="btnNext" resx="Next" runat="server" submitform="true" visible="false" />
	</div>
</asp:Panel>

<asp:Panel id="pnlWindows" runat="server">
	<amp:ContentWebPart id="ampContent" runat="server" />
</asp:Panel>

</body>
</html>
