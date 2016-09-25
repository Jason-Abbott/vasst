<%@ Page Language="vb" AutoEventWireup="false" Codebehind="products.aspx.vb" Inherits="AMP.Pages.Administration.Products" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:panel id="pnlBody" runat="server">
	<fieldset runat="server">
		<legend>Product</legend>
		
		<div id="classification">
			<fieldset>
				<legend><%=Say("Label.Categories")%></legend>
				<amp:validation target="lbCategories" type="Select" message="Category" required="true" runat="server" ID="Validation1" NAME="Validation1"/>
				<amp:listbox id="lbCategories" shownote="false" showlink="false" datatextfield="Name" selectionmode="Multiple" datavaluefield="ID" rows="6" runat="server" />
			</fieldset>
			<fieldset>
				<legend><%=Say("Label.Sections")%></legend>
				<amp:validation target="ampSectionList" type="Select" message="Section" required="true" runat="server" ID="Validation2" NAME="Validation2"/>
				<amp:EnumList id="ampSectionList" shownote="false" showlink="false" runat="server" rows="5" />
			</fieldset>
		</div>
		
		<amp:field id="fldTitle" type="text" style="width: 40%;" resx="Title"
			validate="PlainText" maxlength="50" required="true" shownote="false" runat="server" /><br/>
	
		<amp:Upload ID="fldPicture" Resx="Picture" Runat="server" ShowNote="False" />
		
		<fieldset id="dates">
			<legend><%=Say("Title.Dates")%></legend>
			<amp:field id="fldShowOn" type="text" resx="StartDate" maxlength="10"
				validate="Date" required="true" inline="true" runat="server" /><br/>
			<amp:field id="fldHideAfter" type="text" resx="EndDate" maxlength="10"
				validate="Date" required="true" inline="true" runat="server" /><br/>
		</fieldset>
	
		<fieldset id="description">
			<legend><%=Say("Label.Description")%></legend>
			<amp:validation target="tbDescription" type="PlainText" message="Description (plain text only)" required="true" runat="server" ID="Validation3" NAME="Validation3"/>
			<asp:textbox id="tbDescription" maxlength="1000" textmode="MultiLine" rows="5" runat="server" />
			<br/><span class="note" style="padding:0;"><%=Say("Note.PlainText")%></span>
		</fieldset>
	
	</fieldset>
	
	<div id="buttons">
		<amp:button id="btnSave" runat="server" resx="Save" submitform="true" />
		<amp:button id="btnCancel" runat="server" resx="Cancel" onclick="history.back()" />
	</div>
</asp:panel>

<asp:panel id="pnlWindows" runat="server" />
</body>
</html>