<%@ Page Language="vb" AutoEventWireup="false" Codebehind="categories.aspx.vb" Inherits="AMP.Pages.Administration.Categories" %>
<%@ Register TagPrefix="amp" TagName="ContentWebPart" Src="~/control/Content.ascx"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:panel id="pnlBody" runat="server">
	<fieldset runat="server">
		<legend>Category Administration</legend>
		<fieldset id="categories">
			<legend>Categories</legend>
			<asp:Repeater ID="rptCategories" Runat="server">
				<HeaderTemplate>
					<div id="categoryBox">
						<input type="text" id="new" maxlength="30" onclick="Category.Add(this)" value="Add New">
				</HeaderTemplate>
				<ItemTemplate>
					<input type="text" onclick="Category.Select(this)"
						maxlength="30"
						id="<%#DirectCast(Container.DataItem, AMP.Category).ID%>"
						value="<%#DirectCast(Container.DataItem, AMP.Category).Name%>">
				</ItemTemplate>
				<FooterTemplate></div></FooterTemplate>
			</asp:Repeater>
		</fieldset>
		
		<fieldset id="sections"><legend>In <%=Say("Title.Sections")%></legend>
			<amp:enumcheckbox id="ampSectionList" name="section" runat="server" />
		</fieldset>
	
		<fieldset id="entities"><legend>Valid For</legend>
			<amp:EntityList id="ampEntityList" runat="server" />
		</fieldset>
	</fieldset>
	
	<div id="buttons">
		<amp:button id="btnDelete" runat="server" onclick="Category.Delete()" resx="Delete" visible="true" />
		<amp:button id="btnSave" runat="server" onclick="Category.Save()" resx="Save" visible="true" />
	</div>
</asp:panel>

<asp:panel id="pnlWindows" runat="server" >
	<amp:contentwebpart file="help/Categories.txt" title="Category Tips" id="ampContent" runat="server" />
</asp:panel>
</body>
</html>