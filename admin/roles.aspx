<%@ Page Language="vb" AutoEventWireup="false" Codebehind="roles.aspx.vb" Inherits="AMP.Pages.Administration.Roles" %>
<%@ Register TagPrefix="amp" TagName="ContentWebPart" Src="~/control/Content.ascx"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:panel id="pnlBody" runat="server">
	<fieldset id="roleAdmin">
		<legend>Role Administration</legend>

		<fieldset id="permissions">
			<legend>Permissions</legend>
			<amp:permissionBox id="smgPermissionBox" runat="server"/>
		</fieldset>
		
		<fieldset id="roles">
			<legend>Roles</legend>
			<amp:roletree id="smgRoleTree" runat="server"/>
		</fieldset>
	</fieldset>
</asp:panel>

<asp:panel id="pnlWindows" runat="server">
	<amp:contentwebpart file="help/Roles.txt" title="Tips & Help" id="ampContent" runat="server" />
</asp:panel>

</body>
</html>