<%@ Page Language="vb" AutoEventWireup="false" Codebehind="account.aspx.vb" Inherits="AMP.Pages.Account" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<%@ Register TagPrefix="amp" TagName="NewAssets" Src="~/control/NewAssets.ascx"%>
<%@ Register TagPrefix="amp" TagName="Tours" Src="~/control/Tours.ascx"%>
<%@ Register TagPrefix="amp" TagName="Contests" Src="~/control/Contests.ascx"%>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:Panel id="pnlBody" runat="server">
	<fieldset class="container" runat="server">
		<legend>Your Account</legend>
		<asp:Label ID="lblRole" style="display: none;" runat="server"/>
		
		<div id="sections">
			<fieldset>
				<legend><%=Say("Title.Sections")%></legend>
				<amp:enumcheckbox id="ampSectionList" name="section" runat="server" />
				<span class="note" style="padding:0;"><%=Say("Note.Sections")%></span>
			</fieldset>
		</div>
	
		<label class="required">
			<%=Say("Label.RealName")%></label>
			<amp:validation target="tbFirstName" type="Name" message="First Name (plain text only)" required="true" runat="server" />
			<asp:textbox id="tbFirstName" width="75" MaxLength="30" runat="server" />
			<amp:validation target="tbLastName" type="Name" message="Last Name (plain text only)" required="true" runat="server" />
			<asp:textbox id="tbLastName" width="75" MaxLength="30" runat="server" />
			<div class="note"><%=Say("Note.RealName")%></div>
			
		<amp:field id="fldScreenName" type="text" style="width: 100px;" resx="ScreenName"
			validate="PlainText" MaxLength="50" runat="server" />
		
		<fieldset id="credentials">
			<legend><%=Say("Title.Credentials")%></legend>
			
			<amp:Field id="fldEmail" Type="text" Resx="EmailAddress" Required="True"
				Validate="Email" MaxLength="50" Runat="server" />
				<div id="emailNote" class="note">if changed, a new confirmation code will be sent</div>
				
			<div id="confirm">
				<amp:Field ID="fldConfirmationCode" type="text" Resx="ConfirmationCode"
					ShowNote="False" Runat="server" />
				<div id="confirmNote" class="note"></div>
			</div>
				
			<div id="password">
				<amp:Field ID="fldPassword" type="password" Resx="Password"
					Validate="Password" MaxLength="50" ShowNote="False" Runat="server" />
				<input type="password" id="fldPasswordAgain" maxlength="50" />
				<div id="passwordNote" class="note">leave blank except to change; enter twice to verify</div>
			</div>
			
		</fieldset>
		
		<div id="errata">
			<amp:field id="fldPrivacy" type="checkbox" resx="HideEmail" runat="server" />
			
			<amp:Upload ID="fldPicture" Resx="Picture" Runat="server" ShowNote="False" /><br/>
			
			<amp:Field ID="fldWebSite" Type="text" Validate="URL" Resx="WebSite"
				MaxLength="100" Style="width: 20em;" Value="http://" ShowNote="False" Runat="server" />
			
			<fieldset id="description">
				<legend><%=Say("Label.AboutYou")%></legend>
				<amp:Validation Target="tbDescription" Type="PlainText" Message="Description (plain text only)" Required="False" runat="server" />
				<asp:textbox id="tbDescription" textmode="MultiLine" MaxLength="1000" runat="server" /><br/>
				<span class="note" style="padding:0;"><%=Say("Note.PlainText")%></span>
			</fieldset>
		</div>

		<input type="hidden" runat="server" id="fldOldEmail" />
		<input type="hidden" runat="server" id="fldConfirmation" />
	</fieldset>
	
	<div id="buttons">
		<amp:button id="btnSave" runat="server" resx="Save" submitform="true" />
		<amp:button id="btnCancel" runat="server" resx="Cancel" onclick="history.back()" />
	</div>
</asp:Panel>

<asp:Panel id="pnlWindows" runat="server">
	<amp:Contests id="Contests" runat="server" />
	<amp:NewAssets count="10" id="ampNewListings" runat="server" />
	<amp:Tours NewerThan="1/1/04" id="ampTours" runat="server" />
</asp:Panel>

</body>
</html>