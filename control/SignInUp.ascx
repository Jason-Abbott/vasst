<%@ Control Language="vb" AutoEventWireup="false" Codebehind="SignInUp.ascx.vb" Inherits="AMP.Controls.SignInUp" TargetSchema="http://www.w3.org/1999/xhtml" %>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>

<fieldset xmlns:amp="urn:http://www.amp.com/schemas">
	<legend id="lgdTitle" runat="server"></legend>
	
	<amp:Field ID="fldEmail" Type="text" Resx="EmailAddress" Required="True"
		Validate="Email" MaxLength="50" Runat="server" />
		<div id="emailNote" class="note"></div>
	
	<div id="password">
		<amp:Field ID="fldPassword" type="password" Resx="Password" Required="True"
			Validate="Password" MaxLength="50" ShowNote="False" Runat="server" />
		<input type="password" id="fldPasswordAgain" maxlength="50" />
		<div id="passwordNote" class="note"></div>
	</div>
	
	<div id="signup">	
		<amp:Field ID="fldConfirmationCode" type="text" Resx="ConfirmationCode"
			ShowNote="False" Runat="server" />
			<div id="confirmNote" class="note"><%=Say("Note.ConfirmationCode")%></div>
		
		<fieldset id="name">
			<legend><%=Say("Label.Name")%></legend>
			<label class="required"><%=Say("Label.RealName")%></label>
				<amp:validation target="tbFirstName" type="Name" message="First Name (plain text only)" required="true" runat="server" />
				<asp:textbox id="tbFirstName" MaxLength="30" runat="server" />
				<amp:validation target="tbLastName" type="Name" message="Last Name (plain text only)" required="true" runat="server" />
				<asp:textbox id="tbLastName" MaxLength="30" runat="server" />
				<div class="note"><%=Say("Note.RealName")%></div>
			<amp:field id="fldScreenName" type="text" resx="ScreenName"
				validate="PlainText" MaxLength="50" runat="server" />
		</fieldset>
		
		<amp:field id="fldPrivacy" type="checkbox" resx="HideEmail" runat="server" />
		<amp:Field ID="fldWebSite" Type="text" Validate="URL" Resx="WebSite"
			MaxLength="100" Style="width: 20em;" Value="http://" ShowNote="False" Runat="server" />
	</div>
	
	<input type="hidden" runat="server" id="fldConfirmation"/>
	<input type="hidden" runat="server" id="fldSignup" value="false" />
	<input type="hidden" runat="server" id="fldClientTime" />
	<input type="hidden" runat="server" id="fldServerTime" />
	<input type="hidden" runat="server" id="fldTipType" value="welcome" />
	<input type="hidden" name="step" value="4" />
</fieldset>

<div id="links">
	<a href="javascript:Login.ShowSignup(true)"><%=Say("Action.Register")%></a><br/>
	<a href="javascript:Login.Broker.SendPassword()"><%=Say("Action.NewPassword")%></a>
</div>

<asp:Label ID="lblWarning" CssClass="warning" Runat="server">
We are pleased to support the Opera browser.  We want to let you know, however, that some insurmountable quirks were experienced during testing, the same that <a href="http://www.scss.com.au/family/andrew/opera/gmail/">prevent Opera from working with Gmail</a>.
<br/><br/>
Some of the functionality on this page may be delayed for several seconds while <a href="http://groups-beta.google.com/group/opera.general/browse_frm/thread/2192fd26964e5f5b">Opera reports "request queued."</a>
<br/><br/>
These issues are resolved in the upcoming 7.6 and 8.0 versions.
</asp:Label>