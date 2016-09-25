<%@ Page Language="vb" AutoEventWireup="false" Codebehind="contests.aspx.vb" Inherits="AMP.Pages.Administration.Contests" %>
<%@ Register TagPrefix="amp" TagName="ContentWebPart" Src="~/control/Content.ascx"%>
<%@ Register TagPrefix="amp" Namespace="AMP.Controls" Assembly="AMP.VASST" %>
<html><head><meta name="vs_targetSchema" content="http://www.w3.org/1999/xhtml"></head>
<body xmlns:amp="urn:http://www.amp.com/schemas">

<asp:panel id="pnlBody" runat="server">
	<h1>Contests</h1>
	
	<asp:DropDownList ID="ddlContests" CssClass="select"
		DataValueField="ID" DataTextField="Title" Runat="server" />
	
	<fieldset class="container" runat="server">
		<legend runat="server" id="topLabel">New Contest</legend>
		
		<amp:field id="fldTitle" type="text" style="width: 280px;" resx="Title" shownote="false" 
			validate="PlainText" required="true" maxlength="75" runat="server" />
		
		<fieldset id="sections">
			<legend><%=Say("Title.Sections")%></legend>
			<amp:validation target="ampSectionList" type="Select" message="Section" required="true" runat="server" />
			<amp:enumlist id="ampSectionList" runat="server" rows="6" shownote="false" showlink="false" />
		</fieldset>
		
		<fieldset id="dates">
			<legend><%=Say("Title.Dates")%></legend>
			<amp:field id="fldStartDate" type="text" resx="StartDate" maxlength="10"
				validate="Date" required="true" inline="true" runat="server" /><br/>
			<amp:field id="fldEndVoteDate" type="text" resx="EndVoteDate" maxlength="10"
				validate="Date" required="true" shownote="false" runat="server" /><br/>
			<amp:field id="fldStopDate" type="text" resx="EndDate" maxlength="10"
				validate="Date" required="true" inline="true" runat="server" /><br/>
		</fieldset>
		
		<div id="voting">
			<amp:field id="fldWinners" type="text" resx="Winners" maxlength="2"
				validate="NonZero" required="true" inline="true" runat="server" /><br/>
			<amp:field id="fldVotes" type="text" resx="VotesAllowed" maxlength="2"
				validate="NonZero" required="true" inline="true" runat="server" /><br/>
			<amp:field id="fldVoteWeight" type="text" resx="VoteWeight" maxlength="2"
				validate="NonZero" required="true" inline="true" runat="server" /><br/>
		</div>
			
		<fieldset id="restrictions">
			<legend><%=Say("Title.EntryRestrictions")%></legend>
			
			<fieldset id="rules">
				<legend>Additional Rules</legend>
				<amp:validation target="tbRules" type="PlainText" message="Rules" required="false" runat="server" />
				<asp:textbox id="tbRules" textmode="MultiLine" MaxLength="600" runat="server" />
				<div class="note">each rule on a new line with hyphen</div>           
			</fieldset>
			
			<amp:field id="fldEntriesAllowed" type="text" resx="AllowedContestEntries" style="width: 2em;"
				validate="NonZero" required="true" shownote="false" maxlength="2" runat="server" /><br/>
			
			<label class="required"><%=Say("Label.FileTypes")%></label>
				<amp:validation target="ampFileList" type="Select" message="File Type (select at least one)" required="true" runat="server" />
				<amp:enumlist id="ampFileList" runat="server" shownote="false" showlink="false" /><br/>
			
			<amp:field id="fldMaxFileSize" type="text" resx="MaxContestFileSize" style="width: 40px;"
				validate="NonZero" required="true" inline="true" runat="server" maxlength="6" /><br/>

			<div id="checkboxes">
				<amp:field id="fldAnonymous" type="checkbox" resx="AnonymousEntries" runat="server" />
				<span id="forProjects">
					<amp:field id="fldFreePlugins" type="checkbox" resx="FreePlugins" runat="server" />
					<amp:field id="fldGeneratedMediaOnly" type="checkbox" resx="GeneratedMediaOnly" runat="server" />
				</span>
			</div>
		</fieldset>
		
		<fieldset id="prizes">
			<legend>Prizes</legend>
			<amp:validation target="tbPrizes" type="PlainText" message="Prizes" required="false" runat="server" ID="Validation1"/>
			<asp:textbox id="tbPrizes" textmode="MultiLine" MaxLength="600" runat="server" />
			<div class="note">each prize in order on a new line with hyphen</div>           
		</fieldset>
		
		<fieldset id="description">
			<legend><%=Say("Label.Description")%></legend>
			<amp:validation target="tbDescription" type="PlainText" message="Description" required="true" runat="server" />
			<asp:textbox id="tbDescription" textmode="MultiLine" MaxLength="600" runat="server" />
		</fieldset>
		
		
	</fieldset>
	
	<div id="buttons">
		<amp:button runat="server" resx="Save" submitform="true" />
		<amp:button id="btnCancel" runat="server" resx="Cancel" onclick="history.back()" />
	</div>
</asp:panel>

<asp:panel id="pnlWindows" runat="server" >
	<amp:contentwebpart file="help/ContestWeighting.txt" title="About Weighting" id="ampContent" runat="server" />
</asp:panel>
</body>
</html>