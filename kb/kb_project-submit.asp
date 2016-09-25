<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_projects_cls.asp"-->
<%
Const m_sFORM_NAME = "frmSubmission"
dim m_oLayout
dim m_oProject
%>
<html>
<title><%=g_sORG_NAME%>: Submit Your Project</title>
<head>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<style>
TD.Terms {
	font-size: 8pt;
	width: 150px;
	padding-left: 20px;
	text-align: justify;
}
DIV.TermsTitle {
	font-size: 9pt;
	text-align: center;
	font-weight: bold;
	border-bottom: 1px solid <%=g_sCOLOR_EDGE%>;
	margin-bottom: 4px;
}
</style>
</head>
<body>
<% Set m_oLayout = New kbLayout %>
<!--#include file="./include/kb_header_inc.asp"-->
<!--#include file="./include/kb_ads_inc.asp"-->
<!--#include file="./include/kb_message.inc"-->
<% Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<center>
<form name='<%=m_sFORM_NAME%>' action='kb_project-save.asp' method='post' enctype='multipart/form-data' onSubmit='return IsValidUpload();'>
<table border='0'>
<tr>
	<td valign='top'>
	<% Call m_oLayout.WriteTitleBoxTop("Project File Submission", "", "") %>
	<table cellspacing='0' cellpadding='0' border='0'>
	<tr>
		<td class='Required'>File:</td>
		<td class='FormInput' valign='top'><input type='file' name='fldFile' size='45' onChange='getName(this);'></td>
	<tr><td></td><td class='FormNote'>maximum file size is <%=g_MAX_FILE_KB%> kilobytes</td>
	<tr>
		<td class='Required'>Friendly Name:</td>
		<td class='FormInput' valign='top'>
        <input type='text' name='fldFriendlyName' maxlength='30' size='30'></td>
	<tr>
		<td class='Required'>Requires:</td>
		<td class='FormInput' valign='top'>
			<% Call m_oLayout.WriteVersionList("fldVersion", "", g_ITEM_PROJECT) %> software required to open the project</td>
	<tr>
		<td class='FormLabel'>Format:</td>
		<td class='FormInput' valign='top'>
			<% Set m_oProject = New kbProjects : Call m_oProject.WriteFormatList("fldFormat", GetSessionValue(g_USER_FILE_FORMAT)) %> &nbsp;</td>
	<tr>
		<td class='FormLabel'>Rendered URL:</td>
		<td class='FormInput' valign='top'><input type='text' name='fldRendered' maxlength='75' size='30' value='http://'> link to rendered version</td>
	<tr>
		<td class='FormLabel'>Media URL:</td>
		<td class='FormInput' valign='top'><input type='text' name='fldMedia' maxlength='75' size='30' value='http://'> link to media used</td>
	<tr>
		<td></td>
		<td class='FormNote'><%=g_sMSG_MEDIA_HINT%></td>
	<tr>
		<td class='Required' valign='top'>Description:</td>
		<td class='FormInput' valign='top'><textarea rows='6' cols='52' name='fldDescription'></textarea></td>
	<tr><td></td><td class='FormNote'><%=g_sMSG_HTML_LIMIT%>&nbsp;</td>
	<tr>
		<td></td>
		<td class='FormInput' valign='top'>
		<table cellspacing='0' cellpadding='4' border='0'>
		<tr>
			<td valign='top' width='50%'>
			<b>Categories:</b><br>
			<% Call m_oLayout.WriteCategoryList("fldCategories", "", 4, g_ITEM_PROJECT) %>
			</td>
			<td valign='top' width='50%'>
			<b>Plugins:</b><br>
			<% Call m_oProject.WritePluginList("fldPlugins", "", 4) : Set m_oProject = Nothing %>
			</td>
		<tr>
			<td colspan='2' class='FormNote' align='center'><%=g_sMSG_MULTI_SELECT%></td>
		</table>
		</td>
	<tr>
		<td class='Required' style='font-size: 8pt; text-align: center;'>(required)</td>
		<td align='right'>
			<% Call m_oLayout.WriteToggleImage("btn_upload-file", "", "Upload File", "class='Image' width='80' height='14'", true) %>
		</td>
	</table>
	<% Call m_oLayout.WriteBoxBottom("") : Set m_oLayout = nothing %>
	</td>
	<td class='Terms' valign='top'>
		<div class='TermsTitle'>Terms & Conditions</div>
		<!--#include file="./include/upload_terms_sundance.inc"-->
	</td>
<% if MakeNumber(g_lSiteID) = g_SITE_VEGAS then %>
<tr>
	<td align='left'>
	<br>
	<img align='left' src='images/vegas-properties.jpg' width='200' height='154'>
	<div style='width: 220px; text-align: left; color: #aaaaaa;'>
	<b>Uploading a Vegas file?</b>
	Let people know who made it under File|Properties|Summary.  Put in your name, project title and anything in the comments field that you'd like others to know about the project or yourself.
	</div>
	</td>
<% end if %>
</table>
<input type='hidden' name='fldAction' value='<%=g_ACT_FILE_ADD%>'>
</form>
</center>
<!--#include file="./sundance/sundance_footer.inc"-->
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language="javascript" src="./script/kb_validation.js"></script>
<script language="javascript">
var m_oForm = document.<%=m_sFORM_NAME%>;
var m_oFields = {
	fldFile:{desc:"Project file (Browse...)",type:"String",req:1},
	fldFriendlyName:{desc:"Friendly project name",type:"String",req:1},
	fldVersion:{desc:"Software needed to open it",type:"Select",req:1},
	fldDescription:{desc:"Description",type:"Posting",req:1},
	fldRendered:{desc:"Rendered URL",type:"URL",req:0},
	fldMedia:{desc:"Additional Media URL",type:"URL",req:0}};
	
function IsValidUpload() {
	if (IsValid('<%=m_sFORM_NAME%>', m_oFields)) {
		if (IsValidType("<%=g_sPROJECT_TYPES%>", m_oForm.fldFile.value, "project")) { return true; }
	}
	return false;
}

function getName(r_oField) {
	var sName = r_oField.value;
	var oTargetFld = m_oForm.fldFriendlyName;
	if (sName.length != 0 && oTargetFld.value.length == 0) {
		// only overwrite friendly name if none exists
		sName = sName.substr(sName.lastIndexOf("\\") + 1, sName.length);
		oTargetFld.value = sName.substr(0, sName.lastIndexOf("."));
	}
}
</script>