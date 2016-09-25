<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_scripts_cls.asp"-->
<%
Const m_sFORM_NAME = "frmSubmission"
dim m_oLayout
%>
<html>
<title><%=g_sORG_NAME%>: Submit Your Script</title>
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
<form name='<%=m_sFORM_NAME%>' action='kb_script-save.asp' method='post' enctype='multipart/form-data' onSubmit='return IsValidUpload();'>
<table border='0'>
<tr>
	<td valign='top'>
	<% Call m_oLayout.WriteTitleBoxTop("Script File Submission", "", "") %>
	<table cellspacing='0' cellpadding='0' border='0'>
	<tr>
		<td class='Required'>Script:</td>
		<td class='FormInput' valign='top'><input type='file' name='fldFile' size='45' onChange='getName(this);'></td>
	<tr><td></td><td class='FormNote'>maximum file size is <%=g_MAX_FILE_KB%> kilobytes</td>
	<tr>
		<td class='Required'>Friendly Name:</td>
		<td class='FormInput' valign='top'>
        <input type='text' name='fldFriendlyName' maxlength='30' size='30'></td>
	<tr>
		<td class='Required'>Requires:</td>
		<td class='FormInput' valign='top'>
			<% Call m_oLayout.WriteVersionList("fldVersion", "", g_ITEM_SCRIPT) %> software required to run the script</td>
	<tr>
		<td class='FormLabel'>Resources URL:</td>
		<td class='FormInput' valign='top'><input type='text' name='fldResources' maxlength='75' size='30' value='http://'> link to any extra files needed</td>
	<tr>
		<td></td>
		<td class='FormNote'><%=g_sMSG_MEDIA_HINT%></td>
	<tr>
		<td class='Required' valign='top'>Description:</td>
		<td class='FormInput' valign='top'><textarea rows='6' cols='52' name='fldDescription'></textarea></td>
	<tr><td></td><td class='FormNote'><%=g_sMSG_HTML_LIMIT%>&nbsp;</td>
	
	<tr>
		<td class='FormLabel' valign="top">Categories:</td>
		<td class='FormInput' valign='top'>
		<% Call m_oLayout.WriteCategoryList("fldCategories", "", 4, g_ITEM_SCRIPT) %>
	<tr>
		<td></td>
		<td class="FormNote"><%=g_sMSG_MULTI_SELECT%></td>
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
	fldFile:{desc:"Script file (Browse...)",type:"String",req:1},
	fldFriendlyName:{desc:"Friendly script name",type:"String",req:1},
	fldVersion:{desc:"Software needed to run it",type:"Select",req:1},
	fldDescription:{desc:"Description",type:"Posting",req:1},
	fldResources:{desc:"Additional files needed",type:"URL",req:0}};

function IsValidUpload() {
	if (IsValid('<%=m_sFORM_NAME%>', m_oFields)) {
		if (IsValidType("<%=g_sSCRIPT_TYPES%>", m_oForm.fldFile.value, "script")) { return true; }
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