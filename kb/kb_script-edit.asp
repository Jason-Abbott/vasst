<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_contest_cls.asp"-->
<!--#include file="./include/kb_scripts_cls.asp"-->
<!--#include file="./include/kb_script-data_cls.asp"-->
<%
' save comment
Const m_sFORM_NAME = "frmFileEdit"
dim m_oLayout
dim m_oScripts
dim m_oScriptData
dim m_lScriptID

With Request
	m_lScriptID = MakeNumber(.QueryString("id"))
	If IsNumber(.Form("fldDelete")) Then
		Set m_oScriptData = New kbScriptData
		Call m_oScriptData.Delete(Trim(.Form("fldDelete")))
		Set m_oScriptData = Nothing
		response.redirect "kb_projects.asp"
	ElseIf .Form("fldFriendlyName") <> "" Then
		Set m_oScriptData = New kbScriptData
		Call m_oScriptData.Update(m_lScriptID, .Form("fldFriendlyName"), .Form("fldDescription"), _
			.Form("fldFormat"), .Form("fldVersion"), .Form("fldMedia"), _
			.Form("fldPlugins"), .Form("fldCategories"))
		Set m_oScriptData = Nothing
		response.redirect "kb_scripts.asp"
	End If
End With
%>
<html>
<title><%=g_sORG_NAME%>: Script Edit</title>
<head>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
</head>
<body>
<% Set m_oLayout = New kbLayout %>
<!--#include file="./include/kb_header_inc.asp"-->
<!--#include file="./include/kb_ads_inc.asp"-->
<!--#include file="./include/kb_message.inc"-->
<% Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) : Set m_oLayout = Nothing %>
<center>
<form name='<%=m_sFORM_NAME%>' action='kb_script-save.asp' method='post' onSubmit="return IsValidUpload();" enctype='multipart/form-data'>
<% Set m_oScripts = New kbScripts : Call m_oScripts.WriteEditForm(m_lScriptID) : Set m_oScripts = Nothing %>
<input type='hidden' name='fldAction' value='<%=g_ACT_FILE_UPDATE%>'>
<input type='hidden' name='fldScriptID' value='<%=m_lScriptID%>'>
<input type='hidden' name='fldURL' value='<%=Request.QueryString("url")%>'>
</form>
</center>
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language="javascript" src="./script/kb_validation.js"></script>
<script language='javascript'>
var m_oForm = document.<%=m_sFORM_NAME%>;
var m_oFields = {
	fldFriendlyName:{desc:"Friendly project name",type:"String",req:1},
	fldVersion:{desc:"Software needed to run it",type:"Select",req:1},
	fldDescription:{desc:"Description",type:"Posting",req:1},
	fldMedia:{desc:"Additional Resources URL",type:"URL",req:0}};

SelectCategories();

function IsValidUpload() {
	if (IsValid("<%=m_sFORM_NAME%>", m_oFields)) {
		if (m_oForm.fldFile.value.length > 0) {
			if (IsValidType("<%=g_sSCRIPT_TYPES%>", m_oForm.fldFile.value, "script")) { return true; }
		} else {
			return true;
		}
	}
	return false;
}

function DeleteFile() {
	if (confirm("You are about to permanently remove this file from our listing. \nAre you sure you want to continue?")) {
		with (m_oForm) {
			fldAction.value = '<%=g_ACT_FILE_DELETE%>';
			submit();
		}
	}
}
</script>