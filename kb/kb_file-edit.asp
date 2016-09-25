<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_contest_cls.asp"-->
<!--#include file="./include/kb_files_cls.asp"-->
<!--#include file="./include/kb_file-data_cls.asp"-->
<%
' save comment
Const m_sFORM_NAME = "frmFileEdit"
dim m_oLayout
dim m_oFiles
dim m_oFileData
dim m_lFileID

With Request
	m_lFileID = MakeNumber(.QueryString("id"))
	If IsNumber(.Form("fldDelete")) Then
		Set m_oFileData = New kbFileData
		Call m_oFileData.Delete(Trim(.Form("fldDelete")))
		Set m_oFileData = Nothing
		response.redirect "kb_files.asp"
	ElseIf .Form("fldFriendlyName") <> "" Then
		Set m_oFileData = New kbFileData
		Call m_oFileData.Update(m_lFileID, .Form("fldFriendlyName"), .Form("fldDescription"), _
			.Form("fldFormat"), .Form("fldVersion"), .Form("fldRendered"), .Form("fldMedia"), _
			.Form("fldPlugins"), .Form("fldCategories"))
		Set m_oFileData = Nothing
		response.redirect "kb_files.asp"
	End If
End With
%>
<html>
<title><%=g_sORG_NAME%>: File Edit</title>
<head>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="./sundance/sundance_header.inc"-->
<!--#include file="./include/kb_message.inc"-->
<% Set m_oLayout = New kbLayout : Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) : Set m_oLayout = Nothing %>
<center>
<form name='<%=m_sFORM_NAME%>' action='kb_file-save.asp' method='post' onSubmit="return isValid('<%=m_sFORM_NAME%>',m_oFields);" enctype='multipart/form-data'>
<% Set m_oFiles = New kbFiles : Call m_oFiles.WriteEditForm(m_lFileID) : Set m_oFiles = Nothing %>
<input type='hidden' name='fldAction' value='<%=m_ACT_FILE_UPDATE%>'>
<input type='hidden' name='fldFileID' value='<%=m_lFileID%>'>
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
	fldFriendlyName:{desc:"Listed Name",type:"String",req:1},
	fldDescription:{desc:"Description",type:"Posting",req:1}};

SelectPlugins();
SelectCategories();
	
function SelectPlugins() {
	var sPlugins = m_oForm.fldPluginList.value;
	if (sPlugins != "") {
		var aPlugin = sPlugins.split(",");
		var oPlugin = m_oForm.fldPlugins;
		for (var x = 0; x < aPlugin.length; x++) {
			for (var y = 0; y < oPlugin.options.length; y++) {
				if (oPlugin.options[y].value == aPlugin[x]) {
					oPlugin.options[y].selected = true;
				}
			}
		}
	}
}
	
function DeleteFile() {
	if (confirm("You are about to permanently remove this file from our listing. \nAre you sure you want to continue?")) {
		with (m_oForm) {
			fldAction.value = '<%=m_ACT_FILE_DELETE%>';
			submit();
		}
	}
}
</script>