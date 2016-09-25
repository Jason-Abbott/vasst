<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_contest_cls.asp"-->
<!--#include file="./include/kb_projects_cls.asp"-->
<!--#include file="./include/kb_project-data_cls.asp"-->
<%
' save comment
Const m_sFORM_NAME = "frmFileEdit"
dim m_oLayout
dim m_oProjects
dim m_oProjectData
dim m_lProjectID

With Request
	m_lProjectID = MakeNumber(.QueryString("id"))
	If IsNumber(.Form("fldDelete")) Then
		Set m_oProjectData = New kbProjectData
		Call m_oProjectData.Delete(Trim(.Form("fldDelete")))
		Set m_oProjectData = Nothing
		response.redirect "kb_projects.asp"
	ElseIf .Form("fldFriendlyName") <> "" Then
		Set m_oProjectData = New kbProjectData
		Call m_oProjectData.Update(m_lProjectID, .Form("fldFriendlyName"), .Form("fldDescription"), _
			.Form("fldFormat"), .Form("fldVersion"), .Form("fldRendered"), .Form("fldMedia"), _
			.Form("fldPlugins"), .Form("fldCategories"))
		Set m_oProjectData = Nothing
		response.redirect "kb_projects.asp"
	End If
End With
%>
<html>
<title><%=g_sORG_NAME%>: Project Edit</title>
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
<form name='<%=m_sFORM_NAME%>' action='kb_project-save.asp' method='post' onSubmit="return IsValidUpload();" enctype='multipart/form-data'>
<% Set m_oProjects = New kbProjects : Call m_oProjects.WriteEditForm(m_lProjectID) : Set m_oProjects = Nothing %>
<input type='hidden' name='fldAction' value='<%=g_ACT_FILE_UPDATE%>'>
<input type='hidden' name='fldProjectID' value='<%=m_lProjectID%>'>
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
	fldVersion:{desc:"Software needed to open it",type:"Select",req:1},
	fldDescription:{desc:"Description",type:"Posting",req:1},
	fldRendered:{desc:"Rendered URL",type:"URL",req:0},
	fldMedia:{desc:"Additional Media URL",type:"URL",req:0}};

SelectPlugins();
SelectCategories();
	
function IsValidUpload() {
	if (IsValid("<%=m_sFORM_NAME%>", m_oFields)) {
		if (m_oForm.fldFile.value.length > 0) {
			
			if (IsValidType("<%=g_sPROJECT_TYPES%>", m_oForm.fldFile.value, "project")) { return true; }
		} else {
			return true;
		}
	}
	return false;
}

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
			fldAction.value = '<%=g_ACT_FILE_DELETE%>';
			submit();
		}
	}
}
</script>