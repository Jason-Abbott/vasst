<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_verify-admin_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_files_cls.asp"-->
<!--#include file="./include/kb_file-data_cls.asp"-->
<!--#include file="./include/kb_tutorials_cls.asp"-->
<!--#include file="./include/kb_tutorial-data_cls.asp"-->
<!--#include file="./include/kb_mail_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<%
dim m_lItemID
dim m_oLayout
dim m_oFile
dim m_oTutorial

m_lItemID = Request.QueryString("id")
If IsNumber(m_lItemID) Then
	Set m_oFile = New kbFileData
	Set m_oTutorial = New kbTutorialData
	select case Request.QueryString("do")
		case "approvefile"
			Call m_oFile.ApprovePending(m_lItemID)
		case "denyfile"
			Call m_oFile.DenyPending(m_lItemID)
		case "approvetut"
			Call m_oTutorial.ApprovePending(m_lItemID)
		case "denytut"
			Call m_oTutorial.DenyPending(m_lItemID)
	end select
	Set m_oFile = Nothing
	Set m_oTutorial = Nothing
End If
%>
<html>
<head>
<title>Administration: File Uploads</title>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_admin.css" rel="stylesheet" type="text/css">
<style>
TD.UploadsHead {
	text-align: center;
	font-size: 9pt;
	font-weight: bold;
	padding: 5px;
}
TD.UploadName {
	font-size: 9pt;
	padding-right: 8px;
	padding-top: 2px;
	border-top: 1px solid <%=g_sCOLOR_EDGE%>;
}
TD.UploadDate {
	font-size: 9pt;
	color: <%=g_sCOLOR_EDGE%>;
	border-top: 1px solid <%=g_sCOLOR_EDGE%>;
	padding-right: 8px;
}
TD.UploadOwner {
	font-size: 9pt;
	padding-right: 8px;
	text-align: center;
	border-top: 1px solid <%=g_sCOLOR_EDGE%>;
}
TD.UploadDescription {
	font-size: 9pt;
	padding-left: 25px;
	padding-bottom: 15px;
}
TD.UploadFriendlyName {
	font-size: 9pt;
	text-align: center;
	border-top: 1px solid <%=g_sCOLOR_EDGE%>;
	padding-right: 8px;
	
}
TD.UploadAction {
	font-size: 9pt;
	text-align: right;
	font-weight: bold;
	padding-left: 10px;
	border-top: 1px solid <%=g_sCOLOR_EDGE%>;
}
TD.UploadNotes {
	font-size: 8pt;
	color: #999999;
	padding-left: 10px;
	padding-bottom: 10px;
}
DIV.Head {
	font-family: Impact, Arial, Helvetica;
	font-size: 18pt;
	text-align: center;
	color: <%=g_sCOLOR_EDGE%>;
}
</style>
</head>
<body>
<!--#include file="./sundance/sundance_header.inc"-->
<!--#include file="./include/kb_message.inc"-->
<% Set m_oLayout = New kbLayout : Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<% Call m_oLayout.WriteMenuBar(m_sMENU_ADMIN) : Set m_oLayout = Nothing %>
<center>
<div class='Head'>Files</div>
<% Set m_oFile = New kbFiles : Call m_oFile.WritePending() : Set m_oFile = Nothing%>
<div class='Head'>Tutorials</div>
<% Set m_oTutorial = New kbTutorials : Call m_oTutorial.WritePending() : Set m_oTutorial = Nothing%>
</center>
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>