<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_verify-admin_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_projects_cls.asp"-->
<!--#include file="./include/kb_project-data_cls.asp"-->
<!--#include file="./include/kb_scripts_cls.asp"-->
<!--#include file="./include/kb_script-data_cls.asp"-->
<!--#include file="./include/kb_tutorials_cls.asp"-->
<!--#include file="./include/kb_tutorial-data_cls.asp"-->
<!--#include file="./include/kb_reviews_cls.asp"-->
<!--#include file="./include/kb_review-data_cls.asp"-->
<!--#include file="./include/kb_file-system_cls.asp"-->
<!--#include file="./include/kb_mail_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<%
dim m_oLayout

Call ProcessItem()

'-------------------------------------------------------------------------
'	Name: 		ProcessItem()
'	Purpose: 	approve or deny given item
'Modifications:
'	Date:		Name:	Description:
'	7/21/04		JEA		Created
'-------------------------------------------------------------------------
Sub ProcessItem()
	dim lItemID
	dim sAction
	dim oItem

	lItemID = Request.QueryString("id")
	sAction = Request.QueryString("do")
	
	If IsNumber(lItemID) And Not IsVoid(sAction) Then
		Select Case Right(sAction, 6)
			Case "roject" : Set oItem = New kbProjectData
			Case "script" : Set oItem = New kbScriptData
			Case "torial" : Set oItem = New kbTutorialData
			Case "review" : Set oItem = New kbReviewData
		End Select
		
		If Left(sAction, 4) = "deny" then
			oItem.DenyPending(lItemID)
		Else
			oItem.ApprovePending(lItemID)
		End If
		
		Set oItem = Nothing
	End If
End Sub

'-------------------------------------------------------------------------
'	Name: 		WritePendingItems()
'	Purpose: 	output all pendings items as HTML
'Modifications:
'	Date:		Name:	Description:
'	7/21/04		JEA		Created
'-------------------------------------------------------------------------
Sub WritePendingItems()
	dim oItem
	dim sItem
	dim aItemNames
	
	aItemNames = Array("Projects","Scripts","Tutorials","Reviews")
	
	with response
		.Write "<center>"
		for each sItem in aItemNames
			.Write "<div class='Head'>"
			.Write sItem
			.Write "</div>"
			Set oItem = Eval("New kb" & sItem)
			Call oItem.WritePending()
			Set oItem = Nothing
		next
		.Write "</center>"
	end with
End Sub
%>
<html>
<head>
<title>Administration: Submissions</title>
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
	margin-top: 10px;
	color: <%=g_sCOLOR_EDGE%>;
}
</style>
</head>
<body>
<% Set m_oLayout = New kbLayout %>
<!--#include file="./include/kb_header_inc.asp"-->
<!--#include file="./include/kb_message.inc"-->
<% Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<% Call m_oLayout.WriteMenuBar(m_sMENU_ADMIN) : Set m_oLayout = Nothing %>
<% Call WritePendingItems() %>
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>