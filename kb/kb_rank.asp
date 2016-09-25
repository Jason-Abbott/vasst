<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_projects_cls.asp"-->
<!--#include file="./include/kb_project-data_cls.asp"-->
<!--#include file="./include/kb_scripts_cls.asp"-->
<!--#include file="./include/kb_script-data_cls.asp"-->
<!--#include file="./include/kb_tutorials_cls.asp"-->
<!--#include file="./include/kb_tutorial-data_cls.asp"-->
<!--#include file="./include/kb_reviews_cls.asp"-->
<!--#include file="./include/kb_review-data_cls.asp"-->
<!--#include file="./include/kb_forums_cls.asp"-->
<!--#include file="./include/kb_forum-data_cls.asp"-->
<!--#include file="./include/kb_rank_cls.asp"-->
<%
Const m_sFORM_NAME = "frmRank"
dim m_oLayout
dim m_oItem
dim m_oRank
dim m_lItemID
dim m_lItemTypeID
dim m_sItemObject

with Request
	m_lItemID = MakeNumber(.QueryString("id"))
	m_lItemTypeID = MakeNumber(.QueryString("type"))
	If .Form("fldStars") <> "" Then
		Set m_oRank = New kbRank
		Call m_oRank.Save(m_lItemID, m_lItemTypeID, .Form("fldStars"), .Form("fldComment"))
		Set m_oRank = Nothing
	ElseIf .QueryString("do") = "delete" Then
		Set m_oRank = New kbRank
		Call m_oRank.Delete(m_lItemID, m_lItemTypeID, .QueryString("user"))
		Set m_oRank = Nothing
	End If
end with

Select Case m_lItemTypeID
	Case g_ITEM_PROJECT : m_sItemObject = "kbProjects"
	Case g_ITEM_SCRIPT : m_sItemObject = "kbScripts"
	Case g_ITEM_TUTORIAL : m_sItemObject = "kbTutorials"
	Case g_ITEM_FORUM : m_sItemObject = "kbForums"
	Case g_ITEM_REVIEW : m_sItemObject = "kbReviews"
	Case Else : response.redirect g_sDEFAULT_PAGE
End Select
%>
<html>
<title><%=g_sORG_NAME%>: Rankings</title>
<head>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_rank.css" rel="stylesheet" type="text/css">
</head>
<body>
<% Set m_oLayout = New kbLayout %>
<!--#include file="./include/kb_header_inc.asp"-->
<!--#include file="./include/kb_ads_inc.asp"-->
<!--#include file="./include/kb_message.inc"-->
<% Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) : Set m_oLayout = Nothing %>
<form name='<%=m_sFORM_NAME%>' action='kb_rank.asp?id=<%=m_lItemID%>&type=<%=m_lItemTypeID%>' method='post' onSubmit="return IsValid('<%=m_sFORM_NAME%>', m_oFields);">
<center>
<% Set m_oItem = Eval("New " & m_sItemObject) : Call m_oItem.WriteAsHeader(m_lItemID) : Set m_oItem = Nothing %>
<br>
<% Set m_oRank = new kbRank : Call m_oRank.WriteItemRankings(m_lItemID, m_lItemTypeID) : Set m_oRank = Nothing %>
</center>
</form>
<!--#include file="./sundance/sundance_footer.inc"-->
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language="javascript" src="./script/kb_validation.js"></script>
<script language="javascript">
var m_oForm = document.<%=m_sFORM_NAME%>;
var m_oFields = { fldComment:{desc:"rank comment",type:"Posting",req:1} }
</script>
