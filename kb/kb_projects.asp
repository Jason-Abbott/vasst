<% Option Explicit %>
<% Response.Buffer = False %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_projects_cls.asp"-->
<!--#include file="./include/kb_project-data_cls.asp"-->
<!--#include file="./include/kb_file-system_cls.asp"-->
<!--#include file="./include/kb_contest_cls.asp"-->
<%
dim m_oLayout
dim m_oFiles
dim m_oContest
dim m_lPassedSiteID

m_lPassedSiteID = Request.QueryString("s")
If IsNumber(m_lPassedSiteID) And m_lPassedSiteID <> g_lSiteID Then
	g_lSiteID = m_lPassedSiteID
	Call SetSessionValue(g_USER_SITE, m_lPassedSiteID)
End If

Set m_oFiles = New kbProjects
With Request
	m_oFiles.SortBy = .QueryString("sort")
	m_oFiles.Category = .QueryString("cat")
	m_oFiles.Page = .QueryString("page")
	m_oFiles.Software = .QueryString("sw")
	m_oFiles.Author = .QueryString("author")
End With
%>
<html>
<title><%=g_sORG_NAME%>: Free Files!</title>
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
<% Set m_oContest = New kbContest : Call m_oContest.WriteContests() : Set m_oContest = Nothing %>
<center>
<% Call m_oFiles.WritePublic() : Set m_oFiles = Nothing %>
</center>
<!--#include file="./sundance/sundance_footer.inc"-->
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>