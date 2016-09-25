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
<!--#include file="./include/kb_contest_cls.asp"-->
<%
dim m_lSortID
dim m_lPage
dim m_oLayout
dim m_oFiles
dim m_oContest

m_lSortID = ReplaceNull(Request.QueryString("sort"), g_SORT_DATE_DESC)
m_lPage = ReplaceNull(Request.QueryString("page"), 1)
%>
<html>
<title><%=g_sORG_NAME%>: Free Files!</title>
<head>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./sundance/sundance_style.css" rel="stylesheet" type="text/css">
<style>
IMG.SortArrow {	margin-left: 3px; }
</style>
</head>
<body>
<!--#include file="./include/kb_header_inc.asp"-->
<!--#include file="./include/kb_message.inc"-->
here
</body>
</html>