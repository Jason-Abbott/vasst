<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_forums_cls.asp"-->
<!--#include file="./include/kb_forum-data_cls.asp"-->
<%
dim m_oLayout
dim m_oForums

Set m_oForums = New kbForums
With Request
	m_oForums.SortBy = .QueryString("sort")
	m_oForums.Category = .QueryString("cat")
	m_oForums.Page = .QueryString("page")
	m_oForums.Software = .QueryString("sw")
	m_oForums.Author = .QueryString("author")
End With
%>
<html>
<title><%=g_sORG_NAME%>: Tutorials</title>
<head>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<style>
TD.ItemOwner {
	font-size: 9pt;
	font-weight: bold;
	text-align: right;
	font-weight: normal;
	padding-right: 3px;
	padding-top: 3px;
}
TD.ItemName {
	font-size: 10pt;
}
DIV.ItemAction {
	text-align: center;
	font-weight: bold;
}	
</style>
</head>
<body>
<% Set m_oLayout = New kbLayout %>
<!--#include file="./include/kb_header_inc.asp"-->
<!--#include file="./include/kb_ads_inc.asp"-->
<!--#include file="./include/kb_message.inc"-->
<% Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) : Set m_oLayout = Nothing %>
<center>
<% Call m_oForums.WritePublic() : Set m_oForums = Nothing %>
</center>
<!--#include file="./sundance/sundance_footer.inc"-->
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>