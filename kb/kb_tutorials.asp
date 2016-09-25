<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_tutorials_cls.asp"-->
<!--#include file="./include/kb_tutorial-data_cls.asp"-->
<%
dim m_oLayout
dim m_oTutorials

Set m_oTutorials = New kbTutorials
With Request
	m_oTutorials.SortBy = .QueryString("sort")
	m_oTutorials.Category = .QueryString("cat")
	m_oTutorials.Page = .QueryString("page")
	m_oTutorials.Software = .QueryString("sw")
	m_oTutorials.Author = .QueryString("author")
End With
%>
<html>
<title><%=g_sORG_NAME%>: Tutorials</title>
<head>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="./sundance/sundance_header.inc"-->
<!--#include file="./include/kb_message.inc"-->
<% Set m_oLayout = New kbLayout : Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) : Set m_oLayout = Nothing %>
<center>
<% Call m_oTutorials.WritePublic() : Set m_oTutorials = Nothing %>
</center>
<!--#include file="./sundance/sundance_footer.inc"-->
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>