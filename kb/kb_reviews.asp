<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_reviews_cls.asp"-->
<!--#include file="./include/kb_review-data_cls.asp"-->
<%
dim m_oLayout
dim m_oReviews

Set m_oReviews = New kbReviews
With Request
	m_oReviews.SortBy = .QueryString("sort")
	m_oReviews.Category = .QueryString("cat")
	m_oReviews.Page = .QueryString("page")
	m_oReviews.Software = .QueryString("sw")
	m_oReviews.Author = .QueryString("author")
End With
%>
<html>
<title><%=g_sORG_NAME%>: Tutorials</title>
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
<% Call m_oReviews.WritePublic() : Set m_oReviews = Nothing %>
</center>
<!--#include file="./sundance/sundance_footer.inc"-->
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>