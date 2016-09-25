<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_verify-admin_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_cache_cls.asp"-->
<%
dim m_oCache
dim m_oLayout

Set m_oLayout = New kbLayout
Set m_oCache = New kbCache

Call m_oCache.RefreshSelected(Request.Form("fldRefresh"))
%>
<html>
<head>
<title>Administration: Cache</title>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_admin.css" rel="stylesheet" type="text/css">
<style>
INPUT { border: none; }
</style>
</head>
<body>
<!--#include file="./include/kb_header_inc.asp"-->
<!--#include file="./include/kb_message.inc"-->
<% Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<% Call m_oLayout.WriteMenuBar(m_sMENU_ADMIN) %>
<p><center><% Call m_oCache.WriteRefreshList("kb_admin-cache.asp") %></center>
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<% Set m_oCache = Nothing : Set m_oLayout = Nothing %>

<script language="javascript" src="./script/kb_functions.js"></script>