<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_verify-admin_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_files_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_cache_cls.asp"-->
<%
Const m_sFORM_NAME = "frmCache"
dim m_sMessage
dim m_oCache
dim m_oData
dim m_oLayout

Set m_oLayout = New kbLayout
Set m_oCache = New kbCache

select case Trim(Request.QueryString("do"))
	case "reset"
		Call ClearCache(Request.QueryString("type"))
		Call SetSessionValue(g_USER_MSG, "The cache has been reset")
	case "header"
		Call m_oCache.ResetHeader()
		Call SetSessionValue(g_USER_MSG, "The header has been reset")
end select
%>
<html>
<head>
<title>Administration: Cache</title>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_admin.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="./sundance/sundance_header.inc"-->
<!--#include file="./include/kb_message.inc"-->
<% Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<% Call m_oLayout.WriteMenuBar(m_sMENU_ADMIN) %>
<center>
<p>
Reset cache for<br>
<% Call m_oCache.WriteRefreshList("kb_admin-cache.asp") %>

</center>
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<% Set m_oCache = Nothing : Set m_oLayout = Nothing %>

<script language="javascript" src="./script/kb_functions.js"></script>
<script language='javascript'>
var m_oForm = document.<%=m_sFORM_NAME%>;

function Restore() {
	var oField = m_oForm.fldBackup;
	var sDatabase = oField.options[oField.selectedIndex].text;
	alert("You selected " + sDatabase + "\nNot coded yet");
}

</script>