<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_verify-admin_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<%
Const m_sFORM_NAME = "frmUser"

dim m_aData
dim m_oLayout
dim m_oUser
%>
<html>
<head>
<title>Administration: Users</title>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_admin.css" rel="stylesheet" type="text/css">
</head>
<body>
<% Set m_oLayout = New kbLayout %>
<!--#include file="./include/kb_header_inc.asp"-->
<!--#include file="./include/kb_message.inc"-->
<% Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<% Call m_oLayout.WriteMenuBar(m_sMENU_ADMIN) %>
<center>
<form name='<%=m_sFORM_NAME%>' method='post' action='kb_admin-users.asp'>

<% Set m_oUser = New kbUser : Call m_oUser.WriteUserList("fldUsers", "", "", "") : Set m_oUser = Nothing%>
<br>
will have ability to ban, promote to admin, etc.
</form>

</center>
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<% Set m_oLayout = Nothing %>

<script language="javascript" src="./script/kb_validation.js"></script>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language='javascript'>
var m_oForm = document.<%=m_sFORM_NAME%>;
var m_oFields = {
	fldName:{desc:"Contest Name",type:"String",req:1},
	fldStart:{desc:"Contest Start Date",type:"Date",req:1},
	fldEnd:{desc:"Contest End Date",type:"Date",req:1},
	fldVotes:{desc:"Votes Allowed",type:"Numeric",req:1},
	fldWeight:{desc:"Vote Weighting",type:"Numeric",req:1},
	fldDescription:{desc:"Contest Description",type:"String",req:1}};
</script>