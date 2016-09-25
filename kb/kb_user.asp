<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_mail_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<%
dim m_oLayout
dim m_oUser
dim m_oMail
dim m_lUserID

With Request
	m_lUserID = ReplaceNull(.QueryString("id"), GetSessionValue(g_USER_ID))
	if Trim(.Form("fldSubject")) <> "" then
		Set m_oMail = New kbMail
		Call m_oMail.SendUserToUserEmail(.Form("fldFrom"), .Form("fldTo"), .Form("fldSubject"), _
			.Form("fldBody"), CBool(.Form("fldBcc") = "on"))
		Set m_oMail = Nothing
	end if
End With
%>
<html>
<title><%=g_sORG_NAME%>: User</title>
<head>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<style>
DIV.Name {
	font-family: Impact, Arial, Helvetica;
	font-size: 18pt;
	margin-right: 7px;
}
DIV.RealName {
	font-size: 9pt;
	font-weight: bold;
	margin-left: 7px;
	color: #aaaaaa;
}
DIV.Email {
	font-size: 9pt;
	font-weight: bold;
	margin-left: 7px;
}
DIV.WebURL {
	font-size: 9pt;
	font-weight: bold;
	margin-left: 7px;
	margin-right: 5px;
}
IMG.UserImage { border: 1px solid <%=g_sCOLOR_EDGE%>; margin-bottom: 8px; }
DIV.About { margin-top: 4px; }
</style>
</head>
<body>
<!--#include file="./sundance/sundance_header.inc"-->
<!--include file="./sundance/sundance_ad-upper-middle.inc"-->
<!--#include file="./include/kb_message.inc"-->
<% Set m_oLayout = New kbLayout : Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) : Set m_oLayout = Nothing %>
<center>
<br>
<% Set m_oUser = New kbUser : Call m_oUser.WriteUser(m_lUserID) : Set m_oUser = Nothing %>
</center>
<!--include file="./sundance/sundance_ad-bottom-middle.inc"-->
<!--#include file="./sundance/sundance_footer.inc"-->
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language="javascript" src="./script/kb_validation.js"></script>
<script language='javascript'>
var m_oFields = {
	fldSubject:{desc:"e-mail Subject",type:"String",req:1},
	fldBody:{desc:"e-mail Message",type:"Posting",req:1}};
</script>