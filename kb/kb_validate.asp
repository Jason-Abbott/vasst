<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_mail_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<%
Const m_sFORM_NAME = "frmValidate"
dim g_lSiteID
dim m_sMessage
dim m_sURL
dim m_sValidate
dim m_lUserID
dim m_oUser
dim m_oMail
dim m_oLayout

g_lSiteID = GetSessionValue(g_USER_SITE)
m_sValidate = Trim(Request.QueryString("v"))
If m_sValidate = "" Then
	m_sValidate = Trim(Request.Form("fldValidate"))
	m_lUserID = GetSessionValue(g_USER_ID)
Else
	m_lUserID = Trim(Request.QueryString("id"))
End If
If Not IsNumber(m_lUserID) Then response.redirect "kb_login.asp"

If Request.QueryString("do") = "resend" Then
	Set m_oMail = New kbMail
	Call m_oMail.SendValidationEmail(m_lUserID)
	Set m_oMail = Nothing
ElseIf m_sValidate <> "" Then
	Set m_oUser = New kbUser : m_sMessage = m_oUser.Validate(m_sValidate, m_lUserID) : Set m_oUser = Nothing
	m_sURL = ReplaceNull(Request.QueryString("url"), g_sDEFAULT_PAGE)
	Call SetSessionValue(g_USER_MSG, ReplaceNull(m_sMessage, g_sMSG_REGISTER_THANKS))
	If IsVoid(m_sMessage) Then Response.Redirect m_sURL
End If
%>
<html>
<head>
<title>Validate</title>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<style>
DIV.Resend {
	margin-top: 10px;
	font-size: 8pt;
}
DIV.Note {
	margin-top: 10px;
	font-size: 8pt;
	width: 250px;
}
</style>
<meta name="Microsoft Border" content="none, default">
</head>
<body>
<!--#include file="./sundance/sundance_header.inc"-->
<!--include file="./sundance/sundance_ad-upper-middle.inc"-->
<!--#include file="./include/kb_message.inc"-->
<div class='AppName'><%=g_sORG_NAME%>&nbsp;<%=g_sAPP_NAME%></div>
<center>
<form name="<%=m_sFORM_NAME%>" action="kb_validate.asp?url=<%=m_sURL%>" method="post" onSubmit="return isValid('<%=m_sFORM_NAME%>', m_oFields);">

<% Set m_oLayout = New kbLayout : Call m_oLayout.WriteBoxTop("", "") %>

<table cellspacing='0' cellpadding='0' border='0'>
<tr>
	<td class='FormLabel'>Validation Code:</td>
	<td class='FormInput'><input type='text' name='fldValidate' size='16' value='<%=m_sValidate%>'></td>
<tr>
	<td></td><td class='FormNote'>enter the code you received in e-mail</td>
</table>
<p>
<div align='center'>
<% Call m_oLayout.WriteToggleImage("btn_validate-my-registration", "", "Validate Registration", "height='14' width='161' class='Image'", true) %>
</div>
<% Call m_oLayout.WriteBoxBottom("") : Set m_oLayout = Nothing %>
<div class='Resend'>Please <a href='kb_validate.asp?url=<%=m_sURL%>&do=resend'>resend my validation code</a></div>
<div class='Note'>If you do not receive the validation code it usually indicates that the e-mail address you signed up with is not valid or temporarily inaccessible.</div>
</form>
</center>
<!--#include file="./sundance/sundance_footer.inc"-->
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language="javascript" src="./script/kb_validation.js"></script>
<script language="javascript">
var m_oFields = {
	fldValidate:{desc:"Validation Code",type:"String",req:1}};
</script>