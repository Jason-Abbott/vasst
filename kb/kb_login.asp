<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_mail_cls.asp"-->
<!--#include file="./include/kb_encryption_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<%
Const m_sFORM_NAME = "frmLogin"
dim m_sEmail
dim m_sPassword
dim m_sMessage
dim m_lTimeShift	' client-server time difference in seconds
dim g_lSiteID
dim m_sURL
dim m_bRemind
dim m_oUser
dim m_oMail
dim m_oLayout

With Request
	If .QueryString("logout") = "yes" Then
		Set m_oUser = New kbUser : Call m_oUser.Logout() : Set m_oUser = Nothing
	End If
	m_sURL = ReplaceNull(.QueryString("url"), g_sDEFAULT_PAGE)
	m_sEmail = Trim(.Form("fldEmail"))
	m_sPassword = Trim(.Form("fldPassword"))
	m_bRemind = CBool(.Form("fldRemind") = "yes")
	g_lSiteID = MakeNumber(ReplaceNull(.QueryString("s"), .Cookies(g_sSITE_COOKIE)))
End With

'If g_lSiteID = 0 Then response.redirect [to forum select page]

m_sMessage = ""

If m_bRemind Then
	Set m_oMail = New kbMail
	If m_oMail.SendPasswordEmail(m_sEmail) Then
		m_sMessage = "Your password has been e-mailed to " & m_sEmail
	Else
		m_sMessage = "Sorry, there was a problem sending to " & m_sEmail
	End If
	Set m_oMail = Nothing
ElseIf m_sEmail <> "" And m_sPassword <> "" Then
	m_lTimeShift = DateDiff("s", ReplaceNull(Request.Form("fldServerTime"), Now()), ReplaceNull(Request.Form("fldClientTime"), Now()))
	Set m_oUser = New kbUser
	m_sMessage = m_oUser.Login(m_sEmail, m_sPassword, m_lTimeShift, g_lSiteID)
	Set m_oUser = Nothing
	If IsVoid(m_sMessage) Then Response.Redirect m_sURL
End If
%>
<html>
<head>
<title>Sign in or up</title>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<style>
TD.LoginBox { font-size: 9pt; }
TD.LoginLabel { font-size: 9pt; text-align: right; font-weight: bold; }
DIV.LoginTitle { font-family: Impact, Arial; font-size: 20pt; margin-bottom: 5px; color: #6699FF; }
TD.LoginSignup { font-size: 9pt; border-left: 1px solid #6699FF; }
</style>
</head>
<body>
<% Set m_oLayout = New kbLayout %>
<!--#include file="./include/kb_header_inc.asp"-->
<div class='AppName'><%=g_sORG_NAME%>&nbsp;<%=g_sAPP_NAME%></div>
<center>
<p>
<% Call m_oLayout.WriteBoxTop("width='70%'", "") %>

<table width='100%' cellpadding='0' cellspacing='0' border='0'>
<form name="<%=m_sFORM_NAME%>" action="kb_login.asp?s=<%=g_lSiteID%>&url=<%=m_sURL%>" method="post" onSubmit="return IsValid('<%=m_sFORM_NAME%>', m_oFields);">
<tr>
	<td width='50%' align='center' valign='top'>
	<div class='LoginTitle'>Sign in</div>
	<table cellspacing='0' cellpadding='4' border='0'>
	<tr><td class='LoginLabel'>E-mail:</td><td><input type="text" name="fldEmail" size='20' value="<%=m_sEmail%>"></td>
	<tr><td class='LoginLabel'>Password:</td><td><input type="password" name="fldPassword" size='20' value='<%=m_sPassword%>'></td>
	</table>
	
	</td>
	<input type='hidden' name='fldRemind'>
	<td width='50%' class='LoginSignup' align='center' valign='top'>
	<div class='LoginTitle'>Sign up</div>
	Wanna' see great tutorials<br>and sample projects?
	<div style='font-size: 12pt; font-weight: bold; font-family: Arial, Tahoma;'>Then sign up now.</div>
	<div style='font-size: 11pt; color: #ffff00;'>It's fast and free!</div>
	<br>
	</td>
<tr>
	<td>
	<table width='100%' cellspacing='0' cellpadding='0' border='0'>
	<tr><td width='54'><% Call m_oLayout.WriteToggleImage("btn_sign-in", "", "Sign in", "width='80' height='14' class='Image'", true) %></td>
	<td align='right' style='font-size: 8pt;'><a href='javascript:Forgot();'>Forgot your password?</a>  &nbsp; </td>
	</table>
	</td>
	<td align='right' class='LoginSignup'><a href='kb_register.asp?s=<%=g_lSiteID%>&url=<%=m_sURL%>'><% Call m_oLayout.WriteToggleImage("btn_sign-up", "", "Sign up", "width='80' height='14'", false) %></a></td>
<input type='hidden' name='fldServerTime' value='<%=Now()%>'>
<input type='hidden' name='fldClientTime' value=''>
</form>	
</table>
<% Call m_oLayout.WriteBoxBottom("") : Set m_oLayout = Nothing %>



</center>
<!--#include file="./sundance/sundance_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language="javascript" src="./script/kb_validation.js"></script>
<script language="javascript">
var m_sMessage = '<%=m_sMessage%>';
var m_oForm = document.<%=m_sFORM_NAME%>;
var m_oFields = {
	fldEmail:{desc:"e-mail address",type:"Email",req:1},
	fldPassword:{desc:"password",type:"String",req:1}};

m_oForm.fldClientTime.value = GetClientTime();
	
if (m_sMessage != "") { alert(m_sMessage); }

function Forgot() {
	var sEmail = m_oForm.fldEmail.value;
	if (isEmail(m_oForm.fldEmail)) {
		if (confirm("Press OK if you'd like us to e-mail your password to you")) {
			m_oForm.fldRemind.value = "yes";
			m_oForm.submit();
		}
	} else {
		sEmail = prompt("Enter your registered e-mail address below\nand we will send your password to that address","");
		m_oForm.fldEmail.value = sEmail;
		if (isEmail(m_oForm.fldEmail)) {
			m_oForm.fldRemind.value = "yes";
			m_oForm.submit();
		} else {
			alert("Please enter a valid e-mail address ");
		}
	}
}
</script>