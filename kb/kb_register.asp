<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_mail_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<%
Const m_sFORM_NAME = "frmRegister"
dim m_sMessage
dim m_sFirstName
dim m_sLastName
dim m_sScreenName
dim m_sEmail
dim m_sPassword
dim m_sHomePage
dim m_bNotify
dim m_bPrivacy
dim m_bRemind
dim m_bUsedEmail
dim m_lTimeShift	' client-server time difference in seconds
dim g_lSiteID
dim m_oLayout
dim m_oUser
dim m_oMail

m_bUsedEmail = false
With Request
	m_sFirstName = Trim(.Form("fldFirstName"))
	m_sLastName = Trim(.Form("fldLastName"))
	m_sScreenName = Trim(.Form("fldScreenName"))
	m_sEmail = Trim(.Form("fldEmail"))
	m_sPassword = Trim(.Form("fldPassword"))
	m_sHomePage = Trim(.Form("fldHomePage"))
	m_bRemind = CBool(.Form("fldRemind") = "yes")
	m_bNotify = CBool(.Form("fldNotify") = "on" Or IsVoid(m_sFirstName))
	m_bPrivacy = CBool(.Form("fldPrivacy") = "on")
	g_lSiteID = MakeNumber(.QueryString("s"))
End With

'If g_lSiteID = 0 Then response.redirect [to forum select]

If m_bRemind Then
	Set m_oMail = New kbMail : Call m_oMail.SendPasswordEmail(m_sEmail) : Set m_oMail = Nothing
	m_sMessage = "Your password has been e-mailed to " & m_sEmail
ElseIf m_sFirstName <> "" Then
	m_lTimeShift = DateDiff("s", Request.Form("fldServerTime"), Request.Form("fldClientTime"))
	Set m_oUser = New kbUser
	m_sMessage = m_oUser.Register(m_sFirstName, m_sLastName, m_sScreenName, m_sEmail, m_sPassword, _
		m_sHomePage, m_bPrivacy, m_bNotify, m_lTimeShift, g_lSiteID)
	Set m_oUser = Nothing
	If IsVoid(m_sMessage) Then
		Call SetSessionValue(g_USER_MSG, g_sMSG_NEED_VALIDATION)
		Response.Redirect "kb_validate.asp?s=" & g_lSiteID & "&url=" & Request.QueryString("url")
	Else
		If InStr(m_sMessage, "e-mail") Then m_bUsedEmail = true
	End If
End If
%>
<html>
<head>
<title><%=g_sORG_NAME%>: Sign up</title>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
</head>
<body>
<% Set m_oLayout = New kbLayout %>
<!--#include file="./include/kb_header_inc.asp"-->
<!--#include file="./include/kb_message.inc"-->
<div class='AppName'><%=g_sORG_NAME%>&nbsp;<%=g_sAPP_NAME%></div>
<center>
<form name="<%=m_sFORM_NAME%>" action="kb_register.asp?s=<%=g_lSiteID%>&url=<%=Request.QueryString("url")%>" method="post" onSubmit='return SignUp();'>
<br>
<% Call m_oLayout.WriteTitleBoxTop(g_sAPP_NAME & " Sign Up", "", "") %>
<br>
<table cellspacing='0' cellpadding='0' border='0'>
<tr>
	<td class='Required'>Real Name:</td>
	<td class='FormInput'><input type='text' name='fldFirstName' size='15' value='<%=m_sFirstName%>' maxlength='50'>
		<input type='text' name='fldLastName' size='15' value='<%=m_sLastName%>' maxlength='50'></td>
<tr><td></td><td class='FormNote'>First and Last</td>
<tr>
	<td class='FormLabel'>Screen Name:</td>
	<td class='FormInput'><input type='text' name='fldScreenName' size='30' maxlength='30' value='<%=m_sScreenName%>'></td>
<tr><td></td><td class='FormNote'>Leave blank to use only your real name</td>
<tr>
	<td class='Required'>e-mail:</td>
	<td class='FormInput'><input type='text' name='fldEmail' size='25' value='<%=m_sEmail%>' maxlength='50'></td>
<tr><td></td><td class='FormNote'>A validation code will be sent to this address</td>
<tr>
	<td class='Required'>Password:</td>
	<td class='FormInput'><input type='password' name='fldPassword' value='<%=m_sPassword%>' maxlength='50'></td>
<tr><td></td><td class='FormNote'>At least <%=g_MIN_PASSWORD_LENGTH%> characters</td>
<tr>
	<td class='Required'>Confirm:</td>
	<td class='FormInput'><input type='password' name='fldConfirm' value='<%=m_sPassword%>' maxlength='50'></td>
<tr><td></td><td class='FormNote'>Please type your password again</td>
<tr>
	<td class='FormLabel'>Privacy:</td>
	<td class='FormInput'><input type='checkbox' name='fldPrivacy' style='border: none;' <% if m_bPrivacy then %> checked<% end if %>></td>
<tr><td></td><td class='FormNote'>Hide e-mail address from other members</td>
<input type='hidden' name='fldNotify' value='on'>
<!-- <tr>
	<td class='FormLabel'>Notify:</td>
	<td class='FormInput'><input type='checkbox' name='fldNotify' style='border: none;' <% if m_bNotify then %>checked<% end if %>></td>
<tr><td></td><td class='FormNote'>Receive occasional news from <%=g_sORG_NAME%></td> -->
<tr>
	<td class='FormLabel'>Web page:</td>
	<td class='FormInput'>http://<input type='text' name='fldHomePage' size='25' value='<%=m_sHomePage%>' maxlength='50'></td>
<tr>
	<td class='Required' style='font-size: 8pt; text-align: center;'>(required)</td>
	<td class='FormInput' align='right'><% Call m_oLayout.WriteToggleImage("btn_sign-up", "", "Sign up", "width='80' height='14' class='Image'", true) %></td>
</table>

<input type='hidden' name='fldRemind'>
<input type='hidden' name='fldServerTime' value='<%=Now()%>'>
<input type='hidden' name='fldClientTime' value=''>

<% Call m_oLayout.WriteBoxBottom("") : Set m_oLayout = Nothing %>

<% if m_bUsedEmail then %>
<p>
<a href='JavaScript:SendPassword()'>Send the password for my account to <%=m_sEmail%></a>
<% end if %>

</form>
</center>
<!--#include file="./sundance/sundance_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language="javascript" src="./script/kb_validation.js"></script>
<script language="javascript">
var m_sMessage = '<%=m_sMessage%>'
var m_oForm = document.<%=m_sFORM_NAME%>;
var m_oFields = {
	fldFirstName:{desc:"First Name",type:"Name",req:1},
	fldLastName:{desc:"Last Name",type:"Name",req:1},
	fldEmail:{desc:"e-mail address",type:"Email",req:1},
	fldPassword:{desc:"Password",type:"Password",req:1},
	fldConfirm:{desc:"Password Confirmation",type:"Password",req:1},
	fldHomePage:{desc:"Web Page",type:"URL",req:0}};

m_oForm.fldClientTime.value = GetClientTime();
	
if (m_sMessage != "") { alert(m_sMessage); }

function SignUp() {
	if (IsValid('<%=m_sFORM_NAME%>', m_oFields)) {
		if (m_oForm.fldPassword.value.length < <%=g_MIN_PASSWORD_LENGTH%>) {
			alert("Your password must be at least six characters in length  "); return false;
		}
		if (m_oForm.fldPassword.value != m_oForm.fldConfirm.value) {
			alert("Your password and password confirmation do not match  "); return false;
		}
		return true;
	}
	return false;
}
function SendPassword() { m_oForm.fldRemind.value = "yes"; m_oForm.submit(); }
</script>