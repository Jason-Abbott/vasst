<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_tutorials_cls.asp"-->
<!--#include file="./include/kb_tutorial-data_cls.asp"-->
<%
Const m_sFORM_NAME = "frmSubmission"
dim m_oLayout
dim m_oTutorial
dim m_oUser
dim m_aData
dim m_bAdmin

Set m_oTutorial = New kbTutorialData
m_aData = m_oTutorial.GetArray()
If Trim(Request.Form("fldName")) <> "" Then	Call m_oTutorial.Save(m_aData)
m_bAdmin = CBool(GetSessionValue(g_USER_TYPE) = CStr(g_USER_ADMIN))
%>
<html>
<title><%=g_sORG_NAME%>: Submit your tutorial</title>
<head>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="./sundance/sundance_header.inc"-->
<!--include file="./sundance/sundance_ad-upper-middle.inc"-->
<!--#include file="./include/kb_message.inc"-->
<% Set m_oLayout = New kbLayout : Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<center>
<form name='<%=m_sFORM_NAME%>' action='kb_tutorial-edit.asp?id=<%=m_aData(m_TUTORIAL_ID)%>' method='post' onSubmit="return isValid('<%=m_sFORM_NAME%>',m_oFields);">
<% Call m_oLayout.WriteTitleBoxTop("Tutorial Submission", "", "") %>
<table cellspacing='0' cellpadding='0' border='0'>
<tr>
	<td class='Required'>Tutorial Name:</td>
	<td class='FormInput' valign='top'><input type='text' name='fldName' size='30' maxlength='75' value='<%=m_aData(m_TUTORIAL_NAME)%>'></td>
<tr><td></td><td class='FormNote'>maximum of 75 characters</td>
<tr>
	<td class='Required'>URL:</td>
	<td class='FormInput' valign='top'><input type='text' name='fldURL' maxlength='75' size='30' value='<%=m_aData(m_TUTORIAL_URL)%>'></td>
<tr><td></td><td class='FormNote'>link to web address where pages are hosted</td>
<tr>
	<td class='FormLabel'>Author:</td>
	<td class='FormInput' valign='top'>
		<% Set m_oUser = New kbUser : Call m_oUser.WriteUserList("fldAuthor", m_aData(m_TUTORIAL_AUTHOR_ID), g_STATUS_APPROVED, "not listed") %>
		<% if m_bAdmin then %>
			<a href='kb_user-edit.asp?new=yes&url=<%=GetURL(false)%>'><% Call m_oLayout.WriteToggleImage("btn_add-user", "", "Add User", "", false) %></a>
		<% end if %>
	</td>
<tr>
	<td class='Required' valign='top'>Description:</td>
	<td class='FormInput' valign='top'><textarea rows='6' cols='45' name='fldDescription'><%=m_aData(m_TUTORIAL_TEXT)%></textarea></td>
<tr><td></td><td class='FormNote'><%=g_sMSG_HTML_LIMIT%>&nbsp;</td>
<tr>
	<td class='FormLabel' valign='top'>Categories:</td>
	<td>
	<table width='100%' cellspacing='0' cellpadding='2' border='0'>
	<tr>
		<td rowspan='2' class='FormInput'>
		<% Call m_oLayout.WriteCategoryList("fldCategories", m_aData(m_TUTORIAL_CATS), 6, g_ITEM_TUTORIAL) %>
		</td>
		<td valign='top' class='FormNote'><br><%=g_sMSG_MULTI_SELECT%></td>
	<tr>
		<td align='right' valign='bottom'>
		<% Call m_oLayout.WriteToggleImage("btn_save", "", "Save Tutorial", "class='Image'", true) %>
		</td>
	</table>
	</td>
</table>
<% Call m_oLayout.WriteBoxBottom("") %>
<% Set m_oLayout = Nothing : Set m_oTutorial = Nothing %>

</form>
</center>
<!--include file="./sundance/sundance_ad-bottom-middle.inc"-->
<!--#include file="./sundance/sundance_footer.inc"-->
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language="javascript" src="./script/kb_validation.js"></script>
<script language="javascript">
var m_oForm = document.<%=m_sFORM_NAME%>;
var m_oFields = {
	fldName:{desc:"Tutorial Name",type:"String",req:1},
	fldURL:{desc:"URL for tutorial",type:"URL",req:1},
	fldDescription:{desc:"Description",type:"Posting",req:1}};

SelectCategories();
</script>