<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_verify-admin_inc.asp"-->
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_forums_cls.asp"-->
<!--#include file="./include/kb_forum-data_cls.asp"-->
<%
Const m_sFORM_NAME = "frmEdit"
dim m_oLayout
dim m_oForum
dim m_oUser
dim m_aData
dim m_bAdmin

Set m_oForum = New kbForumData
m_aData = m_oForum.GetArray()
If Trim(Request.Form("fldName")) <> "" Then Call m_oForum.Save(m_aData)
Set m_oForum = Nothing
m_bAdmin = CBool(GetSessionValue(g_USER_TYPE) = CStr(g_USER_ADMIN))
%>
<html>
<title><%=g_sORG_NAME%>: Submit your tutorial</title>
<head>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
</head>
<body>
<% Set m_oLayout = New kbLayout %>
<!--#include file="./include/kb_header_inc.asp"-->
<!--#include file="./include/kb_ads_inc.asp"-->
<!--#include file="./include/kb_message.inc"-->
<% Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<center>
<form name='<%=m_sFORM_NAME%>' action='kb_forum-edit.asp?id=<%=m_aData(m_FORUM_ID)%>' method='post' onSubmit="return IsValid('<%=m_sFORM_NAME%>',m_oFields);">
<% Call m_oLayout.WriteTitleBoxTop("Forum Edit", "", "") %>
<table cellspacing='0' cellpadding='0' border='0'>
<tr>
	<td class='Required'>Forum Name:</td>
	<td class='FormInput' valign='top'><input type='text' name='fldName' size='30' maxlength='50' value='<%=m_aData(m_FORUM_NAME)%>'></td>
<tr><td></td><td class='FormNote'>maximum of 50 characters</td>
<tr>
	<td class='Required'>URL:</td>
	<td class='FormInput' valign='top'><input type='text' name='fldURL' maxlength='75' size='30' value='<%=m_aData(m_FORUM_URL)%>'></td>
<tr><td></td><td class='FormNote'>link to web address where pages are hosted</td>
<tr>
	<td class='FormLabel'>Host:</td>
	<td class='FormInput' valign='top'>
		<% Set m_oForum = New kbForums : Call m_oForum.WriteHostList("fldHost", m_aData(m_FORUM_ID)) : Set m_oForum = Nothing %>
		<% if m_bAdmin then %>
			&nbsp; <a href='kb_host-edit.asp?new=yes&url=<%=GetURL(false)%>'><% Call m_oLayout.WriteToggleImage("btn_add-host", "", "Add Host", "", false) %></a>
		<% end if %>
	</td>
<tr>
	<td class='Required' valign='top'>Description:</td>
	<td class='FormInput' valign='top'><textarea rows='6' cols='45' name='fldDescription'><%=m_aData(m_FORUM_TEXT)%></textarea></td>
<tr><td></td><td class='FormNote'><%=g_sMSG_HTML_LIMIT%>&nbsp;</td>
<tr>
	<td class='FormLabel' valign='top'>Categories:</td>
	<td>
	<table width='100%' cellspacing='0' cellpadding='2' border='0'>
	<tr>
		<td rowspan='2' class='FormInput'>
		<% Call m_oLayout.WriteCategoryList("fldCategories", m_aData(m_FORUM_CATS), 6, g_ITEM_FORUM) %>
		</td>
		<td valign='top' class='FormNote'><br><%=g_sMSG_MULTI_SELECT%></td>
	<tr>
		<td align='right' valign='bottom'>
		<% Call m_oLayout.WriteToggleImage("btn_save", "", "Save Tutorial", "class='Image'", true) %>
		</td>
	</table>
	</td>
</table>
<% Call m_oLayout.WriteBoxBottom("") : Set m_oLayout = Nothing %>

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