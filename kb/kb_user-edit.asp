<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<%
Const m_sFORM_NAME = "frmUser"
dim m_oLayout
dim m_oUser
dim m_bNew
dim m_aData
dim m_sTitle

m_bNew = CBool(Request.QueryString("new") = "yes" And GetSessionValue(g_USER_TYPE) = CStr(g_USER_ADMIN))
Set m_oUser = New kbUser
m_aData = m_oUser.GetArray(m_bNew)
Set m_oUser = Nothing
m_sTitle = IIf(m_bNew, "Add User", "Account Settings")
%>
<html>
<title><%=g_sORG_NAME%>: My Account</title>
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
<form name='<%=m_sFORM_NAME%>' action='kb_user-save.asp' method='post' onSubmit='return IsValidUser();' enctype="multipart/form-data">
<% Call m_oLayout.WriteTitleBoxTop(m_sTitle, "", "") %>

<table cellspacing='0' cellpadding='0' border='0'>
<tr>
	<td class='Required'>Real Name:</td><td class='FormInput'>
	<input type='text' name='fldFirstName' maxlength='50' value='<%=m_aData(m_FIRST_NAME)%>' size='15'>
	<input type='text' name='fldLastName' maxlength='50' value='<%=m_aData(m_LAST_NAME)%>' size='15'>
	</td>
<tr><td></td><td class='FormNote'>first and last name</td>
<tr>
	<td class='FormLabel'>Screen Name:</td><td class='FormInput'>
	<input type='text' name='fldScreenName' maxlength='30' value='<%=m_aData(m_SCREEN_NAME)%>' size='30'>
	</td>
<tr><td></td><td class='FormNote'>leave blank to use only your real name</td>
<tr>
	<td class='FormLabel'>Web site:</td><td class='FormInput'>
	<input type='text' name='fldWebURL' value='http://<%=m_aData(m_USER_WEB_URL)%>' maxlength='50' size='30'>
	</td>
<tr>
	<td class='Required'>e-mail:</td><td class='FormInput'>
	<input type='text' name='fldEmail' size='30' maxlength='50' value='<%=m_aData(m_USER_EMAIL)%>'>
	</td>
<%if Not m_bNew then%>
<tr><td></td><td class='FormNote'>changing your e-mail address will require re-verification
	<input type='hidden' name='fldOldEmail' value='<%=m_aData(m_USER_EMAIL)%>'>
	</td>
<%end if%>
<tr>
	<td class='Required'>Password:</td><td class='FormInput'>
	<input type='password' name='fldPassword' maxlength='20' value='' size='10'>
	<input type='password' name='fldConfirm' maxlength='20' value='' size='10'>
</td><tr><td></td><td class='FormNote'><%if Not m_bNew then%>leave blank for no change / <%end if%>enter twice to confirm</td>
<tr>
	<td class='FormLabel'>Privacy:</td><td class='FormInput'>
	<input type='checkbox' name='fldPrivacy' style='border: none;' <% if m_aData(m_USER_PRIVATE) then %> checked<% end if %>>
	hide e-mail address from other members
	</td>
<input type='hidden' name='fldNotify' value='on'>
<!-- <tr>
	<td class='FormLabel'>Notify:</td><td class='FormInput'>
	<input type='checkbox' name='fldNotify' style='border: none;' <% if m_aData(m_USER_SPAM) then %> checked<% end if %>>
	receive occasional news from <%=g_sORG_NAME%>
	</td>
 --><tr>
	<td class='FormLabel'>Sorting:</td><td class='FormInput'>
	<% Set m_oUser = New kbUser : Call m_oUser.WriteSortList("fldSortBy", m_aData(m_ITEM_SORT)) %>
	</td>
<tr><td></td><td class='FormNote'>default sorting of lists</td>
<tr>
	<td class='FormLabel'>Page size:</td><td class='FormInput'>
	<% Call m_oUser.WritePageSizeList("fldItemsPerPage", m_aData(m_ITEMS_PER_PAGE)) %> items per page
	</td>
<tr>
	<td class='FormLabel'>Format:</td><td class='FormInput'>
	<% Call m_oUser.WriteFormatList("fldFormat", m_aData(m_USER_FILE_FORMAT)) : Set m_oUser = Nothing %>
	default format for file uploads</td>
<tr>
	<td class='FormLabel' valign='top'>About you:</td><td class='FormInput'>
	<textarea name='fldBio' cols='45' rows='4'><%=m_aData(m_USER_ABOUT)%></textarea>
	</td>
<tr><td></td><td class='FormNote'><%=g_sMSG_HTML_LIMIT%></td>
<tr>
	<td class='FormLabel' valign='top'>Picture:</td>
	<td class='FormInput'><input type='file' name='fldPicture' size='30'></td>
<tr><td></td><td class='FormNote'>JPEG or GIF: maximum file size is <%=g_MAX_PICTURE_KB%> kilobytes<br>leave blank for no change</td>
<tr>
	<td class='Required' style='font-size: 8pt; text-align: center;'>(required)</td>
	<td align='right'><br>
	<% Call m_oLayout.WriteToggleImage("btn_save", "", "Save Changes", "class='Image'", true) %>
	</td>
</table>
<% Call m_oLayout.WriteBoxBottom("") : Set m_oLayout = Nothing %>
<input type='hidden' name='fldUserID' value='<%=m_aData(m_USER_ID)%>'>
<input type='hidden' name='fldNew' value='<%=m_bNew%>'>
<input type='hidden' name='fldReturnURL' value='<%=Request.QueryString("url")%>'>
</form>
</center>
<!--include file="./sundance/sundance_ad-bottom-middle.inc"-->
<!--#include file="./sundance/sundance_footer.inc"-->
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language="javascript" src="./script/kb_validation.js"></script>
<script language='javascript'>
var m_oForm = document.<%=m_sFORM_NAME%>;
var m_bNeedPassword = <%=VBtoSQLBoolean(m_bNew)%>;
var m_oFields = {
	fldFirstName:{desc:"first name",type:"Name",req:1},
	fldLastName:{desc:"last name",type:"Name",req:1},
	fldWebURL:{desc:"web site",type:"URL",req:0},
	fldEmail:{desc:"e-mail address",type:"Email",req:0},
	fldPassword:{desc:"password",type:"Password",req:m_bNeedPassword},
	fldConfirm:{desc:"password confirmation",type:"Password",req:m_bNeedPassword},
	fldBio:{desc:"about you",type:"Posting",req:0},
	fldPicture:{desc:"picture (JPEG or GIF)",type:"Image",req:0}};
	
function IsValidUser() {
	if (m_oForm.fldWebURL.value == "http://") { m_oForm.fldWebURL.value = ""; }
	if (IsValid('<%=m_sFORM_NAME%>', m_oFields)) {
		if (m_oForm.fldPassword.value != m_oForm.fldConfirm.value) {
			alert("the passwords do not match "); return false;
		}
		return true;
	}
	return false;
}
</script>