<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_verify-admin_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_categories_cls.asp"-->
<%
Const m_sFORM_NAME = "frmCategories"

dim m_aData
dim m_oLayout
dim m_oCat

With Request
	if Trim(.Form("fldCatName")) <> "" then
		Set m_oCat = New kbCategories
		Call m_oCat.Save(.Form("fldCatID"), .Form("fldCatName"), .Form("fldCatItems"))
		Set m_oCat = Nothing
	end if
End With
%>
<html>
<head>
<title>Administration: Categories</title>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_admin.css" rel="stylesheet" type="text/css">
<style>
TD.CatLabel {
	text-align: right;
	padding-right: 4px;
	padding-top: 4px;
}
TD.CatData {
	padding-top: 4px;
}
</style>
</head>
<body>
<% Set m_oLayout = New kbLayout %>
<!--#include file="./include/kb_header_inc.asp"-->
<!--#include file="./include/kb_message.inc"-->
<% Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<% Call m_oLayout.WriteMenuBar(m_sMENU_ADMIN) %>
<center>
<form name='<%=m_sFORM_NAME%>' method='post' action='kb_admin-categories.asp' onSubmit="return IsValid('<%=m_sFORM_NAME%>',m_oFields);">
<% Call m_oLayout.WriteTitleBoxTop("Categories", "", "") %>
<table cellspacing='0' cellpadding='0' border='0'>
<tr>
	<td class='CatLabel'>Category:</td>
	<td class='CatData'>
	<% Set m_oCat = New kbCategories : Call m_oCat.WriteOptionList("fldCategory", "", true, "onChange='changeCat(this);'") %>
	</td>
<tr>
	<td class='CatLabel'>Name:</td>
	<td class='CatData'><input type='text' name='fldCatName' maxlength='25'></td>
<tr>
	<td class='CatLabel' valign='top'>Valid for:</td>
	<td class='CatData'>
	<table width='100%' cellspacing='0' cellpadding='0' border='0'>
	<tr>
		<td><select name='fldCatItems' multiple size='5'>
		<%=MakeList("SELECT lItemTypeID, vsDescription FROM tblItemTypes ORDER BY vsDescription", "") %>
		</select>
		</td>
		<td align='right' valign='bottom'>
		<% Call m_oLayout.WriteToggleImage("btn_save", "", "Save Category", "class='Image'", true) %>
		</td>
	</table>
	</td>
</table>

<% Call m_oLayout.WriteBoxBottom("") %>
<input type='hidden' name='fldCatID'>
</form>

</center>
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_validation.js"></script>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language='javascript'>
var m_aCategories = [<% m_oCat.WriteJSArray() %>];
var m_oForm = document.<%=m_sFORM_NAME%>;
var m_oFields = {
	fldCatName:{desc:"Category Name",type:"String",req:1},
	fldCatItems:{desc:"Items category applies to",type:"Select",req:1}};

function newCat() {
	with (m_oForm) {
		fldCatName.value = ""
		fldCatID.value = ""
		for (var x = 0; x < fldCatItems.options.length; x++) {
			fldCatItems.options[x].selected = false;
		}
	}
}
	
function changeCat(r_oField) {
	var CAT_ID = 0;
	var CAT_NAME = 1;
	var CAT_ITEM = 2;
	var lCatID = parseInt(r_oField.options[r_oField.selectedIndex].value);
	if (lCatID == 0) { newCat(); }
	else {
		for (var x = 0; x < m_aCategories.length; x++) {
			if (parseInt(m_aCategories[x][CAT_ID]) == lCatID) {
				m_oForm.fldCatName.value = m_aCategories[x][CAT_NAME];
				m_oForm.fldCatID.value = lCatID
				selectItem(m_aCategories[x][CAT_ITEM]);
			}
		}
	}
}

function selectItem(r_lItemID) {
	var oField = m_oForm.fldCatItems;
	for (var x = 0; x < oField.options.length; x++) {
		if (oField.options[x].value == r_lItemID) { oField.options[x].selected = true; return; }
	}
}
</script>
<% Set m_oCat = Nothing : Set m_oLayout = Nothing %>