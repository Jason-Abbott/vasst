<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_verify-admin_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_contest_cls.asp"-->
<%
Const m_sFORM_NAME = "frmContest"

dim m_aData
dim m_oLayout
dim m_oContest

Set m_oContest = New kbContest
m_aData = m_oContest.GetContestArray(CBool(Request.QueryString("new") = "yes"))
if Request.Form("fldName") <> "" then Call m_oContest.SaveContest(m_aData)
%>
<html>
<head>
<title>Administration: Contests</title>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_admin.css" rel="stylesheet" type="text/css">
<style>
TD.ContestLabel {
	font-size: 9pt;
	text-align: right;
	padding-right: 6px;
	padding-top: 5px;
}
TD.ContestData {
	font-size: 8pt;
	color: #cccccc;
	padding-top: 5px;
}
TD.ContestNote {
	font-size: 8pt;
	color: #cccccc;
}
DIV.Factor {
	font-size: 8pt;
	color: #cccccc;
	margin-left: 3px;
	margin-bottom: 3px;
}
TD.SideBar {
	font-size: 8pt;
	width: 150px;
	padding-left: 10px;
}
DIV.SideTitle {
	font-size: 9pt;
	font-weight: bold;
	text-align: center;
	border-bottom: 1px solid <%=g_sCOLOR_EDGE%>;
	margin-bottom: 2px;
}
DIV.SideP {
	margin-bottom: 5px;
	text-indent: 10px;
}
TD.WeightHead {
	font-size: 9pt;
	border-bottom: 1px solid <%=g_sCOLOR_EDGE%>;
	border-right: 1px solid <%=g_sCOLOR_EDGE%>;
	text-align: center;
}
TD.WeightLabel {
	font-size: 9pt;
	border-bottom: 1px solid <%=g_sCOLOR_EDGE%>;
	border-right: 1px solid <%=g_sCOLOR_EDGE%>;
	text-align: center;
}
TD.WeightData {
	font-size: 9pt;
	border-bottom: 1px solid <%=g_sCOLOR_EDGE%>;
	border-right: 1px solid <%=g_sCOLOR_EDGE%>;
	padding-left: 4px;
	padding-right: 4px;
}
</style>
<meta name="Microsoft Border" content="none, default">
</head>
<body>
<!--#include file="./sundance/sundance_header.inc"-->
<!--#include file="./include/kb_message.inc"-->
<% Set m_oLayout = New kbLayout : Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<% Call m_oLayout.WriteMenuBar(m_sMENU_ADMIN) %>
<center>

<form name='<%=m_sFORM_NAME%>' method='post' action='kb_admin-contests.asp?id=<%=m_aData(m_CONTEST_ID)%>' onSubmit="return isValid('<%=m_sFORM_NAME%>',m_oFields);">
<% Call m_oContest.WriteContestList("fldContests", m_aData(m_CONTEST_ID)) %>
&nbsp;<a href='kb_admin-contests.asp?new=yes'><% Call m_oLayout.WriteToggleImage("btn_add-contest", "", "Add a new contest", "", false) %></a>
&nbsp;<a href='javascript:ViewContest();'><% Call m_oLayout.WriteToggleImage("btn_view", "", "View contest", "", false) %></a>
<p>
<table>
<tr>
	<td valign='top'>

<% Call m_oLayout.WriteTitleBoxTop(ReplaceNull(m_aData(m_CONTEST_NAME), "New") & " Contest", "", "") %>
<table cellspacing='0' cellpadding='0' border='0'>
<tr>
	<td class='ContestLabel'>Name:</td>
	<td class='ContestData'><input type='text' name='fldName' value='<%=m_aData(m_CONTEST_NAME)%>' maxlength='50'></td>
<tr>
	<td class='ContestLabel'>Site:</td>
	<td class='ContestData'><select name='fldSite'><%Call m_oContest.WriteSiteList(m_aData(m_CONTEST_SITE))%></select></td>
<tr>
	<td class='ContestLabel'>Start Date:</td>
	<td class='ContestData'><input type='text' name='fldStart' value='<%=m_aData(m_CONTEST_START)%>' maxlength='10' size='10'> (month/day/year)</td>
<tr>
	<td class='ContestLabel'>Vote by Date:</td>
	<td class='ContestData'><input type='text' name='fldVoteBy' value='<%=m_aData(m_CONTEST_VOTE_BY)%>' maxlength='10' size='10'></td>
<tr>
	<td class='ContestLabel'>End Date:</td>
	<td class='ContestData'><input type='text' name='fldEnd' value='<%=m_aData(m_CONTEST_END)%>' maxlength='10' size='10'> Contest will disappear after this date</td>
<tr>
	<td class='ContestLabel'>Winners:</td>
	<td class='ContestData'><input type='text' name='fldWinners' value='<%=m_aData(m_CONTEST_WINNERS)%>' maxlength='2' size='2'>
	Number of winners to be selected</td>
<tr>
	<td class='ContestLabel'>Votes Allowed:</td>
	<td class='ContestData'><input type='text' name='fldVotes' value='<%=m_aData(m_CONTEST_VOTES)%>' maxlength='2' size='2'>
	Number of <b>different</b> files a user can vote for</td>
<tr>
	<td class='ContestLabel'>Vote Weight:</td>
	<td class='ContestData'><input type='text' name='fldWeight' value='<%=m_aData(m_CONTEST_WEIGHT)%>' size='2' maxlength='1'> See sidebar for explanation</td>
<tr>
	<td class='ContestLabel'>Entries Allowed:</td>
	<td class='ContestData'><input type='text' name='fldEntries' value='<%=m_aData(m_CONTEST_MAX_ENTRIES)%>' maxlength='2' size='2'>
	Number of files a user can enter into the contest</td>
<tr>
	<td class='ContestLabel'>Restrictions:</td>
	<td class='ContestData'><input type='checkbox' name='fldFreePlugins' style='border: none;' <%if m_aData(m_CONTEST_FREE_PLUGINS_ONLY) then%> checked<%end if%>>Only free plugins
	<input type='checkbox' name='fldExternalMedia' style='border: none;' <%if m_aData(m_CONTEST_NO_EXTERNAL_MEDIA) then%> checked<%end if%>>No external media</td>
	
<tr>
	<td class='ContestLabel' valign='top'>Description:</td>
	<td class='ContestData'><textarea cols='55' rows='10' name='fldDescription'><%=m_aData(m_CONTEST_TEXT)%></textarea></td>
<tr><td></td><td class='ContestNote'>HTML is allowed</td>
</table>
<div align='right'><% Call m_oLayout.WriteToggleImage("btn_save", "", "Save Contest", "class='Image'", true) %></div>

<% Call m_oLayout.WriteBoxBottom("") %>

	</td>
	<td valign='top' class='SideBar'>
	<div class='SideTitle'><nobr>About Weighting</nobr></div>
	<div class='SideP'>The "weight" setting controls how multiple user votes are counted, if allowed.</div>
	<div class='SideP'>If translated to points, the last choice always receives one point with successively higher choices differentiated by the weight factor.</div>
	<div class='SideP'>For example, if three votes are allowed, the files voted for would receive points as follows:</div>
	<center>
	<table cellspacing='0' cellpadding='1' border='0'>
	<tr>
		<td class='WeightHead'><sub>Weight</sub> <sup>Vote</sup></td>
		<td class='WeightHead'>1</td><td class='WeightHead'>2</td><td class='WeightHead'>3</td>
	<tr>
		<td class='WeightLabel'>0</td>
		<td class='WeightData'>1</td><td class='WeightData'>1</td><td class='WeightData'>1</td>
	<tr>
		<td class='WeightLabel'>1</td>
		<td class='WeightData'>3</td><td class='WeightData'>2</td><td class='WeightData'>1</td>
	<tr>
		<td class='WeightLabel'>2</td>
		<td class='WeightData'>5</td><td class='WeightData'>3</td><td class='WeightData'>1</td>
	<tr>
		<td class='WeightLabel'>3</td>
		<td class='WeightData'>7</td><td class='WeightData'>4</td><td class='WeightData'>1</td>
	<tr>
		<td></td>
		<td colspan='3' align='center'>points</td>
	</table>
	</center>
	</td>
</table>
</form>

</center>
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<% Set m_oLayout = Nothing : Set m_oContest = Nothing %>

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

function SwitchContest(r_oField) {
	location.href = "kb_admin-contests.asp?id=" + r_oField.options[r_oField.selectedIndex].value;
}
function ViewContest() {
	var oField = m_oForm.fldContests;
	location.href = "kb_contest.asp?id=" + oField.options[oField.selectedIndex].value;
}
</script>