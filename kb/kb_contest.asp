<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_contest_cls.asp"-->
<%
Const m_sFORM_NAME = "frmContest"
Const m_sACT_VOTE = 1
Const m_sACT_ADD = 2
Const m_sACT_REMOVE = 3
dim m_aData
dim m_oLayout
dim m_oContest
dim m_bExpired

Set m_oContest = New kbContest
m_aData = m_oContest.GetContestArray(false)
m_bExpired = CBool(m_aData(m_CONTEST_VOTE_BY) < Date)

With Request
	Select Case MakeNumber(.Form("fldAction"))
		Case m_sACT_VOTE
			Call m_oContest.CastVote(m_aData(m_CONTEST_ID), m_aData(m_CONTEST_WEIGHT), _
				m_aData(m_CONTEST_VOTES), .Form("fldVotes"), g_ITEM_FILE)
		Case m_sACT_ADD
			Call m_oContest.AddItem(m_aData(m_CONTEST_ID), .Form("fldFiles"), g_ITEM_FILE)
		Case m_sACT_REMOVE
			Call m_oContest.RemoveItem(m_aData(m_CONTEST_ID), .Form("fldInContest"), g_ITEM_FILE)
	End Select
End With
%>
<html>
<title><%=g_sORG_NAME%>: Contest</title>
<head>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<style>
DIV.ContestNote {
	font-size: 8pt;
	text-align: center;
	margin-bottom: 4px;
}
TD.EntryLabel {
	font-size: 9pt;
	text-align: right;
	font-weight: bold;
	padding-top: 4px;
	border-top: 1px solid <%=g_sCOLOR_TITLE%>;
}
TD.EntryData {
	font-size: 9pt;
	padding-top: 4px;
	border-top: 1px solid <%=g_sCOLOR_TITLE%>;
}
TD.EntryButton {
	padding-bottom: 4px;
}
TD.VoteBar {
	background-color: <%=g_sCOLOR_EDGE%>;
	border: 1px solid <%=g_sCOLOR_TITLE%>;
	font-weight: bold;
	font-size: 9pt;
	margin-bottom: 2px;
	text-align: center;
	position: relative;
}
TD.VoteLabel {
	font-size: 9pt;
	text-align: right;
	font-weight: bold;
	padding-right: 8px;
	width: 10%;
}
TD.VoteHead {
	text-align: center;
	font-size: 18pt;
	padding-bottom: 4px;
	font-family: Impact, Arial, Helvetical;
	color: <%=g_sCOLOR_EDGE%>;
}
DIV.ContestName {
	text-align: center;
	font-size: 18pt;
	font-family: Impact, Arial, Helvetical;
	padding-top: 15px;
}
DIV.Congrats {
	text-align: center;
	font-weight: bold;
	font-size: 12pt;
	border-bottom: 1px solid <%=g_sCOLOR_EDGE%>;
}
TD.Winner {
	text-align: right;
	font-size: 12pt;
	padding-right: 7px;
	font-family: Arial, Helvetica;
}
TD.FileName {
	font-size: 12pt;
	font-family: Arial, Helvetica;
}
DIV.EditButton {
	text-align: right;
	margin-right: 10px;
}	
</style>
</head>
<body>
<!--#include file="./sundance/sundance_header.inc"-->
<!--#include file="./include/kb_message.inc"-->
<% Set m_oLayout = New kbLayout : Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<center>
<% if m_bExpired then %>
	<% Call m_oContest.WriteConclusion(m_aData) %>
	<p>
<% else %>

<form name='<%=m_sFORM_NAME%>' action='kb_contest.asp?id=<%=m_aData(m_CONTEST_ID)%>' method='post' onSubmit='return isValidVote();'>
<table>
<tr>
	<td valign='top'>
<% Call m_oContest.WriteContestVote(m_aData) %>
	</td>
	<td valign='top' width='220'>
<% Call m_oLayout.WriteTitleBoxTop(m_aData(m_CONTEST_NAME) & " Entry", "width='100%'", "") %>
<% Call m_oContest.WriteQualifyingItemList(m_aData) %>
<% Call m_oLayout.WriteBoxBottom("") %>
	</td>
</table>
<input type='hidden' name='fldAction' value=''>
</form>
<% end if %>
<% if g_bAdmin then %>
<div class='EditButton'><a href='kb_admin-contests.asp?id=<%=m_aData(m_CONTEST_ID)%>'><% Call m_oLayout.WriteToggleImage("btn_edit", "", "Edit contest", "", false) %></a></div>
<% end if %>
<% Call m_oContest.WriteContestStatus(m_aData(m_CONTEST_ID)) : Set m_oContest = Nothing : Set m_oLayout = Nothing %>
</center>
<!--#include file="./sundance/sundance_footer.inc"-->
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>
<script language="javascript" src="./script/kb_functions.js"></script>
<script language='javascript'>
var m_oForm = document.<%=m_sFORM_NAME%>;

function isValidVote() {
	var oField;
	var lValue;
	var bValid = false;
	for (var x = 1; x <= m_oForm.fldMaxVotes.value; x++) {
		oField = eval('m_oForm.fldVote' + x);
		lValue = parseInt(oField.options[oField.selectedIndex].value);
		if (lValue > 0) {
			if (m_oForm.fldVotes.value.indexOf("," + lValue + ",") != -1) {
				m_oForm.fldVotes.value = ",";
				alert("Please select files only once ");
				return false;
			}
			m_oForm.fldVotes.value += lValue + ",";
			bValid = true;
		}
	}
	if (bValid) {
		m_oForm.fldAction.value = '<%=m_sACT_VOTE%>';
	} else {
		alert("Please select at least one file to vote for ");
	}
	return bValid;
}

function AddToContest() { with (m_oForm) { fldAction.value = "<%=m_sACT_ADD%>"; submit(); } }
function RemoveFromContest() { 
	if (confirm("This will cancel any votes placed for your file.\nAre you sure you want to remove it?")) {
		with (m_oForm) { fldAction.value = "<%=m_sACT_REMOVE%>"; submit(); }
	}
}
</script>
