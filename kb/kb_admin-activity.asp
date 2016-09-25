<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_verify-admin_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_log_cls.asp"-->
<%
Const m_sFORM_NAME = "frmActivity"
dim m_sStartDate
dim m_sEndDate
dim m_lActivityID
dim m_lUserID
dim m_lFileID
dim m_lSiteID
dim m_oLog
dim m_oLayout
dim m_oUser

m_sStartDate = ReplaceNull(Request.Form("fldStartDate"), DateAdd("d", -1, Date))
m_sEndDate = ReplaceNull(Request.Form("fldEndDate"), Date)
m_lActivityID = ReplaceNull(Request.Form("fldActivity"), 0)
m_lUserID = ReplaceNull(Request.Form("fldUser"), 0)
m_lFileID = ReplaceNull(Request.Form("fldFile"), 0)
m_lSiteID = ReplaceNull(Request.Form("fldSite"), GetSessionValue(g_USER_SITE))
%>
<html>
<head>
<title>Administration: Activity Log</title>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_admin.css" rel="stylesheet" type="text/css">
<style>
TD.LogHead {
	font-size: 9pt;
	text-align: center;
	font-weight: bold;
}
TD.LogDate {
	font-size: 9pt;
	text-align: center;
	font-weight: bold;
	padding-top: 2px;
	padding-bottom: 2px;
	border-top: 1px solid <%=g_sCOLOR_EDGE%>;
	background-color: #002266;
}
TD.LogTime {
	font-size: 9pt;
	padding-right: 8px;
	text-align: right;
}
TD.LogActivity { font-size: 9pt; padding-right: 8px; }
TD.LogUser { font-size: 9pt; padding-right: 8px; }
TD.LogSite { font-size: 9pt; }
TD.LogIP { padding-left: 8px; font-size: 9pt; }
TD.FormLabel { text-align: right; font-size: 9pt; }
</style>
</head>
<body>
<!--#include file="./sundance/sundance_header.inc"-->
<% Set m_oLayout = New kbLayout : Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<% Call m_oLayout.WriteMenuBar(m_sMENU_ADMIN) %>
<center>

<form name='<%=m_sFORM_NAME%>' action='kb_admin-activity.asp' method='post'>
<table>
<tr>
	<td rowspan='5' valign='middle'><% Call m_oLayout.WriteToggleImage("btn_show-activity", "", "Show Activity", "width='94' height='14' class='Image'", true) : Set m_oLayout = Nothing %></td>
	<td class='FormLabel'>between</td>
	<td><input type='text' name='fldStartDate' size='8' value='<%=m_sStartDate%>'>
and <input type='text' name='fldEndDate' size='8' value='<%=m_sEndDate%>'> (inclusive)
	</td>
<tr>
	<td class='FormLabel'>activity</td>
	<td><select name='fldActivity'>
	<option value='0'>All Activities
	<%=MakeList("SELECT lActivityID, vsActivityName FROM tblLoggedActivities ORDER BY vsActivityName", m_lActivityID)%>
	</select></td>
<tr>
	<td class='FormLabel'>user</td>
	<td>
	<% Set m_oUser = New kbUser : Call m_oUser.WriteUserList("fldUser", m_lUserID, "", "all users") : Set m_oUser = Nothing %>
	</td>
<tr>
	<td class='FormLabel'>file</td>
	<td><select name='fldFile'>
	<option value='0'>All Files
	<%=MakeList("SELECT lFileID, vsFriendlyName FROM tblFiles ORDER BY vsFriendlyName", m_lFileID)%>
	</select></td>
<tr>
	<td class='FormLabel'>site</td>
	<td><select name='fldSite'>
	<option value='0'>All Sites
	<%=MakeList("SELECT lSiteID, vsSiteName FROM tblSite ORDER BY vsSiteName", m_lSiteID)%>
	</select></td>
</table>
	
</form>

<%
Set m_oLog = New kbLog
Call m_oLog.WriteActivity(m_sStartDate, m_sEndDate, m_lActivityID, m_lUserID, m_lFileID, m_lSiteID)
Set m_oLog = Nothing
%>

</center>
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>

<script language="javascript" src="./script/kb_functions.js"></script>