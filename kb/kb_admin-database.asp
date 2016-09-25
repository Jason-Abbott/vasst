<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_verify-admin_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_layout_cls.asp"-->
<!--#include file="./include/kb_files_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_database_cls.asp"-->
<%
Const m_sFORM_NAME = "frmDatabase"
dim m_oDatabase
dim m_oData
dim m_oLayout
dim m_bRunQuery

m_bRunQuery = false
Set m_oDatabase = New kbDatabase
Set m_oLayout = New kbLayout

select case Trim(Request.QueryString("do"))
	case "compact"
		Call m_oDatabase.CompactDatabase(false)
		Call SetSessionValue(g_USER_MSG, "The database has been compacted")
	case "backup"
		Call m_oDatabase.CompactDatabase(true)
		Call SetSessionValue(g_USER_MSG, "The database has been backed up")
	case "save"
		Call m_oDatabase.SaveQuery(Request.Form("fldSaveQuery"), Request.Form("fldQuery"))
	case "run"
		m_bRunQuery = true
end select
%>
<html>
<head>
<title>Administration: Database</title>
<link href="./style/kb_common.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_site.css" rel="stylesheet" type="text/css">
<link href="./style/<%=g_lSiteID%>/kb_admin.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="./sundance/sundance_header.inc"-->
<!--#include file="./include/kb_message.inc"-->
<% Call m_oLayout.WriteMenuBar(m_sMENU_COMMON) %>
<% Call m_oLayout.WriteMenuBar(m_sMENU_ADMIN) %>
<center>

<form name='<%=m_sFORM_NAME%>' method='post' action='kb_admin-database.asp'>

<div style='font-size: 9pt; border-top: 1px solid #6699FF; border-bottom: 1px solid #6699FF; width: 275px;'>
<div style='font-size: 8pt; color: #999999;'><%=Server.Mappath(g_sDB_LOCATION)%></div>
<nobr>The current database size	is <% Call m_oDatabase.WriteDatabaseSize()%> bytes</nobr></div>
<p>
<table cellspacing='0' cellpadding='4' border='0' width='70%'>
<tr>
	<td class='dbAction' valign='top'><nobr><a href='kb_admin-database.asp?do=compact'><%=m_oLayout.WriteToggleImage("btn_compact-database", "", "Compact Database", "width='124' height='14'", false)%></a></nobr></td>
	<td class='dbInfo' valign='top'>The database <b><% Call m_oDatabase.WriteLastCompactDate() %></b>.  Deleted and modified records continue to take up space in the database and, over time, can impact performance.  It is recommended that you backup the database before compacting it.</td>
<tr>
	<td class='dbAction' valign='top'><a href='kb_admin-database.asp?do=backup'><%=m_oLayout.WriteToggleImage("btn_backup-database", "", "Backup Database", "width='124' height='14'", false)%></a></td>
	<td class='dbInfo' valign='top'>The database <b><% Call m_oDatabase.WriteLastBackupDate() %></b>.  Backing up the database consists of making a copy of the live file with the current date appended to the file name.</td>
<!-- <tr>
	<td class='dbAction' valign='top'><nobr><a href='javascript:Restore();'><%'m_oLayout.WriteToggleImage("btn_restore-database", "", "Restore Database", "width='124' height='14'", false)%></a></nobr></td>
	<td class='dbInfo' valign='top'><% 'Call m_oDatabase.WriteBackupList() %><br>Select the backup you'd like to restore.  <b>This will overwrite the live database</b>.  If the live database appears to be corrupt, try compacting it first as that process will sometimes correct errors.</td>
 --><tr>
	<td class='dbAction' valign='top'><a href='javascript:RunQuery();'><%=m_oLayout.WriteToggleImage("btn_run-query", "", "Run Query", "width='124' height='14'", false)%></a></td>
	<td class='dbInfo' valign='top'><textarea cols='70' rows='6' name='fldQuery'><%=Request.Form("fldQuery")%></textarea><br>
	<a href='javascript:SaveQuery();'><%=m_oLayout.WriteToggleImage("btn_save-query", "", "Save Query", "width='81' height='14'", false)%></a>
	<input type='text' name='fldSaveQuery' size='15' maxlength='25'>
	<a href='javascript:LoadQuery();'><%=m_oLayout.WriteToggleImage("btn_load-query", "", "Load Query", "width='81' height='14'", false)%></a>
	<select name='fldLoadQuery'></select>
	</td>
</table>
</form>

<% If m_bRunQuery Then Call m_oDatabase.WriteQueryResults(Request.Form("fldQuery")) %>

</center>
<!--#include file="./include/kb_footer.inc"-->
</body>
</html>

<script language="javascript" src="./script/kb_functions.js"></script>
<script language='javascript'>
var m_oForm = document.<%=m_sFORM_NAME%>;
var m_aQueries = [<%=m_oDatabase.WriteQueryArray()%>];
var m_QUERY_ID = 0;
var m_QUERY_NAME = 1;
var m_QUERY = 2;

BuildQueryList();

function SaveQuery() {
	with (m_oForm) { 
		if (fldSaveQuery.value != "" && fldQuery != "") {
			action += "?do=save"; submit();
		} else {
			alert("Please enter a query and name for your query  ");
		}
	}
}

function LoadQuery() {
	var oList = m_oForm.fldLoadQuery;
	var lQueryID = oList.options[oList.selectedIndex].value;
	for (var x = 0; x < m_aQueries.length; x++) {
		if (lQueryID == m_aQueries[x][m_QUERY_ID]) {
			m_oForm.fldQuery.value = m_aQueries[x][m_QUERY];
		}	
	}
}

function BuildQueryList() {
	var oList = m_oForm.fldLoadQuery;
	for (var x = 0; x < m_aQueries.length; x++) {
		oList.options[x] = new Option(m_aQueries[x][m_QUERY_NAME], m_aQueries[x][m_QUERY_ID]);
	}
}

function RunQuery() { m_oForm.action += "?do=run"; m_oForm.submit(); }

function Restore() {
	var oField = m_oForm.fldBackup;
	var sDatabase = oField.options[oField.selectedIndex].text;
	alert("You selected " + sDatabase + "\nNot coded yet");
}

</script>
<% Set m_oDatabase = Nothing : Set m_oLayout = Nothing %>