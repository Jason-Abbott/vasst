<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_script-data_cls.asp"-->
<!--#include file="./include/kb_file-system_cls.asp"-->
<meta HTTP-EQUIV="content-type" CONTENT="application/octet-stream">
<%
dim m_oScriptData
dim m_sDownloadURL

Set m_oScriptData = New kbScriptData
m_sDownloadURL = m_oScriptData.GetDownloadURL(Request.QueryString("id"))
Set m_oScriptData = Nothing

response.redirect m_sDownloadURL
%>