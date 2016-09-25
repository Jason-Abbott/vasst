<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_file-data_cls.asp"-->
<meta HTTP-EQUIV="content-type" CONTENT="application/octet-stream">
<%
dim m_oFileData
dim m_sDownloadURL

Set m_oFileData = New kbFileData
m_sDownloadURL = m_oFileData.GetDownloadURL(Request.QueryString("id"))
Set m_oFileData = Nothing

response.redirect m_sDownloadURL
%>