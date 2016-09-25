<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_form_cls.asp"-->
<!--#include file="./include/kb_files_cls.asp"-->
<!--#include file="./include/kb_file-data_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<%
dim m_oFileData
Set m_oFileData = New kbFileData
Call m_oFileData.Save("fldFile")
Set m_oFileData = Nothing
%>