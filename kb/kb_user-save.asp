<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_verify_inc.asp"-->
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_form_cls.asp"-->
<!--#include file="./include/kb_file-data_cls.asp"-->
<!--#include file="./include/kb_user_cls.asp"-->
<!--#include file="./include/kb_encryption_cls.asp"-->
<!--#include file="./include/kb_mail_cls.asp"-->
<%
dim m_oUser
Set m_oUser = New kbUser
Call m_oUser.Save("fldPicture")
Set m_oUser = Nothing
%>