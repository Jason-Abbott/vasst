<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/kb_constants_inc.asp"-->
<!--#include file="./include/kb_functions_inc.asp"-->
<!--#include file="./include/kb_data-access_cls.asp"-->
<!--#include file="./include/kb_encryption_cls.asp"-->
<%
dim squery
dim x
dim adata
dim odata
dim enc

squery = "select luserid, vspassword from tblusers"
set odata = new kbdataaccess
adata = odata.getarray(squery)

set enc = New kbEncryption
for x = 0 to ubound(adata, 2)
	squery = "update tblusers set vspassword = '" & enc.Encrypt(adata(1,x)) & "' where luserid = " & adata(0,x)
	call odata.executeonly(squery)
	response.write "converted " & adata(0,x)
next
set odata = nothing
set enc = nothing

%>
