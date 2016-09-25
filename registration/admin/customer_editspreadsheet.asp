<% pageTitle = "Customer Spreadsheet Editor" %>
<!--#include file="functions.asp"-->
<% 
If Not (getUserInfo(getUserID,"accesslevel") = "Admin") Then 
	Response.Write "<CENTER><B>You do not have access to this section of the CMS, your access is Read-Only, contact your admin to get rights.</B></CENTER>"
	Response.End
End If
%>
<FORM METHOD=post>
<SCRIPT>
this.focus()
</SCRIPT>
<% 
printHeader 
If Not (dontShowListOrEdit) Then

If (Request.Form("saveButton") = "") Then
	printCustomers "edit"
Else
	printCustomers "update"
	Response.Redirect("customer_spreadsheet.asp")
End If

Else
%>
<br>
<br>
<br>
<center>There are no customers in the database, please register a customer, or add one manually.</center>
<%
End If
printFooter
%>
</FORM>