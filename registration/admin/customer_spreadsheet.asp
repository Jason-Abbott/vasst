<% subPageTitle = "Customer Spreadsheet" %>
<!--#include file="functions.asp"-->
<FORM METHOD=post>
<SCRIPT>
this.focus()
</SCRIPT>
<% 
printHeader 
If Not (dontShowListOrEdit) Then

If (Request.Form("editButton") = "") Then
	printCustomers "view"
Else
	Response.Redirect("customer_editspreadsheet.asp")
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