<% pageTitle = "Purge Customers" %>
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

If (Request.Form("purgeButton") = "Purge Customers") Then
	purgeCustomers
	Response.Redirect("customer_purge.asp")
End If
%>
<b>WARNING:</b> This utility is used to remove all Deleted? flagged customer from the customer database, this is a permanent change.<br>
<br>
<b>The Following Customers will be Purged:</b><br>
<blockquote>
<% printToBePurged %>
</blockquote>
<center><input type="submit" name="purgeButton" value="Purge Customers"></center>
<%
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