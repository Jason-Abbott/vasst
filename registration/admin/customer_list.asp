<!--#include file="functions.asp"-->
<%
Response.Write("Test:" & Request.QueryString("searchFor"))
printHeader
If Not (dontShowListOrEdit) Then
%>
<FORM METHOD=post>
<%
'Response.Write("Form<BR>" & SplitForm & "Cookies<BR>" & SplitCookies)
If (Request.Form("edit") = "") Then
	printCustomers ""
	If (dontShowListOrEdit) Then
		Response.Redirect("customer_list.asp")
	End If
Else
	Response.Cookies("customerPass")("editID") = Request.Form("editID")
	Response.Redirect("customer_edit.asp")
End If
%>
<table border='0' cellpadding='0' cellspacing='0' width='100%' height='73'>
	<tr>
		<td width='25' class='bar' valign="top"><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center' valign="top">Account Information</td>
		<td width='100%' class='bar' valign="top"><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Database ID#:&nbsp;</td>
		<td colspan='2' valign="top"><%=editCustID%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Account Status:&nbsp;</td>
		<td colspan='2' valign="top"><%=editStatus%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Signed Up Date:&nbsp;</td>
		<td colspan='2' valign="top"><%=editSignupDate%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Deleted?:&nbsp;</td>
		<td colspan='2' valign="top"><%=editDeleted%></td>
	</tr>
	<tr>
		<td width='25' class='bar' valign="top"><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center' valign="top">Personal Information</td>
		<td width='100%' class='bar' valign="top"><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Name:&nbsp;</td>
		<td colspan='2' valign="top"><%=editFirstName%>&nbsp;<%=editLastName%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Company:&nbsp;</td>
		<td colspan='2' valign="top"><%=editCompanyName%></td>
	</tr>
    <tr>
		<td colspan='2' align='right' valign="top">Title:&nbsp;</td>
		<td colspan='2' valign="top"><%=editTitleName%></td>
	</tr>
	<tr>
		<td width='25' class='bar' valign="top"><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center' valign="top">Contact Information</td>
		<td width='100%' class='bar' valign="top"><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Phone:&nbsp;</td>
		<td colspan='2' valign="top"><%=editPhoneName%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Email Address:&nbsp;</td>
		<td colspan='2' valign="top"><%=editEmailName%></td>
	</tr>
	<tr>
		<td width='25' class='bar' valign="top"><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center' valign="top">Billing Address</td>
		<td width='100%' class='bar' valign="top"><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Address1:&nbsp;</td>
		<td colspan='2' valign="top"><%=editAddress1Name%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Address2:&nbsp;</td>
		<td colspan='2' valign="top"><%=editAddress2Name%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">City, State, Zip:&nbsp;</td>
		<td colspan='2' valign="top"><%=editCityName%>, <%=editStateName%>, <%=editZipName%></td>
	</tr>
	<tr>
		<td width='25' class='bar' valign="top"><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center' valign="top">Seminar Information</td>
		<td width='100%' class='bar' valign="top"><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' nowrap align='right' valign="top">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Seminar Name, Date, City:&nbsp;</td>
		<td colspan='2' valign="top">
			<!--<%=editSeminarName%>, (<%=editSeminarDate%>), <%=editSeminarCity%></td>-->
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Seminar Cost:&nbsp;</td>
		<td colspan='2' valign="top">$<%=editSeminarCost%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Payment Type:&nbsp;</td>
		<td colspan='2' valign="top"><%=editPaymentType%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Discount?:&nbsp;</td>
		<td colspan='2' valign="top"><%=editDiscount%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Paid?:&nbsp;</td>
		<td colspan='2' valign="top"><%=editPaid%></td>
	</tr>
	<tr>
		<td width='25' class='bar' valign="top"><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center' valign="top">Miscellaneous</td>
		<td width='100%' class='bar' valign="top"><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Retail1:&nbsp;</td>
		<td colspan='2' valign="top"><%=editRetail1%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Retail2:&nbsp;</td>
		<td colspan='2' valign="top"><%=editRetail2%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Retail3:&nbsp;</td>
		<td colspan='2' valign="top"><%=editRetail3%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Comments1:&nbsp;</td>
		<td colspan='2' valign="top"><%=editComments1%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Comments2:&nbsp;</td>
		<td colspan='2' valign="top"><%=editComments2%></td>
	</tr>
	<tr>
		<td colspan='2' align='right' valign="top">Comments3:&nbsp;</td>
		<td colspan='2' valign="top"><%=editComments3%></td>
	</tr>
	<tr>
		<td width='25' class='bar'><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' align='center' class='bartitle' nowrap>
			&nbsp;<% If (getUserInfo(getUserID,"accesslevel") = "Admin") Then %><input type='submit' name='edit' value='Edit'><% End If %>
		</td>
		<td class='bar'>
			<img src='blank.gif' width=25 height=1>
			<input type=submit name=first value='<<'>
			<input type=submit name=prev value='<'>
			<input name='editID' size='5' value='<%=editID%>' style="text-align:center;">
			<input type=submit name=goto value='Go'>
<!--			<input name='editPos' size='6' value='<%=editPos%>' style="text-align:center;"> -->
			<input type=submit name=next value='>'>
			<input type=submit name=last value='>>'>
		</td>
	</tr>
</table>
</FORM>
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