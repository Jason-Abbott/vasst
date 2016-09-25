<!--#include file="functions.asp"-->
<%
If Not (getUserInfo(getUserID,"accesslevel") = "Admin") Then
	Response.Write "<CENTER><B>You do not have access to this section of the CMS, your access is Read-Only, contact your admin to get rights.</B></CENTER>"
	Response.End
End If
%>
<% printHeader
If Not (dontShowListOrEdit) Then
%>
<FORM METHOD=post>
<%
'Response.Write("Form<BR>" & SplitForm & "Cookies<BR>" & SplitCookies)
If (Request.Form("save") = "") Then
	printCustomers ""
Else
	modifyCustomer "update"
	Response.Cookies("customerPass")("editID") = Request.Form("editID")
	'Response.Redirect("customer_list.asp")
End If
%>
<script>
	var lastStatus;
	function paidToggle() {
		if (document.forms[0].editPaid.checked) {
			lastStatus = document.forms[0].editStatus.value;
			document.forms[0].editStatus.value = "Paid";
		} else {
			document.forms[0].editStatus.value = "Waiting For Payment";
		}
	}
	function deletedToggle() {
		if (document.forms[0].editDeleted.checked) {
			lastStatus = document.forms[0].editStatus.value;
			document.forms[0].editStatus.value = "Canceled";
		} else {
			document.forms[0].editStatus.value = lastStatus;
		}
	}
</script>
<table border='0' cellpadding='0' cellspacing='0' width='100%'>
	<tr>
		<td width='25' class='bar'><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center'>Account Information</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Database ID#:</td>
		<td colspan='2'>&nbsp;<input name='editCustID' size='20' value='<%=editCustID%>'>&nbsp;
    	    Account Status:
	        <select name="editStatus" size='1'>
				<%
				statuses = Array("New Signup","Waiting For Payment","Paid","Canceled","Denied")
				For stat = 0 to Ubound(statuses)
					If (editStatus = statuses(stat)) Then
						statusSelected = "SELECTED"
					Else
						statusSelected = ""
					End If
					Response.Write("<option value='" & statuses(stat) & "' " & statusSelected & ">" & statuses(stat) & "</option>")
				Next
				%>
        	</select>
        </td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Signed Up Date:</td>
		<td colspan='2'>&nbsp;<input type="text" name="editSignupDate" size="20" value="<%=editSignupDate%>">&nbsp;
			<%
			If (editDeleted = "True") Then
				deletedChecked = "CHECKED"
			Else
				deletedChecked = ""
			End If
			%>Deleted?:
        <INPUT CLASS='normal' TYPE='checkbox' NAME='editDeleted' <%=deletedChecked%> onClick="deletedToggle()" value="ON">
        </td>
	</tr>
	<tr>
		<td width='25' class='bar'><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center'>Personal Information</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Name:</td>
		<td colspan='2'>&nbsp;<input name='editFirstName' size='28' value='<%=editFirstName%>'><IMG SRC='blank.gif' WIDTH=5 HEIGHT=1><input name='editLastName' size='27' value='<%=editLastName%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Company:</td>
		<td colspan='2'>&nbsp;<input name='editCompanyName' size='60' value='<%=editCompanyName%>'></td>
	</tr>
    <tr>
		<td colspan='2' align='right'>Title:</td>
		<td colspan='2'>&nbsp;<input name='editTitleName' size='60' value='<%=editTitleName%>'></td>
	</tr>
	<tr>
		<td width='25' class='bar'><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center'>Contact Information</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Phone:</td>
		<td colspan='2'>&nbsp;<input name='editPhoneName' size='20' value='<%=editPhoneName%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Email Address:</td>
		<td colspan='2'>&nbsp;<input name='editEmailName' size='60' value='<%=editEmailName%>'></td>
	</tr>
	<tr>
		<td width='25' class='bar'><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center'>Billing Address</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Address1:</td>
		<td colspan='2'>&nbsp;<input name='editAddress1Name' size='60' value='<%=editAddress1Name%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Address2:</td>
		<td colspan='2'>&nbsp;<input name='editAddress2Name' size='60' value='<%=editAddress2Name%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>City, State, Zip:</td>
		<td colspan='2'>&nbsp;<input name='editCityName' size='20' value='<%=editCityName%>'>,&nbsp;<input name='editStateName' size='10' value='<%=editStateName%>'>,&nbsp;<input name='editZipName' size='10' value='<%=editZipName%>'></td>
	</tr>
	<tr>
		<td width='25' class='bar'><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center'>Seminar Information</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<%
		seminarID = getSeminarID(editSeminarName,editSeminarDate,editSeminarCity)
		If (Cint(seminarID) = -1) Or (seminarID = "") Then
			seminarManualChecked = "CHECKED"
		Else
			seminarListChecked = "CHECKED"
		End If
		%>
		<td colspan='2' nowrap align='right'>&nbsp;&nbsp;&nbsp;Seminar Name/Date/City:&nbsp;<!--<input class='normal' type='radio' name='seminarSelectMethod' value='manual' <%=seminarManualChecked%>>--></td>
		<td colspan='2'>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr bgcolor="#666666">
					<td><font color="#FFFFFF">&nbsp;</font></td>
					<td><font color="#FFFFFF"><b>City</b></font></td>
					<td><font color="#FFFFFF"><b>Dates</b></font></td>
				</tr>
			<!--<select size="1" name="selectSeminar">
				<option value="select">Select a seminar you would like to attend...</option>-->
				<% printSeminarOptions2 seminarID, editSeminarName, "seminarSelect" %>
			</table>
<!--
			&nbsp;<input name='editSeminarName' size='20' value='<%=editSeminarName%>' onKeyDown="this.form.seminarSelectMethod[0].checked = true"> /
			<input name='editSeminarDate' size='20' value='<%=editSeminarDate%>' onKeyDown="this.form.seminarSelectMethod[0].checked = true"> /
			<input name='editSeminarCity' size='20' value='<%=editSeminarCity%>' onKeyDown="this.form.seminarSelectMethod[0].checked = true">
-->
		</td>
	</tr>
<!--
	<tr>
		<td colspan='2' align='right'>&nbsp;<input class='normal' type='radio' name='seminarSelectMethod' value='list' <%=seminarListChecked%>></td>
		<td colspan='2'>
			&nbsp;<select name='seminarSelect' size=1 onChange="this.form.seminarSelectMethod[1].checked = true">
				<option value='select'>Choose a seminar...</option>
				<% printSeminarOptions(seminarID) %>
			</select>
		</td>
	</tr>
-->
	<tr>
		<td colspan='2' align='right'>Seminar Cost: $</td>
		<td colspan='2'>&nbsp;<input name='editSeminarCost' size='20' value='<%=editSeminarCost%>'>&nbsp;
    	    Payment Type:
	        <select name="editPaymentType" size='1'>
		        <option value="select">Choose a payment type...</option>
				<%
				payments = Array("Credit","PayPal","Check")
				For stat = 0 to Ubound(payments)
					If (editPaymentType = payments(stat)) Then
						paymentSelected = "SELECTED"
					Else
						paymentSelected = ""
					End If
					Response.Write("<option value='" & payments(stat) & "' " & paymentSelected & ">" & payments(stat) & "</option>")
				Next
				%>
        	</select>

		</td>
	</tr>
	<tr>
		<%
		If (editDiscount = "True") Then
			discountChecked = "CHECKED"
		Else
			discountChecked = ""
		End If
		%>
		<td colspan='2' align='right'>Discount?:</td>
		<td colspan='2'>&nbsp;<INPUT CLASS='normal' TYPE='checkbox' NAME='editDiscount' <%=discountChecked%> value="ON">&nbsp;
			<%
			If (editPaid = "True") Then
				paidChecked = "CHECKED"
			Else
				paidChecked = ""
			End If
			%>
        	Paid?:
        <INPUT CLASS='normal' TYPE='checkbox' NAME='editPaid' <%=paidChecked%> onClick="paidToggle()" value="ON">
        </td>
	</tr>
	<tr>
		<td width='25' class='bar'><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap align='center'>Miscellaneous</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Retail1:</td>
		<td colspan='2'>&nbsp;<input name='editRetail1' size='60' value='<%=editRetail1%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Retail2:</td>
		<td colspan='2'>&nbsp;<input name='editRetail2' size='60' value='<%=editRetail2%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Retail3:</td>
		<td colspan='2'>&nbsp;<input name='editRetail3' size='60' value='<%=editRetail3%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Comments1:</td>
		<td colspan='2'>&nbsp;<input name='editComments1' size='60' value='<%=editComments1%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Comments2:</td>
		<td colspan='2'>&nbsp;<input name='editComments2' size='60' value='<%=editComments2%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Comments3:</td>
		<td colspan='2'>&nbsp;<input name='editComments3' size='60' value='<%=editComments3%>'></td>
	</tr>
	<tr>
		<td width='25' class='bar'><img src='blank.gif' width=25 height=1></td>
		<td colspan='2' class='bartitle' nowrap>
			<p align='center'>
			<input type='submit' name='save' value='Save/Update'>
			<input type='reset' value='Reset'>
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