<!--#include file="functions.asp"-->
<% 
If Not (getUserInfo(getUserID,"accesslevel") = "Admin") Then 
	Response.Write "<CENTER><B>You do not have access to this section of the CMS, your access is Read-Only, contact your admin to get rights.</B></CENTER>"
	Response.End
End If
%>
<% printHeader %>
<FORM METHOD=post>
<%
'SplitForm
If (Request.Form("add") = "") Then
	printCustomers ""
Else
	formOkay = True

	If Not isDate(Request.Form("editSignupDate")) Then
		formOkay = False
		errorSignupDate = "CLASS='error'"
	End If

	If (Request.Form("editFirstName") = "") Or (Request.Form("editLastName") = "") Then
		formOkay = False
		errorName = "class='error'"
	End If

	If (Request.Form("editAddress1Name") = "") Then
		formOkay = False
		errorAddress1Name = "CLASS='error'"
	End If

	If (Request.Form("editCityName") = "") Then
		formOkay = False
		errorCityName = "CLASS='error'"
	End If

	If (Request.Form("editStateName") = "") Then
		formOkay = False
		errorStateName = "CLASS='error'"
	End If

	If (Request.Form("editZipName") = "") Then
		formOkay = False
		errorZipName = "CLASS='error'"
	End If

	If (Request.Form("editPhoneName") = "") Then
		formOkay = False
		errorPhoneName = "CLASS='error'"
	End If

	If (Request.Form("editEmailName") = "") Then
		formOkay = False
		errorEmailName = "CLASS='error'"
	End If

	If (Request.Form("editEmailName") = "") Then
		formOkay = False
		errorEmailName = "CLASS='error'"
	End If

	If (Request.Form("editEmailName") = "") Then
		formOkay = False
		errorEmailName = "CLASS='error'"
	End If

	If (Request.Form("seminarSelectMethod") = "manual") And ((Request.Form("editSeminarName") = "") Or (Request.Form("editSeminarDate") = "") Or (Request.Form("editSeminarCity") = "")) Then
		formOkay = False
		errorSeminarSelect = "CLASS='error'"
	End If

	If (Request.Form("seminarSelectMethod") = "list") And (Request.Form("seminarSelect") = "select") Then
		formOkay = False
		errorSeminarSelect = "CLASS='error'"
	End If

	If (Request.Form("seminarSelectMethod") = "") Then
		formOkay = False
		errorSeminarSelect = "CLASS='error'"
	End If

	If (formOkay) Then
		modifyCustomer "add"
		Response.Cookies("customerPass")("editID") = Request.Form("editCustID")
		Response.Redirect("customer_list.asp")
	End If
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
		<td width='50' class='bar'><img src='blank.gif' width=50 height=1></td>
		<td width='200' colspan='2' class='bartitle' nowrap align='center'>Account Information</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Database ID#:</td>
		<td colspan='2'>&nbsp;<input name='editCustID' size='20' value='<%=maxRecord+1%>'>&nbsp; 
    	    Account Status:
	        <select name="editStatus" size='1'>
				<%
				statuses = Array("New Signup","Waiting For Payment","Paid","Canceled","Denied")
				For stat = 0 to Ubound(statuses)
					If (Request.Form("editStatus") = statuses(stat)) Then
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
		<td colspan='2' align='right' <%=errorSignupDate%>>Signed Up Date:</td>
		<td colspan='2'>&nbsp;<input type="text" name="editSignupDate" size="20" value='<%=date & " " & time%>'>&nbsp;
        	<%
			If (Request.Form("editDeleted") = "ON") Then
				deletedChecked = "CHECKED"
			Else
				deletedChecked = ""
			End If
			%>Deleted?: 
        <INPUT CLASS='normal' TYPE='checkbox' NAME='editDeleted' value="ON" <%=deletedChecked%> onClick="deletedToggle()">
        </td>
	</tr>
</table>
<table border='0' cellpadding='0' cellspacing='0' width='100%'>
	<tr>
		<td width='50' class='bar'><img src='blank.gif' width=50 height=1></td>
		<td width='200' colspan='2' class='bartitle' nowrap align='center'>Personal Information</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right' <%=errorName%>>Name:</td>
		<td colspan='2'>&nbsp;<input name='editFirstName' size='28' value='<%=Request.Form("editFirstName")%>'><IMG SRC='blank.gif' WIDTH=5 HEIGHT=1><input name='editLastName' size='27' value='<%=Request.Form("editLastName")%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Company:</td>
		<td colspan='2'>&nbsp;<input name='editCompanyName' size='60' value='<%=Request.Form("editCompanyName")%>'></td>
	</tr>
    <tr>
		<td colspan='2' align='right'>Title:</td>
		<td colspan='2'>&nbsp;<input name='editTitleName' size='60' value='<%=Request.Form("editTitleName")%>'></td>
	</tr>
</table>
<table border='0' cellpadding='0' cellspacing='0' width='100%'>
	<tr>
		<td width='50' class='bar'><img src='blank.gif' width=50 height=1></td>
		<td width='200' colspan='2' class='bartitle' nowrap align='center'>Contact Information</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right' <%=errorPhoneName%>>Phone:</td>
		<td colspan='2'>&nbsp;<input name='editPhoneName' size='20' value='<%=Request.Form("editPhoneName")%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' <%=errorEmailName%>>Email Address:</td>
		<td colspan='2'>&nbsp;<input name='editEmailName' size='60' value='<%=Request.Form("editEmailName")%>'></td>
	</tr>
</table>
<table border='0' cellpadding='0' cellspacing='0' width='100%'>
	<tr>
		<td width='50' class='bar'><img src='blank.gif' width=50 height=1></td>
		<td width='200' colspan='2' class='bartitle' nowrap align='center'>Billing Address</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right' <%=errorAddress1Name%>>Address1:</td>
		<td colspan='2'>&nbsp;<input name='editAddress1Name' size='60' value='<%=Request.Form("editAddress1Name")%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Address2:</td>
		<td colspan='2'>&nbsp;<input name='editAddress2Name' size='60' value='<%=Request.Form("editAddress2Name")%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'><span <%=errorCityName%>>City</span>, <span <%=errorStateName%>>State</span>, <span <%=errorZipName%>>Zip</span>:</td>
		<td colspan='2'>&nbsp;<input name='editCityName' size='20' value='<%=Request.Form("editCityName")%>'>,&nbsp;<input name='editStateName' size='10' value='<%=Request.Form("editStateName")%>'>,&nbsp;<input name='editZipName' size='10' value='<%=Request.Form("editZipName")%>'></td>
	</tr>
</table>
<table border='0' cellpadding='0' cellspacing='0' width='100%'>
	<tr>
		<td width='50' class='bar'><img src='blank.gif' width=50 height=1></td>
		<td width='200' colspan='2' class='bartitle' nowrap align='center'>Seminar Information</td>
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
		<td colspan='2' nowrap align='right' <%=errorSeminarSelect%>>&nbsp;&nbsp;&nbsp;Seminar Name/Date/City:&nbsp;<input class='normal' type='radio' name='seminarSelectMethod' value='manual' <%=seminarManualChecked%>></td>
		<td colspan='2'>
			&nbsp;<input name='editSeminarName' size='20' value='<%=Request.Form("editSeminarName")%>' onKeyDown="this.form.seminarSelectMethod[0].checked = true"> /
			<input name='editSeminarDate' size='20' value='<%=Request.Form("editSeminarDate")%>' onKeyDown="this.form.seminarSelectMethod[0].checked = true"> /
			<input name='editSeminarCity' size='20' value='<%=Request.Form("editSeminarCity")%>' onKeyDown="this.form.seminarSelectMethod[0].checked = true">
		</td>
	</tr>
	<tr>
		<td colspan='2' align='right'>&nbsp;<input class='normal' type='radio' name='seminarSelectMethod' value='list' <%=seminarListChecked%>></td>
		<td colspan='2'>
			&nbsp;<select name='seminarSelect' size=1 onChange="this.form.seminarSelectMethod[1].checked = true">
				<option value='select'>Choose a seminar...</option>
				<% printSeminarOptions(Request.Form("seminarSelect")) %>
			</select>
		</td>
	</tr>
	<tr>
		<td colspan='2' align='right' <%=errorSeminarCost%>>Seminar Cost: $</td>
		<td colspan='2'>&nbsp;<input name='editSeminarCost' size='20' value='<% If (Request.Form("editSeminarCost") = "") Then Response.Write(seminarCost) Else Response.Write(Request.Form("editSeminarCost")) End If %>'>&nbsp; 
    	    Payment Type:
	        <select name="editPaymentType" size='1'>
		        <option value="select">Choose a payment type...</option>
				<%
				payments = Array("Credit","PayPal","Check")
				For stat = 0 to Ubound(payments)
					If (Request.Form("editPaymentType") = payments(stat)) Then
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
		If (Request.Form("editDiscount") = "ON") Then
			discountChecked = "CHECKED"
		Else
			discountChecked = ""
		End If
		%>
		<td colspan='2' align='right'>Discount?:</td>
		<td colspan='2'>&nbsp;<INPUT CLASS='normal' TYPE='checkbox' NAME='editDiscount' value="ON" <%=discountChecked%>>&nbsp;
        	<%
			If (Request.Form("editPaid") = "ON") Then
				paidChecked = "CHECKED"
			Else
				paidChecked = ""
			End If
			%>Paid?: 
        <INPUT CLASS='normal' TYPE='checkbox' NAME='editPaid' value="ON" <%=paidChecked%> onClick="paidToggle()">
        </td>
	</tr>
</table>
<table border='0' cellpadding='0' cellspacing='0' width='100%'>
	<tr>
		<td width='50' class='bar'><img src='blank.gif' width=50 height=1></td>
		<td width='200' colspan='2' class='bartitle' nowrap align='center'>Miscellaneous</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Retail1:</td>
		<td colspan='2'>&nbsp;<input name='editRetail1' size='60' value='<%=Request.Form("editRetail1")%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Retail2:</td>
		<td colspan='2'>&nbsp;<input name='editRetail2' size='60' value='<%=Request.Form("editRetail2")%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Retail3:</td>
		<td colspan='2'>&nbsp;<input name='editRetail3' size='60' value='<%=Request.Form("editRetail3")%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Comments1:</td>
		<td colspan='2'>&nbsp;<input name='editComments1' size='60' value='<%=Request.Form("editComments1")%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Comments2:</td>
		<td colspan='2'>&nbsp;<input name='editComments2' size='60' value='<%=Request.Form("editComments2")%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right'>Comments3:</td>
		<td colspan='2'>&nbsp;<input name='editComments3' size='60' value='<%=Request.Form("editComments3")%>'></td>
	</tr>
</table>
<table border='0' cellpadding='0' cellspacing='0' width='100%'>
	<tr>
		<td width='50' class='bar'><img src='blank.gif' width=50 height=1></td>
		<td width='200' colspan='2' class='bartitle' nowrap align='center'>
			<p align='center'>
			<input type='submit' name='add' value='Add'>
			<input type='reset' value='Reset'>
		</td>
		<td class='bar'><img src='blank.gif' width=<%=displayWidth-200-50%> height=1></td>
	</tr>
</table>
</FORM>
<%
printFooter 
%>