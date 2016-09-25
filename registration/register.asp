<% authNotNeeded = True %>
<!--#include file="admin/functions.asp"-->
<%
'If (Request.ServerVariables("REMOTE_ADDR") <> "205.208.240.201") Then
'	Response.Write("This section is down for maintenance, please check back shortly.")
'	Response.End
'End If

Function RenderCountrySelect(strFormName)
	%><!--#include virtual="/registration/rendercountryselect.asp"--><%
End Function

'If (Request.QueryString("tour") = "Sony Media Software Showcase") Then
'	Response.Redirect "/registration/?sony"
'End If
If Not (Request.QueryString("tour") = "") Then

errStartFormat = "<BR><FONT COLOR=red><small><B><I>"
errEndFormat = "</B></I></small></FONT>"

'Response.Write(Request.Form)

isFormGood = True

'Remove spaces and setup variables for checks.
scrollTop = Request.Form("scrollTop")
firstName = Trim(Request.Form("firstName"))
lastName = Trim(Request.Form("lastName"))
companyName = Trim(Request.Form("companyName"))
companyTitleName = Trim(Request.Form("companyTitleName"))
address1Name = Trim(Request.Form("address1Name"))
address2Name = Trim(Request.Form("address2Name"))
cityName = Trim(Request.Form("cityName"))
stateName = Trim(Request.Form("stateName"))
zipName = Trim(Request.Form("zipName"))
countryName = Trim(Request.Form("countryName"))
phoneName = Trim(Request.Form("phoneName"))
emailName = Trim(Request.Form("emailName"))
'optInEmails = Trim(Request.Form("optInEmails"))
'optInChecked = "CHECKED"
discountCode = Ucase(Trim(Request.Form("discountCode")))
seminarCost = getSeminarCost(Request.QueryString("tour"),discountCode)
normalCost = getSeminarCost(Request.QueryString("tour"),"This discount code will obviously not work, so it will return the normal cost of this tour.")
If (seminarCost < normalCost) Then
	wasDiscounted = True
Else
	wasDiscounted = False
End If
If (seminarCost = "n/a") Then
	Response.Redirect "index.asp?noPrice=true"
'	Response.Write("Discount Code: " & discountCode & "<br />")
'	Response.Write("Seminar Cost: " & seminarCost & "<br />")
'	Response.Write("Normal Cost: " & normalCost & "<br />")
'	Response.End()
End If
paymentType = Request.Form("paymentType")
selectSeminar = Request.Form("selectSeminar")
comments1 = Request.Form("comments1")
hearAbout = Request.Form("hearAbout")
hearAboutOther = Trim(Request.Form("hearAboutOther"))

Function writeFocus(focusField)
	Response.Write("<script>" & vbNewline)
	Response.Write("function focusCode() {" & vbNewline)
	Response.Write("  document.forms.regform." & focusField & ".focus();" & vbNewline)
	If (isNumeric(scrollTop)) And (Request.Form("method") = "checkCode") Then
		Response.Write("  window.scrollTo(0," & scrollTop & ");" & vbNewline)
	End If
	Response.Write("}" & vbNewline)
	Response.Write("window.onload = focusCode;" & vbNewline)
	Response.WritE("</script>")
End Function

If (Request.Form("method") = "processRegistration") Or (Request.Form("method") = "checkCode") Then
	If (Request.Form("method") = "checkCode") Then
		writeFocus "discountCode"
		isFormGood = False
	Else
		writeFocus "firstName"
	End If

	'Check first name and report error if field is blank.
	If (firstName = "") Then
		isFormGood = False
		errFirstName = errStartFormat & "Your first name is required." & errEndFormat
	End If

	'Check last name and report error if field is blank.
	If (lastName = "") Then
		isFormGood = False
		errLastName = errStartFormat & "Your last name is required." & errEndFormat
	End If

	'Check address and report error if field is blank.
	If (address1Name = "") Then
		isFormGood = False
		errAddress1Name = errStartFormat & "Your address is required." & errEndFormat
	End If

	'Check city  and report error if field is blank.
	If (cityName = "") Then
		isFormGood = False
		errCityName = errStartFormat & "Your city is required." & errEndFormat
	End If

	'Check zip and report error if field is blank.
	If (countryName = "United States of America") Then
		'Check state and report error if field is blank.
		If (stateName = "") Then
			isFormGood = False
			errStateName = errStartFormat & "Your state is required." & errEndFormat
		End If
		If (Len(zipName) = 0) Then
			isFormGood = False
			errZipName = errStartFormat & "Your zip code is required." & errEndFormat
		ElseIf (Len(zipName) < 5) Then
			isFormGood = False
			errZipName = errStartFormat & "A 5 digit zip code is required." & errEndFormat
		End If
	ElseIf (countryName = "Canada") Then
		If (stateName = "") Then
			isFormGood = False
			errStateName = errStartFormat & "Your province is required." & errEndFormat
		End If
		If (Len(zipName) = 0) Then
			isFormGood = False
			errZipName = errStartFormat & "Your postal code is required." & errEndFormat
		ElseIf (Len(zipName) < 7) Then
			isFormGood = False
			errZipName = errStartFormat & "Your postal code should be formated 'A1A 1A1'." & errEndFormat
		End If
	End If
	
	If (Len(countryName) = 0) Then
		isFormGood = False
		errCountryName = errStartFormat & "You must select your country." & errEndFormat
	End If

	'Check phone and report error if field is blank.
	If (Len(phoneName) = 0) Then
		isFormGood = False
		errPhoneName = errStartFormat & "Your phone number is required." & errEndFormat
	ElseIf (Len(phoneName) < 10) Then
		isFormGood = False
		errPhoneName = errStartFormat & "Your phone number must be at least 10 digits long, please include area code and country code if outside US & Canada." & errEndFormat
	ElseIf Not (isNumeric(phoneName)) Then
		isFormGood = False
		errPhoneName = errStartFormat & "Your phone number should be a numeric number, please no symbols." & errEndFormat
	End If

	'Check email address and report error if field is blank.
	If (emailName = "") Then
		isFormGood = False
		errEmailName = errStartFormat & "Your email address is required." & errEndFormat
	Else
		If Not (checkAddress(emailName)) Then
			isFormGood = False
			errEmailName = errStartFormat & "Email address needs to be in this format: john@abc.com" & errEndFormat
		End If
	End If

	Function checkAddress(emailAddress)
		If (emailAddress = "") Then
			checkAddress = false
			exit function
		Else
			addressOk = true
			For X = 1 To Len(emailAddress)
				letter = mid(emailAddress, X, 1)
				If Not ((letter >= "0") And (letter <= "9") Or (letter >= "A") And (letter <= "Z") Or (letter >= "a") And (letter <= "z") Or (letter = "-") Or (letter = "_") Or (letter = "@") Or (letter = ".")) Then
					checkAddress = false
					exit function
				End If
			Next
			whereIsAt = Cint(InStr(emailAddress,"@"))

			If (whereIsAt > 0) Then
				leftOfAddress = Left(emailAddress,whereIsAt-1)
				leftLength = Cint(Len(leftOfAddress))

				rightOfAddress = Right(emailAddress,Len(emailAddress)-whereIsAt)
				rightLength = Cint(Len(rightOfAddress))

				firstLeftPeriod = Cint(InStr(leftOfAddress,"."))
				lastLeftPeriod = Cint(InStrRev(leftOfAddress,"."))
				firstRightPeriod = Cint(InStr(rightOfAddress,"."))
				lastRightPeriod = Cint(InStrRev(rightOfAddress,"."))
			End If

			'These are all the checks used to make sure the email address is valid.
			If (whereIsAt = 0) Then
				checkAddress = false
				exit function
			End If

			If (leftLength < 1) Or _
			(rightLength < 3) Then
				checkAddress = false
				exit function
			End If

			If (firstLeftPeriod = 1) Or _
			(lastLeftPeriod = leftLength) Or _
			(firstRightPeriod = 1) Or _
			(lastRightPeriod = rightLength) Or _
			(firstRightPeriod = 0) And (lastRightPeriod = 0) Then
				checkAddress = false
				exit function
			End If

			checkAddress = true
			exit function
		End If
	End Function

'	If (optInEmails = "yes") Then
'		optInChecked = "CHECKED"
'	Else
'		optInChecked = ""
'	End If

	'Check payment type to make sure it is valid and understood.
	If (seminarCost > 0) Then
		If (paymentType = "Paypal") Then
			paypalChecked = "CHECKED"
		ElseIF (paymentType = "Check") Then
			checkChecked = "CHECKED"
		ElseIf (paymentType = "Credit") Then
			creditChecked = "CHECKED"
		Else
			isFormGood = False
			errPaymentType = errStartFormat & "Invalid payment type, please select another type." & errEndFormat
		End If
	Else
		paymentType = "free"
	End If
	
	'Check to make sure its a valid seminar selected.
	If (selectSeminar = "") Then
		isFormGood = False
		errSelectSeminar = Replace(errStartFormat,"<BR>","") & "A seminar is required." & errEndFormat
	End If

	If (hearAbout = "select") Then
		isFormGood = False
		errSelectHearAbout = errStartFormat & "We want to know how you heard about us." & errEndFormat
	End If

	If (hearAbout = "Other") And (hearAboutOther = "") Then
		isFormGood = False
		errSelectHearAboutOther = errStartFormat & "Please enter a response." & errEndFormat
	End If

	'If we passed all the checks, process the form now and enter into database.
	If (isFormGood) Then
		If (hearAbout = "Other") Then
			hearString = "Other - " & hearAboutOther
		Else
			hearString = hearAbout
		End If

		'Subscribe to opt-in mailing.
		'If (optInEmails = "yes") Then
			subscribeName = Replace(Trim(firstName & " " & lastName),"'","&#39;")
			subscribeEmail = Replace(Trim(emailName),"'","&#39;")
			subscribeToList 26, subscribeName, emailName
		'End If

		'Add the customer to the database.
		custID = registerCustomer(Replace(firstName,"'","&#39;"), Replace(lastName,"'","&#39;"), Replace(companyName,"'","&#39;"), Replace(companyTitleName,"'","&#39;"), Replace(address1Name,"'","&#39;"), Replace(address2Name,"'","&#39;"), Replace(cityName,"'","&#39;"), Replace(stateName,"'","&#39;"), Replace(zipName,"'","&#39;"), Replace(phoneName,"'","&#39;"), Replace(emailName,"'","&#39;"), seminarCost, Replace(paymentType,"'","&#39;"), Replace(selectSeminar,"'","&#39;"), Replace(comments1,"'","&#39;"), Replace(hearString,"'","&#39;"), wasDiscounted, Replace(countryName,"'","&#39;"))

		'Send the customer a email confirming their signup to the email address they entered, and send a notification to register@sundancemediagroup.com of the customers registration.
		Call sendConfirmationOfSignup(custID)

		If (seminarCost > 0) Then
			If (paymentType = "Paypal") Then
				'Redirect for PayPal payment.
				Response.Cookies("verisign")("custID") = custID
				Response.Redirect("paypal.asp")
			ElseIF (paymentType = "Check") Then
				'Redirect for Check payment information.
				Response.Redirect("confirm.asp")
			ElseIf (paymentType = "Credit") Then
				'Redirect to Verisign for Secure Credit Card Transaction.
				Response.Cookies("verisign")("custID") = custID
				Response.Redirect("verisign.asp")
			End If
		Else
			markAsPaid custID
			Response.Redirect("free.asp")
		End If
	End If
	submitted = true
End If

If Not (submitted) Or Not (isFormGood) Then
%>
<html>
	<head>
		<title>Seminar Registration Form</title>
        <style>
			TD { Color: #800000; Font-Size: 12pt; }
			.notes { Font-Size: 10pt; }
		</style>
		<script>
		  function updatePrice() {
		    document.forms.regform.method.value = 'checkCode';
		    document.forms.regform.scrollTop.value = document.body.scrollTop;
		    document.forms.regform.submit();
		  }
		</script>
	<meta name="Microsoft Theme" content="modified-powerplugs-web-templates-art3dblue 011, default">
</head>
	<body background="../_themes/modified-powerplugs-web-templates-art3dblue/background.gif" bgcolor="#C0C0C0" text="#000000" link="#6A6A6A" vlink="#808080" alink="#FFFFFF"><!--mstheme--><font face="Arial, Arial, Helvetica">
        <form method="POST" name="regform" id="regform">
<!--#include file="top.html"-->
        	<input type=hidden name=method value=processRegistration>
        	<input type=hidden name=scrollTop value="<%=scrollTop%>">
<!--mstheme--></font><table border="0" cellpadding="3" cellspacing="0" width="100%%">
    			<tr>
      				<td width="100%" colspan="2" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica">
      					<b><font face="Arial" color="#800000">V.A.S.S.T.
                        Registration Form<br>
                        </font></b>
      					<span class=notes>&nbsp;&quot;*&quot; indicates required information</span>
                    <!--mstheme--></font></td>
				</tr>
	    		<tr>
      				<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>First Name:*</b><!--mstheme--></font></td>
      				<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="firstName" value="<%=firstName%>" size="40"> <%=errFirstName%><!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Last Name:*</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="lastName" value="<%=lastName%>" size="40"> <%=errLastName%><!--mstheme--></font></td>
				</tr>
				<tr>
					<td colspan="2"><!--mstheme--><font face="Arial, Arial, Helvetica">&nbsp;<!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Company Name:</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="companyName" value="<%=companyName%>" size="40"><!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Position/Title:</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="companyTitleName" value="<%=companyTitleName%>" size="40"><!--mstheme--></font></td>
				</tr>
				<tr>
					<td colspan="2"><!--mstheme--><font face="Arial, Arial, Helvetica">&nbsp;<!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Address 1:*</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="address1Name" value="<%=address1Name%>" size="40" maxlength="50"> <%=errAddress1Name%><!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Address 2:<!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="address2Name" value="<%=address2Name%>" size="40"><!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>City:*</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="cityName" value="<%=cityName%>" size="40" maxlength="25"> <%=errCityName%><!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>State / Province:*</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="stateName" value="<%=stateName%>" size="20" maxlength="20"> <small><sup>(US & CA Only)</sup></small> <%=errStateName%><!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Zip / Postal Code:*</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="zipName" value="<%=zipName%>" size="7" maxlength="7"> <small><sup>(US & CA Only)</sup></small> <%=errZipName%><!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Country:*</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><% RenderCountrySelect "countryName" %> <%=errCountryName%><!--mstheme--></font></td>
				</tr>
				<tr>
					<td colspan="2"><!--mstheme--><font face="Arial, Arial, Helvetica">&nbsp;<!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Phone:*</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="phoneName" value="<%=phoneName%>" size="40" maxlength="40"> <%=errPhoneName%><!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Email Address:*</b><!--mstheme--></font></td>
					<td width="60%"><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="emailName" value="<%=emailName%>" size="20" maxlength="40"> <%=errEmailName%><br><!--mstheme--></font></td>
				</tr>
				<tr>
					<td colspan="2"><!--mstheme--><font face="Arial, Arial, Helvetica">&nbsp;<!--mstheme--></font></td>
				</tr>
<% If (seminarCost > 0) Then %>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Payment Amount:</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica">$<%=seminarCost%><!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Discount Code:</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" name="discountCode" value="<%=discountCode%>" size="20" maxlength="20">&nbsp;<a href="javascript:updatePrice();" style="font-size: 10pt;">Update Price</a><!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="middle" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica">&nbsp;<!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><small><small><small>(Only a correct code will change price.)</small></small></small><!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="top" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Payment Type:*</b><!--mstheme--></font></td>
					<td width="60%"><!--mstheme--><font face="Arial, Arial, Helvetica">
						<dl>
							<dt><input type="radio" name="paymentType" value="Credit" <%=creditChecked%>>Credit Card (online)</dt>
							<dt><input type="radio" name="paymentType" value="Paypal" <%=payPalChecked%>>PayPal (online)</dt>
							<dt><input type="radio" name="paymentType" value="Check" <%=checkChecked%>>Mail-in Check</dt>
						</dl><%=errPaymentType%>
					<!--mstheme--></font></td>
				</tr>
				<tr>
					<td colspan="2"><!--mstheme--><font face="Arial, Arial, Helvetica">&nbsp;<!--mstheme--></font></td>
				</tr>
<% End If %>
				<tr>
					<td width="40%" align="right" valign="top" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Seminar Location/Date:*</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica">
						<!--mstheme--></font><table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr bgcolor="#666666">
								<td><!--mstheme--><font face="Arial, Arial, Helvetica"><font color="#FFFFFF">&nbsp;</font><!--mstheme--></font></td>
								<td><!--mstheme--><font face="Arial, Arial, Helvetica"><font color="#FFFFFF"><b>City</b></font><!--mstheme--></font></td>
								<td><!--mstheme--><font face="Arial, Arial, Helvetica"><font color="#FFFFFF"><b>Dates</b></font><!--mstheme--></font></td>
							</tr>
						<!--<select size="1" name="selectSeminar">
							<option value="select">Select a seminar you would like to attend...</option>-->
							<% printSeminarOptions2 selectSeminar, Request.QueryString("tour"), "selectSeminar" %>
						</table><!--mstheme--><font face="Arial, Arial, Helvetica">
						<%=errSelectSeminar%>
						<!--</select>-->
                    <!--mstheme--></font></td>
				</tr>
				<tr>
					<td colspan="2"><!--mstheme--><font face="Arial, Arial, Helvetica">&nbsp;<!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="top" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Where did you hear about us?*</b><br><i>Other:</i><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica">
						<select size="1" name="hearAbout">
							<option value="select">Select where you heard about VASST...</option>
							<% printHearAboutOptions hearAbout,Request.QueryString("tour") %>
							<option value="Other" <% If hearAbout = "Other" Then Response.Write("selected") %>>Other... (Enter on line directly below)</option>
						</select><%=errSelectHearAbout%><br>
						<input type="text" name="hearAboutOther" value="<%=hearAboutOther%>" size="40" maxlength="255"><%=errSelectHearAboutOther%>
                    <!--mstheme--></font></td>
				</tr>
				<tr>
					<td colspan="2"><!--mstheme--><font face="Arial, Arial, Helvetica">&nbsp;<!--mstheme--></font></td>
				</tr>
				<tr>
					<td width="40%" align="right" valign="top" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Comments/Questions:</b><!--mstheme--></font></td>
					<td width="60%" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica"><textarea rows="4" name="comments1" cols="34"><%=comments1%></textarea><!--mstheme--></font></td>
				</tr>
			</table><!--mstheme--><font face="Arial, Arial, Helvetica"><!--mstheme--></font><table border=0 cellpadding=0 cellspacing=0 align="center" width="500">
				<tr>
					<td align="center"><!--mstheme--><font face="Arial, Arial, Helvetica">
						<input type="submit" value="Submit"><input type="reset" value="Reset">
                    <!--mstheme--></font></td>
				</tr>
			</table><!--mstheme--><font face="Arial, Arial, Helvetica">
		</form>
<!--#include file="bottom.html"-->
<%
End If
Else
	Response.Redirect "index.asp"
End If
%> <!--mstheme--></font></body>
</html>