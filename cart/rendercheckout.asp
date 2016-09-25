<%
'If Not (Request.ServerVariables("REMOTE_ADDR") = "205.208.240.201") Then
'	Response.Write("Product checkout is temporarily offline, please check back again later.")
'	Exit Function
'End If
	
	If (GetCartCount < 1) Then
		%><center><b>You have no items in your cart.  <br />You cannot checkout unless you have items in your cart.</b></center><br /><%
	Else
		lblContinue = "Continue"
		lblCancel = "Cancel Checkout"
	
		step1Completed = False
		step2Completed = False
		step3Completed = False
		step4Completed = False

		formFailed = False
		
		If (Request.QueryString("debug") = "true") Then
			postBackURL = Request.ServerVariables("URL") & "?function=checkout&debug=true"
		Else
			postBackURL = Request.ServerVariables("URL") & "?function=checkout"
		End If
		
		strFunction = makeSafe(Request.Form("function"))
		strCheckout = makeSafe(Request.Form("checkout"))

		'Billing Information
		checkoutName = makeSafe(Request.Form("checkoutName"))
		checkoutTitle = makeSafe(Request.Form("checkoutTitle"))
		checkoutCompany = makeSafe(Request.Form("checkoutCompany"))
		checkoutEmailAddress = makeSafe(Request.Form("checkoutEmailAddress"))
		checkoutEmailAddress2 = makeSafe(Request.Form("checkoutEmailAddress2"))
		checkoutPhoneNumber = makeSafe(Request.Form("checkoutPhoneNumber"))
		checkoutAddress = makeSafe(Request.Form("checkoutAddress"))
		checkoutAddress2 = makeSafe(Request.Form("checkoutAddress2"))
		checkoutCity = makeSafe(Request.Form("checkoutCity"))
		checkoutState = makeSafe(Request.Form("checkoutState"))
		checkoutZip = makeSafe(Request.Form("checkoutZip"))
		checkoutCountry = makeSafe(Request.Form("checkoutCountry"))

		'Shipping Information
		sameasbilling = makeSafe(Request.Form("sameasbilling"))
		checkoutShippingName = makeSafe(Request.Form("checkoutShippingName"))
		checkoutShippingTitle = makeSafe(Request.Form("checkoutShippingTitle"))
		checkoutShippingCompany = makeSafe(Request.Form("checkoutShippingCompany"))
		checkoutShippingAddress = makeSafe(Request.Form("checkoutShippingAddress"))
		checkoutShippingAddress2 = makeSafe(Request.Form("checkoutShippingAddress2"))
		checkoutShippingCity = makeSafe(Request.Form("checkoutShippingCity"))
		checkoutShippingState = makeSafe(Request.Form("checkoutShippingState"))
		checkoutShippingZip = makeSafe(Request.Form("checkoutShippingZip"))
		checkoutShippingCountry = makeSafe(Request.Form("checkoutShippingCountry"))

		'Step 2
		shippingMethod = makeSafe(Request.Form("shippingMethod"))

		'Step 2
		paymentMethod = makeSafe(Request.Form("paymentMethod"))

		If (GetOrderSessionID > 0) Then
			step1Completed = IsStep1Completed
			step2Completed = IsStep2Completed
			step3Completed = IsStep3Completed
			step4Completed = IsStep4Completed
		End If
		
		If (Session("savestep4") = "yes") Then
			strFunction = lblContinue
			Session("savestep4") = ""
		End If
		
		If (strFunction = lblCancel) Then
			KillOrderSession
			RedirectToCheckout
		ElseIf (strFunction = lblContinue) Then
			If Not (step1Completed) Then
				If (Len(checkoutName) < 4) Then	
					formFailed = True
					checkoutNameError = PrintError("A name is required.")
				End If
		
				If (checkoutEmailAddress = "") Then
					formFailed = True
					checkoutEmailAddressError = PrintError("An email address is required.")
				Else
					If (checkoutEmailAddress2 = "") Then
						formFailed = True
						checkoutEmailAddress2Error = PrintError("A matching email address is required.")
					Else
						If Not (IsValidEmailAddress(checkoutEmailAddress)) Then
							formFailed = True
							checkoutEmailAddressError = PrintError("An email address needs to be in this format: john@domain.com")
						Else
							If Not (IsValidEmailAddress(checkoutEmailAddress)) Then
								formFailed = True
								checkoutEmailAddress2Error = PrintError("An email address needs to be in this format: john@domain.com")
							Else
								If (checkoutEmailAddress <> checkoutEmailAddress2) Then
									formFailed = True
									checkoutEmailAddress2Error = PrintError("Both email address boxes must match.")
								End If
							End If
						End If
					End If
				End If
				
				If (Len(checkoutPhoneNumber) < 7) Then	
					formFailed = True
					checkoutPhoneNumberError = PrintError("A phone number is required.")
				End If
		
				If (Len(checkoutAddress) < 5) Then	
					formFailed = True
					checkoutAddressError = PrintError("An address is required.")
				End If
		
				If (Len(checkoutCity) = 0) Then	
					formFailed = True
					checkoutCityError = PrintError("An city is required.")
				End If
		
				If (checkoutCountry = "United States of America") Then
					If (Len(checkoutState) = 0) Then	
						formFailed = True
						checkoutStateError = PrintError("A state is required.")
					End If
		
					If (Len(checkoutZip) < 5) Then	
						formFailed = True
						checkoutZipError = PrintError("An zip code is required.")
					End If
				ElseIf (checkoutCountry = "Canada") Then
					If (Len(checkoutState) = 0) Then	
						formFailed = True
						checkoutStateError = PrintError("A provence is required.")
					End If
		
					If (Len(checkoutZip) < 6) Then	
						formFailed = True
						checkoutZipError = PrintError("An postal code is required. Ex. 'A1A 1A1'")
					End If
				End If

				If (Len(checkoutCountry) = 0) Then
					formFailed = True
					checkoutCountryError = PrintError("A country selection is requried.")
				End If
		
				If (NeedToShip) And ((isNull(sameasbilling)) Or (Len(sameasbilling) = 0)) Then
					If (Len(checkoutShippingName) < 4) Then	
						formFailed = True
						checkoutShippingNameError = PrintError("A name is required.")
					End If
		
					If (Len(checkoutShippingAddress) < 5) Then	
						formFailed = True
						checkoutShippingAddressError = PrintError("An address is required.")
					End If
			
					If (Len(checkoutShippingCity) = 0) Then	
						formFailed = True
						checkoutShippingCityError = PrintError("An city is required.")
					End If
			
					If (checkoutShippingCountry = "United States of America") Then
						If (Len(checkoutShippingState) = 0) Then	
							formFailed = True
							checkoutShippingStateError = PrintError("A state is required.")
						End If
			
						If (Len(checkoutShippingZip) < 5) Then	
							formFailed = True
							checkoutShippingZipError = PrintError("An zip code is required.")
						End If
					ElseIf (checkoutShippingCountry = "Canada") Then
						If (Len(checkoutShippingState) = 0) Then	
							formFailed = True
							checkoutShippingStateError = PrintError("A provence is required.")
						End If
			
						If (Len(checkoutShippingZip) < 7) Then	
							formFailed = True
							checkoutShippingZipError = PrintError("An postal code is required. Ex. 'A1A 1A1'")
						End If
					End If
	
					If (Len(checkoutShippingCountry) = 0) Then
						formFailed = True
						checkoutShippingCountryError = PrintError("A country selection is requried.")
					End If
				End If
					
				If Not (formFailed) Then
					If (sameasbilling = "yes") Or (Not (NeedToShip)) Then
						SaveStep1 checkoutName, checkoutTitle, checkoutCompany, checkoutEmailAddress, checkoutPhoneNumber, checkoutAddress, checkoutAddress2, checkoutCity, checkoutState, checkoutZip, checkoutCountry, checkoutName, checkoutTitle, checkoutCompany, checkoutAddress, checkoutAddress2, checkoutCity, checkoutState, checkoutZip, checkoutCountry
					Else
						SaveStep1 checkoutName, checkoutTitle, checkoutCompany, checkoutEmailAddress, checkoutPhoneNumber, checkoutAddress, checkoutAddress2, checkoutCity, checkoutState, checkoutZip, checkoutCountry, checkoutShippingName, checkoutShippingTitle, checkoutShippingCompany, checkoutShippingAddress, checkoutShippingAddress2, checkoutShippingCity, checkoutShippingState, checkoutShippingZip, checkoutShippingCountry
					End If
				End If
			ElseIf (step1Completed) And Not (step2Completed) Then
				formFailed = True
				If (shippingMethod = "bypass") And (Request.Form("bypasscode") = "freeship") Then
					formFailed = False
				ElseIf (shippingMethod = "ground") Or (shippingMethod = "2day") Or (shippingMethod = "overnight") Or (shippingMethod = "canadamexico") Or (shippingMethod = "international") Then
					formFailed = False
				End If
				
				If Not (formFailed) Then
					SaveStep2 shippingMethod
				End If
			ElseIf (step1Completed) And (step2Completed) And Not (step3Completed) Then
				formFailed = True
				If (paymentMethod = "") Or (isNull(paymentMethod)) Then
					If (GetTotal("totalcost") = 0) Then
						formFailed = False
						paymentMethod = "exempt"
					Else
						formFailed = True
					End If
				ElseIf (paymentMethod = "check") Or (paymentMethod = "creditdebit") Or (paymentMethod = "paypal") Then
					formFailed = False
				Else
					formFailed = True
				End If

				If Not (formFailed) Then
					SaveStep3 paymentMethod
				End If
			ElseIf (step1Completed) And (step2Completed) And (step3Completed) And Not (step4Completed) Then

				If Not (formFailed) Then
					SaveStep4
				End If
			End If
		End If
		
		step1Image = "step1incomplete.jpg"
		step2Image = "step2incomplete.jpg"
		step3Image = "step3incomplete.jpg"
		step4Image = "step4incomplete.jpg"
		
		If Not (step1Completed) Then
			step1Image = "step1inprogress.jpg"
		Else
			step1Image = "step1complete.jpg"
			If Not (step2Completed) Then
				step2Image = "step2inprogress.jpg"
			Else
				step2Image = "step2complete.jpg"
				If Not (step3Completed) Then
					step3Image = "step3inprogress.jpg"
				Else
					step3Image = "step3complete.jpg"
					If Not (step4Completed) Then
						step4Image = "step4inprogress.jpg"
						If (GetPaymentMethod = "creditdebit") Then 
							postBackURL = "https://payflowlink.verisign.com/payflowlink.cfm"
						ElseIf (GetPaymentMethod = "paypal") Then
							postBackURL = "https://www.paypal.com/cgi-bin/webscr"
						End If
					Else
						step4Image = "step4complete.jpg"
					End If
				End If
			End If
		End If
		
		debug("Step1Complete: " & step1Completed & "<br />")
		debug("Step2Complete: " & step2Completed & "<br />")
		debug("Step3Complete: " & step3Completed & "<br />")
		debug("Step4Complete: " & step4Completed & "<br />")
		
		%>
			<table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse;">
				<form name="checkout" id="checkout" method="post" action="<%=postBackURL%>">
				<tr>
					<td bgcolor="#c0c0c0">
						<table border="0" cellpadding="0" cellspacing="0" align="right">
							<tr bgcolor="#c0c0c0">
								<td><img src="<%=step1Image%>"></td>
								<td><img src="<%=step2Image%>"></td>
								<td><img src="<%=step3Image%>"></td>
								<td><img src="<%=step4Image%>"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table border="0" cellpadding="3" cellspacing="0" width="100%">

<% If Not (step1Completed) Then %>
							<tr>
								<td colspan="2" nowrap background="/images/banner_back.gif"><font style="color: white;"><b>&nbsp;&nbsp;Billing Information</b></font></td>
							</tr>
							<tr>
								<td colspan="2"><font size="-1">An astrisk (*) denotes the field is required.&nbsp;&nbsp;This information needs to be the same as it appears on your billing statement.</font></td>
							</tr>
							<tr>
								<td width="35%" align="right"><font size="-1">Name:</font></td>
								<td valign="top"><input type="text" name="checkoutName" size="30" value="<%=checkoutName%>">&nbsp;*</td>
							</tr>
							<%=checkoutNameError%>
							
							<tr>
								<td align="right"><font size="-1">Title:</font></td>
								<td valign="top"><input type="text" name="checkoutTitle" size="30" value="<%=checkoutTitle%>"></td>
							</tr>
							<%=checkoutTitleError%>
							
							<tr>
								<td align="right"><font size="-1">Company:</font></td>
								<td valign="top"><input type="text" name="checkoutCompany" size="30" value="<%=checkoutCompany%>"></td>
							</tr>
							<%=checkoutCompanyError%>
							
							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>

							<tr>
								<td colspan="2"><font size="-1">Please make sure you type a valid email address.  We use this address to contact you with the invoice and shipping information.</font></td>
							</tr>
							
							<tr>
								<td align="right"><font size="-1">Email Address:</font></td>
								<td valign="top"><input type="text" name="checkoutEmailAddress" size="30" value="<%=checkoutEmailAddress%>">&nbsp;*</td>
							</tr>
							<%=checkoutEmailAddressError%>
							
							<tr>
								<td align="right"><font size="-1">Re-type Email:</font></td>
								<td valign="top"><input type="text" name="checkoutEmailAddress2" size="30" value="<%=checkoutEmailAddress2%>">&nbsp;*</td>
							</tr>
							<%=checkoutEmailAddress2Error%>

							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
						
							<tr>
								<td align="right"><font size="-1">Phone Number:</font></td>
								<td valign="top"><input type="text" name="checkoutPhoneNumber" size="30" value="<%=checkoutPhoneNumber%>">&nbsp;*</td>
							</tr>
							<%=checkoutPhoneNumberError%>
							
							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
							
							<tr>
								<td align="right"><font size="-1">Address:</font></td>
								<td valign="top">
									<input type="text" name="checkoutAddress" size="30" value="<%=checkoutAddress%>">&nbsp;*<br />
									<span style="font-size: 10px;"><i>Street address, P.O. box, company name, c/o</i></span>
								</td>
							</tr>
							<%=checkoutAddressError%>

							<tr>
								<td align="right"><font size="-1">Address 2:</font></td>
								<td valign="top">
									<input type="text" name="checkoutAddress2" size="30" value="<%=checkoutAddress2%>"><br />
									<span style="font-size: 10px;"><i>Apartment, suite, unit, building, floor, etc.</i></span>
								</td>
							</tr>
							
							<tr>
								<td align="right"><font size="-1">City:</font></td>
								<td valign="top"><input type="text" name="checkoutCity" size="30" value="<%=checkoutCity%>">&nbsp;*</td>
							</tr>
							<%=checkoutCityError%>
							
							<tr>
								<td align="right"><font size="-1">State / Provence:</font></td>
								<td valign="top">
									<% RenderStateSelect "checkoutState" %>&nbsp;* <font size="-1">(US/CA Only)</font>
								</td>
							</tr>
							<%=checkoutStateError%>
							
							<tr>
								<td align="right"><font size="-1">Zip / Postal Code:</font></td>
								<td valign="top"><input type="text" name="checkoutZip" size="30" value="<%=checkoutZip%>">&nbsp;* <font size="-1">(US/CA Only)</font></td>
							</tr>
							<%=checkoutZipError%>
							
							<tr>
								<td align="right"><font size="-1">Country:</font></td>
								<td valign="top">
									<% RenderCountrySelect "checkoutCountry" %>&nbsp;*
								</td>
							</tr>
							<%=checkoutCountryError%>
							
							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
						
<% If (NeedToShip) Then %>
							<tr>
								<td colspan="2" nowrap background="/images/banner_back.gif"><font style="color: white;"><b>&nbsp;&nbsp;Shipping Information</b></font></td>
							</tr>
							
							<script>
								function toggleShipping(id) 
								{
									var isChecked = (id.checked) ? true : false;
									document.checkout.checkoutShippingName.disabled = isChecked;
									document.checkout.checkoutShippingTitle.disabled = isChecked;
									document.checkout.checkoutShippingCompany.disabled = isChecked;
									document.checkout.checkoutShippingAddress.disabled = isChecked;
									document.checkout.checkoutShippingAddress2.disabled = isChecked;
									document.checkout.checkoutShippingCity.disabled = isChecked;
									document.checkout.checkoutShippingState.disabled = isChecked;
									document.checkout.checkoutShippingZip.disabled = isChecked;
									document.checkout.checkoutShippingCountry.disabled = isChecked;
									/*
									if (isChecked)
									{
										toggleCountry(document.checkout.checkoutCountry);
									}
									else
									{
										toggleCountry(document.checkout.checkoutShippingCountry);
									}
									*/
								}
								/*
								var selectedMethod = "<%=shippingMethod%>";
								
								var shippingChoose = new Array(
									Array("Please choose a shipping country.",""));
								
								var shippingUSA = new Array(
									Array("Ground","ground"),
									Array("2 Day","2day"),
									Array("Overnight","overnight"));

								var shippingCanada = new Array(
									Array("Canada","canadamexico"));

								var shippingMexico = new Array(
									Array("Mexico","canadamexico"));

								var shippingInternational = new Array(
									Array("International","international"));
								*/
								function toggleCountry(id)
								{
									/*
									var thisID = id.name;
									if ((thisID == "checkoutShippingCountry") || ((thisID == "checkoutCountry") && (document.checkout.sameasbilling.checked)))
									{
										var thisCountry = id.options[id.selectedIndex].value;
										if (thisCountry == "") 
										{
											changeShipping(shippingChoose);
										}
										else if (thisCountry == "United States of America")
										{
											changeShipping(shippingUSA);
										}
										else if (thisCountry == "Canada")
										{
											changeShipping(shippingCanada);
										}
										else if (thisCountry == "Mexico")
										{
											changeShipping(shippingMexico);
										}
										else 
										{
											changeShipping(shippingInternational);
										}
									}
									*/
								}
								
								function changeShipping(newItemArray)
								{
									document.checkout.shippingMethod.length = 0;

									document.checkout.shippingMethod.length = newItemArray.length;
									for (x = 0; x < newItemArray.length; x++)
									{
										document.checkout.shippingMethod.options[x].text = newItemArray[x][0];
										document.checkout.shippingMethod.options[x].value = newItemArray[x][1];
										if (selectedMethod == newItemArray[x][1]) 
										{
											document.checkout.shippingMethod.selectedIndex = x;
										}
									}
								}
								
								function pageInit()
								{
									toggleShipping(document.checkout.sameasbilling);
									/*if (document.checkout.sameasbilling.checked) 
									{
										toggleCountry(document.checkout.checkoutCountry);
									}
									else
									{
										toggleCountry(document.checkout.checkoutShippingCountry);
									}*/
								}
								
								window.onload = pageInit;
							</script>
							<tr>
								<td align="right"><input type="checkbox" name="sameasbilling" id="sameasbilling" value="yes" onclick="toggleShipping(this)" <% If (sameasbilling = "yes") Then print("checked") %>></td>
								<td><label for="sameasbilling"><font size="-1">Shipping information is the same as the billing information.</font></label></td>
							</tr>
							
							<span id="shippingData">
							<tr>
								<td width="20%" align="right"><font size="-1">Name:</font></td>
								<td valign="top" width="80%"><input type="text" name="checkoutShippingName" size="30" value="<%=checkoutShippingName%>">&nbsp;*</td>
							</tr>
							<%=checkoutShippingNameError%>
							
							<tr>
								<td align="right"><font size="-1">Title:</font></td>
								<td valign="top"><input type="text" name="checkoutShippingTitle" size="30" value="<%=checkoutShippingTitle%>"></td>
							</tr>
							<%=checkoutShippingTitleError%>
							
							<tr>
								<td align="right"><font size="-1">Company:</font></td>
								<td valign="top"><input type="text" name="checkoutShippingCompany" size="30" value="<%=checkoutShippingCompany%>"></td>
							</tr>
							<%=checkoutShippingCompanyError%>
							
							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
							
							<tr>
								<td align="right"><font size="-1">Address:</font></td>
								<td valign="top">
									<input type="text" name="checkoutShippingAddress" size="30" value="<%=checkoutShippingAddress%>">&nbsp;*<br />
									<span style="font-size: 10px;"><i>Street address, P.O. box, company name, c/o</i></span>
								</td>
							</tr>
							<%=checkoutShippingAddressError%>

							<tr>
								<td align="right"><font size="-1">Address 2:</font></td>
								<td valign="top">
									<input type="text" name="checkoutShippingAddress2" size="30" value="<%=checkoutShippingAddress2%>">&nbsp;*<br />
									<span style="font-size: 10px;"><i>Apartment, suite, unit, building, floor, etc.</i></span>
								</td>
							</tr>
							
							<tr>
								<td align="right"><font size="-1">City:</font></td>
								<td valign="top"><input type="text" name="checkoutShippingCity" size="30" value="<%=checkoutShippingCity%>">&nbsp;*</td>
							</tr>
							<%=checkoutShippingCityError%>
							
							<tr>
								<td align="right"><font size="-1">State / Province:</font></td>
								<td valign="top">
									<% RenderStateSelect "checkoutShippingState" %>&nbsp;* <font size="-1">(US/CA Only)</font>
								</td>
							</tr>
							<%=checkoutShippingStateError%>
							
							<tr>
								<td align="right"><font size="-1">Zip:</font></td>
								<td valign="top"><input type="text" name="checkoutShippingZip" size="30" value="<%=checkoutShippingZip%>">&nbsp;* <font size="-1">(US/CA Only)</font></td>
							</tr>
							<%=checkoutShippingZipError%>
							
							<tr>
								<td align="right"><font size="-1">Country:</font></td>
								<td valign="top">
									<% RenderCountrySelect "checkoutShippingCountry" %>&nbsp;*
								</td>
							</tr>
							<%=checkoutShippingCountryError%>
							</span>

							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
<% End If %>
<% Else %>
	<% If Not (step2Completed) Then %>							
							<tr>
								<td colspan="2" nowrap background="/images/banner_back.gif"><font style="color: white;"><b>&nbsp;&nbsp;Shipping Options</b></font></td>
							</tr>
							<%
								showUSA = False
								showCanada = False
								showMexico = False
								showInternational = False
								
								shippingCountry = GetShippingCountry
'								Response.Write("Shipping to '" & shippingCountry & "'<br />")
'								If (shippingCountry = "United States of America") Then
'									Response.Write("Matches america.<br />")
'								Else
'									Response.Write("Doesn't match america.<br />")
'								End If
								If (shippingCountry = "United States of America") Or (shippingCountry = "Virgin Islands (USA)") Then
									showUSA = True
								ElseIf (shippingCountry = "Canada") Then
									showCanada = True
								ElseIf (shippingCountry = "Mexico") Then
									showMexico = True
								Else
									showInternational = True
								End If
							%>
							<%
							'<tr>
							'	<td align="right"><font size="-1">Total Products:</font></td>
							'	<td><font size="-1"><%=GetCartCount%'></font></td>
							'</tr>
							%>

							<% If (showUSA) Then %>
							<tr>
								<td align="right" width="25%"><font size="-1">Ground Shipping:</font></td>
								<td width="75%"><font size="-1">$<%=FormatNumber(GetShippingCost("ground"),2)%> (US Only)</font></td>
							</tr>

							<tr>
								<td align="right"><font size="-1">2 Day Shipping:</font></td>
								<td><font size="-1">$<%=FormatNumber(GetShippingCost("twoday"),2)%> (US Only)</font></td>
							</tr>

							<tr>
								<td align="right"><font size="-1">Overnight Shipping:</font></td>
								<td><font size="-1">$<%=FormatNumber(GetShippingCost("overnight"),2)%> (US Only)</font></td>
							</tr>
							<% End If %>

							<% If (showCanada) Or (showMexico) Then %>
							<tr>
								<td align="right"><font size="-1">Canada/Mexico Shipping:</font></td>
								<td><font size="-1">$<%=FormatNumber(GetShippingCost("canadamexico"),2)%> (Canada/Mexico Only)</font></td>
							</tr>
							<% End If %>

							<% If (showInternational) Then %>
							<tr>
								<td align="right"><font size="-1">International Shipping:</font></td>
								<td><font size="-1">$<%=FormatNumber(GetShippingCost("international"),2)%></font></td>
							</tr>
							<% End If %>

							<tr>
								<td align="right"><font size="-1">Shipping Schedule:</font></td>
								<td>
									<small><small>All expedited (overnight, 2nd Day, and Int'l) orders placed before 2pm eastern, will ship the same day. Expedited orders placed after 2pm eastern will ship the following day. All ground orders will ship on Tuesday and Friday of each week. Orders are shipped from our warehousing facility in Orlando, FL.</small></small>
									<!--
									
									<% 'If (showUSA) Then %>
									If ordered by 2 PM EST, and you have chosen 2 day or overnight shipping, your order will be shipped out today.<br /><br />
									<% 'End If %>
									<% 
'										Dim shipArray(3)
'										shipPos = 0
'										If (showUSA) Then
'											shipArray(shipPos) = "Ground"
'											shipPos = shipPos + 1
'										End If
'										If (showCanada) Or (showMexico) Then
'											shipArray(shipPos) = "Canada/Mexico"
'											shipPos = shipPos + 1
'										End If
'										If (showInternational) Then
'											shipArray(shipPos) = "International"
'											shipPos = shipPos + 1
'										End If
'										ReDim Preserve shipArray(shipPos)
'										shipString = Join(shipPos, ", ")
									%>
									<%'=shipString%> orders are shipped on Tuesday and Friday every week.-->
								</td>
							</tr>
							<tr>
								<td align="right"><font size="-1">Shipping Method:</font></td>
								<td>
									<select id="shippingMethod" name="shippingMethod" size="1">
										<% If (showUSA) Then %>
											<option value="ground">Ground</option>
											<option value="2day">2nd Day</option>
											<option value="overnight">Overnight</option>
										<% End If %>
										<% If (showCanada) Then %>
											<option value="canadamexico">Canada</option>
										<% End If %>
										<% If (showMexico) Then %>
											<option value="canadamexico">Mexico</option>
										<% End If %>
										<% If (showInternational) Then %>
											<option value="international">International</option>
										<% End If %>
											<option value="bypass">By-Pass Shipping</option>
									</select>&nbsp;*
								</td>
							</tr>
							<tr>
								<td align="right"><font size="-1">By-Pass Code:</font></td>
								<td>
									<input id="bypasscode" name="bypasscode" type="password" />
								</td>
							</tr>
	<% Else %>
		<% If Not (step3Completed) Then %>
			<tr>
				<td colspan="2">

			<% RenderOrderSummary GetOrderSessionID, False %>
			<br />
			<br />
<% If (GetTotal("totalcost") > 0) Then %>
			<table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse;">
				<tr>
					<td align="center"><b>Select Payment Method</b></td>
				</tr>
				<tr>
					<td>
						<font size="-1">
							Please select the payment method you would like to use:	<br />
							<input type="radio" name="paymentMethod" value="creditdebit" id="creditdebit"><label for="creditdebit">&nbsp;Credit/Debit</label><br />
							<input type="radio" name="paymentMethod" value="paypal" id="paypal"><label for="paypal">&nbsp;PayPal</label><br />
							<input type="radio" name="paymentMethod" value="check" id="check"><label for="check">&nbsp;Check</label><br />
						</font>
					</td>
				</tr>
			</table><br />
			<br />
<% End If %>
			<br />
			<br />

				</td>
			</tr>

		<% Else %>
			<% If Not (step4Completed) Then %>
				<%
					If (GetPaymentMethod = "creditdebit") Then
						Session("cartorder") = "yes"
						Session("returnurl") = Request.ServerVariables("URL")
						%>
							<font color=#800000>
							<form method="POST" action="https://payflowlink.verisign.com/payflowlink.cfm">
								<input type="hidden" name="LOGIN" value="vrn227251412">
								<input type="hidden" name="PARTNER" value="wfb">
								<input type="hidden" name="AMOUNT" value="<%=GetTotal("total")%>">
								<input type="hidden" name="TYPE" value="S">
								<input type="hidden" name="DESCRIPTION" value="Shopping Cart Purchase #C<%=PadNumber(GetInfo("ordersessionid"))%>">
								<input type="hidden" name="NAME" value="<%=GetInfo("name")%>">
								<input type="hidden" name="ADDRESS" value="<%=GetInfo("address")%>">
								<input type="hidden" name="CITY" value="<%=GetInfo("city")%>">
								<input type="hidden" name="STATE" value="<%=GetInfo("state")%>">
								<input type="hidden" name="ZIP" value="<%=GetInfo("zip")%>">
								<input type="hidden" name="COUNTRY" value="<%=GetInfo("country")%>">
								<input type="hidden" name="PHONE" value="<%=ParseInt(GetInfo("phonenumber"))%>">
								<input type="hidden" name="EMAIL" value="<%=GetInfo("emailaddress")%>">
								<input type="hidden" name="INVOICE" value="C<%=PadNumber(GetInfo("ordersessionid"))%>">
								<br />
								<b><%=GetInfo("name")%>, your order is ready to be paid for.</b><br>
								<br>
								Click the button below to continue to the Secure Payment site operated by Verisign.<br>
								<br>
								<center><input type="submit" value="Continue"></center>
								<br />
							</form>
							</font>
						<%						
					ElseIf (GetPaymentMethod = "paypal") Then
						%>
							<font color=#800000>
							<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
								<input type="hidden" name="cmd" value="_xclick">
								<input type="hidden" name="business" value="paypal@sundancemediagroup.com">
								<input type="hidden" name="item_name" value="Shopping Cart Purchase #C<%=PadNumber(GetInfo("ordersessionid"))%>">
								<input type="hidden" name="item_number" value="C<%=PadNumber(GetInfo("ordersessionid"))%>">
								<input type="hidden" name="amount" value="<%=GetTotal("total")%>">
								<input type="hidden" name="no_note" value="1">
								<input type="hidden" name="currency_code" value="USD">
								<b><%=GetInfo("name")%>, your order is ready to be paid for.</b><br>
								<br>
								Click the button below to continue to the Secure Payment site operated by PayPal.<br />
								<br />
								<br>
								<center><input type="submit" value="Continue"></center>
							</form>
							</font>
						<%
					Else
						%>
							<font color=#800000>
								<b>Thank you for your purchase!</b><br>
								<br>
								Please check your email for details about how you can complete the payment process.<br>
								<br>
								Merchandise will not be shipped until we receive the payment, you will receive an email confirmation once we receive your payment.<br />
								<br>
								Kind regards,<br>
								<br>
								VASST<br>
								<br>
								<br>
								<center><input type="submit" name="function" value="Continue"></center>
							</font>
						<%
					End If
				%>
				<b><i>NOTE: Currency is based on US Dollars.</i></b>
				<% 
					SaveStep4 
				%>
			<% Else %>
				<br />
				<br />
				<b><%=GetInfo("name")%>, Your order has been completed.</b><br />
				<br />
				<br />
				<a href="<%=Request.ServerVariables("URL")%>?function=browse">Browse Products</a>
				<% ClearOrder %>
			<% End If %>	
		<% End If %>
	<% End If %>
<% End If %>

<% If Not (step3Completed) Then %>
							<tr>
								<td colspan="2" align="right"><input type="submit" name="function" value="<%=lblContinue%>">&nbsp;&nbsp;<input type="submit" name="function" value="<%=lblCancel%>"></td>
							</tr>
<% Else %>
<% End If %>
							
						</table>
					</td>
				</tr>
				<input type="hidden" name="checkout" value="true">
				</form>
			</table><br />
		<%
	End If
%>
