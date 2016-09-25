<!--#include virtual="/cart/admin/plugins/regcode.asp"-->
<%
'	continueLoading = True
'	If (Request.ServerVariables("REMOTE_ADDR") = "216.126.195.179") Or _
'		(Request.ServerVariables("REMOTE_ADDR") = "205.208.240.201") Or _
'		(Request.ServerVariables("REMOTE_ADDR") = "12.5.79.2") Then
'		continueLoading = true
'	End If
	
'	If Not (continueLoading) Then
'		Response.Write("The cart is offline temporarily, please check back soon.")
'		Response.End()
'	End If
	
	Dim bolDebug
	Dim strSpacer

	Dim cachedSessionID
	cachedSessionID = -1
	Dim cachedOrderSessionID
	cachedOrderSessionID = -1

	intSessionID = Session.SessionID
	strIPAddress = Request.ServerVariables("REMOTE_ADDR")

	' Set to true to enable debug output, set to false to disable debug output.
	bolDebug = false

	strSpacer = "&nbsp;&nbsp;&nbsp;"

	Dim dbFile
	Dim cartDB
	Dim intOpenCount
	intOpenCount = 0
	
	Randomize

	' Database file location.
	dbFile = "d:\inetpub\vasst_com\cart\admin\data\cartZRQ.mdb"

	' From email address.
	strEmailFrom = "support@vasst.com"
	outgoingMailServer = "mail.sisna.com"

	Function openDB
		If (intOpenCount > 0) Then
			debug("Database was already opened.")
		Else
			debug("Opening connection to database.")
			Set cartDB = Server.CreateObject("ADODB.Connection")
			cartDB.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & dbFile & ";USER ID=Admin;PASSWORD=;"
			debug("Connection to database opened.")
		End If
		intOpenCount = intOpenCount + 1
		debug("Adding Database Count: " & intOpenCount)
	End Function

	Function closeDB
		If (intOpenCount > 1) Then
			debug("Database was open before, we are not going to close it yet.")
		Else
			debug("Closing connection to database.")
			cartDB.Close
			Set cartDB = Nothing
			debug("Connection to database closed.")
		End If
		intOpenCount = intOpenCount - 1
		debug("Removing Database Count: " & intOpenCount)
	End Function
	
	Function cycleDatabase
		If (bolDbOpen) Then
			tmpCount = intOpenCount 
			intOpenCount = 0
			debug("Cycling database.")
			closeDB
			openDB
			intOpenCount = tmpCount
		Else
			debug("Database cycle called when database was closed.")
		End If
	End Function

	Dim globalSQLs : globalSQLs = ""
	Function SQL(strSQL)
		debug("Executing SQL: <font color=""blue"">" & strSQL & "</font>")
		globalSQLs = globalSQLs & "[" & Now & "]: " & strSQL & "<br />"
		SQL = strSQL
	End Function

	Function Message(strMessage)
		debug("Message: <font color=""green"">" & strMessage & "</font>")
		Message = Server.URLEncode(strMessage)
	End Function

	Function printSpacer(intCount)
		For X = 1 To intCount
			printSpacer = printSpacer & strSpacer
		Next
	End Function

	Sub print(strText)
		Response.Write(strText & vbNewline)
	End Sub

	Sub redirect(strURL)
		debug("Redirect has been called, <a href=""" & strURL & """>click here</a> to complete.")
		If Not (bolDebug) Then
			Response.Redirect(strURL)
		End If
	End Sub

	Function PrintError(strText)
		PrintError = "<tr><td></td><td><font size=""-1"" style=""color: red;"">" & strText & "</font></td></tr>"
	End Function

	Sub debug(strText)
		If (bolDebug) Then
			print("DEBUG: " & strText & "<br />")
		End If
	End Sub

	Function makeSafe(strText)
		If (isEmpty(strText)) Or (isNull(strText)) Then
			strText = Null
		Else
			strText = Trim(strText)
			'strText = Replace(strText,"""","""""")
			strText = Replace(strText,"'","&#39;")
		End If
		makeSafe = strText
	End Function

	Function ReportError(strMessage)
		On Error Resume Next
		viewID = GetRandomString
		ReportID = SaveIncident(strMessage, viewID)
		SMTPhtml strEmailFrom, "syntax@sisna.com", "", "", "Automated Error Report Details -- " & viewID & "", strMessage & "<br /><br />" & Replace(GetSessionInformation,vbNewline,"<br />")
		'SMTP strEmailFrom, "8019187449@messaging.sprintpcs.com", "", "", "Automated Error Report", "URL: http://www.vasst.com/cart/admin/incident.asp?view=" & viewID & ""
		'Response.Write("A critical error has occured, the admin has been notified, please try again later.<br />")
		'Response.End
	End Function
	
	Function SaveIncident(strMessage, viewID)
		On Error Resume Next
		openDB 
		cartDB.Execute(SQL("INSERT INTO errors ( reported, message, viewid ) VALUES ( #" & date & " " & time & "#, '" & makeSafe(strMessage & vbNewline & vbNewline & GetSessionInformation) & "', '" & makeSafe(viewID) & "' )"))
		cycleDatabase
		Set dbError = cartDB.Execute(SQL("SELECT id FROM errors WHERE viewid = '" & viewID & "'"))
		SaveIncident = Fix(dbError("id"))
		Set dbError = Nothing
		closeDB
	End Function
	
	Function GetSessionInformation
		strReturn = "<b><u>Session Information</u></b>" & vbNewline & vbNewline
		strReturn = strReturn & "<b>Cart Data</b>" & vbNewline
		strReturn = strReturn & "CartSessionID: " & GetCartSessionID & vbNewline
		strReturn = strReturn & "OrderSessionID: " & GetOrderSessionID & vbNewline

		strReturn = strReturn & vbNewline

		strReturn = strReturn & "<b>POST Data</b>" & vbNewline
		For Each frm In Request.Form
			If (Len(Request.Form(frm)) > 0) Then
				strReturn = strReturn & frm & vbNewline & "------------------" & vbNewline & Request.Form(frm) & vbNewline & vbNewline
			End IF
		Next

		strReturn = strReturn & "<b>GET Data</b>" & vbNewline
		For Each qs In Request.QueryString
			If (Len(Request.QueryString(qs)) > 0) Then
				strReturn = strReturn & qs & vbNewline & "------------------" & vbNewline & Request.QueryString(qs) & vbNewline & vbNewline
			End If
		Next

		strReturn = strReturn & "<b>Environment Data</b>" & vbNewline
		For Each env In Request.ServerVariables
			If (Len(Request.ServerVariables(env)) > 0) Then
				strReturn = strReturn & env & vbNewline & "------------------" & vbNewline & Request.ServerVariables(env) & vbNewline & vbNewline
			End If
		Next
	
		GetSessionInformation = strReturn
	End Function
		
	Function SMTP(mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody)
'		CDOMailer mailFrom, mailTo, mailCC, mailBCC & ";syntax-cart@sisna.com", mailSubject, mailBody, "text"
''		CDOMailer mailFrom, mailTo, mailCC, mailBCC & ";syntax-cart@sisna.com", mailSubject, mailBody, "text"
		CDOMailer mailFrom, mailTo, mailCC, mailBCC & "", mailSubject, mailBody, "text"
'		EasyMail mailFrom, mailTo, mailCC, mailBCC & ";syntax-cart@sisna.com", mailSubject, mailBody, "text"
		SaveMailToHistory mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody
	End Function

	Function SMTPAttachment(mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody, mailAttachment)
		CDOMailerAttachment mailFrom, mailTo, mailCC, mailBCC & "", mailSubject, mailBody, "text", mailAttachment
'		CDOMailerAttachment mailFrom, mailTo, mailCC, mailBCC & ";syntax-cart@sisna.com", mailSubject, mailBody, "text", mailAttachment
'		EasyMailAttachment mailFrom, mailTo, mailCC, mailBCC & ";syntax-cart@sisna.com", mailSubject, mailBody, "text", mailAttachment
		SaveMailToHistory mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody
	End Function

	Function SMTPhtml(mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody)
'		CDOMailer mailFrom, mailTo, mailCC, mailBCC & ";syntax-cart@sisna.com", mailSubject, mailBody, "html"
		CDOMailer mailFrom, mailTo, mailCC, mailBCC & "", mailSubject, mailBody, "html"
		SaveMailToHistory mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody
	End Function
	
	Function CDOMailer(mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody, mailType)
		Set objSendMail = CreateObject("CDO.Message") 
		With objSendMail
			.Subject = mailSubject 
			.From = mailFrom
			.To = Replace(mailTo,";",",")
			.CC = Replace(mailCC,";",",")
			.BCC = Replace(mailBCC,";",",")
			If (mailType = "html") Then
				.HTMLBody = mailBody
			Else
				.TextBody = mailBody
			End If
			.Send()
		End With
		Set objSendMail = Nothing
	End Function
	
	Function CDOMailerAttachment(mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody, mailType, mailAttachment)
		Set objSendMail = CreateObject("CDO.Message") 
		With objSendMail
			.Subject = mailSubject 
			.From = mailFrom
			.To = Replace(mailTo,";",",")
			.CC = Replace(mailCC,";",",")
			.BCC = Replace(mailBCC,";",",")
			.AddAttachment mailAttachment
			If (mailType = "html") Then
				.HTMLBody = mailBody
			Else
				.TextBody = mailBody
			End If
			.Send()
		End With
		Set objSendMail = Nothing
	End Function	
	
	Function SaveMailToHistory(mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody)
		openDB
		cartDB.Execute("INSERT INTO emailhistory ( mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody, mailSent ) VALUES ( '" & makeSafe(mailFrom) & "', '" & makeSafe(mailTo) & "', '" & makeSafe(mailCC) & "', '" & makeSafe(mailBCC) & "', '" & makeSafe(mailSubject) & "', '" & makeSafe(mailBody) & "', #" & date & " " & time & "# )") 
		closeDB
	End Function
	
	Function EasyMail(mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody, mailType)
		aMailTo = Split(mailTo,";")
		aMailCC = Split(mailCC,";")
		aMailBCC = Split(mailBCC,";")
		
		Set ezMailer = Server.CreateObject("EasyMail.SMTP.5")
		ezMailer.LicenseKey = "Unregistered User/S10I510R1AX70C0Rb600"
		ezMailer.MailServer = outgoingMailServer
		ezMailer.FromAddr = mailFrom
		If (mailType = "html") Then
			ezMailer.BodyFormat = 1
		Else
			ezMailer.BodyFormat = 0
		End If
		For Each addr In aMailTo
			If (Len(Trim(addr)) > 0) Then
				ezMailer.AddRecipient "", addr, 1
			End If
		Next
		For Each addr In aMailCC
			If (Len(Trim(addr)) > 0) Then
				ezMailer.AddRecipient "", addr, 2
			End If
		Next
		For Each addr In aMailBCC
			If (Len(Trim(addr)) > 0) Then
				ezMailer.AddRecipient "", addr, 3
			End If
		Next
		ezMailer.Subject = mailSubject
		ezMailer.BodyText = mailBody
		ezMailer.AddCustomHeader "Return-Path", "<support@vasst.com>"
		EasyMail = ezMailer.Send
'		If ezReturn = 0 Then
'		  Response.Write "Message sent successfully."
'		Else
'		  Response.Write "There was an error sending your message.  Error: " & CStr(ezReturn)
'		End If
		Set ezMailer = Nothing
	End Function

	Function EasyMailAttachment(mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody, mailType, mailAttachment)
		aMailTo = Split(mailTo,";")
		aMailCC = Split(mailCC,";")
		aMailBCC = Split(mailBCC,";")
		
		Set ezMailer = Server.CreateObject("EasyMail.SMTP.5")
		ezMailer.LicenseKey = "Unregistered User/S10I510R1AX70C0Rb600"
		ezMailer.MailServer = outgoingMailServer
		ezMailer.FromAddr = mailFrom
		If (mailType = "html") Then
			ezMailer.BodyFormat = 1
		Else
			ezMailer.BodyFormat = 0
		End If
		ezMailer.AddAttachment mailAttachment, 0
		For Each addr In aMailTo
			If (Len(Trim(addr)) > 0) Then
				ezMailer.AddRecipient "", addr, 1
			End If
		Next
		For Each addr In aMailCC
			If (Len(Trim(addr)) > 0) Then
				ezMailer.AddRecipient "", addr, 2
			End If
		Next
		For Each addr In aMailBCC
			If (Len(Trim(addr)) > 0) Then
				ezMailer.AddRecipient "", addr, 3
			End If
		Next
		ezMailer.Subject = mailSubject
		ezMailer.BodyText = mailBody
		ezMailer.AddCustomHeader "Return-Path", "<support@vasst.com>"
		EasyMailAttachment = ezMailer.Send
'		If ezReturn = 0 Then
'		  Response.Write "Message sent successfully."
'		Else
'		  Response.Write "There was an error sending your message.  Error: " & CStr(ezReturn)
'		End If
		Set ezMailer = Nothing
	End Function

	Function IsValidEmailAddress(emailAddress)
		If (emailAddress = "") Or (isNull(emailAddress)) Then
			IsValidEmailAddress = false
			exit function
		Else
			addressOk = true
			For X = 1 To Len(emailAddress)
				letter = mid(emailAddress, X, 1)
				If Not ((letter >= "0") And (letter <= "9") Or (letter >= "A") And (letter <= "Z") Or (letter >= "a") And (letter <= "z") Or (letter = "-") Or (letter = "_") Or (letter = "@") Or (letter = ".")) Then
					IsValidEmailAddress = false
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
				IsValidEmailAddress = false
				exit function
			End If

			If (leftLength < 1) Or _
			(rightLength < 3) Then
				IsValidEmailAddress = false
				exit function
			End If

			If (firstLeftPeriod = 1) Or _
			(lastLeftPeriod = leftLength) Or _
			(firstRightPeriod = 1) Or _
			(lastRightPeriod = rightLength) Or _
			(firstRightPeriod = 0) And (lastRightPeriod = 0) Then
				IsValidEmailAddress = false
				exit function
			End If

			IsValidEmailAddress = true
			exit function
		End If
	End Function

	Function GetRandomString
		Randomize
		strReturn = ""
		For X = 1 To 12
			intRandom = Round(2 * Rnd) + 1
			'Response.Write("Random: " & intRandom & "<br />")
			If (intRandom = 1) Then
				strChar = Chr(Round(9 * Rnd) + 48)
			ElseIf (intRandom = 2) Then
				strChar = Chr(Round(25 * Rnd) + 65)
			ElseIf (intRandom = 3) Then
				strChar = Chr(Round(25 * Rnd) + 97)
			End If
			strReturn = strReturn & strChar
		Next
		GetRandomString = strReturn
	End Function				

	Function PadNumber(intNumber)
		sNumber = CStr(intNumber)
		For X = 0 To 8-Len(sNumber)
			PadNumber = PadNumber & "0"
		Next
		PadNumber = PadNumber & intNumber
	End Function
	
	Function JustifyField(strText,strJustification,strWidth)
		strLeftPad = ""
		strRightPad = ""
		If (Lcase(strJustification) = "right") Then
			For xCount = 0 To strWidth - Len(strText) - 1
				strLeftPad = strLeftPad & " "
			Next
		ElseIf (Lcase(strJustification) = "center") Then
		
		Else
			For xcount = 0 To strWidth - Len(strText) - 1
				strRightPad = strRightPad & " "
			Next
		End If
		JustifyField = strLeftPad & strText & strRightPad
	End Function
	
	Function ReplaceField(strText,strField,strValue)
		If (InStr(strText,strField) > 0) Then
			If (strField = "%INVOICE%") Then
'				Response.WritE("Text: " & strText & ", Field: " & strField & " Value Eval(" & strValue & ")<br />")
				strText = Replace(strText,strField,"C" & eval(strValue))
			Else
				strText = Replace(strText,strField,eval(strValue))
			End If
		End If
		ReplaceField = strText
	End Function
	
	Function ParseAndReplaceEmail(strEmail, intOrderSessionID, intVendorID)
		openDB
		strEmail = ReplaceField(strEmail,"%NAME%","GetInfoFromOrder(""name""," & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%TITLE%","GetInfoFromOrder(""title""," & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%COMPANY%","GetInfoFromOrder(""company""," & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%ADDRESS%","GetInfoFromOrder(""address""," & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%CITY%","GetInfoFromOrder(""city""," & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%STATE%","GetInfoFromOrder(""state""," & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%ZIP%","GetInfoFromOrder(""zip""," & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%PHONENUMBER%","GetInfoFromOrder(""phonenumber""," & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%EMAILADDRESS%","GetInfoFromOrder(""emailaddress""," & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%INVOICE%","PadNumber(GetInfoFromOrder(""ordersessionid""," & intOrderSessionID & "))")
		strEmail = ReplaceField(strEmail,"%BILLINGINFO%","GetBillingInfo(" & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%SHIPPINGINFO%","GetShippingInfo(" & intOrderSessionID & ")")
		
		Set dbOrderCost = cartDB.Execute(SQL("SELECT subtotalcost, shippingcost, salestaxcost, totalcost FROM ordersession WHERE id = " & intOrderSessionID & ""))
		If Not (dbOrderCost.EOF) Then
			fSubTotal = dbOrderCost("subtotalcost")
			fShipping = dbOrderCost("shippingcost")
			fSalesTax = dbOrderCost("salestaxcost")
			fTotal = dbOrderCost("totalcost")
		Else
			ReportError "There was a problem getting the total information from the order to replace fields with."
		End If
		Set dbOrderCost = Nothing

		strEmail = Replace(strEmail,"%SUBTOTAL%",JustifyField(FormatNumber(fSubTotal,2),"right",8))
		strEmail = Replace(strEmail,"%TAX%",JustifyField(FormatNumber(fSalesTax,2),"right",8))
		strEmail = Replace(strEmail,"%SHIPPING%",JustifyField(FormatNumber(fShipping,2),"right",8))
		strEmail = Replace(strEmail,"%TOTAL%",JustifyField(FormatNumber(fTotal,2),"right",8))
		
		strEmail = ReplaceField(strEmail,"%SHIPPINGMETHOD%","GetOrderInfo(""shippingmethod""," & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%PAYMENTMETHOD%","GetOrderInfo(""paymentmethod""," & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%ORDERPRODUCTS%","GetOrderProducts(" & intOrderSessionID & ")")
		strEmail = ReplaceField(strEmail,"%FULFILLMENTPRODUCTS%","GetFulfillmentProducts(" & intOrderSessionID & "," & intVendorID & ")")

		strEmail = ReplaceField(strEmail,"%VENDORNAME%","GetVendorNameFromID(" & intVendorID & ")")

		closeDB

		ParseAndReplaceEmail = strEmail
	End Function
	
	Function GetVendorNameFromID(intVendorID)
		openDB
		Set dbVendor = cartDB.Execute(SQL("SELECT name FROM vendor WHERE id = " & intVendorID & ""))
		If (dbVendor.EOF) Then
			GetVendorNameFromID = "N/A"
		Else
			GetVendorNameFromID = dbVendor("name")
		End If
		closeDB
	End Function
	
	Function GetOrderProducts(intOrderID)
		openDB

		Set dbOrderData = cartDB.Execute(SQL("SELECT productid, pricingid, quantity FROM orderdata WHERE orderdata.ordersessionid = " & intOrderID & ""))
		If (dbOrderData.EOF) Then
			GetOrderProducts = "No products were found in this order."
		Else
			GetOrderProducts = ""
			Do Until dbOrderData.EOF
				Set dbPrice = cartDB.Execute(SQL("SELECT price FROM pricing WHERE id = " & dbOrderData("pricingid") & ""))
				Set dbProduct = cartDB.Execute(SQL("SELECT name FROM product WHERE id = " & dbOrderData("productid") & ""))

				GetOrderProducts = GetOrderProducts & _
					JustifyField(dbOrderData("quantity"),"right",4) & " | " & _
					JustifyField(dbProduct("name") & " @ $" & FormatNumber(dbPrice("price"),2) & "/ea.","left",56) & " | " & _
					"$" & JustifyField(FormatNumber(dbPrice("price")*dbOrderData("quantity"),2),"right",8)

				Set dbProduct = Nothing
				Set dbPrice = Nothing
				
				dbOrderData.MoveNext
			Loop
		End If
		Set dbOrderData = Nothing
		
		closeDB
	End Function
	
	Function GetFulfillmentProducts(intOrderID, intVendorID)
		openDB
		Set dbOrderData = cartDB.Execute(SQL("SELECT productid, pricingid, quantity FROM orderdata WHERE orderdata.ordersessionid = " & intOrderID & ""))
		If (dbOrderData.EOF) Then
			GetFulfillmentProducts = "No proudcts were found in this order, or there was a problem getting the products."
		Else
			GetFulfillmentProducts = ""
			Do Until dbOrderData.EOF
				Set dbProduct = cartDB.Execute(SQL("SELECT name, sku, isbnupc, vendorid FROM product WHERE id = " & dbOrderData("productid") & ""))

				productCount = 0
				If (Fix(dbProduct("vendorid")) = Fix(intVendorID)) Or (Fix(intVendorID) = -2) Then
					productCount = productCount + 1
					GetFulfillmentProducts = GetFulfillmentProducts & _
						"Quantity: " & dbOrderData("quantity") & vbNewline & _
						"Description: " & dbProduct("name") & vbNewline & _
						"SKU: " & dbProduct("sku") & vbNewline & _
						"ISBN/UPC: " & dbProduct("isbnupc") & vbNewline & _
						vbNewline
				End If
				
				Set dbProduct = Nothing
				
				dbOrderData.MoveNext
			Loop
			If (productCount = 0) Then
				GetFulfillmentProducts = "No proudcts were found in this order, or there was a problem getting the products."
			End If
		End If
		Set dbOrderProduct = Nothing
		
		closeDB
	End Function
	
	Function GetOrderInfo(strWhich,intOrderSessionID)
		openDB
		Set dbOrder = cartDB.Execute("SELECT " & strWhich & " FROM ordersession WHERE id = " & intOrderSessionID & "")
		If (dbOrder.EOF) Then
			ReportError "Tried to pull order info for '" & strWhich & "' for order id '" & intOrderSessionID & "' and failed."
		Else
			If (strWhich = "paymentmethod") Then
			End If
			GetOrderInfo = dbOrder(strWhich)
		End If
		closeDB
	End Function

	Function GetProductInfo(strWhich, intProductID)
		openDB
		Set dbProd = cartDB.Execute("SELECT " & strWhich & " FROM product WHERE id = " & intProductID & "")
		If (dbProd.EOF) Then
			ReportError "Tried to pull product info for '" & strWhich & "' for product id '" & intProductID & "' and failed."
		Else
			GetProductInfo = dbProd(strWhich)
		End If
		closeDB
	End Function
	
	Function GetBillingInfo(intOrderSessionID)
		strReturn = GetInfoFromOrder("name",intOrderSessionID)
		If (Len(GetInfoFromOrder("title",intOrderSessionID)) > 0) Then
			strReturn = strReturn & ", " & GetInfoFromOrder("title",intOrderSessionID)
		End If
		strReturn = strReturn & vbNewline
		If (Len(GetInfoFromOrder("company",intOrderSessionID)) > 0) Then
			strReturn = strReturn & GetInfoFromOrder("company",intOrderSessionID) & vbNewline
		End If
		strReturn = strReturn & GetInfoFromOrder("address",intOrderSessionID) & vbNewline
		strReturn = strReturn & GetInfoFromOrder("city",intOrderSessionID) & ", " & GetInfoFromOrder("state",intOrderSessionID) & " " & GetInfoFromOrder("zip",intOrderSessionID) & vbNewline
		strReturn = strReturn & GetInfoFromOrder("country",intOrderSessionID) & vbNewline
		strReturn = strReturn & "Phone: " & GetInfoFromOrder("phonenumber",intOrderSessionID) & vbNewline
		strReturn = strReturn & "Email: " & GetInfoFromOrder("emailaddress",intOrderSessionID) & vbNewline

		GetBillingInfo = strReturn
	End Function

	Function GetShippingInfo(intOrderSessionID)
		strReturn = GetInfoFromShipping("name",intOrderSessionID)
		If (Len(GetInfoFromShipping("title",intOrderSessionID)) > 0) Then
			strReturn = strReturn & ", " & GetInfoFromShipping("title",intOrderSessionID)
		End If
		strReturn = strReturn & vbNewline
		If (Len(GetInfoFromShipping("company",intOrderSessionID)) > 0) Then
			strReturn = strReturn & GetInfoFromShipping("company",intOrderSessionID) & vbNewline
		End If
		strReturn = strReturn & GetInfoFromShipping("address",intOrderSessionID) & vbNewline
		strReturn = strReturn & GetInfoFromShipping("city",intOrderSessionID) & ", " & GetInfoFromShipping("state",intOrderSessionID) & " " & GetInfoFromShipping("zip",intOrderSessionID) & vbNewline
		strReturn = strReturn & GetInfoFromShipping("country",intOrderSessionID) & vbNewline

		GetShippingInfo = strReturn
	End Function

	Function GetInfo(strWhich)
		If (strWhich = "ordersessionid") Then
			strWhich = "ordersession.id"
		End If
		openDB
		Set dbOrderCost = cartDB.Execute(SQL("SELECT " & strWhich & " FROM orderbilling, ordersession WHERE ordersession.billingid = orderbilling.id AND ordersession.id = " & GetOrderSessionID & ""))
		If (strWhich = "ordersession.id") Then
			strWhich = "id"
		End If
		GetInfo = dbOrderCost(strWhich)
		Set dbOrderCost = Nothing
		closeDB
	End Function

	Function GetInfoFromOrder(strWhich,intOrderID)
		If (strWhich = "ordersessionid") Then
			strWhich = "ordersession.id"
		End If
		openDB
		Set dbOrderCost = cartDB.Execute(SQL("SELECT " & strWhich & " FROM orderbilling, ordersession WHERE ordersession.billingid = orderbilling.id AND ordersession.id = " & intOrderID & ""))
		If (strWhich = "ordersession.id") Then
			strWhich = "id"
		End If
		GetInfoFromOrder = dbOrderCost(strWhich)
		Set dbOrderCost = Nothing
		closeDB
	End Function

	Function GetInfoFromShipping(strWhich,intOrderID)
		If (strWhich = "ordersessionid") Then
			strWhich = "ordersession.id"
		End If
		openDB
		Set dbOrderCost = cartDB.Execute(SQL("SELECT " & strWhich & " FROM ordershipping, ordersession WHERE ordersession.shippingid = ordershipping.id AND ordersession.id = " & intOrderID & ""))
		If (strWhich = "ordersession.id") Then
			strWhich = "id"
		End If
		GetInfoFromShipping = dbOrderCost(strWhich)
		Set dbOrderCost = Nothing
		closeDB
	End Function

	Function ParseInt(strText)
		newText = ""
		If (strText <> "") And Not (isNull(strText)) Then
			For X = 1 To Len(strText)
				letter = Right(Left(strText,X),1)
				If ((letter = "0") Or _
					(letter = "1") Or _
					(letter = "2") Or _
					(letter = "3") Or _
					(letter = "4") Or _
					(letter = "5") Or _
					(letter = "6") Or _
					(letter = "7") Or _
					(letter = "8") Or _
					(letter = "-") Or _
					(letter = "9")) Then
					newText = newText & letter
				End If
			Next
		End If
		ParseInt = newText
	End Function

	Function ParseFloat(strText)
		newText = ""
		If (strText <> "") And Not (isNull(strText)) Then
			For X = 1 To Len(strText)
				letter = Right(Left(strText,X),1)
				If ((letter = ".") Or _
					(letter = "0") Or _
					(letter = "1") Or _
					(letter = "2") Or _
					(letter = "3") Or _
					(letter = "4") Or _
					(letter = "5") Or _
					(letter = "6") Or _
					(letter = "7") Or _
					(letter = "8") Or _
					(letter = "9")) Then
					newText = newText & letter
				End If
			Next
		End If
		ParseFloat = newText
	End Function

	Function GetOrderSessionID
		debug("Entering GetOrderSessionID (CachedOrderSessionID: " & cachedOrderSessionID & ")")
		If (cachedOrderSessionID = -1) Or (isEmpty(cachedOrderSessionID)) Then
			openDB
			Set dbSession = cartDB.Execute(SQL("SELECT id FROM ordersession WHERE sessionid = '" & intSessionID & "' AND ipaddress = '" & strIPAddress & "' AND isdeleted = False AND iscompleted = False"))
			If (dbSession.EOF) Then
				cachedOrderSessionID = -1
			Else
				cachedOrderSessionID = dbSession("id")
			End If
			Set dbSession = Nothing
			closeDB
		End If
		GetOrderSessionID = cachedOrderSessionID
	End Function

	Function GetCartSessionID
		debug("Entering GetCartSessionID (CachedSessionID: " & cachedSessionID & ")")
		If (cachedSessionID = -1) Or (isEmpty(cachedSessionID)) Then
			openDB
			Set dbSession = cartDB.Execute(SQL("SELECT id FROM cartsession WHERE sessionid = '" & intSessionID & "' AND ipaddress = '" & strIPAddress & "'"))
			If (dbSession.EOF) Then
				cachedSessionID = -1
				'print("<b>Session not found.  Critical Error.</b><br />Please make sure you have cookies enabled and allowed for this site.")
				'Response.End
			Else
				cachedSessionID = dbSession("id")
			End If
			Set dbSession = Nothing
			closeDB
		End If
		GetCartSessionID = cachedSessionID
	End Function
	
	Function GetPaymentMethod
		openDB
		Set dbPayment = cartDB.Execute("SELECT paymentmethod FROM ordersession WHERE id = " & GetOrderSessionID & "")
		If (dbPayment.EOF) Then
			RedirectToCheckout
		Else
			GetPaymentMethod = dbPayment("paymentmethod")
		End If
		Set dbPayment = Nothing
		closeDB
	End Function

	Function SendCustomerReceipt(intOrderID)
		openDB

		Set dbEmailAddress = cartDB.Execute("SELECT emailaddress FROM orderbilling, ordersession WHERE orderbilling.id = ordersession.billingid AND ordersession.id = " & intOrderID & "")
		strEmailAddress = dbEmailAddress("emailaddress")
		Set dbEmailAddress = Nothing

		' Parse Send Customer A Receipt Of Order
		Set dbEmail = cartDB.Execute(SQL("SELECT subject, body FROM email WHERE name = 'Customer Receipt'"))
		If (dbEmail.EOF) Then
			ReportError "There was getting the customer receipt from the database."
		Else
			mSubject = ParseAndReplaceEmail(dbEmail("subject"),intOrderID,-1)
			mMessage = ParseAndReplaceEmail(dbEmail("body"),intOrderID,-1)
		End If
		Set dbEmail = Nothing

		SMTP strEmailFrom, strEmailAddress, "", "", mSubject, mMessage
		
		closeDB	
	End Function
	
	Function SendCustomerPaid(intOrderID)
		openDB

		Set dbEmailAddress = cartDB.Execute("SELECT emailaddress FROM orderbilling, ordersession WHERE orderbilling.id = ordersession.billingid AND ordersession.id = " & intOrderID & "")
		strEmailAddress = dbEmailAddress("emailaddress")
		Set dbEmailAddress = Nothing

		' Parse Send Customer A Receipt Of Order
		Set dbEmail = cartDB.Execute(SQL("SELECT subject, body FROM email WHERE name = 'Customer Paid'"))
		If (dbEmail.EOF) Then
			ReportError "There was getting the customer receipt from the database."
		Else
			mSubject = ParseAndReplaceEmail(dbEmail("subject"),intOrderID,-1)
			mMessage = ParseAndReplaceEmail(dbEmail("body"),intOrderID,-1)
		End If
		Set dbEmail = Nothing

		SMTP strEmailFrom, strEmailAddress, "", "mannie@sundancemediagroup.com", mSubject, mMessage
		
		closeDB	
	End Function
	
	Function SendFulfillmentRequest(intOrderID)
		openDB
		
		' Send Fullfillment Email
		Set dbEmail = cartDB.Execute(SQL("SELECT subject, body FROM email WHERE name = 'Fulfillment Request'"))
		If (dbEmail.EOF) Then
			ReportError "There was a problem getting the fulfillment request email from the database."
		Else

			Set dbVendors = cartDB.Execute("SELECT product.vendorid FROM orderdata, product WHERE orderdata.productid = product.id AND orderdata.ordersessionid = " & intOrderID & " GROUP BY product.vendorid")
			If (dbVendors.EOF) Then
				ReportError "Tried to get fulfillment products, but there were no products found."
			Else
				Do Until dbVendors.EOF

					mSubject = ParseAndReplaceEmail(dbEmail("subject"),intOrderID,dbVendors("vendorid"))
					mMessage = ParseAndReplaceEmail(dbEmail("body"),intOrderID,dbVendors("vendorid"))
					
					Set dbVendor = cartDB.Execute("SELECT email, emailcc, emailbcc FROM vendor WHERE id = " & dbVendors("vendorid") & "")
					
'					SMTP strEmailFrom, dbVendor("email"), dbVendor("emailcc"), dbVendor("emailbcc"), mSubject, mMessage
					SMTP strEmailFrom, "mannie@sundancemediagroup.com", "", "", mSubject, mMessage

					dbVendors.MoveNext
					
				Loop
			End If

		End If
		Set dbEmail = Nothing

		closeDB
	End Function
	
	Function SendMerchantReceipt(intOrderID)
		openDB
		
		' Send Mannie A Receipt Of Order
		Set dbEmail = cartDB.Execute(SQL("SELECT subject, body FROM email WHERE name = 'Merchant Receipt'"))
		If (dbEmail.EOF) Then
			ReportError "There was getting the customer receipt from the database."
		Else
			mSubject = ParseAndReplaceEmail(dbEmail("subject"),intOrderID,-2)
			mMessage = ParseAndReplaceEmail(dbEmail("body"),intOrderID,-2)
		End If
		Set dbEmail = Nothing

		SMTP strEmailFrom, "mannie@sundancemediagroup.com", "", "", mSubject, mMessage

		closeDB
	End Function
	
	Function CompleteOrder(intOrderID)
		openDB
		
'		SendCustomerReceipt intOrderID
'		SendMerchantReceipt intOrderID
		If (CBool(GetInfoFromOrder("ispaid",intOrderID))) Then
			SendCustomerPaid intOrderID
			SendFulfillmentRequest intOrderID
			Call DoRegistrationLaunch(intOrderID)
			DoSendAttachment intOrderID
		End If
		
		If (GetInfoFromOrder("paymentmethod",intOrderID) <> "creditdebit") Then
			cartDB.Execute(SQL("UPDATE ordersession SET completed = #" & date & " " & time & "#, iscompleted = True, paid = #" & date & " " & time & "#, ispaid = true WHERE id = " & intOrderID & ""))
		Else
			cartDB.Execute(SQL("UPDATE ordersession SET completed = #" & date & " " & time & "#, iscompleted = True WHERE id = " & intOrderID & ""))
		End If
		closeDB
	End Function
	
	Function DoSendAttachment(orderID)
		openDB
		Set dbOrder = cartDB.Execute("SELECT productid FROM orderdata WHERE ordersessionid = " & orderID & "")
		If Not dbOrder.EOF Then
			Do Until dbOrder.EOF
				Set dbSend = cartDB.Execute("SELECT filename FROM productattachment WHERE productid = " & dbOrder("productid") & "")
				If Not (dbSend.EOF) Then

					Set dbEmail = cartDB.Execute(SQL("SELECT subject, body FROM email WHERE name = 'Order Attachment'"))
					If (dbEmail.EOF) Then
						ReportError "There was getting the from the database."
					Else
						mSubject = ParseAndReplaceEmail(dbEmail("subject"),orderID,-1)
						mMessage = ParseAndReplaceEmail(dbEmail("body"),orderID,-1)
					End If
					Set dbEmail = Nothing
					closeDB
					
					toAddress = GetInfoFromOrder("emailaddress", orderID)
					
					Call SMTPAttachment(strEmailFrom, toAddress, "", "", mSubject, mMessage, "D:\inetpub\vasst_com" & dbSend("filename"))
				End If
									
				dbOrder.MoveNext
			Loop
		End If
		
	End Function
	
	Function ClearOrder
		openDB
		
		cartDB.Execute(SQL("UPDATE cartsession SET expired = true WHERE id = " & GetCartSessionID & ""))
		Session.Abandon
		
		closeDB
	End Function
	
	Function NeedToShip
		openDB
		
		NeedToShip = False
		Set dbToShip = cartDB.Execute(SQL("SELECT product.shippingexempt FROM product, cartdata WHERE cartdata.productid = product.id AND cartdata.cartsessionid = " & GetCartSessionID & ""))
		If Not (dbToShip.EOF) Then
			Do Until (dbToShip.EOF)
'				Response.Write("Product: " & dbToShip("id") & " Exempt: " & dbToShip("shippingexempt")
				If (CBool(dbToShip("shippingexempt")) = False) Then
					NeedToShip = True
					Exit Do
				End If
				dbToShip.MoveNext
			Loop
		End If
		Set dbToShip = Nothing
		
		closeDB
	End Function
	
'	Function NeedToPay
'		openDB
'		Set dbOrderCost = cartDB.Execute(SQL("SELECT subtotalcost, shippingcost, salestaxcost, totalcost FROM ordersession WHERE id = " & intOrderSessionID & ""))
'		If Not (dbOrderCost.EOF) Then
''			fSubTotal = dbOrderCost("subtotalcost")
''			fShipping = dbOrderCost("shippingcost")
''			fSalesTax = dbOrderCost("salestaxcost")
'			fTotal = CDbl(dbOrderCost("totalcost"))
'			If (fTotal = 0) Then
'				NeedToPay = False
'			Else
'				NeedToPay = True
'			End If
'		Else
'			NeedToPay = True
'		End If
'		Set dbOrderCost = Nothing
'		closeDB
'	End Function
%>