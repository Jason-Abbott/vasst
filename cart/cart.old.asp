<!--#include virtual="/cart/admin/includes.asp"-->
<%
	bolDebug = false
	
	strFunction = Lcase(makeSafe(Request.QueryString("function")))
	intID = Fix(Request.QueryString("id"))
	strSearch = Lcase(makeSafe(Request.QueryString("search")))
	strFrmFunction = Lcase(makeSafe(Request.Form("function")))
	strFrmRemove = Request.Form("remove")
	strFrmRemoveX = Request.Form("remove.x")
	strFrmRemoveY = Request.Form("remove.y")
	strFrmUpdate = Request.Form("update")
	strFrmUpdateX = Request.Form("update.x")
	strFrmUpdateY = Request.Form("update.y")
	intFrmID = Fix(Request.Form("id"))
	intFrmPrice = Fix(Request.Form("price"))
	strDiscountCode = Lcase(makeSafe(Request.Form("discountcode")))
	intQuantity = Fix(Request.Form("quantity"))
	intSessionID = Session.SessionID
	strIPAddress = Request.ServerVariables("REMOTE_ADDR")
	dateNow = CDate(Date & " " & Time)

	Dim cachedSessionID
	cachedSessionID = -1
	Dim cachedOrderSessionID
	cachedOrderSessionID = -1

	If (Len(strFunction) <= 0) Then
		strFunction = "browse"
	End If

'Page Handler
	%><!--include virtual="/cart/top.html"--><%
	RenderViewCartLink

	debug("Function: " & strFunction & "")
	debug("ID: " & intID & "")
	debug("Search: " & strSearch & "")
	debug("Discount Code: " & strDiscountCode & "")
	debug("Form Function: " & strFrmFunction & "")
	debug("Quantity: " & intQuantity & "")
	debug("SessionID: " & intSessionID & "")
	debug("Date: " & dateNow & "")

	debug("strFrmRemove=" & strFrmRemove & "(" & isEmpty(strFrmRemove) & ")")
	debug("strFrmRemoveX=" & strFrmRemoveX & "(" & isEmpty(strFrmRemoveX) & ")")
	debug("strFrmRemoveY=" & strFrmRemoveY & "(" & isEmpty(strFrmRemoveY) & ")")
	debug("strFrmUpdate=" & strFrmUpdate & "(" & isEmpty(strFrmUpdate) & ")")
	debug("strFrmUpdateX=" & strFrmUpdateX & "(" & isEmpty(strFrmUpdateX) & ")")
	debug("strFrmUpdateY=" & strFrmUpdateY & "(" & isEmpty(strFrmUpdateY) & ")")

	TrackSession
	
	debug("Session: " & GetCartSessionID)
	
	Select Case (strFunction)
		Case "browse"
			If (strFrmFunction = "add to cart") Then
				If (intQuantity < 1) Then
					RenderBrowser
				Else
					addToCart intFrmID, intQuantity
					If (Len(strDiscountCode) > 0) Then
						applyDiscountToCart
					End If
					redirect(Request.ServerVariables("URL") & "?function=cart")
				End If
			Else
				If (strFrmFunction = "apply code") Then
					applyDiscountToCart
				End If
				RenderBrowser
			End If
		Case "cart"
			If (strFrmFunction = "update") Or (Not (isEmpty(strFrmUpdate))) Or ((Not isEmpty(strFrmUpdateX)) And (Not isEmpty(strFrmUpdateY))) Then
				If (intQuantity > 0) Then
					updateCart intFrmID, intQuantity
				End If
				redirect(Request.ServerVariables("URL") & "?function=cart")
			ElseIf (strFrmFunction = "remove") Or (Not (isEmpty(strFrmRemove))) Or ((Not isEmpty(strFrmRemoveX)) And (Not isEmpty(strFrmRemoveY))) Then
				removeFromCart intFrmID
				redirect(Request.ServerVariables("URL") & "?function=cart")
			Else
				RenderCart
			End If
		Case "checkout"
			RenderCheckout
		Case Else
			Response.Redirect Request.ServerVariables("URL") & "?function=browse"
	End Select
	RenderViewCartLink
	%><!--include virtual="/cart/bottom.html"--><%

'Rendering Functions
	Function RenderBrowser
		%><!--#include virtual="/cart/renderbrowser.asp"--><%
	End Function
	
	Function RenderCart
		%><!--#include virtual="/cart/rendercart.asp"--><%
	End Function

	Function RenderCheckout
		postBackURL = Request.ServerVariables("URL") & "?function=checkout"
		strCheckout = makeSafe(Request.Form("checkout"))


		%>
<b>The checkout is currently unavailable.  Please give us your name and email address, and we will notify you the moment the checkout is available.</b>
<table border="0" cellpadding="0" cellspacing="0">
  <form name="subscribe" method="post" action="http://www.vasst.com/registration/subscribe.asp">
  <input type="hidden" name="list" value="45">
  <tr>
    <td align="center" colspan="2">Product Notification Signup</td>
  </tr>
  <tr>
    <td align="right">Name:</td>
    <td><input name="name" size="20"></td>
  </tr>
  <tr>
    <td align="right">Email:</td>
    <td><input name="email" size="20"></td>
  </tr>
  <tr>
    <td align="center" colspan="2">
      <input type="submit" name="add" value="Subscribe">
    </td>
  </tr>
  </form>
</table><br />
		<%

		%><!--include virtual="/cart/rendercheckout.asp"--><%
	End Function
	
	Function RenderCountrySelect(strFormName)
		%><!--#include virtual="/cart/rendercountryselect.asp"--><%
	End Function
			
	Function RenderViewCartLink
		%>
			<table border="1" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse;">
				<tr>
					<td align="right">
						<a href="<%=Request.ServerVariables("URL")%>?function=browse">Browse Products</a> |
						<a href="<%=Request.ServerVariables("URL")%>?function=cart">View Cart</a> |
						<a href="<%=Request.ServerVariables("URL")%>?function=checkout">Checkout</a>
					</td>
				</tr>
			</table><br />
		<%
	End Function

'Data Gathering Functions
	Function GetNormalPriceID(productID)
		openDatabase
		Set dbPrice = dbConnection.Execute("SELECT id FROM pricing WHERE productid = " & productID & " AND (discountcode IS NULL OR discountcode = '')")
		If (dbPrice.EOF) Then
			GetNormalPriceID = -1
		Else
			GetNormalPriceID = dbPrice("id")
		End If
	End Function

	Function GetCartSessionID
		debug("Entering GetCartSessionID (CachedSessionID: " & cachedSessionID & ")")
		If (cachedSessionID = -1) Or (isEmpty(cachedSessionID)) Then
			openDatabase
			Set dbSession = dbConnection.Execute(SQL("SELECT id FROM cartsession WHERE sessionid = '" & intSessionID & "' AND ipaddress = '" & strIPAddress & "'"))
			If (dbSession.EOF) Then
				print("<b>Session not found.  Critical Error.</b><br />Please make sure you have cookies enabled and allowed for this site.")
				Response.End
			Else
				cachedSessionID = dbSession("id")
			End If
			Set dbSession = Nothing
			closeDatabase
		End If
		GetCartSessionID = cachedSessionID
	End Function
	
	Function GetCartCount
		openDatabase
		Set dbCart = dbConnection.Execute(SQL("SELECT sum(quantity) as totalCount FROM cartdata WHERE cartsessionid = " & GetCartSessionID & ""))
		If (dbCart.EOF) Then
			GetCartCount = 0
		Else
			If (isNull(dbCart("totalCount"))) Then
				GetCartCount = 0
			Else
				GetCartCount = Fix(dbCart("totalCount"))
			End If
		End If
		Set dbCart = Nothing
		closeDatabase
	End Function
	
	Function GetShippingCost(strShippingType)
		floatShippingTotal = 0.0
		openDatabase
		Set dbCart = dbConnection.Execute(SQL("SELECT productid, quantity FROM cartdata WHERE cartsessionid = " & GetCartSessionID & ""))
		If Not (dbCart.EOF) Then
			Do Until (dbCart.EOF)
				strShippingType = Replace(strShippingType, "2", "two")
				Set dbShipping = dbConnection.Execute(SQL("SELECT TOP 1 " & strShippingType & ", quantity FROM shipping WHERE productid = " & dbCart("productid") & " AND quantity <= " & dbCart("quantity") & " ORDER BY quantity DESC"))
				If (dbShipping.EOF) Then
					Response.Write("Shipping for this product is not found, removing from cart.")
					removeFromCart dbCart("productid")
				Else
					debug("Quantity: " & dbCart("quantity") & ">=" & dbShipping("quantity") & " -- $" & dbShipping(strShippingType) & " by " & strShippingType & "<br />")
					floatShippingTotal = floatShippingTotal + dbShipping(strShippingType)
				End If
				dbCart.MoveNext
			Loop
		End If
		Set dbCart = Nothing
		closeDatabase
		GetShippingCost = floatShippingTotal
	End Function
	
	Function GetCheapestPriceID(intID, intPriceID1, intPriceID2)
		openDatabase
		Set dbLowestPrice = dbConnection.Execute("SELECT TOP 1 id FROM pricing WHERE productid = " & intID & " AND id IN (" & intPriceID1 & "," & intPriceID2 & ") ORDER BY price")
		GetCheapestPriceID = dbLowestPrice("id")
		Set dbLowestPrice = Nothing
		closeDatabase
	End Function

	Function isInCart(productID)
		openDatabase
		Set dbCart = dbConnection.Execute(SQL("SELECT id FROM cartdata WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & productID & ""))
		If (dbCart.EOF) Then
			isInCart = False
		Else
			isInCart = True
		End If
		closeDatabase
	End Function

	Function isCartDiscounted(productID)
		openDatabase
		Set dbPricing = dbConnection.Execute(SQL("SELECT pricing.discountcode FROM pricing, cartdata WHERE pricing.id = cartdata.pricingid AND cartdata.cartsessionid = " & GetCartSessionID & " AND cartdata.productid = " & productID & ""))
		If (dbPricing.EOF) Then
			isCartDiscounted = False
		Else
			If Not (isNull(dbPricing("discountcode"))) And Not (Trim(dbPricing("discountcode")) = "") Then
				isCartDiscounted = True
			Else
				isCartDiscounted = False
			End If
		End If
		closeDatabase
	End Function
	
	Function GetCartPrice(productID)
		openDatabase
		Set dbPricing = dbConnection.Execute(SQL("SELECT pricing.price FROM cartdata, pricing WHERE cartdata.pricingid = pricing.id AND cartdata.cartsessionid = " & GetCartSessionID & " AND cartdata.productid = " & productID & ""))
		If (dbPricing.EOF) Then
			print("There was a problem with the pricing for an item in your cart, we have removed it for you.")
			removeFromCart(productID)
			GetCartPrice = -1
		Else
			GetCartPrice = dbPricing("price")
		End If
		closeDatabase
	End Function

	Function GetCartPriceID(productID)
		openDatabase
		Set dbPricing = dbConnection.Execute(SQL("SELECT pricing.id FROM cartdata, pricing WHERE cartdata.pricingid = pricing.id AND cartdata.cartsessionid = " & GetCartSessionID & " AND cartdata.productid = " & productID & ""))
		If (dbPricing.EOF) Then
			print("There was a problem with the pricing for an item in your cart, we have removed it for you.")
			removeFromCart(productID)
			GetCartPriceID = -1
		Else
			GetCartPriceID = dbPricing("id")
		End If
		closeDatabase
	End Function

	Function GetOrderSessionID
		debug("Entering GetOrderSessionID (CachedOrderSessionID: " & cachedOrderSessionID & ")")
		If (cachedOrderSessionID = -1) Or (isEmpty(cachedOrderSessionID)) Then
			openDatabase
			Set dbSession = dbConnection.Execute(SQL("SELECT id FROM ordersession WHERE sessionid = '" & intSessionID & "' AND ipaddress = '" & strIPAddress & "'"))
			If (dbSession.EOF) Then
				cachedOrderSessionID = -1
			Else
				cachedOrderSessionID = dbSession("id")
			End If
			Set dbSession = Nothing
			closeDatabase
		End If
		GetOrderSessionID = cachedOrderSessionID
	End Function

	Function GetRandomString
		strReturn = ""
		For X = 1 To 12
			intRandom = Round(3 * Rnd) + 1
			Response.Write("Random: " & intRandom & "<br />")
			If (intRandom = 1) Then
				strChar = Round(10 * Rnd) + 48
			ElseIf (intRandom = 2) Then
				strChar = Round(26 * Rnd) + 65
			ElseIf (intRandom = 3) Then
				strChar = Round(26 * Rnd) + 97
			End If
			strReturn = strReturn + strChar
		Next
		GetRandomString = strReturn
	End Function				
	
	Function IsStep1Completed
		IsStep1Completed = IsStepCompleted(1)
	End Function

	Function IsStep2Completed
		IsStep2Completed = IsStepCompleted(2)
	End Function

	Function IsStep3Completed
		IsStep3Completed = IsStepCompleted(3)
	End Function

	Function IsStep4Completed
		IsStep4Completed = IsStepCompleted(4)
	End Function
	
	Function IsStepCompleted(intStep)
		openDatabase
		Set dbOrderSession = dbConnection.Execute(SQL("SELECT step" & intStep & "completed FROM ordersession WHERE id = " & GetOrderSessionID & " AND step" & intStep & "completed IS NOT NULL"))
		If (dbOrderSession.EOF) Then
			IsStepCompleted = False
		Else
			IsStepCompleted = True
		End If
		Set dbOrderSession = Nothing
		closeDatabase
	End Function

'Database Manipulation Functions
	Function addToCart(productID, productQuantity)
		openDatabase
		Set dbCart = dbConnection.Execute(SQL("SELECT pricingid FROM cartdata WHERE cartsessionid = " & GetCartSessionID & " AND cartdata.productid = " & productID & ""))
		If (dbCart.EOF) Then
			dbConnection.Execute(SQL("INSERT INTO cartdata ( cartsessionid, productid, quantity, pricingid ) VALUES ( " & GetCartSessionID & ", " & productID & ", " & productQuantity & ", " & GetNormalPriceID(productID) & " )"))
		Else
			intCheapestPrice = GetCheapestPriceID(intFrmID, dbCart("pricingid"), GetNormalPriceID(productID))
			dbConnection.Execute(SQL("UPDATE cartdata SET quantity = " & productQuantity & ", pricingid = " & intCheapestPrice & " WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & productID & ""))
		End If
		Set dbCart = Nothing
		closeDatabase
	End Function

	Function updateCart(itemID, intNewQuantity)
		openDatabase
		Set dbCart = dbConnection.Execute(SQL("SELECT * FROM cartdata WHERE cartsessionid = " & GetCartSessionID & ""))
		If (dbCart.EOF) Then
			%><center><b>Your cart has expired, please create it again.</b></center><%
		Else
			dbConnection.Execute(SQL("UPDATE cartdata SET quantity = '" & intNewQuantity & "' WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & itemID & ""))
		End If
		Set dbCart = Nothing
		closeDatabase
	End Function

	Function removeFromCart(itemID)
		openDatabase
		Set dbCart = dbConnection.Execute(SQL("SELECT * FROM cartdata WHERE cartsessionid = " & GetCartSessionID & ""))
		If (dbCart.EOF) Then
			%><center><b>Your cart has expired, please create it again.</b></centeR><%
		Else
			dbConnection.Execute(SQL("DELETE FROM cartdata WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & itemID & ""))
		End If
		Set dbCart = Nothing
		closeDatabase
	End Function

	Function applyDiscountToCart
		openDatabase
		Set dbProduct = dbConnection.Execute(SQL("SELECT id, productid FROM pricing WHERE discountcode = '" & strDiscountCode & "'"))
		If Not (dbProduct.EOF) Then
			intProductID = Fix(dbProduct("productid"))
			intDiscountID = Fix(dbProduct("id"))

			Set dbCart = dbConnection.Execute(SQL("SELECT pricingid FROM cartdata WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & intProductID & ""))
			If Not (dbCart.EOF) Then
				dbConnection.Execute(SQL("UPDATE cartdata SET pricingid = " & intDiscountID & " WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & intProductID & ""))
			End If
			Set dbCart = Nothing
		End If
		Set dbProduct = Nothing
		closeDatabase
	End Function
	
	Function SaveStep1
		strStep1Completed = date & " " & time
		openDatabase
		
		print("1 Order Session: " & GetOrderSessionID & "<br />")
		If (GetOrderSessionID = -1) Then
			CreateNewOrderSession
		End If
		print("2 Order Session: " & GetOrderSessionID & "<br />")

		Set dbOrderSession = dbConnection.Execute("SELECT billingid, shippingid WHERE ")
		If (IsStep1Completed) Then
			print("update")
		Else
			print("insert")
		End If
				
		closeDatabase
		cycleDatabase
	End Function
	
	Function IsStep1Completed
		If (GetOrderSessionID = -1) Then
			IsStep1Completed = False
			Exit Function
		End If
		
		
		
	End Function

	Function CreateNewOrderSession
		strDate = date & " " & time
		openDatabase
		dbConnection.Execute("INSERT INTO ordersession ( sessionid, ipaddress, created, cartsessionid ) VALUES ( '" & intSessionID & "', '" & strIPAddress & "', #" & strDate & "#, " & GetCartSessionID & " )")
		cycleDatabase
		closeDatabase
	End Function
	
'Maintenance Functions
	'Function: TrackSession
	'Purpose: To create, and update, and destroy sessions.
	Function TrackSession
		debug("Enter session tracking.")
		openDatabase
		cycleDatabase
		'Expire old sessions.
		dbConnection.Execute(SQL("UPDATE cartsession SET expired = True WHERE expires < #" & dateNow & "#"))
		'Find current session.
		Set dbSession = dbConnection.Execute(SQL("SELECT * FROM cartsession WHERE sessionid = '" & intSessionID & "'"))
		If (dbSession.EOF) Then
			'Create one if this is a new session.
			dbConnection.Execute(SQL("INSERT INTO cartsession ( sessionid, ipaddress, created, expires ) VALUES ( '" & intSessionID & "', '" & strIPAddress & "', #" & dateNow & "#, #" & DateAdd("d",7,dateNow) & "# )"))
		Else
			'Update the current session.
			dbConnection.Execute(SQL("UPDATE cartsession SET expires = #" & DateAdd("d",7,dateNow) & "# WHERE sessionid = '" & intSessionID & "'"))
		End If
		Set dbSession = Nothing
		cycleDatabase
		closeDatabase
	End Function
	
	'[14:28:49] <Mannie Frances> wfb
	'[14:28:51] <Mannie Frances> Your Login name is: vrn227251412
	'[14:28:57] <Mannie Frances> vasst123
%>