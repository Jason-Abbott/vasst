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
	dateNow = CDate(Date & " " & Time)

	If (Len(strFunction) <= 0) Then
		strFunction = "browse"
	End If

'Page Handler

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
	
	If (InStr(Request.ServerVariables("URL"),"printer.asp") = 0) Then
		RenderViewCartLink
		Select Case (strFunction)
			Case "browse"
				If (strFrmFunction = "add to cart") Then
					If (intQuantity < 1) Then
						RenderBrowser
					Else
						addToCart intFrmID, intQuantity
						If (Len(strDiscountCode) > 0) Then
							applyDiscountToCart strDiscountCode, intFrmID
						ElseIf (Len(Session("product_" & intFrmID)) > 0) Then
							applyDiscountToCart Session("product_" & intFrmID), intFrmID
						End If
						redirect(Request.ServerVariables("URL") & "?function=cart")
					End If
				Else
					If (strFrmFunction = "apply code") Then
						applyDiscountToCart strDiscountCode, intFrmID
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
				If (Len(strSearch) > 0) Then
					Response.Redirect Request.ServerVariables("URL") & "?function=browse&search=" & strSearch
				Else
					Response.Redirect Request.ServerVariables("URL") & "?function=browse"
				End If
		End Select
		RenderViewCartLink
	End If
	
'Rendering Functions
	Dim printCategoriesCurrentLevel : printCategoriesCurrentLevel = 0
	Function printSearchCategories(parentID, intSelectedID)
		openDB
		Set dbBranchCount = cartDB.Execute("SELECT count(*) as branchCount FROM category WHERE parentid = " & parentID & " AND isdeleted = false")
		If (Fix(dbBranchCount("branchCount")) > 0) Then
			printCategoriesCurrentLevel = printCategoriesCurrentLevel + 1
'			Set dbBranchData = cartDB.Execute("SELECT category.id, category.name, count(product.id) as productcount FROM category, product WHERE category.id = product.categoryid AND product.parentid = " & parentID & " AND product.deleted = False GROUP BY category.name, category.id ORDER BY category.name")
			Set dbBranchData = cartDB.Execute("SELECT id, name FROM category WHERE parentid = " & parentID & " AND isdeleted = False ORDER BY name")

			Do Until dbBranchData.EOF
				Set dbThisBranchCount = cartDB.Execute("SELECT count(*) as branchCount FROM category WHERE parentid = " & dbBranchData("id") & " AND isdeleted = false")
				Set dbThisProductCount = cartDB.Execute(SQL("SELECT count(id) as productCount FROM product WHERE categoryid = " & dbBranchData("id") & " AND deleted = false AND hidden = false"))

				If (Fix(dbThisBranchCount("branchCount")) > 0) Or (Fix(dbThisProductCount("productCount")) > 0) Then
					If (Fix(intSelectedID) = Fix(dbBranchData("id"))) Then
						thisSelected = "selected"
					Else
						thisSelected = ""
					End If
					print("<option value=""" & dbBranchData("id") & """ " & thisSelected & " " & thisColor & ">" & printSpacer(printCategoriesCurrentLevel-1) & dbBranchData("name") & "</option>")
				End If
				
				printSearchCategories dbBranchData("id"), intSelectedID
				dbBranchData.MoveNext
			Loop
			printCategoriesCurrentLevel = printCategoriesCurrentLevel - 1
		End If
		closeDB
	End Function
	
	Function RenderBrowser
		%><!--#include virtual="/cart/renderbrowser.asp"--><%
	End Function
	
	Function RenderCart
		%><!--#include virtual="/cart/rendercart.asp"--><%
	End Function

	Function RenderCheckout
		%><!--#include virtual="/cart/rendercheckout.asp"--><%
	End Function
	
	Function RenderStateSelect(strFormName)
		%><!--#include virtual="/cart/renderstateselect.asp"--><%
	End Function
			
	Function RenderCountrySelect(strFormName)
		%><!--#include virtual="/cart/rendercountryselect.asp"--><%
	End Function
			
	Function RenderViewCartLink
		openDB
		Set dbCartCount = cartDB.Execute("SELECT count(productid) as cartcount FROM cartdata WHERE cartsessionid = " & GetCartSessionId & "")
		cartCount = Fix(dbCartCount("cartcount"))
		Set dbCartCount = Nothing
		closeDB
		%>
			<table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse;">
				<tr>
<!--					<td><a href="<%=Request.ServerVariables("URL")%>?function=cart"><%=cartCount%> items in cart.</a></td>-->
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
		openDB
		Set dbPricePreorder = cartDB.Execute("SELECT id, validuntil FROM pricing WHERE productid = " & productid & " AND ispreorder = true AND validuntil >= #" & Date & " 11:59:59 PM# AND isdeleted = false")
		If (dbPricePreorder.EOF) Then
			Set dbPrice = cartDB.Execute("SELECT id FROM pricing WHERE productid = " & productID & " AND isnormal = true AND isdeleted = FALSE")
			If (dbPrice.EOF) Then
				GetNormalPriceID = -1
			Else
				GetNormalPriceID = dbPrice("id")
			End If
		Else
			GetNormalPriceID = dbPricePreorder("id")
		End If
		Set dbPricePreorder = Nothing
	End Function

	Function GetCartCount
		openDB
		Set dbCart = cartDB.Execute(SQL("SELECT sum(quantity) as totalCount FROM cartdata WHERE cartsessionid = " & GetCartSessionID & ""))
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
		closeDB
	End Function
	
	Function GetShippingCost(strShippingType)
		floatShippingTotal = 0.0
		If (strShippingType = "bypass") Or (Not (NeedToShip)) Then
		Else
			If (Len(strShippingType) > 0) Then
				openDB
		
				Set dbOrderCount = cartDB.Execute(SQL("SELECT sum(quantity) as totalQuantity FROM cartdata WHERE cartsessionid = " & GetCartSessionID & ""))
				totalQuantity = dbOrderCount("totalQuantity")
				Set dbOrderCount = Nothing
		
		'		Response.Write("Total Quantity: " & totalQuantity & "<br />")
		
				If (strShippingType = "2day") Then
					strShippingType = "twoday"
				End If
				Set dbShipping = cartDB.Execute(SQL("SELECT TOP 1 " & strShippingType & " FROM shipping WHERE quantity <= " & totalQuantity & " AND productid = 1 ORDER BY quantity DESC"))
				floatShippingTotal = dbShipping(strShippingType)
				Set dbShipping = Nothing
		
				closeDB
			End If
		End If
		GetShippingCost = floatShippingTotal
	End Function

	Function GetShippingCostOld(strShippingType)
		floatShippingTotal = 0.0
		openDB
		Set dbCart = cartDB.Execute(SQL("SELECT productid, quantity FROM orderdata WHERE ordersessionid = " & GetOrderSessionID & ""))
		If Not (dbCart.EOF) Then
			Do Until (dbCart.EOF)
				strShippingType = Replace(strShippingType, "2", "two")
				Set dbShipping = cartDB.Execute(SQL("SELECT TOP 1 " & strShippingType & ", quantity FROM shipping WHERE productid = " & dbCart("productid") & " AND quantity <= " & dbCart("quantity") & " ORDER BY quantity DESC"))
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
		closeDB
		GetShippingCost = floatShippingTotal
	End Function
	
	Function GetCheapestPriceID(intID, intPriceID1, intPriceID2)
		openDB
		Set dbLowestPrice = cartDB.Execute(SQL("SELECT TOP 1 id FROM pricing WHERE productid = " & intID & " AND id IN (" & intPriceID1 & "," & intPriceID2 & ") AND isdeleted = false ORDER BY price"))
		GetCheapestPriceID = dbLowestPrice("id")
		Set dbLowestPrice = Nothing
		closeDB
	End Function

	Function isADiscount(discountcode, productid)
		openDB
		Set dbDiscount = cartDB.Execute(SQL("SELECT id FROM pricing WHERE isdiscount = true AND discountcode = '" & discountcode & "' AND productid = " & productid & " AND price IS NOT NULL AND isdeleted = false"))
		If (dbDiscount.EOF) Then
			isADiscount = False
		Else
			isADiscount = True
		End If
	End Function
	
	Function GetDiscountPriceID(discountcode, productid)
		openDB
		Set dbDiscount = cartDB.Execute(SQL("SELECT id FROM pricing WHERE isdiscount = true AND discountcode = '" & discountcode & "' AND productid = " & productid & " AND price IS NOT NULL AND isdeleted = false"))
		If (dbDiscount.EOF) Then
			GetDiscountPriceID = -1
		Else
			If (isNull(dbDiscount("id"))) Then
				GetDiscountPriceID = -1
			Else
				GetDiscountPriceID = Fix(dbDiscount("id"))
			End If
		End If
		closeDB
	End Function
	
	Function GetDiscountPrice(discountcode, productid)
		openDB
		Set dbDiscount = cartDB.Execute(SQL("SELECT price FROM pricing WHERE isdiscount = true AND discountcode = '" & discountcode & "' AND productid = " & productid & " AND price IS NOT NULL AND isdeleted = false"))
		If (dbDiscount.EOF) Then
			GetDiscountPrice = -1
		Else
			If (isNull(dbDiscount("price"))) Then
				GetDiscountPrice = -1
			Else
				GetDiscountPrice = dbDiscount("price")
			End If
		End If
		closeDB
	End Function

	Function isInCart(productID)
		openDB
		Set dbCart = cartDB.Execute(SQL("SELECT id FROM cartdata WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & productID & ""))
		If (dbCart.EOF) Then
			isInCart = False
		Else
			isInCart = True
		End If
		closeDB
	End Function

	Function isCartDiscounted(productID)
		openDB
		Set dbPricing = cartDB.Execute(SQL("SELECT pricing.discountcode FROM pricing, cartdata WHERE pricing.id = cartdata.pricingid AND cartdata.cartsessionid = " & GetCartSessionID & " AND cartdata.productid = " & productID & " AND pricing.isdeleted = false"))
		If (dbPricing.EOF) Then
			isCartDiscounted = False
		Else
			If Not (isNull(dbPricing("discountcode"))) And Not (Trim(dbPricing("discountcode")) = "") Then
				isCartDiscounted = True
			Else
				isCartDiscounted = False
			End If
		End If
		closeDB
	End Function
	
	Function GetCartPrice(productID)
		openDB
		Set dbPricing = cartDB.Execute(SQL("SELECT pricing.price FROM cartdata, pricing WHERE cartdata.pricingid = pricing.id AND cartdata.cartsessionid = " & GetCartSessionID & " AND cartdata.productid = " & productID & " AND pricing.isdeleted = false"))
		If (dbPricing.EOF) Then
			print("There was a problem with the pricing for an item in your cart, we have removed it for you.")
			removeFromCart(productID)
			GetCartPrice = -1
		Else
			GetCartPrice = dbPricing("price")
		End If
		closeDB
	End Function

	Function GetCartPriceID(productID)
		openDB
		Set dbPricing = cartDB.Execute(SQL("SELECT pricing.id FROM cartdata, pricing WHERE cartdata.pricingid = pricing.id AND cartdata.cartsessionid = " & GetCartSessionID & " AND cartdata.productid = " & productID & " AND pricing.isdeleted = false"))
		If (dbPricing.EOF) Then
			print("There was a problem with the pricing for an item in your cart, we have removed it for you.")
			removeFromCart(productID)
			GetCartPriceID = -1
		Else
			GetCartPriceID = dbPricing("id")
		End If
		closeDB
	End Function

	Function GetShippingCountry
		openDB
		Set dbCountry = cartDB.Execute("SELECT ordershipping.country FROM ordershipping, ordersession WHERE ordershipping.id = ordersession.shippingid AND ordersession.id = " & GetOrderSessionID & "")
		If (dbCountry.EOF) Then
			GetShippingCountry = ""
		Else
			GetShippingCountry = dbCountry("country")
		End If
		Set dbCountry = Nothing
		closeDB
	End Function

'Database Manipulation Functions
	Function addToCart(productID, productQuantity)
		openDB
		Set dbCart = cartDB.Execute(SQL("SELECT pricingid FROM cartdata WHERE cartsessionid = " & GetCartSessionID & " AND cartdata.productid = " & productID & ""))
		If (dbCart.EOF) Then
			cartDB.Execute(SQL("INSERT INTO cartdata ( cartsessionid, productid, quantity, pricingid ) VALUES ( " & GetCartSessionID & ", " & productID & ", " & productQuantity & ", " & GetNormalPriceID(productID) & " )"))
		Else
			intCheapestPrice = GetCheapestPriceID(intFrmID, dbCart("pricingid"), GetNormalPriceID(productID))
			cartDB.Execute(SQL("UPDATE cartdata SET quantity = " & productQuantity & ", pricingid = " & intCheapestPrice & " WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & productID & ""))
		End If
		Set dbCart = Nothing
		closeDB
	End Function

	Function updateCart(itemID, intNewQuantity)
		openDB
		Set dbCart = cartDB.Execute(SQL("SELECT * FROM cartdata WHERE cartsessionid = " & GetCartSessionID & ""))
		If (dbCart.EOF) Then
			%><center><b>Your cart has expired, please create it again.</b></center><%
		Else
			cartDB.Execute(SQL("UPDATE cartdata SET quantity = '" & intNewQuantity & "' WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & itemID & ""))
		End If
		Set dbCart = Nothing
		closeDB
	End Function

	Function removeFromCart(itemID)
		openDB
		Set dbCart = cartDB.Execute(SQL("SELECT * FROM cartdata WHERE cartsessionid = " & GetCartSessionID & ""))
		If (dbCart.EOF) Then
			%><center><b>Your cart has expired, please create it again.</b></centeR><%
		Else
			cartDB.Execute(SQL("DELETE FROM cartdata WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & itemID & ""))
		End If
		Set dbCart = Nothing
		closeDB
	End Function

	Function applyDiscountToCart(discountcode, productid)
		openDB
		Set dbProduct = cartDB.Execute(SQL("SELECT id, productid FROM pricing WHERE discountcode = '" & discountcode & "' AND productid = " & productid & "AND isdiscount = true AND isdeleted = false"))
		If Not (dbProduct.EOF) Then
			intProductID = Fix(dbProduct("productid"))
			intDiscountID = Fix(dbProduct("id"))

			Set dbCart = cartDB.Execute(SQL("SELECT pricingid FROM cartdata WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & productid & ""))
			If Not (dbCart.EOF) Then
				cartDB.Execute(SQL("UPDATE cartdata SET pricingid = " & intDiscountID & " WHERE cartsessionid = " & GetCartSessionID & " AND productid = " & intProductID & ""))
			Else
				Session("product_" & productid) = discountcode
			End If
			Set dbCart = Nothing
		End If
		Set dbProduct = Nothing
		closeDB
	End Function
	
	Function SaveStep1(checkoutName, checkoutTitle, checkoutCompany, checkoutEmailAddress, checkoutPhoneNumber, checkoutAddress, checkoutAddress2, checkoutCity, checkoutState, checkoutZip, checkoutCountry, checkoutShippingName, checkoutShippingTitle, checkoutShippingCompany, checkoutShippingAddress, checkoutShippingAddress2, checkoutShippingCity, checkoutShippingState, checkoutShippingZip, checkoutShippingCountry)
		strStep1Completed = date & " " & time
		openDB
		
		'print("1 Order Session: " & GetOrderSessionID & "<br />")
		If (GetOrderSessionID = -1) Then
			CreateNewOrderSession
		End If
		'print("2 Order Session: " & GetOrderSessionID & "<br />")

		Set dbOrderSession = cartDB.Execute("SELECT billingid, shippingid FROM ordersession WHERE id = " & GetOrderSessionID & "")
		If (dbOrderSession.EOF) Then
			print("<b>The order failed to create, please notify administrator.  Please check to make sure you have cookies enabled.</b>")
			SaveStep1 = False
		Else
			'checkoutName, checkoutTitle, checkoutCompany, checkoutEmailAddress, checkoutPhoneNumber, checkoutAddress, checkoutCity, checkoutState, checkoutZip, checkoutCountry
			'name, title, company, emailaddress, phonenumber, address, city, state, zip, country
			'checkoutShippingName, checkoutShippingTitle, checkoutShippingCompany, checkoutShippingAddress, checkoutShippingCity, checkoutShippingState, checkoutShippingZip, checkoutShippingCountry
			'name, title, company, address, city, state, zip, country
			If (Fix(dbOrderSession("billingid")) = 0) Then
				'print("Insert new billing entry.")
				cartDB.Execute(SQL("INSERT INTO orderbilling ( name, title, company, emailaddress, phonenumber, address, address2, city, state, zip, country ) VALUES ( '" & checkoutName & "', '" & checkoutTitle & "', '" & checkoutCompany & "', '" & checkoutEmailAddress & "', '" & checkoutPhoneNumber & "', '" & checkoutAddress & "', '" & checkoutAddress2 & "', '" & checkoutCity & "', '" & checkoutState & "', '" & checkoutZip & "', '" & checkoutCountry & "' ) "))
				cycleDatabase
				Set dbBilling = cartDB.Execute(SQL("SELECT max(id) as updateID FROM orderbilling WHERE name = '" & checkoutName & "' AND title = '" & checkoutTitle & "' AND company = '" & checkoutCompany & "' AND emailaddress = '" & checkoutEmailAddress & "' AND phonenumber = '" & checkoutPhoneNumber & "' AND address = '" & checkoutAddress & "' AND address2 = '" & checkoutAddress2 & "' AND city = '" & checkoutCity & "' AND state = '" & checkoutState & "' AND zip = '" & checkoutZip & "' AND country = '" & checkoutCountry & "'"))
				cartDB.Execute(SQL("UPDATE ordersession SET billingid = " & dbBilling("updateID") & " WHERE id = " & GetOrderSessionID & ""))
				Set dbBilling = Nothing
			Else
				'print("Update current billing entry.")
				cartDB.Execute(SQL("UPDATE orderbilling SET name = '" & checkoutName & "', title = '" & checkoutTitle & "', company = '" & checkoutCompany & "', emailaddress = '" & checkoutEmailAddress & "', phonenumber = '" & checkoutPhoneNumber & "', address = '" & checkoutAddress & "', address2 = '" & checkoutAddress2 & "', city = '" & checkoutCity & "', state = '" & checkoutState & "', zip = '" & checkoutZip & "', country = '" & checkoutCountry & "' WHERE id = " & dbOrderSession("billingid") & ""))
			End If
			
			If (Fix(dbOrderSession("shippingid")) = 0) Then
				'print("Insert new shipping entry.")
				cartDB.Execute(SQL("INSERT INTO ordershipping ( name, title, company, address, address2, city, state, zip, country ) VALUES ( '" & checkoutShippingName & "', '" & checkoutShippingTitle & "', '" & checkoutShippingCompany & "', '" & checkoutShippingAddress & "', '" & checkoutShippingAddress2 & "', '" & checkoutShippingCity & "', '" & checkoutShippingState & "', '" & checkoutShippingZip & "', '" & checkoutShippingCountry & "' ) "))
				cycleDatabase
				Set dbShipping = cartDB.Execute(SQL("SELECT max(id) as updateID FROM ordershipping WHERE name = '" & checkoutShippingName & "' AND title = '" & checkoutShippingTitle & "' AND company = '" & checkoutShippingCompany & "' AND address = '" & checkoutShippingAddress & "' AND address2 = '" & checkoutShippingAddress2 & "' AND city = '" & checkoutShippingCity & "' AND state = '" & checkoutShippingState & "' AND zip = '" & checkoutShippingZip & "' AND country = '" & checkoutShippingCountry & "'"))
				cartDB.Execute(SQL("UPDATE ordersession SET shippingid = " & dbShipping("updateID") & " WHERE id = " & GetOrderSessionID & ""))
				Set dbShipping = Nothing
			Else
				'print("Update currnet billing entry.")
				cartDB.Execute(SQL("UPDATE ordershipping SET name = '" & checkoutShippingName & "', title = '" & checkoutShippingTitle & "', company = '" & checkoutShippingCompany & "', address = '" & checkoutShippingAddress & "', address2 = '" & checkoutShippingAddress2 & "', city = '" & checkoutShippingCity & "', state = '" & checkoutShippingState & "', zip = '" & checkoutShippingZip & "', country = '" & checkoutShippingCountry & "' WHERE id = " & dbOrderSession("billingid") & ""))
			End If
			
			CopyCartToOrder GetCartSessionID, GetOrderSessionID
			cycleDatabase

			CalculateOrderCost
			cycleDatabase

			cartDB.Execute(SQL("UPDATE ordersession SET step1completed = #" & strStep1Completed & "# WHERE id = " & GetOrderSessionID & ""))
			SaveStep1 = True
			
			If Not (NeedToShip) Then
				SaveStep2 "exempt"
			End If
		End If
		

		closeDB
		cycleDatabase
		RedirectToCheckout
	End Function
	
	Function SaveStep2(shippingMethod)
		strStep2Completed = date & " " & time
		openDB

		'print("Order Session: " & GetOrderSessionID & "<br />")
		If (GetOrderSessionID = -1) Then
			SaveStep2 = False
			RedirectToCheckout
		Else
			cartDB.Execute(SQL("UPDATE ordersession SET step2completed = #" & strStep2Completed & "#, shippingmethod = '" & shippingMethod & "' WHERE id = " & GetOrderSessionID & ""))
			cycleDatabase

			CopyCartToOrder GetCartSessionID, GetOrderSessionID
			cycleDatabase

			CalculateOrderCost
			cycleDatabase
			
			SaveStep2 = True
		End If

		closeDB
		cycleDatabase
		RedirectToCheckout
	End Function

	Function CopyCartToOrder(intCartSessionID, intOrderSessionID)
		openDB
		Set dbCart = cartDB.Execute(SQL("SELECT productid, quantity, pricingid FROM cartdata WHERE cartsessionid = " & intCartSessionID & ""))
		If (dbCart.EOF) Then
			RedirectToCart
		Else
			Do Until (dbCart.EOF)
				Set dbOrderData = cartDB.Execute(SQL("SELECT id FROM orderdata WHERE productid = " & dbCart("productid") & " AND ordersessionid = " & intOrderSessionID & ""))
				If (dbOrderData.EOF) Then
					cartDB.Execute(SQL("INSERT INTO orderdata ( productid, quantity, pricingid, ordersessionid ) VALUES ( " & dbCart("productid") & ", " & dbCart("quantity") & ", " & dbCart("pricingid") & ", " & intOrderSessionID & " ) "))
				Else
					cartDB.Execute(SQL("UPDATE orderdata SET quantity = " & dbCart("quantity") & ", pricingid = " & dbCart("pricingid") & " WHERE id = " & dbOrderData("id") & ""))
				End If
				dbCart.MoveNext
			Loop
		End If
		Set dbCart = Nothing
		cartDB.Execute(SQL("UPDATE cartsession SET movedtoorder = True WHERE id = " & GetCartSessionID & ""))
		closeDB
	End Function
	
	Function SaveStep3(paymentMethod)
		strStep3Completed = date & " " & time
		openDB

		'print("Order Session: " & GetOrderSessionID & "<br />")
		If (GetOrderSessionID = -1) Then
			SaveStep3 = False
			RedirectToCheckout
		Else
			cartDB.Execute(SQL("UPDATE ordersession SET step3completed = #" & strStep3Completed & "#, paymentmethod = '" & paymentMethod & "' WHERE id = " & GetOrderSessionID & ""))
			SaveStep3 = True
			
			If (GetPaymentMethod = "exempt") And (GetTotal("totalcost") = 0) Then
				SaveStep4
				CompleteOrder GetOrderSessionID
			End If
		End If

		closeDB
		cycleDatabase
		RedirectToCheckout
	End Function
	
	Function SaveStep4
		strStep4Completed = date & " " & time
		openDB

		'print("Order Session: " & GetOrderSessionID & "<br />")
		If (GetOrderSessionID = -1) Then
			SaveStep4 = False
			RedirectToCheckout
		Else
			cartDB.Execute(SQL("UPDATE ordersession SET step4completed = #" & strStep4Completed & "# WHERE id = " & GetOrderSessionID & ""))
			SaveStep4 = True
		End If
		
		SendCustomerReceipt GetOrderSessionID
		SendMerchantReceipt GetOrderSessionID

		closeDB
		cycleDatabase
'		RedirectToCheckout
	End Function
	
	Function RedirectToCheckout
		redirect Request.ServerVariables("URL") & "?function=checkout"
	End Function

	Function RedirectToCart
		KillOrderSession
		redirect Request.ServerVariables("URL") & "?function=cart"
	End Function
	
	Function KillOrderSession
		If (GetOrderSessionID > 0) Then
			openDB
			cartDB.Execute(SQL("UPDATE ordersession SET isdeleted = True WHERE id = " & GetOrderSessionID & ""))
			closeDB
		End If
	End Function

	Function IsStep1Completed
		openDB
		If (GetOrderSessionID = -1) Then
			IsStep1Completed = False
			Exit Function
		End If
		
		Set dbOrderSession = cartDB.Execute(SQL("SELECT * FROM ordersession WHERE id = " & GetOrderSessionID & ""))
		If (dbOrderSession.EOF) Then
			IsStep1Completed = False
		Else
			If (Fix(dbOrderSession("billingid")) > 0) And (Fix(dbOrderSession("shippingid")) > 0) And (isDate(dbOrderSession("step1completed"))) Then
				IsStep1Completed = True
			Else	
				IsStep1Completed = False
			End If
		End If
		Set dbOrderSession = Nothing
		closeDB
	End Function

	Function IsStep2Completed
		openDB
		If (GetOrderSessionID = -1) Then
			IsStep2Completed = False
			Exit Function
		End If
		
		Set dbOrderSession = cartDB.Execute(SQL("SELECT * FROM ordersession WHERE id = " & GetOrderSessionID & ""))
		If (dbOrderSession.EOF) Then
			IsStep2Completed = False
		Else
			If (Len(dbOrderSession("shippingmethod")) > 0) And (isDate(dbOrderSession("step2completed"))) Then
				IsStep2Completed = True
			Else	
				IsStep2Completed = False
			End If
		End If
		Set dbOrderSession = Nothing
		closeDB
	End Function

	Function IsStep3Completed
		openDB
		If (GetOrderSessionID = -1) Then
			IsStep3Completed = False
			Exit Function
		End If
		
		Set dbOrderSession = cartDB.Execute(SQL("SELECT * FROM ordersession WHERE id = " & GetOrderSessionID & ""))
		If (dbOrderSession.EOF) Then
			IsStep3Completed = False
		Else
			If (Len(dbOrderSession("shippingmethod")) > 0) And (isDate(dbOrderSession("step3completed"))) Then
				IsStep3Completed = True
			Else	
				IsStep3Completed = False
			End If
		End If
		Set dbOrderSession = Nothing
		closeDB
	End Function

	Function IsStep4Completed
		openDB
		If (GetOrderSessionID = -1) Then
			IsStep4Completed = False
			Exit Function
		End If
		
		Set dbOrderSession = cartDB.Execute(SQL("SELECT * FROM ordersession WHERE id = " & GetOrderSessionID & ""))
		If (dbOrderSession.EOF) Then
			IsStep4Completed = False
		Else
			If (isDate(dbOrderSession("step4completed"))) Then
				If (GetPaymentMethod = "check") Or (GetPaymentMethod = "paypal") Then
					IsStep4Completed = True
				ElseIf (GetPaymentMethod = "creditdebit") And (CBool(dbOrderSession("ispaid"))) Then
					IsStep4Completed = True
				ElseIf (GetPaymentMethod = "exempt") And (GetTotal("totalcost") = 0) Then
					IsStep4Completed = True
				Else
					IsStep4Completed = False
				End If
			Else
				IsStep4Completed = False
			End If
		End If
		Set dbOrderSession = Nothing
		closeDB
	End Function

	Function CreateNewOrderSession
		strDate = date & " " & time
		openDB
		cartDB.Execute(SQL("INSERT INTO ordersession ( sessionid, ipaddress, created, cartsessionid, agent ) VALUES ( '" & intSessionID & "', '" & strIPAddress & "', #" & strDate & "#, " & GetCartSessionID & ", '" & Request.ServerVariables("HTTP_USER_AGENT") & "' )"))
		cycleDatabase
		closeDB
	End Function
	
	Function ResetToStep1
		openDB
		cartDB.Execute(SQL("UPDATE ordersession SET step2completed = null, step3completed = null, step4completed = null, paymentmethod = null, shippingmethod = null WHERE id = " & GetOrderSessionID & ""))
		closeDB
	End Function
	
	Function CalculateOrderCost
		fRunningTotal = 0.0
		fThisTotal = 0.0
		fShippingCost = 0.0
		fSalesTax = 0.0
		fSubTotal = 0.0
		
		openDB
		Set dbOrderSession = cartDB.Execute(SQL("SELECT shippingid, shippingmethod FROM ordersession WHERE id = " & GetOrderSessionID & ""))
		strShippingMethod = dbOrderSession("shippingmethod")
		intShippingID = dbOrderSession("shippingid")
		Set dbOrderSession = Nothing
		
		Set dbShipping = cartDB.Execute(SQL("SELECT state FROM ordershipping WHERE id = " & intShippingID & ""))
		strShippingState = Ucase(dbShipping("state"))
		Set dbShipping = Nothing
		
		Set dbOrder = cartDB.Execute(SQL("SELECT pricing.price, orderdata.quantity FROM orderdata, pricing, product WHERE orderdata.productid = pricing.productid AND orderdata.pricingid = pricing.id AND orderdata.productid = product.id AND orderdata.ordersessionid = " & GetOrderSessionID & " AND pricing.isdeleted = false"))
		If Not (dbOrder.EOF) Then
			Do Until (dbOrder.EOF)
				fItemPricing = ParseFloat(dbOrder("price"))
				intQuantity = Fix(dbOrder("quantity"))
				
				fThisTotal = fItemPricing * intQuantity
				fRunningTotal = fRunningTotal + fThisTotal

				dbOrder.MoveNext
			Loop
			
			fSubTotal = fRunningTotal

			If (strShippingState = "UTAH") Or (strShippingState = "UT") Or (strShippingState = "UTA") Then
				fSalesTax = FormatNumber(fRunningTotal * 0.06,2)
				fRunningTotal = fRunningTotal + fSalesTax 
			End If

			If (isNull(strShippingMethod)) Or (Len(strShippingMethod) > 0) Then
				fShippingCost = GetShippingCost(strShippingMethod)
			End If
			fRunningTotal = fRunningTotal + fShippingCost 
		End If

		cartDB.Execute(SQL("UPDATE ordersession SET subtotalcost = " & fSubTotal & ", shippingcost = " & fShippingCost & ", salestaxcost = " & fSalesTax & ", totalcost = " & fRunningTotal & " WHERE id = " & GetOrderSessionID & ""))
		closeDB
	End Function
		
	Function GetTotal(strWhich)
		openDB
		Set dbOrderCost = cartDB.Execute(SQL("SELECT subtotalcost, shippingcost, salestaxcost, totalcost FROM ordersession WHERE id = " & GetOrderSessionID & ""))
		If Not (dbOrderCost.EOF) Then
			If (strWhich = "subtotal") Then
				GetTotal = CDbl(dbOrderCost("subtotalcost"))
			ElseIf (strWhich = "shipping") Then
				GetTotal = CDbl(dbOrderCost("shippingcost"))
			ElseIf (strWhich = "salestax") Then
				GetTotal = CDbl(dbOrderCost("salestaxcost"))
			Else
				GetTotal = CDbl(dbOrderCost("totalcost"))
			End If
		Else
			GetTotal = -1
		End If
		Set dbOrderCost = Nothing
		closeDB
	End Function	

	Function RenderOrderSummary(intOrderSessionID, bolPrinterFriendly)
		If (intOrderSessionID = -1) Then
			Response.Write("Sorry, your order session expired.  Please check your email for your receipt.  If you don't have it, please contact the administrator and request a receipt be sent to you again.")
		Else
			openDB
			Set dbOrderCost = cartDB.Execute(SQL("SELECT subtotalcost, shippingcost, salestaxcost, totalcost FROM ordersession WHERE id = " & intOrderSessionID & ""))
			If Not (dbOrderCost.EOF) Then
				fSubTotal = dbOrderCost("subtotalcost")
				fShipping = dbOrderCost("shippingcost")
				fSalesTax = dbOrderCost("salestaxcost")
				fTotal = dbOrderCost("totalcost")
			Else
				Response.Write("There is no order with this info.")
				Exit Function
			End If
			Set dbOrderCost = Nothing
	
			%>
				<font style="font-size: 10px;">
				<% If Not (bolPrinterFriendly) Then %>
					<a href="/cart/printer.asp" target="_blank">Printer Friendly Version</a>
				<% End If %>
				</font>
				<table border="1" cellpadding="2" cellspacing="0" style="background: white; font-size: 10px; border-collapse: collapse;" <% If (bolPrinterFriendly) Then %>width="600" align="center"<% Else %>width="100%"<% End If %>>
					<tr>
						<td>
							<br />
							<br />
							<br />
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td valign="top" align="left">
										<font style="font-size: 12px;">
											<b>Merchant Information</b><br />
											Sundance Media Group<br />
											P.O. Box 3 (US Mail Only)<br />
											Stockton, UT 84071<br />
											United States of America<br />
											<br />
											Office: 435-882-8494<br />
											Fax: 435-882-8508<br />
											Email: <a href="mailto:info@sundancemediagroup.com">info@sundancemediagroup.com</a><br />
										</font>
									</td>
									<td valign="top" align="right">
										<font style="font-size: 12px;">
											<b>Customer Information</b><br />
											<% If (Len(GetInfo("title")) > 0) Then %>
												<%=GetInfo("name")%>, <%=GetInfo("title")%><br />
											<% Else %>
												<%=GetInfo("name")%><br />
											<% End If %>
											<% If (Len(GetInfo("company")) > 0) Then %>
												<%=GetInfo("company")%><br />
											<% End If %>
											<%=GetInfo("address")%><br />
											<% If (Len(GetInfo("address2")) > 0) Then %>
												<%=GetInfo("address2")%><br />
											<% End If %>
											<%=GetInfo("city")%>, <%=GetInfo("state")%>&nbsp;<%=GetInfo("zip")%><br />
											<%=GetInfo("country")%><br />
											<br />
											Phone: <%=GetInfo("phonenumber")%><br />
											Email: <%=GetInfo("emailaddress")%><br />
										</font>
									</td>
								</tr>
							</table>
							<br />
							<br />
							<table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse;">
								<tr>
									<td colspan="4" align="center"><b>Final Order Receipt #C<%=PadNumber(GetInfo("ordersessionid"))%></b></td>
								</tr>
								<tr>
									<td width="5%" nowrap background="/images/banner_back.gif"><font style="color: white;"><b>Qty</b></font></td>
									<td width="85%" nowrap background="/images/banner_back.gif"><font style="color: white;"><b>Item</b></font></td>
									<td colspan="2" width="10%" nowrap background="/images/banner_back.gif"><font style="color: white;"><b>Total</b></font></td>
	<!--
									<td width="8%" nowrap background="/images/banner_back.gif"><font style="color: white;"><b>Total</b></font></td>
									<td width="8%" nowrap background="/images/banner_back.gif"><font style="color: white;"><b>Remove</b></font></td>
									<td width="16%" nowrap background="/images/banner_back.gif"><font style="color: white;"><b>Notes</b></font></td>
	-->
								</tr>
			<%
			
			Set dbOrder = cartDB.Execute(SQL("SELECT product.id, product.name, pricing.price, orderdata.quantity FROM orderdata, pricing, product WHERE orderdata.productid = pricing.productid AND orderdata.pricingid = pricing.id AND orderdata.productid = product.id AND orderdata.ordersessionid = " & intOrderSessionID & " AND pricing.isdeleted = false"))
			If (dbOrder.EOF) Then
				%>
					<tr>
						<td colspan="4" align="center">You have no items in your order.</td>
					</tr>
				<%
			Else
				If (bolDebug) Then
					For Each frm In Request.Form
						debug(frm & "=" & Request.Form(frm) & "")
					Next
				End If
	
				intRunningTotal = 0.0
				intRunningCount = 0
				Do Until (dbOrder.EOF)
					intItemID = Fix(dbOrder("id"))
					fItemPricing = ParseFloat(dbOrder("price"))
					intQuantity = Fix(dbOrder("quantity"))
					strItemName = dbOrder("name")
					
					fThisTotal = fItemPricing * intQuantity
	
					If (strBGColor = "#c0c0c0") Then strBGColor = "#dddddd" Else strBGColor = "#c0c0c0"
					%>
						<tr bgcolor="<%=strBGColor%>">
							<td nowrap><font size="-1"><%=intQuantity%></font></td>
							<td nowrap><font size="-1"><a href="<%=Request.ServerVariables("URL")%>?function=browse&id=<%=intItemID%>"><%=strItemName%></a></font></td>
							<td nowrap><font size="-1">$</font></td>
							<td align="right"><font size="-1"><%=FormatNumber(fThisTotal,2)%></font></td>
						</tr>
					<%
					dbOrder.MoveNext
				Loop
	
				%>
					<tr>
						<td colspan="2" align="right"><font size="-1">Sub-Total:&nbsp;</font></td>
						<td><font size="-1">$</font></td>
						<td align="right"><font size="-1"><%=FormatNumber(fSubTotal,2)%></font></td>
					</tr>
				<%
	
				If (CDbl(fSalesTax) > 0.0) Then
					%>
					<tr>
						<td colspan="2" align="right"><font size="-1">Sales Tax:&nbsp;</font></td>
						<td><font size="-1">$</font></td>
						<td align="right"><font size="-1"><%=FormatNumber(fSalesTax,2)%></font></td>
					</tr>
					<%
				End If
	
				%>
					<tr>
						<td colspan="2" align="right"><font size="-1">Shipping:&nbsp;</font></td>
						<td><font size="-1">$</font></td>
						<td align="right"><font size="-1"><%=FormatNumber(fShipping,2)%></font></td>
					</tr>
					<tr>
						<td colspan="2" align="right"><font size="-1"><b>Total:</b>&nbsp;</font></td>
						<td><font size="-1">$</font></td>
						<td align="right"><font size="-1"><b><%=FormatNumber(fTotal,2)%></b></font></td>
					</tr>
				<%
			End If
			%>
				</table>
				<br />
				<br />
				<br />
				</td>
				</tr>
				</table>
				<font style="font-size: 10px;">
				<% If Not (bolPrinterFriendly) Then %>
					<a href="/cart/printer.asp" target="_blank">Printer Friendly Version</a>
				<% End If %>
				</font>
			<%
			Set dbOrder = Nothing
			closeDB
		End If
	End Function
	
'Maintenance Functions
	'Function: TrackSession
	'Purpose: To create, and update, and destroy sessions.
	Function TrackSession
		debug("Enter session tracking.")
		openDB
		cycleDatabase
		'Expire old sessions.
		cartDB.Execute(SQL("UPDATE cartsession SET expired = True WHERE expires < #" & dateNow & "#"))
		'Find current session.
		Set dbSession = cartDB.Execute(SQL("SELECT * FROM cartsession WHERE sessionid = '" & intSessionID & "'"))
		If (dbSession.EOF) Then
			'Create one if this is a new session.
			cartDB.Execute(SQL("INSERT INTO cartsession ( sessionid, ipaddress, created, expires, agent ) VALUES ( '" & intSessionID & "', '" & strIPAddress & "', #" & dateNow & "#, #" & DateAdd("d",7,dateNow) & "#, '" & Request.ServerVariables("HTTP_USER_AGENT") & "' )"))
		Else
			'Update the current session.
			cartDB.Execute(SQL("UPDATE cartsession SET expires = #" & DateAdd("d",7,dateNow) & "#, agent = '" & Request.ServerVariables("HTTP_USER_AGENT") & "' WHERE sessionid = '" & intSessionID & "'"))
		End If
		Set dbSession = Nothing
		cycleDatabase
		closeDB
	End Function
	

%>
