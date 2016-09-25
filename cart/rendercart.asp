<%
		If (GetOrderSessionID > 0) Then
			ResetToStep1
		End If

		postBackURL = Request.ServerVariables("URL") & "?function=cart"
		postBackCheckoutURL = Request.ServerVariables("URL") & "?function=checkout"
'id
'session
'created
'saved
'savedemail
'expires
'data

'ProductID:Quantity:PriceID;...
		%>
			<table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse;">
				<form method="post" action="<%=postBackURL%>">
				<tr>
					<td>
						<br />
						<br />
						<table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse;">
							<tr>
								<td colspan="6" align="center"><b>Shopping Cart Items</b></td>
							</tr>
							<tr>
								<td width="60%" nowrap background="/images/banner_back.gif"><font style="color: white; font-size: 10px;"><b>Item</b></font></td>
								<td width="8%" nowrap background="/images/banner_back.gif"><font style="color: white; font-size: 10px;"><b>Price</b></font></td>
								<td width="8%" nowrap background="/images/banner_back.gif"><font style="color: white; font-size: 10px;"><b>Quantity</b></font></td>
								<td width="8%" nowrap background="/images/banner_back.gif"><font style="color: white; font-size: 10px;"><b>Total</b></font></td>
								<td width="8%" nowrap background="/images/banner_back.gif"><font style="color: white; font-size: 10px;"><b>Remove</b></font></td>
<!--								<td width="16%" nowrap background="/images/banner_back.gif"><font style="color: white; font-size: 10px;"><b>Notes</b></font></td>-->
							</tr>
		<%
		openDB
		Set dbCart = cartDB.Execute(SQL("SELECT cartsession.id as sessionid, product.id, product.name, pricing.price, cartdata.quantity FROM cartsession, cartdata, pricing, product WHERE cartsession.id = cartdata.cartsessionid AND cartdata.productid = pricing.productid AND cartdata.pricingid = pricing.id AND cartdata.productid = product.id AND cartsession.sessionid = '" & intSessionID & "' AND pricing.isdeleted = false"))
		If (dbCart.EOF) Then
			%>
				<tr>
					<td colspan="6" align="center">You have no items in your cart, or your session timed out.</td>
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
			Do Until (dbCart.EOF)
				intItemID = Fix(dbCart("id"))
				intItemPricing = ParseFloat(dbCart("price"))
				intQuantity = Fix(dbCart("quantity"))
				strItemName = dbCart("name")
				
				intThisTotal = intItemPricing * intQuantity
				intRunningTotal = intRunningTotal + intThisTotal

				intRunningCount = intRunningCount + intQuantity

				If (strBGColor = "#c0c0c0") Then strBGColor = "#dddddd" Else strBGColor = "#c0c0c0"
				%>
					<tr bgcolor="<%=strBGColor%>">
						<form method="post" action="<%=postBackURL%>">
						<input type="hidden" name="id" value="<%=intItemID%>">
						<td nowrap><font size="-1"><a href="<%=Request.ServerVariables("URL")%>?function=browse&id=<%=intItemID%>"><%=strItemName%></a></font></td>
						<td nowrap><font size="-1">$<%=FormatNumber(intItemPricing,2)%></font></td>
						<td nowrap><font size="-1"><input type="text" size="2" name="quantity" value="<%=intQuantity%>"><input type="image" name="update" src="/cart/icon_check.gif" border="0" value="image" title="Update Quantity"></font></td>
						<td nowrap><font size="-1">$<%=FormatNumber(intThisTotal,2)%></font></td>
						<td nowrap><font size="-1"><input type="image" name="remove" src="/cart/icon_remove.gif" border="0" value="remove" title="Remove"></font></td>
<!--						<td nowrap><font size="-1"><%=strNotes%></font></td>-->
						</form>
					</tr>
				<%
				dbCart.MoveNext
			Loop

			%>
				<tr>
					<td colspan="3" align="right">Total:&nbsp;</td>
					<td>$<%=FormatNumber(intRunningTotal,2)%></td>
					<td><!--Blank--></td>
				</tr>
				<tr>
					<td colspan="4" align="right"><font size="-1">This total does not include tax or shipping charges.</font></td>
					<td><!--Blank--></td>
				</tr>
				<tr>
					<td colspan="6">
						<br />
						<br />
						<b><i>Checkout Checklist</i></b>
						<ul>
							<li>Apply discounts from the browse product page.
							<li>Update quantities by first changing the value, then click the checkmark.  (Only updates one at a time.)
							<li>You are ready to checkout.
						</ul>
					</td>
				</tr>
				<tr>
					<form method="post" action="<%=postBackCheckoutURL%>">
					<td colspan="6" align="right"><br />Checkout:&nbsp;<input type="submit" name="function" value="Checkout"></td>
					</form>
				</tr>
			<%


		End If
		%>
						
						</table>
						<br />
						<br />
					</td>
				</tr>
			</table><br />
		<%
		closeDB
%>