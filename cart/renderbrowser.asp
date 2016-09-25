<%
	If (Len(Request.QueryString("categoryid")) > 0) Then
		searchCategoryID = ParseInt(Request.QueryString("categoryid"))
	Else
		searchCategoryID = -1
	End IF

	If (searchCategoryID > 0) Then
		searchFor = "AND product.categoryid = " & searchCategoryID
		postBackURL = Request.ServerVariables("URL") & "?function=browse&categoryid=" & searchCategoryID
	ElseIf (intID > 0) Then
		searchFor = "AND product.id = " & intID
		postBackURL = Request.ServerVariables("URL") & "?function=browse&id=" & intID
	ElseIf (Len(strSearch) > 0) And (strSearch <> "") Then
		searchFor = "AND product.keyword = '" & strSearch & "'"
		postBackURL = Request.ServerVariables("URL") & "?function=browse&search=" & strSearch
	Else
		searchFor = ""
	End If

	If (intQuantity < 1) Then
		intQuantity = 1
	End If

'	If (Request.ServerVariables("REMOTE_ADDR") = "205.208.240.201") Then
		%>
					<table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse;">
						<form method="get" action="<%=postBackURL%>">
						<tr>
							<td>
			Show Products from <select size="1" name="categoryid">
			<option value="-1">All</option>
		<%
		printSearchCategories -1, Request.QueryString("categoryid")
		%>
			</select><input type="submit" value="Show" />
			<input type="hidden" name="function" value="browse" />
							</td>
						</tr>
						</form>
					</table><br />
		<%
'	End If
	
		openDB
		Set dbProduct = cartDB.Execute(SQL("SELECT product.id, product.name, product.description, product.thumbnail, product.added FROM pricing LEFT JOIN product ON product.id = pricing.productid WHERE (pricing.isnormal = true) AND product.deleted = False AND product.hidden = False " & searchFor & " AND pricing.isdeleted = false ORDER BY product.ordinal DESC, product.added DESC, product.categoryid, product.name"))
		If (dbProduct.EOF) Then
			print("<br /><br />There are no products available in this listing, or the product you were searching for is not found.  Please check the URL and try again.<br /><br />")
			print("<a href=""/cart/"">Click Here To Browse Listing</a><br /><br /><br />")
		Else
			Do Until dbProduct.EOF
				intProductID = dbProduct("id")
				strProductName = dbProduct("name")
				strProductDescriptoin = dbProduct("description")
				strProductThumbnail = dbProduct("thumbnail")
				
				Set dbNormalPrice = cartDB.Execute(SQL("SELECT id, price FROM pricing WHERE productid = " & intProductID & " AND isnormal = true AND isdeleted = false"))
				intNormalPriceID = dbNormalPrice("id")
				floatNormalPrice = dbNormalPrice("price")
				Set dbNormalPrice = Nothing
				
				Set dbPreorderPrice = cartDB.Execute(SQL("SELECT id, price, validuntil FROM pricing WHERE productid = " & intProductID & " AND ispreorder = true AND isdeleted = false"))
				If (dbPreorderPrice.EOF) Then
					hasPreorderPricing = False
				Else
					If (DateDiff("d",dbPreorderPrice("validuntil"),Date) > 0) Then
						hasPreorderPricing = False
					Else
						hasPreorderPricing = True
						intPreorderPriceID = dbPreorderPrice("id")
						floatPreorderPrice = dbPreorderPrice("price")
						datePreorderUntil = dbPreorderPrice("validuntil")
					End If
				End If
				Set dbPreorderPrice = Nothing
				
				If (isInCart(intProductID)) Then
					If (isCartDiscounted(intProductID)) Then
						isDiscounted = True
						intPriceID = GetCartPriceID(intProductID)
						discountPrice = GetCartPrice(intProductID)
						If (discountPrice = -1) then
							isDiscounted = False
						End If
					Else
						isDiscounted = False
					End If
				Else
					If (Len(Session("product_" & intProductID)) > 0) Then
						If (isADiscount(Session("product_" & intProductID), intProductID)) Then
							isDiscounted = True
							intPriceID = GetDiscountPriceID(Session("product_" & intProductID), intProductID)
							discountPrice = GetDiscountPrice(Session("product_" & intProductID), intProductID)
							If (intPriceID = -1) Or (discountPrice = -1) Then
								ReportError "An error was found in the discount check, it passed the discount check, however there was no price or id available."
								isDiscounted = False
							End If
						Else
							isDiscounted = False
						End If
					Else
						isDiscounted = False
					End If
				End If
				
				If (bolDebug) Then
					debug("Is in cart: " & isInCart(intProductID) & "")
					debug("Is discounted: " & isCartDiscounted(intProductID) & "")
					debug("Cart Price: " & GetCartPrice(intProductID) & "")
				End If
				
				%>
					<table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse;">
						<form method="post" action="<%=postBackURL%>#product_<%=intProductID%>">
						<tr>
							<td>
								<a name="product_<%=intProductID%>">
								<table border="0" cellpadding="20" cellspacing="0" width="100%">
									<tr>
										<td align="center" valign="middle">
											<% If (Len(strProductThumbnail) > 0) Then %>
											<img src="<%=strProductThumbnail%>" width="100"/>
											<% Else %>
											No product image<br /> is avilable.
											<% End If %>
										</td>
										<td>
											<font style="font-size: 13pt; font-weight: bold;"><center><%=strProductName%></center></font>
											<% If (DateDiff("d", dbProduct("added"), Now) < 30) Then %>
											<center><span style="background: #FFCC00; color: black; font-weight: bold;">&nbsp;&nbsp;NEW!&nbsp;&nbsp;</span></center>
											<% End If %>
											<br /><font style="font-size: 10pt;"><%=Replace(strProductDescriptoin,vbNewline,"<br />")%><br />
											<br />
											<% If Not (isDiscounted) And (hasPreorderPricing) Then %>
												<b><i>Order for $<%=FormatNumber(floatPreorderPrice,2)%> until <%=datePreorderUntil%>.</i></b><br />
											<% End If %>
											<b>Our Price: 
											<% If (isDiscounted) Or (hasPreorderPricing) Then %>
												<font style="color: #333333; text-decoration: line-through;">$<%=FormatNumber(floatNormalPrice,2)%></font>
												<% If (isDiscounted) Then %>
													$<%=FormatNumber(discountPrice,2)%> (discounted)
												<% End If %>
											<% Else %>
												$<%=FormatNumber(floatNormalPrice,2)%>
											<% End If %>
											</b><br />
											<%
												If Not (isDiscounted) Then
													If (Fix(intProductID) = intFrmID) Then
														thisDiscountCode = strDiscountCode
													Else
														thisDiscountCode = ""
													End If
													%>
														Discount Code: <input type="text" size="10" name="discountcode" value="<%=thisDiscountCode%>"><input type="submit" name="function" value="Apply Code"><br />
													<%
													If (strDiscountCode <> "") And (Fix(intProductID) = intFrmID) Then
														%>
															<font style="color: red;">This code isn't a valid discount code.</font><br />
														<%
													End If
												Else
													%>
														<input type="hidden" name="discountcode" value="<%=strDiscountCode%>">
													<%
												End If
											%>
											Quantity: <input type="text" size="3" name="quantity" value="<%=intQuantity%>"><input type="submit" name="function" value="Add To Cart">
											<%
												If (Fix(intQuantity) < 1) Then
													%>
														<font style="color: red;">This isn't a valid quantity.</font><br />
													<%
												End If
											%>
											</font>
											<input type="hidden" name="id" value="<%=intProductID%>">
<!--											<input type="hidden" name="price" value="<%=priceID%>">-->
										</td>
									</tr>
								</table>
							</td>
						</tr>
						</form>
					</table><br />
				<%
				dbProduct.MoveNext
			Loop
		End If
		Set dbProduct = Nothing
		closeDB
%>