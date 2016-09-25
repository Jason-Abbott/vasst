<!--#include virtual="/cart/admin/includes.asp"-->
<%
	'[14:28:49] <Mannie Frances> wfb
	'[14:28:51] <Mannie Frances> Your Login name is: vrn227251412
	'[14:28:57] <Mannie Frances> vasst1234

'ResetDatabase -- Handler
	Function ResetDatabaseHandler
		debug("Entering ResetDatabaseHandler handler.")
		resetDatabase
		redirect(Request.ServerVariables("URL") & "?message=Database has been cleared.")
	End Function

	Function resetDatabase
		resetTable("category")
	End Function

	Sub resetTable(strTable)
		sqlDeleteFromTable = "DELETE FROM " & strTable & ""
		sqlResetIdentity = "ALTER TABLE " & strTable & " ALTER COLUMN id IDENTITY(1,1)"

		openDB
		cartDB.Execute(SQL(sqlDeleteFromTable))
		debug("Cleared all records from table <b>" & strTable & "</b>.")
		cartDB.Execute(SQL(sqlResetIdentity))
		debug("Reset identity on column <b>id</b> for table <b>" & strTable & "</b>.")
		closeDB
	End Sub

' Add Category
	Function AddCategoryHandler
		debug("Entering AddCategoryHandler handler.")

		strCategoryName = Trim(Request.Form("strCategoryName"))
		intParentID = Fix(Request.Form("intParentID"))
		queryParentID = Fix(Request.QueryString("parentid"))

		If (intParentID < 1) Then
			If (queryParentID < 1) Then
				intParentID = -1
			Else
				intParentID = queryParentID
			End If
		End If

		If (Request.Form("process") = "true") Then
			If (strCategoryName = "") Then
				strBlankMessage = Message("You must enter a category name.")
				Response.Redirect Request.ServerVariables("URL") & "?function=AddCategory&success=false&message=" & strBlankMessage & "&parentid=" & intParentID
			Else
				categoryID = addCategory(strCategoryName, intParentID)
				If (categoryID = -1) Then
					strFailedMessage = Message("The category already exists, please try a new name.")
					Response.Redirect Request.ServerVariables("URL") & "?function=AddCategory&success=false&message=" & strFailedMessage
				Else
					strSuccessMessage = Message("Added category '" & strCategoryName & "' as ID# " & categoryID & ".")
					Response.Redirect Request.ServerVariables("URL") & "?success=true&message=" & strSuccessMessage
				End If
			End If
		Else
			renderAddCategoryInput strCategoryName, intParentID
		End If
	End Function

	Function renderAddCategoryInput(strCategoryName, intParentID)
		%>
Enter the name of the category you would like to add.
<table border="0" cellpadding="2" cellspacing="0">
	<tr>
		<td>Parent Category</td>
		<td>
			<select size="1" name="intParentID">
				<option value="-1"> ... No Parent ... </option>
				<% printCategoryTreeOptions intParentID %>
			</select>
		</td>
	</tr>
	<tr>
		<td>Category Name</td>
		<td><input type="text" name="strCategoryName" value="<%=strCategoryName%>"/></td>
	</tr>
	<tr>
		<td><!--Blank--></td>
		<td>
			<input type="submit" name="function" value="Add">
		</td>
	</tr>
</table>
		<%
	End Function

	Function addCategory(strCategoryName, intParentID)
		sqlFindCategory = "SELECT id FROM category WHERE name = '" & strCategoryName & "' AND parentid = " & intParentID & ""
		sqlAddCategory = "INSERT INTO category ( name, parentid ) VALUES ( '" & strCategoryName & "', " & intParentID & " )"

		openDB
		Set categoryID = cartDB.Execute(SQL(sqlFindCategory))
		If (categoryID.EOF) Then
			cartDB.Execute(SQL(sqlAddCategory))
			closeDB
			openDB
			Set categoryID = cartDB.Execute(SQL(sqlFindCategory))
			intResult = categoryID("id")
		Else
			intResult = -1
		End If
		closeDB

		addCategory = intResult
	End Function

' Edit Category
	Function EditCategoryHandler
		debug("Entering AddCategoryHandler handler.")

		strCategoryName = Trim(Request.Form("strCategoryName"))
		intID = Fix(Request.Form("intID"))
		intParentID = Fix(Request.Form("intParentID"))

		If (Request.Form("process") <> "true") Then
			renderEditCategorySelection
		Else
			If (Lcase(Request.Form("function")) = "edit") Then
				If (isNumeric(intID)) And (intID > 0) Then
					openDB
					Set dbCategory = cartDB.Execute("SELECT name, parentid FROM category WHERE id = " & intID & "")
					If (dbCategory.EOF) Then
						debug("DBCategory EOF in EditCategoryHandler")
						print("Invalid entry, please click edit category again on the left, and try again.<br />")
					Else
						strCategoryName = dbCategory("name")
						intParentID = Fix(dbCategory("parentid"))
					End If
					Set dbCategory = Nothing
					closeDB
					renderEditCategory intID, strCategoryName, intParentID
					'closeDB
				Else
					strFailedMessage = "Please select a category to edit."
					redirect Request.ServerVariables("URL") & "?function=EditCategory&success=false&message=" & strFailedMessage
				End If
			ElseIf (Lcase(Request.Form("function")) = "save") Then
				saveCategory intID, strCategoryName, intParentID
			End If
		End If
	End Function

	Function renderEditCategorySelection
		%>
Select the name of the category you would like to edit.
<table border="0" cellpadding="2" cellspacing="0">
	<tr>
		<td>Edit Category</td>
		<td>
			<select size="1" name="intID">
				<option value="-1"> ... Select a category to edit. ... </option>
				<% printCategoryTreeOptions -1 %>
			</select>
		</td>
	</tr>
	<tr>
		<td><!--Blank--></td>
		<td>
			<input type="submit" name="function" value="Edit">
		</td>
	</tr>
</table>
		<%
	End Function

	Function renderEditCategory(intID, strCategoryName, intParentID)
		%>
Select the name of the category you would like to edit.
<table border="0" cellpadding="2" cellspacing="0">
	<tr>
		<td>New Parent Category</td>
		<td>
			<select size="1" name="intParentID">
				<option value="-1" <% If (intParentID = -1) Then Response.Write("selected") %>> ... No Parent ... </option>
				<% printCategoryTreeOptionsWithoutCurrent intID, intParentID %>
			</select>
		</td>
	</tr>
	<tr>
		<td>New Category Name</td>
		<td><input type="text" name="strCategoryName" value="<%=strCategoryName%>"/></td>
	</tr>
	<tr>
		<td><!--Blank--></td>
		<td>
			<input type="hidden" name="intID" value="<%=intID%>">
			<input type="submit" name="function" value="Save">
		</td>
	</tr>
</table>
		<%
	End Function

	Function saveCategory(intID, strCategoryName, intParentID)
		debug("Entering Save Category")
		openDB
		cartDB.Execute("UPDATE category SET name = '" & strCategoryName & "', parentid = " & intParentID & " WHERE id = " & intID & "")
		strSuccessMessage = "Saved category '" & strCategoryName & "'."
		redirect Request.ServerVariables("URL") & "?message=" & strSuccessMessage
		closeDB
	End Function
	
' Delete Category
	Function DeleteCategoryHandler
		debug("Entering DeleteCategoryHandler")

		intID = Fix(Request.Form("intID"))

		If (Request.Form("process") = "true") Then
			If (intID < 1) Then
				strFailedMessage = "Please select a category to delete."
				redirect Request.ServerVariables("URL") & "?function=DeleteCategory&success=false&message=" & strFailedMessage
			Else
				strCategoryName = deleteCategory(intID)
				strSuccessMessage = message("Deleted category '" & strCategoryName & "' as ID# " & intID & ".")
				Response.Redirect Request.ServerVariables("URL") & "?success=true&message=" & strSuccessMessage
			End If
		Else
			renderDeleteCategorySelection
		End If
	End Function
	
	Function renderDeleteCategorySelection
		%>
Select the name of the category you would like to edit.
<table border="0" cellpadding="2" cellspacing="0">
	<tr>
		<td>Delete Category</td>
		<td>
			<select size="1" name="intID">
				<option value="-1"> ... Select a category to delete. ... </option>
				<% printCategoryTreeOptions -1 %>
			</select>
		</td>
	</tr>
	<tr>
		<td><!--Blank--></td>
		<td>
			<input type="submit" name="function" value="Delete">
		</td>
	</tr>
</table>
		<%
	End Function

	Dim strChildrenIDs
	Function deleteCategory(intID)
		openDB
		
		Set dbCategory = cartDB.Execute("SELECT name FROM category WHERE id = " & intID & "")
		deleteCategory = dbCategory("name")
		Set dbCategory = Nothing
		
		strDeleteIDs = GetChildren(intID,intID)
		aDeleteIDs = Split(strDeleteIDs,";")

		For Each id In aDeleteIDs
			cartDB.Execute(SQL("UPDATE category SET isdeleted = true WHERE id = " & intID & ""))
			cartDB.Execute(SQL("UPDATE product SET deleted = true WHERE categoryid = " & intID & ""))
		Next
		
		closeDB
	End Function
	
	Function GetChildren(intID,rootID)
		If (Fix(intID) = Fix(rootID)) Then
			strChildrenIDs = rootID
		End If
		Set dbCategory = cartDB.Execute("SELECT count(id) as branchCount FROM category WHERE parentid = " & intID & " AND isdeleted = False")
		If (Fix(dbCategory("branchCount")) > 0) Then
			Set dbBranchData = cartDB.Execute("SELECT id FROM category WHERE parentid = " & intID & " AND isdeleted = False")
			Do Until dbBranchData.EOF
				strChildrenIDs = strChildrenIDs & ";" & dbBranchData("id")
				GetChildren dbBranchData("id"), rootID
				dbBranchData.MoveNext
			Loop
			Set dbBranchData = Nothing
		End If
		Set dbCategory = Nothing
		If (Fix(intID) = Fix(rootID)) Then
			GetChildren = strChildrenIDs
		End If
	End Function
		
' Category Tree
	Dim currentLevel

	Function CategoryTreeHandler
		debug("Entering CategoryTreeHandler handler.")
		currentLevel = -1
		print("<table border=""1"" cellpadding=""2"" cellspacing=""0"")")
		print("<tr bgcolor=""#000000"">")
		print("<td><font style=""color: white;"">Category Name</font></td>")
		print("<td><font style=""color: white;"">Product Count</font></td>")
		print("</tr>")
		printChildren -1, -1, "table", -1
		print("</table>")
	End Function

	Function printCategoryTreeOptions(intParentID)
		currentLevel = -1
		printChildren -1, -1, "option", intParentID
	End Function

	Function printCategoryTreeOptionsWithoutCurrent(intID, intParentID)
		currentLevel = -1
		printChildren intID, -1, "option", intParentID
	End Function

	Function printChildren(skipID,parentID,strType,intSelectedID)
		openDB
		Set dbBranchCount = cartDB.Execute("SELECT count(*) as branchCount FROM category WHERE parentid = " & parentID & " AND isdeleted = false")
		If (Fix(dbBranchCount("branchCount")) > 0) Then
			currentLevel = currentLevel + 1
'			Set dbBranchData = cartDB.Execute("SELECT category.id, category.name, count(product.id) as productcount FROM category, product WHERE category.id = product.categoryid AND parentid = " & parentID & " AND isdeleted = False GROUP BY category.name, category.id ORDER BY category.name")
			Set dbBranchData = cartDB.Execute("SELECT id, name FROM category WHERE parentid = " & parentID & " AND isdeleted = False ORDER BY name")

			Do Until dbBranchData.EOF
				Set dbProductCount = cartDB.Execute(SQL("SELECT count(id) as productCount FROM product WHERE categoryid = " & dbBranchData("id") & " AND deleted = false"))
				If Not ((Fix(skipID) <> -1) And (Fix(skipID) = Fix(dbBranchData("id")))) Then
					skipBranch = False
					thisColor = "style=""color:#000000"""
				Else
					skipBranch = True
					thisColor = "style=""color:#ffc0c0"""
				End If

				If Not (skipBranch) Then
					If (strType = "text") Then
						print(printSpacer(currentLevel) & dbBranchData("name") & "<br />")
					ElseIf (strType = "table") Then
						If (bgColor = "#f0f0f0") Then bgColor = "#ffffff" Else bgColor = "#f0f0f0"
						print("<tr bgcolor=""" & bgColor & """><td>" & printSpacer(currentLevel) & dbBranchData("name") & "</td><td>" & dbProductCount("productCount") & "</td></tr>")
					ElseIf (strType = "products") Then
						If (Fix(intSelectedID) = Fix(dbBranchData("id"))) Then
							thisSelected = "selected"
						Else
							thisSelected = ""
						End If
						thisColor = "style=""text-decoration: underline; color: #666666;"""
						print("<option value=""-1"" " & thisSelected & " " & thisColor & ">" & printSpacer(currentLevel) & dbBranchData("name") & " (" & dbProductCount("productCount") & " products)</option>")
					Else
						If (Fix(intSelectedID) = Fix(dbBranchData("id"))) Then
							thisSelected = "selected"
						Else
							thisSelected = ""
						End If
						print("<option value=""" & dbBranchData("id") & """ " & thisSelected & " " & thisColor & ">" & printSpacer(currentLevel) & dbBranchData("name") & "</option>")
					End If
				End If

				If Not (skipBranch) Then
					If (strType = "products") Then
						printProductOptionsFromParent intSelectedID, currentLevel, dbBranchData("id")
					End If
'					Set dbChildrenCount = cartDB.Execute("SELECT count(*) as childrenCount FROM category WHERE parentid = " & dbBranchData("id") & "")
'					If (Fix(dbChildrenCount("childrenCount")) > 0) Then
						printChildren skipID, dbBranchData("id"), strType, intSelectedID
'					End If
				End If
				dbBranchData.MoveNext
			Loop
			currentLevel = currentLevel - 1
		End If
		closeDB
	End Function
	
	Function printProductOptionsFromParent(intSelectedID, intLevel, intParentID)
		openDB
		Set dbProducts = cartDB.Execute(SQL("SELECT id, name FROM product WHERE categoryid = " & intParentID & " AND deleted = false"))
		If Not (dbProducts.EOF) Then
			Do Until dbProducts.EOF
				If (Fix(intSelectedID) = Fix(dbProducts("id"))) Then
					thisSelected = "selected"
				Else
					thisSelected = ""
				End If
				print("<option value=""" & dbProducts("id") & """ " & thisSelected & ">" & printSpacer(intLevel+1) & " * " & dbProducts("name") & "</option>")
				dbProducts.MoveNext
			Loop
		End If
		Set dbProducts = Nothing
		closeDB
	End Function

'Add Product
	Function thisError(strText)
		thisError = """ style=""background: #FF9999"" title=""" & strText & ""
	End Function

	Function AddProductHandler
		If (Request.Form("process") = "true") And (Request.Form("action") = "Save") Then
			intProductID = PrepareAddProduct("insert")
			strAddedProduct = "Successfully added '" & Request.Form("name") & "' as product id #" & intProductID & "."
			redirect Request.ServerVariables("URL") & "?message=" & strAddedProduct
		Else
			PrepareAddProduct "editor"
		End If
	End Function
	
	Function EditProductHandler
		If (Request.Form("process") = "true") Then
			If (Request.Form("action") = "Edit") Then
				If (Fix(Request.Form("productid")) > 0) Then
					PrepareEditProduct Fix(Request.Form("productid"))
				Else
					strSelectAProduct = "Please select a product to edit."
					redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&message=" & strSelectAProduct
				End If
			ElseIf (Request.Form("action") = "Save") Then
				PrepareAddProduct "update"
				strUpdated = "Saved product."
				redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&message=" & strUpdated
			Else
				PrepareAddProduct "editor"
			End If
		Else
			RenderProductSelect
		End If
	End Function
	
	Function DeleteProductHandler
		If (Request.Form("process") = "true") Then
			If (Request.Form("action") = "Delete") Then
				If (Fix(Request.Form("productid")) > 0) Then
					DeleteProduct Fix(Request.Form("productid"))
					strDeletedProduct = "Deleted product ID#" & Fix(Request.Form("productid"))
					redirect Request.ServerVariables("URL") & "?message=" & strDeletedProduct
				End If
			End If
			strFailedMessage = "Could not delete this product, you need to select a valid product."
			redirect Request.ServerVariables("URL") & "?message=" & strFailedMessage
		Else
			RenderDeleteProductSelect
		End If
	End Function
	
	Function RenderDeleteProductSelect
		%>
		Select the product you would like to delete.<br />
		<select name="productid" size="1">
			<option value="-1">... Select product to delete. ...</option>
			<% PrintProductOptions -1 %>
		</select><input type="submit" name="action" value="Delete">
		<%
	End Function
	
	Function DeleteProduct(intID)
		openDB
		cartDB.Execute("UPDATE product SET deleted = true WHERE id = " & intID & "")
		cartDB.Execute("UPDATE pricing SET isdeleted = true WHERE productid = " & intID & "")
		closeDB
	End Function
	
	Function RenderProductSelect
		%>
		Select the product you would like to edit.<br />
		<select name="productid" size="1">
			<option value="-1">... Select product to edit. ...</option>
			<% PrintProductOptions -1 %>
		</select><input type="submit" name="action" value="Edit">
		<%
	End Function
	
	Function PrintProductOptions(intSelectedID)
		currentLevel = -1
		printChildren -1, -1, "products", intSelectedID
	End Function

	Function PrepareEditProduct(intID)
		openDB
		Set dbProduct = cartDB.Execute(SQL("SELECT * FROM product WHERE id = " & intID & ""))
		If (dbProduct.EOF) Then
			strMessageFailed = "The product you tried to edit was not found in the database."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&message=" & strMessageFailed
		Else
			frmProductID = (dbProduct("id"))
			frmName = (dbProduct("name"))
			frmISBNUPC = (dbProduct("isbnupc"))
			frmSKU = (dbProduct("sku"))
			frmKeyword = (dbProduct("keyword"))
			frmDescription = (dbProduct("description"))
			frmThumbnail = (dbProduct("thumbnail"))
			frmImage = (dbProduct("image"))
			frmCategory = (dbProduct("categoryid"))
			frmHidden = (dbProduct("hidden"))
			If (frmHidden) Then
				frmHidden = " CHECKED"
			Else
				frmHidden = ""
			End If
					
			frmFulfill = (dbProduct("includeinfulfillment"))
			If (frmFulfill) Then
				frmFulfill = " CHECKED"
			Else
				frmFulfill = ""
			End If

			frmVendorID = Fix(dbProduct("vendorid"))
			
			If (isNull(dbProduct("vendorid"))) Then
				frmVendorID = 1
			End If
			
			Set dbVendor = cartDB.Execute(SQL("SELECT * FROM vendor WHERE id = " & frmVendorID & ""))
			frmCurrentVendor = " CHECKED"
			frmNewVendor = ""
			frmVendorName = (dbVendor("name"))
			frmVendorCompany = (dbVendor("company"))
			frmVendorEmail = (dbVendor("email"))
			frmVendorCC = (dbVendor("emailcc"))
			frmVendorBCC = (dbVendor("emailbcc"))
			Set dbVendor = Nothing
		
			Set dbPrice = cartDB.Execute(SQL("SELECT * FROM pricing WHERE productid = " & intID & " AND isnormal = true AND isdeleted = false"))
			frmNormalPrice = FormatNumber((dbPrice("price")),2)
			frmNormalPrice = frmNormalPrice
			Set dbPrice = Nothing
			
			Set dbPrice = cartDB.Execute(SQL("SELECT * FROM pricing WHERE productid = " & intID & " AND ispreorder = true AND isdeleted = false"))
			If Not (dbPrice.EOF) Then
'				Response.Write("PreOrderPrice: " & dbPrice("price") & "<br />")
'				Response.Write("Until: " & dbPrice("validuntil") & "<br />")
				frmPreorderPrice = FormatNumber((dbPrice("price")),2)
				frmPreorderUntil = (dbPrice("validuntil"))
			End If
			Set dbPrice = Nothing
			
			Set dbPrice = cartDB.Execute(SQL("SELECT count(id) as pricecount FROM pricing WHERE productid = " & intID & " AND isdiscount = true AND isdeleted = false"))
			frmPricingOptionCount = Fix(dbPrice("pricecount"))
			Set dbPrice = Nothing
			
'			Response.Write("Options: " & frmPricingOptionCount & "<br />")
			
			Dim frmPricing()
			ReDim frmPricing(frmPricingOptionCount - 1)
			Dim frmDiscountCode()
			ReDim frmDiscountCode(frmPricingOptionCount - 1)

			If (frmPricingOptionCount > 0) Then
				Set dbPrice = cartDB.Execute(SQL("SELECT * FROM pricing WHERE productid = " & intID & " AND isdiscount = true AND isdeleted = false"))
				frmPricingIndex = 0
				Do Until dbPrice.EOF
					thisPrice = (dbPrice("price"))
					frmPricing(frmPricingIndex) = FormatNumber(thisPrice,2)
					frmDiscountCode(frmPricingIndex) = (dbPrice("discountcode"))
'					Response.Write("Price: " & thisPrice & " Discount: " & dbPrice("discountcode") & " ID: " & dbPrice("id") & "<br />")

					frmPricingIndex = frmPricingIndex + 1
					dbPrice.MoveNext
				Loop
			End If
'			Response.End
			
	'		frmShippingOptionCount = makeSafe(Request.Form("shippingoptioncount"))
	'		If (isNumeric(frmShippingOptionCount)) Then
	'			If (Fix(frmShippingOptionCount) < 1) Then
	'				frmShippingOptionCount = 5
	'			End If
	'		Else
	'			frmShippingOptionCount = 5
	'		End If
			
	'		Dim frmQuantity()
	'		ReDim frmQuantity(frmShippingOptionCount)
	'		Dim frmGround()
	'		ReDim frmGround(frmShippingOptionCount)
	'		Dim frmTwoDay()
	'		ReDim frmTwoDay(frmShippingOptionCount)
	'		Dim frmOvernight()
	'		ReDim frmOvernight(frmShippingOptionCount)
	'		Dim frmCanadaMexico()
	'		ReDim frmCanadaMexico(frmShippingOptionCount)
	'		Dim frmInternational()
	'		ReDim frmInternational(frmShippingOptionCount)
			
	'		For frmShippingIndex = 1 To frmShippingOptionCount
	'			frmQuantity(frmShippingIndex) = makeSafe(Request.Form("quantity_" & frmShippingIndex))
	'			frmGround(frmShippingIndex) = makeSafe(Request.Form("ground_" & frmShippingIndex))
	'			frmTwoDay(frmShippingIndex) = makeSafe(Request.Form("twoday_" & frmShippingIndex))
	'			frmOvernight(frmShippingIndex) = makeSafe(Request.Form("overnight_" & frmShippingIndex))
	'			frmCanadaMexico(frmShippingIndex) = makeSafe(Request.Form("canadamexico_" & frmShippingIndex))
	'			frmInternational(frmShippingIndex) = makeSafe(Request.Form("international_" & frmShippingIndex))
	'		Next
	
			RenderProductEditor frmProductID, frmName, frmISBNUPC, frmSKU, frmKeyword, frmDescription, frmThumbnail, frmImage, frmCategory, frmHidden, frmVendor, frmVendorID, frmCurrentVendor, frmNewVendor, frmVendorName, frmVendorCompany, frmVendorEmail, frmVendorCC, frmVendorBCC, frmPricingOptionCount, frmNormalPrice, frmPreorderPrice, frmPreorderUntil, frmPricing, frmDiscountCode, frmFulfill
		End If
	End Function

	Function PrepareAddProduct(strAction)
		frmName = makeSafe(Request.Form("name"))
		frmISBNUPC = makeSafe(Request.Form("isbnupc"))
		frmSKU = makeSafe(Request.Form("sku")) 
		frmKeyword = makeSafe(Request.Form("keyword"))
		frmDescription = makeSafe(Request.Form("description"))
		frmThumbnail = makeSafe(Request.Form("thumbnail"))
		frmImage = makeSafe(Request.Form("image"))
		frmCategory = Fix(Request.Form("categoryid"))
		frmHidden = makeSafe(Request.Form("hidden"))
		If (strAction = "editor") Then
			If (frmHidden = "true") Then
				frmHidden = " CHECKED"
			Else
				frmHidden = ""
			End If
		Else
			If (Len(frmHidden) > 0) Then
				frmHidden = "true"
			Else
				frmHidden = "false"
			End If
		End If
				
		frmFulfill = Request.Form("fulfill")
		If (Len(frmFulfill) > 0) Then
			frmFulfill = "true"
		Else
			frmFulfill = "false"
		End If
		
		frmVendor = makeSafe(Request.Form("vendor"))
		If (frmVendor = "current") Then
			frmVendorID = Fix(Request.Form("vendorid"))
			frmCurrentVendor = " CHECKED"
			frmNewVendor = ""
			frmVendorName = ""
			frmVendorCompany = ""
			frmVendorEmail = ""
			frmVendorCC = ""
			frmVendorBCC = ""
		ElseIf (frmVendor = "new") Then
			frmVendorID = -1
			frmCurrentVendor = ""
			frmNewVendor = " CHECKED"
			frmVendorName = makeSafe(Request.Form("vendorname"))
			frmVendorCompany = makeSafe(Request.Form("vendorcompany"))
			frmVendorEmail = makeSafe(Request.Form("vendoremail"))
			frmVendorCC = makeSafe(Request.Form("vendorcc"))
			frmVendorBCC = makeSafe(Request.Form("vendorbcc"))
		End If

		frmNormalPrice = makeSafe(Request.Form("normalprice"))
		If (isNull(frmNormalPrice)) Or (isEmpty(frmNormalPrice)) Or (Len(frmNormalPrice) = 0) Or Not (isNumeric(frmNormalPrice)) Then
			frmNormalPrice = frmNormalPrice
		Else
			frmNormalPrice = FormatNumber(frmNormalPrice,2)
		End If

		frmPreorderPrice = makeSafe(Request.Form("preorderprice"))
		If (isNull(frmPreorderPrice)) Or (isEmpty(frmPreorderPrice)) Or (Len(frmPreorderPrice) = 0) Or Not (isNumeric(frmPreorderPrice)) Then
			frmPreorderPrice = frmPreorderPrice
		Else
			frmPreorderPrice = FormatNumber(frmPreorderPrice,2)
		End If
		
		frmPreorderUntil = makeSafe(Request.Form("preorderuntil"))

'			Set dbPrice = cartDB.Execute(SQL("SELECT count(id) as pricecount FROM pricing WHERE productid = " & intID & " AND isdiscount = true AND isdeleted = false"))
'			frmPricingOptionCount = Fix(dbPrice("pricecount"))
'			Set dbPrice = Nothing
			
		frmPricingOptionCount = Fix(Request.Form("pricingoptioncount"))
		If Not (isNumeric(frmPricingOptionCount)) Then
			frmPricingOptionCount = 1
		End If
		
		Dim frmPricing()
		ReDim frmPricing(frmPricingOptionCount-1)
		Dim frmDiscountCode()
		ReDim frmDiscountCode(frmPricingOptionCount-1)
		
		For frmPricingIndex = 0 To frmPricingOptionCount-1
			thisPrice = makeSafe(Request.Form("price_" & frmPricingIndex))
			If (isNumeric(thisPrice)) Then
				frmPricing(frmPricingIndex) = FormatNumber(thisPrice,2)
			Else
				frmPricing(frmPricingIndex) = thisPrice
			End If
			frmDiscountCode(frmPricingIndex) = makeSafe(Request.Form("discount_" & frmPricingIndex))
		Next

'		frmShippingOptionCount = makeSafe(Request.Form("shippingoptioncount"))
'		If (isNumeric(frmShippingOptionCount)) Then
'			If (Fix(frmShippingOptionCount) < 1) Then
'				frmShippingOptionCount = 5
'			End If
'		Else
'			frmShippingOptionCount = 5
'		End If
		
'		Dim frmQuantity()
'		ReDim frmQuantity(frmShippingOptionCount)
'		Dim frmGround()
'		ReDim frmGround(frmShippingOptionCount)
'		Dim frmTwoDay()
'		ReDim frmTwoDay(frmShippingOptionCount)
'		Dim frmOvernight()
'		ReDim frmOvernight(frmShippingOptionCount)
'		Dim frmCanadaMexico()
'		ReDim frmCanadaMexico(frmShippingOptionCount)
'		Dim frmInternational()
'		ReDim frmInternational(frmShippingOptionCount)
		
'		For frmShippingIndex = 1 To frmShippingOptionCount
'			frmQuantity(frmShippingIndex) = makeSafe(Request.Form("quantity_" & frmShippingIndex))
'			frmGround(frmShippingIndex) = makeSafe(Request.Form("ground_" & frmShippingIndex))
'			frmTwoDay(frmShippingIndex) = makeSafe(Request.Form("twoday_" & frmShippingIndex))
'			frmOvernight(frmShippingIndex) = makeSafe(Request.Form("overnight_" & frmShippingIndex))
'			frmCanadaMexico(frmShippingIndex) = makeSafe(Request.Form("canadamexico_" & frmShippingIndex))
'			frmInternational(frmShippingIndex) = makeSafe(Request.Form("international_" & frmShippingIndex))
'		Next

		If (strAction = "insert") Then
			frmProductID = -1
		Else
			frmProductID = makeSafe(Request.Form("productid"))
		End If

		If (strAction = "editor") Then
			RenderProductEditor frmProductID, frmName, frmISBNUPC, frmSKU, frmKeyword, frmDescription, frmThumbnail, frmImage, frmCategory, frmHidden, frmVendor, frmVendorID, frmCurrentVendor, frmNewVendor, frmVendorName, frmVendorCompany, frmVendorEmail, frmVendorCC, frmVendorBCC, frmPricingOptionCount, frmNormalPrice, frmPreorderPrice, frmPreorderUntil, frmPricing, frmDiscountCode, frmFulfill
		ElseIf (strAction = "update") Or (strAction = "insert") Then
			PrepareAddProduct = AddOrUpdateProduct(strAction, frmProductID, frmName, frmISBNUPC, frmSKU, frmKeyword, frmDescription, frmThumbnail, frmImage, frmCategory, frmHidden, frmVendor, frmVendorID, frmCurrentVendor, frmNewVendor, frmVendorName, frmVendorCompany, frmVendorEmail, frmVendorCC, frmVendorBCC, frmPricingOptionCount, frmNormalPrice, frmPreorderPrice, frmPreorderUntil, frmPricing, frmDiscountCode, frmFulfill)
		End If
	End Function
	
	Function AddOrUpdateProduct(_
		strAction, _
		frmProductID, _
		frmName, _
		frmISBNUPC, _
		frmSKU, _
		frmKeyword, _
		frmDescription, _
		frmThumbnail, _
		frmImage, _
		frmCategory, _
		frmHidden, _
		frmVendor, _
		frmVendorID, _
		frmCurrentVendor, _
		frmNewVendor, _
		frmVendorName, _
		frmVendorCompany, _
		frmVendorEmail, _
		frmVendorCC, _
		frmVendorBCC, _
		frmPricingOptionCount, _
		frmNormalPrice, _
		frmPreorderPrice, _
		frmPreorderUntil, _
		frmPricing, _
		frmDiscountCode, _
		frmFulfill)

		openDB
		
		Set oRecord = Server.CreateObject("ADODB.RecordSet")

		If (strAction = "update") Then
			If (Fix(frmProductID) > 0) Then
				oRecord.Open "SELECT * FROM product WHERE id = " & frmProductID & "", cartDB, 2, 3
'				strProductSQL = "UPDATE product SET name = '" & frmName & "', isbnupc = '" & frmISBNUPC & "', sku = '" & frmSKU & "', keyword = '" & frmKeyword & "', description = '" & frmDescription & "', thumbnail = '" & frmThumbnail & "', categoryid = " & frmCategory & ", modified = #" & Date & " " & Time & "#, hidden = " & frmHidden & ", vendorid = " & frmVendorID & " WHERE id = " & frmProductID & ""
			Else
				ReportError "Tried to update a product id of -1."
			End If
		Else
			oRecord.Open "SELECT * FROM product", cartDB, 2, 3
			oRecord.AddNew
			oRecord.Fields("added") = Date & " " & Time
'			strProductSQL = "INSERT INTO product ( name, isbnupc, sku, keyword, description, thumbnail, categoryid, added, modified, hidden, vendorid ) VALUES ( '" & frmName & "', '" & frmISBNUPC & "', '" & frmSKU & "', '" & frmKeyword & "', '" & frmDescription & "', '" & frmThumbnail & "', " & frmCategory & ", #" & Date & " " & Time & "#, #" & Date & " " & Time & "#, " & frmHidden & ", " & frmVendorID & " )"
		End If
	
		oRecord.Fields("name") = frmName
		oRecord.Fields("isbnupc") = frmISBNUPC
		oRecord.Fields("sku") = frmSKU
		oRecord.Fields("keyword") = frmKeyword
		oRecord.Fields("description") = frmDescription
		oRecord.Fields("thumbnail") = frmThumbnail
		oRecord.Fields("categoryid") = frmCategory
		oRecord.Fields("modified") = Date & " " & Time
		oRecord.Fields("hidden") = frmHidden
		oRecord.Fields("vendorid") = frmVendorID
		oRecord.Fields("includeinfulfillment") = frmFulfill
				
		oRecord.Update
'		Response.Write("strProductSQL:<br /><small><small><b>" & strProductSQL & "</b></small></small><br />")	
'		cartDB.Execute(SQL(strProductSQL))

		cycleDatabase

		If (strAction = "insert") Then
			Set dbProduct = cartDB.Execute("SELECT max(id) as lastID FROM product WHERE name = '" & frmName & "' AND keyword = '" & frmKeyword & "'")
			frmProductID = dbProduct("lastID")
			Set dbProduct = Nothing
		End If

'		strDeleteNormalSQL = "DELETE FROM pricing WHERE productid = " & frmProductID & " AND isnormal = true"
		strDeleteNormalSQL = "UPDATE pricing SET isdeleted = true WHERE productid = " & frmProductID & " AND isnormal = true"
		strInsertNormalSQL = "INSERT INTO pricing ( productid, price, isnormal ) VALUES ( " & frmProductID & ", " & ParseFloat(frmNormalPrice) & ", true )"

'		Response.Write("strDeleteNormalSQL:<br /><small><small><b>" & strDeleteNormalSQL & "</b></small></small><br />")	
		cartDB.Execute(SQL(strDeleteNormalSQL))
'		Response.Write("strInsertNormalSQL:<br /><small><small><b>" & strInsertNormalSQL & "</b></small></small><br />")	
		cartDB.Execute(SQL(strInsertNormalSQL))

'		strDeletePreorderSQL = "DELETE FROM pricing WHERE productid = " & frmProductID & " AND ispreorder = true"
		strDeletePreorderSQL = "UPDATE pricing SET isdeleted = true WHERE productid = " & frmProductID & " AND ispreorder = true"
		If (Len(frmPreorderUntil) > 0) Then
			strInsertPreorderSQL = "INSERT INTO pricing ( productid, price, validuntil, ispreorder ) VALUES ( " & frmProductID & ", " & ParseFloat(frmPreorderPrice) & ", '" & frmPreorderUntil & "', true )"
			doPreorder = True
		Else
			doPreorder = False
		End If

'		Response.Write("strDeletePreorderSQL:<br /><small><small><b>" & strDeletePreorderSQL & "</b></small></small><br />")	
		cartDB.Execute(SQL(strDeletePreorderSQL))
		If (doPreorder) Then
'			Response.Write("strInsertPreorderSQL:<br /><small><small><b>" & strInsertPreorderSQL & "</b></small></small><br />")
			cartDB.Execute(SQL(strInsertPreorderSQL))
		End If

'		strDeleteDiscountSQL = "DELETE FROM pricing WHERE productid = " & frmProductID & " AND isdiscount = true"
		strDeleteDiscountSQL = "UPDATE pricing SET isdeleted = true WHERE productid = " & frmProductID & " AND isdiscount = true"
		If (Fix(frmPricingOptionCount) > 0) Then
			Dim strInsertDiscountSQL()
			ReDim strInsertDiscountSQL(frmPricingOptionCount)
			For X = 0 To frmPricingOptionCount - 1
				If (Len(frmDiscountCode(X)) > 0) And (Len(frmPricing(X)) > 0) Then
					strInsertDiscountSQL(X) = "INSERT INTO pricing ( productid, price, discountcode, isdiscount ) VALUES ( " & frmProductID & ", " & ParseFloat(frmPricing(X)) & ", '" & frmDiscountCode(X) & "', true )"
				Else
					strInsertDiscountSQL(X) = ""
				End If
			Next
			doDiscounts = True
		Else
			doDiscounts = False
		End If

'		Response.Write("strDeleteDiscountSQL:<br /><small><small><b>" & strDeleteDiscountSQL & "</b></small></small><br />")	
		cartDB.Execute(SQL(strDeleteDiscountSQL))

		If (doDiscounts) Then
			For X = 0 To Ubound(strInsertDiscountSQL)
'				Response.Write("strInsertDiscountSQL:<br /><small><small><b>" & strInsertDiscountSQL(X) & "</b></small></small><br />")
				If (Len(strInsertDiscountSQL(X)) > 0) Then
					cartDB.Execute(SQL(strInsertDiscountSQL(X)))
				End If
			Next
		End If
		
		cycleDatabase
		AddOrUpdateProduct = frmProductID
		closeDB
	End Function
	
	Function RenderProductEditor(frmProductID, frmName, frmISBNUPC, frmSKU, frmKeyword, frmDescription, frmThumbnail, frmImage, frmCategory, frmHidden, frmVendor, frmVendorID, frmCurrentVendor, frmNewVendor, frmVendorName, frmVendorCompany, frmVendorEmail, frmVendorCC, frmVendorBCC, frmPricingOptionCount, frmNormalPrice, frmPreorderPrice, frmPreorderUntil, frmPricing, frmDiscountCode, frmFulfill)
		%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse;">
	<tr bgcolor="#000000">
		<td colspan="3" background="/images/banner_back.gif">&nbsp;&nbsp;<font style="color: white;"><b>Product Description</b></font></td>
	</tr>
	<tr>
		<td align="right" nowrap>Name:</td>
		<td><input type="text" name="name" size="60" value="<%=frmName%>"></td>
		<td><font size="-1">Product title.</td>
	</tr>
	<tr>
		<td align="right" nowrap>ISBN/UPC:</td>
		<td><input type="text" name="isbnupc" size="60" value="<%=frmISBNUPC%>"></td>
		<td><font size="-1">ISBN for books, UPC for other products.</td>
	</tr>
	<tr>
		<td align="right" nowrap>SKU:</td>
		<td><input type="text" name="sku" size="60" value="<%=frmSKU%>"></td>
		<td><font size="-1">Product SKU from barcode.</td>
	</tr>
	<tr>
		<td align="right" nowrap>Keyword:</td>
		<td><input type="text" name="keyword" size="60" value="<%=frmKeyword%>"></td>
		<td><font size="-1">Quick search keyword.</td>
	</tr>
	<tr>
		<td align="right" nowrap>Description:</td>
		<td colspan=""2"">
			<textarea cols="60" rows="7" name="description"><%=frmDescription%></textarea><br />
		</td>
		<td><font size="-1">Enter the product description, can include HTML code.</td>
	</tr>
	<tr>
		<td align="right" nowrap>Thumbnail:</td>
		<td><input type="text" name="thumbnail" size="60" value="<%=frmThumbnail%>"></td>
		<td><font size="-1">Thumbnail image URL.</td>
	</tr>
	<tr bgcolor="#000000">
		<td colspan="3" background="/images/banner_back.gif">&nbsp;&nbsp;<font style="color: white;"><b>Product Attributes</b></font></td>
	</tr>
	<tr>
		<td align="right" nowrap>Category:</td>
		<td>
			<select size="1" name="categoryid">
				<option value="-1"> ... Select a category. ... </option>
				<% printCategoryTreeOptions frmCategory %>
			</select>
		</td>
		<td><font size="-1">Product will show up in this category.</td>
	</tr>
	<tr>
		<td align="right"><input type="checkbox" id="hiddencheck" name="hidden" value="true" <%=frmHidden%>></td>
		<td nowrap><label for="hiddencheck">Product is hidden.</label></td>
		<td><font size="-1">Product will be hidden when added.</td>
	</tr>
	<tr bgcolor="#000000">
		<td colspan="3" background="/images/banner_back.gif">&nbsp;&nbsp;<font style="color: white;"><b>Product Fulfillment Vendor</b></font></td>
	</tr>
	<tr>
		<td align="right"><input type="checkbox" id="fulfillcheck" name="fulfill" value="true" <%=frmFulfill%>></td>
		<td nowrap><label for="fulfillcheck">Included in fulfillment.</label></td>
		<td><font size="-1">Product will be included in automated fulfillment spreadsheets.</td>
	</tr>
	<tr>
		<td align="right"><input type="radio" id="currentvendor" name="vendor" value="current" <%=frmCurrentVendor%>></td>
		<td nowrap><label for="currentvendor">Use existing vendor.</label></td>
		<td><font size="-1">Use an existing vendor for fullfillment.</td>
	</tr>
	<tr>
		<td align="right" nowrap>Existing Vendor:</td>
		<td>
			<select size="1" name="vendorid">
				<option value="-1"> ... Select Fulfillment Vendor ... </option>
				<% printVendorOptions frmVendorID %>
			</select>
		</td>
		<td><font size="-1">Product will be to sent to this vendor.</td>
	</tr>
	<tr>
		<td align="right"><input type="radio" id="newvendor" name="vendor" value="new" <%=frmNewVendor%>></td>
		<td nowrap><label for="newvendor">Use a new vendor. (Enter info below.)</label></td>
		<td><font size="-1">Use a new vendor for fullfillment.</td>
	</tr>
	<tr>
		<td align="right" nowrap>Name:</td>
		<td><input type="text" name="vendorname" size="60" value="<%=frmVendorName%>"></td>
		<td><font size="-1">Contact name for vendor.</td>
	</tr>
	<tr>
		<td align="right" nowrap>Company:</td>
		<td><input type="text" name="vendorcompany" size="60" value="<%=frmVendorCompany%>"></td>
		<td><font size="-1">Vendor company.</td>
	</tr>
	<tr>
		<td align="right" nowrap>To:</td>
		<td><input type="text" name="vendoremail" size="60" value="<%=frmVendorEmail%>"></td>
		<td><font size="-1">To addresses, seperate with semi-color (;).</td>
	</tr>
	<tr>
		<td align="right" nowrap>CC:</td>
		<td><input type="text" name="vendorcc" size="60" value="<%=frmVendorCC%>"></td>
		<td><font size="-1">CC addresses.</td>
	</tr>
	<tr>
		<td align="right" nowrap>BCC:</td>
		<td><input type="text" name="vendorbcc" size="60" value="<%=frmVendorBCC%>"></td>
		<td><font size="-1">BCC addresses.</td>
	</tr>
	<tr bgcolor="#000000">
		<td colspan="3" background="/images/banner_back.gif">&nbsp;&nbsp;<font style="color: white;"><b>Product Pricing</b></font></td>
	</tr>
	<tr>
		<td valign="top" colspan="2">
			<table border="0" cellpadding="2" cellspacing="0" width="100%">
				<tr>
					<td background="/images/banner_back.gif" colspan="3">&nbsp;<font style="color: white;"><b>Normal Pricing</b></font></td>
				</tr>
				<tr>
					<td nowrap width="100">Normal</td>
					<td nowrap width="70">$<input type="text" name="normalprice" size="6" value="<%=frmNormalPrice%>"></td>
					<td width="90%">(Required)</td>
				</tr>
				<tr>
					<td nowrap width="100">Pre-Order</td>
					<td nowrap width="70">$<input type="text" name="preorderprice" size="6" value="<%=frmPreorderPrice%>"></td>
					<td width="90%"><input type="text" name="preorderuntil" size="10" value="<%=frmPreorderUntil%>">(Optional)</td>
				</tr>
				<tr>
					<td background="/images/banner_back.gif" colspan="3">&nbsp;<font style="color: white;"><b>Discount Pricing</b></font></td>
				</tr>
				<tr>
					<td colspan="3" nowrap>Discount Pricing Options:&nbsp;<input type="text" name="pricingoptioncount" size="5" value="<%=frmPricingOptionCount%>"><input type="submit" name="action" value="Update"></td>
				</tr>
				<tr>
					<td background="/images/banner_back.gif">&nbsp;<font style="color: white;"><b>Type</b></font></td>
					<td background="/images/banner_back.gif">&nbsp;<font style="color: white;"><b>Price</b></font></td>
					<td background="/images/banner_back.gif">&nbsp;<font style="color: white;"><b>Discount Code</b></font></td>
				</tr>
				<% For priceIndex = 0 To frmPricingOptionCount - 1 %>
				<tr>
					<td nowrap width="100">Discount</td>
					<td nowrap width="70">$<input type="text" name="price_<%=priceIndex%>" size="6" value="<%=frmPricing(priceIndex)%>"></td>
					<td width="90%"><input type="text" name="discount_<%=priceIndex%>" size="50" value="<%=frmDiscountCode(priceIndex)%>"> (Optional)</td>
				</tr>
				<% Next %>
			</table>
		</td>
		<td>
			<font size="-1">Update the Pricing Options Count number to allow more pricing options.<br />
			<br />
			Only one price can be set as the normal price.  All other prices must have a discount code associated with it.<br />
			<br />
			<br /><b>To remove a price:</b>  Leave the price and discount code entries blank.</font>
		</td>
	</tr>
<!--	<tr bgcolor="#000000">
		<td colspan="3" background="/images/banner_back.gif">&nbsp;&nbsp;<font style="color: white;"><b>Product Shipping</b></font></td>
	</tr>
	<tr>
		<td align="right" nowrap>Shipping Option Count:</td>
		<td><input type="text" name="shippingoptioncount" size="5" value="<%'=frmShippingOptionCount%>"><input type="submit" name="action" value="Update"></td>
		<td rowspan="2"><font size="-1">Update the Pricing Options Count number to allow more pricing options.<br />
			<br />
			Only one price can be set as the normal price.  All other prices must have a discount code associated with it.<br />
			<br />
			<br /><b>To remove a price:</b>  Leave the price and discount code entries blank.</font>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<table border="0" cellpadding="2" cellspacing="0" width="100%">
				<tr>
					<td background="/images/banner_back.gif">&nbsp;<font style="color: white;"><b>Quantity</b></font></td>
					<td background="/images/banner_back.gif">&nbsp;<font style="color: white;"><b>Ground</b></font></td>
					<td background="/images/banner_back.gif">&nbsp;<font style="color: white;"><b>2nd Day</b></font></td>
					<td background="/images/banner_back.gif">&nbsp;<font style="color: white;"><b>Overnight</b></font></td>
					<td background="/images/banner_back.gif">&nbsp;<font style="color: white;"><b>Canada/Mex.</b></font></td>
					<td background="/images/banner_back.gif">&nbsp;<font style="color: white;"><b>Intl.</b></font></td>
				</tr>
				<% 'For shippingIndex = 1 To frmShippingOptionCount %>
				<tr>
					<td width="17%"><input type="text" name="quantity_<%'=shippingIndex%>" size="6" value="<%'=frmQuantity(shippingIndex)%>"></td>
					<td width="17%">$<input type="text" name="ground_<%'=shippingIndex%>" size="6" value="<%'=frmGround(shippingIndex)%>"></td>
					<td width="17%">$<input type="text" name="twoday_<%'=shippingIndex%>" size="6" value="<%'=frmTwoDay(shippingIndex)%>"></td>
					<td width="17%">$<input type="text" name="overnight_<%'=shippingIndex%>" size="6" value="<%'=frmOvernight(shippingIndex)%>"></td>
					<td width="17%">$<input type="text" name="canadamexico_<%'=shippingIndex%>" size="6" value="<%'=frmCanadaMexico(shippingIndex)%>"></td>
					<td width="17%">$<input type="text" name="international_<%'=shippingIndex%>" size="6" value="<%'=frmInternational(shippingIndex)%>"></td>
				</tr>
				<% 'Next %>
			</table>
		</td>
	</tr>
-->
	<tr>
		<td>&nbsp;</td>
		<td>
			<input type="hidden" name="productid" value="<%=frmProductID%>">
			<input type="submit" name="action" value="Save"><input type="reset" value="Reset">
		</td>
		<td>&nbsp;</td>
	</tr>
</table>
		<%
'		For Each frm In Request.Form
'			Response.Write("<b>" & frm & "</b>=" & Request.Form(frm) & "<br />")
'		Next
	End Function
	
	Function printVendorOptions(intSelectedID)
		openDB
		printVendors "option", intSelectedID
		closeDB
	End Function
	
	Function printVendors(strType,intSelectedID)
		Set dbVendors = cartDB.Execute("SELECT id, name, company FROM vendor ORDER BY company, name")

		Do Until dbVendors.EOF
			If (strType = "text") Then
				print(dbVendors("name") & "<br />")
			Else
				If (Fix(intSelectedID) = Fix(dbVendors("id"))) Then
					thisSelected = "selected"
				Else
					thisSelected = ""
				End If
				print("<option value=""" & dbVendors("id") & """ " & thisSelected & " " & thisColor & ">" & dbVendors("company") & " c/o " & dbVendors("name") & "</option>")
			End If
			dbVendors.MoveNext
		Loop
		Set dbVendors = Nothing
	End Function

' View Orders
	lblPaidCompleted = "Paid/Completed"
	lblClearPaid = "Clear Paid"
	lblPaid = "Mark Paid"
	lblClearCompleted = "Clear Completed"
	lblCompleted = "Mark Completed"
	lblDelete = "Delete/Cancel"
	lblEmailCustomerReceipt = "Customer Receipt"
	lblEmailMerchantReceipt = "Merchant Receipt"
	lblEmailPaid = "Customer Paid"
	lblEmailFulfillment = "Fulfillment Request"
	
	lblEditCustomerInfo = "Edit Customer"
	lblSaveCustomerInfo = "Save Customer"
	
	Function ViewOrdersHandler
		If (Len(Request.QueryString("dateStart")) > 0) Then
			dateStart = makeSafe(Request.QueryString("dateStart"))
		ElseIf (Len(Request.Form("dateStart")) > 0) Then
			dateStart = makeSafe(Request.Form("dateStart"))
		Else
			dateStart = DateAdd("d",-6,Date)
'			dateStart = CDate(DatePart("m",Date) & "/01/" & DatePart("yyyy",Date))
		End If

		If (Len(Request.QueryString("dateEnd")) > 0) Then
			dateEnd = makeSafe(Request.QueryString("dateEnd"))
		ElseIf (Len(Request.Form("dateEnd")) > 0) Then
			dateEnd = makeSafe(Request.Form("dateEnd"))
		Else
			dateEnd = Date
'			dateEnd = DateAdd("d",-1,DateAdd("m",1,CDate(DatePart("m",Date) & "/01/" & DatePart("yyyy",Date))))
		End If

		strOrderID = makeSafe(Request.QueryString("orderid"))
		strAction = makeSafe(Request.Form("action"))
		strActionID = makeSafe(Request.Form("id"))

		frmordersession_id = makeSafe(Request.Form("ordersession_id"))
		frmordersession_paymentmethod = makeSafe(Request.Form("ordersession_paymentmethod"))
		frmordersession_shippingmethod = makeSafe(Request.Form("ordersession_shippingmethod"))
		frmorderbilling_id = makeSafe(Request.Form("orderbilling_id"))
		frmorderbilling_name = makeSafe(Request.Form("orderbilling_name"))
		frmorderbilling_company = makeSafe(Request.Form("orderbilling_company"))
		frmorderbilling_title = makeSafe(Request.Form("orderbilling_title"))
		frmorderbilling_address = makeSafe(Request.Form("orderbilling_address"))
		frmorderbilling_address2 = makeSafe(Request.Form("orderbilling_address2"))
		frmorderbilling_city = makeSafe(Request.Form("orderbilling_city"))
		frmorderbilling_state = makeSafe(Request.Form("orderbilling_state"))
		frmorderbilling_zip = makeSafe(Request.Form("orderbilling_zip"))
		frmorderbilling_country = makeSafe(Request.Form("orderbilling_country"))
		frmorderbilling_phonenumber = makeSafe(Request.Form("orderbilling_phonenumber"))
		frmorderbilling_emailaddress = makeSafe(Request.Form("orderbilling_emailaddress"))
		frmordershipping_id = makeSafe(Request.Form("ordershipping_id"))
		frmordershipping_name = makeSafe(Request.Form("ordershipping_name"))
		frmordershipping_company = makeSafe(Request.Form("ordershipping_company"))
		frmordershipping_title = makeSafe(Request.Form("ordershipping_title"))
		frmordershipping_address = makeSafe(Request.Form("ordershipping_address"))
		frmordershipping_address2 = makeSafe(Request.Form("ordershipping_address2"))
		frmordershipping_city = makeSafe(Request.Form("ordershipping_city"))
		frmordershipping_state = makeSafe(Request.Form("ordershipping_state"))
		frmordershipping_zip = makeSafe(Request.Form("ordershipping_zip"))
		frmordershipping_country = makeSafe(Request.Form("ordershipping_country"))
		
		If (strAction = lblPaidCompleted) Then
			MarkOrderCompleted ParseInt(strActionID)
			MarkOrderPaid ParseInt(strActionID)
			CompleteOrder ParseInt(strActionID)
			strPaidCompletedMessage = "This order is now completed.  Paid receipt and fulfillment emails were sent."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&orderid=" & strActionID & "&message=" & strPaidCompletedMessage
		ElseIf (strAction = lblEmailCustomerReceipt) Then
			SendCustomerReceipt ParseInt(strActionID)
			strMessage = "Customer receipt email was sent."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&orderid=" & strActionID & "&message=" & strMessage
		ElseIf (strAction = lblEmailMerchantReceipt) Then
			SendMerchantReceipt ParseInt(strActionID)
			strMessage = "Merchant receipt email was sent."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&orderid=" & strActionID & "&message=" & strMessage
		ElseIf (strAction = lblEmailPaid) Then
			SendCustomerPaid ParseInt(strActionID)
			strMessage = "Customer paid email was sent."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&orderid=" & strActionID & "&message=" & strMessage
		ElseIf (strAction = lblEmailFulfillment) Then
			SendFulfillmentRequest ParseInt(strActionID)
			strMessage = "Fulfillment request email was sent."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&orderid=" & strActionID & "&message=" & strMessage
		ElseIf (strAction = lblPaid) Then
			MarkOrderPaid ParseInt(strActionID)
			strMessage = "Order is now marked as paid."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&orderid=" & strActionID & "&message=" & strMessage
		ElseIf (strAction = lblClearPaid) Then
			ClearOrderPaid ParseInt(strActionID)
			strMessage = "Cleared paid flag on order."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&orderid=" & strActionID & "&message=" & strMessage
		ElseIf (strAction = lblCompleted) Then
			MarkOrderCompleted ParseInt(strActionID)
			strMessage = "Order is now marked as completed."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&orderid=" & strActionID & "&message=" & strMessage
		ElseIf (strAction = lblClearCompleted) Then
			ClearOrderCompleted ParseInt(strActionID)
			strMessage = "Cleared completed flag on order."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&orderid=" & strActionID & "&message=" & strMessage
		ElseIf (strAction = lblDelete) Then
			DeleteOrder strActionID
			strDeleted = "Deleted order ID " & strActionID & "."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&message=" & strDeleted & "&dateStart=" & dateStart & "&dateEnd=" & dateEnd
		ElseIf (strAction = lblSaveCustomerInfo) Then
			'SaveCustomerInfo Request.Form
			UpdateCustomer frmordersession_id, frmordersession_paymentmethod, frmordersession_shippingmethod, frmorderbilling_id, frmorderbilling_name, frmorderbilling_company, frmorderbilling_title, frmorderbilling_address, frmorderbilling_address2, frmorderbilling_city, frmorderbilling_state, frmorderbilling_zip, frmorderbilling_country, frmorderbilling_phonenumber, frmorderbilling_emailaddress, frmordershipping_id, frmordershipping_name, frmordershipping_company, frmordershipping_title, frmordershipping_address, frmordershipping_address2, frmordershipping_city, frmordershipping_state, frmordershipping_zip, frmordershipping_country
			strSaved = "Customer's information was updated."
'			For Each frm In Request.Form
'				Response.Write("rs(""" & Replace(Replace(Replace(Replace(frm,"frm",""),"ordersession_",""),"orderbilling_",""),"ordershipping_","") & """) = frm" & frm & " '" & Request.Form(frm) & "<br />")
'			Next
'			Response.End
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&orderid=" & strActionID & "&message=" & strSaved
		ElseIf (Len(strOrderID) > 0) Then
			RenderOrderReview strOrderID
		Else			
			RenderOrderList dateStart, dateEnd
		End If
	End Function
	
	Function UpdateCustomer(frmordersession_id, frmordersession_paymentmethod, frmordersession_shippingmethod, frmorderbilling_id, frmorderbilling_name, frmorderbilling_company, frmorderbilling_title, frmorderbilling_address, frmorderbilling_address2, frmorderbilling_city, frmorderbilling_state, frmorderbilling_zip, frmorderbilling_country, frmorderbilling_phonenumber, frmorderbilling_emailaddress, frmordershipping_id, frmordershipping_name, frmordershipping_company, frmordershipping_title, frmordershipping_address, frmordershipping_address2, frmordershipping_city, frmordershipping_state, frmordershipping_zip, frmordershipping_country)
		openDB
		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.Open SQL("SELECT * FROM ordersession WHERE id = " & frmordersession_id & ""), cartDB, 2, 3
		rs.MoveFirst
		'rs("id") = frmordersession_id '521
		rs.Fields("paymentmethod") = frmordersession_paymentmethod 'creditdebit
		rs.Fields("shippingmethod") = frmordersession_shippingmethod 'ground
		rs.Update
		rs.Close
		
		rs.Open SQL("SELECT * FROM orderbilling WHERE id = " & frmorderbilling_id & ""), cartDB, 2, 3
		rs.MoveFirst
		'rs("id") = frmorderbilling_id '517
		rs.Fields("name") = frmorderbilling_name 'Stephen Reses
		rs.Fields("company") = frmorderbilling_company 'Stephen Reses Productions
		rs.Fields("title") = frmorderbilling_title 'President
		rs.Fields("address") = frmorderbilling_address '113 Parkwood Place
		rs.Fields("address2") = frmorderbilling_address2 '
		rs.Fields("city") = frmorderbilling_city 'Linwood
		rs.Fields("state") = frmorderbilling_state 'New Jersey
		rs.Fields("zip") = frmorderbilling_zip '08221
		rs.Fields("country") = frmorderbilling_country 'United States of America
		rs.Fields("phonenumber") = frmorderbilling_phonenumber '609-926-7777
		rs.Fields("emailaddress") = frmorderbilling_emailaddress 'kreses@aol.com
		rs.Update
		rs.Close

		rs.Open SQL("SELECT * FROM ordershipping WHERE id = " & frmordershipping_id & ""), cartDB, 2, 3
		rs.MoveFirst
		'rs("id") = frmordershipping_id '523
		rs.Fields("name") = frmordershipping_name 'Stephen Reses
		rs.Fields("company") = frmordershipping_company 'Stephen Reses Productions
		rs.Fields("title") = frmordershipping_title 'President
		rs.Fields("address") = frmordershipping_address '113 Parkwood Place
		rs.Fields("address2") = frmordershipping_address2 '
		rs.Fields("city") = frmordershipping_city 'Linwood
		rs.Fields("state") = frmordershipping_state 'New Jersey
		rs.Fields("zip") = frmordershipping_zip '08221
		rs.Fields("country") = frmordershipping_country 'United States of America
		rs.Update
		rs.Close

		closeDB
	End Function
	
	Function DeleteOrder(intID)
		openDB
		cartDB.Execute(SQL("UPDATE ordersession SET isdeleted = true WHERE id = " & intID & ""))
		closeDB
	End Function

	Function MarkOrderPaid(intID)
		openDB
		cartDB.Execute(SQL("UPDATE ordersession SET ispaid = true, paid = #" & date & " " & time & "# WHERE id = " & intID & ""))
		closeDB
	End Function

	Function ClearOrderPaid(intID)
		openDB
		cartDB.Execute(SQL("UPDATE ordersession SET ispaid = false, paid = NULL WHERE id = " & intID & ""))
		closeDB
	End Function

	Function MarkOrderCompleted(intID)
		openDB
		cartDB.Execute(SQL("UPDATE ordersession SET iscompleted = true, completed = #" & date & " " & time & "# WHERE id = " & intID & ""))
		closeDB
	End Function

	Function ClearOrderCompleted(intID)
		'bolDebug = True
		openDB
		cartDB.Execute(SQL("UPDATE ordersession SET iscompleted = false, completed = NULL WHERE id = " & intID & ""))
		closeDB
	End Function

	Function RenderOrderReview(intID)
		'bolDebug = True

		intID = Right(intID,Len(intID)-1)
		intID = Fix(intID)

		openDB
		Set dbOrder = cartDB.Execute(SQL("SELECT * FROM ordersession WHERE id = " & intID & ""))
		If (dbOrder.EOF) Then
			print("There is no order by this id.")
		Else
			%>
			<input type="hidden" name="id" value="<%=Request.QueryString("orderid")%>">
			<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="100%">
				<tr>
					<td colspan="2"><b>Billing Information</b></td>
				</tr>
			<%
			Set dbPaymentMethods = cartDB.Execute(SQL("SELECT DISTINCT paymentmethod FROM ordersession ORDER BY paymentmethod"))
			Set dbBilling = cartDB.Execute(SQL("SELECT * FROM orderbilling WHERE id = " & dbOrder("billingid") & ""))
			If (dbBilling.EOF) Then
			%>
					<tr>
						<td colspan="2">There is no billing information available for this order.</td>
					</tr>
			<%
			Else
				%>
					<tr>
						<td align="right">Billing ID</td>
						<td><input name="orderbilling_id" type="hidden" value="<%=dbBilling("id")%>" /><%=dbBilling("id")%></td>
					</tr>
					<tr>
						<td valign="top" align="right">Payment Method</td>
						<td valign="top">
							<select name="ordersession_paymentmethod" size="1">
							<% 
							Do Until (dbPaymentMethods.EOF)
								If Not (isNull(dbPaymentMethods("paymentmethod"))) And (Len(dbPaymentMethods("paymentmethod")) > 0) Then
									If (dbOrder("paymentmethod") = dbPaymentMethods("paymentmethod")) Then
										%><option selected><%=dbPaymentMethods("paymentmethod")%></option><%
									Else
										%><option><%=dbPaymentMethods("paymentmethod")%></option><%
									End IF
								End If
								dbPaymentMethods.MoveNext
							Loop
							%>
							</select>
						</td>
					</tr>
					<tr>
						<td align="right">Name</td>
						<td><input name="orderbilling_name" size="30" value="<%=dbBilling("name")%>" /></td>
					</tr>
					<tr>
						<td align="right">Company</td>
						<td><input name="orderbilling_company" size="30" value="<%=dbBilling("company")%>" /></td>
					</tr>
					<tr>
						<td align="right">Title</td>
						<td><input name="orderbilling_title" size="30" value="<%=dbBilling("title")%>" /></td>
					</tr>
					<tr>
						<td align="right">Address</td>
						<td><input name="orderbilling_address" size="30" value="<%=dbBilling("address")%>" /></td>
					</tr>
					<tr>
						<td align="right">Address 2</td>
						<td><input name="orderbilling_address2" size="30" value="<%=dbBilling("address2")%>" /></td>
					</tr>
					<tr>
						<td align="right">City</td>
						<td><input name="orderbilling_city" size="30" value="<%=dbBilling("city")%>" /></td>
					</tr>
					<tr>
						<td align="right">State / Provence</td>
						<td><input name="orderbilling_state" size="30" value="<%=dbBilling("state")%>" /></td>
					</tr>
					<tr>
						<td align="right">Zip / Postal Code</td>
						<td><input name="orderbilling_zip" size="30" value="<%=dbBilling("zip")%>" /></td>
					</tr>
					<tr>
						<td align="right">Country</td>
						<td><input name="orderbilling_country" size="30" value="<%=dbBilling("country")%>" /></td>
					</tr>
					<tr>
						<td align="right">Phone Number</td>
						<td><input name="orderbilling_phonenumber" size="30" value="<%=dbBilling("phonenumber")%>" /></td>
					</tr>
					<tr>
						<td align="right">Email Address</td>
						<td><input name="orderbilling_emailaddress" size="30" value="<%=dbBilling("emailaddress")%>" /></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td><input type="submit" name="action" value="<%=lblSaveCustomerInfo%>" /><input type="reset" value="Reset Changes" /> <i><b>NOTE:</b></i> These buttons effect both billing and shipping information.</td>
					</tr>
				<%
			End If
			%>
				<tr>
					<td colspan="2"><b>Shipping Information</b></td>
				</tr>
			<%
			Set dbShippingMethods = cartDB.Execute(SQL("SELECT DISTINCT shippingmethod FROM ordersession ORDER BY shippingmethod"))
			Set dbShipping = cartDB.Execute(SQL("SELECT * FROM ordershipping WHERE id = " & dbOrder("shippingid") & ""))
			If (dbShipping.EOF) Then
			%>
				<tr>
					<td colspan="2">There is no shipping information available for this order.</td>
				</tr>
			<%
			Else
				%>
					<tr>
						<td align="right">Shipping ID</td>
						<td><input name="ordershipping_id" type="hidden" value="<%=dbShipping("id")%>" /><%=dbShipping("id")%></td>
					</tr>
					<tr>
						<td align="right">Shipping Method</td>
						<td>
							<select name="ordersession_shippingmethod" size="1">
							<% 
							Do Until (dbShippingMethods.EOF)
								If Not (isNull(dbShippingMethods("shippingmethod"))) And (Len(dbShippingMethods("shippingmethod")) > 0) Then
									If (dbOrder("shippingmethod") = dbShippingMethods("shippingmethod")) Then
										%><option selected><%=dbShippingMethods("shippingmethod")%></option><%
									Else
										%><option><%=dbShippingMethods("shippingmethod")%></option><%
									End IF
								End If
								dbShippingMethods.MoveNext
							Loop
							%>
							</select>
						</td>
					</tr>
					<tr>
						<td align="right">Name</td>
						<td><input name="ordershipping_name" size="30" value="<%=dbShipping("name")%>" /></td>
					</tr>
					<tr>
						<td align="right">Company</td>
						<td><input name="ordershipping_company" size="30" value="<%=dbShipping("company")%>" /></td>
					</tr>
					<tr>
						<td align="right">Title</td>
						<td><input name="ordershipping_title" size="30" value="<%=dbShipping("title")%>" /></td>
					</tr>
					<tr>
						<td align="right">Address</td>
						<td><input name="ordershipping_address" size="30" value="<%=dbShipping("address")%>" /></td>
					</tr>
					<tr>
						<td align="right">Address 2</td>
						<td><input name="ordershipping_address2" size="30" value="<%=dbShipping("address2")%>" /></td>
					</tr>
					<tr>
						<td align="right">City</td>
						<td><input name="ordershipping_city" size="30" value="<%=dbShipping("city")%>" /></td>
					</tr>
					<tr>
						<td align="right">State / Provence</td>
						<td><input name="ordershipping_state" size="30" value="<%=dbShipping("state")%>" /></td>
					</tr>
					<tr>
						<td align="right">Zip / Postal Code</td>
						<td><input name="ordershipping_zip" size="30" value="<%=dbShipping("zip")%>" /></td>
					</tr>
					<tr>
						<td align="right">Country</td>
						<td><input name="ordershipping_country" size="30" value="<%=dbShipping("country")%>" /></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td><input type="submit" name="action" value="<%=lblSaveCustomerInfo%>" /><input type="reset" value="Reset Changes" /> <i><b>NOTE:</b></i> These buttons effect both billing and shipping information.</td>
					</tr>
				<%
			End If
			%>
				<tr>
					<td colspan="2"><b>Order Information</b></td>
				</tr>
				<tr>
					<td width="30%" valign="top" align="right">ID</td>
					<td width="70%" valign="top"><input name="ordersession_id" type="hidden" value="<%=dbOrder("id")%>" />C<%=PadNumber(dbOrder("id"))%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Billing ID</td>
					<td valign="top"><%=dbOrder("billingid")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Shipping ID</td>
					<td valign="top"><%=dbOrder("shippingid")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Session ID</td>
					<td valign="top"><%=dbOrder("sessionid")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">IP Address</td>
					<td valign="top"><%=dbOrder("ipaddress")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Browser</td>
					<td valign="top"><%=dbOrder("agent")%></td>
				</tr>
<!--				<script>
					function dateToggle(id)
					{
						id.value = "Is Set To Now";
						id.outerHTML = id.outerHTML.replace("button","text");
					}
				</script>-->
				<tr>
					<td valign="top" align="right">Created</td>
					<td valign="top"><!--<input type="button" name="ordersession_created" value="<%=dbOrder("created")%> - Set To Now" onclick="dateToggle(this);" />--><%=dbOrder("created")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Completed</td>
					<td valign="top"><%=dbOrder("completed")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Cart Session ID</td>
					<td valign="top"><%=dbOrder("cartsessionid")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Step 1 Completed</td>
					<td valign="top"><%=dbOrder("step1completed")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Step 2 Completed</td>
					<td valign="top"><%=dbOrder("step2completed")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Step 3 Completed</td>
					<td valign="top"><%=dbOrder("step3completed")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Step 4 Completed</td>
					<td valign="top"><%=dbOrder("step4completed")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Shipping Method</td>
					<td valign="top"><%=dbOrder("shippingmethod")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Is Completed?</td>
					<td valign="top"><%=dbOrder("iscompleted")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Is Deleted?</td>
					<td valign="top"><%=dbOrder("isdeleted")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Sub Total</td>
					<td valign="top">$<%=FormatNumber(dbOrder("subtotalcost"),2)%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Shipping Cost</td>
					<td valign="top">$<%=FormatNumber(dbOrder("shippingcost"),2)%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Sales Tax</td>
					<td valign="top">$<%=FormatNumber(dbOrder("salestaxcost"),2)%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Total Cost</td>
					<td valign="top">$<%=FormatNumber(dbOrder("totalcost"),2)%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Payment Method</td>
					<td valign="top"><%=dbOrder("paymentmethod")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Is Paid?</td>
					<td valign="top"><%=dbOrder("ispaid")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Paid</td>
					<td valign="top"><%=dbOrder("paid")%></td>
				</tr>
				<tr>
					<td valign="top" align="right">Transaction Data</td>
					<td valign="top"><% If Not (isNull(dbOrder("transactiondata"))) Then print(Replace(dbOrder("transactiondata"),"&","<br />")) Else print("") %></td>
				</tr>
				<tr>
					<td colspan="2"><b>Order Products</b></td>
				</tr>
			<%
			Set dbOrderData = cartDB.Execute(SQL("SELECT product.id, product.name, pricing.price, pricing.discountcode, orderdata.quantity, orderdata.senttofulfillment, orderdata.datesenttofulfillment FROM orderdata, product, pricing WHERE orderdata.productid = product.id AND orderdata.pricingid = pricing.id AND pricing.productid = product.id AND orderdata.ordersessionid = " & dbOrder("id") & " ORDER BY product.name, pricing.id DESC"))
			If (dbOrderData.EOF) Then
				%>
					<tr>
						<td colspan="2">There were no products in this order.</td>
					</tr>
				<%
			Else
				%>
					<tr>
						<td colspan="2">
							<table border="0" cellpadding="2" cellspacing="0" width="100%">
								<tr bgcolor="#000000">
									<td><font color="#FFFFFF">ID</font></td>
									<td><font color="#FFFFFF">Name</font></td>
									<td><font color="#FFFFFF">Price Each</font></td>
									<td><font color="#FFFFFF">Discount Code</font></td>
									<td><font color="#FFFFFF">Quantity</font></td>
									<td><font color="#FFFFFF">Total</font></td>
									<td><font color="#FFFFFF">Sent To Ful.</font></td>
									<td><dont color="#FFFFFF">Date Sent To Ful.</font></td>
								</tr>
				<%
				Do Until dbOrderData.EOF
					%>
								<tr>
									<td><%=dbOrderData("id")%></td>
									<td><%=dbOrderData("name")%></td>
									<td>$<%=FormatNumber(dbOrderData("price"),2)%></td>
									<td><%=dbOrderData("discountcode")%></td>
									<td><%=dbOrderData("quantity")%></td>
									<td>$<%=FormatNumber(Fix(dbOrderData("quantity"))*dbOrderData("price"),2)%></td>
									<td><%=dbOrderData("senttofulfillment")%></td>
									<td><%=dbOrderData("datesenttofulfillment")%></td>
								</tr>
					<%
					dbOrderData.MoveNext
				Loop
				%>
							</table>
						</td>
					</tr>
				<%
			End If
			Set dbOrderData = Nothing
			%>
				<%
			End If
			%>
				<tr>
					<td colspan="2"><b>Account Functions (Generates Emails)</b></td>
				</tr>
				<tr>
					<td valign="top" align="right">
						<% If (dbOrder("ispaid") And dbOrder("iscompleted")) Then thisDisabled = " DISABLED" Else thisDisabled = "" %>
						<input type="submit" name="action" value="<%=lblPaidCompleted%>"<%=thisDisabled%>>
					</td>
					<td valign="top">
						Marks account as paid, and sends paid email to customer, and fulfillment email to vendor.
					</td>
				</tr>
				<tr>
					<td colspan="2"><b>Account Functions (Does Not Generate Emails)</b></td>
				</tr>
				<tr>
					<td valign="top" align="right">
						<% If (dbOrder("ispaid")) Then thisDisabled = "" Else thisDisabled = " DISABLED" %>
						<input type="submit" name="action" value="<%=lblClearPaid%>"<%=thisDisabled%>>
					</td>
					<td valign="top">
						Clears order paid flag and date.
					</td>
				</tr>
				<tr>
					<td valign="top" align="right">
						<% If (dbOrder("ispaid")) Then thisDisabled = " DISABLED" Else thisDisabled = "" %>
						<input type="submit" name="action" value="<%=lblPaid%>"<%=thisDisabled%>>
					</td>
					<td valign="top">
						Marks order as paid, does not generate emails.
					</td>
				</tr>
				<tr>
					<td valign="top" align="right">
						<% If (dbOrder("iscompleted")) Then thisDisabled = "" Else thisDisabled = " DISABLED" %>
						<input type="submit" name="action" value="<%=lblClearCompleted%>"<%=thisDisabled%>>
					</td>
					<td valign="top">
						Clears order completed flag and date.
					</td>
				</tr>
				<tr>
					<td valign="top" align="right">
						<% If (dbOrder("iscompleted")) Then thisDisabled = " DISABLED" Else thisDisabled = "" %>
						<input type="submit" name="action" value="<%=lblCompleted%>"<%=thisDisabled%>>
					</td>
					<td valign="top">
						Marks order as completed, does not generate emails.
					</td>
				</tr>
				<tr>
					<td colspan="2"><b>Generate Customer Emails</b></td>
				</tr>
				<tr>
					<td valign="top" align="right">
						<input type="submit" name="action" value="<%=lblEmailCustomerReceipt%>">
					</td>
					<td valign="top">
						Sends the customer receipt.
					</td>
				</tr>
				<tr>
					<td valign="top" align="right">
						<input type="submit" name="action" value="<%=lblEmailMerchantReceipt%>">
					</td>
					<td valign="top">
						Sends the merchant receipt.
					</td>
				</tr>
				<tr>
					<td valign="top" align="right">
						<% If Not (dbOrder("ispaid")) Then thisDisabled = " DISABLED" Else thisDisabled = "" %>
						<input type="submit" name="action" value="<%=lblEmailPaid%>"<%=thisDisabled%>>
					</td>
					<td valign="top">
						Sends customer paid confirmation.
					</td>
				</tr>
				<tr>
					<td valign="top" align="right">
						<input type="submit" name="action" value="<%=lblEmailFulfillment%>"<%=thisDisabled%>>
					</td>
					<td valign="top">
						Sends the fulfillment request to vendor.
					</td>
				</tr>
			</table>
			<%
'		End If
		Set dbOrder = Nothing
		
		closeDB
	End Function
	
	Function RenderOrderList(dateStart, dateEnd)
		openDB
		Set dbOrderCount = cartDB.Execute(SQL("SELECT count(id) as ordercount FROM ordersession WHERE isdeleted = false"))
		orderCount = dbOrderCount("ordercount")
		Set dbOrderCount = Nothing

		Set dbOrderCount = cartDB.Execute(SQL("SELECT count(id) as ordercount FROM ordersession WHERE isdeleted = false AND created >= #" & dateStart & " 12:00:00 AM# AND created <= #" & dateEnd & " 11:59:59 PM#"))
		visibleOrderCount = dbOrderCount("ordercount")
		Set dbOrderCount = Nothing
		
		Set dbOrders = cartDB.Execute(SQL("SELECT ordersession.id, ordersession.isdeleted, orderbilling.name, orderbilling.emailaddress, ordersession.totalcost, ordersession.created, ordersession.completed, ordersession.ispaid, ordersession.paid, ordersession.step1completed, ordersession.step2completed, ordersession.step3completed, ordersession.step4completed, ordersession.paymentmethod FROM orderbilling, ordersession WHERE orderbilling.id = ordersession.billingid AND ordersession.created >= #" & dateStart & " 12:00:00 AM# AND ordersession.created <= #" & dateEnd & " 11:59:59 PM# AND isdeleted = false ORDER BY ordersession.created DESC, ordersession.completed DESC"))
'		Set dbOrders = cartDB.Execute(SQL("SELECT ordersession.id, ordersession.isdeleted, orderbilling.name, orderbilling.emailaddress, ordersession.totalcost, ordersession.created, ordersession.completed, ordersession.ispaid, ordersession.paid, ordersession.step1completed, ordersession.step2completed, ordersession.step3completed, ordersession.step4completed, ordersession.paymentmethod FROM orderbilling, ordersession WHERE orderbilling.id = ordersession.billingid AND ordersession.created >= #" & dateStart & " 12:00:00 AM# AND ordersession.created <= #" & dateEnd & " 11:59:59 PM# ORDER BY ordersession.created DESC, ordersession.completed DESC"))
		If (dbOrders.EOF) Then
			Response.Write("There are no orders available.")
		Else
			%>
			<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse;" width="100%">
				<tr>
					<td colspan="15">
						Start Date: <input type="text" name="dateStart" value="<%=dateStart%>">
						End Date: <input type="text" name="dateEnd" value="<%=dateEnd%>">
						<input type="submit" name="function" value="Search">
					</td>
				</tr>
				<tr>
					<td colspan="15">
						(<%=visibleOrderCount%>) Visible Orders&nbsp;&nbsp;&nbsp;(<%=orderCount%>) Total Orders&nbsp;&nbsp;&nbsp;Displaying Last (<%=DateDiff("d",dateStart,dateEnd)+1%>) Days
					</td>
				</tr>
				<tr>
					<td colspan="15">
						<table border="0" cellpadding="0" cellspacing="0">
							<tr>
							<%
								colorPaid = "#99FF99"
								colorPaidBorder = "#66CC66"
								colorWaitForPayment = "#FFFF99"
								colorWaitForPaymentBorder = "#CCCC66"
								colorIncompleteInProgress = "#FFFFFF"
								colorIncompleteInProgressBorder = "#CCCCCC"
								colorCanceledDeleted = "#FF9999"
								colorCanceledDeletedBorder = "#CC6666"
							%>
								<td>Legend:&nbsp;&nbsp;</td>
								<td bgcolor="<%=colorPaid%>" style="border: 2px solid <%=colorPaidBorder%>">&nbsp;&nbsp;&nbsp;&nbsp;</td>
								<td>&nbsp; Paid/Completed&nbsp;&nbsp;</td>
								<td>&nbsp;</td>
								<td bgcolor="<%=colorWaitForPayment%>" style="border: 2px solid <%=colorWaitForPaymentBorder%>">&nbsp;&nbsp;&nbsp;&nbsp;</td>
								<td>&nbsp; Waiting For Payment&nbsp;&nbsp;</td>
								<td>&nbsp;</td>
								<td bgcolor="<%=colorIncompleteInProgress%>" style="border: 2px solid <%=colorIncompleteInProgressBorder%>">&nbsp;&nbsp;&nbsp;&nbsp;</td>
								<td>&nbsp; Incomplete/In Progress&nbsp;&nbsp;</td>
								<td>&nbsp;</td>
								<td bgcolor="<%=colorCanceledDeleted%>" style="border: 2px solid <%=colorCanceledDeletedBorder%>">&nbsp;&nbsp;&nbsp;&nbsp;</td>
								<td>&nbsp; Canceled/Deleted&nbsp;&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td background="/images/global_back.gif"><font style="color: white;"><b>ID</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Name</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Email Address</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Qty</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Calc. Total</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Total</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Created</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Paid</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Completed</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Payment</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>1</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>2</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>3</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>4</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>P</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>C</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Delete</b></font></td>
				</tr>
			<%
			Do Until dbOrders.EOF
				Set dbQuantity = cartDB.Execute("SELECT sum(quantity) as quantity FROM orderdata WHERE ordersessionid = " & dbOrders("id") & "")
				If (dbOrders("isdeleted")) Then
					thisColor = colorCanceledDeleted
				Else
					If ((isDate(dbOrders("step1completed"))) And (isDate(dbOrders("step2completed"))) And (isDate(dbOrders("step3completed"))) And (isDate(dbOrders("step4completed")))) Then
						If ((isDate(dbOrders("paid"))) And (isDate(dbOrders("completed")))) Then
							thisColor = colorPaid
						Else
							thisColor = colorWaitForPayment
						End If
					Else
						thisColor = colorIncompleteInProgress
					End If
				End If
				%>
				<tr bgcolor="<%=thisColor%>">
					<td><a href="<%=Request.ServerVariables("URL")%>?function=ViewOrders&orderid=C<%=PadNumber(dbOrders("id"))%>"><%=dbOrders("id")%></a></td>
					<td><%=dbOrders("name")%></td>
					<td><%=dbOrders("emailaddress")%></td>
					<td><%=dbQuantity("quantity")%></td>
					<td align="right">$<%=FormatNumber(dbOrders("totalcost"),2)%></td>
					<td align="right">$<%=FormatNumber(dbOrders("totalcost"),2)%></td>
					<td><%=dbOrders("created")%></td>
					<td><%=dbOrders("paid")%></td>
					<td><%=dbOrders("completed")%></td>
					<td><%=dbOrders("paymentmethod")%></td>
					<td align="center"><% If (isDate(dbOrders("step1completed"))) Then Response.Write("X") %></td>
					<td align="center"><% If (isDate(dbOrders("step2completed"))) Then Response.Write("X") %></td>
					<td align="center"><% If (isDate(dbOrders("step3completed"))) Then Response.Write("X") %></td>
					<td align="center"><% If (isDate(dbOrders("step4completed"))) Then Response.Write("X") %></td>
					<td align="center"><% If (isDate(dbOrders("paid"))) Then Response.Write("X") %></td>
					<td align="center"><% If (isDate(dbOrders("completed"))) Then Response.Write("X") %></td>
					</form><form method="post" action="<%=Request.ServerVariables("URL") & "?function=" & Request.QueryString("function")%>">
					<td align="center">
						<input type="hidden" name="id" value="<%=dbOrders("id")%>">
						<input type="submit" name="action" value="<%=lblDelete%>">
					</td>
					</form>
				</tr>
				<%
				dbOrders.MoveNext
			Loop
			%>
			</table>
			<%
		End If
		Set dbOrders = Nothing
		
		Set dbOrderCount = cartDB.Execute(SQL("SELECT count(id) as ordercount FROM ordersession WHERE isdeleted = true"))
		canceledOrderCount = dbOrderCount("ordercount")
		Set dbOrderCount = Nothing

		Set dbOrderCount = cartDB.Execute(SQL("SELECT count(id) as ordercount FROM ordersession WHERE isdeleted = true AND created >= #" & dateStart & " 12:00:00 AM# AND created <= #" & dateEnd & " 11:59:59 PM#"))
		canceledVisibleOrderCount = dbOrderCount("ordercount")
		Set dbOrderCount = Nothing
		
		Set dbOrders = cartDB.Execute(SQL("SELECT ordersession.id, ordersession.isdeleted, orderbilling.name, orderbilling.emailaddress, ordersession.totalcost, ordersession.created, ordersession.completed, ordersession.ispaid, ordersession.paid, ordersession.step1completed, ordersession.step2completed, ordersession.step3completed, ordersession.step4completed, ordersession.paymentmethod FROM orderbilling, ordersession WHERE orderbilling.id = ordersession.billingid AND ordersession.created >= #" & dateStart & " 12:00:00 AM# AND ordersession.created <= #" & dateEnd & " 11:59:59 PM# AND isdeleted = true ORDER BY ordersession.created DESC, ordersession.completed DESC"))
'		Set dbOrders = cartDB.Execute(SQL("SELECT ordersession.id, ordersession.isdeleted, orderbilling.name, orderbilling.emailaddress, ordersession.totalcost, ordersession.created, ordersession.completed, ordersession.ispaid, ordersession.paid, ordersession.step1completed, ordersession.step2completed, ordersession.step3completed, ordersession.step4completed, ordersession.paymentmethod FROM orderbilling, ordersession WHERE orderbilling.id = ordersession.billingid AND ordersession.created >= #" & dateStart & " 12:00:00 AM# AND ordersession.created <= #" & dateEnd & " 11:59:59 PM# ORDER BY ordersession.created DESC, ordersession.completed DESC"))
		If (dbOrders.EOF) Then
			Response.Write("There are canceled/deleted orders available.")
		Else
			%>
			<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse;" width="100%">
				<tr>
					<td colspan="14">
						<b>Canceled/Deleted Orders</b>
					</td>
				</tr>
				<tr>
					<td colspan="14">
						(<%=canceledVisibleOrderCount%>) Visible Canceled/Deleted Orders&nbsp;&nbsp;&nbsp;(<%=canceledOrderCount%>) Total Canceled/Deleted Orders&nbsp;&nbsp;&nbsp;Displaying Last (<%=DateDiff("d",dateStart,dateEnd)+1%>) Days
					</td>
				</tr>
				<tr>
					<td background="/images/global_back.gif"><font style="color: white;"><b>ID</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Name</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Email Address</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Total</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Created</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Paid</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Completed</b></font></td>
					<td background="/images/global_back.gif"><font style="color: white;"><b>Payment</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>1</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>2</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>3</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>4</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>P</b></font></td>
					<td background="/images/global_back.gif" align="center"><font style="color: white;"><b>C</b></font></td>
				</tr>
			<%
			Do Until dbOrders.EOF
				%>
				<tr bgcolor="<%=colorCanceledDeleted%>">
					<td><a href="<%=Request.ServerVariables("URL")%>?function=ViewOrders&orderid=C<%=PadNumber(dbOrders("id"))%>"><%=dbOrders("id")%></a></td>
					<td><%=dbOrders("name")%></td>
					<td><%=dbOrders("emailaddress")%></td>
					<td align="right">$<%=FormatNumber(dbOrders("totalcost"),2)%></td>
					<td><%=dbOrders("created")%></td>
					<td><%=dbOrders("paid")%></td>
					<td><%=dbOrders("completed")%></td>
					<td><%=dbOrders("paymentmethod")%></td>
					<td align="center"><% If (isDate(dbOrders("step1completed"))) Then Response.Write("X") %></td>
					<td align="center"><% If (isDate(dbOrders("step2completed"))) Then Response.Write("X") %></td>
					<td align="center"><% If (isDate(dbOrders("step3completed"))) Then Response.Write("X") %></td>
					<td align="center"><% If (isDate(dbOrders("step4completed"))) Then Response.Write("X") %></td>
					<td align="center"><% If (isDate(dbOrders("paid"))) Then Response.Write("X") %></td>
					<td align="center"><% If (isDate(dbOrders("completed"))) Then Response.Write("X") %></td>
				</tr>
				<%
				dbOrders.MoveNext
			Loop
			%>
			</table>
			<%
		End If
		Set dbOrders = Nothing
				
		closeDB
	End Function
	
'SQL	
	%><!--#include virtual="/cart/admin/sql.inc.asp"--><%
	Function SQLHandler
		RenderSQL
	End Function

' Email Templates
	Function EmailTemplatesHandler
		selectedTemplate = Fix(Request("templateName"))
		templateSubject = makeSafe(Request.Form("emailSubject"))
		templateEmail = makeSafe(Request.Form("emailTemplate"))
		If (Request.Form("function") = "Save Template") Then
			SaveTemplate selectedTemplate, templateSubject, templateEmail
			strSuccessMessage = "Saved email template."
			redirect Request.ServerVariables("URL") & "?function=EmailTemplates&templateName=" & selectedTemplate & "&message=" & strSuccessMessage
		End If
		RenderEmailTemplateSelect selectedTemplate
		If (Fix(Request.Form("templateName")) > 0) Then
			If (Request.Form("function") = "Load Template") Then
				RenderEmailTemplateInput selectedTemplate
			ElseIf (Request.Form("function") = "Preview Template") Then
				RenderTemplatePreview selectedTemplate
			End If
		End If
	End Function
	
	Function RenderTemplatePreview(intID)
		openDB
		Set dbLastID = cartDB.Execute("SELECT max(id) as lastID FROM ordersession")
		lastID = dbLastID("lastID")
		Set dbLastID = Nothing
		
		Set dbEmail = cartDB.Execute("SELECT subject, body FROM email WHERE id = " & intID & "")
		If (dbEmail.EOF) Then
			Response.Write("Could not find this template.")
		Else
			mSubject = ParseAndReplaceEmail(dbEmail("subject"),lastID,1)
			mMessage = ParseAndReplaceEmail(dbEmail("body"),lastID,1)
			
			%>
			<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse;">
				<tr>
					<td>
						Subject: <%=mSubject%><br />
						<textarea cols="80" rows="30" wrap="soft" readonly><%=mMessage%></textarea>
					</td>
				</tr>
			</table>
			<%
		End If
		closeDB
	End Function
	
	Function SaveTemplate(intID,strSubject,strEmail)
		openDB
		cartDB.Execute("UPDATE email SET subject = '" & strSubject & "', body = '" & strEmail & "' WHERE id = " & intID & "")
		closeDB
	End Function
	
	Function RenderEmailTemplateInput(selectedTemplate)
		openDB
		Set dbEmails = cartDB.Execute(SQL("SELECT subject, body FROM email WHERE id = " & selectedTemplate & ""))
		If (dbEmails.EOF) Then
			ReportError "Tried to load an invalid email template."
		Else
			%>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td width="1" valign="top">
						Subject: <input type="text" name="emailSubject" size="80" value="<%=dbEmails("subject")%>"><br />
						<textarea cols="80" rows="30" wrap="soft" name="emailTemplate"><%=dbEmails("body")%></textarea><br />
						<input type="submit" name="function" value="Save Template">
					</td>
					<td valign="top">
						<font style="font-family: Arial; font-size: 11px;">
						<b>Dynamic Fields</b><br />
						%NAME%<br />&nbsp;&nbsp;&nbsp;Billing Name<br />
						%TITLE%<br />&nbsp;&nbsp;&nbsp;Billing Title<br />
						%COMPANY%<br />&nbsp;&nbsp;&nbsp;Billing Company<br />
						%ADDRESS%<br />&nbsp;&nbsp;&nbsp;Billing Address<br />
						%CITY%<br />&nbsp;&nbsp;&nbsp;Billing City<br />
						%STATE%<br />&nbsp;&nbsp;&nbsp;Billing State<br />
						%ZIP%<br />&nbsp;&nbsp;&nbsp;Billing Zip<br />
						%PHONENUMBER%<br />&nbsp;&nbsp;&nbsp;Phone Number<br />
						%EMAILADDRESS%<br />&nbsp;&nbsp;&nbsp;Email Address<br />
						%INVOICE%<br />&nbsp;&nbsp;&nbsp;Invoice Number<br />
						%BILLINGINFO%<br />&nbsp;&nbsp;&nbsp;Billing Information<br />
						%SHIPPINGINFO%<br />&nbsp;&nbsp;&nbsp;Shipping Information<br />
						%PAYMENTMETHOD%<br />&nbsp;&nbsp;&nbsp;Payment Method<br />
						%SHIPPINGMETHOD%<br />&nbsp;&nbsp;&nbsp;Shipping Method<br />
						</font>
					</td>
				</tr>
			</table>
			<%
		End If
		closeDB
	End Function
	
	Function RenderEmailTemplateSelect(selectedTemplate)
		openDB
		Set dbEmails = cartDB.Execute(SQL("SELECT id, name FROM email ORDER BY name"))
		If (dbEmails.EOF) Then
			ReportError "There are no email templates setup."
		Else
			%>
			<select name="templateName" size="1">
				<option value="-1"> ... Select an email template to edit. ... </option>
			<%
			Do Until (dbEmails.EOF)
				If (Fix(dbEmails("id")) = selectedTemplate) Then
					thisSelected = " selected"
				Else
					thisSelected = ""
				End If
				%><option value="<%=dbEmails("id")%>"<%=thisSelected%>><%=dbEmails("name")%></option><%
				dbEmails.MoveNext
			Loop
			%>
			</option><input type="submit" name="function" value="Load Template"><input type="submit" name="function" value="Preview Template">
			<%
		End If
		closeDB
	End Function
	
' Email History
	Function EmailHistoryHandler
		If (Len(Request.QueryString("dateStart")) > 0) Then
			dateStart = makeSafe(Request.QueryString("dateStart"))
		ElseIf (Len(Request.Form("dateStart")) > 0) Then
			dateStart = makeSafe(Request.Form("dateStart"))
		Else
			dateStart = Date
		End If
		If (Len(Request.QueryString("dateEnd")) > 0) Then
			dateEnd = makeSafe(Request.QueryString("dateEnd"))
		ElseIf (Len(Request.Form("dateEnd")) > 0) Then
			dateEnd = makeSafe(Request.Form("dateEnd"))
		Else
			dateEnd = Date
		End If

		If (Request.Form("action") = "Resend") Then
			strTo = ResendEmail(Fix(Request.Form("id")))
			strResent = "Resent message to '" & strTo & "' from message ID# " & Request.Form("id") & "."
			redirect Request.ServerVariables("URL") & "?function=EmailHistory&success=true&dateStart=" & dateStart & "&dateEnd=" & dateEnd & "&message=" & strResent
		Else
			RenderEmailHistoryInput dateStart, dateEnd
			
			If (isDate(dateStart)) And (isDate(dateEnd)) Then
				RenderEmailHistory dateStart, dateEnd
			ElseIf Not (isDate(dateStart)) And (isDate(dateEnd)) Then
				strStartDateInvalid = "Please enter a valid date for the start date."
				If (Request.QueryString("success") <> "false") Then
					redirect Request.ServerVariables("URL") & "?function=EmailHistory&success=false&dateStart=" & dateStart & "&dateEnd=" & dateEnd & "&message=" & strStartDateInvalid
				End If
			ElseIf (isDate(dateStart)) And Not (isDate(dateEnd)) Then
				strEndDateInvalid = "Please enter a valid date for the end date."
				If (Request.QueryString("success") <> "false") Then
					redirect Request.ServerVariables("URL") & "?function=EmailHistory&success=false&dateStart=" & dateStart & "&dateEnd=" & dateEnd & "&message=" & strEndDateInvalid
				End If
			ElseIf Not (isDate(dateStart)) And Not (isDate(dateEnd)) Then
				strStartEndDateInvalid = "Please enter a valid date for the start and end dates."
				If (Request.QueryString("success") <> "false") Then
					redirect Request.ServerVariables("URL") & "?function=EmailHistory&success=false&dateStart=" & dateStart & "&dateEnd=" & dateEnd & "&message=" & strStartEndDateInvalid
				End If
			End If
		End If
	End Function
	
	Function RenderEmailHistoryInput(dateStart, dateEnd)
		print("Start Date: <input type=""text"" name=""dateStart"" value=""" & dateStart & """>")
		print("End Date: <input type=""text"" name=""dateEnd"" value=""" & dateEnd & """>")
		print("<input type=""submit"" name=""function"" value=""Show History"">")
	End Function
	
	Function RenderEmailHistory(dateStart, dateEnd)
		openDB
		print("<table border=""1"" cellpadding=""2"" cellspacing=""0"")")
		print("<tr background=""/images/global_back.gif"">")
		print("<td><font style=""color: white;"">ID</font></td>")
'		print("<td><font style=""color: white;"">From</font></td>")
		print("<td><font style=""color: white;"">To</font></td>")
'		print("<td><font style=""color: white;"">CC</font></td>")
'		print("<td><font style=""color: white;"">BCC</font></td>")
		print("<td><font style=""color: white;"">Sent</font></td>")
		print("<td><font style=""color: white;"">Subject</font></td>")
		print("<td><font style=""color: white;"">Resend</font></td>")
		print("</tr>")
		Set dbEmails = cartDB.Execute("SELECT * FROM emailhistory WHERE mailSent >= #" & dateStart & " 12:00:00 AM# AND mailSent <= #" & dateEnd & " 11:59:59 PM# ORDER BY mailSent DESC")
		Do Until dbEmails.EOF
			print("<tr bgcolor=""#FFFFFF"">")
			print("<td>" & dbEmails("id") & "</td>")
'			print("<td>" & dbEmails("mailFrom") & "</td>")
			print("<td>" & Replace(dbEmails("mailTo"),";","<br />") & "</td>")
'			print("<td>" & Replace(dbEmails("mailCC"),";","<br />") & "</td>")
'			print("<td>")
'			If (Len(dbEmails("mailBCC")) = 0) Then 
'				print("&nbsp;") 
'			Else 
'				print(Replace(dbEmails("mailBCC"),";","<br />"))
'			End If
'			print("</td>")
			print("<td>" & dbEmails("mailSent") & "</td>")
			print("<td>" & dbEmails("mailSubject") & "</td>")
			print("</form><form method=""post"" action=""" & Request.ServerVariables("URL") & "?function=EmailHistory"">")
			print("<input type=""hidden"" name=""id"" value=""" & dbEmails("id") & """>")
			print("<td><input type=""submit"" name=""action"" value=""Resend""></td>")
			print("</form>")
			print("</tr>")
			dbEmails.MoveNext
		Loop
		print("</table>")
		closeDB
	End Function
	
	Function ResendEmail(intID)
		openDB
		Set dbEmail = cartDB.Execute("SELECT * FROM emailhistory WHERE id = " & intID & "")
		ResendEmail = dbEmail("mailTo")
		SMTP dbEmail("mailFrom"), dbEmail("mailTo"), dbEmail("mailCC"), dbEmail("mailBCC"), dbEmail("mailSubject"), dbEmail("mailBody")
		Set dbEmail = Nothing
		closeDB
	End Function
	
' Report Error
	Function ReportErrorHandler
		ReportError "TEST REPORT:  Something is messed up."
	End Function

' View Incidents
	Function ViewIncidentsHandler
		RenderViewIncidents
	End Function
	
	Function RenderViewIncidents
		openDB
		print("<table border=""1"" cellpadding=""2"" cellspacing=""0"")")
		print("<tr bgcolor=""#000000"">")
		print("<td><font style=""color: white;"">ID</font></td>")
		print("<td><font style=""color: white;"">ViewID</font></td>")
		print("<td><font style=""color: white;"">Reported</font></td>")
		print("<td><font style=""color: white;"">Message</font></td>")
		print("<td><font style=""color: white;"">Delete</font></td>")
		print("</tr>")
		Set dbErrors = cartDB.Execute("SELECT * FROM errors ORDER BY reported DESC")
		Do Until dbErrors.EOF
			print("<tr bgcolor=""#FFFFFF"">")
			print("<td><a href=""/cart/admin/incident.asp?view=" & dbErrors("viewid") & """>" & dbErrors("id") & "</a></td>")
			print("<td><a href=""/cart/admin/incident.asp?view=" & dbErrors("viewid") & """>" & dbErrors("viewid") & "</a></td>")
			print("<td><a href=""/cart/admin/incident.asp?view=" & dbErrors("viewid") & """>" & dbErrors("reported") & "</a></td>")
			print("<td><a href=""/cart/admin/incident.asp?view=" & dbErrors("viewid") & """>" & Left(dbErrors("message"),InStr(dbErrors("message"),vbNewline)) & "...</a></td>")
			print("</form><form method=""post"" action=""/cart/admin/incident.asp?view=" & dbErrors("viewid") & """>")
			print("<input type=""hidden"" name=""id"" value=""" & dbErrors("id") & """>")
			print("<td><input type=""submit"" name=""action"" value=""Delete""></td>")
			print("</form>")
			print("</tr>")
			dbErrors.MoveNext
		Loop
		print("</table>")
		closeDB
	End Function

' Fulfillment Data
	Function FulfillmentDataHandler
		If Not (isDate(Request.QueryString("dateStart"))) Then
			dateStart = DateAdd("d",-89,Date)
		ElseIf (Len(Request.QueryString("dateStart")) > 0) Then
			dateStart = makeSafe(Request.QueryString("dateStart"))
		ElseIf (Len(Request.Form("dateStart")) > 0) Then
			dateStart = makeSafe(Request.Form("dateStart"))
		Else
			dateStart = DateAdd("d",-89,Date)
'			dateStart = CDate(DatePart("m",Date) & "/01/" & DatePart("yyyy",Date))
		End If

		If Not (isDate(Request.QueryString("dateEnd"))) Then
			dateEnd = Date
		ElseIf (Len(Request.QueryString("dateEnd")) > 0) Then
			dateEnd = makeSafe(Request.QueryString("dateEnd"))
		ElseIf (Len(Request.Form("dateEnd")) > 0) Then
			dateEnd = makeSafe(Request.Form("dateEnd"))
		Else
			dateEnd = Date
'			dateEnd = DateAdd("d",-1,DateAdd("m",1,CDate(DatePart("m",Date) & "/01/" & DatePart("yyyy",Date))))
		End If
		
'		Response.Write("IDs: " & Request.QueryString("productid") & "")
		If (Request.QueryString("action") = "Generate Fulfillment Report") And (Len(Request.QueryString("productid")) = 0) Then
			strMessage = "Please select one or more products to generate this report for."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&message=" & strMessage & "&dateStart=" & dateStart & "&dateEnd=" & dateEnd
		ElseIf (Request.QueryString("action") = "Generate Fulfillment Report") And (Len(Request.QueryString("productid")) > 0) Then
			RenderFulfillmentReport dateStart, dateEnd, Request.QueryString("productid")
		Else
			RenderFulfillmentSelect dateStart, dateEnd
		End If
	End Function
	
	Function RenderFulfillmentReport(dateStart, dateEnd, sProductIDs)
		openDB
		%>
		<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse;">
			<tr>
				<td>Invoice</td>
				<td>Ship Method</td>
				<td>Name</td>
				<td>Address</td>
				<td>Address 2</td>
				<td>City</td>
				<td>State</td>
				<td>Zip</td>
				<td>Country</td>
				<td>Phone Number</td>
		<%
		Set dbProducts = cartDB.Execute(SQL("SELECT name FROM product WHERE id IN ( " & sProductIDs & " ) ORDER BY name"))
		Do Until dbProducts.EOF
		%>
				<td><%=dbProducts("name")%></td>
		<%
			dbProducts.MoveNext
		Loop
		Set dbProducts = Nothing
		%>
			</tr>
		<%
		
		Set dbOrders = cartDB.Execute(SQL("SELECT ordersession.id FROM ordersession, orderdata WHERE ordersession.id = orderdata.ordersessionid AND ordersession.ispaid = True AND orderdata.productid IN ( " & sProductIDs & " ) AND ordersession.completed >= #" & dateStart & " 12:00:00 AM# AND ordersession.completed <= #" & dateEnd & " 11:59:59 PM# GROUP BY ordersession.id ORDER BY ordersession.id"))
		Do Until dbOrders.EOF
'			Set dbProduct = cartDB.Execute(SQL("SELECT product.name, product.sku, pricing.discountcode FROM product, pricing WHERE product.id = pricing.productid AND product.id = " & Fix(dbOrders("productid")) & " AND pricing.id = " & Fix(dbOrders("pricingid")) & ""))
			Set dbOrderIDs = cartDB.Execute(SQL("SELECT id, created, totalcost, shippingcost, shippingmethod, shippingid, billingid FROM ordersession WHERE id = " & dbOrders("id") & ""))
			Set dbShipping = cartDB.Execute(SQL("SELECT * FROM ordershipping WHERE id = " & dbOrderIDs("shippingid") & ""))
			Set dbBilling = cartDB.Execute(SQL("SELECT * FROM orderbilling WHERE id = " & dbOrderIDs("billingid") & ""))
			%>
			<tr>
				<td><%=dbOrderIDs("id")%></td>
				<td><%=dbOrderIDs("shippingmethod")%></td>
				<td><%=dbShipping("name")%></td>
				<td><%=dbShipping("address")%></td>
				<td></td>
				<td><%=dbShipping("city")%></td>
				<td><%=dbShipping("state")%></td>
				<td><%=dbShipping("zip")%></td>
				<td><% If (dbShipping("country") <> "United States of America") Then %><%=dbShipping("country")%><% End If %></td>
				<td><%=dbBilling("phonenumber")%></td>
		<%
		Set dbProducts = cartDB.Execute(SQL("SELECT id FROM product WHERE id IN ( " & sProductIDs & " ) ORDER BY name"))
		Do Until dbProducts.EOF
			Set dbQuantity = cartDB.Execute(SQL("SELECT quantity FROM orderdata WHERE ordersessionid = " & dbOrders("id") & " AND productid = " & dbProducts("id") & ""))
			If (dbQuantity.EOF) Then
			%>
				<td><!--Blank--></td>
			<%
			Else
			%>
				<td><% If (Fix(dbQuantity("quantity")) > 0) Then %><%=dbQuantity("quantity")%><% End If %></td>
			<%
			End If
			Set dbQuantity = Nothing
			dbProducts.MoveNext
		Loop
		Set dbProducts = Nothing
		%>
			</tr>
			<%
			dbOrders.MoveNext
		Loop
		closeDB
	End Function

	Function RenderFulfillmentSelect(dateStart, dateEnd)
		%>
		</form><form action="<%=Request.ServerVariables("URL")%>?function=<%=Request.QueryString("function")%>" method="get">
		<table border="0" cellpadding="2" cellspacing="0">
			<tr>
				<td rowspan="4" valign="top">
					Select the products you would like to view fulfillment data for:<br />
					<select name="productid" size="20" multiple>
						<% PrintProductOptions -1 %>
					</select><br />
					Use Ctrl+Click to select multiple products.<br />
					Use Shift+Click to select a range of products.<br />
				</td>
				<td align="right">Start Date:&nbsp;</td>
				<td><input type="text" name="dateStart" value="<%=dateStart%>"></td>
			</tr>
			<tr>
				<td align="right">End Date:&nbsp;</td>
				<td><input type="text" name="dateEnd" value="<%=dateEnd%>"></td>
			</tr>
			<tr>
				<td colspan="2" align="center">
					<input type="hidden" name="function" value="FulfillmentData">
					<input type="submit" name="action" value="Generate Fulfillment Report">
				</td>
			</tr>
			<tr>
				<td height="90%" colspan="2">
					<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
				</td>
			</tr>
		</table>
		<%
	End Function

'Spreadsheet
	Function SpreadsheetHandler
		If Not (isDate(Request.QueryString("dateStart"))) Then
			dateStart = DateAdd("d",-89,Date)
		ElseIf (Len(Request.QueryString("dateStart")) > 0) Then
			dateStart = makeSafe(Request.QueryString("dateStart"))
		ElseIf (Len(Request.Form("dateStart")) > 0) Then
			dateStart = makeSafe(Request.Form("dateStart"))
		Else
			dateStart = DateAdd("d",-89,Date)
'			dateStart = CDate(DatePart("m",Date) & "/01/" & DatePart("yyyy",Date))
		End If

		If Not (isDate(Request.QueryString("dateEnd"))) Then
			dateEnd = Date
		ElseIf (Len(Request.QueryString("dateEnd")) > 0) Then
			dateEnd = makeSafe(Request.QueryString("dateEnd"))
		ElseIf (Len(Request.Form("dateEnd")) > 0) Then
			dateEnd = makeSafe(Request.Form("dateEnd"))
		Else
			dateEnd = Date
'			dateEnd = DateAdd("d",-1,DateAdd("m",1,CDate(DatePart("m",Date) & "/01/" & DatePart("yyyy",Date))))
		End If
		
'		Response.Write("IDs: " & Request.QueryString("productid") & "")
		If (Request.QueryString("action") = "Generate Spreadsheet") And (Len(Request.QueryString("productid")) = 0) Then
			strMessage = "Please select one or more products to generate this report for."
			redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&message=" & strMessage & "&dateStart=" & dateStart & "&dateEnd=" & dateEnd
		ElseIf (Request.QueryString("action") = "Generate Spreadsheet") And (Len(Request.QueryString("productid")) > 0) Then
			RenderSpreadsheet dateStart, dateEnd, Request.QueryString("productid")
		Else
			RenderSpreadsheetSelect dateStart, dateEnd
		End If
	End Function
	
	Function RenderSpreadsheet(dateStart, dateEnd, sProductIDs)
		openDB
		%>
		<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse;">
			<tr>
				<td>Invoice</td>
				<td>Date Ordered</td>
				<td>Discount Code</td>
				<td>Total Cost</td>
				<td>Shipping Cost</td>
				<td>Ship Method</td>
				<td>Name</td>
				<td>Address</td>
				<td>City</td>
				<td>State</td>
				<td>Zip</td>
				<td>Country</td>
				<td>Phone Number</td>
				<td>Email Address</td>
				<td>Quantity</td>
				<td>Product Name</td>
			</tr>
		<%
		
		Set dbOrders = cartDB.Execute(SQL("SELECT * FROM ordersession, orderdata WHERE ordersession.id = orderdata.ordersessionid AND ordersession.ispaid = True AND orderdata.productid IN ( " & sProductIDs & " ) AND ordersession.completed >= #" & dateStart & " 12:00:00 AM# AND ordersession.completed <= #" & dateEnd & " 11:59:59 PM# ORDER BY ordersession.id"))
		Do Until dbOrders.EOF
			Set dbProduct = cartDB.Execute(SQL("SELECT product.name, product.sku, pricing.discountcode FROM product, pricing WHERE product.id = pricing.productid AND product.id = " & Fix(dbOrders("productid")) & " AND pricing.id = " & Fix(dbOrders("pricingid")) & ""))
			Set dbShipping = cartDB.Execute(SQL("SELECT * FROM ordershipping WHERE id = " & dbOrders("shippingid") & ""))
			Set dbBilling = cartDB.Execute(SQL("SELECT * FROM orderbilling WHERE id = " & dbOrders("billingid") & ""))
			%>
			<tr>
				<td><%=dbOrders("ordersession.id")%></td>
				<td><%=dbOrders("created")%></td>
				<td><%=dbProduct("discountcode")%></td>
				<td>$<%=FormatNumber(dbOrders("totalcost"),2)%></td>
				<td>$<%=FormatNumber(dbOrders("shippingcost"),2)%></td>
				<td><%=dbOrders("shippingmethod")%></td>
				<td><%=dbShipping("name")%></td>
				<td><%=dbShipping("address")%></td>
				<td><%=dbShipping("city")%></td>
				<td><%=dbShipping("state")%></td>
				<td><%=dbShipping("zip")%></td>
				<td><% If (dbShipping("country") <> "United States of America") Then %><%=dbShipping("country")%><% End If %></td>
				<td><%=dbBilling("phonenumber")%></td>
				<td><%=dbBilling("emailaddress")%></td>
				<td><%=dbOrders("quantity")%></td>
				<td><%=dbProduct("name")%></td>
			</tr>
			<%
			dbOrders.MoveNext
		Loop
		closeDB
	End Function
		
	Function RenderSpreadsheetSelect(dateStart, dateEnd)
		%>
		</form><form action="<%=Request.ServerVariables("URL")%>?function=<%=Request.QueryString("function")%>" method="get">
		<table border="0" cellpadding="2" cellspacing="0">
			<tr>
				<td rowspan="4" valign="top">
					Select the products you would like to view in spreadsheet:<br />
					<select name="productid" size="20" multiple>
						<% PrintProductOptions -1 %>
					</select><br />
					Use Ctrl+Click to select multiple products.<br />
					Use Shift+Click to select a range of products.<br />
				</td>
				<td align="right">Start Date:&nbsp;</td>
				<td><input type="text" name="dateStart" value="<%=dateStart%>"></td>
			</tr>
			<tr>
				<td align="right">End Date:&nbsp;</td>
				<td><input type="text" name="dateEnd" value="<%=dateEnd%>"></td>
			</tr>
			<tr>
				<td colspan="2" align="center">
					<input type="hidden" name="function" value="Spreadsheet">
					<input type="submit" name="action" value="Generate Spreadsheet">
				</td>
			</tr>
			<tr>
				<td height="90%" colspan="2">
					<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
				</td>
			</tr>
		</table>
		<%
	End Function

	Function SalesReportHandler
		frmAction = Request.Form("action")
		frmSent = Request.Form("senttofulfillment")
		qsShow = Request.QueryString("show")
		qsFulfilldata = Request.QueryString("fulfilldata")
		If (Len(qsFulfilldata) > 0) Then
			qsFulfilldata = True
		Else
			qsFulfilldata = False
		End If
		
		If (Len(qsShow) > 0) Then
			RenderSalesReport qsShow, qsFulfilldata
		Else
			If (frmAction = "Save Changes") Then
				If (Len(frmSent) > 0) Then
					openDB
					frmSent = Split(frmSent, ",")
					For Each frm In frmSent
						frmOrder = Split(Trim(frm), "_")
						cartDB.Execute("UPDATE orderdata SET senttofulfillment = true, datesenttofulfillment = DATE() + TIME() WHERE ordersessionid = " & frmOrder(0) & " AND productid = " & frmOrder(1) & "")
					Next
					closeDB
					strMessage = "Saved changes to fulfillment history."
				Else
					strMessage = "No changes needed to be made to fulfillment history."
				End If
				Response.Redirect Request.ServerVariables("URL") & "?function=" & qsFunction & "&message=" & strMessage
			ElseIf (frmAction = "Cancel Changes") Then
				strMessage = "Canceled saving any changes to fulfillment history."
				Response.Redirect Request.ServerVariables("URL") & "?function=" & qsFunction & "&message=" & strMessage
			Else
				RenderSalesReportSelector
			End If
		End If
	End Function

	Function RenderSalesReportSelector
		%>
		</form><form action="<%=Request.ServerVariables("URL")%>?function=<%=Request.QueryString("function")%>" method="get">
		<table border="0" cellpadding="2" cellspacing="0">
			<tr>
				<td rowspan="4" valign="top">
					Select the products you would like to view fulfillment history:<br />
					<select name="show" size="20" multiple>
						<% PrintProductOptions -1 %>
					</select><br />
					Use Ctrl+Click to select multiple products.<br />
					Use Shift+Click to select a range of products.<br />
				</td>
				<td valign="top">
					<table border="0" cellpadding="2" cellspacing="0" width="100%">
						<tr>
							<td align="right"><input type="checkbox" id="fulfilldata" name="fulfilldata" value="true" /></td>
							<td><label for="fulfilldata">Include fulfillment history.</label></td>
						</tr>
						<tr>
							<td colspan="2" align="center">
								<input type="hidden" name="function" value="SalesReport">
								<input type="submit" name="action" value="Generate Report">
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<%
	End Function
	
	Function RenderSalesReport(showIDs, showFulfill)
		openDB
		Set dbProducts = cartDB.Execute(SQL("SELECT id, name FROM product WHERE deleted = false AND id IN (" & showIDs & ") ORDER BY name"))
		%>
		<script>
			var warnedBox = false;
			function doCheckBox(src, id)
			{
				if (!(warnedBox))
				{
					alert("Make sure you save the changes by clicking Save Changes.");
					warnedBox = true;
				}
			}
		</script>
		<table border="1" cellpadding="2" cellspacing="0" width="100%">
		<%
		
		Dim fGrandTotal : fGrandTotal = 0.0
		Dim fSubTotal : fSubTotal = 0.0
		Dim iCount : iCount = 0
		Dim iSubCount : iSubCount = 0
		
		Do Until dbProducts.EOF
			%>
			<tr>
				<td colspan="4"><b><%=dbProducts("name")%></b></td>
			</tr>
			<tr>
				<td><b>ID#</b></td>
				<td><b>Name</b></td>
				<td><b>Date Sold</b></td>
				<td><b>Qty</b></td>
				<td><b>Price Sold</b></td>
				<td><b>Discount Code</b></td>
				<% If (showFulfill) Then %>
				<td colspan="2"><b>Sent to fulfillment?</b></td>
				<% End If %>
			</tr>
			<%
'				"SELECT orderdata.quantity, pricing.price, pricing.discountcode " & _
			Set dbSales = cartDB.Execute(SQL("" & _
				"SELECT ordersession.billingid, orderdata.senttofulfillment, orderdata.datesenttofulfillment, ordersession.id as orderid, orderdata.quantity, ordersession.completed, orderdata.pricingid " & _
				"FROM ordersession " & _
				"  LEFT JOIN orderdata ON orderdata.ordersessionid = ordersession.id " & _
				"WHERE orderdata.productid = " & dbProducts("id") & " AND ordersession.ispaid = true AND ordersession.iscompleted = true AND ordersession.isdeleted = false ORDER BY ordersession.completed" _
				)) '  & _
				
			fSubTotal = 0.0
			iSubCount = 0
  			Do Until dbSales.EOF
				Set dbName = cartDB.Execute(SQL("SELECT name FROM orderbilling WHERE id = " & dbSales("billingid") & ""))
				
				Set dbPrice = cartDB.Execute(SQL("SELECT price, discountcode FROM pricing WHERE id = " & dbSales("pricingid") & ""))
				fSubTotal = fSubTotal + (CDbl(dbPrice("price")) * Fix(dbSales("quantity")))
				fGrandTotal = fGrandTotal + (CDbl(dbPrice("price")) * Fix(dbSales("quantity")))
				iSubCount = iSubCount + Fix(dbSales("quantity"))
				iCount = iCount + Fix(dbSales("quantity"))
				If (bgColor = "#FFFFFF") Then bgColor = "#EEEEEE" Else bgColor = "#FFFFFF"
				%>
				<tr bgcolor="<%=bgColor%>">
					<td><%=dbSales("orderid")%></td>
					<td><%=dbName("name")%></td>
					<td><%=dbSales("completed")%></td>
					<td><%=dbSales("quantity")%></td>
					<td align="right">$<%=FormatNumber((dbPrice("price") * Fix(dbSales("quantity"))),2)%></td>
					<td><%=dbPrice("discountcode")%></td>
					<% If (showFulfill) Then %>
						<% If (dbSales("senttofulfillment")) Then %>
						<td colspan="2"><% If (Len(dbSales("datesenttofulfillment")) > 0) Then Response.Write(dbSales("datesenttofulfillment")) Else Response.Write("<span style=""color: gray;"">Already Sent</span>") %></td>
						<% Else %>
						<td width="1"><input type="checkbox" id="senttofulfillment_<%=dbSales("orderid")%>_<%=dbProducts("id")%>" name="senttofulfillment" value="<%=dbSales("orderid")%>_<%=dbProducts("id")%>" onClick="doCheckBox(this, -1)" /></td>
						<td width="100"><label for="senttofulfillment_<%=dbSales("orderid")%>_<%=dbProducts("id")%>">Mark Fulfilled</label></td>
						<% End If %>
					<% End If %>
				</tr>
				<%
				dbSales.MoveNext
			Loop
			%>
			<tr>
				<td align="right"><b>Sub-Total</b></td>
				<td><%=iSubCount%></td>
				<td align="right">$<%=FormatNumber(fSubTotal,2)%></td>
				<td>&nbsp;</td>
				<% If (showFulfill) Then %>
				<td colspan="2">
				<input type="submit" name="action" value="Save Changes" /><br />
				<input type="submit" name="action" value="Cancel Changes" />
				</td>
				<% End If %>
			</tr>
			<%
			dbProducts.MoveNext
		Loop
		%>
		<tr>
			<td align="right"><b>Total</b></td>
			<td><%=iCount%></td>
			<td align="right">$<%=FormatNumber(fGrandTotal,2)%></td>
			<td>&nbsp;</td>
		</tr>
	</table>
		<%
	End Function
	
	Function ManualOrderEntryHandler
		openDB
		Set dbProduct = cartDB.Execute("SELECT product.id, product.name, pricing.id, pricing.price, pricing.discountcode, pricing.validuntil FROM product, pricing WHERE product.id = pricing.productid AND pricing.isdeleted = false ORDER BY product.id, pricing.isnormal")
		If (dbProduct.EOF) Then
			%>
			There are no products availble to add to this order.<br />
			<%
		Else
			%>
			<script>
				var productPricing = new Array();
			<%

			For Each frm In Request.Form
				If (Left(frm, Len("product_price_")) = "product_price_") Then
					thisProductID = Replace(frm, "product_price_", "")
					If (Len(Request.Form(frm)) > 0) And (Request.Form(frm) <> "other") Then
						thisPriceID = Request.Form(frm)
					Else
						thisPriceID = -1
					End If
					%>
				var product_price_<%=thisProductID%>_selected = <%=thisPriceID%>;<%
				End If
			Next

			lastProduct = 0
			pricingString = ""
			Do Until dbProduct.EOF
				If (lastProduct <> Fix(dbProduct("product.id"))) Then
					If (Len(pricingString) > 0) Then
						%>
				productPricing[<%=lastProduct%>] = "<%=Left(pricingString,Len(pricingString)-1)%>";<%
					End If					
					lastProduct = Fix(dbProduct("product.id"))
					pricingString = ""
				End If

				If (Len(dbProduct("validuntil")) > 0) Then
					thisCode = "*X*PREORDER*X*"
				Else
					thisCode = dbProduct("discountcode")
				End If
				pricingString = pricingString & dbProduct("pricing.id") & ":" & FormatNumber(dbProduct("price"),2) & ":" & thisCode & ";"
				dbProduct.MoveNext
			Loop
			
			Set dbProduct = Nothing
			%>
				function changeProductSelection(product, num)
				{
					window.status = "Value Selected: " + product.value;
					if ((product.value > 0) && (typeof productPricing[product.value] != "undefined"))
					{
						clearOptions(num);
						pushOptions(product.value, num);
						eval("document.forms[0].product_price_" + num + "").style.display = "inline";
						eval("document.forms[0].product_price_" + num + "").disabled = false;
					}
					else
					{
						eval("document.forms[0].product_price_" + num + "").style.display = "none";
						eval("document.forms[0].product_price_" + num + "").disabled = true;
					}
				}
				
				function changePricingSelection(product, num)
				{
					if (product.value == "other")
					{
						changeToManualPrice(num);
					}
				}
				
				function changeToManualPrice(num)
				{
					eval("document.forms[0].product_price_" + num + "").outerHTML = "<input type=\"text\" name=\"product_price_" + num + "\" size=\"10\" /><input type=\"button\" value=\"&nbsp;X&nbsp;\" onClick=\"changeToSelectPrice(this, " + num + ");\" style=\"font-weight: bold;\" />";
					eval("document.forms[0].product_price_" + num + "").focus();
				}
				
				function changeToSelectPrice(xbox, num)
				{
					xbox.outerHTML = "";
					eval("document.forms[0].product_price_" + num + "").outerHTML = "<select name=\"product_price_" + num + "\" size=\"1\" onChange=\"changePricingSelection(this, " + num + ");\"></select>";
					pushOptions(eval("document.forms[0].product_" + num + "").value, num);
					eval("document.forms[0].product_price_" + num + "").focus();
				}
				
				function clearOptions(num)
				{
					// Clear Options
					eval("document.forms[0].product_price_" + num + "").options.length = 0;
				}
				
				function pushOptions(productid, num)
				{
					// Break Up Each ID,Price Pair
					var pricingArray = productPricing[productid].split(";");
					
					// Loop Through Pairs
					for (var p = 0; p < pricingArray.length; p++)
					{
						// Break Up Each ID And Price
						var priceIDPair = pricingArray[p].split(":");
						
						if (p == 0)
							thisComment = " (Normal)";
						else if (priceIDPair[2] == "*X*PREORDER*X*")
							thisComment = " (Pre-Order)";
						else if (priceIDPair[2].length > 0)
							thisComment = " (Code: " + priceIDPair[2] + ")";
						else
							thisComment = "";
						
						eval("document.forms[0].product_price_" + num + "").options[p] = new Option( "$" + priceIDPair[1] + thisComment, priceIDPair[0] );
					}
					eval("document.forms[0].product_price_" + num + "").options[p] = new Option( " -- Other -- ", "other" );
				}
			
			</script>
			
		<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse;">
			<tr>
				<td background="/images/global_back.gif"><font style="color: white;"><b>Product Selection</b></font></td>
			</tr>
			<tr>
				<td>
					<table border="0" cellpadding="3" cellspacing="0" width="100%">
						<tr>
							<% thisID = 1 %>
							<td>Product <%=thisID%>:</td>
							<td><% RenderProductSelection thisID %></td>
							<td><select name="product_price_<%=thisID%>" size="1" onChange="changePricingSelection(this, <%=thisID%>);" style="display: none;" disabled></select></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
			<%
		End If
		%>
		<input type="submit" name="submit" value="submit" /><br />
		<b>Form</b><hr />
		<%
		For Each frm In Request.Form
			%>
			<b><%=frm%></b>=<%=Request.Form(frm)%><br />
			<%
		Next
	End Function

	Function RenderProductSelection(id)
		%>
		<select name="product_<%=id%>" size="1" onChange="changeProductSelection(this, <%=id%>);">
			<option value="-1">... Select Product To Add To Order ...</option>
			<option value="352345">... Test ...</option>
			<% PrintProductOptions Fix(Request.Form("product_" & id)) %>
		</select>
		<%
	End Function
	
	Function ManualEmailHandler
		strFrom = Request.Form("from")
		strTo = Request.Form("to")
		strCC = Request.Form("cc")
		strBCC = Request.Form("bcc")
		strSubject = Request.Form("subject")
		strBody = Request.Form("body")
		strAction = Request.Form("action")
		If (strAction = "Send") Then
			Call SMTP(strFrom, strTo, strCC, strBCC, strSubject, strBody)
			Response.Write("Sent email to " & strTo & "")
		End If
		RenderEmailForm
	End Function
	
	Function RenderEmailForm
		%>
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td><b>From:</b></td>
				<td><input type="text" name="from" /></td>
			</tr>
			<tr>
				<td><b>To:</b></td>
				<td><input type="text" name="to" /></td>
			</tr>
			<tr>
				<td><b>CC:</b></td>
				<td><input type="text" name="cc" /></td>
			</tr>
			<tr>
				<td><b>BCC:</b></td>
				<td><input type="text" name="bcc" /></td>
			</tr>
			<tr>
				<td><b>Subject:</b></td>
				<td><input type="text" name="subject" /></td>
			</tr>
			<tr>
				<td><b>Body:</b></td>
				<td><textarea rows="4" cols="50" name="body"></textarea></td>
			</tr>
			<tr>
				<td></td>
				<td><input type="submit" value="Send" /><input type="hidden" name="action" value="Send" /></td>
			</tr>
		</table>
		<%
	End Function
		
	Function CodeEntryHandler
		openDB

		formProductID = Request.Form("productid")
		If (InStr(formProductID, ",") > 0) Then 
			formProductID = Split(formProductID, ",")
			insertSQL = "productid, altproductid"
			insertValuesSQL = formProductID(0) & ", " & formProductID(1)
		Else
			insertSQL = "productid"
			insertValuesSQL = formProductID
		End If

		formAction = Request.Form("action")
		If (formAction = "add") Then
			frmCodes = Split(Request.Form("codes"), vbNewline)
			If (Ubound(frmCodes) > -1) Then
				codesAdded = 0
				For Each code In frmCodes
					If (Len(Trim(code)) > 0) Then
						codesAdded = codesAdded + 1
						cartDB.Execute("INSERT INTO regcodes ( code, added, isUsed, used, ordersessionid, " & insertSQL & " ) VALUES ( '" & Trim(code) & "', date() + time(), false, null, null, " & insertValuesSQL & " )")
					End If
				Next
				strMessage = "Added " & codesAdded & " codes."
				redirect Request.ServerVariables("URL") & "?function=" & Request.QueryString("function") & "&message=" & strMessage & ""
			End If
		End If
		
		cycleDatabase
		closeDB

		RenderCodeEntry
	End Function
	
	Function RenderCodeEntry
		openDB

		Set dbProducts = cartDB.Execute("SELECT id, name FROM product WHERE id IN (SELECT productid FROM regcodes GROUP BY productid) ORDER BY id")
		If (dbProducts.EOF) Then
			%>
				<b>There are no products setup with registration codes.</b>
			<%
		Else
			%>
				<table border="0" cellpadding="3" cellspacing="0">
					<thead>
						<th style="background: black; color: white;">Product</th>
						<th style="background: black; color: white;">Available</th>
						<th style="background: black; color: white;">Used</th>
						<th style="background: black; color: white;">Total</th>
						<th style="background: black; color: white;">Downloads</th>
						<th style="background: black; color: white;">View Emails</th>
					</thead>
			<%
			productSelect = ""
			Do Until (dbProducts.EOF)
				Set dbProd = cartDB.Execute("SELECT productid, altproductid FROM regcodes WHERE (productid = " & dbProducts("id") & " OR altproductid = " & dbProducts("id") & ")")
				Set dbName1 = cartDB.Execute("SELECT name FROM product WHERE id = " & dbProd("productid") & "")
				productName = dbName1("name")

				If (isNumeric(dbProd("altproductid"))) Then
					Set dbName2 = cartDB.Execute("SELECT name FROM product WHERE id = " & dbProd("altproductid") & "")
					If Not (dbName2.EOF) Then
						productName = productName & "/" & dbName2("name")
						productSelect = productSelect & "<option value=""" & dbProd("productid") & "," & dbProd("altproductid") & """>" & productName & "</option>"
					Else
						productSelect = productSelect & "<option value=""" & dbProd("productid") & """>" & productName & "</option>"
					End If
				Else
					productSelect = productSelect & "<option value=""" & dbProd("productid") & """>" & productName & "</option>"
				End If
				
				Set dbCodes = cartDB.Execute("SELECT count(code) as codecount FROM regcodes WHERE (productid = " & dbProducts("id") & " OR altproductid = " & dbProducts("id") & ")")
				totalCodes = dbCodes("codecount")
				Set dbCodes = Nothing
				
				Set dbCodes = cartDB.Execute("SELECT count(code) as codecount FROM regcodes WHERE isused = true AND (productid = " & dbProducts("id") & " OR altproductid = " & dbProducts("id") & ")")
				usedCodes = dbCodes("codecount")
				Set dbCodes = Nothing
		
				Set dbCodes = cartDB.Execute("SELECT count(code) as codecount FROM regcodes WHERE isused = false AND (productid = " & dbProducts("id") & " OR altproductid = " & dbProducts("id") & ")")
				availableCodes = dbCodes("codecount")
				Set dbCodes = Nothing
		
				Set dbCodes = cartDB.Execute("SELECT count(id) as downloads FROM downloademails WHERE productid = " & dbProducts("id") & " OR productid = " & dbProd("altproductid") & "")
				downloads = dbCodes("downloads")
				Set dbCodes = Nothing
			%>
					<tr>
						<td><%=productName%></td>
						<td><%=availableCodes%></td>
						<td><%=usedCodes%></td>
						<td><%=totalCodes%></td>
						<td><%=downloads%></td>
						<td><a href="<%=Request.ServerVariables("URL")%>?function=<%=Request.QueryString("function")%>&viewid=">View Emails</a></td>
					</tr>
			<%
				dbProducts.MoveNext
			Loop
			
			%>
				</table>
			<%
		End If

		%>
		<br />
		<br />
		Product To Add Keys To: <select name="productid" size="1">
			<%=productSelect%>
		</select><br />
		<br />
		Code entry (one per line):<br />
		<textarea name="codes" cols="70" rows="15"></textarea><br />
		<input type="hidden" name="action" value="add" />
		<input type="submit" value="Add Codes" />
		<%
		closeDB
	End Function
	
	Function UltimateEmailsHandler
		RenderUltimateEmails
	End Function
	
	Function RenderUltimateEmails
		openDB
		Set dbEmails = cartDB.Execute("SELECT * FROM ultimateemails ORDER BY id")
		If Not (dbEmails.EOF) Then
			%>
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<th><b>Count</b></th>
						<th><b>Name</b></th>
						<th><b>Address</b></th>
						<th><b>Downloaded</b></th>
					</tr>
			<%
			Do Until (dbEmails.EOF)
				%>
					<tr>
						<td><%=dbEmails("id")%></td>
						<td><%=dbEmails("name")%></td>
						<td><%=dbEmails("email")%></td>
						<td><%=dbEmails("added")%></td>
					</tr>
				<%
				dbEmails.MoveNext
			Loop
			%>
				</table>
			<%
		End If
		Set dbEmails = Nothing
		closeDB
	End Function
		
' Admin Page
	If (Request.Cookies("loggedin") <> "vr8735slkdj8e#!421%2598326^#s") Then
		If (Request.Form("username") = "!!vasstcart$$") And (Request.Form("password") = "tooele947") Then
			Response.Cookies("loggedin") = "vr8735slkdj8e#!421%2598326^#s"
			redirect Request.ServerVariables("URL") & "?" & Request.QueryString
		End If
		%>
		<form action="<%=Request.ServerVariables("URL")%>?<%=Request.QueryString%>" method="post">
		Username: <input type="text" name="username" /><br />
		Password: <input type="password" name="password" /><br />
		<input type="submit" name="action" value="Login">
		</form>
		<%
	Else
		AdminHandler
	End If
	
	Function AdminHandler
		debug("Entering AdminHandler handler.")

		cartFunction = Trim(Request.QueryString("function"))
		cartMessage = Trim(Request.QueryString("message"))

		If (Len(cartMessage) > 0) Then
			bolMessage = True
		Else
			bolMessage = False
		End If

		If (Len(cartFunction) > 0) Then
			bolFunction = True
		Else
			bolFunction = False
		End If

		renderAdmin bolFunction, cartFunction, bolMessage, cartMessage
	End Function

	Function renderStyle
		%>
		<title> | Vasst Cart Manager | </title>
		<style>
		body, td, font, input, select {
			font-family: "Arial";
			font-size: 11px;
		}
		table {
			border-collapse: collapse;
		}
		a {
			color: black;
			text-decoration: underline;
		}
		a:hover {
			color: gray;
			text-decoration: none;
		}
		</style>
		<%
	End Function

	Function TestHandler
	bolDebug = true
		openDB
		Set dbVendors = cartDB.Execute(SQL("SELECT vendor.id, vendor.email, vendor.emailcc, vendor.emailbcc, count(orderdata.id) as productcount FROM product, orderdata, vendor WHERE vendor.id = product.vendorid AND product.id = orderdata.productid AND orderdata.ordersessionid = 26 GROUP BY vendor.id, vendor.email, vendor.emailcc, vendor.emailbcc"))		
		If (dbVendors.EOF) Then
			ReportError "There was a problem locating the vendor for a product."
		Else
			Do Until dbVendors.EOF
				Response.Write("Vendor ID: " & dbVendors("id") & "(" & dbVendors("productcount") & ")<br />")
				Response.Write("To: " & dbVendors("email") & " CC: " & dbVendors("emailcc") & " BCC: " & dbVendors("emailbcc") & "<br />")

				Set dbProducts = cartDB.Execute(SQL("SELECT product.name, orderdata.quantity FROM product, orderdata WHERE product.id = orderdata.productid AND product.vendorid = " & dbVendors("id") & " AND orderdata.ordersessionid = 26"))
				If (dbProducts.EOF) Then
					ReportError "There was a problem getting the product list for a perticular vendor."
				Else
					Do Until dbProducts.EOF
						Response.Write("Product Name: " & dbProducts("name") & " Quantity: " & dbProducts("quantity") & "<br />")
						dbProducts.MoveNext
					Loop
				End If
				dbVendors.MoveNext
			Loop
		End If		
		closeDB		
	End Function
	
	Function renderAdmin( bolHandler, strHandler, bolMessage, strMessage )
		If (Request.ServerVariables("REMOTE_ADDR") = "205.208.240.21") Then
			thisIsMe = True
		ElseIf (Request.ServerVariables("REMOTE_ADDR") = "205.208.240.201") Then
			thisIsMe = True
		ElseIf (Request.ServerVariables("REMOTE_ADDR") = "205.208.226.5") Then
			thisIsMe = True
		Else
			thisIsMe = False
		End If

		renderStyle

		%>
		<table border="1" cellpadding="0" cellspacing="0" width="100%" height="100%" style="border-collapse: collapse;">
			<form method="post" action="<%=Request.ServerVariables("URL")%>?function=<%=Request.QueryString("function")%>">
			<tr>
				<td background="/images/banner_back.gif" height="23" nowrap align="center">
					<font style="color: white"><b>Cart Manager</b></font>
				</td>
				<td background="/images/banner_back.gif">
					<% If (bolMessage) Then %>
						<font style="color: white;"><b>Status: </b><i><%=strMessage%></i></font><br />
					<% Else %>
						<!--Blank-->
					<% End If %>
				</td>
			</tr>

			<tr>
				<td width="150" bgcolor="#c0c0c0" valign="top" nowrap>
					<table border="0" cellpadding="2" cellspacing="0" width="100%">
					<tr>
						<td background="/images/global_back.gif"><font style="color: white;"><b>Categories</b></font></td>
					</tr>
					<tr>
						<td>
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=AddCategory">Add</a><br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=EditCategory">Edit</a><br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=DeleteCategory">Delete</a><br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=CategoryTree">Product Counts</a><br />
						</td>
					</tr>
					<tr>
						<td background="/images/global_back.gif"><font style="color: white;"><b>Products</b></font></td>
					</tr>
					<tr>
						<td>
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=AddProduct">Add</a><br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=EditProduct">Edit</a><br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=DeleteProduct">Delete</a><br />
							<br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=CodeEntry">Code Entry</a><br />
						</td>
					</tr>
					<tr>
						<td background="/images/global_back.gif"><font style="color: white;"><b>Orders</b></font></td>
					</tr>
					<tr>
						<td>
				<% If (thisIsMe) Then %>
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=ManualOrderEntry">Manual Order Entry</a><br />
				<% End If %>
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=ViewOrders">View Orders</a><br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=Spreadsheet">View Spreadsheet</a><br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=FulfillmentData">Fulfillment Data</a><br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=SalesReport">Sales Report</a><br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=UltimateEmails">Ultimate Emails</a><br />
						</td>
					</tr>
					<tr>
						<td background="/images/global_back.gif"><font style="color: white;"><b>Email</b></font></td>
					</tr>
					<tr>
						<td>
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=EmailTemplates">Edit Templates</a><br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=EmailHistory">View History</a><br />
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=ManualEmail">Manual Email</a><br />
						</td>
					</tr>
					<tr>
						<td background="/images/global_back.gif"><font style="color: white;"><b>Errors</b></font></td>
					</tr>
					<tr>
						<td>
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=ViewIncidents">View Incidents</a><br />
						</td>
					</tr>
				<% If (thisIsMe) Then %>
					<tr>
						<td background="/images/global_back.gif"><font style="color: white;"><b>Admin</b></font></td>
					</tr>
					<tr>
						<td>
							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=SQL">SQL</a><br />
<!--							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=ReportError">Generate Error</a><br />-->
<!--							&nbsp;<a href="<%=Request.ServerVariables("URL")%>?function=ResetDatabase">Reset Database</a><br />-->
						</td>
					</tr>
				<% End If %>
				</table>
			</td>
			<td valign="top">
				<% 
					If (bolHandler) Then
						debug("Calling handler for <b>" & strHandler & "</b>.")
						Eval(strHandler & "Handler")
					End If
				%>
			</td>
		</tr>
		<input type="hidden" name="process" value="true">
	</form>
	</table>
	<%
	End Function
%>
