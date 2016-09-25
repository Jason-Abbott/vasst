<!--#include virtual="/cart/admin/includes.asp"-->
<b>Fulfillment Batch</b><br />
<%
'	Response.Write("Aborted.")
'	Response.End()
	
'	Dim aDay : aDay = Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat")
	Dim aDay : aDay = Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
	Dim sToday : sToday = aDay(DatePart("w",Now)-1)
	Dim sHour : sHour = Fix(DatePart("h",Now))
	
	Dim sIncludeInFulfillment : sIncludeInFulfillment = "AND hidden = false" '"AND includeinfulfillment = true"
	
	Dim bTest : bTest = True		' Do a test run and output entries to fulfill, don't actually mark fulfilled or send to fulfillment house.
	
	Response.Write("Today is " & sToday & ".<br />")
	Response.Write("The hour is " & sHour & ".<br />")
	
	' Ground orders, twice weekly, 6 AM MST.
	If (((sToday = "Tuesday") Or (sToday = "Friday")) And ((sHour = 6) Or (sHour = 5) Or (sHour = 7))) Or (Request.QueryString("function") = "ground") Then
		Response.Write("Checking for ground orders only.<br />")
		iVendorID = 2
		bWithGround = True
'		createAndMailReport 2, True

	' Expedited orders, every day, 12 PM MST.
	ElseIf ((sHour = 12) Or (sHour = 11) Or (sHour = 13)) Or (Request.QueryString("function") = "expedited") Then
		Response.Write("Checking for any expedited orders only.<br />")
		iVendorID = 2
		bWithGround = False
'		createAndMailReport 2, False

	Else
		Response.Write("It is not time to run any orders.<br />")
		Response.End()
	End If
'	printNeededToBeFulfilled

	iOrdersFound = Fix(getOrderCount(iVendorID, bWithGround))
	Response.Write("Found a total of " & iOrdersFound & " products needing to be fulfilled.<br />")
	
	If (iOrdersFound > 0) Then
		Response.Write("Generating spreadsheet and mailing.<br />")
		createAndMailReport iVendorID, bWithGround
	Else
		Response.Write("Since there are no products that need to be fulfilled, no action was taken.<br />")
	End If

	Function getOrderCount(iVendorID, bWithGround)
		' If true, do all.  If false, do everything but ground.
		Dim sAndGround
		If (bWithGround) Then
			sAndGround = " AND shippingmethod = 'ground' "
		Else
			sAndGround = " AND shippingmethod <> 'ground' "
		End If
		sSQLQuery = "SELECT count(orderdata.id) as orderCount FROM orderdata LEFT JOIN ordersession ON ordersession.id = orderdata.ordersessionid WHERE orderdata.productid IN (SELECT id FROM product WHERE deleted = false AND vendorid = " & iVendorID & " " & sIncludeInFulfillment & ") AND ordersession.ispaid = true AND orderdata.senttofulfillment = false " & sAndGround & ""
		
		openDB
		Set dbOrders = cartDB.Execute(sSQLQuery)
		If (dbOrders.EOF) Then
			getOrderCount = 0
		Else
			getOrderCount = Fix(dbOrders("orderCount"))
		End If
		closeDB
	End Function

	Function createAndMailReport(iVendorID, bWithGround)
		' If true, do all.  If false, do everything but ground.
		Dim sAndGround
		Dim sFileName
		If (bWithGround) Then
			sAndGround = " AND shippingmethod = 'ground' "
			sFileName = "_ground"
		Else
			sAndGround = " AND shippingmethod <> 'ground' "
			sFileName = "_expedited"
		End If
		
		Dim sStamp : sStamp = DatePart("yyyy",Now) & "-"
		If (DatePart("m",Now) < 10) Then
			sStamp = sStamp & "0"
		End If
		sStamp = sStamp & DatePart("m",Now) & "-"
		If (DatePart("d",Now) < 10) Then
			sStamp = sStamp & "0"
		End If
		sStamp = sStamp & DatePart("d",Now)
		
		Dim sDir : sDir = Server.MapPath(".") & "\data\excel\"
		Dim sSrc : sSrc = "_blank.xls"
		Dim sDest : sDest = "fulfillment_" & sStamp & "" & sFileName & ".xls"

		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")

'		Response.Write("Dir: " & sDir & "<br />")
'		Response.Write("Src: " & sSrc & "<br />")
'		Response.Write("Dest: " & sDest & "<br />")
'		Response.Write("Excel: " & sDir & sDest & "<br />")
'		Response.Write("<a href=""/cart/admin/data/excel/" & sDest & """>Download Report</a><br />")

		Dim sVendorName : sVendorName = ""
		Dim sVendorTo : sVendorName = ""
		Dim sVendorCC : sVendorName = ""
		Dim sVendorBCC : sVendorName = ""

		openDB
		Set dbVendor = cartDB.Execute("SELECT * FROM vendor WHERE id = " & iVendorID & "")
		If (dbVendor.EOF) Then
			Response.Write("There was no vendor by the id# " & iVendorID & "<br />")
			Response.End()
		Else
			sVendorName = dbVendor("name")
			sVendorTo = dbVendor("email")
			sVendorCC = dbVendor("emailcc")
			sVendorBCC = dbVendor("emailbcc")
			
'			Response.Write("Vendor: " & dbVendor("company") & "<br />")
'			Response.Write("Name: " & dbVendor("name") & "<br />")
'			Response.Write("To: " & dbVendor("email") & "<br />")
'			Response.Write("CC: " & dbVendor("emailcc") & "<br />")
'			Response.Write("BCC: " & dbVendor("emailbcc") & "<br />")
		End If
		closeDB

		If (oFSO.FileExists(sDir & sDest)) Then
			oFSO.DeleteFile sDir & sDest, True
		End If
		oFSO.CopyFile sDir & sSrc, sDir & sDest

		Set oRS = CreateObject("ADODB.Recordset")
		oRS.Open "Select * from [Fulfillment$B1:IV29]", _
				 "Provider=Microsoft.Jet.OLEDB.4.0;" & _
				 "Data Source=" & sDir & sDest & ";" & _
				 "Extended Properties=""Excel 8.0;HDR=NO;""", 1, 3

		openDB
		Dim iPos : iPos = 0
		
		'On Error Resume Next
'		oRS.Fields.Item(iPos).Value = "TRACKING" : iPos = iPos + 1
		oRS.Fields.Item(iPos).Value = "Invoice" : iPos = iPos + 1
		oRS.Fields.Item(iPos).Value = "Ship Method" : iPos = iPos + 1
		oRS.Fields.Item(iPos).Value = "Name" : iPos = iPos + 1
		oRS.Fields.Item(iPos).Value = "Address" : iPos = iPos + 1
		oRS.Fields.Item(iPos).Value = "Address 2" : iPos = iPos + 1
		oRS.Fields.Item(iPos).Value = "City" : iPos = iPos + 1
		oRS.Fields.Item(iPos).Value = "State" : iPos = iPos + 1
		oRS.Fields.Item(iPos).Value = "Zip" : iPos = iPos + 1
		oRS.Fields.Item(iPos).Value = "Country" : iPos = iPos + 1
		oRS.Fields.Item(iPos).Value = "Phone Number" : iPos = iPos + 1

		Set dbProducts = cartDB.Execute(SQL("SELECT name FROM product WHERE deleted = false AND vendorid = " & iVendorID & " " & sIncludeInFulfillment & " ORDER BY id"))
		Do Until dbProducts.EOF
			oRS.Fields(iPos).Value = dbProducts("name") : iPos = iPos + 1
			dbProducts.MoveNext
		Loop
		Set dbProducts = Nothing

'		oRS.MoveNext
		oRS.Update

		' Loop through orders.
'		Set dbOrders = cartDB.Execute(SQL("SELECT ordersession.id FROM ordersession, orderdata WHERE ordersession.id = orderdata.ordersessionid AND ordersession.ispaid = True AND orderdata.senttofulfillment = false " & sAndGround & " GROUP BY ordersession.id ORDER BY ordersession.id"))
		Set dbOrders = cartDB.Execute(SQL("SELECT ordersession.id FROM orderdata LEFT JOIN ordersession ON ordersession.id = orderdata.ordersessionid WHERE orderdata.productid IN (SELECT id FROM product WHERE deleted = false AND vendorid = " & iVendorID & " " & sIncludeInFulfillment & ") AND ordersession.ispaid = true AND orderdata.senttofulfillment = false " & sAndGround & " GROUP BY ordersession.id ORDER BY ordersession.id"))
		Do Until dbOrders.EOF
'			Set dbProduct = cartDB.Execute(SQL("SELECT product.name, product.sku, pricing.discountcode FROM product, pricing WHERE product.id = pricing.productid AND product.id = " & Fix(dbOrders("productid")) & " AND pricing.id = " & Fix(dbOrders("pricingid")) & ""))
			Set dbOrderIDs = cartDB.Execute(SQL("SELECT id, created, totalcost, shippingcost, shippingmethod, shippingid, billingid FROM ordersession WHERE id = " & dbOrders("id") & ""))
			Set dbShipping = cartDB.Execute(SQL("SELECT * FROM ordershipping WHERE id = " & dbOrderIDs("shippingid") & ""))
			Set dbBilling = cartDB.Execute(SQL("SELECT * FROM orderbilling WHERE id = " & dbOrderIDs("billingid") & ""))

			iPos = 0
			
			oRS.AddNew
			'oRS.Fields.Item(iPos).Value = "" : iPos = iPos + 1
			oRS.Fields.Item(iPos).Value = dbOrderIDs("id") : iPos = iPos + 1
			oRS.Fields.Item(iPos).Value = dbOrderIDs("shippingmethod") : iPos = iPos + 1
			oRS.Fields.Item(iPos).Value = dbShipping("name") : iPos = iPos + 1
			oRS.Fields.Item(iPos).Value = dbShipping("address") : iPos = iPos + 1
			oRS.Fields.Item(iPos).Value = dbShipping("address2") : iPos = iPos + 1
			oRS.Fields.Item(iPos).Value = dbShipping("city") : iPos = iPos + 1
			oRS.Fields.Item(iPos).Value = dbShipping("state") : iPos = iPos + 1
			oRS.Fields.Item(iPos).Value = dbShipping("zip") : iPos = iPos + 1
			oRS.Fields.Item(iPos).Value = dbShipping("country") : iPos = iPos + 1
			oRS.Fields.Item(iPos).Value = dbBilling("phonenumber") : iPos = iPos + 1

			' Loop through product quantities.
			Set dbProducts = cartDB.Execute(SQL("SELECT id FROM product WHERE deleted = false AND vendorid = " & iVendorID & " " & sIncludeInFulfillment & " ORDER BY id"))
			Do Until dbProducts.EOF
				Set dbQuantity = cartDB.Execute(SQL("SELECT quantity FROM orderdata WHERE ordersessionid = " & dbOrders("id") & " AND productid = " & dbProducts("id") & ""))
				If (dbQuantity.EOF) Then
					oRS.Fields.Item(iPos).Value = "" : iPos = iPos + 1
				Else
					If (Fix(dbQuantity("quantity")) > 0) Then
'						Response.Write("UPDATE orderdata SET senttofulfillment = true WHERE ordersessionid = " & dbOrders("id") & " AND productid = " & dbProducts("id") & "<br />")
						cartDB.Execute(SQL("UPDATE orderdata SET senttofulfillment = true, datesenttofulfillment = #" & date & " " & time & "# WHERE ordersessionid = " & dbOrders("id") & " AND productid = " & dbProducts("id") & ""))
						oRS.Fields.Item(iPos).Value = dbQuantity("quantity") : iPos = iPos + 1
					Else
						oRS.Fields.Item(iPos).Value = "" : iPos = iPos + 1
					End If
				End If
				Set dbQuantity = Nothing
				dbProducts.MoveNext
			Loop
			Set dbProducts = Nothing
			
			oRS.Update
			oRS.MoveNext
			dbOrders.MoveNext
		Loop

		closeDB

'		oRS.Update
	
		oRS.Close
		Set oRS = Nothing
		
		Response.Write("Marked all new orders as sent to fulfillment.<br />")
		Response.Write("Generated spreadsheet, you can download it by <a href=""/cart/admin/data/excel/" & sDest & """>clicking here</a>.<br />")

		SMTPAttachment _
			"support@vasst.com", _ 
			sVendorTo, _
			sVendorCC, _
			sVendorBCC, _
			"Vendor Fulfillment Update - " & sStamp & "", _
			sVendorName & "," & vbNewline & _
				vbNewline & _
				"Here is an Excel spreadsheet containing the orders that we need to be fulfilled." & vbNewline & _
				vbNewline & _
				"Thanks," & vbNewline & _
				"VASST.com Management", _
			sDir & sDest

'		SMTPAttachment _
'			"support@vasst.com", _ 
'			"syntax@modpro.net", _
'			"", _
'			"", _
'			"Vendor Fulfillment Update - " & sStamp & "", _
'			sVendorName & "," & vbNewline & _
'				vbNewline & _
'				"Here is an Excel spreadsheet containing the orders that we need to be fulfilled." & vbNewline & _
'				vbNewline & _
'				"Thanks," & vbNewline & _
'				"VASST.com Management", _
'			sDir & sDest
			
		Response.Write("Mailed spreadsheet...<br />")
		Response.Write("&nbsp;&nbsp;&nbsp;To: " & sVendorTo & "<br />")
		Response.Write("&nbsp;&nbsp;&nbsp;CC: " & sVendorCC & "<br />")
		Response.Write("&nbsp;&nbsp;&nbsp;BCC: " & sVendorBCC & "<br />")
			
'		If (oFSO.FileExists(sDir & sDest)) Then
'			oFSO.DeleteFile sDir & sDest, True
'		End If
	End Function	
%>