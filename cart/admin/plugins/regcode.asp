<%
'<!--#include virtual="/cart/admin/includes.asp"-->
'' HEADER ''
'' NAME: regcode.asp
'' DESCRIPTION: Ultimate Vegas Scripting Tool Registration
'' ACTIVE: true
'' FUNCTION: DoRegistrationLaunch
'' PARAMETERS: plugin_ordersessionid
'' HEADER END ''

'If (VarType(Session("plugin_ordersessionid")) = vbNull) Then
'	Response.Write("<b>There was an error executing a plugin.</b><br />")
'	Call ReportError("There was a problem, the plugin_ordersessionid value was null.")
'	Response.End()
'Else
'	Call DoRegistrationLaunch(Session("plugin_ordersessionid"))
'End If

'Call DoRegistrationLaunch(GetOrderSessionId)

Function DoRegistrationLaunch(ordersessionid)
	If (ordersessionid > 0) Then
		openDB
		
	'	Response.Write("Order ID: " & ordersessionid & " Email: " & GetInfoFromOrder("emailaddress",ordersessionid) & "<br />")
		Set dbProducts = cartDB.Execute("SELECT productid, quantity, quantity FROM orderdata WHERE ordersessionid = " & ordersessionid & "")
		If Not (dbProducts.EOF) Then
			Do Until (dbProducts.EOF)
				Set dbNeedKey = cartDB.Execute("SELECT count(*) as keycounts FROM regcodes WHERE productid = " & dbProducts("productid") & " OR altproductid = " & dbProducts("productid") & "")
				If (Fix(dbNeedKey("keycounts")) > 0) Then
					thisProduct = GetProductInfo("name", dbProducts("productid"))
'					Response.Write("Product: " & thisProduct & " ID: " & dbProducts("productid") & "<br />")
					Set dbCheckAlready = cartDB.Execute("SELECT count(code) as codecount FROM regcodes WHERE ordersessionid = " & ordersessionid & "")
					alreadyAssigned = dbCheckAlready("codecount")
					Set dbCheckAlready = Nothing
	
'					Response.Write("Already assigned codes: " & alreadyAssigned & "<br />")
					If (alreadyAssigned > 0) Then
						Set dbCodes = cartDB.Execute("SELECT code FROM regcodes WHERE ordersessionid = " & ordersessionid & " AND (productid = " & dbProducts("productid") & " OR altproductid = " & dbProducts("productid") & ")")
						origKeys = ""
						Do Until dbCodes.EOF
							origKeys = origKeys & dbCodes("code") & vbNewline
							dbCodes.MoveNext
						Loop
						Set dbCodes = Nothing
						
						Call SendKey(GetInfoFromOrder("emailaddress",ordersessionid), thisProduct, origKeys, ordersessionid)
					Else
				
		'			Response.Write("Checking ID# " & dbProducts("productid") & " """ & thisProduct & """<br />")
						Set dbKeys = cartDB.Execute("SELECT count(*) as availablekeys FROM regcodes WHERE isUsed = false AND productid = " & dbProducts("productid") & " OR altproductid = " & dbProducts("productid") & "")
						availableKeys = Fix(dbKeys("availablekeys"))
						Set dbKeys = Nothing
		
						If (availableKeys <= 0) Then
							Call SendNoKeyWarning("support@vasst.com", thisProduct)
						Else
							If (availableKeys - Fix(dbProducts("quantity")) <= 0) Then
								Call SendNoKeyWarning("support@vasst.com", thisProduct)
							ElseIf (availableKeys <= 10) Then
								Call SendLowKeyWarning("support@vasst.com", availableKeys, thisProduct)
							End If
							
							productKeys = ""
							For X = 0 To Fix(dbProducts("quantity")) - 1
								Set dbNextKey = cartDB.Execute("SELECT min(id) as nextid FROM regcodes WHERE isUsed = false AND (productid = " & dbProducts("productid") & " OR altproductid = " & dbProducts("productid") & ")")
								If (dbNextKey.EOF) Then
		'							Response.Write("<b>FAILED ON GETTING A KEY FOR THIS PRODUCT.</b><br />")
								Else
									Set dbProdName = cartDB.Execute("SELECT name FROM product WHERE id = " & dbProducts("productid") & "")
									Set dbKey = cartDB.Execute("SELECT code FROM regcodes WHERE id = " & dbNextKey("nextid") & "")
									productKeys = productKeys & dbProdName("name") & " -- " & dbKey("code") & vbNewline
									cartDB.Execute("UPDATE regcodes SET isused = true, used = DATE() + TIME(), ordersessionid = " & ordersessionid & " WHERE id = " & dbNextKey("nextid") & "")
									cycleDatabase
								End If
							Next
							Call SendKey(GetInfoFromOrder("emailaddress",ordersessionid), thisProduct, productKeys, ordersessionid)
						End If
					End If
				End If
				Set dbNeedKey = Nothing
						
				dbProducts.MoveNext
			Loop
		End If
		Set dbProducts = Nothing
	
		closeDB
	End If
End Function

Function SendKey(toAddress, productTitle, keyCode, ordersessionid)
'	Response.Write("Sending key """ & keyCode & """ for """ & productTitle & """.<br />")
	openDB
	Set dbEmail = cartDB.Execute(SQL("SELECT subject, body FROM email WHERE name = 'Registration Code'"))
	If (dbEmail.EOF) Then
		ReportError "There was getting the from the database."
	Else
		mSubject = ParseAndReplaceEmail(dbEmail("subject"),ordersessionid,-1)
		mMessage = ParseAndReplaceEmail(dbEmail("body"),ordersessionid,-1)
		mMessage = Replace(mMessage, "%REGISTRATIONCODES%", keyCode)
	End If
	Set dbEmail = Nothing
	closeDB
		
	Call SMTP(strEmailFrom, toAddress, "", "", mSubject, mMessage)
End Function

Function SendLowKeyWarning(toAddress, availableKeys, productTitle)
'	Response.Write("Low key warning.<br />")

	Call SMTP(strEmailFrom, toAddress, "", "", "Low Product Key Warning [" & productTitle & "]", "We just sent out a key for """ & productTitle & """, there are only " & (availableKeys - 1) & "key(s) left.")
End Function

Function SendNoKeyWarning(toAddress, productTitle)
'	Response.Write("No key warning.<br />")
	
	Call SMTP(strEmailFrom, toAddress, "", "", "NO KEY WARNING [" & productTitle & "]", "There are not keys left for """ & productTitle & """, please add some more keys.")
End Function

%>