<%
'-------------------------------------------------------------------------
'	Name: 		kbCategories class
'	Purpose: 	encapsulate functions for managing and displaying categories
'Modifications:
'	Date:		Name:	Description:
'	1/8/02		JEA		Creation
'-------------------------------------------------------------------------
Class kbCategories
	'Private Sub Class_Initialize()
	'End Sub
	
	'Private Sub Class_Terminate()
	'End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteOptionList()
	'	Purpose: 	write list of categories
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/02		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteOptionList(ByVal V_sFieldName, ByVal v_lSelectedID, ByVal v_bNew, ByVal v_sOptions)
		dim sQuery
		sQuery = "SELECT lCategoryID, vsCategoryName FROM tblCategories ORDER BY vsCategoryName"
		with response
			.write "<select name='"
			.write v_sFieldName
			.write "' "
			.write v_sOptions
			.write ">"
			if v_bNew then .write "<option value='0'>--add new--"
			.write MakeList(sQuery, v_lSelectedID)
			.write "</select>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteJSArray()
	'	Purpose: 	write list of categories
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/02		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteJSArray()
		dim sQuery
		dim oData
		sQuery = "SELECT C.lCategoryID, C.vsCategoryName, CIT.lItemTypeID " _
			& "FROM tblCategories C INNER JOIN tblCategoryItemTypes CIT " _
			& "ON C.lCategoryID = CIT.lCategoryID ORDER BY C.lCategoryID"
		Set oData = New kbDataAccess
		response.write oData.GetJSArray(sQuery)
		Set oData = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteOptionList()
	'	Purpose: 	write list of categories
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/02		JEA		Creation
	'	7/23/04		JEA		Add message
	'-------------------------------------------------------------------------
	Public Sub Save(ByVal v_lCategoryID, ByVal v_sCategoryName, ByVal v_sCategoryItems)
		dim sQuery
		dim oData
		dim oRS
		dim aItems
		dim sMessage
		dim x
		
		Set oData = New kbDataAccess
		Call oData.BeginTrans()
		
		If IsNumber(v_lCategoryID) Then
			' update existing category
			sQuery = "UPDATE tblCategories SET vsCategoryName = '" & CleanForSQL(v_sCategoryName) _
				& "' WHERE lCategoryID = " & v_lCategoryID
			Call oData.ExecuteOnly(sQuery)
			' clear old item associations
			sQuery = "DELETE FROM tblCategoryItemTypes WHERE lCategoryID = " & v_lCategoryID
			Call oData.ExecuteOnly(sQuery)
			
			sMessage = "The """ & v_sCategoryName & """ category has been updated"
		Else
			' insert new category
			Set oRS = Server.CreateObject("ADODB.Recordset")
			With oRS
				.Open "tblCategories", oData.Connection, adOpenStatic, adLockOptimistic, adCmdTable
				.AddNew
				.Fields("vsCategoryName") = CleanForSQL(v_sCategoryName)
				.Update
				v_lCategoryID = .Fields("lCategoryID")
				.Close
			End With
			Set oRS = nothing
			
			sMessage = "The category """ & v_sCategoryName & """ has been added"
		End If
		
		aItems = Split(v_sCategoryItems, ",")
		for x = 0 to UBound(aItems)
			sQuery = "INSERT INTO tblCategoryItemTypes (lCategoryID, lItemTypeID) VALUES (" _
				& v_lCategoryID & ", " & aItems(x) & ")"
			Call oData.ExecuteOnly(sQuery)
		next
		Call oData.CommitTrans()
		Call ClearFilterCache(g_FILTER_CATEGORY)
		Call SetSessionValue(g_USER_MSG, sMessage)
		Set oData = Nothing
	End Sub
End Class
%>