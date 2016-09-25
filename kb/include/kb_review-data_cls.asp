<%
'-------------------------------------------------------------------------
'	Name: 		kbReviewData class
'	Purpose: 	methods for saving and retrieving review data
'Modifications:
'	Date:		Name:	Description:
'	7/21/04		JEA		Copied from tutorial data class
'-------------------------------------------------------------------------
Class kbReviewData
	Private m_sBaseSQL
	Private m_oData

	Private Sub Class_Initialize()
		m_sBaseSQL = "SELECT T.lReviewID, T.vsReviewURL, T.vsReviewName, " _
			& "T.vsDescription, T.dtApproveDate, IIf(U.lUserID IS NULL, 0, U.lUserID), " _
			& "IIf(U.vsScreenName IS NULL, " _
			& "U.vsFirstName + ' ' + U.vsLastName, U.vsScreenName), CR.fRank, T.lSubmitterID " _
			& "FROM ((tblReviews T LEFT JOIN tblUsers U ON T.lSubmitterID = U.lUserID) " _
			& "INNER JOIN (SELECT lItemID FROM tblItemSites WHERE lItemTypeID = " _
			& g_ITEM_REVIEW & " AND lSiteID = " & GetSessionValue(g_USER_SITE) _
			& ") tIS ON tIS.lItemID = T.lReviewID) " _
			& "LEFT JOIN (SELECT lItemID, fRank FROM tblComputedRankings WHERE lItemTypeID = " _
			& g_ITEM_REVIEW & ") CR ON CR.lItemID = T.lReviewID "
		Set m_oData = New kbDataAccess
	End Sub
	
	Private Sub Class_Terminate()
		Set m_oData = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetItem()
	'	Purpose: 	get review data
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Public Function GetItem(ByVal v_lReviewID, ByVal v_bOwnerOnly)
		dim sReviewSQL
		dim sCatSQL
		dim aData
		
		sReviewSQL = m_sBaseSQL & "WHERE T.lReviewID = " & v_lReviewID
		If v_bOwnerOnly And CStr(GetSessionValue(g_USER_TYPE)) <> CStr(g_USER_ADMIN) Then
			sReviewSQL = sReviewSQL & " AND T.lAuthorID = " & GetSessionValue(g_USER_ID)
		End If
		sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
				& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
				& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_REVIEW _
				& " AND IC.lItemID = " & v_lReviewID
		aData = m_oData.GetArray(sReviewSQL)
		aData = JoinArray(aData, m_REVIEW_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
		GetItem = aData
	End Function

	'-------------------------------------------------------------------------
	'	Name: 		GetPublic()
	'	Purpose: 	get review data
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Public Function GetPublic(ByVal v_aFilter)
		dim sReviewSQL
		dim sCatSQL
		dim aData
		dim sKey
		
		sKey = MakeKey(g_ITEM_REVIEW, v_aFilter)
		aData = Application(sKey)
		If Not IsArray(aData) or true Then 'DEBUG
			sReviewSQL = m_sBaseSQL & "WHERE T.lStatusID = " & g_STATUS_APPROVED & " "
			If v_aFilter(g_FILTER_AUTHOR) <> 0 Then
				sReviewSQL = sReviewSQL & "AND T.lAuthorID = " & v_aFilter(g_FILTER_AUTHOR) & " "
			End If
			If v_aFilter(g_FILTER_SOFTWARE) <> 0 Then
				sReviewSQL = sReviewSQL & "AND SV.lSoftwareID = " & v_aFilter(g_FILTER_SOFTWARE) & " "
			End If
			sReviewSQL = sReviewSQL & MakeSortSQL(v_aFilter(g_FILTER_SORT))
			sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
					& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
					& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_REVIEW
			aData = m_oData.GetArray(sReviewSQL)
			'response.Write sReviewSQL : response.end
			aData = JoinArray(aData, m_REVIEW_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
			If v_aFilter(g_FILTER_CATEGORY) <> 0 Then
				aData = FilterArray(aData, m_REVIEW_CATS, g_CAT_ID, v_aFilter(g_FILTER_CATEGORY))
			End If
			Application.Lock
			Application(sKey) = aData
			Application.Unlock
		End If
		GetPublic = aData
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		MakeSortSQL()
	'	Purpose: 	generate SQL for sorting reviews
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Private Function MakeSortSQL(ByVal v_lSortID)
		dim sSQL
		sSQL = "ORDER BY "
		select case v_lSortID
			case g_SORT_NAME_ASC
				sSQL = sSQL & "T.vsReviewName"
			case g_SORT_NAME_DESC
				sSQL = sSQL & "T.vsReviewName DESC"
			case g_SORT_DATE_ASC
				sSQL = sSQL & "T.dtApproveDate, T.vsReviewName"
			case g_SORT_DATE_DESC
				sSQL = sSQL & "T.dtApproveDate DESC, T.vsReviewName"
			case g_SORT_OWNER_ASC
				sSQL = sSQL & "U.vsFirstName, U.vsLastName, T.vsReviewName"
			case g_SORT_OWNER_DESC
				sSQL = sSQL & "U.vsFirstName DESC, U.vsLastName DESC, T.vsReviewName"
			case else
				sSQL = sSQL & "T.vsReviewName"
		end select
		MakeSortSQL = sSQL
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetArray()
	'	Purpose: 	get review data, one way or another
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Public Function GetArray()
		dim aData
		dim x
		dim aReturn(9)
		
		aReturn(m_REVIEW_ID) = Trim(Request.QueryString("id"))
		If Request.Form("fldName") <> "" Then
			With Request
				aReturn(m_REVIEW_NAME) = Trim(.Form("fldName"))
				aReturn(m_REVIEW_URL) = .Form("fldURL")
				aReturn(m_REVIEW_TEXT) = .Form("fldDescription")
				aReturn(m_REVIEW_AUTHOR_ID) = .Form("fldAuthor")
				aReturn(m_REVIEW_CATS) = .Form("fldCategories")
			End With
		ElseIf IsNumber(aReturn(m_REVIEW_ID)) Then
			aData = GetItem(aReturn(m_REVIEW_ID), false)
			If IsArray(aData) Then
				for x = 0 to UBound(aData)
					aReturn(x) = aData(x, 0)
				next
			End If
		End If
		aReturn(m_REVIEW_AUTHOR_ID) = ReplaceNull(aReturn(m_REVIEW_AUTHOR_ID), GetSessionValue(g_USER_ID))
		GetArray = aReturn
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		Save()
	'	Purpose: 	save review to database
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Public Sub Save(ByVal v_aData)
		dim oRS
		dim sQuery
		dim lReviewID
		dim sMessage
		
		lReviewID = v_aData(m_REVIEW_ID)
		Call m_oData.BeginTrans()
		
		If IsNumber(lReviewID) Then
			' update existing
			sQuery = "UPDATE tblReviews SET " _
				& "vsReviewURL = '" & Replace(v_aData(m_REVIEW_URL), "http://", "") & "', " _
				& "lAuthorID = " & v_aData(m_REVIEW_AUTHOR_ID) & ", " _
				& "vsReviewName = '" & CleanForSQL(v_aData(m_REVIEW_NAME)) & "', " _
				& "vsDescription = '" & CleanForSQL(v_aData(m_REVIEW_TEXT)) & "' " _
				& "WHERE lReviewID = " & lReviewID
			Call m_oData.ExecuteOnly(sQuery)
			Call ClearCache(g_ITEM_REVIEW)
			
			sMessage = "The review has been saved"
		Else
			' insert new
			Set oRS = Server.CreateObject("ADODB.Recordset")
			With oRS
				.Open "tblReviews", m_oData.Connection, adOpenStatic, adLockOptimistic, adCmdTable
				.AddNew
				.Fields("vsReviewURL") = Replace(v_aData(m_REVIEW_URL), "http://", "")
				.Fields("lAuthorID") = v_aData(m_REVIEW_AUTHOR_ID)
				.Fields("lSubmitterID") = GetSessionValue(g_USER_ID)
				.Fields("vsReviewName") = v_aData(m_REVIEW_NAME)
				.Fields("vsDescription") = v_aData(m_REVIEW_TEXT)
				.Fields("dtSubmitDate") = Now()
				.Fields("lStatusID") = g_STATUS_PENDING
				.Fields("bLocalFile") = false
				.Update
				lReviewID = .Fields("lReviewID")
				.Close
			End With
			Set oRS = Nothing
			
			sQuery = "INSERT INTO tblItemSites (lSiteID, lItemTypeID, lItemID) VALUES (" _
				& GetSessionValue(g_USER_SITE) & ", " _
				& g_ITEM_REVIEW & ", " _
				& lReviewID & ")"
			Call m_oData.ExecuteOnly(sQuery)
			
			sMessage = "The review has been submitted for approval"
		End If
		
		Call SaveCategories(lReviewID, Trim(v_aData(m_REVIEW_CATS)), g_ITEM_REVIEW)
		Call m_oData.CommitTrans()
		Call SetSessionValue(g_USER_MSG, sMessage)
		response.redirect "kb_reviews.asp"
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetPending()
	'	Purpose: 	get array of uploaded reviews
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Public Function GetPending()
		dim sQuery
		dim aData
		sQuery = Replace(m_sBaseSQL, "T.dtApproveDate", "T.dtSubmitDate") _
			& "WHERE T.lStatusID = " & g_STATUS_PENDING _
			& " ORDER BY T.dtSubmitDate"
		GetPending = m_oData.GetArray(sQuery)
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		ApprovePending()
	'	Purpose: 	update review status
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Public Sub ApprovePending(ByVal v_lReviewID)
		dim sQuery
		sQuery = "UPDATE tblReviews SET lStatusID = " & g_STATUS_APPROVED _
			& ", dtApproveDate = " & g_sSQL_DATE_DELIMIT & Now() & g_sSQL_DATE_DELIMIT _
			& " WHERE lReviewID = " & v_lReviewID
		Call m_oData.ExecuteOnly(sQuery)
		Call ClearCache(g_ITEM_TUTORIAL)
		'Call m_oData.LogActivity(g_ACT_APPROVE_UPLOAD, "", "", "", "", "")
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		DenyPending()
	'	Purpose: 	update review status
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Public Sub DenyPending(ByVal v_lReviewID)
		dim sQuery
		sQuery = "UPDATE tblReviews SET lStatusID = " & g_STATUS_REJECTED _
			& " WHERE lReviewID = " & v_lReviewID
		Call m_oData.ExecuteOnly(sQuery)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		SaveCategories()
	'	Purpose: 	save review category data to database
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Public Sub SaveCategories(ByVal v_lItemID, ByVal v_sCategories, ByVal v_lItemType)
		dim aCategories
		dim sQuery
		dim x
		
		If v_sCategories <> "" Then
			' delete any old plugins listed
			sQuery = "DELETE FROM tblItemCategories WHERE lItemID = " & v_lItemID _
				& " AND lItemTypeID = " & v_lItemType
			Call m_oData.ExecuteOnly(sQuery)
			
			' insert new plugins
			sQuery = "INSERT INTO tblItemCategories (lItemID, lItemTypeID, lCategoryID) " _
				& "VALUES (" & v_lItemID & ", " & v_lItemType & ", "
			aCategories = Split(v_sCategories, ",")
			for x = 0 to UBound(aCategories)
				Call m_oData.ExecuteOnly(sQuery & Trim(aCategories(x)) & ")")
			next
		End If
	End Sub
End Class
%>