<%
'-------------------------------------------------------------------------
'	Name: 		kbTutorialData class
'	Purpose: 	methods for saving and retrieving tutorial data
'Modifications:
'	Date:		Name:	Description:
'	1/4/02		JEA		Creation
'	4/25/03		JEA		Add site condition
'-------------------------------------------------------------------------
Class kbTutorialData
	Private m_sBaseSQL
	Private m_oData

	Private Sub Class_Initialize()
		m_sBaseSQL = "SELECT T.lTutorialID, T.vsTutorialURL, T.vsTutorialName, " _
			& "T.vsDescription, T.dtApproveDate, IIf(U.lUserID IS NULL, 0, U.lUserID), " _
			& "IIf(U.vsScreenName IS NULL, " _
			& "U.vsFirstName + ' ' + U.vsLastName, U.vsScreenName), CR.fRank, T.lSubmitterID " _
			& "FROM ((tblTutorials T LEFT JOIN tblUsers U ON T.lAuthorID = U.lUserID) " _
			& "INNER JOIN (SELECT lItemID FROM tblItemSites WHERE lItemTypeID = " _
			& g_ITEM_TUTORIAL & " AND lSiteID = " & GetSessionValue(g_USER_SITE) _
			& ") tIS ON tIS.lItemID = T.lTutorialID) " _
			& "LEFT JOIN (SELECT lItemID, fRank FROM tblComputedRankings WHERE lItemTypeID = " _
			& g_ITEM_TUTORIAL & ") CR ON CR.lItemID = T.lTutorialID "
		Set m_oData = New kbDataAccess
	End Sub
	
	Private Sub Class_Terminate()
		Set m_oData = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetItem()
	'	Purpose: 	get tutorial data
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Function GetItem(ByVal v_lTutorialID, ByVal v_bOwnerOnly)
		dim sTutorialSQL
		dim sCatSQL
		dim aData
		
		sTutorialSQL = m_sBaseSQL & "WHERE T.lTutorialID = " & v_lTutorialID
		If v_bOwnerOnly And CStr(GetSessionValue(g_USER_TYPE)) <> CStr(g_USER_ADMIN) Then
			sTutorialSQL = sTutorialSQL & " AND T.lAuthorID = " & GetSessionValue(g_USER_ID)
		End If
		sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
				& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
				& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_TUTORIAL _
				& " AND IC.lItemID = " & v_lTutorialID
		aData = m_oData.GetArray(sTutorialSQL)
		aData = JoinArray(aData, m_TUTORIAL_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
		GetItem = aData
	End Function

	'-------------------------------------------------------------------------
	'	Name: 		GetPublic()
	'	Purpose: 	get tutorial data
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Function GetPublic(ByVal v_aFilter)
		dim sTutorialSQL
		dim sCatSQL
		dim aData
		dim sKey
		
		sKey = MakeKey(g_ITEM_TUTORIAL, v_aFilter)
		aData = Application(sKey)
		If Not IsArray(aData) Then
			sTutorialSQL = m_sBaseSQL & "WHERE T.lStatusID = " & g_STATUS_APPROVED & " "
			If v_aFilter(g_FILTER_AUTHOR) <> 0 Then
				sTutorialSQL = sTutorialSQL & "AND T.lAuthorID = " & v_aFilter(g_FILTER_AUTHOR) & " "
			End If
			If v_aFilter(g_FILTER_SOFTWARE) <> 0 Then
				sTutorialSQL = sTutorialSQL & "AND SV.lSoftwareID = " & v_aFilter(g_FILTER_SOFTWARE) & " "
			End If
			sTutorialSQL = sTutorialSQL & MakeSortSQL(v_aFilter(g_FILTER_SORT))
			sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
					& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
					& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_TUTORIAL
			aData = m_oData.GetArray(sTutorialSQL)
			aData = JoinArray(aData, m_TUTORIAL_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
			If v_aFilter(g_FILTER_CATEGORY) <> 0 Then
				aData = FilterArray(aData, m_TUTORIAL_CATS, g_CAT_ID, v_aFilter(g_FILTER_CATEGORY))
			End If
			Application.Lock
			Application(sKey) = aData
			Application.Unlock
		End If
		GetPublic = aData
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		MakeSortSQL()
	'	Purpose: 	generate SQL for sorting tutorials
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function MakeSortSQL(ByVal v_lSortID)
		dim sSQL
		sSQL = "ORDER BY "
		select case v_lSortID
			case g_SORT_NAME_ASC
				sSQL = sSQL & "T.vsTutorialName"
			case g_SORT_NAME_DESC
				sSQL = sSQL & "T.vsTutorialName DESC"
			case g_SORT_DATE_ASC
				sSQL = sSQL & "T.dtApproveDate, T.vsTutorialName"
			case g_SORT_DATE_DESC
				sSQL = sSQL & "T.dtApproveDate DESC, T.vsTutorialName"
			case g_SORT_OWNER_ASC
				sSQL = sSQL & "U.vsFirstName, U.vsLastName, T.vsTutorialName"
			case g_SORT_OWNER_DESC
				sSQL = sSQL & "U.vsFirstName DESC, U.vsLastName DESC, T.vsTutorialName"
			case else
				sSQL = sSQL & "T.vsTutorialName"
		end select
		MakeSortSQL = sSQL
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetArray()
	'	Purpose: 	get tutorial data, one way or another
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/7/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Function GetArray()
		dim aData
		dim x
		dim aReturn(9)
		
		aReturn(m_TUTORIAL_ID) = Trim(Request.QueryString("id"))
		If Request.Form("fldName") <> "" Then
			With Request
				aReturn(m_TUTORIAL_NAME) = Trim(.Form("fldName"))
				aReturn(m_TUTORIAL_URL) = .Form("fldURL")
				aReturn(m_TUTORIAL_TEXT) = .Form("fldDescription")
				aReturn(m_TUTORIAL_AUTHOR_ID) = .Form("fldAuthor")
				aReturn(m_TUTORIAL_CATS) = .Form("fldCategories")
			End With
		ElseIf IsNumber(aReturn(m_TUTORIAL_ID)) Then
			aData = GetItem(aReturn(m_TUTORIAL_ID), false)
			If IsArray(aData) Then
				for x = 0 to UBound(aData)
					aReturn(x) = aData(x, 0)
				next
			End If
		End If
		aReturn(m_TUTORIAL_AUTHOR_ID) = ReplaceNull(aReturn(m_TUTORIAL_AUTHOR_ID), GetSessionValue(g_USER_ID))
		GetArray = aReturn
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		Save()
	'	Purpose: 	save tutorial to database
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/03		JEA		Creation
	'	4/25/03		JEA		Save site
	'-------------------------------------------------------------------------
	Public Sub Save(ByVal v_aData)
		dim oRS
		dim sQuery
		dim lTutorialID
		
		lTutorialID = v_aData(m_TUTORIAL_ID)
		Call m_oData.BeginTrans()
		
		If IsNumber(lTutorialID) Then
			' update existing
			sQuery = "UPDATE tblTutorials SET " _
				& "vsTutorialURL = '" & Replace(v_aData(m_TUTORIAL_URL), "http://", "") & "', " _
				& "lAuthorID = " & v_aData(m_TUTORIAL_AUTHOR_ID) & ", " _
				& "vsTutorialName = '" & CleanForSQL(v_aData(m_TUTORIAL_NAME)) & "', " _
				& "vsDescription = '" & CleanForSQL(v_aData(m_TUTORIAL_TEXT)) & "' " _
				& "WHERE lTutorialID = " & lTutorialID
			Call m_oData.ExecuteOnly(sQuery)
			Call ClearCache(g_ITEM_TUTORIAL)
		Else
			' insert new
			Set oRS = Server.CreateObject("ADODB.Recordset")
			With oRS
				.Open "tblTutorials", m_oData.Connection, adOpenStatic, adLockOptimistic, adCmdTable
				.AddNew
				.Fields("vsTutorialURL") = Replace(v_aData(m_TUTORIAL_URL), "http://", "")
				.Fields("lAuthorID") = v_aData(m_TUTORIAL_AUTHOR_ID)
				.Fields("lSubmitterID") = GetSessionValue(g_USER_ID)
				.Fields("vsTutorialName") = v_aData(m_TUTORIAL_NAME)
				.Fields("vsDescription") = v_aData(m_TUTORIAL_TEXT)
				.Fields("dtSubmitDate") = Now()
				.Fields("lStatusID") = g_STATUS_PENDING
				.Fields("bLocalFile") = false
				.Update
				lTutorialID = .Fields("lTutorialID")
				.Close
			End With
			Set oRS = Nothing
			
			sQuery = "INSERT INTO tblItemSites (lSiteID, lItemTypeID, lItemID) VALUES (" _
				& GetSessionValue(g_USER_SITE) & ", " _
				& g_ITEM_TUTORIAL & ", " _
				& lTutorialID & ")"
			Call m_oData.ExecuteOnly(sQuery)
		End If
		
		Call SaveCategories(lTutorialID, Trim(v_aData(m_TUTORIAL_CATS)), g_ITEM_TUTORIAL)
		Call m_oData.CommitTrans()
		Call SetSessionValue(g_USER_MSG, "The tutorial has been saved")
		response.redirect "kb_tutorials.asp"
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetPending()
	'	Purpose: 	get array of uploaded files
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
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
	'	Purpose: 	update tutorial status
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub ApprovePending(ByVal v_lTutorialID)
		dim sQuery
		sQuery = "UPDATE tblTutorials SET lStatusID = " & g_STATUS_APPROVED _
			& ", dtApproveDate = " & g_sSQL_DATE_DELIMIT & Now() & g_sSQL_DATE_DELIMIT _
			& " WHERE lTutorialID = " & v_lTutorialID
		Call m_oData.ExecuteOnly(sQuery)
		Call ClearCache(g_ITEM_TUTORIAL)
		'Call m_oData.LogActivity(g_ACT_APPROVE_UPLOAD, "", "", "", "", "")
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		DenyPending()
	'	Purpose: 	update tutorial status
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub DenyPending(ByVal v_lTutorialID)
		dim sQuery
		sQuery = "UPDATE tblTutorials SET lStatusID = " & g_STATUS_REJECTED _
			& " WHERE lTutorialID = " & v_lTutorialID
		Call m_oData.ExecuteOnly(sQuery)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		SaveCategories()
	'	Purpose: 	save tutorial category data to database
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/03		JEA		Creation
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