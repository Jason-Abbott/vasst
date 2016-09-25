<%
Const m_FORUM_ID = 0
Const m_FORUM_NAME = 1
Const m_FORUM_URL = 2
Const m_FORUM_TEXT = 3
Const m_FORUM_HOST_ID = 4
Const m_FORUM_HOST_NAME = 5
Const m_FORUM_HOST_URL = 6
Const m_FORUM_RANK = 7
Const m_FORUM_CATS = 8

'-------------------------------------------------------------------------
'	Name: 		kbForumData class
'	Purpose: 	methods for saving and retrieving forum data
'Modifications:
'	Date:		Name:	Description:
'	1/10/02		JEA		Creation
'	4/25/03		JEA		Add site condition
'-------------------------------------------------------------------------
Class kbForumData
	Private m_sBaseSQL
	Private m_oData

	Private Sub Class_Initialize()
		m_sBaseSQL = "SELECT F.lForumID, F.vsForumName, F.vsForumURL, F.vsForumDescription, " _
			& "F.lForumHostID, FH.vsHostName, FH.vsHostURL, CR.fRank FROM ((tblForums F INNER JOIN " _
			& "tblForumHosts FH ON FH.lForumHostID = F.lForumHostID) " _
			& "INNER JOIN (SELECT lItemID FROM tblItemSites WHERE lItemTypeID = " _
			& g_ITEM_FORUM & " AND lSiteID = " & GetSessionValue(g_USER_SITE) _
			& ") tIS ON tIS.lItemID = F.lForumID) " _
			& "LEFT JOIN (SELECT lItemID, fRank FROM tblComputedRankings WHERE lItemTypeID = " _
			& g_ITEM_FORUM & ") CR ON CR.lItemID = F.lForumID "
		Set m_oData = New kbDataAccess
	End Sub
	
	Private Sub Class_Terminate()
		Set m_oData = Nothing
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		GetPublic()
	'	Purpose: 	get forum data
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Function GetPublic(ByVal v_aFilter)
		dim sForumSQL
		dim sCatSQL
		dim aData
		dim sKey
		
		sKey = MakeKey(g_ITEM_FORUM, v_aFilter)
		aData = Application(sKey)
		If Not IsArray(aData) Then
			sForumSQL = m_sBaseSQL & MakeSortSQL(v_aFilter(g_FILTER_SORT))
			sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
					& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
					& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_FORUM
			aData = m_oData.GetArray(sForumSQL)
			aData = JoinArray(aData, m_FORUM_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
			If v_aFilter(g_FILTER_CATEGORY) <> 0 Then
				aData = FilterArray(aData, m_FORUM_CATS, g_CAT_ID, v_aFilter(g_FILTER_CATEGORY))
			End If
			Application.Lock
			Application(sKey) = aData
			Application.Unlock
		End If
		GetPublic = aData
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetItem()
	'	Purpose: 	get single forum data
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Function GetItem(ByVal v_lForumID, ByVal v_bOwnerOnly)
		dim sForumSQL
		dim sCatSQL
		dim aData
		
		sForumSQL = m_sBaseSQL & "WHERE F.lForumID = " & v_lForumID
		sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
				& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
				& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_FORUM _
				& " AND IC.lItemID = " & v_lForumID
		aData = m_oData.GetArray(sForumSQL)
		aData = JoinArray(aData, m_FORUM_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
		GetItem = aData
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		MakeSortSQL()
	'	Purpose: 	generate SQL for sorting forums
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function MakeSortSQL(ByVal v_lSortID)
		dim sSQL
		sSQL = "ORDER BY "
		select case v_lSortID
			case g_SORT_NAME_ASC
				sSQL = sSQL & "F.vsForumName"
			case g_SORT_NAME_DESC
				sSQL = sSQL & "F.vsForumName DESC"
			case g_SORT_OWNER_ASC
				sSQL = sSQL & "FH.vsHostName, F.vsForumName"
			case g_SORT_OWNER_DESC
				sSQL = sSQL & "FH.vsHostName DESC, F.vsForumName"
			case else
				sSQL = sSQL & "F.vsForumName"
		end select
		MakeSortSQL = sSQL
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetArray()
	'	Purpose: 	get forum data, one way or another
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Function GetArray()
		dim aData
		dim x
		dim aReturn(8)
		
		aReturn(m_FORUM_ID) = Trim(Request.QueryString("id"))
		If Request.Form("fldName") <> "" Then
			With Request
				aReturn(m_FORUM_NAME) = Trim(.Form("fldName"))
				aReturn(m_FORUM_URL) = .Form("fldURL")
				aReturn(m_FORUM_TEXT) = .Form("fldDescription")
				aReturn(m_FORUM_HOST_ID) = .Form("fldHost")
				aReturn(m_FORUM_CATS) = .Form("fldCategories")
			End With
		ElseIf IsNumber(aReturn(m_FORUM_ID)) Then
			aData = GetItem(aReturn(m_FORUM_ID), false)
			If IsArray(aData) Then
				for x = 0 to UBound(aData)
					aReturn(x) = aData(x, 0)
				next
			End If
		End If
		GetArray = aReturn
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		Save()
	'	Purpose: 	save forum to database
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'	4/25/03		JEa		Save site information
	'-------------------------------------------------------------------------
	Public Sub Save(ByVal v_aData)
		dim oRS
		dim sQuery
		dim lForumID
		
		lForumID = v_aData(m_FORUM_ID)
		Call m_oData.BeginTrans()
		
		If IsNumber(lForumID) Then
			' update existing
			sQuery = "UPDATE tblForums SET " _
				& "vsForumURL = '" & Replace(v_aData(m_FORUM_URL), "http://", "") & "', " _
				& "vsForumName = '" & CleanForSQL(v_aData(m_FORUM_NAME)) & "', " _
				& "vsForumDescription = '" & CleanForSQL(v_aData(m_FORUM_TEXT)) & "', " _
				& "lForumHostID = " & v_aData(m_FORUM_HOST_ID) _
				& " WHERE lForumID = " & lForumID
			Call m_oData.ExecuteOnly(sQuery)
		Else
			' insert new
			Set oRS = Server.CreateObject("ADODB.Recordset")
			With oRS
				.Open "tblForums", m_oData.Connection, adOpenStatic, adLockOptimistic, adCmdTable
				.AddNew
				.Fields("vsForumURL") = Replace(v_aData(m_FORUM_URL), "http://", "")
				.Fields("vsForumName") = v_aData(m_FORUM_NAME)
				.Fields("vsForumDescription") = v_aData(m_FORUM_TEXT)
				.Fields("lForumHostID") = v_aData(m_FORUM_HOST_ID)
				.Fields("lStatusID") = g_STATUS_APPROVED
				.Fields("bModerated") = true
				.Update
				lForumID = .Fields("lForumID")
				.Close
			End With
			Set oRS = Nothing
			
			sQuery = "INSERT INTO tblItemSites (lSiteID, lItemTypeID, lItemID) VALUES (" _
				& GetSessionValue(g_USER_SITE) & ", " _
				& g_ITEM_FORUM & ", " _
				& lForumID & ")"
			Call m_oData.ExecuteOnly(sQuery)
		End If
		Call SaveCategories(lForumID, Trim(v_aData(m_FORUM_CATS)), g_ITEM_FORUM)
		Call m_oData.CommitTrans()
		Call SetSessionValue(g_USER_MSG, "The forum has been saved")
		Call ClearCache(g_ITEM_FORUM)
		response.redirect "kb_forums.asp"
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		SaveCategories()
	'	Purpose: 	save forum category data to database
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
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