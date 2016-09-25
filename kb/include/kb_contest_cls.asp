<%
Const m_CONTEST_ID = 0
Const m_CONTEST_NAME = 1
Const m_CONTEST_START = 2
Const m_CONTEST_END = 3
Const m_CONTEST_VOTE_BY = 4
Const m_CONTEST_TEXT = 5
Const m_CONTEST_VOTES = 6
Const m_CONTEST_WEIGHT = 7
Const m_CONTEST_MAX_ENTRIES = 8
Const m_CONTEST_NO_EXTERNAL_MEDIA = 9
Const m_CONTEST_FREE_PLUGINS_ONLY = 10
Const m_CONTEST_WINNERS = 11
Const m_CONTEST_SITE = 12

Class kbContest
	Private m_sBaseSQL
	
	Private Sub Class_Initialize()
		m_sBaseSQL = "lContestID, vsContestName, dtStartDate, dtEndDate, dtVoteByDate, " _
			& "vsDescription, lVotesAllowed, lWeightFactor, lMaxEntries, bNoExternalMedia, " _
			& "bFreePluginsOnly, lWinners, lSiteID FROM tblContests WHERE lSiteID = " _
			& GetSessionValue(g_USER_SITE)
	End Sub
	
	Private Sub Class_Terminate()
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		CastVote()
	'	Purpose: 	record user's vote
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub CastVote(ByVal v_lContestID, ByVal v_lWeight, ByVal v_lMaxVotes, ByVal v_sItems, _
		ByVal v_lItemTypeID)
		
		Const ITEM_ID = 0
		Const OLD_RANK = 1
		Const NEW_RANK = 2
		dim sQueryAdd
		dim sQueryPoints
		dim oData
		dim aVote
		dim x
		
		sQueryAdd = "INSERT INTO tblContestVotes (lContestID, lVoterID, dtDateVoted, " _
			& "lItemTypeID, lItemID, lRank) VALUES (" _
			& v_lContestID & ", " _
			& GetSessionValue(g_USER_ID) & ", '" _
			& Now & "', " _
			& v_lItemTypeID & ", "

		sQueryPoints = " FROM tblContestVotes WHERE lContestID = " & v_lContestID _
			& " AND lVoterID = " & GetSessionValue(g_USER_ID) & " AND lItemTypeID = " _
			& v_lItemTypeID
			
		If v_sItems <> "," Then
			v_sItems = Mid(v_sItems, 2, Len(v_sItems) - 2)
			Set oData = New kbDataAccess
			aVote = MakeUserVoteArray(Split(v_sItems, ","), oData.GetArray("SELECT lItemID, lRank" & sQueryPoints))
			Call oData.BeginTrans()
			Call oData.ExecuteOnly("DELETE" & sQueryPoints)
			For x = 0 to UBound(aVote, 2)
				if aVote(NEW_RANK, x) > 0 then
					Call oData.ExecuteOnly(sQueryAdd & aVote(ITEM_ID, x) & ", " & aVote(NEW_RANK, x) & ")")
					Call oData.LogActivity(g_ACT_VOTE, aVote(ITEM_ID, x), v_lContestID, "", "", "")
				end if
			Next
			Call UpdateVoteTotals(v_lContestID, v_lWeight, v_lMaxVotes, v_lItemTypeID, aVote, oData)
			Call oData.CommitTrans()
			Set oData = Nothing
			Call SetSessionValue(g_USER_MSG, g_sMSG_AFTER_VOTING)
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		MakeUserVoteArray()
	'	Purpose: 	make array of current votes and any previous votes for same items
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function MakeUserVoteArray(ByVal v_aVote, ByVal v_aOldVote)
		Const ITEM_ID = 0
		Const OLD_RANK = 1
		Const NEW_RANK = 2
		dim aData()
		dim lNewBound
		dim bExists
		dim x, y, z

		ReDim aData(2, UBound(v_aVote))
		
		If IsArray(v_aOldVote) Then
			for x = 0 to UBound(aData, 2)
				aData(ITEM_ID, x) = v_aVote(x)
				aData(NEW_RANK, x) = x + 1
				aData(OLD_RANK, x) = 0			' default value
				
				for y = 0 to UBound(v_aOldVote, 2)
					' find any old votes matching a new vote
					if v_aOldVote(ITEM_ID, y) <> 0 then
						' only consider unprocessed files
						bExists = false
						
						if CStr(aData(ITEM_ID, x)) = CStr(v_aOldVote(ITEM_ID, y)) then
							aData(OLD_RANK, x) = v_aOldVote(OLD_RANK, y)
							v_aOldVote(ITEM_ID, y) = 0
							Exit For
						else
							' see if we need to add the old vote
							for z = 0 to UBound(aData, 2)
								if CStr(aData(ITEM_ID, z)) = CStr(v_aOldVote(ITEM_ID, y)) then bExists = true
							next
							
							if Not bExists then
								' the old vote doesn't match ANY new vote
								lNewBound = UBound(aData, 2) + 1
								ReDim Preserve aData(2, lNewBound)
								aData(ITEM_ID, lNewBound) = v_aOldVote(ITEM_ID, y)
								aData(OLD_RANK, lNewBound) = v_aOldVote(OLD_RANK, y)
								aData(NEW_RANK, lNewBound) = 0
								v_aOldVote(ITEM_ID, y) = 0
							end if
						end if
					end if
				next
			next
		Else
			' no old votes so old rank defaults to zero
			for x = 0 to UBound(aData, 2)
				aData(ITEM_ID, x) = v_aVote(x)
				aData(OLD_RANK, x) = 0
				aData(NEW_RANK, x) = x + 1
			next
		End If
		MakeUserVoteArray = aData
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		UpdateVoteTotals()
	'	Purpose: 	update the computed vote totals
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub UpdateVoteTotals(ByVal v_lContestID, ByVal v_lWeight, ByVal v_lMaxVotes, _
		ByVal v_lItemTypeID, ByVal v_aVote, ByRef r_oData)
		
		Const ITEM_ID = 0
		Const POINTS = 1
		Const OLD_RANK = 1
		Const NEW_RANK = 2
		dim sQuery
		dim lOldPoints
		dim lPointChange	' value could be negative if old vote had higher rank
		dim lNewTotalPoints
		dim aPoints
		dim x, y
		
		sQuery = "SELECT lItemID, lPoints FROM tblComputedVotePoints " _
			& "WHERE lContestID = " & v_lContestID & " AND lItemTypeID = " _
			& v_lItemTypeID & " AND lItemID IN ("
		for x = 0 to UBound(v_aVote, 2)
			sQuery = sQuery & v_aVote(ITEM_ID, x) & ","
		next
		sQuery = Left(sQuery, Len(sQuery) - 1) & ")"
		aPoints = r_oData.GetArray(sQuery)

		for x = 0 to UBound(v_aVote, 2)
			' go through each item
			lOldPoints = 0
			lPointChange = GetVotePoints(v_aVote(NEW_RANK, x), v_lWeight, v_lMaxVotes) - GetVotePoints(v_aVote(OLD_RANK, x), v_lWeight, v_lMaxVotes)
			
			If IsArray(aPoints) then
				' contest already has items with points--check for current item points
				for y = 0 to UBound(aPoints, 2)
					if CStr(aPoints(ITEM_ID, y)) = CStr(v_aVote(ITEM_ID, x)) then lOldPoints = aPoints(POINTS, y)
				next
			end if
			

			If lPointChange <> 0 then
				If lOldPoints > 0 then
					' this item already has votes
					lNewTotalPoints = lPointChange + lOldPoints
					If lNewTotalPoints > 0 Then
						sQuery = "UPDATE tblComputedVotePoints SET lPoints = " & lNewTotalPoints _
							& " WHERE lContestID = " & v_lContestID & " AND lItemID = " _
							& v_aVote(ITEM_ID, x) & " AND lItemTypeID = " & v_lItemTypeID
					Else
						' all votes have been removed for this item
						sQuery = "DELETE FROM tblComputedVotePoints WHERE lContestID = " _
							& v_lContestID & " AND lProjectID = " & v_aVote(ITEM_ID, x) _
							& " AND lItemTypeID = " & v_lItemTypeID
					End If
				else
					' this is the first vote for this file
					sQuery = "INSERT INTO tblComputedVotePoints (lItemID, lItemTypeID, lContestID, lPoints)" _
						& " VALUES (" _
						& v_aVote(ITEM_ID, x) & ", " _
						& v_lItemTypeID & ", " _
						& v_lContestID & ", " _
						& lPointChange & ")"
				end if
				Call r_oData.ExecuteOnly(sQuery)		
			end if
		next
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetVotePoints()
	'	Purpose: 	get weighted points for vote
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function GetVotePoints(ByVal v_lRank, ByVal v_lWeight, ByVal v_lMaxVotes)
		GetVotePoints = IIf((v_lRank = 0), 0, (v_lWeight * (v_lMaxVotes - v_lRank)) + 1)
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteQualifyingItemsList()
	'	Purpose: 	write list of user's items qualified for contest
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteQualifyingItemList(ByVal v_aData)
		Const ITEM_ID = 0
		Const FRIENDLY_NAME = 1
		Const CONTEST_ID = 2
		dim bNoExternalMedia
		dim bFreePluginsOnly
		dim lInContestCount
		dim oLayout
		dim aData
		dim sList
		dim x
		
		bNoExternalMedia = v_aData(m_CONTEST_NO_EXTERNAL_MEDIA)
		bFreePluginsOnly = v_aData(m_CONTEST_FREE_PLUGINS_ONLY)
		lInContestCount = 0
		
		aData = GetUserContestItems(v_aData(m_CONTEST_ID))
		with response
			.write "<div class='ContestNote'>This contest allows up to "
			.write v_aData(m_CONTEST_MAX_ENTRIES)
			.write " file entries"
			if bNoExternalMedia Or bFreePluginsOnly then
				.write " that "
				if bNoExternalMedia then .write "<b>do not</b> require external media"
				if bNoExternalMedia And bFreePluginsOnly then .write " and "
				if bFreePluginsOnly then .write "<b>use only</b> built-in or free plugins"
			end if
			.write ".</div>"
		
			if IsArray(aData) then
				sList = ""
				.write "<table cellspacing='0' cellpadding='2' border='0'>"
				Set oLayout = New kbLayout
				for x = 0 to UBound(aData, 2)
					if IsVoid(aData(CONTEST_ID, x)) then
						sList = sList & "<option value='" & aData(ITEM_ID, x) _
							& "'>" & aData(FRIENDLY_NAME, x)
					else
						if lInContestCount = 0 then
							.write "<tr><td class='EntryLabel'><nobr>In contest:</nobr></td>"
							.write "<td class='EntryData'><select name='fldInContest'>"
						end if
						.write "<option value='"
						.write aData(ITEM_ID, x)
						.write "'>"
						.write aData(FRIENDLY_NAME, x)
						lInContestCount = lInContestCount + 1
					end if
				next
				if lInContestCount > 0 then
					 .write "</select></td><tr><td></td><td class='EntryButton'><a href='javascript:RemoveFromContest();'>"
					 Call oLayout.WriteToggleImage("btn_remove-from-contest", "", "Remove this file from the contest", "width='143' height='14'", false)
					 .write "</a></td>"
				end if
				if sList <> "" then
					.write "<tr><td class='EntryLabel'><nobr>Eligible for contest:</nobr></td>"
					.write "<td class='EntryData'><select name='fldFiles'>"
					.write sList
					.write "</select></td><tr><td></td><td class='EntryButton'><a href='javascript:AddToContest();'>"
					Call oLayout.WriteToggleImage("btn_add-to-contest", "", "Add this file to the contest", "height='14'", false)
					.write "</a></td>"
				end if
				Set oLayout = Nothing
				.write "<input type='hidden' name='fldEntries' value='"
				.write lInContestCount
				.write "'><input type='hidden' name='fldMaxEntries' value='"
				.write v_aData(m_CONTEST_MAX_ENTRIES)
				.write "'></table>"
			else
				.write "<br><center>Add your files to our listing and return here to to enter them in this contest</center>"
			end if
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetUserContestItems()
	'	Purpose: 	get user items (currently specific to files) qualified for contest
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/03		JEA		Creation
	'	10/7/03		JEA		Return only items for current site
	'-------------------------------------------------------------------------
	Private Function GetUserContestItems(ByVal v_lContestID)
		dim sQuery
		dim oData
		sQuery = "SELECT F.lProjectID, F.vsFriendlyName, CI.lContestID FROM (tblProjects F " _
			& "INNER JOIN (SELECT lItemID FROM tblItemSites WHERE lItemTypeID = " _
			& g_ITEM_PROJECT & " AND lSiteID = " & GetSessionValue(g_USER_SITE) _
			& ") tIS ON tIS.lItemID = F.lProjectID) " _
			& "LEFT JOIN tblContestItems CI ON CI.lItemID = F.lProjectID " _
			& "WHERE F.lUserID = " & GetSessionValue(g_USER_ID) _
			& " AND (CI.lContestID = " & v_lContestID & " OR CI.lContestID IS NULL) " _
			& "AND F.lStatusID = " & g_STATUS_APPROVED _
			& " ORDER BY F.vsFriendlyName"
		Set oData = New kbDataAccess
		GetUserContestItems = oData.GetArray(sQuery)
		Set oData = Nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		AddItem()
	'	Purpose: 	make item available for contest voting
	'Modifications:
	'	Date:		Name:	Description:
	'	1/3/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub AddItem(ByVal v_lContestID, ByVal v_lItemID, ByVal v_lItemTypeID)
		dim sQuery
		dim oData
		sQuery = "INSERT INTO tblContestItems (lContestID, lItemID, lItemTypeID, dtAddedDate) " _
			& "VALUES (" _
			& v_lContestID & ", " _
			& v_lItemID & ", " _
			& v_lItemTypeID & ", " & g_sSQL_DATE_DELIMIT _
			& Now() & g_sSQL_DATE_DELIMIT & ")"
		Set oData = New kbDataAccess
		Call oData.ExecuteOnly(sQuery)
		Set oData = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		RemoveItemGlobally()
	'	Purpose: 	remove item from all contests
	'Modifications:
	'	Date:		Name:	Description:
	'	1/5/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub RemoveItemGlobally(ByVal v_lItemID, ByVal v_lItemTypeID)
		Const CONTEST_ID = 0
		dim sQuery
		dim oData
		dim aData
		dim x
		
		sQuery = "SELECT lContestID FROM tblContestItems WHERE lItemID = " _
			& v_lItemID & " AND lItemTypeID = " & v_lItemTypeID
		Set oData = New kbDataAccess
		aData = oData.GetArray(sQuery)
		Set oData = Nothing
		If IsArray(aData) Then
			' item is in contests
			for x = 0 to UBound(aData, 2)
				Call RemoveItem(aData(CONTEST_ID,x), v_lItemID, v_lItemTypeID)
			next
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		RemoveItem()
	'	Purpose: 	remove item from contest
	'Modifications:
	'	Date:		Name:	Description:
	'	1/3/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub RemoveItem(ByVal v_lContestID, ByVal v_lItemID, ByVal v_lItemTypeID)
		dim sQuery
		dim oData
		dim aData
		
		sQuery = " WHERE lItemID = " & v_lItemID & " AND lItemTypeID = " _
			& v_lItemTypeID & " AND lContestID = " & v_lContestID
		Set oData = New kbDataAccess
		aData = oData.GetArray("SELECT lItemID FROM tblContestVotes" & sQuery)
		Call oData.BeginTrans
		If IsArray(aData) Then
			' votes have already been cast for this file--ugh
			Call RemoveItemVotes(v_lContestID, v_lItemID, v_lItemTypeID, oData)
			Call RecomputePoints(v_lContestID, oData)
		End If
		Call oData.ExecuteOnly("DELETE FROM tblContestItems" & sQuery)
		Call oData.CommitTrans()
		Set oData = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		RecomputePoints()
	'	Purpose: 	recompute all item rankings for given contest
	'				should be contained in ADO transaction
	'Modifications:
	'	Date:		Name:	Description:
	'	1/3/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub RecomputePoints(ByVal v_lContestID, ByRef r_oData)
		Const ITEM_ID = 0
		Const ITEM_TYPE_ID = 1
		Const RANK = 2
		Const MAX_VOTES = 0
		Const WEIGHT = 1
		dim sQuery
		dim lMaxVotes
		dim lWeight
		dim lItemID
		dim lItemTypeID
		dim lPoints
		dim aData
		dim x
		
		sQuery = "SELECT lVotesAllowed, lWeightFactor FROM tblContests WHERE lContestID = " & v_lContestID
		aData = r_oData.GetArray(sQuery)
		If IsArray(aData) Then
			lMaxVotes = aData(MAX_VOTES, 0)
			lWeight = aData(WEIGHT, 0)
		Else
			Exit Sub
		End If
		
		sQuery = "SELECT lItemID, lItemTypeID, lRank FROM tblContestVotes WHERE lContestID = " _
			& v_lContestID & " ORDER BY lItemID"
		aData = r_oData.GetArray(sQuery)
		
		If IsArray(aData) Then
			sQuery = "DELETE FROM tblComputedVotePoints WHERE lContestID = " & v_lContestID
			Call r_oData.ExecuteOnly(sQuery)				' delete all old points
			lItemID = 0
			lItemTypeID = 0
			for x = 0 to UBound(aData, 2)
				If lItemID <> aData(ITEM_ID, x) Or lItemTypeID <> aData(ITEM_TYPE_ID, x) Then
					if x > 0 then Call InsertComputedVote(v_lContestID, lItemID, lItemTypeID, lPoints, r_oData) 
					lItemID = aData(ITEM_ID, x)
					lItemTypeID = aData(ITEM_TYPE_ID, x)
					lPoints = 0
				End If
				lPoints = lPoints + GetVotePoints(aData(RANK, x), lWeight, lMaxVotes)
			next
			If lPoints > 0 Then Call InsertComputedVote(v_lContestID, lItemID, lItemTypeID, lPoints, r_oData) 
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		InsertComputedVote()
	'	Purpose: 	add a computed vote
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub InsertComputedVote(ByVal v_lContestID, ByVal v_lItemID, ByVal v_lItemTypeID, _
		ByVal v_lPoints, ByRef r_oData)

		dim sQuery
		sQuery = "INSERT INTO tblComputedVotePoints (lItemID, lItemTypeID, " _
			& "lContestID, lPoints) VALUES (" & v_lItemID & ", " & v_lItemTypeID _
			& ", " & v_lContestID & ", " & v_lPoints & ")"
		Call r_oData.ExecuteOnly(sQuery)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		RemoveItemVotes()
	'	Purpose: 	remove votes for contest item and re-rank remaining items
	'Modifications:
	'	Date:		Name:	Description:
	'	1/3/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub RemoveItemVotes(ByVal v_lContestID, ByVal v_lItemID, ByVal v_lItemTypeID, ByRef r_oData)
		Const ITEM_ID = 0
		Const USER_ID = 1
		Const RANK = 2
		dim sQuery
		dim sQueryDel
		dim lUserID
		dim bAdjust
		dim aData
		dim x
		
		sQueryDel = "DELETE FROM tblContestVotes WHERE lContestID = " & v_lContestID _
			& " AND lItemTypeID = " & v_lItemTypeID & " AND lItemID = " _
			& v_lItemID & " AND lVoterID = "
		sQuery = "SELECT lItemID, lVoterID, lRank FROM tblContestVotes WHERE lContestID = " _
			& v_lContestID & " AND lItemTypeID = " & v_lItemTypeID & " ORDER BY lVoterID, lRank"
		aData = r_oData.GetArray(sQuery)
		If IsArray(aData) Then
			lUserID = 0
			bAdjust = false
			for x = 0 to UBound(aData, 2)
				if CStr(lUserID) <> CStr(aData(USER_ID, x)) then
					bAdjust = false
					lUserID = aData(USER_ID, x)
				end if
				If CStr(v_lItemID) = CStr(aData(ITEM_ID, x)) Then
					' remove user's vote for this item
					bAdjust = true
					Call r_oData.ExecuteOnly(sQueryDel & lUserID)
				ElseIf bAdjust Then
					' increase ranking by one to replace file
					sQuery = "UPDATE tblContestVotes SET lRank = " & (aData(RANK, x) - 1) _
						& " WHERE lContestID = " & v_lContestID & " AND lItemID = " & aData(ITEM_ID, x) _
						& " AND lVoterID = " & lUserID & " AND lItemTypeID = " & v_lItemTypeID
					Call r_oData.ExecuteOnly(sQuery)
				End If
			next
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteContestVote()
	'	Purpose: 	write voting form
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/03		JEA		Creation
	'	3/23/03		JEA		Check for contestants
	'-------------------------------------------------------------------------
	Public Sub WriteContestVote(ByVal v_aData)
		Const ITEM_ID = 0
		Const RANK = 1
		dim x
		dim oData
		dim oLayout
		dim sQuery
		dim aVotes
		dim aFile
		dim sFileList
		
		sQuery = "SELECT CI.lItemID, F.vsFriendlyName FROM tblContestItems CI " _
			& "INNER JOIN tblProjects F ON F.lProjectID = CI.lItemID WHERE CI.lContestID = " _
			& v_aData(m_CONTEST_ID) & " AND F.lStatusID = " & g_STATUS_APPROVED _
			& " ORDER BY F.vsFriendlyName"
		
		sFileList = MakeList(sQuery, "")
		
		If IsVoid(sFileList) Then
			response.write "There are not yet<br>any files to vote on"
			exit sub
		End If
		
		' get array of previous votes to pre-select drop-downs
		sQuery = "SELECT lItemID, lRank FROM tblContestVotes WHERE lContestID = " _
			& v_aData(m_CONTEST_ID) & " AND lVoterID = " & GetSessionValue(g_USER_ID) _
			& " ORDER BY lRank"
		Set oData = New kbDataAccess
		aVotes = oData.GetArray(sQuery)
		Set oData = Nothing
		
		ReDim aFile(v_aData(m_CONTEST_VOTES))
		If IsArray(aVotes) Then
			for x = 1 to v_aData(m_CONTEST_VOTES)
				if (x - 1) <= UBound(aVotes, 2) then aFile(x) = aVotes(ITEM_ID, x - 1)
			next
		End If		
		
		with response
			Set oLayout = New kbLayout
			Call oLayout.WriteTitleBoxTop(v_aData(m_CONTEST_NAME) & " Vote", "width='100%'", "")
			.write "<center>"
			if v_aData(m_CONTEST_VOTES) > 1 then
				.write "You may vote for up to "
				.write v_aData(m_CONTEST_VOTES)
				.write " different files.<br>The order <b>is "
				.write IIf((v_aData(m_CONTEST_WEIGHT) > 0), "significant", "not</b> significant")
				.write ":"
			end if
			.write "<ol>"
			for x = 1 to v_aData(m_CONTEST_VOTES)
				.write "<li><select name='fldVote"
				.write x
				.write "'><option value='0'>no vote"
				.write MakeSelected(sFileList, aFile(x))
				.write "</select>"
			next
			.write "</ol><input type='hidden' name='fldMaxVotes' value='"
			.write v_aData(m_CONTEST_VOTES)
			.write "'><input type='hidden' name='fldVotes' value=','>"
			Call oLayout.WriteToggleImage("btn_cast-vote", "", "Cast Vote", "class='Image'", true)
			.write "</center>"
			Call m_oLayout.WriteBoxBottom("")
			Set oLayout = Nothing
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteContests()
	'	Purpose: 	write contests to screen
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteContests()
		dim aData
		dim oLayout
		dim x
		aData = GetContests()
		If IsArray(aData) Then
			with Response
				.write "<center><table cellspacing='0' cellpadding='0' border='0'><tr><td class='BoxLeft'>&nbsp;</td>"
				.write "<td class='BoxBody'>"
				for x = 0 to UBound(aData, 2)
					.write "<div class='Contest'>"
					.write FormatAsHTML(aData(m_CONTEST_TEXT, 0))
					.write "<br><b><a href='kb_contest.asp?id="
					.write aData(m_CONTEST_ID, 0)
					.write "'>"
					If Date <= aData(m_CONTEST_VOTE_BY, 0) Then
						.write "Vote or enter your file in the contest</a> by "
						If aData(m_CONTEST_VOTE_BY, 0) = Date Then
							.write "the end of today!"
						Else
							.write FormatDate(aData(m_CONTEST_VOTE_BY, 0))
						End If
					Else
						.write "Voting ended "
						.write FormatDate(aData(m_CONTEST_VOTE_BY, 0))
						.write ". View the Results."
					End If
					.write "</b></div>"
				next
				Set oLayout = New kbLayout
				Call oLayout.WriteBoxBottom("")
				Set oLayout = Nothing
				.write "</center>"
			end with
		End If
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		GetContests()
	'	Purpose: 	get array of current contests
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function GetContests()
		dim sQuery
		dim oData
		sQuery = "SELECT " & m_sBaseSQL & " AND dtEndDate > " & g_sSQL_DATE_DELIMIT _
			& Date & g_sSQL_DATE_DELIMIT & " ORDER BY dtVoteByDate DESC"
		Set oData = New kbDataAccess
		GetContests = oData.GetArray(sQuery)
		Set oData = Nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		SaveContest()
	'	Purpose: 	save contest data
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/03		JEA		Creation
	'	4/25/03		JEA		Include site with contest
	'-------------------------------------------------------------------------
	Public Sub SaveContest(ByRef r_aData)
		dim sQuery
		dim oRS
		dim oData
		
		Set oData = New kbDataAccess
		If IsNumber(r_aData(m_CONTEST_ID)) Then
			' contest already exists
			sQuery = "UPDATE tblContests SET " _
				& "vsContestName = '" & CleanForSQL(r_aData(m_CONTEST_NAME)) & "', " _
				& "dtStartDate = '" & r_aData(m_CONTEST_START) & "', " _
				& "dtEndDate = '" & r_aData(m_CONTEST_END) & "', " _
				& "dtVoteByDate = '" & r_aData(m_CONTEST_VOTE_BY) & "', " _
				& "vsDescription = '" & CleanForSQL(r_aData(m_CONTEST_TEXT)) & "', " _
				& "lVotesAllowed = " & r_aData(m_CONTEST_VOTES) & ", " _
				& "lWeightFactor = " & r_aData(m_CONTEST_WEIGHT) & ", " _
				& "lMaxEntries = " & r_aData(m_CONTEST_MAX_ENTRIES) & ", " _
				& "bFreePluginsOnly = " & r_aData(m_CONTEST_FREE_PLUGINS_ONLY) & ", " _
				& "bNoExternalMedia = " & r_aData(m_CONTEST_NO_EXTERNAL_MEDIA) & ", " _
				& "lWinners = " & r_aData(m_CONTEST_WINNERS) & ", " _
				& "lSiteID = " & r_aData(m_CONTEST_SITE) _
				& " WHERE lContestID = " & r_aData(m_CONTEST_ID)
				
			Call oData.BeginTrans()
			Call oData.ExecuteOnly(sQuery)
			Call RecomputePoints(r_aData(m_CONTEST_ID), oData)
			Call oData.CommitTrans()
		Else
			' creating new contest
			Set oRS = Server.CreateObject("ADODB.Recordset")
			With oRS
				.Open "tblContests", oData.Connection, adOpenStatic, adLockOptimistic, adCmdTable
				.AddNew
				.Fields("vsContestName") = CleanForSQL(r_aData(m_CONTEST_NAME))
				.Fields("dtStartDate") = r_aData(m_CONTEST_START)
				.Fields("dtEndDate") = r_aData(m_CONTEST_END)
				.Fields("dtVoteByDate") = r_aData(m_CONTEST_VOTE_BY)
				.Fields("vsDescription") = CleanForSQL(r_aData(m_CONTEST_TEXT))
				.Fields("lVotesAllowed") = r_aData(m_CONTEST_VOTES)
				.Fields("lWeightFactor") = r_aData(m_CONTEST_WEIGHT)
				.Fields("lMaxEntries") = r_aData(m_CONTEST_MAX_ENTRIES)
				.Fields("bFreePluginsOnly") = r_aData(m_CONTEST_FREE_PLUGINS_ONLY)
				.Fields("bNoExternalMedia") = r_aData(m_CONTEST_NO_EXTERNAL_MEDIA)
				.Fields("lWinners") = r_aData(m_CONTEST_WINNERS)
				.Fields("lSiteID") = r_aData(m_CONTEST_SITE)
				.Update
				r_aData(m_CONTEST_ID) = .Fields("lContestID")
				.Close
			End With
			Set oRS = nothing
		End If
		Call SetSessionValue(g_USER_MSG, "The contest """ & r_aData(m_CONTEST_NAME) & """ has been saved")
		Call oData.LogActivity(g_ACT_SAVE_CONTEST, "", "", r_aData(m_CONTEST_ID), "", "", "")
		Set oData = Nothing
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		GetContestArray()
	'	Purpose: 	get contest data
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'	4/25/03		JEA		Get contest site
	'-------------------------------------------------------------------------
	Public Function GetContestArray(ByVal v_bNew)
		dim sQuery
		dim aData
		dim oData
		dim x
		dim aReturn(12)
		
		If Not v_bNew Then
			With Request
				aReturn(m_CONTEST_ID) = .QueryString("id")
				aReturn(m_CONTEST_NAME) = Trim(.Form("fldName"))
				If aReturn(m_CONTEST_NAME) <> "" Then
					aReturn(m_CONTEST_START) = .Form("fldStart")
					aReturn(m_CONTEST_END) = .Form("fldEnd")
					aReturn(m_CONTEST_VOTE_BY) = .Form("fldVoteBy")
					aReturn(m_CONTEST_TEXT) = .Form("fldDescription")
					aReturn(m_CONTEST_VOTES) = .Form("fldVotes")
					aReturn(m_CONTEST_WEIGHT) = .Form("fldWeight")
					aReturn(m_CONTEST_MAX_ENTRIES) = .Form("fldEntries")
					aReturn(m_CONTEST_FREE_PLUGINS_ONLY) = CBool(.Form("fldFreePlugins") = "on")
					aReturn(m_CONTEST_NO_EXTERNAL_MEDIA) = CBool(.Form("fldExternalMedia") = "on")
					aReturn(m_CONTEST_WINNERS) = .Form("fldWinners")
					aReturn(m_CONTEST_SITE) = .Form("fldSite")
					
				Else
					If IsNumber(aReturn(m_CONTEST_ID)) Then
						sQuery = "SELECT " & m_sBaseSQL & " AND lContestID = " & aReturn(m_CONTEST_ID)
					Else
						sQuery = "SELECT TOP 1 " & m_sBaseSQL & " ORDER BY dtEndDate DESC"
					End If
					Set oData = New kbDataAccess
					aData = oData.GetArray(sQuery)
					Set oData = Nothing
					
					If IsArray(aData) Then
						for x = 0 to UBound(aData)
							aReturn(x) = aData(x, 0)
						next
					End If
				End If
			End With
		End If
		GetContestArray = aReturn
	End Function

	'-------------------------------------------------------------------------
	'	Name: 		WriteContestList()
	'	Purpose: 	write option list with all contests
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteContestList(ByVal v_sFieldName, ByVal v_lSelectedID)
		dim sQuery
		sQuery = "SELECT lContestID, vsContestName FROM tblContests ORDER BY dtStartDate DESC"
		with response
			.write "<select name='"
			.write v_sFieldName
			.write "' onChange='SwitchContest(this);'>"
			.write MakeList(sQuery, v_lSelectedID)
			.write "</select>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteSiteList()
	'	Purpose: 	write option list with all sites
	'Modifications:
	'	Date:		Name:	Description:
	'	4/25/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteSiteList(ByVal v_lSelectedID)
		dim sQuery
		sQuery = "SELECT lSiteID, vsSiteName FROM tblSite ORDER BY vsSiteName"
		response.write MakeList(sQuery, v_lSelectedID)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetContestPoints()
	'	Purpose: 	get item points for contest
	'Modifications:
	'	Date:		Name:	Description:
	'	1/3/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function GetContestPoints(ByVal v_lContestID)
		dim sQuery
		dim oData
		sQuery = "SELECT CV.lItemID, F.vsFriendlyName, CV.lPoints FROM tblComputedVotePoints CV " _
			& "INNER JOIN tblProjects F ON F.lProjectID = CV.lItemID WHERE CV.lContestID = " _
			& v_lContestID & " ORDER BY CV.lPoints DESC"
		Set oData = New kbDataAccess
		GetContestPoints = oData.GetArray(sQuery)
		Set oData = Nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteContestStatus()
	'	Purpose: 	write progress bars for contest votes
	'Modifications:
	'	Date:		Name:	Description:
	'	1/3/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteContestStatus(ByVal v_lContestID)
		Const ITEM_ID = 0
		Const FRIENDLY_NAME = 1
		Const POINTS = 2
		dim aData
		dim x
		dim lMaxPoints
		dim lPercent
		dim lTotalPoints
	
		lMaxPoints = 0
		lTotalPoints = 0
		aData = GetContestPoints(v_lContestID)
		
		If IsArray(aData) Then
			lMaxPoints = aData(POINTS, 0)	' was sorted by points so first is max
			for x = 0 to UBound(aData, 2)
				lTotalPoints = lTotalPoints + aData(POINTS, x)
			next
			with response
				.write "<center><table cellspacing='0' cellpadding='0' border='0' width='70%'>"
				.write "<tr><td colspan='2' class='VoteHead'>Vote Status</td>"
				for x = 0 to UBound(aData, 2)
					lPercent = Int((aData(POINTS, x)/lMaxPoints) * 100)
					.write "<tr><td class='VoteLabel'><nobr><a href='kb_download.asp?id="
					.write aData(ITEM_ID, x)
					.write "'>"
					.write aData(FRIENDLY_NAME, x)
					.write "</a></nobr></td><td><table align='left' cellspacing='2' cellpadding='0' border='0' width='"
					.write (lPercent - 10)
					.write "%'><tr><td class='VoteBar'>"
					.write Round(((aData(POINTS, x)/lTotalPoints) * 100), 0)
					.write "%</td></table>"
					.write aData(POINTS, x)
					.write "</td>"
				next
				.write "</table></center>"
			end with
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteConclusion()
	'	Purpose: 	write contest conclusion
	'Modifications:
	'	Date:		Name:	Description:
	'	1/5/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteConclusion(ByVal v_aData)
		Const FILE_ID = 0
		Const FILE_NAME = 1
		Const USER_ID = 2
		Const USER_FIRST = 3
		Const USER_LAST = 4
		dim sQuery
		dim oData
		dim aData
		dim x
		
		sQuery = "SELECT TOP " & v_aData(m_CONTEST_WINNERS) _
			& " F.lProjectID, F.vsFriendlyName, U.lUserID, U.vsFirstName, U.vsLastName " _
			& "FROM (tblComputedVotePoints CVP INNER JOIN tblProjects F " _
			& "ON CVP.lItemID = F.lProjectID) INNER JOIN tblUsers U " _
			& "ON U.lUserID = F.lUserID WHERE CVP.lContestID = " _
			& v_aData(m_CONTEST_ID) & " ORDER BY CVP.lPoints DESC, U.vsFirstName"
		Set oData = New kbDataAccess
		aData = oData.GetArray(sQuery)
		Set oData = Nothing
		
		If IsArray(aData) Then
			with response
				.write "<table cellspacing='0' cellpadding='0' border='0'>"
				.write "<tr><td colspan='2'><div class='ContestName'>"
				.write v_aData(m_CONTEST_NAME)
				.write "</div><div class='Congrats'>congratulations to</div></td>"
				for x = 0 to UBound(aData, 2)
					.write "<tr><td class='Winner'><a href='kb_user.asp?id="
					.write aData(USER_ID, x)
					.write "'>"
					.write aData(USER_FIRST, x)
					.write " "
					.write aData(USER_LAST, x)
					.write "</a></td><td class='FileName'>for <a href='kb_download.asp?id="
					.write aData(FILE_ID, x)
					.write "'>"
					.write aData(FILE_NAME, x)
					.write "</a></td>"
				next
				.write "</table>"
			end with
		End If
	End Sub
End Class
%>