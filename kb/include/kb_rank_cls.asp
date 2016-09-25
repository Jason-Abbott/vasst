<%
Class kbRank
	Private m_oData

	Private Sub Class_Initialize()
		Set m_oData = New kbDataAccess
	End Sub
	
	Private Sub Class_Terminate()
		Set m_oData = nothing
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		Save()
	'	Purpose: 	save file ranking
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub Save(ByVal v_lItemID, ByVal v_lItemTypeID, ByVal v_lStars, ByVal v_sComment)
		dim sQuery
		
		Call m_oData.BeginTrans()

		' remove any old ranking
		sQuery = "DELETE FROM tblRankings WHERE lItemID = " & v_lItemID _
			& " AND lItemTypeID = " & v_lItemTypeID & " AND lUserID = " _
			& GetSessionValue(g_USER_ID)
		Call m_oData.ExecuteOnly(sQuery)
		
		' insert new ranking
		sQuery = "INSERT INTO tblRankings (lItemID, lItemTypeID, lUserID, lRank, " _
			& "vsComment, dtRankDate) VALUES (" _
			& v_lItemID & ", " _
			& v_lItemTypeID & ", " _
			& GetSessionValue(g_USER_ID) & ", " _
			& v_lStars & ", '" _
			& CleanForSQL(v_sComment) & "', " & g_sSQL_DATE_DELIMIT _
			& Now() & g_sSQL_DATE_DELIMIT & ")"
		Call m_oData.ExecuteOnly(sQuery)
		If UpdateComputed(v_lItemID, v_lItemTypeID) Then ClearCache(v_lItemTypeID)
		Call m_oData.CommitTrans()
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Delete()
	'	Purpose: 	delete a ranking
	'Modifications:
	'	Date:		Name:	Description:
	'	1/22/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub Delete(ByVal v_lItemID, ByVal v_lItemTypeID, ByVal v_lUserID)
		dim sQuery
		
		If Not g_bAdmin Then Exit Sub
		
		Call m_oData.BeginTrans()
		sQuery = "DELETE FROM tblRankings WHERE lItemID = " & v_lItemID _
			& " AND lItemTypeID = " & v_lItemTypeID & " AND lUserID = " & v_lUserID
		Call m_oData.ExecuteOnly(sQuery)
		If UpdateComputed(v_lItemID, v_lItemTypeID) Then Call ClearCache(v_lItemTypeID)
		Call m_oData.CommitTrans()
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		UpdateComputed()
	'	Purpose: 	update total ranking for item; should run in transaction
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function UpdateComputed(ByVal v_lItemID, ByVal v_lItemTypeID)
		dim sQuery
		dim lOldRank
		dim lRank
		dim aData
		dim bUpdated
		
		bUpdated = false
		lOldRank = 0
		sQuery = "Rankings WHERE lItemID = " & v_lItemID & " AND lItemTypeID = " & v_lItemTypeID
		aData = m_oData.GetArray("SELECT fRank FROM tblComputed" & sQuery)
		If IsArray(aData) Then lOldRank = MakeNumber(aData(0,0))
		aData = m_oData.GetArray("SELECT AVG(lRank) FROM tbl" & sQuery)
		If IsArray(aData) Then
			lRank = Round((MakeNumber(aData(0,0)) * 2), 0) / 2
			If lRank <> lOldRank Then
				bUpdated = true
				Call m_oData.ExecuteOnly("DELETE FROM tblComputed" & sQuery)
				sQuery = "INSERT INTO tblComputedRankings (lItemID, lItemTypeID, fRank)" _
					& " VALUES (" & v_lItemID & ", " & v_lItemTypeID & ", " & lRank & ")"
				Call m_oData.ExecuteOnly(sQuery)
			End If
		End If
		UpdateComputed = bUpdated
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteItemRankings()
	'	Purpose: 	write rank log for item
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Function WriteItemRankings(ByVal v_lItemID, ByVal v_lItemTypeID)
		Const USER_ID = 0
		Const USER_NAME = 1
		Const RANK = 2
		Const COMMENT = 3
		Const RANK_DATE = 4
		Const MAX_STARS = 5
		dim aRanking
		dim lRank
		dim sComment
		dim oLayout
		dim x
		
		aRanking = GetItemRankings(v_lItemID, v_lItemTypeID)
		lRank = 5
		
		If IsArray(aRanking) Then
			sComment = ""
			for x = 0 to UBound(aRanking, 2)
				if CStr(aRanking(USER_ID, x)) = CStr(GetSessionValue(g_USER_ID)) then
					lRank = MakeNumber(aRanking(RANK, x))
					sComment = aRanking(COMMENT, x)
					exit for
				end if
			next
		End If
		
		Set oLayout = New kbLayout
		with response
			.write "<table cellspacing='0' border='0' cellpadding='1' width='500'>"
			.write "<tr><td>Stars: <select name='fldStars'>"
			for x = 1 to MAX_STARS
				.write "<option"
				if lRank = x then .write " selected"
				.write ">"
				.write x
			next
			.write "</select></td><td>Rank and comment</td>"
			.write "<td align='right'>"
			Call oLayout.WriteToggleImage("btn_save", "", "Save Ranking", "class='Image'", true)
			.write "<tr><td colspan='3' class='Comment'><textarea class='Comment' "
			.write "name='fldComment' cols='100' rows='3'>"
			.write sComment
			.write "</textarea><div class='Note'>"
			.write g_sMSG_HTML_LIMIT
			.write "</div></td>"
			
			If IsArray(aRanking) Then
				for x = 0 to UBound(aRanking, 2)
					.write "<tr><td class='Rank'>"
					Call oLayout.WriteStars(aRanking(RANK, x), false)
					.write "</td><td class='UserName'><a href='kb_user.asp?id="
					.write aRanking(USER_ID, x)
					.write "'>"
					.write aRanking(USER_NAME, x)
					.write "</a></td><td align='right' class='Date'>"
					.write FormatDate(aRanking(RANK_DATE, x))
					.write "</td><tr><td colspan='3' class='Comment'>"
					if g_bAdmin then
						' allow rankings to be deleted
						.write "<a href='kb_rank.asp?id="
						.write v_lItemID
						.write "&type="
						.write v_lItemTypeID
						.write "&user="
						.write aRanking(USER_ID, x)
						.write "&do=delete'>"
						Call oLayout.WriteToggleImage("btn_delete", "", "Delete Ranking", "class='Image' align='right'", false)
						.write "</a>"
					end if
					.write FormatAsHTML(aRanking(COMMENT, x))
					.write "</td>"
				next
			End If
			.write "</table>"
		end with
		Set oLayout = Nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetItemRankings()
	'	Purpose: 	get array of rankings for this item
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function GetItemRankings(ByVal v_lItemID, ByVal v_lItemTypeID)
		dim sQuery
		sQuery = "SELECT U.lUserID, IIf(U.vsScreenName IS NULL, " _
			& "U.vsFirstName + ' ' + U.vsLastName, U.vsScreenName), R.lRank, R.vsComment, R.dtRankDate " _
			& "FROM tblRankings R INNER JOIN tblUsers U ON U.lUserID = R.lUserID " _
			& "WHERE R.lItemID = " & v_lItemID & " AND R.lItemTypeID = " & v_lItemTypeID _
			& " ORDER BY R.dtRankDate"
		GetItemRankings = m_oData.GetArray(sQuery)
	End Function
End Class
%>