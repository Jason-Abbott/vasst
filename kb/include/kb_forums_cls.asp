<%
Class kbForums
	Private m_aFilter(4)
	
	'Private Sub Class_Initialize()
	'	Set m_oFileData = New kbProjectData
	'End Sub
	
	'Private Sub Class_Terminate()
	'	Set m_oFileData = Nothing
	'End Sub
	
	Public Property Let SortBy(v_lSortID)
		m_aFilter(g_FILTER_SORT) = MakeNumber(ReplaceNull(v_lSortID, g_SORT_OWNER_ASC))
	End Property
	
	Public Property Let Category(v_lCategoryID)
		m_aFilter(g_FILTER_CATEGORY) = MakeNumber(v_lCategoryID)
	End Property
	
	Public Property Let Page(ByVal v_lPage)
		m_aFilter(g_FILTER_PAGE) = MakeNumber(ReplaceNull(v_lPage, 1))
	End Property

	Public Property Let Software(ByVal v_lSoftwareID)
		m_aFilter(g_FILTER_SOFTWARE) = MakeNumber(v_lSoftwareID)
	End Property
	
	Public Property Let Author(ByVal v_lAuthorID)
		m_aFilter(g_FILTER_AUTHOR) = MakeNumber(v_lAuthorID)
	End Property
	
	'-------------------------------------------------------------------------
	'	Name: 		WritePublic()
	'	Purpose: 	display forums
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WritePublic()
		Const sPAGE_NAME = "kb_forums.asp"
		dim aData
		dim lItemsPerPage
		dim oForumData
		dim oLayout
		dim lItemCount
		dim lStart
		dim lEnd
		dim x
		
		lItemsPerPage = GetSessionValue(g_USER_ITEMS_PER_PAGE)
		lEnd = (m_aFilter(g_FILTER_PAGE) * lItemsPerPage) - 1
		lStart = lEnd - (lItemsPerPage - 1)
		
		Set oForumData = New kbForumData
		aData = oForumData.GetPublic(m_aFilter)
		Set oForumData = Nothing
		
		Set oLayout = New kbLayout
		with response
			If IsArray(aData) Then
				lItemCount = UBound(aData, 2)
				If lEnd > lItemCount Then lEnd = lItemCount
				
				.write "<table width='80%' cellspacing='0' cellpadding='0' border='0' class='List'>"
				.write "<tr><td colspan='3'>"
				Call oLayout.WriteOptionHead(g_ITEM_FORUM, m_aFilter, true)
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "forums", oLayout)
				.write "</td>"
				Call WriteItemListHead(sPAGE_NAME, m_aFilter(g_FILTER_SORT))

				For x = lStart to lEnd
					aData(m_FORUM_RANK,x) = MakeNumber(aData(m_FORUM_RANK,x))
					' forum host
					.write "<tr><td class='ItemOwner' valign='top'><nobr><a href='http://"
					.write aData(m_FORUM_HOST_URL,x)
					.write "' target='_new'>"
					.write aData(m_FORUM_HOST_NAME,x)
					.write "</a></nobr>:</td>"
					' forum name
					.write "<td class='ItemName' valign='top'><a href='http://"
					.write aData(m_FORUM_URL,x)
					.write "' target='_new'>"
					.write aData(m_FORUM_NAME,x)
					.write "</a></td>"
					' forum description
					.write "<td rowspan='2' class='ItemDate'>"
					.write "<table width='100%' cellspacing='0' cellpadding='0' border='0'>"
					.write "<tr><td class='ItemDescription'>"
					.write FormatAsHTML(aData(m_FORUM_TEXT,x))
					.write "</td><td valign='bottom' class='ItemButton' rowspan='2'>"
					If g_bAdmin Then
						.write "<a href='kb_forum-edit.asp?id="
						.write aData(m_FORUM_ID,x)
						.write "'>"
						Call oLayout.WriteToggleImage("btn_edit", "", "Edit Forum", "", false)
						.write "</a>"
					End If
					.write "</td><tr><td class='ItemNotes'>"
					Call oLayout.WriteCategories(aData(m_FORUM_CATS,x))
					.write "</td></table></td>"
					' ranking
					.write "<tr><td colspan='2'><div class='ItemAction'>"
					.write "<a style='color: #777777;' href='kb_rank.asp?id="
					.write aData(m_FORUM_ID,x)
					.write "&type="
					.write g_ITEM_FORUM
					.write "'>"
					If aData(m_FORUM_RANK,x) = 0 Then
						.write "rank this forum"
					Else
						Call oLayout.WriteStars(aData(m_FORUM_RANK,x), true)
					End If
					.write "</a></div></td>"
				Next
				.write "<tr><td colspan='3'><div class='ListBottom'>"
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "forums", oLayout)
				.write "</div></td></table>"
			Else
				.write "<p>No forums were found"
				If g_bAdmin Then
					.write " <a href='kb_forum-edit.asp'>"
					Call oLayout.WriteToggleImage("btn_add-forum", "", "Add " & g_ITEM_FORUM, "", false)
					.write "</a>"
				End If
			End If
		end with
		Set oLayout = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteItemListHead()
	'	Purpose: 	write heading for forum list in three column table
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'	4/29/03		JEA		Make site-specific arrow
	'-------------------------------------------------------------------------
	Private Sub WriteItemListHead(ByVal v_sPageURL, ByVal v_lSortID)
		Const NAME = 0
		Const SORT_ASC = 1
		Const SORT_DESC = 2
		dim sArrow
		dim aHeadings
		dim x
		
		aHeadings = Array( _
			Array("Host", g_SORT_OWNER_ASC, g_SORT_OWNER_DESC), _
			Array("Name", g_SORT_NAME_ASC, g_SORT_NAME_DESC))

		sArrow = "<img class='SortArrow' src='./images/" & g_lSiteID & "/arrow_" & IIf((v_lSortID Mod 2), "down", "up") & ".gif'>"
		with response
			.write "<tr>"
			for x = 0 to UBound(aHeadings)
				.write "<td class='ItemHead'><a class='ItemHead' href='"
				.write v_sPageURL
				.write "?sort="
				.write IIf((v_lSortID = aHeadings(x)(SORT_ASC)), aHeadings(x)(SORT_DESC), aHeadings(x)(SORT_ASC))
				.write "'>"
				.write aHeadings(x)(NAME)
				.write "</a>"
				if MatchesOne(v_lSortID, Array(aHeadings(x)(SORT_DESC), aHeadings(x)(SORT_ASC)), true) then .write sArrow
				.write "</td>"
			next
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteAsHeader()
	'	Purpose: 	write summary for given forum
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteAsHeader(ByVal v_lForumID)
		dim aForum
		dim oForumData
		dim oLayout
		dim lRank
		dim sComment
		dim x
		
		Set oForumData = New kbForumData
		aForum = oForumData.GetItem(v_lForumID, false)
		Set oForumData = Nothing
		
		with response
			If IsArray(aForum) Then
				' file info
				aForum(m_FORUM_RANK, 0) = MakeNumber(aForum(m_FORUM_RANK, 0))
				.write "<table cellspacing='0' cellpadding='3' border='0' width='450'>"
				.write "<td valign='top'><div class='ItemName'><nobr>"
				.write aForum(m_FORUM_NAME, 0)
				.write "</nobr></div><div class='ItemOwner'>by <a href='http://"
				.write aForum(m_FORUM_URL, 0)
				.write "'>"
				.write aForum(m_FORUM_HOST_NAME, 0)
				.write "</a></div></td><td rowspan='2' class='ItemText' valign='top'>"
				.write aForum(m_FORUM_TEXT, 0)
				.write "</td><tr><td class='ItemRank' valign='bottom'>"
				if aForum(m_FORUM_RANK, 0) > 0 then
					Set oLayout = New kbLayout
					Call oLayout.WriteStars(aForum(m_FORUM_RANK, 0), false)
					Set oLayout = Nothing
				end if
				.write "</td></table>"
			Else
				.redirect "kb_forums.asp"
			End If
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteHostList()
	'	Purpose: 	write option list with forum hosts
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteHostList(ByVal v_sFieldname, ByVal v_lSelectedID)
		dim sQuery
		sQuery = "SELECT lForumHostID, vsHostName FROM tblForumHosts ORDER BY vsHostName"
		with response
			.write "<select name='"
			.write v_sFieldName
			.write "'>"
			.write MakeList(sQuery, v_lSelectedID)
			.write "</select>"
		end with
	End Sub
End Class
%>