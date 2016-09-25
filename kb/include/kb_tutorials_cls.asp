<%
Const m_TUTORIAL_ID = 0
Const m_TUTORIAL_URL = 1
Const m_TUTORIAL_NAME = 2
Const m_TUTORIAL_TEXT = 3
Const m_TUTORIAL_DATE = 4
Const m_TUTORIAL_AUTHOR_ID = 5
Const m_TUTORIAL_AUTHOR_NAME = 6
Const m_TUTORIAL_RANK = 7
Const m_TUTORIAL_SUBMITTER = 8
Const m_TUTORIAL_CATS = 9		' generated from join

'-------------------------------------------------------------------------
'	Name: 		kbTutorials class
'	Purpose: 	methods for displaying tutorial information
'Modifications:
'	Date:		Name:	Description:
'	1/4/02		JEA		Creation
'-------------------------------------------------------------------------
Class kbTutorials
	Private m_aFilter(4)
	
	Public Property Let SortBy(v_lSortID)
		m_aFilter(g_FILTER_SORT) = MakeNumber(ReplaceNull(v_lSortID, g_SORT_DATE_DESC))
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
	'	Purpose: 	display tutorials
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WritePublic()
		Const sPAGE_NAME = "kb_tutorials.asp"
		dim aData
		dim lItemsPerPage
		dim oTutorialData
		dim oLayout
		dim lItemCount
		dim lStart
		dim lEnd
		dim x
		
		lItemsPerPage = GetSessionValue(g_USER_ITEMS_PER_PAGE)
		lEnd = (m_aFilter(g_FILTER_PAGE) * lItemsPerPage) - 1
		lStart = lEnd - (lItemsPerPage - 1)
		
		Set oTutorialData = New kbTutorialData
		aData = oTutorialData.GetPublic(m_aFilter)
		Set oTutorialData = Nothing
		
		with response
			If IsArray(aData) Then
				Set oLayout = New kbLayout
				lItemCount = UBound(aData, 2)
				If lEnd > lItemCount Then lEnd = lItemCount
				
				.write "<table width='80%' cellspacing='0' cellpadding='0' border='0' class='List'>"
				.write "<tr><td colspan='3'>"
				Call oLayout.WriteOptionHead(g_ITEM_TUTORIAL, m_aFilter, false)
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "tutorials", oLayout)
				.write "</td>"
				Call oLayout.WriteItemListHead(sPAGE_NAME, m_aFilter)

				For x = lStart to lEnd
					aData(m_TUTORIAL_RANK,x) = MakeNumber(aData(m_TUTORIAL_RANK,x))
					.write "<tr><td rowspan='2' class='ItemName' valign='top'><a href='"
					If InStr(aData(m_TUTORIAL_URL,x), "frame-it") = 0 Then .write "http://"
					.write aData(m_TUTORIAL_URL,x)
					.write "' target='_new'>"
					.write aData(m_TUTORIAL_NAME,x)
					.write "</a>"
					.write "<div class='ItemAction'><a style='color: #777777;' href='kb_rank.asp?id="
					.write aData(m_TUTORIAL_ID,x)
					.write "&type="
					.write g_ITEM_TUTORIAL
					.write "'>"
					If aData(m_TUTORIAL_RANK,x) = 0 Then
						.write "rank this tutorial"
					Else
						Call oLayout.WriteStars(aData(m_TUTORIAL_RANK,x), true)
					End If
					.write "</a></div>"
					.write "</td><td align='center' class='ItemDate' valign='top'>"
					.write FormatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), aData(m_TUTORIAL_DATE,x)))
					.write "</td><td align='center' class='ItemOwner' valign='top'><a href='kb_user.asp?id="
					.write aData(m_TUTORIAL_AUTHOR_ID,x)
					.write "'>"
					.write ReplaceNull(aData(m_TUTORIAL_AUTHOR_NAME,x), "&nbsp;")
					.write "</a></td><tr><td colspan='2'>"
					.write "<table width='100%' cellspacing='0' cellpadding='0' border='0'>"
					.write "<tr><td class='ItemDescription'>"
					.write FormatAsHTML(aData(m_TUTORIAL_TEXT,x))
					.write "</td>"
					.write "<tr><td class='ItemNotes'>"
					Call oLayout.WriteCategories(aData(m_TUTORIAL_CATS,x))
					.write "</td><td valign='bottom' class='ItemButton'>"
					If g_bAdmin Or _
						MatchesOne(GetSessionValue(g_USER_ID), Array(aData(m_TUTORIAL_AUTHOR_ID,x), aData(m_TUTORIAL_SUBMITTER,x)), true) Then
						
						.write "<div class='ItemAction'><a href='kb_tutorial-edit.asp?id="
						.write aData(m_TUTORIAL_ID,x)
						.write "'>"
						Call oLayout.WriteToggleImage("btn_edit", "", "Edit Tutorial", "", false)
						.write "</a></div>"
					End If
					.write "</td></table></td>"
				Next
				.write "<tr><td colspan='3'><div class='ListBottom'>"
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "tutorials", oLayout)
				.write "</div></td></table>"
				
				Set oLayout = Nothing
			Else
				.write "No tutorials were found"
			End If
		End With
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WritePending()
	'	Purpose: 	show tutorials needing approval
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/04		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WritePending()
		dim aData
		dim oLayout
		dim oData
		dim x

		Set oData = New kbTutorialData
		aData = oData.GetPending()
		Set oData = Nothing
		
		With Response
			If IsArray(aData) Then
				.write "<table cellspacing='0' cellpadding='0' border='0' width='700'><tr>"
				.write "<td class='UploadsHead'>URL</td>"
				.write "<td class='UploadsHead'>Submit Date</td>"
				.write "<td class='UploadsHead'>Submitted By</td>"
				.write "<td class='UploadsHead'>Name</td>"
				.write "<td class='UploadsHead'>Action</td>"
				Set oLayout = New kbLayout
				For x = 0 to UBound(aData,2)
					.write "<tr><td class='UploadName'><nobr><a href='http://"
					.write aData(m_TUTORIAL_URL,x)
					.write "'>"
					.write aData(m_TUTORIAL_URL,x)
					.write "</a></nobr></td><td align='center' class='UploadDate'>"
					.write FormatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), aData(m_TUTORIAL_DATE, x)))
					.write "</td><td align='center' class='UploadOwner'><a href='kb_user.asp?id="
					.write aData(m_TUTORIAL_SUBMITTER,x)
					.write "'>"
					.write aData(m_TUTORIAL_AUTHOR_NAME,x)
					.write "</a></td><td class='UploadFriendlyName'>"
					.write aData(m_TUTORIAL_NAME,x)
					.write "</td><td class='UploadAction' valign='middle' rowspan='2'>"
					.write "<a href='kb_admin-uploads.asp?do=approvetut&id="
					.write aData(m_TUTORIAL_ID,x)
					.write "'>"
					Call oLayout.WriteToggleImage("btn_approve", "", "Approve", "", false)
					.write "</a> <a href='kb_admin-uploads.asp?do=denytut&id="
					.write aData(m_TUTORIAL_ID,x)
					.write "'>"
					Call oLayout.WriteToggleImage("btn_deny", "", "Approve", "", false)
					.write "</a></td><tr><td colspan='4' class='UploadDescription'>"
					.write FormatAsHTML(ReplaceNull(aData(m_TUTORIAL_TEXT,x), "(no description given)"))
					.write "</td>"
				Next
				Set oLayout = Nothing
				.write "</table>"
			Else
				.write "There are no pending tutorials"
			End If
		End With
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteAsHeader()
	'	Purpose: 	write summary for given tutorial
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteAsHeader(ByVal v_lTutorialID)
		dim aTutorial
		dim oTutorialData
		dim oLayout
		dim lRank
		dim sComment
		dim x
		
		Set oTutorialData = New kbTutorialData
		aTutorial = oTutorialData.GetItem(v_lTutorialID, false)
		Set oTutorialData = Nothing
		
		with response
			If IsArray(aTutorial) Then
				' file info
				aTutorial(m_TUTORIAL_RANK, 0) = MakeNumber(aTutorial(m_TUTORIAL_RANK, 0))
				.write "<table cellspacing='0' cellpadding='3' border='0' width='450'>"
				.write "<td valign='top'><div class='ItemName'><nobr>"
				.write aTutorial(m_TUTORIAL_NAME, 0)
				.write "</nobr></div><div class='ItemOwner'>by <a href='kb_user.asp?id="
				.write aTutorial(m_TUTORIAL_AUTHOR_ID, 0)
				.write "'>"
				.write aTutorial(m_TUTORIAL_AUTHOR_NAME, 0)
				.write "</a></div></td><td rowspan='2' class='ItemText' valign='top'>"
				.write aTutorial(m_TUTORIAL_TEXT, 0)
				.write "</td><tr><td class='ItemRank' valign='bottom'>"
				if aTutorial(m_TUTORIAL_RANK, 0) > 0 then
					Set oLayout = New kbLayout
					Call oLayout.WriteStars(aTutorial(m_TUTORIAL_RANK, 0), false)
					Set oLayout = Nothing
				end if
				.write "</td></table>"
			Else
				.redirect "kb_files.asp"
			End If
		end with
	End Sub
End Class
%>