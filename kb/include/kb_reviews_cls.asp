<%
Const m_REVIEW_ID = 0
Const m_REVIEW_URL = 1
Const m_REVIEW_NAME = 2
Const m_REVIEW_TEXT = 3
Const m_REVIEW_DATE = 4
Const m_REVIEW_AUTHOR_ID = 5
Const m_REVIEW_AUTHOR_NAME = 6
Const m_REVIEW_RANK = 7
Const m_REVIEW_SUBMITTER = 8
Const m_REVIEW_CATS = 9		' generated from join

'-------------------------------------------------------------------------
'	Name: 		kbReview class
'	Purpose: 	methods for displaying review information
'Modifications:
'	Date:		Name:	Description:
'	7/21/04		JEA		Copied from tutorial class
'-------------------------------------------------------------------------
Class kbReviews
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
	'	Purpose: 	display reviews
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Public Sub WritePublic()
		Const sPAGE_NAME = "kb_reviews.asp"
		dim aData
		dim lItemsPerPage
		dim oReviewData
		dim oLayout
		dim lItemCount
		dim lStart
		dim lEnd
		dim x
		
		lItemsPerPage = GetSessionValue(g_USER_ITEMS_PER_PAGE)
		lEnd = (m_aFilter(g_FILTER_PAGE) * lItemsPerPage) - 1
		lStart = lEnd - (lItemsPerPage - 1)
		
		Set oReviewData = New kbReviewData
		aData = oReviewData.GetPublic(m_aFilter)
		Set oReviewData = Nothing
		
		with response
			If IsArray(aData) Then
				Set oLayout = New kbLayout
				lItemCount = UBound(aData, 2)
				If lEnd > lItemCount Then lEnd = lItemCount
				
				.write "<table width='80%' cellspacing='0' cellpadding='0' border='0' class='List'>"
				.write "<tr><td colspan='3'>"
				Call oLayout.WriteOptionHead(g_ITEM_REVIEW, m_aFilter, false)
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "reviews", oLayout)
				.write "</td>"
				Call oLayout.WriteItemListHead(sPAGE_NAME, m_aFilter)

				For x = lStart to lEnd
					aData(m_REVIEW_RANK,x) = MakeNumber(aData(m_REVIEW_RANK,x))
					.write "<tr><td rowspan='2' class='ItemName' valign='top'><a href='"
					If InStr(aData(m_REVIEW_URL,x), "frame-it") = 0 Then .write "http://"
					.write aData(m_REVIEW_URL,x)
					.write "' target='_new'>"
					.write aData(m_REVIEW_NAME,x)
					.write "</a>"
					.write "<div class='ItemAction'><a style='color: #777777;' href='kb_rank.asp?id="
					.write aData(m_REVIEW_ID,x)
					.write "&type="
					.write g_ITEM_REVIEW
					.write "'>"
					If aData(m_REVIEW_RANK,x) = 0 Then
						.write "rank this review"
					Else
						Call oLayout.WriteStars(aData(m_REVIEW_RANK,x), true)
					End If
					.write "</a></div>"
					' date
					.write "</td><td align='center' class='ItemDate' valign='top'>"
					.write FormatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), aData(m_REVIEW_DATE,x)))
					' author
					.write "</td><td align='center' class='ItemOwner' valign='top'><a href='kb_user.asp?id="
					.write aData(m_REVIEW_AUTHOR_ID,x)
					.write "'>"
					.write ReplaceNull(aData(m_REVIEW_AUTHOR_NAME,x), "&nbsp;")
					.write "</a></td><tr><td colspan='2'>"
					' description
					.write "<table width='100%' cellspacing='0' cellpadding='0' border='0'>"
					.write "<tr><td class='ItemDescription'>"
					.write FormatAsHTML(aData(m_REVIEW_TEXT,x))
					.write "</td><td class='ItemDescription'>&nbsp;</td>"
					.write "<tr><td class='ItemNotes'>"
					Call oLayout.WriteCategories(aData(m_REVIEW_CATS,x))
					.write "&nbsp;</td><td valign='bottom' class='ItemButton'>"
					If g_bAdmin Or _
						MatchesOne(GetSessionValue(g_USER_ID), Array(aData(m_REVIEW_AUTHOR_ID,x), aData(m_REVIEW_SUBMITTER,x)), true) Then
						
						.write "<div class='ItemAction'><a href='kb_review-edit.asp?id="
						.write aData(m_REVIEW_ID,x)
						.write "'>"
						Call oLayout.WriteToggleImage("btn_edit", "", "Edit Review", "", false)
						.write "</a></div>"
					End If
					.write "</td></table></td>"
				Next
				.write "<tr><td colspan='3'><div class='ListBottom'>"
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "reviews", oLayout)
				.write "</div></td></table>"
				
				Set oLayout = Nothing
			Else
				.write "No reviews were found"
			End If
		End With
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WritePending()
	'	Purpose: 	show reviews needing approval
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Public Sub WritePending()
		dim aData
		dim oLayout
		dim oData
		dim x

		Set oData = New kbReviewData
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
					.write aData(m_REVIEW_URL,x)
					.write "'>"
					.write aData(m_REVIEW_URL,x)
					.write "</a></nobr></td><td align='center' class='UploadDate'>"
					.write FormatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), aData(m_REVIEW_DATE, x)))
					.write "</td><td align='center' class='UploadOwner'><a href='kb_user.asp?id="
					.write aData(m_REVIEW_SUBMITTER,x)
					.write "'>"
					.write aData(m_REVIEW_AUTHOR_NAME,x)
					.write "</a></td><td class='UploadFriendlyName'>"
					.write aData(m_REVIEW_NAME,x)
					.write "</td><td class='UploadAction' valign='middle' rowspan='2'>"
					.write "<a href='kb_admin-uploads.asp?do=approvereview&id="
					.write aData(m_REVIEW_ID,x)
					.write "'>"
					Call oLayout.WriteToggleImage("btn_approve", "", "Approve", "", false)
					.write "</a> <a href='kb_admin-uploads.asp?do=denyreview&id="
					.write aData(m_REVIEW_ID,x)
					.write "'>"
					Call oLayout.WriteToggleImage("btn_deny", "", "Approve", "", false)
					.write "</a></td><tr><td colspan='4' class='UploadDescription'>"
					.write FormatAsHTML(ReplaceNull(aData(m_REVIEW_TEXT,x), "(no description given)"))
					.write "</td>"
				Next
				Set oLayout = Nothing
				.write "</table>"
			Else
				.write "There are no pending reviews"
			End If
		End With
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteAsHeader()
	'	Purpose: 	write summary for given review
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from tutorial class
	'-------------------------------------------------------------------------
	Public Sub WriteAsHeader(ByVal v_lReviewID)
		dim aReview
		dim oReviewData
		dim oLayout
		dim lRank
		dim sComment
		dim x
		
		Set oReviewData = New kbReviewData
		aReview = oReviewData.GetItem(v_lReviewID, false)
		Set oReviewData = Nothing
		
		with response
			If IsArray(aReview) Then
				' file info
				aReview(m_REVIEW_RANK, 0) = MakeNumber(aReview(m_REVIEW_RANK, 0))
				.write "<table cellspacing='0' cellpadding='3' border='0' width='450'>"
				.write "<tr><td valign='top'><div class='ItemName'><nobr>"
				.write aReview(m_REVIEW_NAME, 0)
				.write "</nobr></div><div class='ItemOwner'>by <a href='kb_user.asp?id="
				.write aReview(m_REVIEW_AUTHOR_ID, 0)
				.write "'>"
				.write aReview(m_REVIEW_AUTHOR_NAME, 0)
				.write "</a></div></td><tr><td class='ItemText' valign='top'>"
				.write FormatAsHTML(aReview(m_REVIEW_TEXT, 0))
				.write "</td><tr><td class='ItemRank' valign='bottom'>"
				if aReview(m_REVIEW_RANK, 0) > 0 then
					Set oLayout = New kbLayout
					Call oLayout.WriteStars(aReview(m_REVIEW_RANK, 0), false)
					Set oLayout = Nothing
				end if
				.write "</td></table>"
			Else
				.redirect "kb_projects.asp"
			End If
		end with
	End Sub
End Class
%>