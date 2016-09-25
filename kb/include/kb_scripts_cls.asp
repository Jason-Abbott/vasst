<%
'-------------------------------------------------------------------------
'	Name: 		kbScripts class
'	Purpose: 	methods for displaying script information
'Modifications:
'	Date:		Name:	Description:
'	7/21/04		JEA		Copied from files class
'-------------------------------------------------------------------------
Class kbScripts
	Private m_aFilter(4)
	
	'Private Sub Class_Initialize()
	'	Set m_oScriptData = New kbScriptData
	'End Sub
	
	'Private Sub Class_Terminate()
	'	Set m_oScriptData = Nothing
	'End Sub
	
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
	'	Purpose: 	output scripts as HTML
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Sub WritePublic()
		Const sPAGE_NAME = "kb_scripts.asp"
		dim aData
		dim lItemsPerPage
		dim oLayout
		dim lItemCount
		dim lStart
		dim lEnd
		dim oScriptData
		dim sLastLogin
		dim bBreak
		dim x, y
		
		lItemsPerPage = GetSessionValue(g_USER_ITEMS_PER_PAGE)
		lEnd = (m_aFilter(g_FILTER_PAGE) * lItemsPerPage) - 1
		lStart = lEnd - (lItemsPerPage - 1)
		sLastLogin = GetSessionValue(g_USER_LAST_LOGIN)

		Set oScriptData = New kbScriptData
		aData = oScriptData.GetPublic(m_aFilter)
		Set oScriptData = Nothing
		
		With Response
			Set oLayout = New kbLayout
			.write "<table width='80%' cellspacing='0' cellpadding='0' border='0'>"
			.write "<tr><td colspan='3'>"
			Call oLayout.WriteOptionHead(g_ITEM_SCRIPT, m_aFilter, false)
		
			If IsArray(aData) Then
				lItemCount = UBound(aData, 2)
				If lEnd > lItemCount Then lEnd = lItemCount
				.write "</td>"
				Call oLayout.WriteItemListHead(sPAGE_NAME, m_aFilter)
				.Write "<tr><td colspan='3' class='PagingRow'>"
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "scripts", oLayout)
				.Write "</td>"
				
				For x = lStart to lEnd
					bBreak = false
					aData(m_SCRIPT_RANK,x) = MakeNumber(aData(m_SCRIPT_RANK,x))
					.write "<tr><td class='ItemName' rowspan='3' valign='top'><nobr><a href='kb_script-download.asp?id="
					.write aData(m_SCRIPT_ID,x)
					.write "'>"
					.write aData(m_SCRIPT_FRIENDLY_NAME,x)
					.write "</a></nobr>"
					' downloads
					.write "<div class='ItemDownloads'>"
					.write aData(m_SCRIPT_DOWNLOADS,x)
					.write " downloads</div>"
					' ranking
					.write "<div class='ItemAction'><a style='color: #777777;' href='kb_rank.asp?id="
					.write aData(m_SCRIPT_ID,x)
					.write "&type="
					.write g_ITEM_SCRIPT
					.write "'>"
					If aData(m_SCRIPT_RANK,x) = 0 Then
						.write "rank this script"
					Else
						Call oLayout.WriteStars(aData(m_SCRIPT_RANK,x), true)
					End If
					.write "</a></div>"
					' edit button
					If g_bAdmin Or CStr(aData(m_SCRIPT_USER_ID,x)) = CStr(GetSessionValue(g_USER_ID)) Then
						.write "<div class='ItemEdit'><a href='kb_script-edit.asp?id="
						.write aData(m_SCRIPT_ID,x)
						.write "&url="
						.write Server.URLEncode(GetURL(false))
						.write "'>"
						Call oLayout.WriteToggleImage("btn_edit", "", "Edit Entry", "", false)
						.write "</a></div>"
					End If
					' item date
					.write "</td><td align='center' class='ItemDate'>"
					.write FormatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), aData(m_SCRIPT_DATE_ADDED,x)))
					If aData(m_SCRIPT_VERSION,x) > 1 And DateDiff("d", aData(m_SCRIPT_DATE_ADDED,x), Date()) <= g_SHOW_AS_NEW_DAYS Then
						.write "<img src='./images/"
						.write g_lSiteID
						.write "/new-version.gif' width='61' height='11'>"
					ElseIf DateDiff("d", aData(m_SCRIPT_DATE_ADDED,x), sLastLogin) < 0 Then
						.write "<img src='./images/"
						.write g_lSiteID
						.write "/new.gif' width='29' height='11'>"
					End If
					.write "</td><td align='center' class='ItemOwner'><a href='kb_user.asp?id="
					.write aData(m_SCRIPT_USER_ID,x)
					.write "'>"
					.write aData(m_SCRIPT_CREATOR,x)
					.write "</a></td><tr><td colspan='2' class='ItemDescription'>"
					.write FormatAsHTML(aData(m_SCRIPT_DESCRIPTION,x))
					if Trim(aData(m_SCRIPT_MEDIA, x)) <> "" then
						.write " (requires <a href='http://"
						.write aData(m_SCRIPT_MEDIA, x)
						.write "'>additional media</a>)"
					end if
					' item notes
					.write "</td><tr><td class='ItemNotes'>"
					Call oLayout.WriteCategories(aData(m_SCRIPT_CATS,x))
					.write "&nbsp;</td><td align='right' class='ItemVersion'>"
					' version and icon
					.write "<table cellspacing='0' cellpadding='0' border='0'><tr>"
					.write "<td align='right' style='font-size: 8pt;'>"
					If aData(m_SCRIPT_SOFTWARE_NAME,x) <> "" Then
						.write Replace(aData(m_SCRIPT_SOFTWARE_NAME,x), " (any)", "")
						bBreak = true
					End If
					' product icon
					.write "</td></table></td>"
				Next
				.write "<tr><td colspan='3'><div class='ListBottom'>"
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "scripts", oLayout)
				.write "</div></td></table>"
			Else
				.write "<br><center>No scripts match your criteria</center></td></table>"
			End If
			Set oLayout = Nothing
		End With
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WritePending()
	'	Purpose: 	output files as HTML
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Sub WritePending()
		dim aData
		dim oLayout
		dim oScriptData
		dim bBreak
		dim x

		Set oScriptData = New kbScriptData
		aData = oScriptData.GetPending()
		Set oScriptData = Nothing
		
		With Response
			If IsArray(aData) Then
				.write "<table cellspacing='0' cellpadding='0' border='0' width='700'><tr>"
				.write "<td class='UploadsHead'>File Name</td>"
				.write "<td class='UploadsHead'>Upload Date</td>"
				.write "<td class='UploadsHead'>Submitted By</td>"
				.write "<td class='UploadsHead'>Friendly Name</td>"
				.write "<td class='UploadsHead'>Action</td>"
				Set oLayout = New kbLayout
				For x = 0 to UBound(aData,2)
					.write "<tr><td class='UploadName'><nobr><a href='kb_script-download.asp?id="
					.write aData(m_SCRIPT_ID,x)
					.write "'>"
					.write aData(m_SCRIPT_NAME,x)
					.write "</a></nobr></td><td align='center' class='UploadDate'>"
					.write FormatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), aData(m_SCRIPT_DATE_ADDED, x)))
					.write "</td><td align='center' class='UploadOwner'><a href='kb_user.asp?id="
					.write aData(m_SCRIPT_USER_ID,x)
					.write "'>"
					.write aData(m_SCRIPT_CREATOR,x)
					.write "</a></td><td class='UploadFriendlyName'>"
					.write aData(m_SCRIPT_FRIENDLY_NAME,x)
					.write "</td><td class='UploadAction' valign='middle' rowspan='2'><nobr>"
					.write "<a href='kb_admin-uploads.asp?do=approvescript&id="
					.write aData(m_SCRIPT_ID,x)
					.write "'>"
					Call oLayout.WriteToggleImage("btn_approve", "", "Approve", "", false)
					.write "</a> <a href='kb_admin-uploads.asp?do=denyscript&id="
					.write aData(m_SCRIPT_ID,x)
					.write "'>"
					Call oLayout.WriteToggleImage("btn_deny", "", "Approve", "", false)
					.write "</a></nobr></td><tr><td colspan='4' class='UploadDescription'>"
					.write FormatAsHTML(ReplaceNull(aData(m_SCRIPT_DESCRIPTION,x), "(no description given)"))
					if Trim(aData(m_SCRIPT_MEDIA, x)) <> "" then
						.write " (requires <a href='http://"
						.write aData(m_SCRIPT_MEDIA, x)
						.write "'>additional media</a>)"
					end if
					.write "</td>"
				Next
				Set oLayout = Nothing
				.write "</table>"
			Else
				.write "There are no pending scripts"
			End If
		End With
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteEditForm()
	'	Purpose: 	write box for editing single script file
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Sub WriteEditForm(ByVal v_lScriptID)
		dim aData
		dim oLayout
		dim oScriptData
		dim x
		
		Set oScriptData = New kbScriptData
		aData = oScriptData.GetItem(v_lScriptID, true)
		Set oScriptData = Nothing
		
		with response
			If IsArray(aData) Then
				Set oLayout = New kbLayout
				Call oLayout.WriteTitleBoxTop("Script Edit", "", "")
				.write "<table cellspacing='0' cellpadding='0' border='0'>"
				' file name
				.write "<tr><td class='FormLabel'>Script:</td><td class='FormInput'>"
				.write "<input type='file' name='fldFile' value='"
				.write aData(m_SCRIPT_NAME, 0)
				.write "' size='35'></td><tr><td></td><td class='FormNote'>Leave blank"
				.write " except to upload a new version of the file</td>"
				' listed name
				.write "</td><tr><td class='Required'>Listed Name:</td><td class='FormInput'>"
				.write "<input type='text' name='fldFriendlyName' value='"
				.write aData(m_SCRIPT_FRIENDLY_NAME, 0)
				.write "' maxlength='30' size='30'></td>"
				' software
				.write "<tr><td class='Required'>Requires:</td><td class='FormInput'>"
				Call oLayout.WriteVersionList("fldVersion", aData(m_SCRIPT_SOFTWARE_ID, 0), g_ITEM_SCRIPT)
				.write " software required to run the script</td>"
				' media URL				
				.write "<tr><td class='FormLabel'>Resources URL:</td><td class='FormInput'>"
				.write "<input type='text' name='fldMedia' maxlength='75' size='30' value='"
				.write aData(m_SCRIPT_MEDIA, 0)
				.write "'> link to any extra files needed</td><tr><td></td><td class='FormNote'><nobr>"
				.write g_sMSG_MEDIA_HINT
				.write "</nobr></td>"
				' description
				.write "<tr><td class='Required' valign='top'>Description:</td>"
				.write "<td class='FormInput' valign='top'><textarea rows='6' cols='52' name='fldDescription'>"
				.write aData(m_SCRIPT_DESCRIPTION, 0)
				.write "</textarea></td><tr><td></td><td class='FormNote'>"
				.write g_sMSG_HTML_LIMIT
				' categories
				.write "</td><tr><td class='FormLabel' valign='top'>Categories:</td>"
				.Write "<td class='FormInput' valign='top'>"
				Call oLayout.WriteCategoryList("fldCategories", aData(m_SCRIPT_CATS, x), 4, g_ITEM_SCRIPT)
				.write "</td>"
				' buttons
				.write "<tr><td class='Required' valign='bottom' style='font-size: 8pt; text-align: center;'>"
				.write "(required)</td><td align='right' valign='bottom'><a href='javascript:DeleteFile()'>"
				Call oLayout.WriteToggleImage("btn_delete", "", "Delete This File", "width='53' height='14'", false)
				.write "</a> "
				Call oLayout.WriteToggleImage("btn_save", "", "Save Changes", "width='53' height='14' class='Image'", true)
				.write "</td></table>"
				Call oLayout.WriteBoxBottom("")
				Set oLayout = Nothing
			Else
				.write "Script not found or not permissable"
			End If
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteVersionList()
	'	Purpose: 	write option list with software versions
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Sub WriteVersionList(ByVal v_sFieldName, ByVal v_lSelectedID, ByVal v_lItemTypeID)
		dim sQuery
	
		sQuery = "SELECT SV.lVersionID, vsSoftwareName + ' ' + vsVersionText " _
			& "FROM ((tblSoftware S " _
			& "INNER JOIN tblSoftwareVersions SV ON S.lSoftwareID = SV.lSoftwareID) " _
			& "INNER JOIN tblPublishers P ON P.lPublisherID = S.lPublisherID) " _
			& "INNER JOIN (SELECT lItemID FROM tblItemSites " _
			& 	"WHERE lItemTypeID = " & g_ITEM_PUBLISHER & " AND lSiteID = " & GetSessionValue(g_USER_SITE) _
			&	") tIS ON tIS.lItemID = P.lPublisherID " _
			& "ORDER BY vsSoftwareName, vsVersionText"

		with response
			.write "<select name='"
			.write v_sFieldName
			.write "'>"
			.write MakeList(sQuery, v_lSelectedID)
			.write "</select>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteAsHeader()
	'	Purpose: 	write summary for given script
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Sub WriteAsHeader(ByVal v_lScriptID)
		dim aFile
		dim oScriptData
		dim oLayout
		dim lRank
		dim sComment
		dim x
		
		Set oScriptData = New kbScriptData
		aFile = oScriptData.GetItem(v_lScriptID, false)
		Set oScriptData = Nothing
		
		with response
			If IsArray(aFile) Then
				' file info
				aFile(m_SCRIPT_RANK, 0) = MakeNumber(aFile(m_SCRIPT_RANK, 0))
				.write "<table cellspacing='0' cellpadding='3' border='0' width='450'>"
				.write "<td valign='top' colspan='2'><div class='ItemName'><nobr>"
				.write aFile(m_SCRIPT_FRIENDLY_NAME, 0)
				.write "</nobr></div></td><tr><td class='ItemRank'>"
				if aFile(m_SCRIPT_RANK, 0) > 0 then
					Set oLayout = New kbLayout
					Call oLayout.WriteStars(aFile(m_SCRIPT_RANK, 0), false)
					Set oLayout = Nothing
				end if
				.write "</td><td><div class='ItemOwner'>by <a href='kb_user.asp?id="
				.write aFile(m_SCRIPT_USER_ID, 0)
				.write "'>"
				.write aFile(m_SCRIPT_CREATOR, 0)
				.write "</a></div></td><tr><td colspan='2' valign='top'>"
				.write FormatAsHTML(aFile(m_SCRIPT_DESCRIPTION, 0))
				.write "</td></table>"
			Else
				.redirect "kb_scripts.asp"
			End If
		end with
	End Sub
End Class
%>