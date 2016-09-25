<%
'-------------------------------------------------------------------------
'	Name: 		kbProjects class
'	Purpose: 	methods for displaying file information
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'-------------------------------------------------------------------------
Class kbProjects
	Private m_aFilter(4)
	
	'Private Sub Class_Initialize()
	'	Set m_oProjectData = New kbProjectData
	'End Sub
	
	'Private Sub Class_Terminate()
	'	Set m_oProjectData = Nothing
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
	'	Purpose: 	output files as HTML
	'Modifications:
	'	Date:		Name:	Description:
	'	12/23/02	JEA		Creation
	'	1/9/03		JEA		Write categories
	'	5/1/03		JEA		Show "new" for files newer than last login
	'	5/30/03		JEA		Write product icon
	'-------------------------------------------------------------------------
	Public Sub WritePublic()
		Const sPAGE_NAME = "kb_projects.asp"
		dim aData
		dim lItemsPerPage
		dim oLayout
		dim lItemCount
		dim lStart
		dim lEnd
		dim oProjectData
		dim sLastLogin
		dim bBreak
		dim x, y
		
		lItemsPerPage = GetSessionValue(g_USER_ITEMS_PER_PAGE)
		lEnd = (m_aFilter(g_FILTER_PAGE) * lItemsPerPage) - 1
		lStart = lEnd - (lItemsPerPage - 1)
		sLastLogin = GetSessionValue(g_USER_LAST_LOGIN)

		Set oProjectData = New kbProjectData
		aData = oProjectData.GetPublic(m_aFilter)
		Set oProjectData = Nothing
		
		With Response
			Set oLayout = New kbLayout
			.write "<table width='80%' cellspacing='0' cellpadding='0' border='0'>"
			.write "<tr><td colspan='3'>"
			Call oLayout.WriteOptionHead(g_ITEM_PROJECT, m_aFilter, false)
		
			If IsArray(aData) Then
				lItemCount = UBound(aData, 2)
				If lEnd > lItemCount Then lEnd = lItemCount
				.write "</td>"
				Call oLayout.WriteItemListHead(sPAGE_NAME, m_aFilter)
				.Write "<tr><td colspan='3' class='PagingRow'>"
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "projects", oLayout)
				.Write "</td>"
				
				For x = lStart to lEnd
					bBreak = false
					aData(m_PROJECT_RANK,x) = MakeNumber(aData(m_PROJECT_RANK,x))
					.write "<tr><td class='ItemName' rowspan='3' valign='top'><nobr><a href='kb_project-download.asp?id="
					.write aData(m_PROJECT_ID,x)
					.write "'>"
					.write aData(m_FRIENDLY_NAME,x)
					.write "</a>"
					If Trim(aData(m_PROJECT_RENDERED, x)) <> "" Then
						.write " <a href='"
						If InStr(aData(m_PROJECT_RENDERED, x), ":") = 0 Then .write "http://"
						.write aData(m_PROJECT_RENDERED, x)
						.write "'>"
						Call oLayout.WriteToggleImage("icon_stream", "", "Play Rendered File", "width='39' height='11'", false)
						.write "</a>"
					End If
					.write "</nobr>"
					' downloads
					.write "<div class='ItemDownloads'>"
					.write aData(m_PROJECT_DOWNLOADS,x)
					.write " downloads</div>"
					' ranking
					.write "<div class='ItemAction'><a style='color: #777777;' href='kb_rank.asp?id="
					.write aData(m_PROJECT_ID,x)
					.write "&type="
					.write g_ITEM_PROJECT
					.write "'>"
					If aData(m_PROJECT_RANK,x) = 0 Then
						.write "rank this project"
					Else
						Call oLayout.WriteStars(aData(m_PROJECT_RANK,x), true)
					End If
					.write "</a></div>"
					' edit button
					If g_bAdmin Or CStr(aData(m_PROJECT_USER_ID,x)) = CStr(GetSessionValue(g_USER_ID)) Then
						.write "<div class='ItemEdit'><a href='kb_project-edit.asp?id="
						.write aData(m_PROJECT_ID,x)
						.write "&url="
						.write Server.URLEncode(GetURL(false))
						.write "'>"
						Call oLayout.WriteToggleImage("btn_edit", "", "Edit Entry", "", false)
						.write "</a></div>"
					End If
					' item date
					.write "</td><td align='center' class='ItemDate'>"
					.write FormatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), aData(m_DATE_ADDED,x)))
					If aData(m_PROJECT_VERSION,x) > 1 And DateDiff("d", aData(m_DATE_ADDED,x), Date()) <= g_SHOW_AS_NEW_DAYS Then
						.write "<img src='./images/"
						.write g_lSiteID
						.write "/new-version.gif' width='61' height='11'>"
					ElseIf DateDiff("d", aData(m_DATE_ADDED,x), sLastLogin) < 0 Then
						.write "<img src='./images/"
						.write g_lSiteID
						.write "/new.gif' width='29' height='11'>"
					End If
					.write "</td><td align='center' class='ItemOwner'><nobr><a href='kb_user.asp?id="
					.write aData(m_PROJECT_USER_ID,x)
					.write "'>"
					.write aData(m_PROJECT_CREATOR,x)
					.write "</a></nobr></td><tr><td colspan='2' class='ItemDescription'>"
					.write FormatAsHTML(aData(m_PROJECT_DESCRIPTION,x))
					if Trim(aData(m_PROJECT_MEDIA, x)) <> "" then
						.write " (requires <a href='http://"
						.write aData(m_PROJECT_MEDIA, x)
						.write "'>additional media</a>)"
					end if
					' item notes
					.write "</td><tr><td class='ItemNotes'>"
					Call WritePlugins(aData(m_PROJECT_PLUGINS,x))
					Call oLayout.WriteCategories(aData(m_PROJECT_CATS,x))
					.write "&nbsp;</td><td align='right' class='ItemVersion'>"
					' version and icon
					.write "<table cellspacing='0' cellpadding='0' border='0'><tr>"
					.write "<td align='right' style='font-size: 8pt;'>"
					If aData(m_PROJECT_SOFTWARE_NAME,x) <> "" Then
						.Write "<nobr>"
						.write Replace(aData(m_PROJECT_SOFTWARE_NAME,x), " (any)", "")
						.Write "</nobr>"
						bBreak = true
					End If
					If aData(m_PROJECT_FORMAT,x) <> g_FORMAT_NA Then
						if bBreak then .write "<br>"
						.write IIf((aData(m_PROJECT_FORMAT,x) = g_FORMAT_NTSC), "NTSC", "PAL")
					End If
					' product icon
					.write "</td><td width='21' height='18' align='right' style='margin-left: 3px;'>"
					.write "<img src='./images/icon_"
					.write aData(m_PROJECT_ICON,x)
					.write ".gif' width='16' height='16'></td></table></td>"
				Next
				.write "<tr><td colspan='3'><div class='ListBottom'>"
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "projects", oLayout)
				.write "</div></td></table>"
			Else
				.write "<br><center>No projects match your criteria</center></td></table>"
			End If
			Set oLayout = Nothing
		End With
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WritePlugins()
	'	Purpose: 	write plugins, if any
	'Modifications:
	'	Date:		Name:	Description:
	'	12/30/02	JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub WritePlugins(ByVal v_aPlugins)
		dim sPluginPackage
		dim x
		
		If IsArray(v_aPlugins) Then
			sPluginPackage = ""
			with response
				.write "<nobr>extra plugins: "
				for x = 0 to UBound(v_aPlugins, 2)
					if sPluginPackage <> v_aPlugins(g_PLUGIN_PACKAGE,x) then
						sPluginPackage = v_aPlugins(g_PLUGIN_PACKAGE,x)
						if x > 0 then .write "</a>; "
						.write "<a href='http://"
						.write v_aPlugins(g_PLUGIN_URL,x)
						.write "'><b>"
						.write v_aPlugins(g_PLUGIN_PACKAGE,x)
						.write "</b> "
					else
						.write ", "
					end if
					.write v_aPlugins(g_PLUGIN_NAME,x)
				next
				.write "</a></nobr>"
			end with
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WritePending()
	'	Purpose: 	output files as HTML
	'Modifications:
	'	Date:		Name:	Description:
	'	12/23/02	JEA		Creation
	'	2/22/03		JEA		Also write linked media
	'-------------------------------------------------------------------------
	Public Sub WritePending()
		dim aData
		dim oLayout
		dim oProjectData
		dim bBreak
		dim x

		Set oProjectData = New kbProjectData
		aData = oProjectData.GetPending()
		Set oProjectData = Nothing
		
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
					.write "<tr><td class='UploadName'><nobr><a href='kb_project-download.asp?id="
					.write aData(m_PROJECT_ID,x)
					.write "'>"
					.write aData(m_PROJECT_NAME,x)
					.write "</a>"
					If Trim(aData(m_PROJECT_RENDERED, x)) <> "" Then
						.write " <a href='"
						If InStr(aData(m_PROJECT_RENDERED, x), ":") = 0 Then .write "http://"
						.write aData(m_PROJECT_RENDERED, x)
						.write "'>"
						Call oLayout.WriteToggleImage("icon_stream", "", "Play Rendered File", "width='39' height='11'", false)
						.write "</a>"
					End If
					.write "</nobr></td><td align='center' class='UploadDate'>"
					.write FormatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), aData(m_DATE_ADDED, x)))
					.write "</td><td align='center' class='UploadOwner'><a href='kb_user.asp?id="
					.write aData(m_PROJECT_USER_ID,x)
					.write "'>"
					.write aData(m_PROJECT_CREATOR,x)
					.write "</a></td><td class='UploadFriendlyName'>"
					.write aData(m_FRIENDLY_NAME,x)
					.write "</td><td class='UploadAction' valign='middle' rowspan='2'><nobr>"
					.write "<a href='kb_admin-uploads.asp?do=approveproject&id="
					.write aData(m_PROJECT_ID,x)
					.write "'>"
					Call oLayout.WriteToggleImage("btn_approve", "", "Approve", "", false)
					.write "</a> <a href='kb_admin-uploads.asp?do=denyproject&id="
					.write aData(m_PROJECT_ID,x)
					.write "'>"
					Call oLayout.WriteToggleImage("btn_deny", "", "Approve", "", false)
					.write "</a></nobr></td><tr><td colspan='4' class='UploadDescription'>"
					.write FormatAsHTML(ReplaceNull(aData(m_PROJECT_DESCRIPTION,x), "(no description given)"))
					if Trim(aData(m_PROJECT_MEDIA, x)) <> "" then
						.write " (requires <a href='http://"
						.write aData(m_PROJECT_MEDIA, x)
						.write "'>additional media</a>)"
					end if
					.write "</td>"
				Next
				Set oLayout = Nothing
				.write "</table>"
			Else
				.write "There are no pending projects"
			End If
		End With
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteEditForm()
	'	Purpose: 	write box for editing single file
	'Modifications:
	'	Date:		Name:	Description:
	'	12/31/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteEditForm(ByVal v_lProjectID)
		dim aData
		dim oLayout
		dim oProjectData
		dim x
		
		Set oProjectData = New kbProjectData
		aData = oProjectData.GetItem(v_lProjectID, true)
		Set oProjectData = Nothing
		
		with response
			If IsArray(aData) Then
				Set oLayout = New kbLayout
				Call oLayout.WriteTitleBoxTop("Project Edit", "", "")
				.write "<table cellspacing='0' cellpadding='0' border='0'>"
				' file name
				.write "<tr><td class='FormLabel'>File:</td><td class='FormInput'>"
				.write "<input type='file' name='fldFile' value='"
				.write aData(m_PROJECT_NAME, 0)
				.write "' size='35'></td><tr><td></td><td class='FormNote'>Leave blank"
				.write " except to upload a new version of the file</td>"
				' listed name
				.write "</td><tr><td class='Required'>Listed Name:</td><td class='FormInput'>"
				.write "<input type='text' name='fldFriendlyName' value='"
				.write aData(m_FRIENDLY_NAME, 0)
				.write "' maxlength='30' size='30'></td>"
				' software
				.write "<tr><td class='Required'>Requires:</td><td class='FormInput'>"
				Call oLayout.WriteVersionList("fldVersion", aData(m_PROJECT_SOFTWARE_ID, 0), g_ITEM_PROJECT)
				.Write " software required to open the project"
				.write "</td><tr><td class='FormLabel'>Format:</td><td class='FormInput'>"
				Call WriteFormatList("fldFormat", aData(m_PROJECT_FORMAT, 0))
				.write "</td><tr><td class='FormLabel'>Rendered URL:</td><td class='FormInput'>"
				' rendered URL
				.write "<input type='text' name='fldRendered' maxlength='75' size='30' value='"
				.write aData(m_PROJECT_RENDERED, 0)
				.write "'>  link to rendered version</td>"
				' media URL				
				.write "<tr><td class='FormLabel'>Media URL:</td><td class='FormInput'>"
				.write "<input type='text' name='fldMedia' maxlength='75' size='30' value='"
				.write aData(m_PROJECT_MEDIA, 0)
				.write "'> link to media used</td><tr><td></td><td class='FormNote'><nobr>"
				.write g_sMSG_MEDIA_HINT
				.write "</nobr></td>"
				' description
				.write "<tr><td class='Required' valign='top'>Description:</td>"
				.write "<td class='FormInput' valign='top'><textarea rows='6' cols='52' name='fldDescription'>"
				.write aData(m_PROJECT_DESCRIPTION, 0)
				.write "</textarea></td><tr><td></td><td class='FormNote'>"
				.write g_sMSG_HTML_LIMIT
				' categories and plugins
				.write "</td><tr><td></td><td class='FormInput' valign='top'>"
				.write "<table cellspacing='0' cellpadding='4' border='0'><tr>"
				.write "<td valign='top' width='50%'>Categories:<br>"
				Call oLayout.WriteCategoryList("fldCategories", aData(m_PROJECT_CATS, x), 4, g_ITEM_PROJECT)
				.write "</td><td valign='top' width='50%'>Plugins:<br>"
				Call WritePluginList("fldPlugins", aData(m_PROJECT_PLUGINS, x), 4)
				.write "</td><tr><td colspan='2' class='FormNote' align='center'>"
				.write g_sMSG_MULTI_SELECT
				.write "</td></table></td>"
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
				.write "File not found or not permissable"
			End If
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteFormatList()
	'	Purpose: 	write option list with video formats
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteFormatList(ByVal v_sFieldName, ByVal v_lSelectedID)
		dim sQuery
		sQuery = "SELECT lFormatID, vsFormatDescription FROM tblFormats ORDER BY vsFormatDescription"
		with response
			.write "<select name='"
			.write v_sFieldName
			.write "'>"
			.write MakeList(sQuery, v_lSelectedID)
			.write "</select>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WritePluginList()
	'	Purpose: 	write option list with plugins
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'	4/25/03		JEA		Add site condition
	'-------------------------------------------------------------------------
	Public Sub WritePluginList(ByVal v_sFieldName, ByVal v_aSelected, ByVal v_lSize)
		dim sQuery
		dim x
	
		sQuery = "SELECT lPluginID, vsPluginPackageName + ' ' + vsPluginName " _
			& "FROM (tblPluginPackages PP INNER JOIN tblPlugins P " _
			& "ON PP.lPluginPackageID = P.lPluginPackageID) " _
			& "INNER JOIN (SELECT lItemID FROM tblItemSites WHERE lItemTypeID = " _
			& g_ITEM_PLUGIN & " AND lSiteID = " & GetSessionValue(g_USER_SITE) _
			& ") tIS ON tIS.lItemID = PP.lPluginPackageID "
		if g_bFREE_PLUGINS_ONLY then sQuery = sQuery & "WHERE PP.bFree = true "
		sQuery = sQuery & "ORDER BY vsPluginPackageName, vsPluginName"
		
		v_lSize = MakeNumber(v_lSize)
		
		with response
			.write "<select align='top' name='"
			.write v_sFieldName
			.write "'"
			if v_lSize > 0 then
				.write " size='"
				.write v_lSize
				.write "' multiple"
			end if
			.write ">"
			.write MakeList(sQuery, "")
			.write "</select><br><a style='font-size: 8pt;' href=""JavaScript:ClearSelection('"
			.write v_sFieldName
			.write "');"">clear selection</a>"
			.write "<input type='hidden' name='fldPluginList' value='"
			If IsArray(v_aSelected) Then
				for x = 0 to UBound(v_aSelected, 2)
					.write v_aSelected(g_CAT_ID, x)
					if x < UBound(v_aSelected, 2) then .write ","
				next
			End If
			.write "'>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteAsHeader()
	'	Purpose: 	write summary for given file
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'	4/29/03		JEA		Adjust layout for longer descriptions
	'-------------------------------------------------------------------------
	Public Sub WriteAsHeader(ByVal v_lProjectID)
		dim aFile
		dim oProjectData
		dim oLayout
		dim lRank
		dim sComment
		dim x
		
		Set oProjectData = New kbProjectData
		aFile = oProjectData.GetItem(v_lProjectID, false)
		Set oProjectData = Nothing
		
		with response
			If IsArray(aFile) Then
				' file info
				aFile(m_PROJECT_RANK, 0) = MakeNumber(aFile(m_PROJECT_RANK, 0))
				.write "<table cellspacing='0' cellpadding='3' border='0' width='450'>"
				.write "<td valign='top' colspan='2'><div class='ItemName'><nobr>"
				.write aFile(m_FRIENDLY_NAME, 0)
				.write "</nobr></div></td><tr><td class='ItemRank'>"
				if aFile(m_PROJECT_RANK, 0) > 0 then
					Set oLayout = New kbLayout
					Call oLayout.WriteStars(aFile(m_PROJECT_RANK, 0), false)
					Set oLayout = Nothing
				end if
				.write "</td><td><div class='ItemOwner'>by <a href='kb_user.asp?id="
				.write aFile(m_PROJECT_USER_ID, 0)
				.write "'>"
				.write aFile(m_PROJECT_CREATOR, 0)
				.write "</a></div></td><tr><td colspan='2' valign='top'>"
				.write FormatAsHTML(aFile(m_PROJECT_DESCRIPTION, 0))
				.write "</td></table>"
			Else
				.redirect "kb_projects.asp"
			End If
		end with
	End Sub
End Class
%>