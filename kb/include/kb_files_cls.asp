<%
Const m_ACT_FILE_ADD = 1
Const m_ACT_FILE_UPDATE = 2
Const m_ACT_FILE_DELETE = 3

'-------------------------------------------------------------------------
'	Name: 		kbFiles class
'	Purpose: 	methods for displaying file information
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'-------------------------------------------------------------------------
Class kbFiles
	Private m_aFilter(4)
	
	'Private Sub Class_Initialize()
	'	Set m_oFileData = New kbFileData
	'End Sub
	
	'Private Sub Class_Terminate()
	'	Set m_oFileData = Nothing
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
	'	Name: 		WritePublicFiles()
	'	Purpose: 	output files as HTML
	'Modifications:
	'	Date:		Name:	Description:
	'	12/23/02	JEA		Creation
	'	1/9/03		JEA		Write categories
	'-------------------------------------------------------------------------
	Public Sub WritePublic()
		Const sPAGE_NAME = "kb_files.asp"
		dim aData
		dim lItemsPerPage
		dim oLayout
		dim lItemCount
		dim lStart
		dim lEnd
		dim oFileData
		dim bBreak
		dim x, y
		
		lItemsPerPage = GetSessionValue(g_USER_ITEMS_PER_PAGE)
		lEnd = (m_aFilter(g_FILTER_PAGE) * lItemsPerPage) - 1
		lStart = lEnd - (lItemsPerPage - 1)

		Set oFileData = New kbFileData
		aData = oFileData.GetPublic(m_aFilter)
		Set oFileData = Nothing
		
		With Response
			Set oLayout = New kbLayout
			.write "<table width='80%' cellspacing='0' cellpadding='0' border='0'>"
			.write "<tr><td colspan='3'>"
			Call oLayout.WriteOptionHead(g_ITEM_FILE, m_aFilter, false)
		
			If IsArray(aData) Then
				lItemCount = UBound(aData, 2)
				If lEnd > lItemCount Then lEnd = lItemCount
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "files", oLayout)
				.write "</td>"
				Call oLayout.WriteItemListHead(sPAGE_NAME, m_aFilter)
				
				For x = lStart to lEnd
					bBreak = false
					aData(m_FILE_RANK,x) = MakeNumber(aData(m_FILE_RANK,x))
					.write "<tr><td class='ItemName' rowspan='3' valign='top'><nobr><a href='kb_download.asp?id="
					.write aData(m_FILE_ID,x)
					.write "'>"
					.write aData(m_FRIENDLY_NAME,x)
					.write "</a>"
					If Trim(aData(m_FILE_RENDERED, x)) <> "" Then
						.write " <a href='"
						If InStr(aData(m_FILE_RENDERED, x), ":") = 0 Then .write "http://"
						.write aData(m_FILE_RENDERED, x)
						.write "'>"
						Call oLayout.WriteToggleImage("icon_stream", "", "Play Rendered File", "width='39' height='11'", false)
						.write "</a>"
					End If
					.write "</nobr>"
					' ranking
					.write "<div class='ItemAction'><a style='color: #777777;' href='kb_rank.asp?id="
					.write aData(m_FILE_ID,x)
					.write "&type="
					.write g_ITEM_FILE
					.write "'>"
					If aData(m_FILE_RANK,x) = 0 Then
						.write "rank this file"
					Else
						Call oLayout.WriteStars(aData(m_FILE_RANK,x), true)
					End If
					.write "</a></div>"
					' edit button
					If g_bAdmin Or CStr(aData(m_FILE_USER_ID,x)) = CStr(GetSessionValue(g_USER_ID)) Then
						.write "<div class='ItemEdit'><a href='kb_file-edit.asp?id="
						.write aData(m_FILE_ID,x)
						.write "&url="
						.write Server.URLEncode(GetURL(false))
						.write "'>"
						Call oLayout.WriteToggleImage("btn_edit", "", "Edit Entry", "", false)
						.write "</a></div>"
					End If
					.write "</td><td align='center' class='ItemDate'>"
					.write FormatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), aData(m_DATE_ADDED,x)))
					.write "</td><td align='center' class='ItemOwner'><a href='kb_user.asp?id="
					.write aData(m_FILE_USER_ID,x)
					.write "'>"
					.write aData(m_FILE_CREATOR,x)
					.write "</a></td><tr><td colspan='2' class='ItemDescription'>"
					.write FormatAsHTML(aData(m_FILE_DESCRIPTION,x))
					if Trim(aData(m_FILE_MEDIA, x)) <> "" then
						.write " (requires <a href='http://"
						.write aData(m_FILE_MEDIA, x)
						.write "'>additional media</a>)"
					end if
					' item notes
					.write "</td><tr><td class='ItemNotes'>"
					Call WritePlugins(aData(m_FILE_PLUGINS,x))
					Call oLayout.WriteCategories(aData(m_FILE_CATS,x))
					.write "</td><td align='right' class='ItemNotes'>"
					If aData(m_FILE_SOFTWARE_NAME,x) <> "" Then
						.write aData(m_FILE_SOFTWARE_NAME,x)
						bBreak = true
					End If
					If aData(m_FILE_FORMAT,x) <> g_FORMAT_NA Then
						if bBreak then .write "<br>"
						.write IIf((aData(m_FILE_FORMAT,x) = g_FORMAT_NTSC), "NTSC", "PAL")
					End If
					.write "</td>"
				Next
				.write "<tr><td colspan='3'><div class='ListBottom'>"
				Call oLayout.WritePaging(lItemsPerPage, lItemCount, m_aFilter, sPAGE_NAME, "files", oLayout)
				.write "</div></td></table>"
			Else
				.write "<br><center>No files match your criteria</center></td></table>"
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
		dim oFileData
		dim bBreak
		dim x

		Set oFileData = New kbFileData
		aData = oFileData.GetPending()
		Set oFileData = Nothing
		
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
					.write "<tr><td class='UploadName'><nobr><a href='kb_download.asp?id="
					.write aData(m_FILE_ID,x)
					.write "'>"
					.write aData(m_FILE_NAME,x)
					.write "</a>"
					If Trim(aData(m_FILE_RENDERED, x)) <> "" Then
						.write " <a href='"
						If InStr(aData(m_FILE_RENDERED, x), ":") = 0 Then .write "http://"
						.write aData(m_FILE_RENDERED, x)
						.write "'>"
						Call oLayout.WriteToggleImage("icon_stream", "", "Play Rendered File", "width='39' height='11'", false)
						.write "</a>"
					End If
					.write "</nobr></td><td align='center' class='UploadDate'>"
					.write FormatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), aData(m_DATE_ADDED, x)))
					.write "</td><td align='center' class='UploadOwner'><a href='kb_user.asp?id="
					.write aData(m_FILE_USER_ID,x)
					.write "'>"
					.write aData(m_FILE_CREATOR,x)
					.write "</a></td><td class='UploadFriendlyName'>"
					.write aData(m_FRIENDLY_NAME,x)
					.write "</td><td class='UploadAction' valign='middle' rowspan='2'><nobr>"
					.write "<a href='kb_admin-uploads.asp?do=approvefile&id="
					.write aData(m_FILE_ID,x)
					.write "'>"
					Call oLayout.WriteToggleImage("btn_approve", "", "Approve", "", false)
					.write "</a> <a href='kb_admin-uploads.asp?do=denyfile&id="
					.write aData(m_FILE_ID,x)
					.write "'>"
					Call oLayout.WriteToggleImage("btn_deny", "", "Approve", "", false)
					.write "</a></nobr></td><tr><td colspan='4' class='UploadDescription'>"
					.write FormatAsHTML(ReplaceNull(aData(m_FILE_DESCRIPTION,x), "(no description given)"))
					if Trim(aData(m_FILE_MEDIA, x)) <> "" then
						.write " (requires <a href='http://"
						.write aData(m_FILE_MEDIA, x)
						.write "'>additional media</a>)"
					end if
					.write "</td>"
				Next
				Set oLayout = Nothing
				.write "</table>"
			Else
				.write "There are no pending files"
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
	Public Sub WriteEditForm(ByVal v_lFileID)
		dim aData
		dim oLayout
		dim oFileData
		dim x
		
		Set oFileData = New kbFileData
		aData = oFileData.GetItem(v_lFileID, true)
		Set oFileData = Nothing
		
		with response
			If IsArray(aData) Then
				Set oLayout = New kbLayout
				Call oLayout.WriteTitleBoxTop("File Edit", "", "")
				.write "<table cellspacing='0' cellpadding='0' border='0'>"
				' file name
				.write "<tr><td class='FormLabel'>File:</td><td class='FormInput'>"
				.write "<input type='file' name='fldFile' value='"
				.write aData(m_FILE_NAME, 0)
				.write "' size='35'></td><tr><td></td><td class='FormNote'>Leave blank"
				.write " except to upload a new version of the file</td>"
				' listed name
				.write "</td><tr><td class='Required'>Listed Name:</td><td class='FormInput'>"
				.write "<input type='text' name='fldFriendlyName' value='"
				.write aData(m_FRIENDLY_NAME, 0)
				.write "' maxlength='30' size='30'></td>"
				' software
				.write "<tr><td class='FormLabel'>Software:</td><td class='FormInput'>"
				Call WriteVersionList("fldVersion", aData(m_FILE_SOFTWARE_ID, 0), g_SOFTWARE_VEGAS)
				.write "</td><tr><td class='FormLabel'>Format:</td><td class='FormInput'>"
				Call WriteFormatList("fldFormat", aData(m_FILE_FORMAT, 0))
				.write "</td><tr><td class='FormLabel'>Rendered URL:</td><td class='FormInput'>"
				' rendered URL
				.write "<input type='text' name='fldRendered' maxlength='75' size='30' value='"
				.write aData(m_FILE_RENDERED, 0)
				.write "'>  link to rendered version</td>"
				' media URL				
				.write "<tr><td class='FormLabel'>Media URL:</td><td class='FormInput'>"
				.write "<input type='text' name='fldMedia' maxlength='75' size='30' value='"
				.write aData(m_FILE_MEDIA, 0)
				.write "'> link to media used</td><tr><td></td><td class='FormNote'><nobr>"
				.write g_sMSG_MEDIA_HINT
				.write "</nobr></td>"
				' description
				.write "<tr><td class='Required' valign='top'>Description:</td>"
				.write "<td class='FormInput' valign='top'><textarea rows='6' cols='52' name='fldDescription'>"
				.write aData(m_FILE_DESCRIPTION, 0)
				.write "</textarea></td><tr><td></td><td class='FormNote'>"
				.write g_sMSG_HTML_LIMIT
				' categories and plugins
				.write "</td><tr><td></td><td class='FormInput' valign='top'>"
				.write "<table cellspacing='0' cellpadding='4' border='0'><tr>"
				.write "<td valign='top' width='50%'>Categories:<br>"
				Call oLayout.WriteCategoryList("fldCategories", aData(m_FILE_CATS, x), 4, g_ITEM_FILE)
				.write "</td><td valign='top' width='50%'>Plugins:<br>"
				Call WritePluginList("fldPlugins", aData(m_FILE_PLUGINS, x), 4)
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
	'	Name: 		WriteVersionList()
	'	Purpose: 	write option list with software versions
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteVersionList(ByVal v_sFieldName, ByVal v_lSelectedID, ByVal v_lSoftwareID)
		dim sQuery
		
		sQuery = "SELECT SV.lVersionID, vsSoftwareName + ' ' + IIf(vsVersionText = '', '(any)', vsVersionText) " _
			& "FROM tblSoftware S INNER JOIN tblSoftwareVersions SV " _
			& "ON S.lSoftwareID = SV.lSoftwareID"
		If v_lSoftwareID <> "" Then sQuery = sQuery & " WHERE S.lSoftwareID = " & v_lSoftwareID
		sQuery = sQuery & " ORDER BY vsVersionText"

		with response
			.write "<select name='"
			.write v_sFieldName
			.write "'>"
			.write MakeList(sQuery, v_lSelectedID)
			.write "</select>"
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
	'-------------------------------------------------------------------------
	Public Sub WriteAsHeader(ByVal v_lFileID)
		dim aFile
		dim oFileData
		dim oLayout
		dim lRank
		dim sComment
		dim x
		
		Set oFileData = New kbFileData
		aFile = oFileData.GetItem(v_lFileID, false)
		Set oFileData = Nothing
		
		with response
			If IsArray(aFile) Then
				' file info
				aFile(m_FILE_RANK, 0) = MakeNumber(aFile(m_FILE_RANK, 0))
				.write "<table cellspacing='0' cellpadding='3' border='0' width='450'>"
				.write "<td valign='top'><div class='ItemName'><nobr>"
				.write aFile(m_FRIENDLY_NAME, 0)
				.write "</nobr></div><div class='ItemOwner'>by <a href='kb_user.asp?id="
				.write aFile(m_FILE_USER_ID, 0)
				.write "'>"
				.write aFile(m_FILE_CREATOR, 0)
				.write "</a></div></td><td rowspan='2' class='ItemText' valign='top'>"
				.write aFile(m_FILE_DESCRIPTION, 0)
				.write "</td><tr><td class='ItemRank' valign='bottom'>"
				if aFile(m_FILE_RANK, 0) > 0 then
					Set oLayout = New kbLayout
					Call oLayout.WriteStars(aFile(m_FILE_RANK, 0), false)
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