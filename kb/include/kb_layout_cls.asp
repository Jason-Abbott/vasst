<%
Const m_sMENU_COMMON = 0
Const m_sMENU_ADMIN = 1
Const m_CORNER_TL = 1
Const m_CORNER_TR = 2
Const m_CORNER_BR = 3
Const m_CORNER_BL = 4

'-------------------------------------------------------------------------
'	Name: 		kbLayout class
'	Purpose: 	encapsulate functions for page layout
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'-------------------------------------------------------------------------
Class kbLayout

	'Private Sub Class_Initialize()
	'End Sub
	
	'Private Sub Class_Terminate()
	'End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetCommonMenuArray()
	'	Purpose: 	get values for common menu
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	12/30/02	JEA		Creation
	'	4/26/03		JEA		Send Site ID for logout
	'-------------------------------------------------------------------------
	Private Function GetCommonMenuArray()
		GetCommonMenuArray = Array( _
			Array("kb_files.asp","Files", "_files", false), _
			Array("kb_tutorials.asp","Tutorials", "_tutorials", false), _
			Array("kb_forums.asp","Forums", "_forums", false), _
			Array("kb_user-edit.asp","My Account", "_user-edit", false), _
			Array("kb_login.asp?s=" & GetSessionValue(g_USER_SITE) & "&logout=yes","Sign out", "_login", false), _
			Array("kb_admin-uploads.asp","Administration", "_admin", true))
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetAdminMenuArray()
	'	Purpose: 	get values for admin menu
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	12/30/02	JEA		Creation
	'	1/22/03		JEA		Add cache menu
	'-------------------------------------------------------------------------
	Private Function GetAdminMenuArray()
		GetAdminMenuArray = Array( _
			Array("kb_admin-uploads.asp","Submissions", "_admin-uploads", true), _
			Array("kb_admin-users.asp","Users", "_admin-users", true), _
			Array("kb_admin-categories.asp","Categories", "_admin-categories", true), _
			Array("kb_admin-activity.asp","Activity Log", "_admin-activity", true), _
			Array("kb_admin-contests.asp","Contests", "_admin-contests", true), _
			Array("kb_admin-cache.asp","Cache", "_admin-cache", true), _
			Array("kb_admin-database.asp","Database", "_admin-database", true))
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteMenuBar()
	'	Purpose: 	write common menu bar
	'Modifications:
	'	Date:		Name:	Description:
	'	12/29/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteMenuBar(ByVal v_lMenuID)
		Const LINK_URL = 0
		Const LINK_NAME = 1
		Const LINK_ID = 2
		Const LINK_RESTRICT = 3		' only allow admin access
		dim sSeparator
		dim sPage
		dim sStyle
		dim bOnPage
		dim aLinks
		dim sClass
		dim x
		
		sSeparator = "<span class='Separator'>&#119;</span>"
		sPage = Request.ServerVariables("SCRIPT_NAME")
		sPage = Right(sPage, Len(sPage) - InStrRev(sPage, "/"))
		
		Select Case MakeNumber(v_lMenuID)
			Case m_sMENU_COMMON
				aLinks = GetCommonMenuArray()
				sClass = "CommonMenuBar"
			Case m_sMENU_ADMIN
				aLinks = GetAdminMenuArray()
				sClass = "AdminMenuBar"
			Case Else
				Exit Sub
		End Select
		
		with response
			if v_lMenuID = m_sMENU_COMMON then
				.write "<div class='AppName'>"
				.write g_sORG_NAME
				.write " "
				.write g_sAPP_NAME
				.write "</div>"
			end if
			.write "<div class='"
			.write sClass
			.write "'>"
			for x = 0 to UBound(aLinks)
				If g_bAdmin Or aLinks(x)(LINK_RESTRICT) = false Then
					bOnPage = CBool(InStr(sPage, aLinks(x)(LINK_ID)) > 0)
					If Not bOnPage Then
						.write "<a href='"
						.write aLinks(x)(LINK_URL)
						.write "'>"
					End If
					.write aLinks(x)(LINK_NAME)
					If Not bOnPage Then .write "</a>"
					If x < UBound(aLinks) Then .write sSeparator
				End If
			next
			.write "</div>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteTabEdge()
	'	Purpose: 	write tab border
	'Modifications:
	'	Date:		Name:	Description:
	'	1/7/03		JEA		Creation
	'	4/26/03		JEA		Add site customization
	'-------------------------------------------------------------------------
	Public Sub WriteTabEdge(ByVal v_lCornerID, ByVal v_sOptions, ByVal v_sCornerImg)
		if v_sCornerImg = "" then v_sCornerImg = "box-on-page"
		with response
			select case v_lCornerID
				case m_CORNER_TL
					.write "<table cellspacing='0' cellpadding='0' border='0' "
					.write v_sOptions
					.write "><td class='BoxCorner'><img src='./images/"
					.write g_lSiteID
					.write "/corner_tl-"
					.write v_sCornerImg
					.write ".gif'><td class='BoxTop' rowspan='2' align='center'>"
				case m_CORNER_TR
					.write "</td><td class='BoxCorner'><img src='./images/"
					.write g_lSiteID
					.write "/corner_tr-"
					.write v_sCornerImg
					.write ".gif'><tr><td class='BoxLeft'>&nbsp;</td><td class='BoxRight'>"
					.write "&nbsp;</td></table>"
				case m_CORNER_BR
				case m_CORNER_BL
			end select
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteTitleBoxTop()
	'	Purpose: 	write title box
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	3/23/03		JEA		Prevent title break
	'	4/26/03		JEA		Customize by site
	'-------------------------------------------------------------------------
	Public sub WriteTitleBoxTop(ByVal v_sTitle, ByVal v_sOptions, ByVal v_sCornerImg)
		if v_sCornerImg = "" then v_sCornerImg = "tall-title-on-page"
		with response
			.write "<table cellspacing='0' cellpadding='0' border='0'"
			If v_sOptions <> "" Then .write " "
			.write v_sOptions
			.write "><tr><td class='BoxTitleCorner'><img src='./images/"
			.write g_lSiteID
			.write "/corner_tl-"
			.write v_sCornerImg
			.write ".gif' width='10' height='20'></td><td class='BoxTitleTop'><nobr>"
			.write v_sTitle
			.write "</nobr></td><td class='BoxTitleCorner'><img src='./images/"
			.write g_lSiteID
			.write "/corner_tr-"
			.write v_sCornerImg
			.write ".gif' width='10' height='20'></td>"
			.write "<tr><td class='BoxLeft'><img src='./images/blank.gif' width='1' height='1'></td>"
			.write "<td class='BoxBody'>"
		end with
	end sub
	
	Public Sub WriteBoxBottom(ByVal v_sCornerImg)
		If v_sCornerImg = "" Then v_sCornerImg = "box-on-page"
		With response
			.write "</td><td class='BoxRight'>"
			.write "<img src='./images/blank.gif' width='1' height='1'></td>"
			.write "<tr><td class='BoxCorner'><img src='./images/"
			.write g_lSiteID
			.write "/corner_bl-"
			.write v_sCornerImg
			.write ".gif' width='10' height='10'></td><td class='BoxBottom'>"
			.write "<img src='./images/blank.gif' width='1' height='1'></td>"
			.write "<td class='BoxCorner'><img src='./images/"
			.write g_lSiteID
			.write "/corner_br-"
			.write v_sCornerImg
			.write ".gif' width='10' height='10'></td></table>"
		End With
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteToggleImage()
	'	Purpose: 	toggle image
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	4/26/03		JEA		Customize by site
	'-------------------------------------------------------------------------
	Public Sub WriteToggleImage(ByVal v_sBaseImage, ByVal v_sPath, ByVal v_sAltText, _
		ByVal v_sOptions, ByVal v_bInput)
		
		If v_sPath = "" Then v_sPath = "./images/" & g_lSiteID & "/"
		With Response
			.write IIf(v_bInput, "<input type='image' ", "<img ")
			.write "src='"
			.write v_sPath
			.write v_sBaseImage
			.write ".gif' onMouseOver=""this.src='"
			.write v_sPath
			.write v_sBaseImage
			.write "_on.gif';"" onMouseOut=""this.src='"
			.write v_sPath
			.write v_sBaseImage
			.write ".gif';"""
			if v_sAltText <> "" then
				.write " alt='"
				.write v_sAltText
				.write "'"
			end if
			if v_sOptions <> "" then
				.write " "
				.write v_sOptions
			end if
			.write " border='0'>"
		End With
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteBoxTop()
	'	Purpose: 	write HTML for rounded box
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	4/26/03		JEA		Customize by site
	'-------------------------------------------------------------------------
	Public Sub WriteBoxTop(ByVal v_sOptions, ByVal v_sCornerImg)
		If v_sCornerImg = "" Then v_sCornerImg = "box-on-page"
		With Response
			.write "<table cellspacing='0' cellpadding='0' border='0'"
			If v_sOptions <> "" Then .write " "
			.write v_sOptions
			.write "><tr><td class='BoxCorner'><img src='./images/"
			.write g_lSiteID
			.write "/corner_tl-"
			.write v_sCornerImg
			.write ".gif' width='10' height='10'></td><td class='BoxTop'>"
			.write "<img src='./images/blank.gif' width='1' height='1'></td>"
			.write "</td><td class='BoxCorner'><img src='./images/"
			.write g_lSiteID
			.write "/corner_tr-"
			.write v_sCornerImg
			.write ".gif' width='10' height='10'></td><tr><td class='BoxLeft'>"
			.write "<img src='./images/blank.gif' width='1' height='1'></td>"
			.write "<td class='BoxBody'>"
		End With
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		WritePaging()
	'	Purpose: 	write links for paging through items (item count is 0-based)
	'Modifications:
	'	Date:		Name:	Description:
	'	12/30/02	JEA		Creation
	'	1/1/03		JEA		Track sorting
	'	1/9/03		JEA		Don't write for single page
	'-------------------------------------------------------------------------
	Public Sub WritePaging(ByVal v_lItemsPerPage, ByVal v_lItemCount, ByVal v_aFilter, _
		ByVal v_sPageName, ByVal v_sItemName, ByRef r_oLayout)
		
		dim lPageCount
		dim sQS
		dim x
		
		lPageCount = (v_lItemCount + 1) /v_lItemsPerPage
		If Int(lPageCount) < lPageCount Then lPageCount = Int(lPageCount) + 1
		If lPageCount > 1 Then
			sQS = "&sort=" & v_aFilter(g_FILTER_SORT) & "&cat=" & v_aFilter(g_FILTER_CATEGORY) _
				& "&sw=" & v_aFilter(g_FILTER_SOFTWARE) & "&author=" & v_aFilter(g_FILTER_AUTHOR)
			with response
				.write "<table width='100%' cellspacing='0' cellpadding='0'><tr><td class='Paging' width='20%'>"
				if v_aFilter(g_FILTER_PAGE) = 1 then
					.write "<nobr>"
					.write v_lItemCount + 1
					.write " "
					.write v_sItemName
					.write " on "
					.write lPageCount
					.write " page"
					if lPageCount > 1 then .write "s:"
					.write "</nobr>"
				else
					.write "<a href='"
					.write v_sPageName
					.write "?page="
					.write v_aFilter(g_FILTER_PAGE) - 1
					.write sQS
					.write "'>"
					Call r_oLayout.WriteToggleImage("btn_back", "", "Previous Page", "width='42' height='14'", false)
					.write "</a>"
				end if
				.write "</td><td align='center'>"
				for x = 1 to lPageCount
					if x = v_aFilter(g_FILTER_PAGE) then
						.write "<img src='./images/"
						.write g_lSiteID
						.write "/btn_"
						.write x
						.write "_on.gif' width='14' height='14'> "
					else
						.write "<a href='"
						.write v_sPageName
						.write "?page="
						.write x
						.write sQS
						.write "'>"
						Call r_oLayout.WriteToggleImage("btn_" & x, "", "go to page " & x, "width='14' height='14'", false)
						.write "</a> "
					end if
				next
				.write "</td><td align='right' width='20%' class='Paging'>"
				if v_aFilter(g_FILTER_PAGE) < lPageCount then
					.write "<a href='"
					.write v_sPageName
					.write "?page="
					.write v_aFilter(g_FILTER_PAGE) + 1
					.write sQS
					.write "'>"
					Call r_oLayout.WriteToggleImage("btn_next", "", "Next Page", "width='42' height='14'", false)
					.write "</a>"
				else
					.write "&nbsp;"
				end if
				.write "</td></table>"
			end with
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetAsFraction()
	'	Purpose: 	write number as fraction
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function GetAsFraction(ByVal v_lNumber)
		dim lFraction
		If IsNumber(v_lNumber) Then
			lFraction = v_lNumber - Int(v_lNumber)
			v_lNumber = Int(v_lNumber)
			select case lFraction
				case .25
					GetAsFraction = v_lNumber & "&frac14;"
				case .5
					GetAsFraction = v_lNumber & "&frac12;"
				case .75
					GetAsFraction = v_lNumber & "&frac34;"
				case else
					GetAsFraction = v_lNumber + lFraction
			end select
		Else
			GetAsFraction = v_lNumber
		End If
		
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteStars()
	'	Purpose: 	write rank display for given value
	'Modifications:
	'	Date:		Name:	Description:
	'	1/4/03		JEA		Creation
	'	4/26/03		JEA		Customize by site
	'-------------------------------------------------------------------------
	Sub WriteStars(ByVal v_lRank, ByVal v_bClickable)
		dim x
		dim sImageName
		with response
			sImageName = "btn_" & v_lRank & "-star"
			if v_bClickable then
				Call WriteToggleImage(sImageName, "", "rank or view comments", "", false)
			else
				.write "<img src='./images/"
				.write g_lSiteID
				.write "/"
				.write sImageName
				.write ".gif' width='62' height='12'>"
			end if
			'.write " "
			'.write GetAsFraction(v_lRank)
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteItemListHead()
	'	Purpose: 	write heading for item list in three column table
	'Modifications:
	'	Date:		Name:	Description:
	'	1/5/03		JEA		Creation
	'	4/26/03		JEA		Customize arrow by site
	'-------------------------------------------------------------------------
	Public Sub WriteItemListHead(ByVal v_sPageURL, ByVal v_aFilter)
		Const NAME = 0
		Const SORT_ASC = 1
		Const SORT_DESC = 2
		dim sArrow
		dim aHeadings
		dim x
		
		aHeadings = Array( _
			Array("Name", g_SORT_NAME_ASC, g_SORT_NAME_DESC), _
			Array("Date Added", g_SORT_DATE_ASC, g_SORT_DATE_DESC), _
			Array("From", g_SORT_OWNER_ASC, g_SORT_OWNER_DESC))
		sArrow = "<img class='SortArrow' src='./images/" & g_lSiteID & "/arrow_" _
			& IIf((v_aFilter(g_FILTER_SORT) Mod 2), "down", "up") & ".gif'>"
		with response
			.write "<tr>"
			for x = 0 to UBound(aHeadings)
				.write "<td class='ItemHead' width='33%'><a class='ItemHead' href='"
				.write v_sPageURL
				.write "?sort="
				.write IIf((v_aFilter(g_FILTER_SORT) = aHeadings(x)(SORT_ASC)), aHeadings(x)(SORT_DESC), aHeadings(x)(SORT_ASC))
				.write "&page=1&cat="
				.write v_aFilter(g_FILTER_CATEGORY)
				.write "&sw=" & v_aFilter(g_FILTER_SOFTWARE)
				.write "&author=" & v_aFilter(g_FILTER_AUTHOR)
				.write "'>"
				.write aHeadings(x)(NAME)
				.write "</a>"
				if MatchesOne(v_aFilter(g_FILTER_SORT), Array(aHeadings(x)(SORT_DESC), aHeadings(x)(SORT_ASC)), true) then .write sArrow
				.write "</td>"
			next
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteOptionHead()
	'	Purpose: 	write heading for item list inside of table row
	'Modifications:
	'	Date:		Name:	Description:
	'	1/7/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteOptionHead(ByVal v_lItemTypeID, ByVal v_aFilter, ByVal v_bRestricted)
		
		dim sCatList
		dim sSoftwareList
		dim sAuthorList
		dim sItemType
		dim sParms
		dim sJSArray
		dim sURL
		dim sAuthor
		Select Case v_lItemTypeID
			Case g_ITEM_FILE : sItemType = "File" : sURL = "submit-file" : sAuthor = "User"
			Case g_ITEM_TUTORIAL : sItemType = "Tutorial" : sURL = "tutorial-edit" : sAuthor = "Author"
			Case g_ITEM_FORUM : sItemType = "Forum" : sURL = "forum-edit" : sAuthor = "User"
		End Select
		sCatList = GetCategoryList(v_lItemTypeID)
		sSoftwareList = GetSoftwareList(v_lItemTypeID)
		sJSArray = "[" & Join(v_aFilter, ",") & "]"
		
		with response
			.write "<div class='OptionHead'>"
			Call WriteTabEdge(m_CORNER_TL, "", "")
			.write "<nobr>"
			If v_lItemTypeID <> g_ITEM_FORUM Then
				sAuthorList = GetAuthorList(v_lItemTypeID, sItemType, sAuthor)
				.write "<select name='fldAuthor' class='TopOption' "
				.write "onChange=""newFilter(this, '"
				.write LCase(sItemType)
				.write "s', "
				.write g_FILTER_AUTHOR
				.write ", "
				.write sJSArray
				.write ");""><option value=''>Any Author"
				.write MakeSelected(sAuthorList, v_aFilter(g_FILTER_AUTHOR))
				.write "</select> "
			End If
			.write "<select name='fldCategory' class='TopOption' "
			.write "onChange=""newFilter(this, '"
			.write LCase(sItemType)
			.write "s', "
			.write g_FILTER_CATEGORY
			.write ", "
			.write sJSArray
			.write ");""><option value=''>Any Category"
			.write MakeSelected(sCatList, v_aFilter(g_FILTER_CATEGORY))
			.write "</select> <select name='fldSoftware' class='TopOption' "
			.write "onChange=""newFilter(this, '"
			.write LCase(sItemType)
			.write "s', "
			.write g_FILTER_SOFTWARE
			.write ", "
			.write sJSArray
			.write ");""><option value=''>Any Software"
			.write MakeSelected(sSoftwareList, v_aFilter(g_FILTER_SOFTWARE))
			.write "</select>"
			If (Not v_bRestricted) Or g_bAdmin Then
				.write " &nbsp; <a href='kb_"
				.write sURL
				.write ".asp'>"
				Call WriteToggleImage("btn_add-" & LCase(sItemType), "", "Add " & sItemType, "", false)
				.write "</a>"
			End If
			.write "</nobr>"
			Call WriteTabEdge(m_CORNER_TR, "", "")
			.write "</div>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetCategoryList()
	'	Purpose: 	get option list of categories for given item type
	'Modifications:
	'	Date:		Name:	Description:
	'	1/12/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function GetCategoryList(ByVal v_lItemTypeID)
		dim sQuery
		dim sList
		dim sKey
		sKey = v_lItemTypeID & "_" & g_FILTER_CATEGORY
		If Application(sKey) = "" Then
			sQuery = "SELECT C.lCategoryID, C.vsCategoryName FROM tblCategories C " _
				& "INNER JOIN tblCategoryItemTypes CIT ON CIT.lCategoryID = C.lCategoryID " _
				& "WHERE CIT.lItemTypeID = " & v_lItemTypeID & " ORDER BY vsCategoryName"
			sList = MakeList(sQuery, "")
			Application.Lock : Application(sKey) = sList : Application.Unlock
		End If
		GetCategoryList = Application(sKey)
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetSoftwareList()
	'	Purpose: 	get option list of software
	'Modifications:
	'	Date:		Name:	Description:
	'	1/12/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function GetSoftwareList(ByVal v_lItemTypeID)
		dim sQuery
		dim sList
		dim sKey
		sKey = v_lItemTypeID & "_" & g_FILTER_SOFTWARE
		If Application(sKey) = "" Then
			sQuery = "SELECT SV.lVersionID, S.vsSoftwareName + ' ' + SV.vsVersionText " _
				& "FROM tblSoftwareVersions SV INNER JOIN tblSoftware S " _
				& "ON SV.lSoftwareID = S.lSoftwareID WHERE SV.lSoftwareID = " & g_SOFTWARE_VEGAS _
				& " AND SV.vsVersionText IS NOT NULL AND SV.vsVersionText <> ''"
			sList = MakeList(sQuery, "")
			Application.Lock : Application(sKey) = sList : Application.Unlock
		End If
		GetSoftwareList = Application(sKey)
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetAuthorList()
	'	Purpose: 	get option list of authors
	'Modifications:
	'	Date:		Name:	Description:
	'	1/12/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function GetAuthorList(ByVal v_lItemTypeID, ByVal v_sItemType, ByVal v_sAuthor)
		dim sQuery
		dim sList
		dim sKey
		sKey = v_lItemTypeID & "_" & g_FILTER_AUTHOR
		If Application(sKey) = "" Then
			sQuery = "SELECT DISTINCT U.lUserID, " _
				& "IIf(U.vsScreenName IS NULL, U.vsLastName + ', ' + U.vsFirstName, U.vsScreenName) " _
				& "FROM tblUsers U INNER JOIN tbl" & v_sItemType & "s I ON I.l" & v_sAuthor & "ID = " _
				& "U.lUserID WHERE I.lStatusID = 2 " _
				& "ORDER BY IIf(U.vsScreenName IS NULL, U.vsLastName + ', ' + U.vsFirstName, U.vsScreenName)"
			sList = MakeList(sQuery, "")
			Application.Lock : Application(sKey) = sList : Application.Unlock
		End If
		GetAuthorList = Application(sKey)
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteCategoryList()
	'	Purpose: 	write option list with categories for files
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteCategoryList(ByVal v_sFieldname, ByVal v_aSelected, ByVal v_lSize, ByVal v_lItemType)
		dim x
		
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
			.write GetCategoryList(v_lItemType)
			.write "</select><br><a style='font-size: 8pt;' href=""JavaScript:ClearSelection('"
			.write v_sFieldName
			.write "');"">clear selection</a>"
			.write "<input type='hidden' name='fldCategoryList' value='"
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
	'	Name: 		WriteCategories()
	'	Purpose: 	write plugins, if any
	'Modifications:
	'	Date:		Name:	Description:
	'	1/9/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteCategories(ByVal v_aCats)
		dim x
		dim lCount
		
		If IsArray(v_aCats) Then
			lCount = UBound(v_aCats, 2)
			with response
				.write "</div><nobr>categor"
				.write IIf(lCount > 0, "ies: ", "y: ")
				.write "<span class='CatNote'>"
				for x = 0 to lCount
					.write v_aCats(g_CAT_NAME,x)
					if x < lCount then .write ", "
				next
				.write "</span></nobr></div>"
			end with
		End If
	End Sub
End Class
%>