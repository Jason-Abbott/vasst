<%
Class kbCache
	Private Sub Class_Initialize()
'		Set m_oData = New kbDataAccess
	End Sub
	
	Private Sub Class_Terminate()
'		Set m_oData = nothing
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		WriteRefreshList()
	'	Purpose: 	write list of item lists that can be refreshed
	'Modifications:
	'	Date:		Name:	Description:
	'	1/22/03		JEA		Creation
	'	5/29/03		JEA		Add advertisements
	'	7/22/04		JEA		Add format box
	'-------------------------------------------------------------------------
	Public Sub WriteRefreshList(ByVal v_sPage)
		Const TYPE_ID = 0
		Const TYPE_NAME = 1
		dim aData
		dim oLayout
		dim x
		
		aData = GetItems()
		If IsArray(aData) Then
			Set oLayout = New kbLayout
			Call oLayout.WriteTitleBoxTop("Cached Lists and Banners", "", "")

			with response
				.Write "<table width='300'><tr><td valign='top' class='Explanation'>"
				.Write g_sMSG_ABOUT_CACHE
				.Write "</td><td valign='top'><form name='frmCache' method='post' action='"
				.Write v_sPage
				.Write "'>"
				for x = 0 to UBound(aData, 2)
					.Write "<input type='checkbox' name='fldRefresh' value='"
					.Write aData(TYPE_ID, x)
					.Write "'>"
					.write ToProperCase(aData(TYPE_NAME, x))
					.write "s<br>"
				next
				.Write "<input type='checkbox' name='fldRefresh' value='header'>Header<br>"
				.Write "<input type='checkbox' name='fldRefresh' value='banners'>Banners<p>"
				Call oLayout.WriteToggleImage("btn_refresh", "", "Refresh selected items", "border=0", true)
				.write "</td></form></table>"
			end with
			
			Call oLayout.WriteBoxBottom("")
			Set oLayout = Nothing
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		RefreshSelected()
	'	Purpose: 	refresh posted items
	'Modifications:
	'	Date:		Name:	Description:
	'	7/22/04		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub RefreshSelected(ByVal v_sRefreshList)
		dim aRefresh
		dim lCount
		dim sMessage
		dim x
		
		If Not IsVoid(v_sRefreshList) Then
			aRefresh = Split(v_sRefreshList, ",")
			lCount = UBound(aRefresh) + 1
			sMessage = IIf((lCount > 1), "items have", "item has")
			
			for x = 0 to UBound(aRefresh)
				aRefresh(x) = Trim(aRefresh(x))
				If IsNumber(aRefresh(x)) Then
					' database cache
					Call ClearCache(aRefresh(x))
				Else
					' manually cached items
					Select Case aRefresh(x)
						case "header"
							Call RefreshContent(g_sCACHE_HEADER)
						case "banners"
							Call RefreshContent(g_sCACHE_BANNERS)
					End Select
				End If
			next
			Call SetSessionValue(g_USER_MSG, "The selected " & sMessage & " been refreshed")
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetItems()
	'	Purpose: 	get item types
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/22/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function GetItems()
		dim oData
		Set oData = New kbDataAccess
		GetItems = oData.GetArray("SELECT lItemTypeID, vsDescription FROM tblItemTypes")
		Set oData = nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		RefreshContent()
	'	Purpose: 	read site-specific content and load into cache
	'Modifications:
	'	Date:		Name:	Description:
	'	5/29/03		JEA		Created
	'-------------------------------------------------------------------------
	Private Sub RefreshContent(ByVal v_sCacheName)
		Application(v_sCacheName & "_" & GetSessionValue(g_USER_SITE)) = ""
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		ResetHeader()
	'	Purpose: 	read FrontPage header border and clean for inclusion with kb
	'Modifications:
	'	Date:		Name:	Description:
	'	1/22/03		JEA		Creation
	'	5/29/03		JEA		DEPRECATED
	'-------------------------------------------------------------------------
	Private Sub ResetHeader()
		Const FOR_READING = 1
		dim oFileSys
		dim oFile
		dim oRegExp
		dim sHeader
		dim lStart
		dim lEnd
		dim sPath
		
		sPath = Server.Mappath("/_borders/") & "\top.htm"
		Set oFileSys = Server.CreateObject(g_sFILE_SYSTEM_OBJECT)
		If oFileSys.FileExists(sPath) Then
			Set oFile = oFileSys.OpenTextFile(sPath, FOR_READING)
			sHeader = oFile.ReadAll
			Set oFile = Nothing
			lStart = InStr(sHeader, "<body")
			If lStart <> 0 Then
				lStart = InStr(lStart, sHeader, ">") + 1
				lEnd = InStr(sHeader, "</body")
				sHeader = Mid(sHeader, lStart, lEnd - lStart)
			End If
			Set oRegExp = New RegExp
			With oRegExp
				.Global = true
				.IgnoreCase = true
				.Pattern = "<td[^>]*>"					' fix td tags
				sHeader = .Replace(sHeader, "<td valign='top'>")
				.Pattern = "<table[^>]*>"				' fix table tag
				sHeader = .Replace(sHeader, "<table width='100%' cellspacing='0' cellpadding='0' border='0'>")
				.Pattern = "\.\./images"				' fix image pathing
				sHeader = .Replace(sHeader, "/images")
				.Pattern = "[\n\r\t]+"					' remove extra spacing
				sHeader = .Replace(sHeader, "")
				.Pattern = ">\s+<"						' remove extra spacing
				sHeader = .Replace(sHeader, "><")
				.Pattern = "\s{2,}"						' remove extra spacing
				sHeader = .Replace(sHeader, " ")
				.Pattern = "href\s*=\s*(['""])([^h])"	' fix page paths
				sHeader = .Replace(sHeader, "href=$1/$2")
			End With
			Set oRegExp = Nothing
		Else
			sHeader = g_sDEFAULT_HEADER
		End If
		Set oFileSys = Nothing
		Application(g_sCACHE_HEADER) = sHeader
	End Sub
End Class
%>