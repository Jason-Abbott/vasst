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
	'-------------------------------------------------------------------------
	Public Sub WriteRefreshList(ByVal v_sPage)
		Const TYPE_ID = 0
		Const TYPE_NAME = 1
		dim aData
		dim x
		
		aData = GetItems()
		If IsArray(aData) Then
			with response
				for x = 0 to UBound(aData, 2)
					.write "<a href='"
					.write v_sPage
					.write "?type="
					.write aData(TYPE_ID, x)
					.write "&do=reset'>"
					.write aData(TYPE_NAME, x)
					.write "s</a><br>"
				next
				.write "<a href='"
				.write v_sPage
				.write "?do=header'>header</a>"	
			end with
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
	'	Name: 		ResetHeader()
	'	Purpose: 	read FrontPage header border and clean for inclusion with kb
	'Modifications:
	'	Date:		Name:	Description:
	'	1/22/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub ResetHeader()
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