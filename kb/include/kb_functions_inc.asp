<%
'-------------------------------------------------------------------------
'	Name: 		MakeKey()
'	Purpose: 	create key for caching arrays
'	Return: 	string
'Modifications:
'	Date:		Name:	Description:
'	1/12/03		JEA		Creation
'	4/25/03		JEA		Add site to key
'-------------------------------------------------------------------------
Function MakeKey(ByVal v_lItemTypeID, ByVal v_aFilter)
	dim sKey
	v_aFilter(g_FILTER_PAGE) = 0
	MakeKey = GetSessionValue(g_USER_SITE) & "-" & v_lItemTypeID & "_" & Join(v_aFilter, "-")
End Function

'-------------------------------------------------------------------------
'	Name: 		ClearCache()
'	Purpose: 	clear cached arrays of given type
'Modifications:
'	Date:		Name:	Description:
'	1/12/03		JEA		Creation
'	4/25/03		JEA		Include site in key
'-------------------------------------------------------------------------
Sub ClearCache(ByVal v_lItemTypeID)
	dim sKey
	dim lSiteID
	lSiteID = GetSessionValue(g_USER_SITE)
	Application.Lock
	For Each sKey In Application.Contents
		If Left(sKey, 4) = lSiteID & "-" & v_lItemTypeID & "_" Then Application(sKey) = ""
	Next
	Application.Unlock
End Sub

'-------------------------------------------------------------------------
'	Name: 		ClearFilterCache()
'	Purpose: 	clear cached filter lists
'Modifications:
'	Date:		Name:	Description:
'	1/12/02		JEA		Creation
'	4/25/03		JEA		Adjust midpoint for site in key
'-------------------------------------------------------------------------
Private Sub ClearFilterCache(ByVal v_lFilterID)
	dim sKey
	Application.Lock
	For Each sKey In Application.Contents
		If Mid(sKey, 4, 2) = "_" & v_lFilterID Then Application(sKey) = ""
	Next
	Application.Unlock
End Sub

'-------------------------------------------------------------------------
'	Name: 		FilterArray()
'	Purpose: 	return array with only matching value for given indices
'	Return: 	string
'Modifications:
'	Date:		Name:	Description:
'	1/9/03		JEA		Creation
'-------------------------------------------------------------------------
Function FilterArray(ByVal v_aData, ByVal v_lParentIndex, ByVal v_lChildIndex, ByVal v_lMatch)
	dim aNewData()
	dim aChild
	dim bMatchChild
	dim bHasMatches
	dim bRowMatch
	dim lNewRow
	dim x, y
	
	bHasMatches = false
	If IsVoid(v_lMatch) Then
		' nothing to match
		FilterArray = v_aData
		Exit Function
	ElseIf IsArray(v_aData) Then
		bMatchChild = IsNumber(v_lChildIndex)
		v_lMatch = CStr(v_lMatch)
		lNewRow = 0
		for x = 0 to UBound(v_aData, 2)
			bRowMatch = false
			if bMatchChild then
				' see if any children match value
				aChild = v_aData(v_lParentIndex, x)
				If IsArray(aChild) Then
					for y = 0 to UBound(aChild, 2)
						if CStr(aChild(v_lChildIndex, y)) = v_lMatch Then
							bRowMatch = True
							exit for
						end if
					next
				End If
			else
				' see if this row matches
				if CStr(v_aData(v_lParentIndex, x)) = v_lMatch Then bRowMatch = true
			end if
			
			if bRowMatch then
				' transfer row to new array
				ReDim Preserve aNewData(UBound(v_aData), lNewRow)
				for y = 0 to UBound(v_aData)
					aNewData(y, lNewRow) = v_aData(y, x)
				next
				lNewRow = lNewRow + 1
				bHasMatches = true
			end if
		next
	End If
	FilterArray = IIf(bHasMatches, aNewData, "")
End Function

'-------------------------------------------------------------------------
'	Name: 		JoinArray()
'	Purpose: 	join child array to parent; both are assumed 2D result of ADO GetArray
'	Return: 	string
'Modifications:
'	Date:		Name:	Description:
'	1/9/03		JEA		Creation
'-------------------------------------------------------------------------
Function JoinArray(ByVal v_aParent, ByVal v_lParentIndex, ByVal v_aChild, ByVal v_lChildIndex)
	dim aJoinParent()
	dim aJoinChild()
	dim bProgeny
	dim lIndex
	dim lParentCols
	dim lChildCols
	dim lCurrentChild
	dim x, y, z
	
	If IsArray(v_aParent) Then
		lParentCols = UBound(v_aParent)
		ReDim aJoinParent(lParentCols + 1, UBound(v_aParent, 2))	' add new column for child link
		If IsArray(v_aChild) Then
			' link children
			lChildCols = UBound(v_aChild)
			for x = 0 to UBound(v_aParent, 2)
				for y = 0 to lParentCols
					' transfer row to new array
					aJoinParent(y, x) = v_aParent(y, x)
				next
				lIndex = aJoinParent(v_lParentIndex, x)
				lCurrentChild = 0
				bProgeny = false
				
				for y = 0 to UBound(v_aChild, 2)
					' find any children
					if v_aChild(v_lChildIndex, y) = lIndex then
						bProgeny = true
						ReDim Preserve aJoinChild(lChildCols, lCurrentChild)
						for z = 0 to lChildCols
							' transfer row to new array
							aJoinChild(z, lCurrentChild) = v_aChild(z, y)
						next
						lCurrentChild = lCurrentChild + 1
					end if
				next
				if bProgeny then
					aJoinParent(lParentCols + 1, x) = aJoinChild
					Erase aJoinChild
				end if
			next
		Else
			for x = 0 to UBound(v_aParent, 2)
				for y = 0 to lParentCols
					' transfer row to new array
					aJoinParent(y, x) = v_aParent(y, x)
				next
			next
		End If
		JoinArray = aJoinParent
	Else
		JoinArray = v_aParent
	End If
End Function

'-------------------------------------------------------------------------
'	Name: 		GetCacheValue()
'	Purpose: 	get Application value
'	Return: 	string
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'-------------------------------------------------------------------------
Function GetCacheValue(ByVal v_sKey)
	If IsArray(Application(g_sCACHE)) Then
		GetCacheValue = Application(g_sCACHE)(v_sKey)
	Else
		GetCacheValue = ""
	End If
End Function

'-------------------------------------------------------------------------
'	Name: 		SetCacheValue()
'	Purpose: 	safe method of checking for value
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'-------------------------------------------------------------------------
Function SetCacheValue(ByVal v_sKey, ByVal v_sValue)
	dim aCache
	aCache = Application(g_sCACHE)
	If Not IsArray(aCache) Then ReDim aCache(6)
	aCache(v_sKey) = v_sValue
	Application(g_sCACHE) = aCache
End Function

'-------------------------------------------------------------------------
'	Name: 		GetSessionValue()
'	Purpose: 	safe method of checking for value
'	Return: 	string
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'-------------------------------------------------------------------------
Function GetSessionValue(ByVal v_sKey)
	If IsArray(Session(g_sSESSION)) Then
		GetSessionValue = CStr(Trim(Session(g_sSESSION)(v_sKey)))
	Else
		GetSessionValue = ""
	End If
End Function

'-------------------------------------------------------------------------
'	Name: 		SetSessionValue()
'	Purpose: 	set session value; setting directly into session object
'				doesn't work, so use intermediate array
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'-------------------------------------------------------------------------
Sub SetSessionValue(ByVal v_sKey, ByVal v_sValue)
	dim aSession
	aSession = Session(g_sSESSION)
	aSession(v_sKey) = v_sValue
	Session(g_sSESSION) = aSession
End Sub

'-------------------------------------------------------------------------
'	Name: 		CleanForSQL()
'	Purpose: 	escape unallowed characters
'Modifications:
'	Date:		Name:	Description:
'	12/24/02	JEA		Created
'-------------------------------------------------------------------------
Function CleanForSQL(ByVal v_sValue)
	CleanForSQL = Replace(v_sValue, "'", "''")
End Function

'-------------------------------------------------------------------------
'	Name: 		FileExists()
'	Purpose: 	does file exist
'	Return: 	boolean
'Modifications:
'	Date:		Name:	Description:
'	12/28/02	JEA		Creation
'-------------------------------------------------------------------------
Function FileExists(ByVal v_sPath, ByVal v_sFile)
	dim oFileSys
	dim bExists
	Set oFileSys = Server.CreateObject(g_sFILE_SYSTEM_OBJECT)
	bExists = oFileSys.FileExists(Server.Mappath(v_sPath & "/" & v_sFile))
	Set oFileSys = nothing
	FileExists = bExists
End Function

'-------------------------------------------------------------------------
'	Name: 		GetURL()
'	Purpose: 	get page name and querystring, optionally fully qualified
'	Return: 	string
'Modifications:
'	Date:		Name:	Description:
'	12/28/02	JEA		Creation
'-------------------------------------------------------------------------
Function GetURL(ByVal v_bComplete)
	dim sQueryString
	dim sPath
	sQueryString = Server.URLEncode(Request.ServerVariables("QUERY_STRING"))
	sPath = Request.ServerVariables("PATH_TRANSLATED")
	sPath = Right(sPath, Len(sPath) - InStrRev(sPath,"\"))
	GetURL = sPath & IIf(IsVoid(sQueryString), "", "?" & sQueryString)
End Function

'-------------------------------------------------------------------------
'	Name: 		GetLastActivityDate()
'	Purpose: 	get most recent date of given activity
'	Return: 	string
'Modifications:
'	Date:		Name:	Description:
'	12/25/02	JEA		Creation
'-------------------------------------------------------------------------
Function GetLastActivityDate(ByVal v_lActivityID)
	dim sQuery
	dim aData
	dim oData

	sQuery = "SELECT TOP 1 dtActivityDate FROM tblLog WHERE lActivityID = " _
		& v_lActivityID & " ORDER BY dtActivityDate DESC"
	Set oData = New kbDataAccess
	aData = oData.GetArray(sQuery)
	Set oData = Nothing
	
	If IsArray(aData) Then
		GetLastActivityDate = aData(0,0)
	Else
		GetLastActivityDate = ""
	End If
End Function

'-------------------------------------------------------------------------
'	Name: 		MakeNumber()
'	Purpose: 	turn string into number
'	Return:		number
'Modifications:
'	Date:		Name:	Description:
'	8/13/02		JEA		Creation
'-------------------------------------------------------------------------
Function MakeNumber(ByVal v_sString)
	v_sString = Trim(v_sString)
	If IsNumber(v_sString) Then
		MakeNumber = CDbl(v_sString)
	Else
		MakeNumber = 0
	End If
End Function

'-------------------------------------------------------------------------
'	Name: 		IsMatch()
'	Purpose: 	run regexp match
'	Return:		boolean
'Modifications:
'	Date:		Name:	Description:
'	5/29/02		JEA		Creation
'-------------------------------------------------------------------------
Function IsMatch(ByVal v_sString, ByVal v_sPattern, ByRef r_oRegExp)
	dim bNewObject
	if Not IsObject(r_oRegExp) then
		Set r_oRegExp = New RegExp
		r_oRegExp.IgnoreCase = true
		r_oRegExp.Global = true
		bNewObject = true
	else
		bNewObject = false
	end if
	r_oRegExp.Pattern = LCase(v_sPattern)
	IsMatch = r_oRegExp.Test(LCase(v_sString))
	if bNewObject then set r_oRegExp = nothing
End Function

'-------------------------------------------------------------------------
'	Name: 		IIf()
'	Purpose: 	I've been missing this
'	Return:		string
'Modifications:
'	Date:		Name:	Description:
'	7/29/02		JEA		Created
'-------------------------------------------------------------------------
Function IIf(v_bCondition, v_sTrue, v_sFalse)
	if v_bCondition then
		IIf = v_sTrue
	else
		IIf = v_sFalse
	end if
End Function

'-------------------------------------------------------------------------
'	Name: 		IsVoid()
'	Purpose: 	check if value is empty in any sense
'	Return:		boolean
'Modifications:
'	Date:		Name:	Description:
'	6/14/02		JEA		Created
'-------------------------------------------------------------------------
Function IsVoid(ByVal v_sValue)
	v_sValue = Trim(v_sValue)
	IsVoid = CBool(v_sValue = "" Or IsEmpty(v_sValue) Or IsNull(v_sValue) Or Len(v_sValue) = 0)
End Function

'-------------------------------------------------------------------------
'	Name: 		IsNumber()
'	Purpose: 	check if value is truly a number
'	Return:		boolean
'Modifications:
'	Date:		Name:	Description:
'	7/24/02		JEA		Created
'-------------------------------------------------------------------------
Function IsNumber(ByVal v_lValue)
	IsNumber = IsNumeric(v_lValue) And Not IsVoid(v_lValue)
End Function

'-------------------------------------------------------------------------
'	Name: 		ReplaceNull()
'	Purpose: 	replace nulls with substitute
'	Return:		string
'Modifications:
'	Date:		Name:	Description:
'	6/18/02		JEA		Creation
'-------------------------------------------------------------------------
Function ReplaceNull(ByVal v_sValue, ByVal v_sReplace)
	ReplaceNull = IIf(IsVoid(v_sValue), v_sReplace, v_sValue)
End Function

'-------------------------------------------------------------------------
'	Name: 		MatchesOne()
'	Purpose: 	does string match any in list
'	Return:		boolean
'Modifications:
'	Date:			Name:	Description:
'	7/30/02			JEA		Creation
'-------------------------------------------------------------------------
Function MatchesOne(ByVal v_sString, ByVal v_aMatchTo, ByVal v_bExact)
	dim bMatch
	dim x
	
	bMatch = false
	
	if IsArray(v_aMatchTo) And Not IsVoid(v_sString) then
		if v_bExact then
			for x = 0 to UBound(v_aMatchTo)
				if CStr(v_sString) = CStr(v_aMatchTo(x)) then
					bMatch = true
					exit for
				end if
			next
		else
			for x = 0 to UBound(v_aMatchTo)
				if InStr(v_sString, v_aMatchTo(x)) > 0 then
					bMatch = true
					exit for
				end if
			next
		end if
	end if
	MatchesOne = bMatch
End Function

'-------------------------------------------------------------------------
'	Name: 		GetObject()
'	Purpose: 	creates object if needed
'	Return:		boolean
'Modifications:
'	Date:		Name:	Description:
'	8/30/00		JEA		Creation
'-------------------------------------------------------------------------
Function GetObject(ByRef r_sObject, ByVal v_sClass)
	if Not IsObject(r_sObject) then
 		' create object if needed
 		Set r_sObject = Server.CreateObject(v_sClass)
 		GetObject = true
	else
		GetObject = false
	end if
End Function

'-------------------------------------------------------------------------
'	Name: 		FormatPhone()
'	Purpose: 	formats number as (xxx) xxx-xxxx
'	Return:		string
'Modifications:
'	Date:		Name:	Description:
'	8/21/00		JEA		Creation
'	7/12/02		JEA		Use new NumbersOnly()
'-------------------------------------------------------------------------
Function FormatPhone(ByVal v_sPhone)
	v_sPhone = NumbersOnly(v_sPhone)
	if Len(v_sPhone) = 10 then
		FormatPhone = "(" & Left(v_sPhone,3) & ") " & Mid(v_sPhone,4,3) & "-" & Right(v_sPhone,4)
	else
		FormatPhone = v_sPhone
	end if
End Function

'-------------------------------------------------------------------------
'	Name: 		FormatDate()
'	Purpose: 	none of FormatDateTime() options are quite right
'	Return:		string
'Modifications:
'	Date:		Name:	Description:
'	8/21/00		JEA		Creation
'-------------------------------------------------------------------------
Function FormatDate(ByVal v_sDate)
	if IsDate(v_sDate) then
		' don't try to format if it's not a date
		FormatDate = MonthName(Month(v_sDate)) & " " & Day(v_sDate) & ", " & Year(v_sDate)
	else
		FormatDate = v_sDate
	end if
End Function

'-------------------------------------------------------------------------
'	Name: 		FormatAsHTML()
'	Purpose: 	replace character codes with HTML
'Modifications:
'	Date:		Name:	Description:
'	1/1/03		JEA		Creation
'-------------------------------------------------------------------------
Function FormatAsHTML(ByVal v_sString)
	If Not IsVoid(v_sString) Then
		v_sString = Replace(v_sString, vbCrLf & vbCrLf, "<p>")
		v_sString = Replace(v_sString, vbCrLf, "<br>")
	End If
	FormatAsHTML = v_sString
End Function

'-------------------------------------------------------------------------
'	Name: 		SafeDate()
'	Purpose: 	write a date that can safely be part of a file name
'	Return:		string
'Modifications:
'	Date:		Name:	Description:
'	12/25/02	JEA		Creation
'-------------------------------------------------------------------------
Function SafeDate(ByVal v_sDate)
	If IsDate(v_sDate) Then
		SafeDate = Year(v_sDate) & "-" & PadNumber(Month(v_sDate), 2) & "-" & PadNumber(Day(v_sDate), 2)
	Else
		SafeDate = v_sDate
	End If
End Function

'-------------------------------------------------------------------------
'	Name: 		PadNumber()
'	Purpose: 	add leading zeros to make given length
'	Return:		string
'Modifications:
'	Date:		Name:	Description:
'	12/25/02	JEA		Creation
'-------------------------------------------------------------------------
Function PadNumber(ByVal v_lNumber, ByVal v_lLength)
	dim lPad
	lPad = v_lLength - Len(v_lNumber)
	If lPad > 0 Then
		PadNumber = String(lPad, "0") & v_lNumber
	Else
		PadNumber = CStr(v_lNumber)
	End If
End Function

'-------------------------------------------------------------------------
'	Name: 		MakeList()
'	Purpose: 	build list from recordset
'				rs must have only id and description fields (two fields)
'	Return:		string
'Modifications:
'	Date:		Name:	Description:
'	12/4/00		JEA		Creation
'	6/19/02		JEA		Check for empty string
'-------------------------------------------------------------------------
Function MakeList(ByVal v_sQuery, ByVal v_sSelect)
	dim oData
	dim sHtml
	
	Set oData = New kbDataAccess
	sHtml = oData.GetString(v_sQuery, "'>", "<option value='")
	Set oData = Nothing
	if sHtml <> "" then
		sHtml = "<option value='" & Left(sHTML, Len(sHTML) - 15)
		sHtml = MakeSelected(sHtml, v_sSelect)
	end if
	MakeList = sHtml
End Function

'-------------------------------------------------------------------------
'	Name: 		WriteListFromArray()
'	Purpose: 	build list from from 1D or 2D array
'	Return:		string
'Modifications:
'	Date:		Name:	Description:
'	7/26/02		JEA		Creation
'-------------------------------------------------------------------------
Sub WriteListFromArray(ByVal v_aValue, ByVal v_aName, ByVal v_sSelect)
	dim sHtml
	dim x
	If IsArray(v_aValue) Then
		sHtml = ""
		if Not IsArray(v_aName) then v_aName = v_aValue
		for x = 0 to UBound(v_aValue)
			sHtml = sHtml & "<option value='" & v_aValue(x) & "'>" & v_aName(x)
		next
		response.write MakeSelected(sHtml, v_sSelect)
	End If
End Sub

'-------------------------------------------------------------------------
'	Name: 		MakeSelected()
'	Purpose: 	selects item in HTML option list
'	Return:		string
'Modifications:
'	Date:		Name:	Description:
'	8/30/00		JEA		Creation
'	6/24/02		JEA		Us IsVoid()
'-------------------------------------------------------------------------
Function MakeSelected(ByVal v_sList, ByVal v_sSelect)
	' assumes <option value='[sSelect]'> with (') delimiters
	if Not IsVoid(v_sSelect) then
		MakeSelected = Replace(v_sList, "'" & v_sSelect & "'", "'" & v_sSelect & "' selected", 1, 1)
	else
		MakeSelected = v_sList
	end if
End Function

'-------------------------------------------------------------------------
'	Name: 		NumbersOnly()
'	Purpose: 	strip non-numbers from string
'	Return:		string
'Modifications:
'	Date:		Name:	Description:
'	6/30/02		JEA		Creation
'	12/6/2002	ZNO		Enhanced to trim before we do anything else
'-------------------------------------------------------------------------
Function NumbersOnly(ByVal v_sNumber)
	dim oRegExp
	If Not IsVoid(v_sNumber) Then
		v_sNumber = Trim(v_sNumber)
		Set oRegExp = New RegExp
		oRegExp.Global = true	
		oRegExp.Pattern = "\D"
		v_sNumber = oRegExp.Replace(v_sNumber, "")
		Set oRegExp = nothing
	end if
	NumbersOnly = v_sNumber
End Function

'-------------------------------------------------------------------------
'	Name: 		MayBeNull()
'	Purpose: 	handle possibly null fields
'	Return:		string
'Modifications:
'	Date:		Name:	Description:
'	6/03/02		JEA		Creation
'	6/18/02		JEA		Use IsVoid()
'	8/22/2002	ZNO		Added replacement if there is a ' in the string
'-------------------------------------------------------------------------
Function MayBeNull(ByVal v_sParameter)
	if IsVoid(v_sParameter) then
		MayBeNull = "null"
	else
		if IsNumeric(v_sParameter) then
			MayBeNull = v_sParameter
		else
			MayBeNull = "'" & CleanForSQL(v_sParameter) & "'"
		end if
	end if
End Function

'-------------------------------------------------------------------------
'	Name: 		VBtoSQLBoolean()
'	Purpose: 	convert VB boolean to SQL boolean
'	Return:		int
'Modifications:
'	Date:		Name:	Description:
'	6/03/02		JEA		Creation
'-------------------------------------------------------------------------
Function VBtoSQLBoolean(ByVal v_bParameter)
	VBtoSQLBoolean = IIf(v_bParameter, 1, 0)
End Function

'-------------------------------------------------------------------------
'	Name: 		SQLtoVBBoolean()
'	Purpose: 	convert SQL boolean to VB boolean
'	Return:		boolean
'Modifications:
'	Date:		Name:	Description:
'	6/06/02		JEA		Creation
'-------------------------------------------------------------------------
Function SQLtoVBBoolean(ByVal v_lParameter)
	SQLtoVBBoolean = IIf((CStr(v_lParameter & "") = "1"), true, false)
End Function

'-------------------------------------------------------------------------
'	Name: 		SayNumber()
'	Purpose: 	replace number with word
'	Return: 	string
'Modifications:
'	Date:		Name:	Description:
'	10/24/02		JEA		Creation
'-------------------------------------------------------------------------
Function SayNumber(ByVal v_lNumber)
	if IsNumber(v_lNumber) then
		select case MakeNumber(v_lNumber)
			case 0 : SayNumber = "zero"
			case 1 : SayNumber = "one"
			case 2 : SayNumber = "two"
			case 3 : SayNumber = "three"
			case 4 : SayNumber = "four"
			case 5 : SayNumber = "five"
			case 6 : SayNumber = "six"
			case 7 : SayNumber = "seven"
			case 8 : SayNumber = "eight"
			case 9 : SayNumber = "nine"
			case 10 : SayNumber = "ten"
			case 11 : SayNumber = "eleven"
			case 12 : SayNumber = "twelve"
			case 13 : SayNumber = "thirteen"
			case 14 : SayNumber = "fourteen"
			case 15 : SayNumber = "fifteen"
			case 16 : SayNumber = "sixteen"
			case 17 : SayNumber = "seventeen"
			case 18 : SayNumber = "eighteen"
			case 19 : SayNumber = "nineteen"
			case 20 : SayNumber = "twenty"
			case 30 : SayNumber = "thirty"
			case 40 : SayNumber = "forty"
			case 50 : SayNumber = "fifty"
			case 60 : SayNumber = "sixty"
			case 70 : SayNumber = "seventy"
			case 80 : SayNumber = "eighty"
			case 90 : SayNumber = "ninety"
			case else : SayNumber = v_lNumber
		end select
	else
		SayNumber = v_lNumber
	end if
End Function
%>