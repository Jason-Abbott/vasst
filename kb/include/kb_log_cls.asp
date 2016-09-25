<%
'-------------------------------------------------------------------------
'	Name: 		kbFiles class
'	Purpose: 	encapsulate activity logging functions
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'-------------------------------------------------------------------------
Class kbLog
	Private m_oData

	Private Sub Class_Initialize()
		Set m_oData = New kbDataAccess
	End Sub
	
	Private Sub Class_Terminate()
		Set m_oData = nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetActivity()
	'	Purpose: 	get array of logged activities
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	4/25/03		JEA		Allow site ID filter
	'-------------------------------------------------------------------------
	Private Function GetActivity(ByVal v_sStartDate, ByVal v_sEndDate, ByVal v_lActivityID, _
		ByVal v_lUserID, ByVal v_lFileID, ByVal v_lSiteID)
		
		dim sQuery
		dim aData
	
		sQuery = "SELECT L.lUserID, L.lFileID, L.lContestID, L.dtActivityDate, L.vsEmail, " _
			& "L.vsPassword, L.vsIPAddress, LA.vsActivityName, U.vsFirstName, U.vsLastName, " _
			& "F.vsFriendlyName, F.vsPath, F.vsFileName, C.vsContestName, S.vsSiteName " _
			& "FROM ((((tblLog L " _
			& "INNER JOIN tblLoggedActivities LA ON L.lActivityID = LA.lActivityID) " _
			& "INNER JOIN tblSite S ON S.lSiteID = L.lSiteID) " _
			& "LEFT JOIN tblUsers U ON U.lUserID = L.lUserID) " _
			& "LEFT JOIN tblFiles F ON F.lFileID = L.lFileID) " _
			& "LEFT JOIN tblContests C ON C.lContestID = L.lContestID " _
			& "WHERE (L.dtActivityDate BETWEEN " & g_sSQL_DATE_DELIMIT _
			& v_sStartDate & g_sSQL_DATE_DELIMIT & " AND " & g_sSQL_DATE_DELIMIT _
			& v_sEndDate & g_sSQL_DATE_DELIMIT & ")"
		if v_lActivityID <> 0 then sQuery = sQuery & " AND LA.lActivityID = " & v_lActivityID
		if v_lUserID <> 0 then sQuery = sQuery & " AND U.lUserID = " & v_lUserID
		if v_lFileID <> 0 then sQuery = sQuery & " AND F.lFileID = " & v_lFileID
		if v_lSiteID <> 0 then sQuery = sQuery & " AND S.lSiteID = " & v_lSiteID
		sQuery = sQuery	& " ORDER BY L.dtActivityDate DESC"
		GetActivity = m_oData.GetArray(sQuery)
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteActivity()
	'	Purpose: 	output activity as HTML
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	4/25/03		JEA		Write Site name
	'-------------------------------------------------------------------------
	Public Sub WriteActivity(ByVal v_sStartDate, ByVal v_sEndDate, ByVal v_lActivityID, _
		ByVal v_lUserID, ByVal v_lFileID, ByVal v_lSiteID)
		
		Const USER_ID = 0
		Const FILE_ID = 1
		Const CONTEST_ID = 2
		Const ACTIVITY_DATE = 3
		Const EMAIL = 4
		Const PASSWORD = 5
		Const IP_ADDRESS = 6
		Const ACTIVITY_NAME = 7
		Const FIRST_NAME = 8
		Const LAST_NAME = 9
		Const FRIENDLY_FILE_NAME = 10
		Const FILE_PATH = 11
		Const FILE_NAME = 12
		Const CONTEST_NAME = 13
		Const SITE_NAME = 14
		dim aData
		dim sDate
		dim sShiftedDate
		dim x
		
		If IsDate(v_sStartDate) And IsDate(v_sEndDate) Then
			v_sStartDate = v_sStartDate & " 12:00 AM"
			v_sEndDate = v_sEndDate & " 11:59 PM"
		Else
			Exit Sub
		End If
		aData = GetActivity(v_sStartDate, v_sEndDate, v_lActivityID, v_lUserID, v_lFileID, v_lSiteID)
		sDate = ""
		With Response
			If IsArray(aData) Then
				.write "<table cellspacing='0' cellpadding='0' border='0'><tr>"
				.write "<td class='LogHead'>Time</td><td class='LogHead'>User</td>"
				.write "<td class='LogHead'>Activity</td><td class='LogHead'>Site</td>"
				.write "<td class='LogHead'>IP</td>"
				For x = 0 to UBound(aData,2)
					sShiftedDate = DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), aData(ACTIVITY_DATE, x))
					' new date row
					if sDate <> FormatDateTime(sShiftedDate, 2) then
						.write "<tr><td class='LogDate' colspan='5'>"
						.write FormatDateTime(sShiftedDate, 1)
						.write "</td>"
						sDate = FormatDateTime(sShiftedDate, 2)
					end if
					' time
					.write "<tr><td class='LogTime'>"
					.write FormatDateTime(sShiftedDate, 3)
					' user
					.write "</td><td class='LogUser' align='center'>"
					If aData(USER_ID, x) = 0 And aData(EMAIL, x) <> "" Then
						.write "<a href='mailto:"
						.write aData(EMAIL, x)
						.write "'>"
						.write aData(EMAIL, x)
					Else
						.write "<a href='kb_user.asp?id="
						.write aData(USER_ID, x)
						.write "'>"
						.write aData(FIRST_NAME, x)
						.write " "
						.write aData(LAST_NAME, x)
					End If
					' activity
					.write "</a></td><td class='LogActivity'>"
					.write aData(ACTIVITY_NAME, x)
					If IsNumber(aData(FILE_ID, x)) And aData(FILE_ID, x) > 0 Then
						.write " (<a href='kb_download.asp?id="
						.write aData(FILE_ID,x)
						.write "'>"
						.write aData(FRIENDLY_FILE_NAME, x)
						.write "</a>)"
					ElseIf IsNumber(aData(CONTEST_ID, x)) And aData(CONTEST_ID, x) > 0 Then
						.write " (<a href='kb_contest.asp?id="
						.write aData(CONTEST_ID,x)
						.write "'>"
						.write aData(CONTEST_NAME, x)
						.write "</a>)"
					End If
					' site
					.write "</td><td class='LogSite'>"
					.write aData(SITE_NAME, x)
					' ip
					.write "</td><td class='LogIP'>"
					.write aData(IP_ADDRESS, x)
					.write "</td>"
					Next
				.write "</table>"
			Else
				.write "There are no logged activities"		
			End If
		End With
	End Sub
End Class
%>