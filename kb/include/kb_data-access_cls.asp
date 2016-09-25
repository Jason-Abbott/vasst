<%
'dim g_sDSN		' connection string
'g_sDSN = g_sDB_CONNECT & Server.Mappath(g_sDB_LOCATION)

'-------------------------------------------------------------------------
'	Name: 		kbData()
'	Purpose: 	encapsulate data functionality
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Created
'-------------------------------------------------------------------------
Class kbDataAccess
	Public Connection
	Private m_bInTransaction

	Private Sub Class_Initialize()
		Set Connection = Server.CreateObject("ADODB.Connection")
		Connection.Open g_sDB_CONNECT & Server.Mappath(g_sDB_LOCATION)
		m_bInTransaction = false
	End Sub
	
	Private Sub Class_Terminate()
		Connection.Close
		Set Connection = nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		*Trans()
	'	Purpose: 	manage ADO transactions
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/02		JEA		Created
	'-------------------------------------------------------------------------
	Public Sub BeginTrans()
	    Connection.beginTrans
	    m_bInTransaction = True
	End Sub
	Public Sub CommitTrans()
	    If m_bInTransaction Then Connection.commitTrans
		m_bInTransaction = False
	End Sub
	Public Sub RollbackTrans()
    	If m_bInTransaction Then Connection.rollbackTrans
	    m_bInTransaction = False
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		ExecuteOnly()
	'	Purpose: 	execute query without results
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Created
	'-------------------------------------------------------------------------
	Public Sub ExecuteOnly(ByVal v_sQuery)
		'response.Write v_sQuery : response.flush
		Connection.Execute v_sQuery, , adExecuteNoRecords
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetArray()
	'	Purpose: 	get array from recordset
	'	Return:		array
	'Modifications:
	'	Date:		Name:	Description:
	'	12/23/02	JEA		Created
	'-------------------------------------------------------------------------
	Public Function GetArray(ByVal v_sQuery)
		dim oRS
		dim aData
		Set oRS = NewRecordSet(v_sQuery)
		If Not oRS.EOF Then aData = oRS.GetRows
		oRS.Close : Set oRS = nothing
		GetArray = aData
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetString()
	'	Purpose: 	get delimited string from query
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/30/02	JEA		Created
	'	3/22/03		JEA		Check for empty recordset
	'-------------------------------------------------------------------------
	Public Function GetString(ByVal v_sQuery, ByVal v_sColDelim, ByVal v_sRowDelim)
		dim oRS
		Set oRS = NewRecordSet(v_sQuery)
		If Not oRS.EOF Then
			GetString = oRS.GetString(adClipString, , v_sColDelim, v_sRowDelim)
			oRS.Close
		Else
			GetString = ""
		End If
		Set oRS = Nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetJSArray()
	'	Purpose: 	create JavaScript array literal from array
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/23/02	JEA		Created
	'-------------------------------------------------------------------------
	Public Function GetJSArray(ByVal v_sQuery)
		dim sJSArray
		sJSArray = GetString(v_sQuery, """,""", """],[""")
		if sJSArray <> "" then sJSArray = "[""" & Left(sJSArray, Len(sJSArray) - 4) & "]"
		GetJSArray = sJSArray 
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		NewRecordSet()
	'	Purpose: 	return open, disconnected recordset
	'Modifications:
	'	Date:		Name:	Description:
	'	12/23/02	JEA		Created
	'-------------------------------------------------------------------------
	Private Function NewRecordSet(ByVal v_sQuery)
		'on error resume next
		dim oRS
		Set oRS = Server.CreateObject("ADODB.Recordset")
		Set oRS.ActiveConnection = Connection
		oRS.CursorLocation = adUseClient
		oRS.Open v_sQuery, , adOpenForwardOnly, adLockReadOnly, adCmdText
		'response.Write v_squery & "<p>" : response.flush
		Set oRS.ActiveConnection = nothing
		Set NewRecordSet = oRS
		Set oRS = nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		RunQuery()
	'	Purpose: 	run raw query; only admins can do this
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/03	JEA		Created
	'-------------------------------------------------------------------------
	Public Function RunQuery(ByVal v_sQuery)
		dim lAffected
		dim oRS
		If GetSessionValue(g_USER_TYPE) <> CStr(g_USER_ADMIN) Then Exit Function
		If LCase(Left(v_sQuery, 6)) = "select" Then
			' get recordset
			Set RunQuery = NewRecordSet(v_sQuery)
		Else
			' indicate rows affected
			Call Connection.Execute(v_sQuery, lAffected, adExecuteNoRecords)
			Set oRS = Server.CreateObject("ADODB.Recordset")
			With oRS
				.Fields.Append "Affected Rows", adInteger	', , , lAffected
				.Open
				.AddNew
				.Fields("Affected Rows") = lAffected
				'.Update
			End With
			Set RunQuery = oRS
			Set oRS = Nothing
		End If
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		LogActivity()
	'	Purpose: 	log activities
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	7/21/04		JEA		Log script ID
	'-------------------------------------------------------------------------
	Public Sub LogActivity(ByVal v_lActivityID, ByVal v_lProjectID, ByVal v_lScriptID, _
		ByVal v_lContestID, ByVal v_lAffectedUserID, ByVal v_sEmail, ByVal v_sPassword)
		dim sQuery
		dim lUserID
		dim sIPAddress
		
		If IsArray(Session("user")) Then
			lUserID = Session("user")(g_USER_ID)
			sIPAddress = Session("user")(g_USER_IP)
		Else
			lUserID = 0
			sIPAddress = Request.ServerVariables("REMOTE_ADDR")
		End If
		
		sQuery = "INSERT INTO tblLog (lUserID, lActivityID, lProjectID, lScriptID, " _
			& "lContestID, lAffectedUserID, " _
			& "vsEmail, vsPassword, dtActivityDate, vsIPAddress, lSiteID) VALUES (" _
			& lUserID & ", " _
			& v_lActivityID & ", " _
			& MakeNumber(v_lProjectID) & ", " _
			& MakeNumber(v_lScriptID) & ", " _
			& MakeNumber(v_lContestID) & ", " _
			& MakeNumber(v_lAffectedUserID) & ", '" _
			& v_sEmail & "', '" _
			& v_sPassword & "', '" _
			& Now() & "', '" _
			& sIPAddress & "', " _
			& MaybeNull(GetSessionValue(g_USER_SITE)) & ")"
		Call ExecuteOnly(sQuery)
	End Sub
End Class
%>