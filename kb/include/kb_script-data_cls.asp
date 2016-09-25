<%
Const m_SCRIPT_ID = 0
Const m_SCRIPT_NAME = 1
Const m_SCRIPT_PATH = 2
Const m_SCRIPT_FRIENDLY_NAME = 3
Const m_SCRIPT_DATE_ADDED = 4
Const m_SCRIPT_DESCRIPTION = 5
Const m_SCRIPT_FORMAT = 6
Const m_SCRIPT_MEDIA = 7
Const m_SCRIPT_STATUS = 8
Const m_SCRIPT_SIZE = 9
Const m_SCRIPT_SOFTWARE_NAME = 10
Const m_SCRIPT_SOFTWARE_ID = 11
Const m_SCRIPT_USER_ID = 12
Const m_SCRIPT_CREATOR = 13
Const m_SCRIPT_RANK = 14
Const m_SCRIPT_VERSION = 15
Const m_SCRIPT_DOWNLOADS = 16
Const m_SCRIPT_CATS = 17			' generated from join

Const m_SCRIPT_NO_VERSION = 5		' from database

'-------------------------------------------------------------------------
'	Name: 		kbScriptData class
'	Purpose: 	methods of getting file data
'Modifications:
'	Date:		Name:	Description:
'	7/21/04		JEA		Copied from files class
'-------------------------------------------------------------------------
Class kbScriptData
	Private m_oData
	Private m_sBaseSQL

	Private Sub Class_Initialize()
		Set m_oData = New kbDataAccess
		m_sBaseSQL = "SELECT F.lScriptID, F.vsFileName, F.vsPath, F.vsFriendlyName, F.dtVersionDate, " _
			& "F.vsDescription, F.lFormatID, F.vsRequiredMediaURL, F.lStatusID, " _
			& "F.lScriptSize, IIf(SV.lVersionID IS NULL, '', S.vsSoftwareName + ' ' + SV.vsVersionText), " _
			& "SV.lVersionID, U.lUserID, IIf(U.vsScreenName IS NULL, " _
			& "U.vsFirstName + ' ' + U.vsLastName, U.vsScreenName), CR.fRank, F.lVersionCount, F.lDownloads " _
			& "FROM ((((tblScripts F INNER JOIN tblUsers U ON U.lUserID = F.lUserID) " _
			& "INNER JOIN (SELECT lItemID FROM tblItemSites WHERE lItemTypeID = " _
			& g_ITEM_SCRIPT & " AND lSiteID = " & GetSessionValue(g_USER_SITE) _
			& ") tIS ON tIS.lItemID = F.lScriptID) " _
			& "LEFT JOIN tblSoftwareVersions SV ON SV.lVersionID = F.lSoftwareVersionID) " _
			& "LEFT JOIN tblSoftware S ON S.lSoftwareID = SV.lSoftwareID) " _
			& "LEFT JOIN (SELECT lItemID, fRank FROM tblComputedRankings WHERE lItemTypeID = " _
			& g_ITEM_SCRIPT & ") CR ON CR.lItemID = F.lScriptID "
	End Sub
	
	Private Sub Class_Terminate()
		Set m_oData = nothing
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		GetPending()
	'	Purpose: 	get array of uploaded scripts
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Function GetPending()
		dim sQuery
		dim aData
		sQuery = Replace(m_sBaseSQL, "F.dtVersionDate", "F.dtSubmitDate") _
			& "WHERE F.lStatusID = " & g_STATUS_PENDING _
			& " ORDER BY F.dtSubmitDate"
		GetPending = m_oData.GetArray(sQuery)
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetPublic()
	'	Purpose: 	get array of scripts visible to public
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Function GetPublic(ByVal v_aFilter)
		dim sScriptsSQL
		dim sCatSQL
		dim aData
		dim sKey
		
		sKey = MakeKey(g_ITEM_SCRIPT, v_aFilter)
		aData = Application(sKey)
		If Not IsArray(aData) Then
			sScriptsSQL = m_sBaseSQL & "WHERE F.lStatusID = " & g_STATUS_APPROVED & " "
			If v_aFilter(g_FILTER_AUTHOR) <> 0 Then
				sScriptsSQL = sScriptsSQL & "AND U.lUserID = " & v_aFilter(g_FILTER_AUTHOR) & " "
			End If
			If v_aFilter(g_FILTER_SOFTWARE) <> 0 Then
				'sScriptsSQL = sScriptsSQL & "AND SV.lVersionID IN (" & v_aFilter(g_FILTER_SOFTWARE) & "," & m_SCRIPT_NO_VERSION & ") "
				sScriptsSQL = sScriptsSQL & "AND SV.lVersionID = " & v_aFilter(g_FILTER_SOFTWARE) & " "
			End If
			sScriptsSQL = sScriptsSQL & MakeSortSQL(v_aFilter(g_FILTER_SORT))
			sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
				& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
				& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_SCRIPT
			aData = m_oData.GetArray(sScriptsSQL)
			aData = JoinArray(aData, m_SCRIPT_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
			If v_aFilter(g_FILTER_CATEGORY) <> 0 Then
				aData = FilterArray(aData, m_SCRIPT_CATS, g_CAT_ID, v_aFilter(g_FILTER_CATEGORY))
			End If
			Application.Lock
			Application(sKey) = aData
			Application.Unlock
		End If
		GetPublic = aData
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		MakeSortSQL()
	'	Purpose: 	generate SQL for sorting scripts
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Private Function MakeSortSQL(ByVal v_lSortID)
		dim sSQL
		sSQL = "ORDER BY "
		select case v_lSortID
			case g_SORT_NAME_ASC
				sSQL = sSQL & "F.vsFriendlyName"
			case g_SORT_NAME_DESC
				sSQL = sSQL & "F.vsFriendlyName DESC"
			case g_SORT_DATE_ASC
				sSQL = sSQL & "F.dtVersionDate, F.vsFriendlyName"
			case g_SORT_DATE_DESC
				sSQL = sSQL & "F.dtVersionDate DESC, F.vsFriendlyName"
			case g_SORT_OWNER_ASC
				sSQL = sSQL & "U.vsFirstName, U.vsLastName, F.vsFriendlyName"
			case g_SORT_OWNER_DESC
				sSQL = sSQL & "U.vsFirstName DESC, U.vsLastName DESC, F.vsFriendlyName"
			case else
				sSQL = sSQL & "F.vsFriendlyName"
		end select
		MakeSortSQL = sSQL
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetDownloadURL()
	'	Purpose: 	create and write download link
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Function GetDownloadURL(ByVal v_lScriptID)
		Const ITEM_NAME = 0
		Const ITEM_PATH = 1
		Const ITEM_USER = 2
		dim oFileSys
		dim sQuery
		dim aData
		
		sQuery = "SELECT vsFileName, vsPath, lUserID FROM tblScripts WHERE lScriptID = " & v_lScriptID
		if Not g_bAdmin Then sQuery = sQuery & " AND lStatusID = " & g_STATUS_APPROVED
		
		aData = m_oData.GetArray(sQuery)
		
		If IsArray(aData) Then
			' increment download count
			If CStr(aData(ITEM_USER, 0)) <> CStr(GetSessionValue(g_USER_ID)) Then
				sQuery = "UPDATE tblScripts SET lDownloads = lDownloads + 1 WHERE lScriptID = " & v_lScriptID
				Call m_oData.ExecuteOnly(sQuery)
				Call UpdateCacheCount(v_lScriptID)
			End If
		
			Set oFileSys = New kbProjectsystem
			GetDownloadURL = oFileSys.CreateDownloadURL(v_lScriptID, g_ITEM_SCRIPT, aData(ITEM_PATH, 0), aData(ITEM_NAME, 0))
			Set oFileSys = Nothing
		End If
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		UpdateCacheCount()
	'	Purpose: 	update download count in cache so entire cache doesn't have to be refreshed
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Private Sub UpdateCacheCount(ByVal v_lItemID)
		dim aData
		dim sPattern
		dim bUpdated
		dim sKey
		dim x

		sPattern = GetSessionValue(g_USER_SITE) & "-" & g_ITEM_SCRIPT & "_"
		Application.Lock
		For Each sKey In Application.Contents
			bUpdated = false
			If Left(sKey, 4) = sPattern Then
				' this is a file cache
				aData = Application(sKey)
				If IsArray(aData) Then
					for x = 0 to UBound(aData, 2)
						if CStr(aData(m_SCRIPT_ID, x)) = CStr(v_lItemID) then
							aData(m_SCRIPT_DOWNLOADS, x) = aData(m_SCRIPT_DOWNLOADS, x) + 1
							bUpdated = true
							exit for
						end if
					next
					if bUpdated then Application(sKey) = aData
				End If
			End If
		Next
		Application.Unlock
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		ApprovePending()
	'	Purpose: 	update script status
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Sub ApprovePending(ByVal v_lScriptID)
		dim sQuery
		dim oMail
		
		sQuery = "UPDATE tblScripts SET " _
			& "lStatusID = " & g_STATUS_APPROVED & ", " _
			& "dtApproveDate = " & g_sSQL_DATE_DELIMIT & Date() & g_sSQL_DATE_DELIMIT & ", " _
			& "dtVersionDate = " & g_sSQL_DATE_DELIMIT & Date() & g_sSQL_DATE_DELIMIT & ", " _
			& "lVersionCount = 1 " _
			& "WHERE lScriptID = " & v_lScriptID
		Call m_oData.ExecuteOnly(sQuery)
		Call m_oData.LogActivity(g_ACT_APPROVE_UPLOAD, "", v_lScriptID, "", "", "", "")
		Set oMail = New kbMail
		Call oMail.SendApprovalEmail(v_lScriptID, g_ITEM_SCRIPT)
		Set oMail = Nothing
		Call ClearCache(g_ITEM_SCRIPT)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		DenyPending()
	'	Purpose: 	update script status and move file
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Sub DenyPending(ByVal v_lScriptID)
		Call Disable(v_lScriptID, g_STATUS_REJECTED, g_ACT_DENY_UPLOAD, g_sDENY_DIR, false)
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		Delete()
	'	Purpose: 	get array for single script
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Sub Delete(ByVal v_lScriptID)
		dim oContest
		Call Disable(v_lScriptID, g_STATUS_DISABLED, g_ACT_DELETE_FILE, g_sDELETE_DIR, true)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Disable()
	'	Purpose: 	update script status and move script
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Private Sub Disable(ByVal v_lScriptID, ByVal v_lNewStatusID, ByVal v_lActionID, _
		ByVal v_sBackupPath, ByVal v_bPurgeCache)
		
		Const ITEM_NAME = 0
		Const ITEM_PATH = 1
		dim sQuery
		dim oFileSys
		dim aData
		
		sQuery = "SELECT vsFileName, vsPath FROM tblScripts WHERE lScriptID = " & v_lScriptID
		aData = m_oData.GetArray(sQuery)
		
		Set oFileSys = New kbProjectsystem
		Call oFileSys.DeleteFromDisk(aData(ITEM_NAME, 0), aData(ITEM_PATH, 0), v_sBackupPath)
		Set oFileSys = Nothing

		sQuery = "UPDATE tblScripts SET lStatusID = " & v_lNewStatusID _
			& ", vsPath = '" & g_sFILES_DIR & "/" & v_sBackupPath & "'" _
			& " WHERE lScriptID = " & v_lScriptID
		Call m_oData.ExecuteOnly(sQuery)
		sQuery = "DELETE FROM tblItemSites WHERE lItemTypeID = " & g_ITEM_SCRIPT _
			& " AND lItemID = " & v_lScriptID
		Call m_oData.ExecuteOnly(sQuery)
		Call m_oData.LogActivity(v_lActionID, "", v_lScriptID, "", "", "", "")
		If v_bPurgeCache Then Call ClearCache(g_ITEM_SCRIPT)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetItem()
	'	Purpose: 	get array for single file
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Function GetItem(ByVal v_lScriptID, ByVal v_bOwnerOnly)
		dim sScriptsSQL
		dim sCatSQL
		dim aData
		
		sScriptsSQL = m_sBaseSQL & "WHERE F.lScriptID = " & v_lScriptID
		If v_bOwnerOnly And Not g_bAdmin Then
			sScriptsSQL = sScriptsSQL & " AND F.lUserID = " & GetSessionValue(g_USER_ID)
		End If
		sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
			& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
			& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_SCRIPT _
			& " AND IC.lItemID = " & v_lScriptID
		aData = m_oData.GetArray(sScriptsSQL)
		aData = JoinArray(aData, m_SCRIPT_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
		GetItem = aData
	End Function

	'-------------------------------------------------------------------------
	'	Name: 		SaveCategories()
	'	Purpose: 	save script plugin data to database
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Sub SaveCategories(ByVal v_lScriptID, ByVal v_sCategories)
		dim aCategories
		dim sQuery
		dim x
		
		If v_sCategories <> "" Then
			' delete any old plugins listed
			sQuery = "DELETE FROM tblItemCategories WHERE lItemID = " & v_lScriptID _
				& " AND lItemTypeID = " & g_ITEM_SCRIPT
			Call m_oData.ExecuteOnly(sQuery)
			
			' insert new plugins
			sQuery = "INSERT INTO tblItemCategories (lItemID, lItemTypeID, lCategoryID) " _
				& "VALUES (" & v_lScriptID & ", " & g_ITEM_SCRIPT & ", "
			aCategories = Split(v_sCategories, ",")
			for x = 0 to UBound(aCategories)
				Call m_oData.ExecuteOnly(sQuery & Trim(aCategories(x)) & ")")
			next
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Save()
	'	Purpose: 	save uploaded script to disk and database
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Sub Save(ByVal v_sFileField)
		dim oForm
		dim oFileSys
		dim sMessage
		dim sFileName
		dim lFileSize
		dim lScriptID
		dim sURL
		
		v_sFileField = LCase(v_sFileField)
		Set oForm = New kbForm
		Set oFileSys = New kbProjectsystem
		Call oForm.ParseFields()
		lScriptID = oForm.Field("fldScriptID")
		sURL = ReplaceNull(oForm.Field("fldURL"), "kb_scripts.asp")
		
		Select Case MakeNumber(oForm.Field("fldAction"))
			Case g_ACT_FILE_ADD
				sURL = "kb_script-submit.asp"
				If oForm.File.Exists(v_sFileField) Then
					sMessage = oFileSys.SaveToDisk(oForm, sFileName, lFileSize, v_sFileField)
					If sMessage = "" Then
						sMessage = g_sMSG_AFTER_UPLOAD
						lScriptID = Insert(sFileName, lFileSize, oForm)
						Call m_oData.LogActivity(g_ACT_FILE_UPLOAD, "", lScriptID, "", "", "", "")
					End If
				Else
					sMessage = "Unable to read file data"
				End If
		
			Case g_ACT_FILE_UPDATE
				If oForm.File.Exists(v_sFileField) Then
					' delete old file and add new
					Call oFileSys.DeleteFromDisk(lScriptID, g_sDELETE_DIR, "")
					sMessage = oFileSys.SaveToDisk(oForm, sFileName, lFileSize, v_sFileField)
				End If
				Call Update(lScriptID, sFileName, lFileSize, oForm)
				
			Case g_ACT_FILE_DELETE
				Call Delete(lScriptID)
		End Select
		
		Set oFileSys = Nothing
		Set oForm = Nothing
		Call SetSessionValue(g_USER_MSG, sMessage)
		response.redirect sURL
	End Sub	
	
	'-------------------------------------------------------------------------
	'	Name: 		Insert()
	'	Purpose: 	save script info to database
	'	Return: 	number
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Private Function Insert(ByVal v_sScriptName, ByVal v_lScriptSize, ByRef r_oForm)
		dim lScriptID
		dim oRS
		dim oScriptData
		dim sQuery

		sQuery = "INSERT INTO tblItemSites (lSiteID, lItemTypeID, lItemID) VALUES (" _
			& GetSessionValue(g_USER_SITE) & ", " & g_ITEM_SCRIPT
		Call m_oData.BeginTrans()
		Set oRS = Server.CreateObject("ADODB.Recordset")
		With oRS
			.Open "tblScripts", m_oData.Connection, adOpenStatic, adLockOptimistic, adCmdTable
			.AddNew
			.Fields("lUserID") = GetSessionValue(g_USER_ID)
			.Fields("vsPath") = g_sFILES_DIR
			.Fields("vsFileName") = v_sScriptName
			.Fields("dtSubmitDate") = Now()
			.Fields("vsFriendlyName") = r_oForm.Field("fldFriendlyName")
			.Fields("vsDescription") = r_oForm.Field("fldDescription")
			.Fields("lSoftwareVersionID") = r_oForm.Field("fldVersion")
			.Fields("vsRequiredMediaURL") = Replace(r_oForm.Field("fldMedia"), "http://", "")
			.Fields("lScriptSize") = v_lScriptSize
			.Fields("lVersionCount") = 1
			.Fields("lDownloads") = 0
			.Fields("lStatusID") = g_STATUS_PENDING
			.Update
			lScriptID = .Fields("lScriptID")
			.Close
		End With
		Set oRS = nothing
		Call SaveCategories(lScriptID, Trim(r_oForm.Field("fldCategories")))
		Call m_oData.ExecuteOnly(sQuery & ", " & lScriptID & ")")
		Call m_oData.CommitTrans()
		Insert = lScriptID
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		Update()
	'	Purpose: 	save script data to database
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Private Sub Update(ByVal v_lScriptID, ByVal v_sScriptName, ByVal v_lScriptSize, ByRef r_oForm)
		dim sQuery
		
		With r_oForm
			sQuery = "UPDATE tblScripts SET " _
				& "vsFriendlyName = '" & CleanForSQL(.Field("fldFriendlyName")) & "', " _
				& "vsDescription = '" & CleanForSQL(.Field("fldDescription")) & "', " _
				& "lSoftwareVersionID = " & .Field("fldVersion") & ", " _
				& "vsRequiredMediaURL = '" & Replace(.Field("fldMedia"), "http://", "") & "'"
			If v_sScriptName <> "" Then
				' update file and version info
				sQuery = sQuery _
					& ", vsFileName = '" & v_sScriptName & "', " _
					& "lScriptSize = " & MakeNumber(v_lScriptSize) & ", " _
					& "dtVersionDate = " & g_sSQL_DATE_DELIMIT & Date() & g_sSQL_DATE_DELIMIT & ", " _
					& "lVersionCount = lVersionCount + 1"
			End If
			sQuery = sQuery & " WHERE lScriptID = " & v_lScriptID
				
			Call m_oData.BeginTrans()
			Call m_oData.ExecuteOnly(sQuery)
			Call SaveCategories(v_lScriptID, Trim(.Field("fldCategories")))
			Call m_oData.CommitTrans()
		End With
		
		Call m_oData.LogActivity(g_ACT_EDIT_FILE_ENTRY, "", v_lScriptID, "", "", "", "")
		Call SetSessionValue(g_USER_MSG, g_sMSG_FILE_EDIT)
		Call ClearCache(g_ITEM_SCRIPT)
	End Sub
End Class
%>