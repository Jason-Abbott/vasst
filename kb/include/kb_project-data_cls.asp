<%
Const m_PROJECT_ID = 0
Const m_PROJECT_NAME = 1
Const m_PROJECT_PATH = 2
Const m_FRIENDLY_NAME = 3
Const m_DATE_ADDED = 4
Const m_PROJECT_DESCRIPTION = 5
Const m_PROJECT_FORMAT = 6
Const m_PROJECT_RENDERED = 7
Const m_PROJECT_MEDIA = 8
Const m_PROJECT_STATUS = 9
Const m_PROJECT_SIZE = 10
Const m_PROJECT_SOFTWARE_NAME = 11
Const m_PROJECT_SOFTWARE_ID = 12
Const m_PROJECT_ICON = 13
Const m_PROJECT_USER_ID = 14
Const m_PROJECT_CREATOR = 15
Const m_PROJECT_RANK = 16
Const m_PROJECT_VERSION = 17
Const m_PROJECT_DOWNLOADS = 18
Const m_PROJECT_PLUGINS = 19		' generated from join
Const m_PROJECT_CATS = 20			' generated from join

Const m_NO_VERSION = 5			' from database

'-------------------------------------------------------------------------
'	Name: 		kbProjectData class
'	Purpose: 	methods of getting file data
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'	4/25/03		JEA		Add site condition
'	4/30/03		JEA		Retrieve download count
'	5/30/03		JEA		Retrieve software icon
'-------------------------------------------------------------------------
Class kbProjectData
	Private m_oData
	Private m_sBaseSQL

	Private Sub Class_Initialize()
		Set m_oData = New kbDataAccess
		m_sBaseSQL = "SELECT F.lProjectID, F.vsFileName, F.vsPath, F.vsFriendlyName, F.dtVersionDate, " _
			& "F.vsDescription, F.lFormatID, F.vsRenderedURL, F.vsRequiredMediaURL, F.lStatusID, " _
			& "F.lFileSize, IIf(SV.lVersionID IS NULL, '', S.vsSoftwareName + ' ' + SV.vsVersionText), " _
			& "SV.lVersionID, SV.vsIcon, U.lUserID, IIf(U.vsScreenName IS NULL, " _
			& "U.vsFirstName + ' ' + U.vsLastName, U.vsScreenName), CR.fRank, F.lVersionCount, F.lDownloads " _
			& "FROM ((((tblProjects F INNER JOIN tblUsers U ON U.lUserID = F.lUserID) " _
			& "INNER JOIN (SELECT lItemID FROM tblItemSites WHERE lItemTypeID = " _
			& g_ITEM_PROJECT & " AND lSiteID = " & GetSessionValue(g_USER_SITE) _
			& ") tIS ON tIS.lItemID = F.lProjectID) " _
			& "LEFT JOIN tblSoftwareVersions SV ON SV.lVersionID = F.lSoftwareVersionID) " _
			& "LEFT JOIN tblSoftware S ON S.lSoftwareID = SV.lSoftwareID) " _
			& "LEFT JOIN (SELECT lItemID, fRank FROM tblComputedRankings WHERE lItemTypeID = " _
			& g_ITEM_PROJECT & ") CR ON CR.lItemID = F.lProjectID "
	End Sub
	
	Private Sub Class_Terminate()
		Set m_oData = nothing
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		GetPending()
	'	Purpose: 	get array of uploaded files
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	4/30/03		JEA		Sub version date
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
	'	Purpose: 	get array of files visible to public
	'Modifications:
	'	Date:		Name:	Description:
	'	12/23/02	JEA		Creation
	'	1/9/03		JEA		create child array with new method
	'	4/30/03		JEA		make more version specific
	'-------------------------------------------------------------------------
	Public Function GetPublic(ByVal v_aFilter)
		dim sFilesSQL
		dim sPluginsSQL
		dim sCatSQL
		dim aData
		dim sKey
		
		sKey = MakeKey(g_ITEM_PROJECT, v_aFilter)
		aData = Application(sKey)
		If Not IsArray(aData) Then
			sFilesSQL = m_sBaseSQL & "WHERE F.lStatusID = " & g_STATUS_APPROVED & " "
			If v_aFilter(g_FILTER_AUTHOR) <> 0 Then
				sFilesSQL = sFilesSQL & "AND U.lUserID = " & v_aFilter(g_FILTER_AUTHOR) & " "
			End If
			If v_aFilter(g_FILTER_SOFTWARE) <> 0 Then
				'sFilesSQL = sFilesSQL & "AND SV.lVersionID IN (" & v_aFilter(g_FILTER_SOFTWARE) & "," & m_NO_VERSION & ") "
				sFilesSQL = sFilesSQL & "AND SV.lVersionID = " & v_aFilter(g_FILTER_SOFTWARE) & " "
			End If
			sFilesSQL = sFilesSQL & MakeSortSQL(v_aFilter(g_FILTER_SORT))
			sPluginsSQL = "SELECT P.lPluginID, P.vsPluginName, PP.vsPluginPackageName, " _
				& "PP.vsPluginPackageURL, FP.lProjectID FROM (tblProjectPlugins FP INNER JOIN tblPlugins P " _
				& "ON P.lPluginID = FP.lPluginID) INNER JOIN tblPluginPackages PP " _
				& "ON PP.lPluginPackageID = P.lPluginPackageID"
			sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
				& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
				& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_PROJECT
			aData = m_oData.GetArray(sFilesSQL)
			aData = JoinArray(aData, m_PROJECT_ID, m_oData.GetArray(sPluginsSQL), g_PLUGIN_ITEM_ID)
			aData = JoinArray(aData, m_PROJECT_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
			If v_aFilter(g_FILTER_CATEGORY) <> 0 Then
				aData = FilterArray(aData, m_PROJECT_CATS, g_CAT_ID, v_aFilter(g_FILTER_CATEGORY))
			End If
			Application.Lock
			Application(sKey) = aData
			Application.Unlock
		End If
		GetPublic = aData
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		MakeSortSQL()
	'	Purpose: 	generate SQL for sorting files
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	4/29/03		JEA		Use version date instead of approve date
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
	'	12/28/02	JEA		Creation
	'	4/30/03		JEA		Track download count
	'	7/21/04		JEA		Separate file system functionality
	'-------------------------------------------------------------------------
	Public Function GetDownloadURL(ByVal v_lProjectID)
		Const ITEM_NAME = 0
		Const ITEM_PATH = 1
		Const ITEM_USER = 2
		dim oFileSys
		dim sQuery
		dim aData
		
		sQuery = "SELECT vsFileName, vsPath, lUserID FROM tblProjects WHERE lProjectID = " & v_lProjectID
		if Not g_bAdmin Then sQuery = sQuery & " AND lStatusID = " & g_STATUS_APPROVED
		
		aData = m_oData.GetArray(sQuery)
		
		If IsArray(aData) Then
			' increment download count
			If CStr(aData(ITEM_USER, 0)) <> CStr(GetSessionValue(g_USER_ID)) Then
				sQuery = "UPDATE tblProjects SET lDownloads = lDownloads + 1 WHERE lProjectID = " & v_lProjectID
				Call m_oData.ExecuteOnly(sQuery)
				Call UpdateCacheCount(v_lProjectID)
			End If
		
			Set oFileSys = New kbProjectsystem
			GetDownloadURL = oFileSys.CreateDownloadURL(v_lProjectID, g_ITEM_PROJECT, aData(ITEM_PATH, 0), aData(ITEM_NAME, 0))
			Set oFileSys = Nothing
		End If
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		UpdateCacheCount()
	'	Purpose: 	update download count in cache so entire cache doesn't have to be refreshed
	'Modifications:
	'	Date:		Name:	Description:
	'	4/30/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub UpdateCacheCount(ByVal v_lItemID)
		dim aData
		dim sPattern
		dim bUpdated
		dim sKey
		dim x

		sPattern = GetSessionValue(g_USER_SITE) & "-" & g_ITEM_PROJECT & "_"
		Application.Lock
		For Each sKey In Application.Contents
			bUpdated = false
			If Left(sKey, 4) = sPattern Then
				' this is a file cache
				aData = Application(sKey)
				If IsArray(aData) Then
					for x = 0 to UBound(aData, 2)
						if CStr(aData(m_PROJECT_ID, x)) = CStr(v_lItemID) then
							aData(m_PROJECT_DOWNLOADS, x) = aData(m_PROJECT_DOWNLOADS, x) + 1
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
	'	Purpose: 	update file status
	'Modifications:
	'	Date:		Name:	Description:
	'	12/28/02	JEA		Creation
	'	4/29/03		JEA		Insert version info
	'-------------------------------------------------------------------------
	Public Sub ApprovePending(ByVal v_lProjectID)
		dim sQuery
		dim oMail
		
		sQuery = "UPDATE tblProjects SET " _
			& "lStatusID = " & g_STATUS_APPROVED & ", " _
			& "dtApproveDate = " & g_sSQL_DATE_DELIMIT & Date() & g_sSQL_DATE_DELIMIT & ", " _
			& "dtVersionDate = " & g_sSQL_DATE_DELIMIT & Date() & g_sSQL_DATE_DELIMIT & ", " _
			& "lVersionCount = 1 " _
			& "WHERE lProjectID = " & v_lProjectID
		Call m_oData.ExecuteOnly(sQuery)
		Call m_oData.LogActivity(g_ACT_APPROVE_UPLOAD, v_lProjectID, "", "", "", "", "")
		Set oMail = New kbMail
		Call oMail.SendApprovalEmail(v_lProjectID, g_ITEM_PROJECT)
		Set oMail = Nothing
		Call ClearCache(g_ITEM_PROJECT)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		DenyPending()
	'	Purpose: 	update file status and move file
	'Modifications:
	'	Date:		Name:	Description:
	'	12/28/02	JEA		Creation
	'	1/1/03		JEA		Move functionality to common method
	'-------------------------------------------------------------------------
	Public Sub DenyPending(ByVal v_lProjectID)
		Call Disable(v_lProjectID, g_STATUS_REJECTED, g_ACT_DENY_UPLOAD, g_sDENY_DIR, false)
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		Delete()
	'	Purpose: 	get array for single file
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'	1/5/03		JEA		Remove from any contests
	'-------------------------------------------------------------------------
	Public Sub Delete(ByVal v_lProjectID)
		dim oContest
		Call Disable(v_lProjectID, g_STATUS_DISABLED, g_ACT_DELETE_FILE, g_sDELETE_DIR, true)
		Set oContest = New kbContest
		Call oContest.RemoveItemGlobally(v_lProjectID, g_ITEM_PROJECT)
		Set oContest = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Disable()
	'	Purpose: 	update file status and move file
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'	4/25/03		JEA		Remove from site table
	'	7/21/04		JEA		Separate file system functionality
	'-------------------------------------------------------------------------
	Private Sub Disable(ByVal v_lProjectID, ByVal v_lNewStatusID, ByVal v_lActionID, _
		ByVal v_sBackupPath, ByVal v_bPurgeCache)
		
		Const ITEM_NAME = 0
		Const ITEM_PATH = 1
		dim sQuery
		dim oFileSys
		dim aData
		
		sQuery = "SELECT vsFileName, vsPath FROM tblProjects WHERE lProjectID = " & v_lProjectID
		aData = m_oData.GetArray(sQuery)

		Set oFileSys = New kbProjectsystem
		Call oFileSys.DeleteFromDisk(aData(ITEM_NAME, 0), aData(ITEM_PATH, 0), v_sBackupPath)
		Set oFileSys = Nothing

		sQuery = "UPDATE tblProjects SET lStatusID = " & v_lNewStatusID _
			& ", vsPath = '" & g_sFILES_DIR & "/" & v_sBackupPath & "'" _
			& " WHERE lProjectID = " & v_lProjectID
		Call m_oData.ExecuteOnly(sQuery)
		sQuery = "DELETE FROM tblItemSites WHERE lItemTypeID = " & g_ITEM_PROJECT _
			& " AND lItemID = " & v_lProjectID
		Call m_oData.ExecuteOnly(sQuery)
		Call m_oData.LogActivity(v_lActionID, v_lProjectID, "", "", "", "", "")
		If v_bPurgeCache Then Call ClearCache(g_ITEM_PROJECT)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetItem()
	'	Purpose: 	get array for single file
	'Modifications:
	'	Date:		Name:	Description:
	'	12/31/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Function GetItem(ByVal v_lProjectID, ByVal v_bOwnerOnly)
		dim sFilesSQL
		dim sPluginsSQL
		dim sCatSQL
		dim aData
		
		sFilesSQL = m_sBaseSQL & "WHERE F.lProjectID = " & v_lProjectID
		If v_bOwnerOnly And Not g_bAdmin Then
			sFilesSQL = sFilesSQL & " AND F.lUserID = " & GetSessionValue(g_USER_ID)
		End If
		sPluginsSQL = "SELECT P.lPluginID, P.vsPluginName, PP.vsPluginPackageName, " _
			& "PP.vsPluginPackageURL, FP.lProjectID FROM (tblProjectPlugins FP INNER JOIN tblPlugins P " _
			& "ON P.lPluginID = FP.lPluginID) INNER JOIN tblPluginPackages PP " _
			& "ON PP.lPluginPackageID = P.lPluginPackageID WHERE FP.lProjectID = " & v_lProjectID
		sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
			& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
			& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_PROJECT _
			& " AND IC.lItemID = " & v_lProjectID
		aData = m_oData.GetArray(sFilesSQL)
		aData = JoinArray(aData, m_PROJECT_ID, m_oData.GetArray(sPluginsSQL), g_PLUGIN_ITEM_ID)
		aData = JoinArray(aData, m_PROJECT_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
		GetItem = aData
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		SavePlugins()
	'	Purpose: 	save file plugin data to database
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub SavePlugins(ByVal v_lProjectID, ByVal v_sPlugins)
		dim aPlugins
		dim sQuery
		dim x
		
		' delete any old plugins listed
		sQuery = "DELETE FROM tblProjectPlugins WHERE lProjectID = " & v_lProjectID
		Call m_oData.ExecuteOnly(sQuery)
		
		If v_sPlugins <> "" Then
			' insert new plugins
			sQuery = "INSERT INTO tblProjectPlugins (lProjectID, lPluginID) VALUES (" _
				& v_lProjectID & ", "
			aPlugins = Split(v_sPlugins, ",")
			for x = 0 to UBound(aPlugins)
				Call m_oData.ExecuteOnly(sQuery & Trim(aPlugins(x)) & ")")
			next
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		SaveCategories()
	'	Purpose: 	save file plugin data to database
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub SaveCategories(ByVal v_lProjectID, ByVal v_sCategories)
		dim aCategories
		dim sQuery
		dim x
		
		If v_sCategories <> "" Then
			' delete any old plugins listed
			sQuery = "DELETE FROM tblItemCategories WHERE lItemID = " & v_lProjectID _
				& " AND lItemTypeID = " & g_ITEM_PROJECT
			Call m_oData.ExecuteOnly(sQuery)
			
			' insert new plugins
			sQuery = "INSERT INTO tblItemCategories (lItemID, lItemTypeID, lCategoryID) " _
				& "VALUES (" & v_lProjectID & ", " & g_ITEM_PROJECT & ", "
			aCategories = Split(v_sCategories, ",")
			for x = 0 to UBound(aCategories)
				Call m_oData.ExecuteOnly(sQuery & Trim(aCategories(x)) & ")")
			next
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Save()
	'	Purpose: 	save uploaded file to disk and database
	'Modifications:
	'	Date:		Name:	Description:
	'	1/5/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub Save(ByVal v_sFileField)
		dim oForm
		dim oFileSys
		dim sMessage
		dim sFileName
		dim lFileSize
		dim lProjectID
		dim sURL
		
		v_sFileField = LCase(v_sFileField)
		Set oForm = New kbForm
		Set oFileSys = New kbProjectsystem
		Call oForm.ParseFields()
		lProjectID = oForm.Field("fldProjectID")
		sURL = ReplaceNull(oForm.Field("fldURL"), "kb_projects.asp")
		
		Select Case MakeNumber(oForm.Field("fldAction"))
			Case g_ACT_FILE_ADD
				sURL = "kb_project-submit.asp"
				If oForm.File.Exists(v_sFileField) Then
					sMessage = oFileSys.SaveToDisk(oForm, sFileName, lFileSize, v_sFileField)
					If sMessage = "" Then
						sMessage = g_sMSG_AFTER_UPLOAD
						lProjectID = Insert(sFileName, lFileSize, oForm)
						Call m_oData.LogActivity(g_ACT_FILE_UPLOAD, lProjectID, "", "", "", "", "")
					End If
				Else
					sMessage = "Unable to read file data"
				End If
		
			Case g_ACT_FILE_UPDATE
				If oForm.File.Exists(v_sFileField) Then
					' delete old file and add new
					Call oFileSys.DeleteFromDisk(lProjectID, g_sDELETE_DIR, "")
					sMessage = oFileSys.SaveToDisk(oForm, sFileName, lFileSize, v_sFileField)
				End If
				Call Update(lProjectID, sFileName, lFileSize, oForm)
				
			Case g_ACT_FILE_DELETE
				Call Delete(lProjectID)
		End Select
		
		Set oForm = Nothing
		Call SetSessionValue(g_USER_MSG, sMessage)
		response.redirect sURL
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Insert()
	'	Purpose: 	save file info to database
	'	Return: 	number
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	12/28/02	JEA		Return file ID for logging
	'	12/31/02	JEA		Save plugin list
	'	1/8/03		JEA		Save category list
	'	4/25/03		JEA		Save site
	'	4/30/03		JEA		Insert counts
	'-------------------------------------------------------------------------
	Private Function Insert(ByVal v_sFileName, ByVal v_lFileSize, ByRef r_oForm)
		dim lProjectID
		dim oRS
		dim oFileData
		dim sQuery

		sQuery = "INSERT INTO tblItemSites (lSiteID, lItemTypeID, lItemID) VALUES (" _
			& GetSessionValue(g_USER_SITE) & ", " _
			& g_ITEM_PROJECT
		Call m_oData.BeginTrans()
		Set oRS = Server.CreateObject("ADODB.Recordset")
		With oRS
			.Open "tblProjects", m_oData.Connection, adOpenStatic, adLockOptimistic, adCmdTable
			.AddNew
			.Fields("lUserID") = GetSessionValue(g_USER_ID)
			.Fields("vsPath") = g_sFILES_DIR
			.Fields("vsFileName") = v_sFileName
			.Fields("dtSubmitDate") = Now()
			.Fields("vsFriendlyName") = r_oForm.Field("fldFriendlyName")
			.Fields("vsDescription") = r_oForm.Field("fldDescription")
			.Fields("lSoftwareVersionID") = r_oForm.Field("fldVersion")
			.Fields("lFormatID") = r_oForm.Field("fldFormat")
			.Fields("vsRenderedURL") = Replace(r_oForm.Field("fldRendered"), "http://", "")
			.Fields("vsRequiredMediaURL") = Replace(r_oForm.Field("fldMedia"), "http://", "")
			.Fields("lFileSize") = v_lFileSize
			.Fields("lVersionCount") = 1
			.Fields("lDownloads") = 0
			.Fields("lStatusID") = g_STATUS_PENDING
			.Update
			lProjectID = .Fields("lProjectID")
			.Close
		End With
		Set oRS = nothing
		Call SavePlugins(lProjectID, Trim(r_oForm.Field("fldPlugins")))
		Call SaveCategories(lProjectID, Trim(r_oForm.Field("fldCategories")))
		Call m_oData.ExecuteOnly(sQuery & ", " & lProjectID & ")")
		Call m_oData.CommitTrans()
		Insert = lProjectID
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		Update()
	'	Purpose: 	save file data to database
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'	1/9/03		JEA		Save categorization, use form object
	'	4/29/03		JEa		Insert version information
	'-------------------------------------------------------------------------
	Private Sub Update(ByVal v_lProjectID, ByVal v_sFileName, ByVal v_lFileSize, ByRef r_oForm)
		dim sQuery
		
		With r_oForm
			sQuery = "UPDATE tblProjects SET " _
				& "vsFriendlyName = '" & CleanForSQL(.Field("fldFriendlyName")) & "', " _
				& "vsDescription = '" & CleanForSQL(.Field("fldDescription")) & "', " _
				& "lFormatID = " & .Field("fldFormat") & ", " _
				& "lSoftwareVersionID = " & .Field("fldVersion") & ", " _
				& "vsRenderedURL = '" & Replace(.Field("fldRendered"), "http://", "") & "', " _
				& "vsRequiredMediaURL = '" & Replace(.Field("fldMedia"), "http://", "") & "'"
			If v_sFileName <> "" Then
				' update file and version info
				sQuery = sQuery _
					& ", vsFileName = '" & v_sFileName & "', " _
					& "lFileSize = " & MakeNumber(v_lFileSize) & ", " _
					& "dtVersionDate = " & g_sSQL_DATE_DELIMIT & Date() & g_sSQL_DATE_DELIMIT & ", " _
					& "lVersionCount = lVersionCount + 1"
			End If
			sQuery = sQuery & " WHERE lProjectID = " & v_lProjectID
				
			Call m_oData.BeginTrans()
			Call m_oData.ExecuteOnly(sQuery)
			Call SavePlugins(v_lProjectID, Trim(.Field("fldPlugins")))
			Call SaveCategories(v_lProjectID, Trim(.Field("fldCategories")))
			Call m_oData.CommitTrans()
		End With
		
		Call m_oData.LogActivity(g_ACT_EDIT_FILE_ENTRY, v_lProjectID, "", "", "", "", "")
		Call SetSessionValue(g_USER_MSG, g_sMSG_FILE_EDIT)
		Call ClearCache(g_ITEM_PROJECT)
	End Sub
End Class
%>