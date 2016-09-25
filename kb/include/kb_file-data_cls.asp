<%
Const m_FILE_ID = 0
Const m_FILE_NAME = 1
Const m_FILE_PATH = 2
Const m_FRIENDLY_NAME = 3
Const m_DATE_ADDED = 4
Const m_FILE_DESCRIPTION = 5
Const m_FILE_FORMAT = 6
Const m_FILE_RENDERED = 7
Const m_FILE_MEDIA = 8
Const m_FILE_STATUS = 9
Const m_FILE_SIZE = 10
Const m_FILE_SOFTWARE_NAME = 11
Const m_FILE_SOFTWARE_ID = 12
Const m_FILE_USER_ID = 13
Const m_FILE_CREATOR = 14
Const m_FILE_RANK = 15
Const m_FILE_PLUGINS = 16		' generated from join
Const m_FILE_CATS = 17			' generated from join

Const m_NO_VERSION = 5			' from database

'-------------------------------------------------------------------------
'	Name: 		kbFileData class
'	Purpose: 	methods of getting file data
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'	4/25/03		JEA		Add site condition
'-------------------------------------------------------------------------
Class kbFileData
	Private m_oData
	Private m_sBaseSQL

	Private Sub Class_Initialize()
		Set m_oData = New kbDataAccess
		m_sBaseSQL = "SELECT F.lFileID, F.vsFileName, F.vsPath, F.vsFriendlyName, F.dtApproveDate, " _
			& "F.vsDescription, F.lFormatID, F.vsRenderedURL, F.vsRequiredMediaURL, F.lStatusID, " _
			& "F.lFileSize, IIf(SV.lVersionID IS NULL, '', S.vsSoftwareName + ' ' + SV.vsVersionText), " _
			& "SV.lVersionID, U.lUserID, IIf(U.vsScreenName IS NULL, " _
			& "U.vsFirstName + ' ' + U.vsLastName, U.vsScreenName), CR.fRank " _
			& "FROM ((((tblFiles F INNER JOIN tblUsers U ON U.lUserID = F.lUserID) " _
			& "INNER JOIN (SELECT lItemID FROM tblItemSites WHERE lItemTypeID = " _
			& g_ITEM_FILE & " AND lSiteID = " & GetSessionValue(g_USER_SITE) _
			& ") tIS ON tIS.lItemID = F.lFileID) " _
			& "LEFT JOIN tblSoftwareVersions SV ON SV.lVersionID = F.lSoftwareVersionID) " _
			& "LEFT JOIN tblSoftware S ON S.lSoftwareID = SV.lSoftwareID) " _
			& "LEFT JOIN (SELECT lItemID, fRank FROM tblComputedRankings WHERE lItemTypeID = " _
			& g_ITEM_FILE & ") CR ON CR.lItemID = F.lFileID "
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
	'-------------------------------------------------------------------------
	Public Function GetPending()
		dim sQuery
		dim aData
		sQuery = Replace(m_sBaseSQL, "F.dtApproveDate", "F.dtSubmitDate") _
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
	'-------------------------------------------------------------------------
	Public Function GetPublic(ByVal v_aFilter)
		dim sFilesSQL
		dim sPluginsSQL
		dim sCatSQL
		dim aData
		dim sKey
		
		sKey = MakeKey(g_ITEM_FILE, v_aFilter)
		aData = Application(sKey)
		If Not IsArray(aData) Then
			sFilesSQL = m_sBaseSQL & "WHERE F.lStatusID = " & g_STATUS_APPROVED & " "
			If v_aFilter(g_FILTER_AUTHOR) <> 0 Then
				sFilesSQL = sFilesSQL & "AND U.lUserID = " & v_aFilter(g_FILTER_AUTHOR) & " "
			End If
			If v_aFilter(g_FILTER_SOFTWARE) <> 0 Then
				sFilesSQL = sFilesSQL & "AND SV.lVersionID IN (" & v_aFilter(g_FILTER_SOFTWARE) & "," & m_NO_VERSION & ") "
			End If
			sFilesSQL = sFilesSQL & MakeSortSQL(v_aFilter(g_FILTER_SORT))
			sPluginsSQL = "SELECT P.lPluginID, P.vsPluginName, PP.vsPluginPackageName, " _
				& "PP.vsPluginPackageURL, FP.lFileID FROM (tblFilePlugins FP INNER JOIN tblPlugins P " _
				& "ON P.lPluginID = FP.lPluginID) INNER JOIN tblPluginPackages PP " _
				& "ON PP.lPluginPackageID = P.lPluginPackageID"
			sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
				& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
				& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_FILE
			aData = m_oData.GetArray(sFilesSQL)
			aData = JoinArray(aData, m_FILE_ID, m_oData.GetArray(sPluginsSQL), g_PLUGIN_ITEM_ID)
			aData = JoinArray(aData, m_FILE_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
			If v_aFilter(g_FILTER_CATEGORY) <> 0 Then
				aData = FilterArray(aData, m_FILE_CATS, g_CAT_ID, v_aFilter(g_FILTER_CATEGORY))
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
				sSQL = sSQL & "F.dtApproveDate, F.vsFriendlyName"
			case g_SORT_DATE_DESC
				sSQL = sSQL & "F.dtApproveDate DESC, F.vsFriendlyName"
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
	'-------------------------------------------------------------------------
	Public Function GetDownloadURL(ByVal v_lFileID)
		Const FILE_NAME = 0
		Const FILE_PATH = 1
		dim sFile
		dim sFolder
		dim sFolderPath
		dim oFileSys
		dim sQuery
		dim aData
		
		sQuery = "SELECT vsFileName, vsPath FROM tblFiles WHERE lFileID = " & v_lFileID
		if Session(g_sSESSION)(g_USER_TYPE) <> g_USER_ADMIN Then sQuery = sQuery & " AND lStatusID = " & g_STATUS_APPROVED
		
		aData = m_oData.GetArray(sQuery)
		
		If IsArray(aData) Then
			Set oFileSys = Server.CreateObject(g_sFILE_SYSTEM_OBJECT)
			sFolder = "./" & g_sDOWNLOAD_DIR & "/" & oFileSys.GetTempName()
			sFolderPath = Server.Mappath(sFolder)
			oFileSys.CreateFolder(sFolderPath)
	
			' log folder creation
			sQuery = "INSERT INTO tblTemporaryFolders (vsFolderName, dtDateCreated, lCreatedForUserID, " _
				& "lCreatedForFileID) VALUES ('" _
				& sFolderPath & "', '" _
				& Now() & "', " _
				& Session(g_sSESSION)(g_USER_ID) & ", " _
				& v_lFileID & ")"
			Call m_oData.ExecuteOnly(sQuery)
	
			' copy requested file to temp folder
			sFile = Server.Mappath("./" & aData(FILE_PATH, 0) & "/" & aData(FILE_NAME, 0))
			Call oFileSys.CopyFile(sFile, sFolderPath & "\", true)
			Call DeleteOldTempFolders(oFileSys)
			Set oFileSys = nothing
			Call m_oData.LogActivity(g_ACT_FILE_DOWNLOAD, v_lFileID, "", "", "", "")
			GetDownloadURL = sFolder & "/" & aData(FILE_NAME, 0)
		End If
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		DeleteOldTempFolders()
	'	Purpose: 	remove previously created temp download folders
	'Modifications:
	'	Date:		Name:	Description:
	'	12/28/02	JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub DeleteOldTempFolders(ByRef r_oFileSys)
		Const FOLDER_ID = 0
		Const FOLDER_NAME = 1
		dim sFolderList
		dim bNewFileSys
		dim sFolderPath
		dim sQuery
		dim aData
		dim x
		
		sQuery = "SELECT lFolderID, vsFolderName FROM tblTemporaryFolders WHERE dtDateCreated < " _
			& g_sSQL_DATE_DELIMIT & DateAdd("h", g_PURGE_TEMP_AFTER * -1, Now()) & g_sSQL_DATE_DELIMIT
		sFolderList = ""	
		aData = m_oData.GetArray(sQuery)
		If IsArray(aData) Then
			bNewFileSys = GetObject(r_oFileSys, g_sFILE_SYSTEM_OBJECT)
			with r_oFileSys
				for x = 0 to UBound(aData, 2)
					sFolderList = sFolderList & aData(FOLDER_ID, x) & ","
					sFolderPath = aData(FOLDER_NAME, x)
					If .FolderExists(sFolderPath) Then .DeleteFolder(sFolderPath)
				next
			end with
			if bNewFileSys then Set r_oFileSys = nothing
			sFolderList = Left(sFolderList, Len(sFolderList) - 1)
			sQuery = "DELETE FROM tblTemporaryFolders WHERE lFolderID IN (" & sFolderList & ")"
			Call m_oData.ExecuteOnly(sQuery)
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		ApprovePending()
	'	Purpose: 	update file status
	'Modifications:
	'	Date:		Name:	Description:
	'	12/28/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub ApprovePending(ByVal v_lFileID)
		dim sQuery
		dim oMail
		
		sQuery = "UPDATE tblFiles SET lStatusID = " & g_STATUS_APPROVED _
			& ", dtApproveDate = " & g_sSQL_DATE_DELIMIT & Date() & g_sSQL_DATE_DELIMIT _
			& " WHERE lFileID = " & v_lFileID
		Call m_oData.ExecuteOnly(sQuery)
		Call m_oData.LogActivity(g_ACT_APPROVE_UPLOAD, v_lFileID, "", "", "", "")
		Set oMail = New kbMail
		Call oMail.SendApprovalEmail(v_lFileID)
		Set oMail = Nothing
		Call ClearCache(g_ITEM_FILE)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		DenyPending()
	'	Purpose: 	update file status and move file
	'Modifications:
	'	Date:		Name:	Description:
	'	12/28/02	JEA		Creation
	'	1/1/03		JEA		Move functionality to common method
	'-------------------------------------------------------------------------
	Public Sub DenyPending(ByVal v_lFileID)
		Call Disable(v_lFileID, g_STATUS_REJECTED, g_ACT_DENY_UPLOAD, g_sDENY_DIR, false)
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		Delete()
	'	Purpose: 	get array for single file
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'	1/5/03		JEA		Remove from any contests
	'-------------------------------------------------------------------------
	Public Sub Delete(ByVal v_lFileID)
		dim oContest
		Call Disable(v_lFileID, g_STATUS_DISABLED, g_ACT_DELETE_FILE, g_sDELETE_DIR, true)
		Set oContest = New kbContest
		Call oContest.RemoveItemGlobally(v_lFileID, g_ITEM_FILE)
		Set oContest = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Disable()
	'	Purpose: 	update file status and move file
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'	4/25/03		JEA		Remove from site table
	'-------------------------------------------------------------------------
	Private Sub Disable(ByVal v_lFileID, ByVal v_lNewStatusID, ByVal v_lActionID, _
		ByVal v_sBackupPath, ByVal v_bPurgeCache)
		dim sQuery
		Call DeleteFromDisk(v_lFileID, v_sBackupPath)
		sQuery = "UPDATE tblFiles SET lStatusID = " & v_lNewStatusID _
			& ", vsPath = '" & g_sFILES_DIR & "/" & v_sBackupPath & "'" _
			& " WHERE lFileID = " & v_lFileID
		Call m_oData.ExecuteOnly(sQuery)
		sQuery = "DELETE FROM tblItemSites WHERE lItemTypeID = " & g_ITEM_FILE _
			& " AND lItemID = " & v_lFileID
		Call m_oData.ExecuteOnly(sQuery)
		Call m_oData.LogActivity(v_lActionID, v_lFileID, "", "", "", "")
		If v_bPurgeCache Then Call ClearCache(g_ITEM_FILE)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetItem()
	'	Purpose: 	get array for single file
	'Modifications:
	'	Date:		Name:	Description:
	'	12/31/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Function GetItem(ByVal v_lFileID, ByVal v_bOwnerOnly)
		dim sFilesSQL
		dim sPluginsSQL
		dim sCatSQL
		dim aData
		
		sFilesSQL = m_sBaseSQL & "WHERE F.lFileID = " & v_lFileID
		If v_bOwnerOnly And Not g_bAdmin Then
			sFilesSQL = sFilesSQL & " AND F.lUserID = " & GetSessionValue(g_USER_ID)
		End If
		sPluginsSQL = "SELECT P.lPluginID, P.vsPluginName, PP.vsPluginPackageName, " _
			& "PP.vsPluginPackageURL, FP.lFileID FROM (tblFilePlugins FP INNER JOIN tblPlugins P " _
			& "ON P.lPluginID = FP.lPluginID) INNER JOIN tblPluginPackages PP " _
			& "ON PP.lPluginPackageID = P.lPluginPackageID WHERE FP.lFileID = " & v_lFileID
		sCatSQL = "SELECT C.lCategoryID, C.vsCategoryName, IC.lItemID " _
			& "FROM tblCategories C INNER JOIN tblItemCategories IC " _
			& "ON C.lCategoryID = IC.lCategoryID WHERE IC.lItemTypeID = " & g_ITEM_FILE _
			& " AND IC.lItemID = " & v_lFileID
		aData = m_oData.GetArray(sFilesSQL)
		aData = JoinArray(aData, m_FILE_ID, m_oData.GetArray(sPluginsSQL), g_PLUGIN_ITEM_ID)
		aData = JoinArray(aData, m_FILE_ID, m_oData.GetArray(sCatSQL), g_CAT_ITEM_ID)
		GetItem = aData
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		SavePlugins()
	'	Purpose: 	save file plugin data to database
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub SavePlugins(ByVal v_lFileID, ByVal v_sPlugins)
		dim aPlugins
		dim sQuery
		dim x
		
		' delete any old plugins listed
		sQuery = "DELETE FROM tblFilePlugins WHERE lFileID = " & v_lFileID
		Call m_oData.ExecuteOnly(sQuery)
		
		If v_sPlugins <> "" Then
			' insert new plugins
			sQuery = "INSERT INTO tblFilePlugins (lFileID, lPluginID) VALUES (" _
				& v_lFileID & ", "
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
	Public Sub SaveCategories(ByVal v_lFileID, ByVal v_sCategories)
		dim aCategories
		dim sQuery
		dim x
		
		If v_sCategories <> "" Then
			' delete any old plugins listed
			sQuery = "DELETE FROM tblItemCategories WHERE lItemID = " & v_lFileID _
				& " AND lItemTypeID = " & g_ITEM_FILE
			Call m_oData.ExecuteOnly(sQuery)
			
			' insert new plugins
			sQuery = "INSERT INTO tblItemCategories (lItemID, lItemTypeID, lCategoryID) " _
				& "VALUES (" & v_lFileID & ", " & g_ITEM_FILE & ", "
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
		dim sMessage
		dim sFileName
		dim lFileSize
		dim lFileID
		dim sURL
		
		v_sFileField = LCase(v_sFileField)
		Set oForm = New kbForm
		Call oForm.ParseFields()
		lFileID = oForm.Field("fldFileID")
		sURL = ReplaceNull(oForm.Field("fldURL"), "kb_files.asp")
		
		Select Case MakeNumber(oForm.Field("fldAction"))
			Case m_ACT_FILE_ADD
				sURL = "kb_submit-file.asp"
				If oForm.File.Exists(v_sFileField) Then
					sMessage = SaveToDisk(oForm, sFileName, lFileSize, v_sFileField)
					If sMessage = "" Then
						sMessage = g_sMSG_AFTER_UPLOAD
						lFileID = Insert(sFileName, lFileSize, oForm)
						Call m_oData.LogActivity(g_ACT_FILE_UPLOAD, lFileID, "", "", "", "")
					End If
				Else
					sMessage = "Unable to read file data"
				End If
		
			Case m_ACT_FILE_UPDATE
				If oForm.File.Exists(v_sFileField) Then
					' delete old file and add new
					Call DeleteFromDisk(lFileID, g_sDELETE_DIR)
					sMessage = SaveToDisk(oForm, sFileName, lFileSize, v_sFileField)
				End If
				Call Update(lFileID, sFileName, lFileSize, oForm)
				
			Case m_ACT_FILE_DELETE
				Call Delete(lFileID)
		End Select
		
		Set oForm = Nothing
		Call SetSessionValue(g_USER_MSG, sMessage)
		response.redirect sURL
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		SaveToDisk()
	'	Purpose: 	save file to disk
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function SaveToDisk(ByRef r_oForm, ByRef r_sFileName, ByRef r_lFileSize, ByVal v_sFileField)
		dim sMessage
		dim oFile

		Set oFile = r_oForm.File.Item(v_sFileField)
		sMessage = oFile.SaveToDisk(Server.MapPath("./" & g_sFILES_DIR), g_MAX_FILE_KB, "", true, false)
		r_lFileSize = oFile.FileSize
		r_sFileName = oFile.FileName
		Set oFile = Nothing
		r_sFileName = Right(r_sFileName, Len(r_sFileName) - InStrRev(r_sFileName, "/"))

		SaveToDisk = sMessage
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		DeleteFromDisk()
	'	Purpose: 	remove file from disk
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub DeleteFromDisk(ByVal v_lFileID, ByVal v_sBackupPath)
		Const FILE_PATH = 0
		Const FILE_NAME = 1
		dim sQuery
		dim sSourceFile
		dim sTargetFile
		dim oFileSys
		dim aData
		
		sQuery = "SELECT vsPath, vsFileName FROM tblFiles WHERE lFileID = " & v_lFileID
		aData = m_oData.GetArray(sQuery)
		sSourceFile = Server.Mappath(aData(FILE_PATH, 0) & "/" & aData(FILE_NAME, 0))
		Set oFileSys = Server.CreateObject(g_sFILE_SYSTEM_OBJECT)
		If v_sBackupPath <> "" Then
			' make backup before deleting
			sTargetFile = Server.Mappath("./" & g_sFILES_DIR & "/" & v_sBackupPath)
			sTargetFile = sTargetFile & "\" & SafeDate(Date) & "_" & aData(FILE_NAME, 0)
			Call oFileSys.CopyFile(sSourceFile, sTargetFile)
		End If
		Call oFileSys.DeleteFile(sSourceFile, true)
		Set oFileSys = nothing
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
	'-------------------------------------------------------------------------
	Private Function Insert(ByVal v_sFileName, ByVal v_lFileSize, ByRef r_oForm)
		dim lFileID
		dim oRS
		dim oFileData
		dim sQuery

		sQuery = "INSERT INTO tblItemSites (lSiteID, lItemTypeID, lItemID) VALUES (" _
			& GetSessionValue(g_USER_SITE) & ", " _
			& g_ITEM_FILE
		Call m_oData.BeginTrans()
		Set oRS = Server.CreateObject("ADODB.Recordset")
		With oRS
			.Open "tblFiles", m_oData.Connection, adOpenStatic, adLockOptimistic, adCmdTable
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
			.Fields("lStatusID") = g_STATUS_PENDING
			.Update
			lFileID = .Fields("lFileID")
			.Close
		End With
		Set oRS = nothing
		Call SavePlugins(lFileID, Trim(r_oForm.Field("fldPlugins")))
		Call SaveCategories(lFileID, Trim(r_oForm.Field("fldCategories")))
		Call m_oData.ExecuteOnly(sQuery & ", " & lFileID & ")")
		Call m_oData.CommitTrans()
		Insert = lFileID
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		Update()
	'	Purpose: 	save file data to database
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'	1/9/03		JEA		Save categorization, use form object
	'-------------------------------------------------------------------------
	Private Sub Update(ByVal v_lFileID, ByVal v_sFileName, ByVal v_lFileSize, ByRef r_oForm)
		dim sQuery
		
		With r_oForm
			sQuery = "UPDATE tblFiles SET " _
				& "vsFriendlyName = '" & CleanForSQL(.Field("fldFriendlyName")) & "', " _
				& "vsDescription = '" & CleanForSQL(.Field("fldDescription")) & "', " _
				& "lFormatID = " & .Field("fldFormat") & ", " _
				& "lSoftwareVersionID = " & .Field("fldVersion") & ", " _
				& "vsRenderedURL = '" & Replace(.Field("fldRendered"), "http://", "") & "', " _
				& "vsRequiredMediaURL = '" & Replace(.Field("fldMedia"), "http://", "") & "'"
			If v_sFileName <> "" Then
				sQuery = sQuery & ", vsFileName = '" & v_sFileName & "', lFileSize = " & MakeNumber(v_lFileSize)
			End If
			sQuery = sQuery & " WHERE lFileID = " & v_lFileID
				
			Call m_oData.BeginTrans()
			Call m_oData.ExecuteOnly(sQuery)
			Call SavePlugins(v_lFileID, Trim(.Field("fldPlugins")))
			Call SaveCategories(v_lFileID, Trim(.Field("fldCategories")))
			Call m_oData.CommitTrans()
		End With
		
		Call m_oData.LogActivity(g_ACT_EDIT_FILE_ENTRY, v_lFileID, "", "", "", "")
		Call SetSessionValue(g_USER_MSG, g_sMSG_FILE_EDIT)
		Call ClearCache(g_ITEM_FILE)
	End Sub
End Class
%>