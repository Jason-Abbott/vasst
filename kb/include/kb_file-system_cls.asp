<%
'-------------------------------------------------------------------------
'	Name: 		kbProjectsystem class
'	Purpose: 	methods for managing file system
'Modifications:
'	Date:		Name:	Description:
'	7/21/04		JEA		Copied from files class
'-------------------------------------------------------------------------
Class kbProjectsystem
	Private Sub Class_Initialize()
	End Sub
	
	Private Sub Class_Terminate()
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		CreateDownloadURL()
	'	Purpose: 	create and write download link
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Function CreateDownloadURL(ByVal v_lItemID, ByVal v_lItemTypeID, _
		ByVal v_sPath, ByVal v_sFileName)
		
		dim sFile
		dim sFolder
		dim sFolderPath
		dim oFileSys
		dim oData
		dim sQuery
		
		Set oFileSys = Server.CreateObject(g_sFILE_SYSTEM_OBJECT)
		sFolder = "./" & g_sDOWNLOAD_DIR & "/" & oFileSys.GetTempName()
		sFolderPath = Server.Mappath(sFolder)
		oFileSys.CreateFolder(sFolderPath)

		' log folder creation
		sQuery = "INSERT INTO tblTemporaryFolders (vsFolderName, dtDateCreated, lCreatedForUserID, " _
			& "lCreatedForItemID, lItemTypeID) VALUES ('" _
			& sFolderPath & "', '" _
			& Now() & "', " _
			& Session(g_sSESSION)(g_USER_ID) & ", " _
			& v_lItemID & ", " _
			& v_lItemTypeID & ")"
			
		Set oData = New kbDataAccess
		Call oData.ExecuteOnly(sQuery)
		
		' copy requested file to temp folder
		sFile = Server.Mappath("./" & v_sPath & "/" & v_sFileName)
		
		If oFileSys.FileExists(sFile) Then Call oFileSys.CopyFile(sFile, sFolderPath & "\", true)
		Call DeleteOldTempFolders(oFileSys)
		Set oFileSys = nothing
		
		Call oData.LogActivity(g_ACT_FILE_DOWNLOAD, v_lItemID, v_lItemTypeID, "", "", "", "")
		Set oData = Nothing
		CreateDownloadURL = sFolder & "/" & v_sFileName
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		DeleteOldTempFolders()
	'	Purpose: 	remove previously created temp download folders
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Private Sub DeleteOldTempFolders(ByRef r_oFileSys)
		Const FOLDER_ID = 0
		Const FOLDER_NAME = 1
		dim sFolderList
		dim bNewFileSys
		dim sFolderPath
		dim sQuery
		dim aData
		dim oData
		dim x
		
		sQuery = "SELECT lFolderID, vsFolderName FROM tblTemporaryFolders WHERE dtDateCreated < " _
			& g_sSQL_DATE_DELIMIT & DateAdd("h", g_PURGE_TEMP_AFTER * -1, Now()) & g_sSQL_DATE_DELIMIT
		sFolderList = ""
		
		Set oData = New kbDataAccess
		aData = oData.GetArray(sQuery)
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
			
			Call oData.ExecuteOnly(sQuery)
		End If
		Set oData = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		SaveToDisk()
	'	Purpose: 	save script to disk
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'-------------------------------------------------------------------------
	Public Function SaveToDisk(ByRef r_oForm, ByRef r_sFileName, ByRef r_lFileSize, ByVal v_sFileField)
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
	'	Purpose: 	remove script from disk
	'Modifications:
	'	Date:		Name:	Description:
	'	7/21/04		JEA		Copied from files class
	'	7/23/04		JEA		Check for existence of file
	'-------------------------------------------------------------------------
	Public Sub DeleteFromDisk(ByVal v_sName, ByVal v_sPath, ByVal v_sBackupPath)
		dim sSourceFile
		dim sTargetFile
		dim oFileSys
		
		sSourceFile = Server.Mappath(v_sPath & "/" & v_sName)
		Set oFileSys = Server.CreateObject(g_sFILE_SYSTEM_OBJECT)
		If oFileSys.FileExists(sSourceFile) Then
			If v_sBackupPath <> "" Then
				' make backup before deleting
				sTargetFile = Server.Mappath("./" & g_sFILES_DIR & "/" & v_sBackupPath)
				sTargetFile = sTargetFile & "\" & SafeDate(Date) & "_" & v_sName
				Call oFileSys.CopyFile(sSourceFile, sTargetFile)
			End If
			Call oFileSys.DeleteFile(sSourceFile, true)
		End If
		Set oFileSys = nothing
	End Sub
End Class
%>