<%
'-------------------------------------------------------------------------
'	Name: 		kbDatabase class
'	Purpose: 	methods for managing the database
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'-------------------------------------------------------------------------
Class kbDatabase

	'Private Sub Class_Initialize()
	'End Sub
	
	'Private Sub Class_Terminate()
	'End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		SaveQuery()
	'	Purpose: 	save query
	'Modifications:
	'	Date:		Name:	Description:
	'	4/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub SaveQuery(ByVal v_sQueryName, ByVal v_sSQL)
		dim oData
		dim sQuery
		
		sQuery = "INSERT INTO tblQueries (vsQueryName, vsSQL, lUserID) VALUES ('" _
			& v_sQueryName & "', '" _
			& Replace(v_sSQL, vbCrLf, " ") & "', " _
			& GetSessionValue(g_USER_ID) & ")"
		Set oData = New kbDataAccess
		Call oData.ExecuteOnly(sQuery)
		Set oData = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteQueryArray()
	'	Purpose: 	write JavaScript array of saved queries
	'Modifications:
	'	Date:		Name:	Description:
	'	4/4/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteQueryArray()
		dim oData
		Set oData = New kbDataAccess
		response.write oData.GetJSArray("SELECT lQueryID, vsQueryName, vsSQL FROM tblQueries ORDER BY vsQueryName")
		Set oData = Nothing
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		WriteDatabaseSize()
	'	Purpose: 	write size of database in bytes
	'Modifications:
	'	Date:		Name:	Description:
	'	12/25/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteDatabaseSize()
		dim oFileSys
		dim oFile
		Set oFileSys = Server.CreateObject(g_sFILE_SYSTEM_OBJECT)
		Set oFile = oFileSys.GetFile(Server.Mappath(g_sDB_LOCATION))
		response.write FormatNumber(oFile.Size, 0)
		Set oFile = nothing
		Set oFileSys = nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		CompactDatabase()
	'	Purpose: 	compact the database
	'Modifications:
	'	Date:		Name:	Description:
	'	12/25/02	JEA		Creation
	'-------------------------------------------------------------------------
	Sub CompactDatabase(ByVal v_bBackup)
		dim oFileSys
		dim sCompactFile
		dim sTempFile
		dim sLiveFile
		dim oData
		dim oJet
		
		sCompactFile = Replace(g_sDB_LOCATION, ".mdb", IIf(v_bBackup, "_" & SafeDate(Date) & ".mdb", "_compacted.mdb"))
		sCompactFile = Server.Mappath(sCompactFile)
		sLiveFile = Server.Mappath(g_sDB_LOCATION)
		Set oFileSys = Server.CreateObject(g_sFILE_SYSTEM_OBJECT)
		If oFileSys.FileExists(sCompactFile) Then Call oFileSys.DeleteFile(sCompactFile)
		Set oJet = CreateObject("JRO.JetEngine")
		Call oJet.CompactDatabase(g_sDB_CONNECT & sLiveFile, g_sDB_CONNECT & sCompactFile)
		Set oJet = nothing
		If Not v_bBackup Then
			' replace live db with compacted file
			sTempFile = Server.Mappath(Replace(g_sDB_LOCATION, ".mdb", "_temp.mdb"))
			Call oFileSys.CopyFile(sLiveFile, sTempFile, true)
			Call oFileSys.CopyFile(sCompactFile, sLiveFile, true)
			Call oFileSys.DeleteFile(sCompactFile)	' delete temporary files
			Call oFileSys.DeleteFile(sTempFile)
		End If
		Set oFileSys = nothing
		
		Set oData = New kbDataAccess
		Call oData.LogActivity(IIf(v_bBackup, g_ACT_BACKUP_DATABASE, g_ACT_COMPACT_DATABASE), "", "", "", "", "")
		Set oData = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteLastCompactDate()
	'	Purpose: 	write date when database was last compacted
	'Modifications:
	'	Date:		Name:	Description:
	'	12/25/02	JEA		Creation
	'	4/4/03		JEA		Check for empty results
	'-------------------------------------------------------------------------
	Public Sub WriteQueryResults(ByVal v_sQuery)
		dim oData
		dim oRS
		dim oField
		dim lColumns
		dim x
		
		v_sQuery = Trim(v_sQuery)
		If IsVoid(v_sQuery) Then Exit Sub
		lColumns = 0
		x = 1
		
		Set oData = New kbDataAccess
		Set oRS = oData.RunQuery(v_sQuery)
		Set oData = Nothing
		
		with response
			.write "<table cellspacing='0' cellpadding='0' border='0'><tr>"
			For Each oField in oRS.Fields
				.write "<td class='QueryHead'>" & oField.Name & "</td>"
				lColumns = lColumns + 1
			Next
			If Not oRS.EOF Then
				.write "<tr><td class='QueryCell'>"
				.write oRS.GetString( , , "</td><td class='QueryCell'>", "</td></tr><tr><td class='QueryCell'>", "&nbsp;")
			Else
				.write "<tr><td align='center' style='font-size: 10pt; font-weight: bold; color: #ee4400;' colspan='"
				.write lColumns
				.write "'>no records found</td>"
			End If
			.write "</table>"
		end with
		oRS.Close
		Set oRS = Nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteLastCompactDate()
	'	Purpose: 	write date when database was last compacted
	'Modifications:
	'	Date:		Name:	Description:
	'	12/25/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteLastCompactDate()
		dim sDate
		sDate = GetLastActivityDate(g_ACT_COMPACT_DATABASE)
		with response
			If IsVoid(sDate) then
				.write "has never been compacted"
			else
				.write "was last compacted on "
				.write FormatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), sDate))
			end if
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteLastBackupDate()
	'	Purpose: 	write date when database was last backed up
	'Modifications:
	'	Date:		Name:	Description:
	'	12/25/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteLastBackupDate()
		dim sDate
		sDate = GetLastActivityDate(g_ACT_BACKUP_DATABASE)
		with response
			If IsVoid(sDate) then
				.write "has never been backed up"
			else
				.write "last had a backup on "
				.write formatDate(DateAdd("s", GetSessionValue(g_USER_TIME_SHIFT), sDate))
			end if
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteLastBackupDate()
	'	Purpose: 	write date when database was last backed up
	'Modifications:
	'	Date:		Name:	Description:
	'	12/25/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteBackupList()
		dim oFileSys
		dim oFolder
		dim oFiles
		dim oFile
		dim sFolder
	
		sFolder = Server.Mappath(g_sDB_LOCATION)
		sFolder = Left(sFolder, InStrRev(sFolder, "\") - 1)
		Set oFileSys = Server.CreateObject("Scripting.FileSystemObject")
		Set oFolder = oFileSys.GetFolder(sFolder)
		Set oFiles = oFolder.Files
		If oFiles.Count > 0 Then
			with response
				.write "<select name='fldBackup'>"
				For Each oFile in oFiles
					.write "<option>"
					.write oFile.name
				Next
				.write "</select>"
			end with
		End If
		Set oFile = nothing
		Set oFiles = nothing
		Set oFolder = nothing
		Set oFileSys = nothing
	End Sub
End Class
%>