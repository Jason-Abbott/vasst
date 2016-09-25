<%
'-------------------------------------------------------------------------
'	Name: 		kbForm
'	Purpose: 	manually process posted form to extract file uploads
'Modifications:
'	Date:		Name:	Description:
'	12/07/00	Jacob "Beezle" Gilley (avis7@airmail.net)
'	12/28/02	JEA		Improved formatting; object cleanup; added methods
'-------------------------------------------------------------------------
Class kbForm
	Public File
	Private m_oFieldHash

	Private Sub Class_Initialize()
		' hold file and form data in hashes
		Set File = Server.CreateObject("Scripting.Dictionary")
		Set m_oFieldHash = Server.CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate()
		File.RemoveAll() : Set File = Nothing
		m_oFieldHash.RemoveAll() : Set m_oFieldHash = Nothing
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		GetFieldValue()
	'	Purpose: 	get field value
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/31/02	JEA		Creation
	'	1/11/03		JEA		fix URLEncoded strings (could be generalized with char loop)
	'-------------------------------------------------------------------------
	Private Function GetFieldValue(ByVal v_sFieldName)
		dim sValue
		sValue = ""
		v_sFieldName = LCase(v_sFieldName)
		If m_oFieldHash.Exists(v_sFieldName) Then
			sValue = m_oFieldHash.Item(v_sFieldName)
			sValue = Replace(sValue, "%3D", "=")
			sValue = Replace(sValue, "%26", "&")
		End If
		GetFieldValue = sValue
	End Function
	
	Public Property Get Field(ByVal v_sFieldName)
		Field = GetFieldValue(v_sFieldName)
	End Property

	'-------------------------------------------------------------------------
	'	Name: 		ParseFields()
	'	Purpose: 	parse uploaded data
	'Modifications:
	'	Date:		Name:	Description:
	'	12/07/00	JG		Creation
	'	12/31/02	JEA		Fixed error with multi-part form field parsing
	'-------------------------------------------------------------------------
	Public Default Sub ParseFields()
		Dim binData			' binary stream
		Dim sFieldName
		Dim sFieldValue
		Dim lStartAt		' start of stream	
		Dim lEndAt			' end of stream
		Dim lFileAt
		Dim oUploadFile
		Dim sFileName
		Dim lCursor
		Dim binUseData		' useable data in stream
		Dim lUseStart
		Dim lCursorMax

		binData = Request.BinaryRead(Request.TotalBytes)
		lStartAt = 1
		lEndAt = InStrB(lStartAt, binData, StrToBin(Chr(13)))
		
		If (lEndAt - lStartAt) <= 0 Then Exit Sub
		 
		binUseData = MidB(binData, lStartAt, lEndAt - lStartAt)
		lUseStart = InStrB(1, binData, binUseData)
		
		Do Until lUseStart = InStrB(binData, binUseData & StrToBin("--"))

			lCursor = InStrB(lUseStart, binData, StrToBin("Content-Disposition"))
			lCursor = InStrB(lCursor, binData, StrToBin("name="))
			lStartAt = lCursor + 6									' get past "name="; start of field name
			lEndAt = InStrB(lStartAt, binData, StrToBin(Chr(34)))	' quote (") ends field name
			sFieldName = LCase(BinToStr(MidB(binData, lStartAt, lEndAt - lStartAt)))
			lFileAt = InStrB(lUseStart, binData, StrToBin("filename="))
			lCursorMax = InStrB(lEndAt, binData, binUseData)
			
			If lFileAt <> 0 And lFileAt < lCursorMax Then
				' file name was found in useable data
				Set oUploadFile = New kbUploadedFile

				' get file name
				lStartAt = lFileAt + 10								' get past "filename="; start of file name
				lEndAt =  InStrB(lStartAt, binData, StrToBin(Chr(34)))
				sFileName = BinToStr(MidB(binData, lStartAt, lEndAt-lStartAt))
				oUploadFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))

				' get file type
				lCursor = InStrB(lEndAt, binData, StrToBin("Content-Type:"))
				lStartAt = lCursor + 14
				lEndAt = InStrB(lStartAt, binData, StrToBin(Chr(13)))
				oUploadFile.ContentType = BinToStr(MidB(binData, lStartAt, lEndAt-lStartAt))
				
				' get file data
				lStartAt = lEndAt + 4
				lEndAt = InStrB(lStartAt, binData, binUseData) - 2
				oUploadFile.FileData = MidB(binData, lStartAt, lEndAt-lStartAt)
				
				If oUploadFile.FileSize > 0 Then Call File.Add(sFieldName, oUploadFile)
				'Set oUploadFile = Nothing
			Else
				' must be a form value
				lCursor = InStrB(lCursor, binData, StrToBin(Chr(13)))
				lStartAt = lCursor + 4
				lEndAt = InStrB(lStartAt, binData, binUseData) - 2
				sFieldValue = BinToStr(MidB(binData, lStartAt, lEndAt-lStartAt))

				If m_oFieldHash.Exists(sFieldName) Then
					' must be multi-part field
					sFieldValue = m_oFieldHash(sFieldName) & "," & sFieldValue
					m_oFieldHash.Remove(sFieldName)
				End If
				Call m_oFieldHash.Add(sFieldName, sFieldValue)
			End If

			lUseStart = InStrB(lUseStart + LenB(binUseData), binData, binUseData)
		Loop
	End Sub

	' String to byte string conversion
	Private Function StrToBin(ByVal v_sString)
		Dim x
		For x = 1 to Len(v_sString)
		   StrToBin = StrToBin & ChrB(AscB(Mid(v_sString, x, 1)))
		Next
	End Function

	' Byte string to string conversion
	Private Function BinToStr(ByVal v_sString)
		Dim x
		BinToStr = ""
		For x = 1 to LenB(v_sString)
		   BinToStr = BinToStr & Chr(AscB(MidB(v_sString, x, 1))) 
		Next
	End Function
End Class

'-------------------------------------------------------------------------
'	Name: 		kbUploadedFile
'	Purpose: 	provide methods for uploaded file
'Modifications:
'	Date:		Name:	Description:
'	12/07/00	Jacob "Beezle" Gilley (avis7@airmail.net)
'	12/28/02	JEA		Improved formatting; object cleanup
'-------------------------------------------------------------------------
Class kbUploadedFile
	Public ContentType	' not used
	Public FileName		' file name
	Public FileData		' binary file contents
	
	Public Property Get FileSize()
		FileSize = LenB(FileData)
	End Property
	
	'-------------------------------------------------------------------------
	'	Name: 		SaveToDisk()
	'	Purpose: 	save file to local disk
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/07/00	JG		Creation
	'	12/28/02	JEA		check if file exists and clean up objects
	'	12/31/02	JEA		allow auto-rename if file exists, or explicit rename
	'-------------------------------------------------------------------------
	Public Function SaveToDisk(ByVal v_sPath, ByVal v_lMaxKB, ByVal v_sNewName, _
		ByVal v_bAutoRename, ByVal v_bReplace)
		
		Dim oFileSys
		Dim oFile
		Dim oData
		Dim lExtAt
		Dim sBaseName
		Dim sExtension
		Dim x

		If v_sPath = "" Or FileName = "" Then
			SaveToDisk = g_sMSG_NO_FILE_OR_PATH
			Exit Function
		End If
		
		If LenB(FileData) > (v_lMaxKB * 1024) Then
			SaveToDisk = g_sMSG_FILE_TOO_BIG & v_lMaxKB & "KB"
			Set oData = New kbDataAccess
			Call oData.LogActivity(g_ACT_TOO_LARGE_FILE_UPLOAD, "", "", "", "", "")
			Set oData = Nothing
			Exit Function
		End if

		lExtAt = InStrRev(FileName, ".") - 1
		sBaseName = Left(FileName, lExtAt)
		sExtension = Right(FileName, Len(FileName) - lExtAt)
		If Mid(v_sPath, Len(v_sPath)) <> "\" Then v_sPath = v_sPath & "\"
		
		If v_sNewName <> "" Then
			FileName = v_sNewName & sExtension
		ElseIf g_sREPLACE_SPACE_WITH <> "" Then
			FileName = Replace(FileName, " ", g_sREPLACE_SPACE_WITH)
		End If
		
		Set oFileSys = Server.CreateObject(g_sFILE_SYSTEM_OBJECT)
		
		If Not oFileSys.FolderExists(v_sPath) Then
			Set oFileSys = nothing
			SaveToDisk = g_sMSG_NO_UPLOAD_DIR
			Exit Function
		End If
		
		If oFileSys.FileExists(v_sPath & FileName) Then
			If v_bAutoRename Then
				' find useable name for file
				x = 10
				FileName = sBaseName & "_" & x & sExtension
				Do While oFileSys.FileExists(v_sPath & FileName)
					x = x + 1
					FileName = sBaseName & "_" & x & sExtension
				Loop
			ElseIf Not v_bReplace Then
				Set oFileSys = nothing
				SaveToDisk = g_sMSG_FILE_EXISTS
				Exit Function
			End If
		End If
		
		Set oFile = oFileSys.CreateTextFile(v_sPath & FileName, True)
		
		For x = 1 to LenB(FileData)
		    oFile.Write Chr(AscB(MidB(FileData,x,1)))
		Next
		
		oFile.Close
		Set oFile = nothing
		Set oFileSys = nothing
		SaveToDisk = ""
	End Function
	
	Public Sub SaveToDatabase(ByRef r_oField)
		If LenB(FileData) = 0 Then Exit Sub
		If IsObject(r_oField) Then r_oField.AppendChunk FileData
	End Sub
End Class
%>