<%
Const m_USER_ID = 0
Const m_FIRST_NAME = 1
Const m_LAST_NAME = 2
Const m_SCREEN_NAME = 3
Const m_USER_TYPE = 4
Const m_USER_STATUS = 5
Const m_ITEM_SORT = 6
Const m_ITEMS_PER_PAGE = 7
Const m_USER_FILE_FORMAT = 8
Const m_USER_EMAIL = 9
Const m_USER_PRIVATE = 10
Const m_USER_SPAM = 11
Const m_USER_IMAGE = 12
Const m_USER_ABOUT = 13
Const m_USER_WEB_URL = 14
Const m_USER_PASSWORD = 15
Const m_USER_SITE = 16
Const m_USER_LAST_LOGIN = 17

'-------------------------------------------------------------------------
'	Name: 		kbUser class
'	Purpose: 	encapsulate user functions
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Creation
'	4/29/03		JEA		Retrieve site id and last login date
'-------------------------------------------------------------------------
Class kbUser
	Private m_sBaseSQL
	Private m_oData
	Private m_lDefaultSortID
	Private m_lDefaultFormatID
	Private m_lItemsPerPage

	Private Sub Class_Initialize()
		Set m_oData = New kbDataAccess
		m_lDefaultSortID = g_SORT_DATE_DESC
		m_lDefaultFormatID = g_FORMAT_NTSC
		m_lItemsPerPage = g_ITEMS_PER_PAGE
		m_sBaseSQL = "SELECT lUserID, vsFirstName, vsLastName, vsScreenName, lUserTypeID, lStatusID, " _
			& "lDefaultSortID, lItemsPerPage, lDefaultFormatID, vsEmail, bPrivateEmail, " _
			& "bEmailNews, vsUserImage, vsAboutMe, vsHomePageURL, vsPassword, lSiteID, dtLastLogin FROM tblUsers "
	End Sub
	
	Private Sub Class_Terminate()
		Set m_oData = nothing
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		Login()
	'	Purpose: 	get array for user
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/23/02	JEA		Creation
	'	4/26/03		JEA		Include Site ID
	'-------------------------------------------------------------------------
	Public Function Login(ByVal v_sEmail, ByVal v_sPassword, ByVal v_lTimeShift, ByVal v_lSiteID)
		dim sQuery
		dim aData
		dim sMessage
		dim oEncrypt
		
		'Set oEncrypt = New kbEncryption
		sQuery = m_sBaseSQL & "WHERE vsEmail = '" & v_sEmail & "' AND vsPassword = '" _
			& v_sPassword & "'"
		'Set oEncrypt = Nothing
		aData = m_oData.GetArray(sQuery)
		If IsArray(aData) Then
			If GoodStatus(aData(m_USER_STATUS, 0)) Then
				' successful login
				Call CreateUserSession(aData, v_lTimeShift, v_lSiteID)
				Call UpdateLastLogin(aData(m_USER_ID, 0))
				Call m_oData.LogActivity(g_ACT_LOGIN, "", "", "", "", "", "")
				sMessage = ""
			Else
				sMessage = "That account is currently unavailable"
			End If
		Else
			sMessage = "That username or password is incorrect "
			Call m_oData.LogActivity(g_ACT_FAILED_LOGIN, "", "", "", "", v_sEmail, v_sPassword)
		End If
		Login = sMessage
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		LoginWithCookie()
	'	Purpose: 	get array for user based on cookie
	'	Return: 	boolean
	'Modifications:
	'	Date:		Name:	Description:
	'	12/28/02	JEA		Creation
	'	4/26/03		JEA		Track site ID
	'-------------------------------------------------------------------------
	Public Function LoginWithCookie(ByVal v_lSiteID)
		Const FIRST_NAME = 0
		Const LAST_NAME = 1
		Const USER_TYPE = 2
		Const USER_STATUS = 3
		Const FILE_SORT = 4
		Const FILES_PER_PAGE = 5
		dim lUserID
		dim bSuccess
		dim sQuery
		dim aData
		
		bSuccess = false
		lUserID = Request.Cookies(g_sUSER_COOKIE)
		If IsVoid(v_lSiteID) Then v_lSiteID = Request.Cookies(g_sSITE_COOKIE)
		If IsNumber(lUserID) Then
			sQuery = m_sBaseSQL & "WHERE lUserID = " & lUserID
			aData = m_oData.GetArray(sQuery)
			If IsArray(aData) Then
				If GoodStatus(aData(m_USER_STATUS, 0)) Then
					Call CreateUserSession(aData, MakeNumber(Request.Cookies(g_sTIME_COOKIE)), v_lSiteID)
					Call UpdateLastLogin(lUserID)
					Call m_oData.LogActivity(g_ACT_COOKIE_LOGIN, "", "", "", "", "", "")
					bSuccess = true
				End If
			End If
		End If
		LoginWithCookie = bSuccess
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GoodStatus()
	'	Purpose: 	check for permissable statuses
	'Modifications:
	'	Date:		Name:	Description:
	'	1/10/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function GoodStatus(ByVal v_lStatusID)
		GoodStatus = CBool(CStr(v_lStatusID) <> CStr(g_STATUS_DISABLED))
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		CreateUserSession()
	'	Purpose: 	persist user values to local session and cookie
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	12/31/02	JEA		Process client-server time difference
	'	4/25/03		JEA		Add site to session and cookie
	'	4/29/03		JEA		More fail-safes for retrieving site id
	'-------------------------------------------------------------------------
	Private Sub CreateUserSession(ByVal v_aData, ByVal v_lTimeShift, ByVal v_lSiteID)
		dim sUserName
		dim sLastLogin
		dim aSession(11)
		
		if IsVoid(v_lSiteID) then v_lSiteID = Request.Cookies(g_sSITE_COOKIE)
		if IsVoid(v_lSiteID) then v_lSiteID = v_aData(m_USER_SITE, 0)
		if IsVoid(v_lSiteID) then v_lSiteID = g_DEFAULT_SITE
		
		sLastLogin = v_aData(m_USER_LAST_LOGIN, 0)
		If Not IsDate(sLastLogin) Then sLastLogin = Date()
		
		sUserName = Trim(v_aData(m_FIRST_NAME, 0) & " " & v_aData(m_LAST_NAME, 0))
		aSession(g_USER_ID) = v_aData(m_USER_ID, 0)
		aSession(g_USER_TYPE) = v_aData(m_USER_TYPE, 0)
		aSession(g_USER_STATUS) = v_aData(m_USER_STATUS, 0)
		aSession(g_USER_IP) = Request.ServerVariables("REMOTE_ADDR")
		aSession(g_USER_NAME) = sUserName
		aSession(g_USER_ITEM_SORT) = v_aData(m_ITEM_SORT, 0)
		aSession(g_USER_ITEMS_PER_PAGE) = v_aData(m_ITEMS_PER_PAGE, 0)
		aSession(g_USER_FILE_FORMAT) = v_aData(m_USER_FILE_FORMAT, 0)
		aSession(g_USER_TIME_SHIFT) = v_lTimeShift
		aSession(g_USER_MSG) = "Welcome back " & sUserName
		aSession(g_USER_SITE) = v_lSiteID
		aSession(g_USER_LAST_LOGIN) = sLastLogin
		Session(g_sSESSION) = aSession
		With Response
			.Cookies(g_sUSER_COOKIE) = v_aData(m_USER_ID, 0)
			.Cookies(g_sUSER_COOKIE).expires = #1/1/2010 00:00:00#
			.Cookies(g_sTIME_COOKIE) = v_lTimeShift
			.Cookies(g_sTIME_COOKIE).expires = #1/1/2010 00:00:00#
			.Cookies(g_sSITE_COOKIE) = aSession(g_USER_SITE)
			.Cookies(g_sSITE_COOKIE).expires = #1/1/2010 00:00:00#
		End With
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		LastLogin()
	'	Purpose: 	get last login and update with current date
	'Modifications:
	'	Date:		Name:	Description:
	'	5/1/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub UpdateLastLogin(v_lUserID)
		dim sQuery
		sQuery = "UPDATE tblUsers SET dtLastLogin = " & g_sSQL_DATE_DELIMIT & Date() & g_sSQL_DATE_DELIMIT _
			& " WHERE lUserID = " & v_lUserID
		Call m_oData.ExecuteOnly(sQuery)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Logout()
	'	Purpose: 	clear user session
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub Logout()
		Session(g_sSESSION) = ""
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Validate()
	'	Purpose: 	see if validation code matches
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Function Validate(ByVal v_sCode, ByVal v_lUserID)
		dim sQuery
		dim aData
		dim sMessage
		
		sQuery = "SELECT sValidationCode FROM tblUsers WHERE lUserID = " & v_lUserID
		aData = m_oData.GetArray(sQuery)
		If IsArray(aData) Then
			If aData(0,0) = v_sCode Then
				Call UpdateUserStatus(v_lUserID)
				Call m_oData.LogActivity(g_ACT_VALIDATE_REGISTRATION, "", "", "", "", "", "")
			Else
				sMessage = g_sMSG_CODE_MISMATCH
				Call m_oData.LogActivity(g_ACT_BAD_VALIDATION, "", "", "", "", "", "")
			End If
		Else
			sMessage = "Your registration could not be found.  Please try signing in again."
		End If
		Validate = sMessage
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		UpdateUserStatus()
	'	Purpose: 	update user's status
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub UpdateUserStatus(ByVal v_lUserID)
		dim sQuery
		sQuery = "UPDATE tblUsers SET lUserTypeID = " & g_USER_VERIFIED _
			& ", lStatusID = " & g_STATUS_APPROVED & " WHERE lUserID = " & v_lUserID
		Call m_oData.ExecuteOnly(sQuery)
		
		Call SetSessionValue(g_USER_TYPE, g_USER_VERIFIED)
		Call SetSessionValue(g_USER_STATUS, g_STATUS_APPROVED)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Register()
	'	Purpose: 	save user data to db
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	1/11/03		JEA		Add screen name
	'	4/26/03		JEA		Include Site ID
	'-------------------------------------------------------------------------
	Public Function Register(ByVal v_sFirstName, ByVal v_sLastName, ByVal v_sScreenName, ByVal v_sEmail, _
		ByVal v_sPassword, ByVal v_sHomePage, ByVal v_bPrivacy, ByVal v_bNotify, ByVal v_lTimeShift, _
		ByVal v_lSiteID)
		
		Const UNVERIFIED_USER = 2
		Const STATUS_PENDING = 1
		dim oRS
		dim oMail
		dim oEncryrpt
		dim lUserID
		dim aData(17,0)
		dim sMessage
		dim sValidationCode
		
		sMessage = ""
	
		If EmailExists(v_sEmail, "") Then
			sMessage = "Someone has already registered with that e-mail address"
			Call m_oData.LogActivity(g_ACT_REGISTER_DUPLICATE_EMAIL, "", "", "", "", v_sEmail, "")
		Else
			sValidationCode = MakeValidationCode(g_VALIDATION_CODE_LENGTH)
			lUserID = Insert(v_sFirstName, v_sLastName, v_sScreenName, v_sHomePage, v_sEmail, _
				v_sPassword, v_bPrivacy, v_bNotify, m_lDefaultSortID, m_lDefaultFormatID, _
				m_lItemsPerPage, "", sValidationCode, g_STATUS_PENDING, v_lSiteID)
			
			
			' generate array for session creation
			aData(m_USER_ID, 0) = lUserID
			aData(m_FIRST_NAME, 0) = v_sFirstName
			aData(m_LAST_NAME, 0) = v_sLastName
			aData(m_SCREEN_NAME, 0) = v_sScreenName
			aData(m_USER_TYPE, 0) = UNVERIFIED_USER
			aData(m_USER_STATUS, 0) = STATUS_PENDING
			aData(m_ITEM_SORT, 0) = m_lDefaultSortID
			aData(m_ITEMS_PER_PAGE, 0) = m_lItemsPerPage
			aData(m_USER_FILE_FORMAT, 0) = m_lDefaultFormatID
			aData(m_USER_LAST_LOGIN, 0) = Null
			
			Call CreateUserSession(aData, v_lTimeShift, v_lSiteID)			
			Call m_oData.LogActivity(g_ACT_REGISTER, "", "", "", "", "", "")

			Set oMail = New kbMail
			Call oMail.SendConfirmationMail(lUserID, v_sFirstName, v_sLastName, v_sEmail, sValidationCode)
			Set oMail = Nothing
		End If
		Register = sMessage
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		EmailExists()
	'	Purpose: 	Check to see if an e-mail address already exists
	'	Return:		boolean
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	1/5/03		JEA		exclude user id option
	'-------------------------------------------------------------------------
	Private Function EmailExists(ByVal v_sEmail, ByVal v_lExcludeID)
		dim sQuery
		dim aData
		sQuery = "SELECT lUserID FROM tblUsers WHERE vsEmail = '" & v_sEmail & "'"
		if v_lExcludeID <> "" then sQuery = sQuery & " AND lUserID <> " & v_lExcludeID
		aData = m_oData.GetArray(sQuery)
		EmailExists = IsArray(aData)
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		MakeValidationCode()
	'	Purpose: 	generate a random validation string
	'	Return:		string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'-------------------------------------------------------------------------
	Private Function MakeValidationCode(ByVal v_lCodeLength)
		dim x
		dim sCode
		dim lRandom
		
		sCode = ""
		Randomize
		for x = 1 to v_lCodeLength
			' alternate random numbers and letters
			lRandom = IIf((x Mod 2), Int((25 * Rnd) + 1) + 65, Int((10 * Rnd) + 1) + 47)
			sCode = sCode & LCase(Chr(lRandom))
		next
		MakeValidationCode = sCode
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		GetUserItemData()
	'	Purpose: 	get user information about item
	'	Return:		array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'	7/21/04		JEA		Abstract for different item types
	'-------------------------------------------------------------------------
	Public Function GetUserItemData(ByVal v_lItemID, ByVal v_lItemTypeID)
		dim sQuery
		
		Select Case v_lItemTypeID
			Case g_ITEM_PROJECT
				sQuery = "SELECT U.vsFirstName, U.vsEmail, F.vsFileName, F.dtSubmitDate " _
					& "FROM tblProjects F INNER JOIN tblUsers U ON F.lUserID = U.lUserID " _
					& "WHERE lProjectID = " & v_lItemID
			Case g_ITEM_SCRIPT
				sQuery = "SELECT U.vsFirstName, U.vsEmail, S.vsFileName, S.dtSubmitDate " _
					& "FROM tblScripts S INNER JOIN tblUsers U ON S.lUserID = U.lUserID " _
					& "WHERE lScriptID = " & v_lItemID
		End Select
		GetUserItemData = m_oData.GetArray(sQuery)
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteSortList()
	'	Purpose: 	write option list for sort options
	'Modifications:
	'	Date:		Name:	Description:
	'	1/5/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteSortList(ByVal v_sFieldName, ByVal v_lSelectedID)
		dim sQuery
		sQuery = "SELECT lSortID, vsSortName FROM tblSortValues"
		with response
			.write "<select name='"
			.write v_sFieldName
			.write "'>"
			.write MakeList(sQuery, v_lSelectedID)
			.write "</select>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteFormatList()
	'	Purpose: 	write option list with video formats
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub WriteFormatList(ByVal v_sFieldName, ByVal v_lSelectedID)
		dim sQuery
		sQuery = "SELECT lFormatID, vsFormatDescription FROM tblFormats ORDER BY vsFormatDescription"
		with response
			.write "<select name='"
			.write v_sFieldName
			.write "'>"
			.write MakeList(sQuery, v_lSelectedID)
			.write "</select>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteUser()
	'	Purpose: 	write user data
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'	4/30/03		JEA		Show e-mail to admins
	'-------------------------------------------------------------------------
	Public Sub WriteUser(ByVal v_lUserID)
		Const sFORM_NAME = "frmEmail"
		dim oLayout
		dim sQuery
		dim aData
		dim x
		
		sQuery = m_sBaseSQL & "WHERE lUserID = " & v_lUserID
		aData = m_oData.GetArray(sQuery)
		If IsArray(aData) Then
			Set oLayout = New kbLayout
			with response
				.write "<table cellspacing='0' cellpadding='0' border='0' width='400'>"
				.write "<tr><td valign='top'>"
				' name
				.write "<div class='Name'><nobr>"
				If Not IsVoid(aData(m_SCREEN_NAME, 0)) Then
					.write aData(m_SCREEN_NAME, 0)
					.write "</nobr></div><div class='RealName'><nobr>"
				End If
				.write aData(m_FIRST_NAME, 0)
				.write " " 
				.write aData(m_LAST_NAME, 0)
				.write "</nobr></div>"
				' e-mail
				if (Not aData(m_USER_PRIVATE, 0)) Or g_bAdmin then
					.write "<div class='Email'><a href='mailto:"
					.write aData(m_USER_EMAIL, 0)
					.write "'>"
					.write aData(m_USER_EMAIL, 0)
					.write "</a>"
					if aData(m_USER_PRIVATE, 0) then .write "<sup>*</sup>"
					.write "</div>"
				end if
				' web site
				if Trim(aData(m_USER_WEB_URL, 0)) <> "" then
					.write "<div class='WebURL'><a href='http://"
					.write Trim(aData(m_USER_WEB_URL, 0))
					.write "' target='_new'>http://"
					.write Trim(aData(m_USER_WEB_URL, 0))
					.write "</a></div>"
				end if
				.write "</td><td align='right' valign='top'>"
				' image
				if Trim(aData(m_USER_IMAGE, 0)) <> "" then
					.write "<img class='UserImage' src='./images/user/"
					.write Trim(aData(m_USER_IMAGE, 0))
					.write "'>"
				end if
				.write "</td><tr><td colspan='2'>"
				' biography
				if Trim(aData(m_USER_ABOUT, 0)) <> "" then
					.write "<div class='About'>"
					.write FormatAsHTML(Trim(aData(m_USER_ABOUT, 0)))
					.write "</div>"
				end if
				' button
				if g_bAdmin Or CStr(GetSessionValue(g_USER_ID)) = CStr(aData(m_USER_ID, 0)) then
					.write "<div align='right'><a href='kb_user-edit.asp?id="
					.write aData(m_USER_ID, 0)
					.write "'>"
					Call oLayout.WriteToggleImage("btn_edit", "", "Edit Account", "", false)
					.write "</a></div>"
				end if
				' e-mail form
				If aData(m_USER_PRIVATE, 0) And Not g_bAdmin Then
					' write e-mailing form
					.write "<br>"
					Call oLayout.WriteTitleBoxTop("Send e-mail", "", "")
					.write "<table cellspacing='0' cellpadding='1' border='0'>"
					.write "<form name='"
					.write sFORM_NAME
					.write "' action='kb_user.asp?id="
					.write aData(m_USER_ID, 0)
					.write "' method='post' onSubmit=""return IsValid('"
					.write sFORM_NAME
					.write "', m_oFields);""><tr><td></td><td class='FormNote'>At their request, this member's address "
					.write "is hidden</td><tr>"
					.write "<td class='Required'>Subject:</td><td class='FormInput'>"
					.write "<input type='text' name='fldSubject'></td><tr>"
					.write "<td class='FormLabel'>Bcc:</td><td class='FormInput'>"
					.write "<input style='border: none;' type='checkbox' name='fldBcc'> send me a copy of this message</td><tr>"
					.write "<td class='Required' valign='top'>Message:</td><td class='FormInput'>"
					.write "<textarea rows='8' cols='50' name='fldBody'></textarea>"
					.write "</td><tr><td class='Required' style='font-size: 8pt; text-align: center;'>"
					.write "(required)</td><td align='right' valign='bottom'>"
					Call oLayout.WriteToggleImage("btn_send-message", "", "Send Message", "class='Image'", true)
					.write "</td><input type='hidden' name='fldFrom' value='"
					.write GetSessionValue(g_USER_ID)
					.write "'><input type='hidden' name='fldTo' value='"
					.write aData(m_USER_ID, 0)
					.write "'></form></td></table>"
					Call oLayout.WriteBoxBottom("")
				ElseIf g_bAdmin And aData(m_USER_PRIVATE, 0) Then
					.write "<div class='FootNote'><sup>*</sup>This address is hidden from non-administrators</div>"
				End If
				.write "</td></table>"
			end with
			Set oLayout = Nothing
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteUserList()
	'	Purpose: 	write option list of users
	'Modifications:
	'	Date:		Name:	Description:
	'	1/7/03		JEA		Creation
	'-------------------------------------------------------------------------	
	Public Sub WriteUserList(ByVal v_sFieldName, ByVal v_lSelectedID, ByVal v_lStatusID, ByVal v_sTopItem)
		dim sQuery
		sQuery = "SELECT lUserID, IIf(vsLastName IS NULL, '', vsLastName + ', ') + vsFirstName FROM tblUsers"
		if IsNumber(v_lStatusID) then sQuery = sQuery & " WHERE lStatusID = " & v_lStatusID
		sQuery = sQuery & " ORDER BY vsLastName, vsFirstName"
		with response
			.write "<select name='"
			.write v_sFieldName
			.write "'>"
			if v_sTopItem <> "" then
				.write "<option value='0'>"
				.write v_sTopItem
			end if
			.write MakeList(sQuery, v_lSelectedID)
			.write "</select>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		WriteUserList()
	'	Purpose: 	write option list of users
	'Modifications:
	'	Date:		Name:	Description:
	'	1/7/03		JEA		Creation
	'-------------------------------------------------------------------------	
	Public Sub WritePageSizeList(ByVal v_sFieldName, ByVal v_lSelectedID)
		dim sList
		dim x
		
		for x = 10 to 50 step 10
			sList = sList & "<option value='" & x & "'>" & x
		next
		with response
			.write "<select name='"
			.write v_sFieldName
			.write "'>"
			.write MakeSelected(sList, v_lSelectedID)
			.write "</select>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Save()
	'	Purpose: 	write user data from custom form parse
	'Modifications:
	'	Date:		Name:	Description:
	'	1/5/03		JEA		Creation
	'	3/31/03		JEA		Update page size session value
	'	4/26/03		JEA		Insert Site ID
	'-------------------------------------------------------------------------	
	Public Sub Save(ByVal v_sPictureField)
		dim oForm
		dim oFile
		dim oMail
		dim sQuery
		dim sMessage
		dim sValidation
		dim lProjectID
		dim bNewEmail
		dim oEncrypt
		dim sPictureName
		dim sFirstName
		dim sLastName
		dim sScreenName
		dim sURL
		dim sEmail
		dim bPrivacy
		dim bNotify
		dim lSortID
		dim lFormatID
		dim lPageSize
		dim sAbout
		dim sPassword
		dim sReturnURL
		dim lUserID
		dim bNewUser
		
		sMessage = ""
		sPictureName = ""
		v_sPictureField = LCase(v_sPictureField)

		Set oForm = New kbForm
		Call oForm.ParseFields()
		bNewUser = CBool(oForm.Field("fldNew") = "True")
		bNewEmail = bNewUser Or CBool(Trim(oForm.Field("fldEmail")) <> Trim(oForm.Field("fldOldEmail")))
		lUserID = IIf(bNewUser, "", ReplaceNull(oForm.Field("fldUserID"), GetSessionValue(g_USER_ID)))
		
		If bNewEmail And EmailExists(oForm.Field("fldEmail"), lUserID) Then
			Call SetSessionValue(g_USER_MSG, "The address " & oForm.Field("fldEmail") _
				& " is already taken by another user")
			response.redirect "kb_user-edit.asp?id=" & lUserID & "&new=" & IIf(bNewUser, "yes", "")
			Exit Sub
		End If
		
		With oForm
			sFirstName = CleanForSQL(.Field("fldFirstName"))
			sLastname = CleanForSQL(.Field("fldLastName"))
			sScreenName = CleanForSQL(.Field("fldScreenName"))
			sURL = Replace(.Field("fldWebURL"), "http://", "")
			sEmail = .Field("fldEmail")
			bPrivacy = CBool(.Field("fldPrivacy") = "on")
			bNotify = CBool(.Field("fldNotify") = "on")
			lSortID = .Field("fldSortBy")
			lFormatID = .Field("fldFormat")
			lPageSize = .Field("fldItemsPerPage")
			sAbout = CleanForSQL(.Field("fldBio"))
			sPassword = Trim(.Field("fldPassword"))
			sReturnURL = .Field("fldReturnURL")
		End With
		
		If bNewUser Then
			' insert user data; admin-created users are automatically approved
			lUserID = Insert(sFirstName, sLastname, sScreenName, sURL, sEmail, sPassword, _
				bPrivacy, bNotify, lSortID, lFormatID, lPageSize, sAbout, "", g_STATUS_APPROVED, _
				GetSessionValue(g_USER_SITE))
			Call m_oData.LogActivity(g_ACT_CREATE_USER, "", "", "", lUserID, "", "")
		End If
		
		If oForm.File.Exists(v_sPictureField) Then
			sPictureName = "user_" & PadNumber(lUserID, 3)
			Set oFile = oForm.File.Item(v_sPictureField)
			sMessage = oFile.SaveToDisk(Server.MapPath("./images/user"), g_MAX_PICTURE_KB, sPictureName, false, true)
			sPictureName = oFile.FileName
			Set oFile = Nothing
		End If

		If sMessage = "" Then
			If bNewUser Then
				' update newly created user's picture name
				sQuery = "UPDATE tblUsers SET vsUserImage = '" & sPictureName & "' "
			Else
				' update all user data
				sQuery = "UPDATE tblUsers SET " _
					& "vsFirstName = '" & sFirstName & "', " _
					& "vsLastName = '" & sLastname & "', " _
					& "vsScreenName = " & MaybeNull(sScreenName) & ", " _
					& "vsHomePageURL = '" & sURL & "', " _
					& "vsEmail = '" & sEmail & "', " _
					& "bPrivateEmail = " & bPrivacy & ", " _
					& "bEmailNews = " & bNotify & ", " _
					& "lDefaultSortID = " & lSortID & ", " _
					& "lDefaultFormatID = " & lFormatID & ", " _
					& "lItemsPerPage = " & lPageSize & ", " _
					& "vsAboutMe = '" & sAbout & "' "
				if sPictureName <> "" then sQuery = sQuery & ", vsUserImage = '" & sPictureName & "' "
				if sPassword <> "" then sQuery = sQuery & ", vsPassword = '" & sPassword & "' "
				
				if bNewEmail then
					sValidation = MakeValidationCode(g_VALIDATION_CODE_LENGTH)
					sQuery = sQuery & ", " _
						& "lStatusID = " & g_STATUS_PENDING & ", " _
						& "sValidationCode = '" & sValidation & "' "
					
					' send confirmation mail for address change
					Set oMail = New kbMail
					Call oMail.SendConfirmationMail(lUserID, oForm.Field("fldFirstName"), _
						oForm.Field("fldLastName"), sEmail, sValidation)
					Set oMail = Nothing
				end if
			End If
			sQuery = sQuery & "WHERE lUserID = " & lUserID
			Call m_oData.ExecuteOnly(sQuery)
		End If
		Set oForm = Nothing
		
		If bNewEmail And Not bNewUser Then
			Call SetSessionValue(g_USER_MSG, g_sMSG_NEED_VALIDATION)
			Call SetSessionValue(g_USER_STATUS, g_STATUS_PENDING)
			response.redirect "kb_validate.asp?" & Server.URLEncode("url=kb_user.asp?id=" & lUserID)
		ElseIf bNewUser Then
			Call SetSessionValue(g_USER_MSG, g_sMSG_ADMIN_ADD_USER)
			response.redirect ReplaceNull(sReturnURL, "kb_user.asp?id=" & lUserID)
		Else
			Call SetSessionValue(g_USER_ITEMS_PER_PAGE, lPageSize)
			Call SetSessionValue(g_USER_MSG, sMessage)
			response.redirect "kb_user.asp?id=" & lUserID
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetArray()
	'	Purpose: 	get user data, one way or another
	'	Return: 	array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/7/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Function GetArray(ByVal v_bNew)
		dim sQuery
		dim aData
		dim oData
		dim x
		dim aReturn(17)
		
		If Not v_bNew Then
			aReturn(m_USER_ID) = ReplaceNull(Trim(Request.QueryString("id")), GetSessionValue(g_USER_ID))
			If IsNumber(aReturn(m_USER_ID)) Then
				sQuery = m_sBaseSQL & "WHERE lUserID = " & aReturn(m_USER_ID)
				aData = m_oData.GetArray(sQuery)
				If IsArray(aData) Then
					for x = 0 to UBound(aData)
						aReturn(x) = aData(x, 0)
					next
				End If
			Else
				With Request
					aReturn(m_FIRST_NAME) = Trim(.Form("fldFirstName"))
					aReturn(m_LAST_NAME) = Trim(.Form("fldLastName"))
					aReturn(m_SCREEN_NAME) = Trim(.Form("fldScreenName"))
					aReturn(m_ITEM_SORT) = .Form("fldSortBy")
					aReturn(m_ITEMS_PER_PAGE) = .Form("fldItemsPerPage")
					aReturn(m_USER_FILE_FORMAT) = .Form("fldFormat")
					aReturn(m_USER_EMAIL) = .Form("fldEmail")
					aReturn(m_USER_PRIVATE) = CBool(.Form("fldPrivacy") = "on")
					aReturn(m_USER_SPAM) = CBool(.Form("fldNotify") = "on")
					aReturn(m_USER_ABOUT) = .Form("fldBio")
					aReturn(m_USER_WEB_URL) = .Form("fldWebURL")
					aReturn(m_USER_PASSWORD) = .Form("fldPassword")
				End With
			End If
		Else
			aReturn(m_USER_SPAM) = true
			aReturn(m_USER_FILE_FORMAT) = 1
			aReturn(m_ITEM_SORT) = g_SORT_DATE_DESC
			aReturn(m_ITEMS_PER_PAGE) = 10
		End If
		GetArray = aReturn
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		Insert()
	'	Purpose: 	save new user data to db
	'	Return: 	number
	'Modifications:
	'	Date:		Name:	Description:
	'	1/7/03		JEA		Creation
	'	4/26/03		JEA		Add Site ID
	'-------------------------------------------------------------------------
	Private Function Insert(ByVal v_sFirstName, ByVal v_sLastname, ByVal v_sScreenName, _
		ByVal v_sURL, ByVal v_sEmail, ByVal v_sPassword, ByVal v_bPrivacy, ByVal v_bNotify, _
		ByVal v_lSortID, ByVal v_lFormatID, ByVal v_lPageSize, ByVal v_sAbout, ByVal v_sCode, _
		ByVal v_lStatusID, ByVal v_lSiteID)
		
		dim oRS
		dim lUserID
	
		Set oRS = Server.CreateObject("ADODB.Recordset")
		With oRS
			.Open "tblUsers", m_oData.Connection, adOpenStatic, adLockOptimistic, adCmdTable
			.AddNew
			.Fields("vsFirstName") = v_sFirstName
			.Fields("vsLastName") = v_sLastName
			.Fields("vsScreenName") = ReplaceNull(v_sScreenName, null)
			.Fields("vsEmail") = v_sEmail
			.Fields("vsPassword") = v_sPassword
			.Fields("vsHomePageURL") = v_sURL
			.Fields("lDefaultSortID") = v_lSortID
			.Fields("lDefaultFormatID") = v_lFormatID
			.Fields("lItemsPerPage") = v_lPageSize
			.Fields("bPrivateEmail") = v_bPrivacy
			.Fields("bEmailNews") = v_bNotify
			.Fields("lUserTypeID") = g_USER_VERIFIED
			.Fields("lStatusID") = v_lStatusID
			.Fields("sValidationCode") = v_sCode
			.Fields("vsAboutMe") = v_sAbout
			.Fields("lSiteID") = ReplaceNull(v_lSiteID, g_DEFAULT_SITE)
			.Fields("dtDateRegistered") = Now()
			.Update
			lUserID = .Fields("lUserID")
			.Close
		End With
		Set oRS = nothing
		Insert = lUserID
	End Function
End Class
%>
