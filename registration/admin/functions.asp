<!--#include file="includes/md5.asp"-->
<%

Response.Buffer = True
Response.Expires = 0

'''''''''''''''''''''''''''''' INFORMATION ''''''''''''''''''''''''''''''
' This is what will keep your server from being hacked by the software, '
' as long as this key stays the same the system will function and the   '
' system will be really hard to be faked or hacked.                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DO NOT CHANGE THIS OR THE CMS WILL STOP FUNCTIONING!                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  uniqueServerID = "ZRQ-SMG-ZRQ-1259-01232003"                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Keep this key a secret.                                               '
'''''''''''''''''''''''''''''' INFORMATION ''''''''''''''''''''''''''''''

'30 minute inactivity timeout.
timeoutSeconds = 30*60

'Location of database file.
dbFileName = "customerZRQ.mdb"
regDBFile = "D:\inetpub\vasst_com\registration\admin\data\" & dbFileName
'regDBFile = "c:\webroot\modprobe\www\sundance\admin\" & dbFileName

Dim outgoingMailServer, fromAddress, registrationConfirmationAddress, titlePos, titleID, titleName, titleDate, titleCity, titleisVisible, titleRemove, errIncorrectDate, errNoEvents, errNoEventsRemove, errNoSeminarsActive, displayWidth, displayHeight, errNoSeminarsAvailable, noteAboutAddingSeminars
Dim isMenuShowing

Function loadOptions
	openDatabase
	Set dbOptions=dbConnection.Execute("SELECT * FROM tblOptions")
	outgoingMailServer = dbOptions("outgoingMailServer")
	fromAddress = dbOptions("fromAddress")
	registrationConfirmationAddress = dbOptions("registrationConfirmationAddress")
	titlePos = dbOptions("titlePos")
	titleID = dbOptions("titleID")
	titleName = dbOptions("titleName")
	titleDate = dbOptions("titleDate")
	titleCity = dbOptions("titleCity")
	titleisVisible = dbOptions("titleisVisible")
	titleRemove = dbOptions("titleRemove")
	errIncorrectDate = dbOptions("errIncorrectDate")
	errNoEvents = dbOptions("errNoEvents")
	errNoEventsRemove = dbOptions("errNoEventsRemove")
	errNoSeminarsActive = dbOptions("errNoSeminarsActive")
	displayWidth = dbOptions("displayWidth")
	displayHeight = dbOptions("displayHeight")
	errNoSeminarsAvailable = "No seminars available, please check back later."
	closeDatabase
End Function

Function saveOptions
	openDatabase
	outgoingMailServer = Replace(Request.Form("optionOutgoingMailServer"),"'","&#39;")
	fromAddress = Replace(Request.Form("optionFromAddress"),"'","&#39;")
	registrationConfirmationAddress = Replace(Request.Form("optionRegistrationConfirmationAddress"),"'","&#39;")
	titlePos = Replace(Request.Form("optionTitlePos"),"'","&#39;")
	titleID = Replace(Request.Form("optionTitleID"),"'","&#39;")
	titleName = Replace(Request.Form("optionTitleName"),"'","&#39;")
	titleDate = Replace(Request.Form("optionTitleDate"),"'","&#39;")
	titleCity = Replace(Request.Form("optionTitleCity"),"'","&#39;")
	titleisVisible = Replace(Request.Form("optionTitleisVisible"),"'","&#39;")
	titleRemove = Replace(Request.Form("optionTitleRemove"),"'","&#39;")
	errIncorrectDate = Replace(Request.Form("optionErrIncorrectDate"),"'","&#39;")
	errNoEvents = Replace(Request.Form("optionErrNoEvents"),"'","&#39;")
	errNoEventsRemove = Replace(Request.Form("optionErrNoEventsRemove"),"'","&#39;")
	errNoSeminarsActive = Replace(Request.Form("optionErrNoSeminarsActive"),"'","&#39;")
	displayWidth = Replace(Request.Form("optionDisplayWidth"),"'","&#39;")
	displayHeight = Replace(Request.Form("optionDisplayHeight"),"'","&#39;")
	dbConnection.Execute("UPDATE tblOptions SET outgoingMailServer = '" & outgoingMailServer & "', fromAddress = '" & fromAddress & "', registrationConfirmationAddress = '" & registrationConfirmationAddress & "', titlePos = '" & titlePos & "', titleID = '" & titleID & "', titleName = '" & titleName & "', titleDate = '" & titleDate & "', titleCity = '" & titleCity & "', titleisVisible = '" & titleisVisible & "', titleRemove = '" & titleRemove & "', errIncorrectDate = '" & errIncorrectDate & "', errNoEvents = '" & errNoEvents & "', errNoEventsRemove = '" & errNoEventsRemove & "', errNoSeminarsActive = '" & errNoSeminarsActive & "', displayWidth = " & displayWidth & ", displayHeight = " & displayHeight & "")
	closeDatabase
End Function

Dim loginMessage
Function checkAuth
	If (authPage) Then
		'Check to see if we are already authenticated.
		If (Request.Cookies("userAuth")("loggedIn") <> "yes") Then
			If Not (Request.Form("Login") = "") Then
				attemptUser = Request.Form("username")
				attemptPass = Request.Form("password")
				If (attemptUser = "") Or (attemptPass = "") Then
					loginMessage = "Enter a username and password."
				Else
					loginID = checkUser(attemptUser, attemptPass)
					If (loginID > -1) Then
						Response.Cookies("userAuth")("loggedIn") = "yes"
						Response.Cookies("userAuth")("userID") = loginID
						sendToMenu = true
					Else
						loginMessage = "Enter a valid username and password."
					End If
				End If
			End If
		Else
			sendToLogin = true
		End If
	Else
		If (Request.Cookies("userAuth")("loggedIn") <> "yes") Then
			sendToLogin = true
		End If
	End If

	'If an invalid sessionID was found or the session has timed out, log the user out.
	If (sendToLogin) Then
		If ((Request.QueryString("i") = "") And (Request.QueryString("c") = "") And (Request.QueryString("o") = "")) Then
			newQueryString = ""
		Else
			newQueryString = "&i=" & Request.QueryString("i") & "&c=" & Request.QueryString("c") & "&o=" & Request.QueryString("o")
		End If
		If (Request.QueryString("page") = "") Then
			Response.Redirect("login.asp?page=" & Right(Request.ServerVariables("URL"),Len(Request.ServerVariables("URL"))-InStrRev(Request.ServerVariables("URL"),"/")) & newQueryString)
		Else
			Response.Redirect("login.asp?page=" & Request.QueryString("page") & newQueryString)
		End If
	End If

	If (sendToMenu) Then
		If ((Request.QueryString("i") = "") And (Request.QueryString("c") = "") And (Request.QueryString("o") = "")) Then
			newQueryString = ""
		Else
			newQueryString = "&i=" & Request.QueryString("i") & "&c=" & Request.QueryString("c") & "&o=" & Request.QueryString("o")
		End If
		If Not (Request.QueryString("page") = "") Then
			Response.Redirect("index.asp?page=" & Request.QueryString("page") & newQueryString)
		Else
			Response.Redirect("index.asp?page=about.asp" & newQueryString)
		End If
	End If

' DEBUGGING CODE '
'	SplitCookies
'	SplitForm
'	SplitServerVariables
' DEBUGGING CODE END '
End Function

Function checkUser( username, password )
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

'	loadUserDB
'	loginCheck = -1

	Set userData = dbConnection.Execute("SELECT * FROM tblUsers WHERE strUsername = '" & username & "'")
	If userData.EOF Then
		result = "Invalid Username"
		loginCheck = -1
	Else
		If (userData("strPassword") = MD5(uniqueServerID & password)) Then
			If (lcase(userData("optAccessLevel")) = "deny") Then
				result = "Access Is Denied"
				loginCheck = -1
			Else
				result = "Login Successful"
				loginCheck = userData("numUserID")
			End If
		Else
			result = "Invalid Password"
			loginCheck = -1
		End If
	End If

	logTheLogin username, MD5(uniqueServerID & password), result
	checkUser = loginCheck

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function checkPassword( password )
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set userData = dbConnection.Execute("SELECT strPassword FROM tblUsers WHERE numUserID = " & getUserID & "")
	If userData.EOF Then
		checkPassword = false
	Else
		If (MD5(uniqueServerID & password) = userData("strPassword")) Then
			checkPassword = true
		Else
			checkPassword = false
		End If
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function checkAvailUser( username )
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set userData = dbConnection.Execute("SELECT numUserID FROM tblUsers WHERE strUsername = '" & username & "'")
	If userData.EOF Then
		checkAvailUser = true
	Else
		checkAvailUser = false
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function createUser( username )
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("INSERT INTO tblUsers ( strUsername, strPassword, strRealName, strEmailAddress, optAccessLevel ) VALUES ( '" & username & "', '" & MD5(uniqueServerID & username) & "', '" & username & "', '" & username & "', 'Allow' )")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function removeUser( userID )
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("DELETE * FROM tblUsers WHERE numUserID = " & userID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function getUserID
	getUserID = Request.Cookies("userAuth")("userID")
End Function

Function logTheLogin( logUsername, logPasswordHash, logResult )
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("INSERT INTO tblAccessLog ( loginDate, loginUsername, loginPasswordHash, loginIP, loginResult ) " & _
	                                       "VALUES ( #" & date & " " & time & "#, '" & logUsername & "', '" & logPasswordHash & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & logResult & "' )")
	Set getlastLoginID = dbConnection.Execute("SELECT max(loginID) as getlastLoginID FROM tblAccessLog WHERE loginUsername = '" & logUsername & "'")
	getlastLoginID = getlastLoginID("getlastLoginID")

	If (logResult = "Login Successful") Then
		dbConnection.Execute("UPDATE tblUsers SET lastLoginID = '" & getlastLoginID & "' WHERE strUsername = '" & logUsername & "'")
	Else
		dbConnection.Execute("UPDATE tblUsers SET lastFailedLoginID = '" & getlastLoginID & "' WHERE strUsername = '" & logUsername & "'")
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printUsers(whichUser)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set userData = dbConnection.Execute("SELECT * FROM tblUsers WHERE numUserID <> " & getUserID & " AND strUsername <> 'jwalker'")
	Do Until userData.EOF
		If (Fix(whichUser) = Fix(userData("numUserID"))) Then
			Response.Write(tabTo(3) & "<OPTION VALUE=""" & userData("numUserID") & """ SELECTED>" & userData("strUsername") & "</OPTION>")
		Else
			Response.Write(tabTo(3) & "<OPTION VALUE=""" & userData("numUserID") & """>" & userData("strUsername") & "</OPTION>")
		End IF
		userData.MoveNext
	Loop

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function getUserInfo(userID,whatInfo)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	If Not (userID = "") Then
		Select Case lcase(whatInfo)
		Case "username"
			whatInfo = "strUsername"
		Case "password"
			whatInfo = "strPassword"
		Case "realname"
			whatInfo = "strRealName"
		Case "email"
			whatInfo = "strEmailAddress"
		Case "accesslevel"
			whatInfo = "optAccessLevel"
		Case "lastlogin"
			whatInfo = "lastLoginID"
		Case "lastfailedlogin"
			whatInfo = "lastFailedLoginID"
		Case Else
			whatInfo = ""
		End Select

		If Not (whatInfo = "") Then
			Set userData = dbConnection.Execute("SELECT " & whatInfo & " AS returnedInfo FROM tblUsers WHERE numUserID = " & userID & "")
			If Not userData.EOF Then
				If Not (userData("returnedInfo") = "") Then
					getUserInfo = userData("returnedInfo")
				Else
					getUserInfo = 0
				End If
			End If
		End If
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function setUserInfo(userID,whatInfo,newInfo)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Select Case lcase(whatInfo)
	Case "username"
		setString = "strUsername = '" & newInfo & "'"
	Case "password"
		setString = "strPassword = '" & MD5(uniqueServerID & newInfo) & "'"
	Case "realname"
		setString = "strRealName = '" & newInfo & "'"
	Case "email"
		setString = "strEmailAddress = '" & newInfo & "'"
	Case "accesslevel"
		setString = "optAccessLevel = '" & newInfo & "'"
	Case "lastlogin"
		setString = "lastLoginID = '" & newInfo & "'"
	Case "lastfailedlogin"
		setString = "lastFailedLoginID = '" & newInfo & "'"
	Case Else
		setString = ""
	End Select

	If Not (setString = "") Then
		dbConnection.Execute("UPDATE tblUsers SET " & setString & " WHERE numUserID = " & userID & "")
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printLastLoginInfo
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set loginData = dbConnection.Execute("SELECT loginDate, loginIP FROM tblAccessLog WHERE loginID < " & getUserInfo(getUserID,"lastlogin") & " AND loginUsername = '" & getUserInfo(getUserID,"username") & "' ORDER BY loginDate DESC")
	If loginData.EOF Then
		Response.Write "Never."
	Else
		Response.Write "Logged in at " & loginData("loginDate") & " from " & loginData("loginIP") & "."
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printLastFailedLoginInfo
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set loginData = dbConnection.Execute("SELECT * FROM tblAccessLog WHERE loginID = " & getUserInfo(getUserID,"lastfailedlogin") & " AND loginUsername = '" & getUserInfo(getUserID,"username") & "'")
	If loginData.EOF Then
		Response.Write "Never."
	Else
		Response.Write "Failed at " & loginData("loginDate") & " from " & loginData("loginIP") & "."
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function logOut
	Response.Cookies("userAuth") = ""
	Response.Redirect("login.asp")
End Function

Function printHeader
	Response.Write(tabTo(0) & "<HTML>")
	Response.Write(tabTo(1) & "<HEAD>")
	Response.Write(tabTo(2) & "<STYLE>")
	Response.Write(tabTo(3) & "BODY { Overflow: auto; Margin: 0px; Font-Family: 'Tahoma'; Font-Size: 10pt; Background: #C0C0C0; Color: black; }")
	Response.Write(tabTo(3) & "TD { Color: black; Font-Family: 'Tahoma'; Font-Size: 10pt; }")
	Response.Write(tabTo(3) & "TD.menu { Color: black; Font-Family: 'Tahoma'; Font-Size: 10pt; }")
	Response.Write(tabTo(3) & "A.menu:visited { Color: blue; Text-Decoration: none; }")
	Response.Write(tabTo(3) & "A.menu:link { Color: blue; Text-Decoration: none; }")
	Response.Write(tabTo(3) & "A.menu:hover { Color: blue; Text-Decoration: underline; }")
	Response.Write(tabTo(3) & "A:visited { Color: black; Text-Decoration: none; }")
	Response.Write(tabTo(3) & "A:link { Color: black; Text-Decoration: none; }")
	Response.Write(tabTo(3) & "A:hover { Color: black; Text-Decoration: underline; }")
'	Response.Write(tabTo(3) & "INPUT { Color: black; Font-Family: 'Tahoma'; Font-Size: 8pt; }")
	Response.Write(tabTo(3) & "SELECT { Background: #CFCFCF; Color: black; Font-Family: 'Tahoma'; Font-Size: 8pt; Border: #000000 1px solid; }")
	Response.Write(tabTo(3) & "INPUT { Background: #CFCFCF; Color: black; Font-Family: 'Tahoma'; Font-Size: 8pt; Border: #000000 1px solid; }")
	Response.Write(tabTo(3) & "INPUT.normal { Background: transparent; Border: 0px; }")
	Response.Write(tabTo(3) & ".topMenu { Color: black; }")
	Response.Write(tabTo(3) & ".topOrangeLink { Color: #A0A0A0; }")
	Response.Write(tabTo(3) & ".orange { Background: #E5E5E5; Color: black; Font-Weight: bold; }")
	Response.Write(tabTo(3) & ".orangebox { Background: #E5E5E5; Color: black; Border: 1px solid black; Font-Weight: bold; }")
	Response.Write(tabTo(3) & ".orangedivider { Padding: 4px; Background: #FFCC66; Color: black; Border-Top: 2px solid #E5E5E5; Border-Bottom: 2px solid #E5E5E5; }")
	Response.Write(tabTo(3) & ".orangeborder { Border: 1px solid black; }")
	Response.Write(tabTo(3) & ".blackborder { Background: #E5E5E5; Color: black; Border: 1px solid black; Font-Weight: bold; Padding: 3px; }")
	Response.Write(tabTo(3) & ".blackbold { Color: black; Font-Weight: bold; }")
	Response.Write(tabTo(3) & ".blackbox { Color: black; Background: white; Border: 1px solid #E5E5E5; Font-Weight: bold; Padding: 3px; }")
	Response.Write(tabTo(3) & ".even { Background: #999999; Color: black; }")
	Response.Write(tabTo(3) & ".odd { Background: #CCCCCC; Color: black; }")
	Response.Write(tabTo(3) & ".header { Background: #000000; Color: white; Font-Weight: bold; }")
	Response.Write(tabTo(3) & ".bartitle { Background: #E5E5E5; Color: black; Border: 1px solid black; Font-Weight: bold; Padding: 3px; }")
	Response.Write(tabTo(3) & ".bar { Background: #E5E5E5; Color: black; Border-Top: 1px #000000 solid; Border-Bottom: 1px #000000 solid; }")
	Response.Write(tabTo(3) & ".error { Color: red; }")
	Response.Write(tabTo(2) & "</STYLE>")
	Response.Write(tabTo(2) & "<TITLE>" & pageTitle & "</TITLE>")
	Response.Write(tabTo(1) & "</HEAD>")
	Response.Write(tabTo(1) & "<BODY>")
End Function

Function printMenu
	If Not (authPage = True) Then
		Response.Write(tabTo(5) & "<SCRIPT>")
		Response.Write(tabTo(6) & "function openPage(page) {")
		Response.Write(tabTo(6) & "  this.location.href = ""index.asp?page="" + page + ""&searchFor="" + escape(document.getElementById(""seminarSelector"").value);")
		Response.Write(tabTo(6) & "}")
		Response.Write(tabTo(6) & "function openWindow(page) {")
		Response.Write(tabTo(6) & "  window.open(page + ""?searchFor="" + escape(document.getElementById(""seminarSelector"").value));")
		Response.Write(tabTo(6) & "}")
		Response.Write(tabTo(5) & "</SCRIPT>")
		Response.Write(tabTo(5) & "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>")
		Response.Write(tabTo(6) & "<TR>")
		Response.Write(tabTo(7) & "<TD ROWSPAN=2 CLASS=orangebox ALIGN=center VALIGN=middle>")

		Response.Write(tabTo(8) & "<HR><U>Restrict By Seminar:</U><BR>")
		printSeminarNames "select"
		Response.Write("*<br>")

		accessLevel = getUserInfo(getUserID,"accesslevel")

		Response.Write(tabTo(8) & "<HR><U>Customer Manager</U><BR>")
		Response.Write(tabTo(8) & "<A HREF=""javascript:openPage('customer_list.asp')"" CLASS=menu TITLE='Lists customers.'><SPAN CLASS=topMenu>List Customers*</SPAN></A><BR>")
		If (accessLevel  = "Admin") Then
			Response.Write(tabTo(8) & "<A HREF=index.asp?page=customer_add.asp CLASS=menu TITLE='Add a customer.'><SPAN CLASS=topMenu>Add Customer</SPAN></A><BR>")
			Response.Write(tabTo(8) & "<A HREF=""javascript:openPage('customer_edit.asp')"" CLASS=menu TITLE='Edit an existing customer.'><SPAN CLASS=topMenu>Edit Customer*</SPAN></A><BR>")
			Response.Write(tabTo(8) & "<A HREF=index.asp?page=customer_purge.asp CLASS=menu TITLE='This option will purge all Deleted? flagged customers.'><SPAN CLASS=topMenu>Purge Customers</SPAN></A><BR>")
		End If
		Response.Write(tabTo(8) & "<A HREF=""javascript:openWindow('customer_spreadsheet.asp')"" CLASS=menu TITLE='View customers in a spreadsheet format.'><SPAN CLASS=topMenu>View Spreadsheet*</SPAN></A><BR>")
		If (accessLevel  = "Admin") Then
			Response.Write(tabTo(8) & "<A HREF=""javascript:openWindow('customer_editspreadsheet.asp')"" CLASS=menu TITLE='Edit customers in a spreadsheet format.'><SPAN CLASS=topMenu>Edit Spreadsheet*</SPAN></A><BR>")
		End If
		Response.Write(tabTo(8) & "<BR>")
		Response.Write(tabTo(8) & "<A HREF=index.asp?page=reports.asp CLASS=menu TITLE='Run queries on the customer database.'><SPAN CLASS=topMenu>Design Reports</SPAN></A><BR>")
		Response.Write(tabTo(8) & "<A HREF=index.asp?page=verisignlog.asp CLASS=menu TARGET=_verisignlog TITLE='Shows all attempted actions with verisign.'><SPAN CLASS=topMenu>Verisign Log</SPAN></A><BR>")
		If (accessLevel  = "Admin") Then
			Response.Write(tabTo(8) & "<BR>")
			Response.Write(tabTo(8) & "<A HREF=index.asp?page=emailer.asp CLASS=menu TITLE='Seminar list mailer.'><SPAN CLASS=topMenu>List Mailer</SPAN></A><BR>")
		End If
		Response.Write(tabTo(8) & "<HR><U>Seminar Manager</U><BR>")
		'printSeminarNames "select"
		'Response.Write("*<br>")
		Response.Write(tabTo(8) & "<A HREF=""javascript:openPage('seminars_list.asp')"" CLASS=menu TITLE='Lists active seminars.'><SPAN CLASS=topMenu>List Seminars*</SPAN></A><BR>")
		If (accessLevel = "Admin") Then
			Response.Write(tabTo(8) & "<A HREF=""javascript:openPage('seminars_add.asp')"" CLASS=menu TITLE='Add a seminar.'><SPAN CLASS=topMenu>Add Seminars*</SPAN></A><BR>")
			Response.Write(tabTo(8) & "<A HREF=""javascript:openPage('seminars_edit.asp')"" CLASS=menu TITLE='Edit an existing seminar.'><SPAN CLASS=topMenu>Edit Seminars*</SPAN></A><BR>")
			Response.Write(tabTo(8) & "<A HREF=""index.asp?page=seminars_hide.asp"" CLASS=menu TITLE='Hide a seminar.'><SPAN CLASS=topMenu>Hide Seminars</SPAN></A><BR>")
			Response.Write(tabTo(8) & "<A HREF=""index.asp?page=seminars_remove.asp"" CLASS=menu TITLE='Remove a seminar permanently.'><SPAN CLASS=topMenu>Remove Seminars</SPAN></A><BR>")
			Response.Write(tabTo(8) & "<BR>")
			Response.Write(tabTo(8) & "<A HREF=""index.asp?page=seminars_prices.asp"" CLASS=menu TITLE='Change seminar pricing.'><SPAN CLASS=topMenu>Seminar Pricing</SPAN></A><BR>")
			Response.Write(tabTo(8) & "<A HREF=""index.asp?page=seminars_info.asp"" CLASS=menu TITLE='Change seminar info URLs that are displayed on the registration seminar select page.'><SPAN CLASS=topMenu>Seminar Info URLs</SPAN></A><BR>")
			Response.Write(tabTo(8) & "<A HREF=""index.asp?page=seminars_time.asp"" CLASS=menu TITLE='Change start and stop times for seminars.'><SPAN CLASS=topMenu>Seminar Times</SPAN></A><BR>")
			Response.Write(tabTo(8) & "<A HREF=""index.asp?page=seminars_hearabout.asp"" CLASS=menu TITLE='Change hear about options.'><SPAN CLASS=topMenu>Hear About Admin</SPAN></A><BR>")
		End If
		Response.Write(tabTo(8) & "<HR><U>Links</U><BR>")
		Response.Write(tabTo(8) & "<A HREF=""../register.asp"" CLASS=menu TARGET=_blank><SPAN CLASS=topMenu>Registration Form</SPAN></A><BR>")
		If (lcase(getUserInfo(getUserID,"username")) = "jwalker") Then
			Response.Write(tabTo(8) & "<HR><U>Administration</U><BR>")
			Response.Write(tabTo(8) & "<A HREF=index.asp?page=sqltest.asp CLASS=menu TARGET=_blank><SPAN CLASS=topMenu>SQL Admin</SPAN></A><BR>")
			Response.Write(tabTo(8) & "<A HREF=index.asp?page=email.asp CLASS=menu><SPAN CLASS=topMenu>Email</SPAN></A><BR>")
		End If
		Response.Write(tabTo(8) & "<HR><U>Options</U><BR>")
		Response.Write(tabTo(8) & "<A HREF=index.asp?page=about.asp CLASS=menu><SPAN CLASS=topMenu>About</SPAN></A><BR>")
		If (accessLevel = "Admin") Then
			Response.Write(tabTo(8) & "<A HREF=index.asp?page=options.asp CLASS=menu><SPAN CLASS=topMenu TITLE='Click to change options.'>Options</SPAN></A><BR>")
		End If
		Response.Write(tabTo(8) & "<A HREF=index.asp?page=edituser.asp CLASS=menu><SPAN CLASS=topMenu TITLE='Click to change username and password.'>Edit User</SPAN></A><BR>")
		Response.Write(tabTo(8) & "<A HREF=index.asp?page=logout.asp CLASS=menu><SPAN CLASS=topMenu>Log Out</SPAN></A><BR>")
		Response.Write(tabTo(7) & "</TD>")
		Response.Write(tabTo(7) & "<TD CLASS=blackborder STYLE='Border-Left: 0px;'>")
		Response.Write(tabTo(8) & "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH='100%'>")
		Response.Write(tabTo(9) & "<TR>")
		Response.Write(tabTo(10) & "<TD CLASS=orange ALIGN=right>")
		Response.Write(tabTo(11) & "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH='100%'>")
		Response.Write(tabTo(12) & "<TR>")
		Response.Write(tabTo(13) & "<TD NOWRAP WIDTH=90% CLASS=blackbold>&nbsp;Customer Management System</TD>")
		Response.Write(tabTo(13) & "<TD NOWRAP CLASS=blackbold>Logged in:</TD>")
		Response.Write(tabTo(13) & "<TD NOWRAP CLASS=blackbox><A HREF='index.asp?page=edituser.asp' TITLE='Click to change username and password.'><SPAN CLASS=topOrangeLink>" & getUserInfo(getUserID,"username") & "</SPAN></A></TD>")
		Response.Write(tabTo(12) & "</TR>")
		Response.Write(tabTo(11) & "</TABLE>")
		Response.Write(tabTo(10) & "</TD>")
		Response.Write(tabTo(9) & "</TR>")
		Response.Write(tabTo(8) & "</TABLE>")
		Response.Write(tabTo(7) & "</TD>")
		Response.Write(tabTo(6) & "</TR>")
		Response.Write(tabTo(6) & "<TR>")
		Response.Write(tabTo(7) & "<TD CLASS=orangebox STYLE='Border-Left: 0px;'>")
	End If
End Function

Function tabTo(times)
	tabTo = vbCRLF
	For xix = 1 to times
		tabTo = tabTo & "  "
	Next
End Function

Function printBox
	Response.Write(tabTo(2) & "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 ALIGN=center>")
	Response.Write(tabTo(3) & "<TR>")
	Response.Write(tabTo(4) & "<TD CLASS=orangeborder>")
End Function

Function printBoxClose
	If Not (authPage = True) Then
		Response.Write(tabTo(7) & "</TD>")
		Response.Write(tabTo(6) & "</TR>")
		Response.Write(tabTo(5) & "</TABLE>")
	End If
	Response.Write(tabTo(4) & "</TD>")
	Response.Write(tabTo(3) & "</TR>")
	Response.Write(tabTo(2) & "</TABLE>")
End Function

Function printFooter
	Response.Write(tabTo(1) & "</BODY>")
	Response.Write(tabTo(0) & "</HTML>")
End Function

Dim dbConnection
Dim isDbOpen
Dim wasDbOpen

Function openDatabase
	set dbConnection = Server.CreateObject("ADODB.Connection")
'	connectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & dbFile & ""
	connectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & regDBFile & ""
	dbConnection.Open connectionString
	isDbOpen = True
End Function

Function closeDatabase
	dbConnection.Close
	isDBOpen = False
End Function

Function printSeminars
	'Open the database for Querys to be run.
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	'Create the heading to display data.
	Response.Write(tabTo(2) & "<TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0 WIDTH='100%'>")
	Response.Write(tabTo(3) & "<TR CLASS=header>")
	Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=1% ALIGN=center NOWRAP>" & titlePos & "</TD>")
'	Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=1% ALIGN=center NOWRAP>" & titleID & "</TD>")
	Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=30% NOWRAP>" & titleName & "</TD>")
	Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=30% NOWRAP>" & titleDate & "</TD>")
	Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=30% NOWRAP>" & titleCity & "</TD>")

	'If request is from Remove page, change the title.
	If (permanentlyRemove) Then
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=4% ALIGN=center NOWRAP>" & titleRemove & "</TD>")
	Else
		tmpPage = tmpPage & _
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=4% ALIGN=center NOWRAP>" & titleisVisible & "</TD>")
	End If

	'Add column if this is a request from the Edit page and end the table line.
	If (addEditButton) Then
    	Response.Write(tabTo(4) & "<TD WIDTH=4% ALIGN=center NOWRAP>Edit</TD>")
    	Response.Write(tabTo(3) & "</TR>")
	Else
		Response.Write(tabTo(3) & "</TR>")
	End If

	'Select the query if this is a request from the Delete page.
	If (permanentlyRemove) Then
		'Load the tblSeminars table and hide the active Seminars
		Set dbEventData=dbConnection.Execute("SELECT * FROM tblSeminars WHERE isVisible = False AND isDeleted = False AND (linkedToSeminar = -1 OR linkedToSeminar = numSeminarID) ORDER BY strSeminarName, dateSeminarDate")
	Else
		If (Request.QueryString("searchFor") = "") Or (Request.QueryString("searchFor") = "All") Then
			'Show everything in the tblSeminars
			Set dbEventData=dbConnection.Execute("SELECT * FROM tblSeminars WHERE isDeleted = False AND (linkedToSeminar = -1 OR linkedToSeminar = numSeminarID) ORDER BY strSeminarName, dateSeminarDate")
		Else
			Set dbEventData=dbConnection.Execute("SELECT * FROM tblSeminars WHERE isDeleted = False AND strSeminarName = '" & Request.QueryString("searchFor") & "' AND (linkedToSeminar = -1 OR linkedToSeminar = numSeminarID) ORDER BY strSeminarName, dateSeminarDate")
		End If
	End If

	'Initialize the Position Counter
	Pos = 0

	If (addDeleteButton = True) Then
		Response.Write("<FORM METHOD=post>")
	End If

	While Not dbEventData.EOF
		Pos = Pos + 1
		If bgclass = "odd" Then bgclass = "even" Else bgclass = "odd" End If
		If (addEditButton = True) and (Cint(Request.Form("edit")) = Cint(dbEventData("numSeminarID"))) Then
			editingThisLine = True
    		Response.Write(tabTo(4) & "<FORM METHOD=post>")
			Response.Write(tabTo(4) & "<INPUT TYPE=hidden NAME=numSeminarID VALUE='" & dbEventData("numSeminarID") & "'>")
		Else
			editingThisLine = False
		End If
		Response.Write(tabTo(3) & "<TR>")
		Response.Write(tabTo(4) & "<TD ALIGN=center NOWRAP CLASS=" & bgclass & ">" & Pos & "</TD>")
'		Response.Write(tabTo(4) & "<TD ALIGN=center NOWRAP CLASS=" & bgclass & ">" & dbEventData("numSeminarID") & "</TD>")

		If (Fix(dbEventData("linkedToSeminar")) = -1) Then
			thisDate = dbEventData("dateSeminarDate")
		Else
			Set datesForEvent = dbConnection.Execute("SELECT * FROM tblSeminars WHERE isDeleted = False AND linkedToSeminar = " & dbEventData("linkedToSeminar") & " ORDER BY dateSeminarDate")

			thisDate = ""
			Do Until datesForEvent.EOF
				If (editingThisLine) Then
					thisDate = thisDate & "," & datesForEvent("dateSeminarDate")
				Else
					thisDate = thisDate & datesForEvent("dateSeminarDate") & "<br />"
				End If
				datesForEvent.MoveNext
			Loop
			If (editingThisLine) Then
				thisDate = Right(thisDate, Len(thisDate)-1)
			End If
		End If

		If (editingThisLine) Then
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=strSeminarName VALUE='" & dbEventData("strSeminarName") & "'></TD>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=dateSeminarDate VALUE='" & thisDate & "'></TD>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=strSeminarCity VALUE='" & dbEventData("strSeminarCity") & "'></TD>")
		Else
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">" & dbEventData("strSeminarName") & "</TD>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">" & thisDate & "</TD>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">" & dbEventData("strSeminarCity") & "</TD>")
		End If
		If (addDeleteButton) or (editingThisLine) Then
			If Not (permanentlyRemove) Then
				If (dbEventData("isVisible")) Then
					Checked = "CHECKED"
				Else
					Checked = ""
				End If
			End If
			If (editingThisLine) Then
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & " ALIGN=center><INPUT CLASS=normal TYPE=checkbox NAME=isVisible " & Checked & ">")
			Else
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & " ALIGN=center><INPUT TYPE=hidden NAME='removeid" & Pos & "' VALUE=" & dbEventData("numSeminarID") & "><INPUT CLASS=normal TYPE=checkbox NAME='cb" & Pos & "' " & Checked & ">")
			End If
		Else
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">" & dbEventData("isVisible") & "</TD>")
		End If

		If (addEditButton) Then
			Response.Write(tabTo(4) & "<TD WIDTH=4% ALIGN=center NOWRAP CLASS=" & bgclass & ">")
	    	If (Cint(Request.Form("edit")) = Cint(dbEventData("numSeminarID"))) Then
		    	Response.Write(tabTo(5) & "<INPUT TYPE=Submit NAME=saveChanges VALUE='Save'>")
			Else
	   			Response.Write(tabTo(5) & "<FORM METHOD=post>")
				Response.Write(tabTo(5) & "<INPUT TYPE=hidden NAME=edit VALUE=" & dbEventData("numSeminarID") & ">")
		    	Response.Write(tabTo(5) & "<INPUT TYPE=hidden NAME=eventName VALUE='" & dbEventData("strSeminarName") & "'>")
		    	Response.Write(tabTo(5) & "<INPUT TYPE=hidden NAME=dateSeminarDate VALUE='" & dbEventData("dateSeminarDate") & "'>")
		    	Response.Write(tabTo(5) & "<INPUT TYPE=hidden NAME=eventCity VALUE='" & dbEventData("strSeminarCity") & "'>")
		    	Response.Write(tabTo(5) & "<INPUT TYPE=hidden NAME=isVisible VALUE=" & dbEventData("isVisible") & ">")
		    	Response.Write(tabTo(5) & "<INPUT TYPE=Submit VALUE='Edit'>")
			End If
	    	Response.Write(tabTo(4) & "</TD>")
	    	Response.Write(tabTo(4) & "</FORM>")
		End If

		Response.Write(tabTo(3) & "</TR>" & vbNewline)
		dbEventData.MoveNext
	Wend

	If (addAddButton) Then
		Response.Write("<TR>")
		Pos = Pos + 1
		If bgclass = "odd" Then bgclass = "even" Else bgclass = "odd" End If
		Response.Write("<TD ALIGN=center NOWRAP CLASS=" & bgclass & ">" & Pos & "</TD>")
'		Response.Write("<TD ALIGN=center NOWRAP CLASS=" & bgclass & "><INPUT TYPE=hidden NAME=numSeminarID VALUE=" & getNextSeminarID & ">" & getNextSeminarID & "</INPUT></TD>")
		If (reloadValues) Then
			strSeminarName = Request.Form("strSeminarName")
			dateSeminarDate = Request.Form("dateSeminarDate")
			strSeminarCity = Request.Form("strSeminarCity")
		End If
		Response.Write("Add")
		If (reloadValues) Then
			If (strSeminarName = "") Then
				If (Request.QueryString("searchFor") = "All") Then
					Response.Write("<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=strSeminarName VALUE=''></INPUT></TD>")
				Else
					Response.Write("<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=strSeminarName VALUE='" & Request.QueryString("searchFor") & "'></INPUT></TD>")
				End If
			Else
				Response.Write("<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=strSeminarName VALUE='" & strSeminarName & "'></INPUT></TD>")
			End If
		Else
			Response.Write("<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=strSeminarName VALUE='" & Request.QueryString("searchFor") & "'></INPUT></TD>")
		End If
		Response.Write("<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=dateSeminarDate VALUE='" & dateSeminarDate & "'></INPUT></TD>")
		Response.Write("<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=strSeminarCity VALUE='" & strSeminarCity & "'></INPUT></TD>")
		Response.Write("<TD NOWRAP CLASS=" & bgclass & "><INPUT TYPE=submit VALUE='Add'></TD>")
		Response.Write("</TR>")
	End If


	If (Pos = 0) Then
		If (permanentlyRemove) Then
			Response.Write("<TR><TD COLSPAN=6 ALIGN=center>" & errNoEventsRemove & "</TD></TR>")
		Else
			Response.Write("<TR><TD COLSPAN=6 ALIGN=center>" & errNoEvents & "</TD></TR>")
		End If
	End If

	If (addDeleteButton = True) Then
		Response.Write("<TR>")
		Response.Write("<TD COLSPAN=6 ALIGN=right>")
		Response.Write("<INPUT TYPE=hidden NAME=howMany VALUE=" & Pos & ">")
		If Not (Pos = 0) Then
			If (permanentlyRemove) Then
				Response.Write("<INPUT TYPE=submit VALUE='Permanently Remove'>")
			Else
				Response.Write("<INPUT TYPE=submit VALUE='Update/Save'>")
			End If
		End If
		Response.Write("</FORM>")
	End If

	Response.Write(tabTo(3) & "</TR>")
	Response.Write(tabTo(2) & "</TABLE>")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printSeminarsPrices(method)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	If (InStr(method,"remove") > 0) Then
		dbConnection.Execute("DELETE * FROM tblPricing WHERE discountID = " & Right(method,Len(method)-InStrRev(method,"_")) & "")
	ElseIf (method = "save") Then
		For Each id In Split(Request.Form("idList"),",")
			If (Cint(id) = Cint(Request.Form("newID"))) Then
				If (Request.Form("priceType_" & id) = "normal") Then
					dbConnection.Execute("INSERT INTO tblPricing ( effectsSeminar, normalPrice ) VALUES ( '" & Request.Form("effectsSeminar_" & id) & "', '" & Request.Form("normalPrice_" & id) & "' )")
				Else
					dbConnection.Execute("INSERT INTO tblPricing ( effectsSeminar, discountPrice, discountCode ) VALUES ( '" & Request.Form("effectsSeminar_" & id) & "', '" & Request.Form("discountPrice_" & id) & "', '" & Ucase(Request.Form("discountCode_" & id)) & "' )")
				End If
			Else
				If (Request.Form("priceType_" & id) = "normal") Then
					dbConnection.Execute("UPDATE tblPricing SET effectsSeminar = '" & Request.Form("effectsSeminar_" & id) & "', normalPrice = " & Request.Form("normalPrice_" & id) & " WHERE discountID = " & id & "")
				Else
					dbConnection.Execute("UPDATE tblPricing SET effectsSeminar = '" & Request.Form("effectsSeminar_" & id) & "', discountPrice = " & Request.Form("discountPrice_" & id) & ", discountCode = '" & Ucase(Request.Form("discountCode_" & id)) & "' WHERE discountID = " & id & "")
				End If
			End If
		Next
	Else
		'Create the heading to display data.
		Response.Write(tabTo(2) & "<TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0 WIDTH='100%'>")
		Response.Write(tabTo(3) & "<TR CLASS=header>")
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=1% ALIGN=center NOWRAP>&nbsp;#&nbsp;</TD>")
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=30% NOWRAP>" & titleName & "</TD>")
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=30% NOWRAP>Normal Cost</TD>")
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=30% NOWRAP>Discount Cost</TD>")
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=30% NOWRAP>Discount Code</TD>")
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=30% NOWRAP>Remove</TD>")
	   	Response.Write(tabTo(3) & "</TR>")

		Set dbEventData = dbConnection.Execute("SELECT * FROM tblPricing ORDER BY effectsSeminar, discountPrice, normalPrice, discountCode")

		'Initialize the Position Counter
		Pos = 0

		Response.Write("<FORM METHOD=post>")

		While Not dbEventData.EOF
			Pos = Pos + 1
			If (Pos = 1) Then
				idList = dbEventData("discountID")
			Else
				idList = idList & "," & dbEventData("discountID")
			End If
			If bgclass = "odd" Then bgclass = "even" Else bgclass = "odd" End If
			Response.Write(tabTo(3) & "<TR>")
			Response.Write(tabTo(4) & "<TD ALIGN=center NOWRAP CLASS=" & bgclass & ">" & Pos & "</TD>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><INPUT TYPE=hidden NAME=effectsSeminar_" & dbEventData("discountID") & " VALUE=""" & dbEventData("effectsSeminar") & """>" & dbEventData("effectsSeminar") & "</TD>")
			If (dbEventData("normalPrice") <> "") Then
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><input type=""hidden"" name=""priceType_" & dbEventData("discountID") & """ value=""normal"">$<INPUT NAME=normalPrice_" & dbEventData("discountID") & " VALUE='" & dbEventData("normalPrice") & "' SIZE=10></TD>")
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">&nbsp;</TD>")
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">&nbsp;</TD>")
			Else
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><input type=""hidden"" name=""priceType_" & dbEventData("discountID") & """ value=""discount"">&nbsp;</TD>")
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">$<INPUT NAME=discountPrice_" & dbEventData("discountID") & " VALUE='" & dbEventData("discountPrice") & "' SIZE=10></TD>")
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=discountCode_" & dbEventData("discountID") & " VALUE='" & dbEventData("discountCode") & "'></TD>")
			End If
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><INPUT TYPE=""submit"" NAME=remove_" & dbEventData("discountID") & " VALUE='Remove'></TD>")
			Response.Write(tabTo(3) & "</TR>" & vbNewline)
			dbEventData.MoveNext
		Wend

		Set dbEventData = Nothing

		If (method = "add") Then
			Pos = Pos + 1
			If bgclass = "odd" Then bgclass = "even" Else bgclass = "odd" End If

			Set dbCheck = dbConnection.Execute("SELECT effectsSeminar FROM tblPricing WHERE effectsSeminar = '" & Request.Form("seminarSelector") & "'")
			If dbCheck.EOF Then
				addType = "normal"
			Else
				addType = "discount"
			End If
			Set dbCheck = Nothing

			Set dbMaxRecord = dbConnection.Execute("SELECT max(discountID) as discountID FROM tblPricing")
			nextPricing = dbMaxRecord("discountID")+1
			Set dbMaxRecord = Nothing
			idList = idList & "," & nextPricing

			Response.Write(tabTo(3) & "<TR>")
			Response.Write(tabTo(4) & "<TD ALIGN=center NOWRAP CLASS=" & bgclass & ">" & Pos & "</TD>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><input type=""hidden"" name=""effectsSeminar_" & nextPricing & """ value=""" & Request.Form("seminarSelector") & """>" & Request.Form("seminarSelector") & "</TD>")
			Response.Write(tabTo(4) & "<input type=""hidden"" name=""newID"" value=""" & nextPricing & """>")
			If (addType = "normal") Then
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><input type=""hidden"" name=""priceType_" & nextPricing & """ value=""normal"">$<INPUT NAME=normalPrice_" & nextPricing & " VALUE='0' SIZE=10></TD>")
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">&nbsp;</TD>")
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">&nbsp;</TD>")
			Else
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><input type=""hidden"" name=""priceType_" & nextPricing & """ value=""discount"">&nbsp;</TD>")
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">$<INPUT NAME=discountPrice_" & nextPricing & " VALUE='0' SIZE=10></TD>")
				Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=discountCode_" & nextPricing & " VALUE=''></TD>")
			End If
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">&nbsp;</TD>")
			Response.Write(tabTo(3) & "</TR>" & vbNewline)
		End If

		If (Pos = 0) Then
			Response.Write("<TR><TD COLSPAN=6 ALIGN=center>" & errNoEvents & "</TD></TR>")
		Else
			Response.Write("<TR>")
			Response.Write("<TD COLSPAN=6 ALIGN=right>")
			Response.Write("<INPUT TYPE=hidden NAME=idList VALUE=" & idList & ">")
			Response.Write("<INPUT TYPE=submit NAME='cancelChanges' VALUE='Cancel'>")
			Response.Write("<INPUT TYPE=reset VALUE='Reset'>")
			Response.Write("<INPUT TYPE=submit NAME='saveChanges' VALUE='Update/Save'>")
			Response.Write(tabTo(3) & "</TR>")
		End If

		If (method <> "add") Then
			Response.Write("<tr><td colspan=6>")
			Response.Write("<center><b>Add a new seminar price:</b><br>")

			printSeminarNameList(-1)

			Response.Write("<INPUT TYPE=submit NAME='addPrice' VALUE='Add New Price'>")
			Response.Write("</center></td></tr>")
		End If

		Response.Write("</FORM>")
		Response.Write(tabTo(2) & "</TABLE>")
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printAboutSeminars(seminar,id)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Response.Write("<select size=1 NAME=effectsSeminar_" & id & ">")
	Set dbSeminars = dbConnection.Execute("SELECT strSeminarName FROM tblSeminars GROUP BY strSeminarName")
	If (seminar = "All") Then
		Response.Write("<option value=""All"" selected>All</option>" & vbNewline)
	Else
		Response.Write("<option value=""All"">All</option>" & vbNewline)
	End If

	Do Until dbSeminars.EOF
		If (dbSeminars("strSeminarName") = seminar) Then
			thisSelected = " SELECTED"
		Else
			thisSelected = ""
		End If
		Response.Write("<option value=""" & dbSeminars("strSeminarName") & """" & thisSelected & ">" & dbSeminars("strSeminarName") & "</option>" & vbNewline)
		dbSeminars.MoveNext
	Loop

	Response.Write("</select>" & vbNewline)
	Set dbSeminars = Nothing

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printHearAboutOptions(selected,tour)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set hearAboutData = dbConnection.Execute("SELECT aboutText FROM tblHearAbout WHERE effectsSeminar = '" & tour & "' OR effectsSeminar = 'All' ORDER BY aboutText")
	Do Until hearAboutData.EOF
		If (selected = hearAboutData("aboutText")) Then
			thisSelected = " SELECTED"
		Else
			thisSelected = ""
		End If
		Response.Write(tabTo(5) & "<option value=""" & hearAboutData("aboutText") & """" & thisSelected & ">" & hearAboutData("aboutText") & "</option>" & vbNewline)
		hearAboutData.MoveNext
	Loop

	Set hearAboutData = Nothing

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printHearAbout(method)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	If (InStr(method,"remove") > 0) Then
		dbConnection.Execute("DELETE * FROM tblHearAbout WHERE aboutID = " & Right(method,Len(method)-InStrRev(method,"_")) & "")
	ElseIf (method = "save") Then
		For Each id In Split(Request.Form("idList"),",")
			If (Cint(id) = Cint(Request.Form("newID"))) Then
				dbConnection.Execute("INSERT INTO tblHearAbout ( effectsSeminar, aboutText ) VALUES ( '" & Request.Form("effectsSeminar_" & id) & "', '" & Request.Form("aboutText_" & id) & "' )")
			Else
				dbConnection.Execute("UPDATE tblHearAbout SET effectsSeminar = '" & Request.Form("effectsSeminar_" & id) & "', aboutText = '" & Request.Form("aboutText_" & id) & "' WHERE aboutID = " & id & "")
			End If
		Next
	Else
		'Create the heading to display data.
		Response.Write(tabTo(2) & "<TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0 WIDTH='100%'>")
		Response.Write(tabTo(3) & "<TR CLASS=header>")
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=1% ALIGN=center NOWRAP>&nbsp;#&nbsp;</TD>")
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=50% NOWRAP>Effective Seminars</TD>")
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=40% NOWRAP>Option Text</TD>")
		Response.Write(tabTo(4) & "<TD CLASS=header WIDTH=10% NOWRAP>Remove</TD>")
	   	Response.Write(tabTo(3) & "</TR>")

		Set dbEventData = dbConnection.Execute("SELECT * FROM tblHearAbout ORDER BY effectsSeminar, aboutText")

		'Initialize the Position Counter
		Pos = 0

		Response.Write("<FORM METHOD=post>")

		While Not dbEventData.EOF
			Pos = Pos + 1
			If (Pos = 1) Then
				idList = dbEventData("aboutID")
			Else
				idList = idList & "," & dbEventData("aboutID")
			End If
			If bgclass = "odd" Then bgclass = "even" Else bgclass = "odd" End If
			Response.Write(tabTo(3) & "<TR>")
			Response.Write(tabTo(4) & "<TD ALIGN=center NOWRAP CLASS=" & bgclass & ">" & Pos & "</TD>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">")
			printAboutSeminars dbEventData("effectsSeminar"), dbEventData("aboutID")
			Response.Write(tabTo(4) & "</TD>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=aboutText_" & dbEventData("aboutID") & " VALUE='" & dbEventData("aboutText") & "' SIZE=30></TD>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><INPUT TYPE=""submit"" NAME=remove_" & dbEventData("aboutID") & " VALUE='Remove'></TD>")
			Response.Write(tabTo(3) & "</TR>" & vbNewline)
			dbEventData.MoveNext
		Wend

		Set dbEventData = Nothing

		If (method = "add") Then
			Pos = Pos + 1
			If bgclass = "odd" Then bgclass = "even" Else bgclass = "odd" End If

			Set dbMaxRecord = dbConnection.Execute("SELECT max(aboutID) as aboutID FROM tblHearAbout")
			nextID = dbMaxRecord("aboutID")+1
			Set dbMaxRecord = Nothing

			idList = idList & "," & nextID

			Response.Write(tabTo(3) & "<TR>")
			Response.Write(tabTo(4) & "<TD ALIGN=center NOWRAP CLASS=" & bgclass & ">" & Pos & "</TD>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">")
			printAboutSeminars Request.Form("effectsSeminar_new"), nextID
			Response.Write(tabTo(4) & "</TD>")
			Response.Write(tabTo(4) & "<input type=""hidden"" name=""newID"" value=""" & nextID & """>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & "><INPUT NAME=aboutText_" & nextID & " VALUE='' SIZE=30></TD>")
			Response.Write(tabTo(4) & "<TD NOWRAP CLASS=" & bgclass & ">&nbsp;</TD>")
			Response.Write(tabTo(3) & "</TR>" & vbNewline)
		End If

		If (Pos = 0) Then
			Response.Write("<TR><TD COLSPAN=6 ALIGN=center>No hear about options available.</TD></TR>")
		Else
			Response.Write("<TR>")
			Response.Write("<TD COLSPAN=6 ALIGN=right>")
			Response.Write("<INPUT TYPE=hidden NAME=idList VALUE=" & idList & ">")
			Response.Write("<INPUT TYPE=submit NAME='cancelChanges' VALUE='Cancel'>")
			Response.Write("<INPUT TYPE=reset VALUE='Reset'>")
			Response.Write("<INPUT TYPE=submit NAME='saveChanges' VALUE='Update/Save'>")
			Response.Write(tabTo(3) & "</TR>")
		End If

		If (method <> "add") Then
			Response.Write("<tr><td colspan=6>")
			Response.Write("<center><b>Add a new hear about option:</b><br>")

			printAboutSeminars "", "new"

			Response.Write("<INPUT TYPE=submit NAME='addOption' VALUE='Add New Option'>")
			Response.Write("</center></td></tr>")
		End If

		Response.Write("</FORM>")
		Response.Write(tabTo(2) & "</TABLE>")
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printSeminarNameList(theSeminar)
	Pos = 0

	Set dbSeminarNames = dbConnection.Execute("SELECT strSeminarName FROM tblSeminars GROUP BY strSeminarName ORDER BY strSeminarName")
	Response.Write(tabTo(3) & "<SELECT ID=""seminarSelector"" NAME=""seminarSelector"" SIZE=""1"">")

	While Not dbSeminarNames.EOF
		Pos = Pos + 1
		If (theSeminar = dbSeminarNames("strSeminarName")) Then
			thisIsSelected = "SELECTED"
		Else
			thisIsSelected = ""
		End If
		Response.Write(tabTo(3) & "<OPTION " & thisIsSelected & " VALUE='" & dbSeminarNames("strSeminarName") & "'>" & dbSeminarNames("strSeminarName") & "</OPTION>")
		dbSeminarNames.MoveNext
	Wend

	Response.Write(tabTo(3) & "</SELECT>")

	If (Pos = 0) Then
		Response.Write(errNoSeminarsAvailable)
	End If
End Function

Function printSeminarNameList2(theSeminar)
	Pos = 0

	Set dbSeminarNames = dbConnection.Execute("SELECT strSeminarName FROM tblSeminars GROUP BY strSeminarName ORDER BY strSeminarName")
	Response.Write(tabTo(3) & "<SELECT ID=""strSeminarName"" NAME=""strSeminarName"" SIZE=""1"">")

	While Not dbSeminarNames.EOF
		Pos = Pos + 1
		If (theSeminar = dbSeminarNames("strSeminarName")) Then
			thisIsSelected = "SELECTED"
		Else
			thisIsSelected = ""
		End If
		Response.Write(tabTo(3) & "<OPTION " & thisIsSelected & " VALUE='" & dbSeminarNames("strSeminarName") & "'>" & dbSeminarNames("strSeminarName") & "</OPTION>")
		dbSeminarNames.MoveNext
	Wend

	Response.Write(tabTo(3) & "</SELECT>")

	If (Pos = 0) Then
		Response.Write(errNoSeminarsAvailable)
	End If
End Function

Function getSeminarCost(semName, discount)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set dbNormalCost = dbConnection.Execute("SELECT normalPrice FROM tblPricing WHERE effectsSeminar = '" & semName & "'")
	Set dbDiscountCost = dbConnection.Execute("SELECT discountPrice FROM tblPricing WHERE effectsSeminar = '" & semName & "' AND discountCode = '" & Ucase(discount) & "'")

	If (dbNormalCost.EOF) Then
		getSeminarCost = "n/a"
	Else
		If (dbDiscountCost.EOF) Then
			getSeminarCost = dbNormalCost("normalPrice")
    	Else
    	   	getSeminarCost = dbDiscountCost("discountPrice")
    	End If
    End If

	Set dbNormalCost = Nothing
	Set dbDiscountCost = Nothing

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printSeminarOptions(selectedSeminar)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set dbEventData = dbConnection.Execute("SELECT * FROM tblSeminars ORDER BY strSeminarName, dateSeminarDate")
	Pos = 0

	While Not dbEventData.EOF
		Pos = Pos + 1
		If Not (selectedSeminar = "select") Then
			If (Cint(dbEventData("numSeminarID")) = Cint(selectedSeminar)) Then
				selected = "SELECTED"
			Else
				selected = ""
			End If
		End If
		Response.Write(tabTo(3) & "<OPTION VALUE='" & dbEventData("numSeminarID") & "' " & selected & ">")
		Response.Write("[" & dbEventData("strSeminarName") & "] ")
		Response.Write(dbEventData("strSeminarCity"))
		Response.Write(" (" & dbEventData("dateSeminarDate") & ")")
		Response.Write("</OPTION>")
		dbEventData.MoveNext
	Wend

	If (Pos = 0) Then
		Response.Write(errNoSeminarsActive)
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Dim monthName
redim monthName(13)
monthName = array("January","February","March","April","May","June","July","August","September","October","November","December")
Function printSeminarOptions2(selectedSeminar, showOnlyTour, formInputName)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set dbEventData = dbConnection.Execute("SELECT * FROM tblSeminars WHERE isVisible = True AND strSeminarName = '" & showOnlyTour & "' AND (linkedToSeminar = -1 OR linkedToSeminar = numSeminarID) ORDER BY strSeminarName, dateSeminarDate")
	Pos = 0

	While Not dbEventData.EOF
		Pos = Pos + 1
		If Not (selectedSeminar = "") Then
			If (Cint(dbEventData("numSeminarID")) = Cint(selectedSeminar)) Then
				selected = "CHECKED"
			Else
				selected = ""
			End If
		End If
		If (bgColor = "#c0c0e0") Then bgColor = "#e0c0c0" Else bgColor = "#c0c0e0" End If
		Response.Write(tabTo(3) & "<tr bgcolor=""" & bgColor & """>")
		Response.Write("<td valign=""top"" align=""center""><input type=""radio"" name=""" & formInputName & """ VALUE='" & dbEventData("numSeminarID") & "' " & selected & "></td>")
'		Response.Write("<td valign=""top"">" & dbEventData("strSeminarName") & "</td>")
		Response.Write("<td valign=""top"">" & dbEventData("strSeminarCity") & "</td>")

		If (Fix(dbEventData("linkedToSeminar")) = -1) Then
			thisDate = monthName(datePart("m",dbEventData("dateSeminarDate")) - 1) & " " & datePart("d",dbEventData("dateSeminarDate")) & ", " & datePart("yyyy",dbEventData("dateSeminarDate"))
		Else
			Set datesForEvent = dbConnection.Execute("SELECT * FROM tblSeminars WHERE linkedToSeminar = " & dbEventData("linkedToSeminar") & " ORDER BY dateSeminarDate")

			thisDate = ""
			Do Until datesForEvent.EOF
				thisDate = thisDate & monthName(datePart("m",datesForEvent("dateSeminarDate")) - 1) & " " & datePart("d",datesForEvent("dateSeminarDate")) & ", " & datePart("yyyy",datesForEvent("dateSeminarDate")) & "<br />"
				datesForEvent.MoveNext
			Loop
		End If

		Response.Write("<td valign=""top"">" & thisDate & "</td>")
		Response.Write("</tr>")
		dbEventData.MoveNext
	Wend

	If (Pos = 0) Then
		Response.Write(errNoSeminarsActive)
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printSeminarNames(theType)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Pos = 0

	If (theType = "select") Then
		Set dbEventData = dbConnection.Execute("SELECT DISTINCT strSeminarName FROM tblSeminars WHERE isDeleted = False ORDER BY strSeminarName")
		Response.Write(tabTo(3) & "<SELECT ID=""seminarSelector"" NAME=""seminarSelector"" SIZE=""1"">")
	ElseIf (theType = "link") Then
		Set dbEventData = dbConnection.Execute("SELECT DISTINCT strSeminarName, iOrd, normalPrice, sOneLiner, sVendor FROM tblSeminars, tblPricing WHERE tblSeminars.strSeminarName = tblPricing.effectsSeminar AND tblPricing.normalPrice IS NOT NULL AND tblSeminars.isVisible = True AND tblSeminars.isDeleted = False ORDER BY iOrd, normalPrice, strSeminarName")
'		Set dbEventData = dbConnection.Execute("SELECT strSeminarName FROM tblSeminars WHERE isVisible = true GROUP BY strSeminarName ORDER BY strSeminarName")
		Response.Write("<table cellSpacing=""1"" cellPadding=""2"" width=""544"" border=""0"">")
		Response.Write("<tbody><tr>")
		Response.Write("<td background=""/images/banner_back.gif"" width=""60%""><!--mstheme--><font face=""Arial, Arial, Helvetica""><b><font color=""#eeeeee"" size=""2"">Event Name</font></b><!--mstheme--></font></td>")
		'Response.Write("<td background=""/images/banner_back.gif"" width=""20%""><!--mstheme--><font face=""Arial, Arial, Helvetica""><b><font color=""#eeeeee"" size=""2"">Begins</font></b><!--mstheme--></font></td>")
		Response.Write("<td background=""/images/banner_back.gif"" width=""30%""><!--mstheme--><font face=""Arial, Arial, Helvetica""><b><font color=""#eeeeee"" size=""2"">Price</font></b><!--mstheme--></font></td>")
		Response.Write("<td background=""/images/banner_back.gif"" width=""10%""><!--mstheme--><font face=""Arial, Arial, Helvetica""><font color=""#ffffff"" size=""2""><b>Info</b></font><!--mstheme--></font></td>")
	    Response.Write("</tr>")
	End If

	If Not (dbEventData.EOF) Then
		If (theType = "select") Then
			Response.Write(tabTo(4) & "<OPTION VALUE=""All"">All</OPTION>")
		End If

		While Not dbEventData.EOF
			Pos = Pos + 1
			If (theType = "link") Then
				'Set dbDate = dbConnection.Execute("SELECT TOP 1 dateSeminarDate FROM tblSeminars WHERE strSeminarName = '" & dbEventData("strSeminarName") & "' ORDER BY dateSeminarDate")
				'If (dbDate.EOF) Then
				'	thisDate = "N/A"
				'Else
				'	thisDate = dbDate("dateSeminarDate")
				'End If
				'Set dbDate = Nothing

				thisEvent = dbEventData("sVendor")

				Set dbInfo = dbConnection.Execute("SELECT infoURL FROM tblInfo WHERE effectsSeminar = '" & dbEventData("strSeminarName") & "'")
				If (dbInfo.EOF) Then
					thisInfoURL = "N/A"
				Else
					thisInfoURL = "<a href=""" & dbInfo("infoURL") & """>Info</a>"
				End If
				Set dbInfo = Nothing

				Set dbPrice = dbConnection.Execute("SELECT normalPrice FROM tblPricing WHERE effectsSeminar = '" & dbEventData("strSeminarName") & "'")
				If (dbPrice.EOF) Then
					thisPrice = "N/A"
				Else
					If (Fix(dbPrice("normalPrice")) = 0) Then
						thisPrice = "Free!"
					Else
						thisPrice = "$" & FormatNumber(dbPrice("normalPrice"),2)
					End If
				End If
				Set dbPrice = Nothing
	
				If (InStr(LCase(thisEvent), LCase(Request.QueryString)) > 0) Or (Len(Request.QueryString) = 0) Or (Request.QueryString = "all") Then
					Response.Write(tabTo(3) & "<tr>")
					Response.Write(tabTo(4) & "<td><a href=""register.asp?tour=" & Server.URLEncode(dbEventData("strSeminarName")) & """>" & dbEventData("strSeminarName") & "</a></td>")
					'Response.Write(tabTo(4) & "<td>" & thisDate & "</td>")
					Response.Write(tabTo(4) & "<td>" & thisPrice  & "</td>")
					Response.Write(tabTo(4) & "<td>" & thisInfoURL & "</td>")
					Response.Write(tabTo(3) & "</tr>")
					If (Len(dbEventData("sOneLiner")) > 0) Then
						Response.Write(tabTo(3) & "<tr>")
						Response.Write(tabTo(4) & "<td colspan=""3""><small><small>" & dbEventData("sOneLiner") & "</small></small><br /></a></td>")
						Response.Write(tabTo(3) & "</tr>")
					End If		
				End If
			ElseIf (theType = "select") Then
				If (Request.QueryString("searchFor") = dbEventData("strSeminarName")) Then
					thisIsSelected = "SELECTED"
				Else
					thisIsSelected = ""
				End If
				Response.Write(tabTo(3) & "<OPTION " & thisIsSelected & " VALUE='" & dbEventData("strSeminarName") & "'>" & dbEventData("strSeminarName") & "</OPTION>")
			End If
			dbEventData.MoveNext
		Wend
	End If

	If (Pos = 0) Then
		If (theType = "select") Then
			Response.Write("<option value=""-1"">No seminars available.</option>")
		Else
			Response.Write(errNoSeminarsAvailable)
		End If
	End If

	If (theType = "select") Then
		Response.Write(tabTo(3) & "</SELECT>")
	ElseIf (theType = "link") Then
		Response.Write(tabTo(3) & "</table>")
		If (Len(Request.QueryString) > 0) And (InStr(Request.QueryString,"all") = 0) Then
			Response.Write("<br /><small><a href=""/registration/?all"">Click here to see all VASST tours.</a></small>")
		End If
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function getSeminarID(byName, byDate, byCity)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	selectedQueries = 0
	queryStr = ""

'	Response.Write("ByName:" & byName & "<BR>")
'	Response.Write("ByDate:" & byDate & "<BR>")
'	Response.Write("ByCity:" & byCity & "<BR>")

	If Not (byName = "") Then
		queryName = "strSeminarName LIKE '" & byName &"'"
		selectedQueries = selectedQueries + 1
		queryStr = queryName
	End If

	If Not (byDate = "") Then
		queryDate = "dateSeminarDate LIKE '" & byDate & "'"
		If (Cint(selectedQueries) > 0) Then
			queryStr = queryStr & " AND " & queryDate
		Else
			queryStr = queryDate
		End If
		selectedQueries = selectedQueries + 1
	End If

	If Not (byCity = "") Then
		queryCity = "strSeminarCity LIKE '" & byCity & "'"
		If (Cint(selectedQueries) > 0) Then
			queryStr = queryStr & " AND " & queryCity
		Else
			queryStr = queryCity
		End If
		selectedQueries = selectedQueries + 1
	End If

	If (Cint(selectedQueries) = 0) Then
		getSeminarID = -1
	Else
'		Response.Write(queryStr)
		queryStr = "SELECT numSeminarID FROM tblSeminars WHERE " & queryStr

		Set dbEventData = dbConnection.Execute(queryStr)
		If (dbEventData.EOF) Then
			retValue = -1
		Else
			retValue = dbEventData("numSeminarID")
		End If
	End If

	queryStr = ""

	If Not (wasDbOpen) Then
		closeDatabase
	End If
	getSeminarID = retValue
End Function

Function getNextSeminarID
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set dbEventData=dbConnection.Execute("SELECT max(numSeminarID) as lastEvent FROM tblSeminars")

	getNextSeminarID = dbEventData("lastEvent") + 1
	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function addSeminar(strName, dateDate, strCity, isDateArray)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	If (isDateArray) Then
		dateArray = Split(dateDate,",")
		totalDates = Ubound(dateArray)

		dbConnection.Execute("INSERT INTO tblSeminars ( strSeminarName, dateSeminarDate, strSeminarCity, isVisible ) VALUES ( '" & strName & "', #" & dateArray(0) & "#, '" & strCity & "', True )")

		closeDatabase
		openDatabase

		Set dbFirstSeminarID = dbConnection.Execute("SELECT max(numSeminarID) as maxID FROM tblSeminars WHERE strSeminarName = '" & strName & "' AND dateSeminarDate = #" & dateArray(0) & "# AND strSeminarCity = '" & strCity & "' AND isVisible = True")
		firstSeminarID = dbFirstSeminarID("maxID")
		Set dbFirstSeminarID = Nothing

		dbConnection.Execute("UPDATE tblSeminars SET linkedToSeminar = " & firstSeminarID & " WHERE numSeminarID = " & firstSeminarID & "")
		For X = 1 To totalDates
			dbConnection.Execute("INSERT INTO tblSeminars ( strSeminarName, dateSeminarDate, strSeminarCity, isVisible, linkedToSeminar ) VALUES ( '" & strName & "', #" & dateArray(X) & "#, '" & strCity & "', True," & firstSeminarID & " )")
		Next
	Else
		dbConnection.Execute("INSERT INTO tblSeminars ( strSeminarName, dateSeminarDate, strSeminarCity, isVisible ) VALUES ( '" & strName & "', #" & dateDate & "#, '" & strCity & "', True )")
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function hideSeminar(eventID, status)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("UPDATE tblSeminars SET isVisible = " & status & " WHERE linkedToSeminar = " & eventID & " OR numSeminarID = " & eventID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function updateSeminar(eventID, strName, dateDate, strCity, isVisible, isDateArray)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set dbCount = dbConnection.Execute("SELECT count(numSeminarID) as totalCurrentDates FROM tblSeminars WHERE linkedToSeminar = " & eventID & " OR numSeminarID = " & eventID & "")
	totalCurrentDates = dbCount("totalCurrentDates")
	Set dbCount = Nothing

	If (isDateArray) Then
		dateArray = Split(dateDate,",")
		neededDates = Ubound(dateArray) + 1
	Else
		neededDates = 1
	End If

	datesDiff = totalCurrentDates - neededDates

	If (datesDiff < 0) Then
		Set dbSeminar = dbConnection.Execute("SELECT * FROM tblSeminars WHERE linkedToSeminar = numSeminarID AND numSeminarID = " & eventID & "")
		Response.Write(datesDiff)
		For X = 1 To (datesDiff * -1)
			dbConnection.Execute("INSERT INTO tblSeminars ( strSeminarName, strSeminarCity, strDiscountCode, isVisible, linkedToSeminar ) VALUES ( '" & dbSeminar("strSeminarName") & "', '" & dbSeminar("strSeminarCity") & "', '" & dbSeminar("strDiscountCode") & "', " & dbSeminar("isVisible") & ", " & eventID & " ) ")
		Next
		Set dbSeminar = Nothing
	ElseIf (datesDiff > 0) Then
		For X = 1 To (datesDiff)
			Set dbLastSeminar = dbConnection.Execute("SELECT max(numSeminarID) as lastSeminarID FROM tblSeminars WHERE linkedToSeminar = " & eventID & "")
			dbConnection.Execute("DELETE * FROM tblSeminars WHERE numSeminarID = " & dbLastSeminar("lastSeminarID") & "")
			Set dbLastSeminar = Nothing
		Next
	End If

	closeDatabase
	openDatabase

	If (isDateArray) Then
		Set dbSeminars = dbConnection.Execute("SELECT numSeminarID FROM tblSeminars WHERE linkedToSeminar = " & eventID & "")
		If (dbSeminars.EOF) Then
			Response.Write("Something failed on the count of the differences.<br />")
			Response.End
		Else
			dateArrayPos = 0
			Do Until dbSeminars.EOF
				dbConnection.Execute("UPDATE tblSeminars SET strSeminarName = '" & strName & "', dateSeminarDate = #" & dateArray(dateArrayPos) & "#, strSeminarCity = '" & strCity & "', isVisible = " & isVisible & " WHERE numSeminarID = " & dbSeminars("numSeminarID") & "")
				dateArrayPos = dateArrayPos + 1
				dbSeminars.MoveNext
			Loop
		End If
	Else
		dbConnection.Execute("UPDATE tblSeminars SET strSeminarName = '" & strName & "', dateSeminarDate = '" & dateDate & "', strSeminarCity = '" & strCity & "', isVisible = " & isVisible & " WHERE numSeminarID = " & eventID)
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function removeSeminar(eventID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

'	dbConnection.Execute("DELETE * FROM tblSeminars WHERE linkedToSeminar = " & eventID & " OR numSeminarID = " & eventID & "")
	dbConnection.Execute("UPDATE tblSeminars SET isDeleted = true WHERE linkedToSeminar = " & eventID & " OR numSeminarID = " & eventID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function Reverse(trueFalse)
	If (trueFalse) Then
		Reverse = False
	Else
		Reverse = True
	End If
End Function

Function registerCustomer(firstName, lastName, companyName, companyTitleName, address1Name, address2Name, cityName, stateName, zipName, phoneName, emailName, paymentCost, paymentType, selectSeminar, comments1, hearAboutText, wasDiscounted, countryName)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set dbEventData = dbConnection.Execute("SELECT * FROM tblSeminars WHERE numSeminarID = " & selectSeminar & "")

	seminarName = dbEventData("strSeminarName")
	seminarDate = dbEventData("dateSeminarDate")
	seminarCity = dbEventData("strSeminarCity")
	linkedToSeminar = dbEventData("linkedToSeminar")

	'Response.Write("Linked To Seminar: " & linkedToSeminar & "<br />")

	If (Fix(linkedToSeminar) = -1) Then
		linkedSeminar = -1
	Else
		linkedSeminar = linkedToSeminar
	End If

	signupDate = date & " " & time

	Call addCustomer(-1, signupDate, firstName, lastName, companyName, companyTitleName, address1Name, address2Name, cityName, stateName, zipName, phoneName, emailName, seminarName, seminarDate, seminarCity, linkedSeminar, paymentCost, paymentType, wasDiscounted, False, False, "New Signup", "", "", hearAboutText, comments1, "", "", countryName)

	closeDatabase
	openDatabase

	strSQL = "SELECT max(numCustID) as lastCustomerID " & _
		"FROM tblCustomers " & _
		"WHERE " & _
		"strFirstName = '" & firstName & "' AND " & _
		"strLastName = '" & lastName & "' AND " & _
		"strEmail = '" & emailName & "' AND " & _
		"strSeminarName = '" & seminarName & "' AND " & _
		"dateSeminarDate = #" & seminarDate & "# AND " & _
		"strSeminarCity = '" & seminarCity & "' " & _
		""
	'Response.Write("<b>EXECUTING:</b> " & strSQL & "<br />")

	Set dbCustomer = dbConnection.Execute(strSQL)
	If (dbCustomer.EOF) Then
		Response.Write("There was a problem finding your registration after it was entered into the database.  Please contact the administrator.<br />")
		Response.End
	Else
		registerCustomer = dbCustomer("lastCustomerID")
	End If

	Set dbEventData = Nothing
	Set dbCustomer = Nothing

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Dim minRecord, maxRecord, countRecord, editID, editCustID, editSignupDate, editFirstName, editLastName, editCompanyName, editTitleName, editAddress1Name, editAddress2Name, editCityName, editStateName, editZipName, editPhoneName, editEmailName, editSeminarName, editSeminarDate, editSeminarCity, editSeminarCost, editPaymentType, editDiscount, editPaid, editDeleted, editStatus, editRetail1, editRetail2, editRetail3, editComments1, editComments2, editComments3

Function dontShowListOrEdit
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	If (Request.QueryString("searchFor") = "") Or (Request.QueryString("searchFor") = "All") Then
		'Show everything in the tblCustomers
		Set dbCount = dbConnection.Execute("SELECT min(numCustID) as minRecord, max(numCustID) as maxRecord, count(numCustID) as countRecord FROM tblCustomers")
	Else
		Set dbCount = dbConnection.Execute("SELECT min(numCustID) as minRecord, max(numCustID) as maxRecord, count(numCustID) as countRecord FROM tblCustomers WHERE strSeminarName = '" & Request.QueryString("searchFor") & "'")
	End If

	minRecord = dbCount("minRecord")
	maxRecord = dbCount("maxRecord")
	countRecord = dbCount("countRecord")

	If Not isNumeric(minRecord) Then
		minRecord = 0
	End If
	If Not isNumeric(maxRecord) Then
		maxRecord = 0
	End If
	If Not isNumeric(countRecord) Then
		countRecord = 0
	End If

	If Not ((minRecord > 0) And (maxRecord > 0) And (countRecord > 0)) Then
		dontShowListOrEdit = true
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printCustomers(spreadsheetMode)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Response.Write("</form><form method=""post"" action=""?searchFor=" & Server.URLEncode(Request.QueryString("searchFor")) & """>")

	If (Request.QueryString("searchFor") = "") Or (Request.QueryString("searchFor") = "All") Then
		'Show everything in the tblCustomers
		Set dbCount = dbConnection.Execute("SELECT min(numCustID) as minRecord, max(numCustID) as maxRecord, count(numCustID) as countRecord FROM tblCustomers")
	Else
		Set dbCount = dbConnection.Execute("SELECT min(numCustID) as minRecord, max(numCustID) as maxRecord, count(numCustID) as countRecord FROM tblCustomers WHERE strSeminarName = '" & Request.QueryString("searchFor") & "'")
	End If

	minRecord = dbCount("minRecord")
	maxRecord = dbCount("maxRecord")
	countRecord = dbCount("countRecord")

	If Not isNumeric(minRecord) Then
		minRecord = 0
	End If
	If Not isNumeric(maxRecord) Then
		maxRecord = 0
	End If
	If Not isNumeric(countRecord) Then
		countRecord = 0
	End If

	If ((minRecord > 0) And (maxRecord > 0) And (countRecord > 0)) Then
		If ((spreadsheetMode = "view") Or (spreadsheetMode = "edit") Or (spreadsheetMode = "update")) Then
			set objRS=server.createObject("ADODB.Recordset")
			If (Request.QueryString("searchFor") = "") Or (Request.QueryString("searchFor") = "All") Then
				'Show everything in the tblCustomers
				strSQL = "SELECT * FROM tblCustomers"
			Else
				strSQL = "SELECT * FROM tblCustomers WHERE strSeminarName = '" & Request.QueryString("searchFor") & "'"
			End If
			objRS.open strSQL, dbConnection, 2, 3

			dbHeadings = array(	  "count",	"numCustID",	"dateSignup",	"strFirstName",	"strLastName",	"strCompanyName",	"strTitle",	"strAddress1",	"strAddress2",	"strCity",	"strState",	"strZip",	"numPhone",	"strEmail",	"strSeminarName",	"dateSeminarDate",	"strSeminarCity",	"linkedSeminar",	"currencyCost",	"optPayment",	"isDiscount",	"isPaid",	"isDeleted",	"optStatus",	"strRetail1",	"strRetail2",	"strRetail3",	"strComments1",	"strComments2",	"strComments3", "strCountry")
'			recordSize = array(		4,			4,					20,				15,				15,				20,					10,		20,			20,			15,		10,		10,		12,		20,		20,			20,			20,			8,			10,		6,		6,		6,		15,		20,		20,		20,		20,			20,			20)
			recordHeadings = array(	"#",	"ID",	"Date of Signup",	"First Name",	"Last Name",	"Company Name",	"Title",	"Address1",	"Address2",	"City",	"State",	"Zip",	"Phone Number",	"Email Address",	"Seminar Name",	"Seminar Date",	"Seminar City",		"Linked Seminar", 	"Seminar Cost",	"Payment Type",	"Discount?",	"Paid?",	"Deleted?",	"Status",	"Retail1",	"Retail2",	"Retail3",	"Comments1",	"Comments2",	"Comments3", "Country")
			recordSize = array(		4,		4,		20,					15,				15,				20,				10,			20,			20,			15,		10,			10,		12,				20,					20,				20,				20,					8,					8,				10,				6,				6,			6,			15,			20,			20,			20,			20,				20,				20, 20)

			If (spreadsheetMode = "update") And Not (countRecord = 0) Then
				Dim records
				numberOfRecords = 0
				numberOfFields = 0

				For Each objForm In Request.Form
					If (InStr(objForm,"field_") > 0) Then
						splitter = Split(objForm,"_")
						id = Cint(splitter(1))
						field = Cint(splitter(2))
						If (id > Cint(numberOfRecords)) Then
							numberOfRecords = id
						End If
						If (field > Cint(numberOfFields)) Then
							numberOfFields = field
						End If
					End If
				Next
				ReDim records(numberOfRecords, numberOfFields)

				For Each objForm In Request.Form
					tmp = Split(objForm,"=")
					If (InStr(objForm,"field_") > 0) Then
'						Response.Write(objForm & "=" & Request.Form(objForm) & "<BR>")
						splitter = Split(objForm,"_")
						id = Cint(splitter(1))
						field = Cint(splitter(2))
'						Response.Write("id:" & id & " field:" & field & "(" & recordHeadings(field+1) & ") =" & Request.Form(objForm) & "<BR>")
						records(id,field) = Request.Form(objForm)
					End If
				Next

				objRS.Move 0
				For x = 1 to numberOfRecords
					For y = 0 to objRS.Fields.Count-1
'						Response.Write("X:" & x & " Y:" & y & "(" & recordHeadings(y+1) & ") --- " & objRS.Fields(y) & "=" & records(x,y))
'						Response.Write("<BR>")
						If Not (y = 0) Then
							objRS.Fields(y) = records(x,y)
						End If
					Next
					objRS.MoveNext
				Next
			Else
				Pos = 0
				Response.Write(tabTo(0) & "<STYLE>")
				Response.Write(tabTo(1) & "TABLE.spreadsheet { Border-Left: 1px solid black; }")
				Response.Write(tabTo(1) & "TR.heading { Background: black; Color: white; Font-Weight: bold; }")
				Response.Write(tabTo(1) & "TD { Font-Size: 9pt; }")
				Response.Write(tabTo(1) & "INPUT.field { Font-Size: 9pt; }")
				Response.Write(tabTo(0) & "</STYLE>")
				Response.Write(tabTo(2) & "<TABLE BORDER=1 CELLPADDING=2 CELLSPACING=0 CLASS=spreadsheet STYLE=""Border-Collapse: collapse; Border-Color: black;"">")
				Response.Write(tabTo(3) & "<TR>")
				Response.Write(tabTo(4) & "<TD COLSPAN=" & objRS.Fields.Count+1 & " CLASS=spreadsheet>")
				Response.Write(tabTo(5) & "<p align=left><font size=+2><b>Customer Database</b></font></p>")
				Response.Write(tabTo(4) & "</TD>")
				Response.Write(tabTo(3) & "</TR>")
				While Not objRS.EOF
					If ((Pos Mod 10) = 0) Then
						Response.Write("<TR>" & vbNewline)
						Response.Write("<TD BGCOLOR=aaaaaa nowrap><font color=333333><b>&nbsp#&nbsp;</b></font></TD>")
						For Each X In objRS.fields
							Response.Write("<TD BGCOLOR=aaaaaa nowrap><font color=333333><b>" & humanize(X.name) & "</b></font></TD>" & vbNewline)
						Next
						Response.Write("</TR>" & vbNewline)
					End If
					Pos = Pos + 1
					Response.Write(tabTo(3) & "<TR>")
					Response.Write(tabTo(4) & "<TD NOWRAP CLASS=spreadsheet>" & Pos & "</TD>")
					for a = 0 to objRS.Fields.Count-1
						If (spreadsheetMode = "view") or (a = 0) Then
							Response.Write(tabTo(4) & "<TD NOWRAP>" & objRS.Fields(a) & "&nbsp;</TD>")
						ElseIf (spreadsheetMode = "edit") Then
							Response.Write(tabTo(4) & "<TD NOWRAP><INPUT CLASS=field NAME=field_" & Pos & "_" & a & " VALUE='" & objRS.Fields(a) & "' SIZE=" & recordSize(a+1) & "></TD>")
						End If
					next
					Response.Write(tabTo(3) & "</TR>")
					objRS.MoveNext
				Wend
				Response.Write(tabTo(3) & "<TR>")
				Response.Write(tabTo(4) & "<TD COLSPAN=" & objRS.Fields.Count+1 & " CLASS=spreadsheet>")
				Response.Write(tabTo(5) & "<p align=left><font size=2>Total # of Records = " & countRecord & "</font></p>")
				Response.Write(tabTo(4) & "</TD>")
				Response.Write(tabTo(3) & "</TR>")
				Response.Write(tabTo(3) & "<TR>")
				Response.Write(tabTo(4) & "<TD COLSPAN=" & objRS.Fields.Count+1 & " CLASS=spreadsheet>")
				Response.Write(tabTo(5) & "<TABLE BORDER=0 WIDTH='100%' ALIGN=center>")
				Response.Write(tabTo(6) & "<TR>")
				For looptimes = 1 to 5
					Response.Write(tabTo(7) & "<TD ALIGN=center NOWRAP>")
					If (spreadsheetMode = "edit") Then
						Response.Write(tabTo(8) & "<INPUT TYPE=submit NAME='saveButton' VALUE='Save' STYLE='Width:49%'><INPUT TYPE=reset VALUE='Reset' STYLE='Width:49%'>")
					ElseIf (spreadsheetMode = "view") Then
						If (getUserInfo(getUserID,"accesslevel") = "Admin") Then
							Response.Write(tabTo(8) & "<INPUT TYPE=submit NAME='editButton' VALUE='Edit' STYLE='Width:98%'>")
						End If
					End If
					Response.Write(tabTo(7) & "</TD>")
				Next
				Response.Write(tabTo(6) & "</TR>")
				Response.Write(tabTo(5) & "</TABLE>")
				Response.Write(tabTo(4) & "</TD>")
				Response.Write(tabTo(3) & "</TR>")
				Response.Write(tabTo(2) & "</TABLE>")
				objRS.close
			End If
		Else
			editID = ""

			If (Request.QueryString("viewCustomer") = "") Then
				If (Request.Cookies("customerPass")("editID") = "") Then
					If (Request.Form("editID") = "") Then
						editID = minRecord
					ElseIf (Request.Form("first") = "<<") Then
						editID = minRecord
					ElseIf (Request.Form("prev") = "<") And (Fix(Request.Form("editID")) > minRecord) Then
						editID = Request.Form("editID") - 1
					ElseIf (Request.Form("next") = ">") Then
						editID = Request.Form("editID") + 1
					ElseIf (Request.Form("last") = ">>") Then
						editID = maxRecord
					Else
						editID = Request.Form("editID")
					End IF
				Else
					editID = Request.Cookies("customerPass")("editID")
					Response.Cookies("customerPass")("editID") = ""
				End If
			Else
				editID = Request.QueryString("viewCustomer")
			End If

			If (Request.QueryString("searchFor") = "") Or (Request.QueryString("searchFor") = "All") Then
				'Show everything in the tblCustomers
				Set dbCustomer = dbConnection.Execute("SELECT * FROM tblCustomers WHERE numCustID = " & editID & "")
			Else
				Set dbCustomer = dbConnection.Execute("SELECT * FROM tblCustomers WHERE numCustID = " & editID & " AND strSeminarName = '" & Request.QueryString("searchFor") & "'")
			End If
			While (dbCustomer.BOF) And Not (noMoreRecords)
				If (Request.Form("prev") = "<") Then
					editID = editID - 1
				Else
					editID = editID + 1
				End If
				If (Cint(editID) > Cint(maxRecord)-1) Then
					editID = maxRecord
				End If
				If (Request.QueryString("searchFor") = "") Or (Request.QueryString("searchFor") = "All") Then
					'Show everything in the tblCustomers
					Set dbCustomer = dbConnection.Execute("SELECT * FROM tblCustomers WHERE numCustID = " & editID & "")
				Else
					Set dbCustomer = dbConnection.Execute("SELECT * FROM tblCustomers WHERE numCustID = " & editID & " AND strSeminarName = '" & Request.QueryString("searchFor") & "'")
				End If
			Wend

			editCustID = dbCustomer("numCustID")
			editSignupDate = dbCustomer("dateSignup")
			editFirstName = dbCustomer("strFirstName")
			editLastName = dbCustomer("strLastName")
			editCompanyName = dbCustomer("strCompanyName")
			editTitleName = dbCustomer("strTitle")
			editAddress1Name = dbCustomer("strAddress1")
			editAddress2Name = dbCustomer("strAddress2")
			editCityName = dbCustomer("strCity")
			editStateName = dbCustomer("strState")
			editZipName = dbCustomer("strZip")
			editPhoneName = dbCustomer("numPhone")
			editEmailName = dbCustomer("strEmail")
			editSeminarName = dbCustomer("strSeminarName")
			editSeminarDate = dbCustomer("dateSeminarDate")
			editSeminarCity = dbCustomer("strSeminarCity")
			editSeminarCost = dbCustomer("currencyCost")
			editPaymentType = dbCustomer("optPayment")
			editDiscount = dbCustomer("isDiscount")
			editPaid = dbCustomer("isPaid")
			editDeleted = dbCustomer("isDeleted")
			editStatus = dbCustomer("optStatus")
			editRetail1 = dbCustomer("strRetail1")
			editRetail2 = dbCustomer("strRetail2")
			editRetail3 = dbCustomer("strRetail3")
			editComments1 = dbCustomer("strComments1")
			editComments2 = dbCustomer("strComments2")
			editComments3 = dbCustomer("strComments3")
		End If
	'Else
	'	dontShowListOrEdit = true
		'Response.Write("No customer's available!")
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printNewSignups
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set loginData = dbConnection.Execute("SELECT loginDate FROM tblAccessLog WHERE loginID < " & getUserInfo(getUserID,"lastlogin") & " AND loginUsername = '" & getUserInfo(getUserID,"username") & "' AND loginResult = 'Login Successful' ORDER BY loginDate DESC")
	If loginData.EOF Then
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;None."
	Else
		Set newSignups = dbConnection.Execute("SELECT * FROM tblCustomers WHERE dateSignup > #" & loginData("loginDate") & "# ORDER BY dateSignup DESC")
		If newSignups.EOF Then
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;None."
		Else
			cnt = 0
			Do Until newSignups.EOF
				cnt = cnt + 1
				Response.Write "<small>&nbsp;&nbsp;&nbsp;&nbsp;Customer ID#: " & newSignups("numCustID") & " - " & newSignups("strLastname") & ", " & newSignups("strFirstname") & " on " & newSignups("dateSignup") & " for " & newSignups("strSeminarName") & "</small><BR>"
				newSignups.MoveNext
			Loop
			Response.Write "<BR>Total # of New Signups:" & cnt
		End If
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printToBePurged
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set toDelete = dbConnection.Execute("SELECT * FROM tblCustomers WHERE isDeleted = true ORDER BY numCustID DESC")
	If toDelete.EOF Then
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;None."
	Else
		cnt = 0
		Do Until toDelete.EOF
			cnt = cnt + 1
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;Customer ID#: " & toDelete("numCustID") & " - " & toDelete("strLastname") & ", " & toDelete("strFirstname") & "<BR>"
			toDelete.MoveNext
		Loop
		Response.Write "<BR>Total # to be removed: " & cnt
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function purgeCustomers
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("DELETE * FROM tblCustomers WHERE isDeleted = true")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function modifyCustomer(updateMethod)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	editCustID = Replace(Request.Form("editCustID"),"'","&#39;")
	editSignupDate = Replace(Request.Form("editSignupDate"),"'","&#39;")
	editDiscount = Replace(Request.Form("editDiscount"),"'","&#39;")
	editPaid = Replace(Request.Form("editPaid"),"'","&#39;")
	editDeleted = Replace(Request.Form("editDeleted"),"'","&#39;")
	editFirstName = Replace(Request.Form("editFirstName"),"'","&#39;")
	editLastName = Replace(Request.Form("editLastName"),"'","&#39;")
	strCompanyName = Replace(Request.Form("editCompanyName"),"'","&#39;")
	editTitleName = Replace(Request.Form("editTitleName"),"'","&#39;")
	editAddress1Name = Replace(Request.Form("editAddress1Name"),"'","&#39;")
	editAddress2Name = Replace(Request.Form("editAddress2Name"),"'","&#39;")
	editCityName = Replace(Request.Form("editCityName"),"'","&#39;")
	editStateName = Replace(Request.Form("editStateName"),"'","&#39;")
	editZipName = Replace(Request.Form("editZipName"),"'","&#39;")
	editCountryName = Replace(Request.Form("editCountryName"),"'","&#39;")
	editPhoneName = Replace(Request.Form("editPhoneName"),"'","&#39;")
	editEmailName = Replace(Request.Form("editEmailName"),"'","&#39;")
	'seminarSelectMethod = Replace(Request.Form("seminarSelectMethod"),"'","&#39;") 'list or manual
	editSeminarCost = Replace(Request.Form("editSeminarCost"),"'","&#39;")
	editPaymentType = Replace(Request.Form("editPaymentType"),"'","&#39;")
	editStatus = Replace(Request.Form("editStatus"),"'","&#39;")
	editRetail1 = Replace(Request.Form("editRetail1"),"'","&#39;")
	editRetail2 = Replace(Request.Form("editRetail2"),"'","&#39;")
	editRetail3 = Replace(Request.Form("editRetail3"),"'","&#39;")
	editComments1 = Replace(Request.Form("editComments1"),"'","&#39;")
	editComments2 = Replace(Request.Form("editComments2"),"'","&#39;")
	editComments3 = Replace(Request.Form("editComments3"),"'","&#39;")

	seminarSelect = Request.Form("seminarSelect")	'list
	Set dbEventData = dbConnection.Execute("SELECT * FROM tblSeminars WHERE numSeminarID = " & seminarSelect & "")

	editSeminarName = dbEventData("strSeminarName")
	editSeminarDate = dbEventData("dateSeminarDate")
	editSeminarCity = dbEventData("strSeminarCity")

	linkedSeminar = dbEventData("linkedToSeminar")

	If (editDiscount = "ON") Then
		editDiscount = "True"
	Else
		editDiscount = "False"
	End If

	If (editPaid = "ON") Then
		editPaid = "True"
	Else
		editPaid = "False"
	End If

	If (editDeleted = "ON") Then
		editDeleted = "True"
	Else
		editDeleted = "False"
	End If

	If (updateMethod = "add") Then
		Call addCustomer(editCustID, editSignupDate, editFirstName, editLastName, editCompanyName, editTitleName, editAddress1Name, editAddress2Name, editCityName, editStateName, editZipName, editPhoneName, editEmailName, editSeminarName, editSeminarDate, editSeminarCity, linkedSeminar, editSeminarCost, editPaymentType, editDiscount, editPaid, editDeleted, editStatus, editRetail1, editRetail2, editRetail3, editComments1, editComments2, editComments3, editCountryName)
	ElseIf (updateMethod = "update") Then
		Call updateCustomer(editCustID, editSignupDate, editFirstName, editLastName, editCompanyName, editTitleName, editAddress1Name, editAddress2Name, editCityName, editStateName, editZipName, editPhoneName, editEmailName, editSeminarName, editSeminarDate, editSeminarCity, linkedSeminar, editSeminarCost, editPaymentType, editDiscount, editPaid, editDeleted, editStatus, editRetail1, editRetail2, editRetail3, editComments1, editComments2, editComments3, editCountryName)
	Else
		Response.Write("<BR><B>Not a valid update method.</B><BR>")
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function addCustomer(funcCustID, funcSignupDate, funcFirstName, funcLastName, funcCompanyName, funcTitleName, funcAddress1Name, funcAddress2Name, funcCityName, funcStateName, funcZipName, funcPhoneName, funcEmailName, funcSeminarName, funcSeminarDate, funcSeminarCity, funcLinkedSeminar, funcSeminarCost, funcPaymentType, funcDiscount, funcPaid, funcDeleted, funcStatus, funcRetail1, funcRetail2, funcRetail3, funcComments1, funcComments2, funcComments3, funcCountryName)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("INSERT INTO tblCustomers ( " & _
		"dateSignup, " & _
		"strFirstName, " & _
		"strLastName, " & _
		"strCompanyName, " & _
		"strTitle, " & _
		"strAddress1, " & _
		"strAddress2, " & _
		"strCity, " & _
		"strState, " & _
		"strZip, " & _
		"strCountry, " & _
		"numPhone, " & _
		"strEmail, " & _
		"strSeminarName, " & _
		"dateSeminarDate, " & _
		"strSeminarCity, " & _
		"linkedSeminar, " & _
		"currencyCost, " & _
		"optPayment, " & _
		"isDiscount, " & _
		"isPaid, " & _
		"isDeleted, " & _
		"optStatus, " & _
		"strRetail1, " & _
		"strRetail2, " & _
		"strRetail3, " & _
		"strComments1, " & _
		"strComments2, " & _
		"strComments3 " & _
		") VALUES ( " & _
		"#" & funcSignupDate & "#, " & _
		"'" & funcFirstName & "', " & _
		"'" & funcLastName & "', " & _
		"'" & funcCompanyName & "', " & _
		"'" & funcTitleName & "', " & _
		"'" & funcAddress1Name & "', " & _
		"'" & funcAddress2Name & "', " & _
		"'" & funcCityName & "', " & _
		"'" & funcStateName & "', " & _
		"'" & funcZipName & "', " & _
		"'" & funcCountryName & "', " & _
		"'" & funcPhoneName & "', " & _
		"'" & funcEmailName & "', " & _
		"'" & funcSeminarName & "', " & _
		"#" & funcSeminarDate & "#, " & _
		"'" & funcSeminarCity & "', " & _
		"" & funcLinkedSeminar & ", " & _
		"" & funcSeminarCost & ", " & _
		"'" & funcPaymentType & "', " & _
		"" & funcDiscount & ", " & _
		"" & funcPaid & ", " & _
		"" & funcDeleted & ", " & _
		"'" & funcStatus & "', " & _
		"'" & funcRetail1 & "', " & _
		"'" & funcRetail2 & "', " & _
		"'" & funcRetail3 & "', " & _
		"'" & funcComments1 & "', " & _
		"'" & funcComments2 & "', " & _
		"'" & funcComments3 & "' " & _
		");")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function updateCustomer(funcCustID, funcSignupDate, funcFirstName, funcLastName, funcCompanyName, funcTitleName, funcAddress1Name, funcAddress2Name, funcCityName, funcStateName, funcZipName, funcPhoneName, funcEmailName, funcSeminarName, funcSeminarDate, funcSeminarCity, funcSeminarCost, funcPaymentType, funcDiscount, funcPaid, funcDeleted, funcStatus, funcRetail1, funcRetail2, funcRetail3, funcComments1, funcComments2, funcComments3, funcCountryName)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	'Set dbSeminar = dbConnection.Execute("SELECT * FROM tblSeminars WHERE numSeminarID = " &

	strSQL = "UPDATE tblCustomers SET dateSignup = #" & funcSignupDate & "#, strFirstName = '" & funcFirstName & "', strLastName = '" & funcLastName & "', strCompanyName = '" & funcCompanyName & "', strTitle = '" & funcTitleName & "', strAddress1 = '" & funcAddress1Name & "', strAddress2 = '" & funcAddress2Name & "', strCity = '" & funcCityName & "', strState = '" & funcStateName & "', strZip = '" & funcZipName & "', strCountry = '" & funcCountryName & "', numPhone = " & funcPhoneName & ", strEmail = '" & funcEmailName & "', strSeminarName = '" & funcSeminarName & "', dateSeminarDate = #" & funcSeminarDate & "#, strSeminarCity = '" & funcSeminarCity & "', currencyCost = " & funcSeminarCost & ", optPayment = '" & funcPaymentType & "', isDiscount = " & funcDiscount & ", isPaid = " & funcPaid & ", isDeleted = " & funcDeleted & ", optStatus = '" & funcStatus & "', strRetail1 = '" & funcRetail1 & "', strRetail2 = '" & funcRetail2 & "', strRetail3 = '" & funcRetail3 & "', strComments1 = '" & funcComments1 & "', strComments2 = '" & funcComments2 & "', strComments3 = '" & funcComments3 & "' WHERE numCustID = " & funcCustID
	Response.Write("<b>Execute:</b> " & strSQL & "<br />")
	Response.Write("Customer was not really updated.  It is being worked on.")
	Response.End
	'dbConnection.Execute(strSQL)

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function GetSeminarTime(whichTime, seminarName)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	If (whichTime = "start") Or (whichTime = "end") Then
		Set dbSeminarTime = dbConnection.Execute("SELECT " & whichTime & "Time FROM tblSeminarTimes WHERE effectsSeminar = '" & seminarName & "'")
		If (dbSeminarTime.EOF) Then
			If (whichTime = "start") Then
				GetSeminarTime = "8:30 AM"
			ElseIf (whichTime = "end") Then
				GetSeminarTime = "5:30 PM"
			End If
		Else
			GetSeminarTime = dbSeminarTime(whichTime & "Time")
		End If
		Set dbSeminarTime = Nothing
	Else
		If (whichTime = "start") Then
			GetSeminarTime = "8:30 AM"
		ElseIf (whichTime = "end") Then
			GetSeminarTime = "5:30 PM"
		End If
	End If
	
	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function sendConfirmationOfSignup(customerID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set dbCustomer = dbConnection.Execute("SELECT * FROM tblCustomers WHERE numCustID = " & customerID & "")

	Dim FSO
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")

	Set confirmEmail  = FSO.OpenTextFile(Server.MapPath("admin/emails/confirm_" & Lcase(dbCustomer("optPayment")) & ".txt"), 1, false)

	While NOT confirmEmail.AtEndOfStream
		emailLine = confirmEmail.ReadLine
		If (InStr(emailLine,"[First Name]") > 0) Then
			emailLine = Replace(emailLine,"[First Name]",dbCustomer("strFirstName"))
		End If
		If (InStr(emailLine,"[Last Name]") > 0) Then
			emailLine = Replace(emailLine,"[Last Name]",dbCustomer("strLastName"))
		End If
		If (InStr(emailLine,"[Company Name]") > 0) Then
			emailLine = Replace(emailLine,"[Company Name]",dbCustomer("strCompanyName"))
		End If
		If (InStr(emailLine,"[Address]") > 0) Then
			addressLine = vbNewline & vbTab & dbCustomer("strAddress1") & vbNewline
			If (Len(dbCustomer("strAddress2")) > 0) Then
				addressLine = addressLine & vbTab & dbCustomer("strAddress2") & vbNewline
			End If
			addressLine = addressLine & vbTab & dbCustomer("strCity") & ", " & dbCustomer("strState") & " " & dbCustomer("strZip") & vbNewline
			addressLine = addressLine & vbTab & dbCustomer("strCountry")
			emailLine = Replace(emailLine,"[Address]",addressLine)
			addressLine = ""
		End If
		If (InStr(emailLine,"[Phone Number]") > 0) Then
			phoneLine = "(" & Left(dbCustomer("numPhone"),3) & ") " & Right(Left(dbCustomer("numPhone"),6),3) & "-" & Right(dbCustomer("numPhone"),Len(dbCustomer("numPhone"))-6)
			emailLine = Replace(emailLine,"[Phone Number]",phoneLine)
			phoneLine = ""
		End If
		If (InStr(emailLine,"[Email Address]") > 0) Then
			emailLine = Replace(emailLine,"[Email Address]",dbCustomer("strEmail"))
		End If
		If (InStr(emailLine,"[Seminar Name]") > 0) Then
			emailLine = Replace(emailLine,"[Seminar Name]",dbCustomer("strSeminarName"))
		End If
		If (InStr(emailLine,"[Seminar Start Time]") > 0) Then
			emailLine = Replace(emailLine,"[Seminar Start Time]",GetSeminarTime("start",dbCustomer("strSeminarName")))
		End If
		If (InStr(emailLine,"[Seminar End Time]") > 0) Then
			emailLine = Replace(emailLine,"[Seminar End Time]",GetSeminarTime("end",dbCustomer("strSeminarName")))
		End If
		If (InStr(emailLine,"[Seminar City]") > 0) Then
			emailLine = Replace(emailLine,"[Seminar City]",dbCustomer("strSeminarCity"))
		End If
		If (InStr(emailLine,"[Seminar Date]") > 0) Then
			If (Fix(dbCustomer("linkedSeminar")) = -1) Then
				emailLine = Replace(emailLine,"[Seminar Date]",dbCustomer("dateSeminarDate"))
			Else
				Set dbDates = dbConnection.Execute("SELECT dateSeminarDate FROM tblSeminars WHERE linkedToSeminar = " & dbCustomer("linkedSeminar") & "")
				If (dbDates.EOF) Then
					Response.Write("There was a problem getting the seminar dates.<br />")
					Response.End
				Else
					strAllDates = ""
					Do Until dbDates.EOF
						strAllDates = strAllDates & ", " & dbDates("dateSeminarDate")
						dbDates.MoveNext
					Loop
					strAllDates = Right(strAllDates,Len(strAllDates)-2)
					emailLine = Replace(emailLine,"[Seminar Date]",strAllDates)
				End If
			End If
		End If
		If (InStr(emailLine,"[Seminar Price]") > 0) Then
			emailLine = Replace(emailLine,"[Seminar Price]",dbCustomer("currencyCost"))
		End If
		If (InStr(emailLine,"[14 days prior]") > 0) Then
			emailLine = Replace(emailLine,"[14 days prior]",DateAdd("d",DateValue(dbCustomer("dateSeminarDate")),-14))
		End If
		emailMessage = emailMessage & emailLine & vbNewline
	Wend

	Call SMTP(fromAddress,dbCustomer("strEmail"),dbCustomer("strSeminarName") & " Registration Confirmation",emailMessage)
	Call SMTP(fromAddress,registrationConfirmationAddress,dbCustomer("strSeminarName") & " Registration Confirmation",emailMessage)

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function markAsPaid(custID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("UPDATE tblCustomers SET strRetail1 = 'Free Seminar', isPaid = true, optStatus = 'Paid' WHERE numCustID = " & custID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function SMTP(mailFrom,mailTo,mailSubject,mailBody)
	'Windows NT4 Emailer
	'lsResult = CDONTSMail(mailFrom,mailTo,mailSubject,mailBody)
	'Windows 2000 Emailer

	'ASPMail Emailer
	'lsResult = ASPMail(mailFrom,mailTo,mailSubject,mailBody)

	'EasyMail Objects v5
	lsResult = EasyMail(mailFrom,mailTo,mailSubject,mailBody)
End Function

Function CDONTSMail(mailFrom,mailTo,mailSubject,mailBody)
	On Error Resume Next
	Dim MyCDONTSMail2
	Dim HTML
	Set MyCDONTSMail2 = CreateObject("CDONTS.NewMail")
	MyCDONTSMail2.From = mailFrom
	MyCDONTSMail2.To = mailTo
	MyCDONTSMail2.Subject = mailSubject
	MyCDONTSMail2.BodyFormat = 0
	MyCDONTSMail2.MailFormat = 0
	MyCDONTSMail2.Body = mailBody
	MyCDONTSMail2.Send
	Set MyCDONTSMail2=nothing
	If Err <> 0 Then
		CDONTSMail = "CDONTS Email Failed<br>" & _
		Err.Description & "<br>" & _
		Err.Source & "<br>"
	Else
		CDONTSMail = "CDONTS Successful<br>"
	End If
End Function

Function ASPMail(mailFrom,mailTo,mailSubject,mailBody)
	Set mailer = Server.CreateObject("ASPMAIL.ASPMailCtrl,1")

	ASPMail = mailer.SendMail(outgoingMailServer, mailTo, mailFrom, mailSubject, mailBody)
'	If (result = "") Then
'		'Mail has been sent.
'		Response.Write("Sent email.")
'		Response.Redirect
'	Else
'		'Mail was not able to be sent.
'		Response.Write("There was a problem sending the mail:<br><br><i>" & result & "</i>")
'	End If
End Function

Function EasyMail(mailFrom, mailTo, mailSubject, mailBody)
	Set ezMailer = Server.CreateObject("EasyMail.SMTP.5")
	ezMailer.LicenseKey = "Unregistered User/S10I510R1AX70C0Rb600"
	ezMailer.MailServer = outgoingMailServer
	ezMailer.FromAddr = mailFrom
	ezMailer.AddRecipient "", mailTo, 1
'	ezMailer.AddRecipient "", "syntax-cart@sisna.com", 3
	ezMailer.Subject = mailSubject
	ezMailer.BodyText = mailBody
	ezMailer.AddCustomHeader "Return-Path", "<support@vasst.com>"
	EasyMail = ezMailer.Send
'	If ezReturn = 0 Then
'	  Response.Write "Message sent successfully."
'	Else
'	  Response.Write "There was an error sending your message.  Error: " & CStr(ezReturn)
'	End If
	Set ezMailer = Nothing
End Function

Dim verisignName, verisignTour, verisignPaymentAmount, verisignAddress, verisignCity, verisignState, verisignZip, verisignPhone, verisignCountry, verisignEmail, verisignCustID
Function loadVerisignForm(custID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

    if (custID = "") Then
		loadVerisignForm = false
	Else
		Set dbCustomer = dbConnection.Execute("SELECT * FROM tblCustomers WHERE numCustID = " & custID)
		if not dbCustomer.EOF then
			verisignName = dbCustomer("strFirstName") & " " & dbCustomer("strLastName")
			verisignTour = dbCustomer("strSeminarName")
			verisignPaymentAmount = dbCustomer("currencyCost")
			verisignAddress = dbCustomer("strAddress1")
			verisignCity = dbCustomer("strCity")
			verisignState = dbCustomer("strState")
			verisignZip = dbCustomer("strZip")
			verisignPhone = dbCustomer("numPhone")
			verisignCountry = dbCustomer("strCountry")
			verisignEmail = dbCustomer("strEmail")
			verisignCustID = custID
			loadVerisignForm = true
		else
			loadVerisignForm = false
		end if
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printVerisignConfirmation(custID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

    if (custID = "") Then
		loadVerisignForm = false
	Else
		Set dbCustomer = dbConnection.Execute("SELECT * FROM tblCustomers WHERE numCustID = " & custID)
		if not dbCustomer.EOF then
		    response.write(dbCustomer("strFirstName") & " " & dbCustomer("strLastName") & ",<BR>")
			response.write("These are the details we have entered into our database:<br><br>")
			response.write("<small>" & dbCustomer("strRetail1") & "<br></small>")

			If (dbCustomer("isPaid")) Then
				Response.Write("<br>Your account is paid for, so everything is in order, we look forward to seeing you at the seminar.</br>")
			Else
				Response.Write("<br>It looks as if something didn't go through correctly, please fix the problem, then goto <a href='/registration/register.asp'>http://www.sundancemediagroup.com/registration/register.asp</a>, and try again.<br>")
			End If
		end if
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printVerisignTransactions
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	set objRS=server.createObject("ADODB.Recordset")
	strSQL = "SELECT * FROM tblTransactionLog ORDER BY transactionDate DESC"
	objRS.open strSQL, dbConnection, 1, 2

	Response.Write(tabTo(0) & "<STYLE>")
	Response.Write(tabTo(1) & "TABLE.spreadsheet { Border-Left: 1px solid black; }")
	Response.Write(tabTo(1) & "TR.heading { Background: black; Color: white; Font-Weight: bold; }")
	Response.Write(tabTo(1) & "TD { Font-Size: 9pt; }")
	Response.Write(tabTo(1) & "INPUT.field { Font-Size: 9pt; }")
	Response.Write(tabTo(0) & "</STYLE>")
	Response.Write(tabTo(2) & "<TABLE BORDER=1 CELLPADDING=2 CELLSPACING=0 CLASS=spreadsheet STYLE=""Border-Collapse: collapse; Border-Color: black;"">")
	Response.Write(tabTo(3) & "<TR>")
	Response.Write(tabTo(4) & "<TD COLSPAN=" & objRS.Fields.Count & " CLASS=spreadsheet>")
	Response.Write(tabTo(5) & "<p align=left><font size=+2><b>Verisign Transaction Log</b></font></p>")
	Response.Write(tabTo(4) & "</TD>")
	Response.Write(tabTo(3) & "</TR>")
	Response.Write(tabTo(3) & "<TR>")
	For Each field In objRS.fields
		Response.Write(tabTo(4) & "<TD BGCOLOR=aaaaaa nowrap><font color=333333><b>" & field.name & "</b></font></TD>")
	Next
	Response.Write(tabTo(3) & "<TR>")

	If (objRS.EOF) Then
		Response.Write("<TR><TD NOWRAP COLSPAN=" & objRS.Fields.Count & ">No verisign transactions to show.</TD></TR>")
	Else
		Do Until objRS.EOF
			Response.Write(tabTo(3) & "<TR>")
			For Each field In objRS.fields
				Response.Write(tabTo(4) & "<TD NOWRAP>" & field & "</TD>")
			Next
			Response.Write(tabTo(3) & "<TR>")
			objRS.MoveNext
		Loop
	End If
	Response.Write(tabTo(2) & "</TABLE>")

	objRS.close
	Set objRS = Nothing

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function SaveTransaction(verisignResult, verisignCustID, verisignRefID, verisignResponseMSG, verisignAuthCode, transactionSuccess)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("INSERT INTO tblTransactionLog ( transactionDate, transactionResultCode, transactionResultMessage, transactionID, transactionAuthCode, transactionCustID, transactionHttpReferer, transactionIsSuccessful) VALUES ( #" & date & " " & time & "#, '" & verisignResult & "', '" & verisignResponseMSG & "', '" & verisignRefID & "', '" & verisignAuthCode & "', '" & verisignCustID & "', '" & Request.ServerVariables("HTTP_REFERER") & "', " & transactionSuccess & " )")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function saveVerisignInfo(verisignResult, verisignCustID, verisignRefID, verisignResponseMSG, verisignAuthCode, transactionSuccess)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dataRetail1 = "Result: " & verisignResult & " - " & verisignResponseMSG
	dataRetail2 = "ID: " & verisignRefID & " AuthCode: " & verisignAuthCode
	dataRetail3 = "Completed at " & date & " " & time
	referer = "Referer: " & Request.ServerVariables("HTTP_REFERER")

	If (transactionSuccess) Then
		dataRetail = "Transaction Successful<BR>" & dataRetail1 & "<BR>" & dataRetail2 & "<BR>" & dataRetail3 & "<BR>" & referer
		dbConnection.Execute("UPDATE tblCustomers SET strRetail1 = '" & dataRetail & "', isPaid = true, optStatus = 'Paid' WHERE numCustID = " & verisignCustID & "")
	Else
		dataRetail = "Transaction Failed<BR>" & dataRetail1 & "<BR>" & dataRetail2 & "<BR>" & dataRetail3 & "<BR>" & referer
		verisignAuthCode = "N/A"
		dbConnection.Execute("UPDATE tblCustomers SET strRetail1 = '" & dataRetail & "', isPaid = false, optStatus = 'Waiting For Payment' WHERE numCustID = " & verisignCustID & "")
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function getVerisignResultCode(theCode)
	Select Case theCode
	Case 0
		codeReturn = "Approved"
	Case 1
		codeReturn = "User authentication failed"
	Case 2
		codeReturn = "Invalid tender type. Your merchant bank account does not support the following credit card type that was submitted."
	Case 3
		codeReturn = "Invalid transaction type. Transaction type is not appropriate for this transaction. For example, you cannot credit an authorization-only transaction."
	Case 4
		codeReturn = "Invalid amount format"
	Case 5
		codeReturn = "Invalid merchant information. Processor does not recognize your merchant account information. Contact your bank account acquirer to resolve this problem."
	Case 7
		codeReturn = "Field format error. Invalid information entered. See RESPMSG."
	Case 8
		codeReturn = "Not a transaction server"
	Case 9
		codeReturn = "Too many parameters or invalid stream"
	Case 10
		codeReturn = "Too many line items"
	Case 11
		codeReturn = "Client time-out waiting for response"
	Case 12
		codeReturn = "Declined. Check the credit card number and transaction information to make sure they were entered correctly. If this does not resolve the problem, have the customer call the credit card issuer to resolve."
	Case 13
		codeReturn = "Referral. Transaction was declined but could be approved with a verbal authorization from the bank that issued the card. Submit a manual Voice Authorization transaction and enter the verbal auth code."
	Case 19
		codeReturn = "Original transaction ID not found. The transaction ID you entered for this transaction is not valid. See RESPMSG."
	Case 20
		codeReturn = "Cannot find the customer reference number"
	Case 22
		codeReturn = "Invalid ABA number"
	Case 23
		codeReturn = "Invalid account number. Check credit card number and re-submit."
	Case 24
		codeReturn = "Invalid expiration date. Check and re-submit."
	Case 25
		codeReturn = "Invalid Host Mapping. Transaction type not mapped to this host"
	Case 26
		codeReturn = "Invalid vendor account"
	Case 27
		codeReturn = "Insufficient partner permissions"
	Case 28
		codeReturn = "Insufficient user permissions"
	Case 29
		codeReturn = "Invalid XML document. This could be caused by an unrecognized XML tag or a bad XML format that cannot be parsed by the system."
	Case 30
		codeReturn = "Duplicate transaction"
	Case 31
		codeReturn = "Error in adding the recurring profile"
	Case 32
		codeReturn = "Error in modifying the recurring profile"
	Case 33
		codeReturn = "Error in canceling the recurring profile"
	Case 34
		codeReturn = "Error in forcing the recurring profile"
	Case 35
		codeReturn = "Error in reactivating the recurring profile"
	Case 36
		codeReturn = "OLTP Transaction failed"
	Case 50
		codeReturn = "Insufficient funds available in account"
	Case 99
		codeReturn = "General error. See RESPMSG."
	Case 100
		codeReturn = "Transaction type not supported by host"
	Case 101
		codeReturn = "Time-out value too small"
	Case 102
		codeReturn = "Processor not available"
	Case 103
		codeReturn = "Error reading response from host"
	Case 104
		codeReturn = "Timeout waiting for processor response. Try your transaction again."
	Case 105
		codeReturn = "Credit error. Make sure you have not already credited this transaction, or that this transaction ID is for a creditable transaction. (For example, you cannot credit an authorization.)"
	Case 106
		codeReturn = "Host not available"
	Case 107
		codeReturn = "Duplicate suppression time-out"
	Case 108
		codeReturn = "Void error. See RESPMSG. Make sure the transaction ID entered has not already been voided. If not, then look at the Transaction Detail screen for this transaction to see if it has settled. (The Batch field is set to a number greater than zero if the transaction has been settled). If the transaction has already settled, your only recourse is a reversal (credit a payment or submit a payment for a credit)."
	Case 109
		codeReturn = "Time-out waiting for host response"
	Case 111
		codeReturn = "Capture error. Only authorization transactions can be captured."
	Case 112
		codeReturn = "Failed AVS check. Address and ZIP code do not match. An authorization may still exist on the cardholder’s account."
'	Case 113
'		codeReturn = "Cannot exceed sales cap. For ACH transactions only."
	Case 113
		codeReturn = "Merchant sale total will exceed the cap with current transaction"
	Case 114
		codeReturn = "Card Security Code (CSC) Mismatch. An authorization may still exist on the cardholder’s account."
	Case 115
		codeReturn = "System busy, try again later"
	Case 116
		codeReturn = "VPS Internal error - Failed to lock terminal number"
	Case 117
		codeReturn = "Failed merchant rule check. An attempt was made to submit a transaction that failed to meet the security settings specified on the VeriSign Manager Security Settings page. See VeriSign Manager User’s Guide."
	Case 118
		codeReturn = "Invalid keywords found in string fields"
	Case 1000
		codeReturn = "Generic host error. See RESPMSG. This is a generic message returned by your credit card processor. The message itself will contain more information describing the error."
	Case Else
		codeReturn = "Unknown"
	End Select
	getVerisignResultCode = theCode & " - " & codeReturn
End Function

Dim paypalTour, paypalName, paypalCustID, paypalPaymentAmount
Function loadPaypalForm(custID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

    if (custID = "") Then
		loadPaypalForm = false
	Else
		Set dbCustomer = dbConnection.Execute("SELECT * FROM tblCustomers WHERE numCustID = " & custID)
		if not dbCustomer.EOF then
			paypalName = dbCustomer("strFirstName") & " " & dbCustomer("strLastName")
			paypalTour = dbCustomer("strSeminarName")
			paypalPaymentAmount = dbCustomer("currencyCost")
			paypalCustID = custID
			loadPaypalForm = true
		else
			loadPaypalForm = false
		end if
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function humanize(fieldName)
	fieldName = Replace(fieldName,"numCustID","Customer ID #")
	fieldName = Replace(fieldName,"dateSignup","Date Of Signup")
	fieldName = Replace(fieldName,"strFirstName","First Name")
	fieldName = Replace(fieldName,"strLastName","Last Name")
	fieldName = Replace(fieldName,"strCompanyName","Company Name")
	fieldName = Replace(fieldName,"strTitle","Company Title")
	fieldName = Replace(fieldName,"strAddress1","Address 1")
	fieldName = Replace(fieldName,"strAddress2","Address 2")
	fieldName = Replace(fieldName,"strCity","City")
	fieldName = Replace(fieldName,"strState","State")
	fieldName = Replace(fieldName,"strZip","Zip")
	fieldName = Replace(fieldName,"strEmail","Email Address")
	fieldName = Replace(fieldName,"numPhone","Phone Number")
	fieldName = Replace(fieldName,"strSeminarName","Seminar Name")
	fieldName = Replace(fieldName,"dateSeminarDate","Seminar Date")
	fieldName = Replace(fieldName,"strSeminarCity","Seminar City")
	fieldName = Replace(fieldName,"currencyCost","Seminar Cost")
	fieldName = Replace(fieldName,"optPayment","Type Of Payment")
	fieldName = Replace(fieldName,"isDiscount","Discounted?")
	fieldName = Replace(fieldName,"isPaid","Paid?")
	fieldName = Replace(fieldName,"isDeleted","Deleted?")
	fieldName = Replace(fieldName,"optStatus","Status")
	fieldName = Replace(fieldName,"strRetail1","Verisign Info")
	fieldName = Replace(fieldName,"strRetail2","Retail 2")
	fieldName = Replace(fieldName,"strRetail3","Hear About")
	fieldName = Replace(fieldName,"strComments1","Comments 1")
	fieldName = Replace(fieldName,"strComments2","Comments 2")
	fieldName = Replace(fieldName,"strComments3","Comments 3")
	humanize = fieldName
End Function

Function subscribeToList(myListID, subName, subEmail)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set checkList = dbConnection.Execute("SELECT * FROM emailer_tblMailingLists WHERE listID = " & myListID & " AND listAllowSubscribe = True")
	If Not (checkList.EOF) Then
		dbConnection.Execute("INSERT INTO emailer_tblMembers ( memberName, memberEmail, memberOf ) VALUES ( '" & subName & "', '" & subEmail & "', " & myListID & " )")
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

'''''''''''''''''''''
'' DEBUG FUNCTIONS ''
'''''''''''''''''''''
Function splitCookies
	tmpHTML = "<B>Cookies</B><BR>" & vbCrLf
	For each objCK in Request.Cookies
		For each objSubCK in Request.Cookies(objCK)
			tmpHTML = tmpHTML & "(" & objCK & ")(" & objSubCK & ")=" & Request.Cookies(objCK)(objSubCK) & "<BR>" & vbCrLf
		Next
	Next
	Response.Write(tmpHTML)
	splitCookies = tmpHTML
	tmpHTML = ""
End Function

Function splitForm
	tmpHTML = "<B>Form</B><BR>" & vbCrLf
	For each objForm in Request.Form
		tmpHTML = tmpHTML & objForm & "=" & Request.Form(objForm) & "<BR>" & vbCrLf
	Next
	Response.Write(tmpHTML)
	splitForm = tmpHTML
	tmpHTML = ""
End Function

Function splitQueryString
	tmpHTML = "<B>QueryString</B><BR>" & vbCrLf
	For each objQS in Request.QueryString
		tmpHTML = tmpHTML & objQS & "=" & Request.QueryString(objQS) & "<BR>" & vbCrLf
	Next
	Response.Write(tmpHTML)
	splitQueryString = tmpHTML
	tmpHTML = ""
End Function

Function splitServerVariables
	tmpHTML = "<B>Server Variables</B><BR>" & vbCrLf
	For Each objSV In Request.ServerVariables
		tmpHTML = tmpHTML & objSV & "=" & Request.ServerVariables(objSV) & "<BR>" & vbCrLf
	Next
	Response.Write(tmpHTML)
	splitServerVariables = tmpHTML
	tmpHTML = ""
End Function

'Loads the strings that are needed throughout the Management System, stored in tblOptions.
loadOptions

'Stuff to run every load of this page.
If Not (authNotNeeded) Then
	'Check to make sure user is logged in, if not forward to login page.
	checkAuth
End If
%>