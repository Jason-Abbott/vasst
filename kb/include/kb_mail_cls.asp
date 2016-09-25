<%
'-------------------------------------------------------------------------
'	Name: 		kbMail class
'	Purpose: 	methods for sending e-mail
'Modifications:
'	Date:		Name:	Description:
'	1/1/03		JEA		Creation
'	1/7/03		JEA		Updated to use syntax for ASPmail object
'-------------------------------------------------------------------------
Class kbMail

	'-------------------------------------------------------------------------
	'	Name: 		SendMail()
	'	Purpose: 	send e-mail
	'	Return:		boolean
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	10/27/04	JEA		Updated to use CDO
	'-------------------------------------------------------------------------
	Private Function SendMail(ByVal v_sMailFrom, ByVal v_sMailTo, ByVal v_sSubject, ByVal v_sBody)
		dim oMail
		Set oMail = Server.CreateObject(g_sEMAIL_OBJECT)
		'SendMail = CBool(oMail.SendMail(g_sEMAIL_SERVER, v_sMailTo, v_sMailFrom, v_sSubject, v_sBody) = "")
		
		With oMail
			.From = v_sMailFrom
			.to = v_sMailTo
			.Subject = v_sSubject
			.TextBody = v_sBody
			'.HTMLBody = v_sBody
		End With
		
		With oMail.Configuration.Fields
			.Item(cdoSchema & "sendusing") = cdoSendUsingPort
			.Item(cdoSchema & "smtpserver") = g_sEMAIL_SERVER
			.Item(cdoSchema & "smtpauthenticate") = cdoBasic
			.Item(cdoSchema & "sendusername") = g_sEMAIL_USERNAME
			.Item(cdoSchema & "sendpassword") = g_sEMAIL_PASSWORD
			.Item(cdoSchema & "smtpserverport") = 25
			.Item(cdoSchema & "smtpusessl") = False
			.Item(cdoSchema & "smtpconnectiontimeout") = 60
			.Update
		End With
		
		oMail.Send
		Set oMail = Nothing
		SendMail = true
	End Function

	'-------------------------------------------------------------------------
	'	Name: 		SendPasswordEmail()
	'	Purpose: 	send e-mail with confirmation string
	'	Return:		boolean
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	10/27/04	JEA		Added extra logging parameter
	'-------------------------------------------------------------------------
	Public Function SendPasswordEmail(ByVal v_sEmail)
		Const FIRST_NAME = 0
		Const LAST_NAME = 1
		Const EMAIL = 2
		Const PASSWORD = 3
		dim bSuccess
		dim oData
		dim aData
		
		bSuccess = False
		aData = GetUserEmailData(GetSessionValue(g_USER_ID), v_sEmail)
		If IsArray(aData) Then
			bSuccess = SendMail(g_sEMAIL_FROM, _
				aData(EMAIL, 0), _
				"Your " & g_sORG_NAME & " Password", _
				aData(FIRST_NAME, 0) & "," & vbCrLf & vbCrLf & "Your password is " _
					& aData(PASSWORD, 0) & "." & vbCrLf & vbCrLf & g_sORG_NAME)
			If bSuccess Then
				Set oData = New kbDataAccess
				Call oData.LogActivity(g_ACT_EMAILED_PASSWORD, "", "", "", "", aData(EMAIL, 0), aData(PASSWORD, 0))
				Set oData = Nothing
			End If
		End If
		SendPasswordEmail = bSuccess
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		SendValidationEmail()
	'	Purpose: 	resend e-mail with confirmation string
	'Modifications:
	'	Date:		Name:	Description:
	'	1/14/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub SendValidationEmail(ByVal v_lUserID)
		Const FIRST_NAME = 0
		Const LAST_NAME = 1
		Const EMAIL = 2
		Const CODE = 4
		dim aData
		dim oData
		dim bSuccess
		
		bSuccess = false
		aData = GetUserEmailData(v_lUserID, "")
		If IsArray(aData) Then
			bSuccess = SendMail(g_sEMAIL_FROM, _
				aData(EMAIL, 0), _
				"Your " & g_sORG_NAME & " Validation Code", _
				aData(FIRST_NAME, 0) & "," & vbCrLf & vbCrLf & "Your code is " _
					& aData(CODE, 0) & "." & vbCrLf & vbCrLf & g_sORG_NAME)
		End If
		If Not bSuccess Then
			Call SetSessionValue(g_USER_MSG, "Sorry, an error was encountered while trying to send to " & aData(EMAIL, 0))
		Else
			Set oData = New kbDataAccess
			Call oData.LogActivity(g_ACT_EMAILED_CODE, "", "", v_lUserID, "", "")
			Set oData = Nothing
			Call SetSessionValue(g_USER_MSG, "Your validation code has been sent to " & aData(EMAIL, 0))
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		SendConfirmationMail()
	'	Purpose: 	send e-mail with confirmation string
	'	Return:		string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub SendConfirmationMail(ByVal v_lUserID, ByVal v_sFirstName, ByVal v_sLastName, _
		ByVal v_sEmail, ByVal v_sConfirmationCode)
		
		dim bSuccess
		bSuccess = SendMail(g_sEMAIL_FROM, _
			v_sEmail, _
			g_sEMAIL_SUBJECT & " with " & g_sORG_NAME, _
			v_sFirstName & "," & vbCrLf & vbCrLf & "Thank you for registering with us. " _
				& "To activate your registration, please sign in and enter this code:" & vbCrLf & vbCrLf _
				& v_sConfirmationCode & vbCrLf & vbCrLf & "Thank You!" & vbCrLf & vbCrLf & g_sORG_NAME)
		If Not bSuccess Then
			Call SetSessionValue(g_USER_MSG, "Sorry, an error was encountered while sending to " & v_sEmail)
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		SendApprovalEmail()
	'	Purpose: 	send e-mail with confirmation string
	'	Return:		string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Creation
	'	7/21/04		JEA		Abstract for various item types
	'-------------------------------------------------------------------------
	Public Sub SendApprovalEmail(ByVal v_lItemID, ByVal v_lItemTypeID)
		Const FIRST_NAME = 0
		Const EMAIL = 1
		Const ITEM_NAME = 2
		Const SUBMIT_DATE = 3
		dim oUser
		dim sQuery
		dim aData
		dim bSuccess

		Set oUser = New kbUser
		aData = oUser.GetUserItemData(v_lItemID, v_lItemTypeID)
		Set oUser = Nothing

		If IsArray(aData) Then
			bSuccess = SendMail(g_sEMAIL_FROM, _
				aData(EMAIL, 0), _
				"Your submission to " & g_sORG_NAME, _
				aData(FIRST_NAME, 0) & "," & vbCrLf & vbCrLf & "The file you submitted on " _
					& FormatDate(aData(SUBMIT_DATE, 0)) & ", " & aData(ITEM_NAME, 0) & ", has been " _
					& "approved.  Thank you again for your submission." & vbCrLf & vbCrLf & g_sORG_NAME)
			If Not bSuccess Then
				Call SetSessionValue(g_USER_MSG, "Sorry, an error was encountered while trying to send e-mail")
			End If
		End If
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		GetUserEmailData()
	'	Purpose: 	get user info for e-mailing
	'	Return:		array
	'Modifications:
	'	Date:		Name:	Description:
	'	1/1/03		JEA		Creation
	'	1/14/03		JEA		return validation code
	'	3/18/03		JEA		only check for e-mail if present
	'-------------------------------------------------------------------------
	Private Function GetUserEmailData(ByVal v_lUserID, ByVal v_sEmail)
		dim sQuery
		dim oData
		sQuery = "SELECT vsFirstName, vsLastName, vsEmail, vsPassword, sValidationCode FROM tblUsers WHERE " _
			& IIf(IsVoid(v_sEmail), "lUserID = " & MakeNumber(v_lUserID), "vsEmail = '" & v_sEmail & "'")
		Set oData = New kbDataAccess
		GetUserEmailData = oData.GetArray(sQuery)
		Set oData = Nothing
	End Function
	
	
	'-------------------------------------------------------------------------
	'	Name: 		SendUserToUserEmail()
	'	Purpose: 	send e-mail to user from user
	'Modifications:
	'	Date:		Name:	Description:
	'	1/5/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub SendUserToUserEmail(ByVal v_lFromID, ByVal v_lToID, ByVal v_sSubject, _
		ByVal v_sBody, ByVal v_bCopySelf)
		
		Const FIRST_NAME = 0
		Const LAST_NAME = 1
		Const EMAIL = 2
		dim aSender
		dim aRecipient
		dim bSuccess
		
		aSender = GetUserEmailData(v_lFromID, "")
		aRecipient = GetUserEmailData(v_lToID, "")
		
		If IsArray(aSender) And IsArray(aRecipient) Then
			bSuccess = SendMail(aSender(EMAIL, 0), _
				aRecipient(EMAIL, 0), _
				v_sSubject, _
				v_sBody & vbCrLf & vbCrLf & "[sent on behalf of " & aSender(FIRST_NAME, 0) _
					& " " & aSender(LAST_NAME, 0) & " by " & g_sORG_NAME & "]")
			If bSuccess Then
				Call SetSessionValue(g_USER_MSG, "Your message has been sent")
			Else
				Call SetSessionValue(g_USER_MSG, "Sorry, an error occurred while sending your message")
			End If
		Else
			Call SetSessionValue(g_USER_MSG, "Sorry, an error occurred while sending your message")
		End If
	End Sub
End Class
%>