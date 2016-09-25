<%
dim g_bAdmin
dim g_lSiteID
if GetSessionValue(g_USER_ID) = "0" or GetSessionValue(g_USER_ID) = "" then Call goLogin()
if GetSessionValue(g_USER_STATUS) = CStr(g_STATUS_PENDING) then response.redirect "kb_validate.asp?url=" & GetURL(false)
g_bAdmin = CBool(GetSessionValue(g_USER_TYPE) = CStr(g_USER_ADMIN))
g_lSiteID = GetSessionValue(g_USER_SITE)

'-------------------------------------------------------------------------
'	Name: 		goLogin()
'	Purpose: 	redirect to login page
'Modifications:
'	Date:		Name:	Description:
'	12/23/02	JEA		Created
'	12/28/02	JEA		Login from cookie if possible
'	4/26/03		JEA		Track site ID
'-------------------------------------------------------------------------
Sub goLogin()
	dim oUser
	dim bCookieLogin
	dim lSiteID
	
	lSiteID = Request.QueryString("s")
	Set oUser = New kbUser
	bCookieLogin = oUser.LoginWithCookie(lSiteID)
	Set oUser = nothing
	If Not bCookieLogin Then response.redirect "kb_login.asp?s=" & ReplaceNull(lSiteID, g_DEFAULT_SITE) & "&url=" & GetURL(false)
End Sub
%>
