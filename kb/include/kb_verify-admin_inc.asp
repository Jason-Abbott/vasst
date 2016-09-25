<%
If CStr(GetSessionValue(g_USER_TYPE)) <> CStr(g_USER_ADMIN) Then
	dim m_oTempData
	Call goLogin()
	Set m_oTempData = New kbDataAccess
	Call m_oTempData.LogActivity(g_ACT_UNAUTHORIZED_ACCESS, "", "", "", "", "")
	Set m_oTempData = Nothing
	response.redirect g_sDEFAULT_PAGE
End If
%>