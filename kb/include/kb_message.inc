<%
Call WriteMessage()

'-------------------------------------------------------------------------
'	Name: 		WriteMessage()
'	Purpose: 	write message
'Modifications:
'	Date:		Name:	Description:
'	12/28/02	JEA		Creation
'-------------------------------------------------------------------------
Sub WriteMessage()
	dim sMessage
	sMessage = GetSessionValue(g_USER_MSG)
	if sMessage <> "" then
		with response
			.write "<center><div class='Message'>"
			.write sMessage
			.write "</div></center>"
		end with
		Call SetSessionValue(g_USER_MSG, "")
	end if
End Sub
%>