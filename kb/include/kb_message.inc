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
			.write "<table cellspacing='0' border='0' cellpadding='0' class='Message' align='center'>"
			.write "<tr><td class='MessageSymbol' valign='middle'>&#118;</td>"
			.write "<td class='Message' valign='top'>"
			.write sMessage
			.write "</td></table>"
		end with
		Call SetSessionValue(g_USER_MSG, "")
	end if
End Sub
%>