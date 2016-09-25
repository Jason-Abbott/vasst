<!--#include file="functions.asp"-->
<FORM METHOD=post>
<% 
If Not (getUserInfo(getUserID,"accesslevel") = "Admin") Then 
	Response.Write "<CENTER><B>You do not have access to this section of the CMS, your access is Read-Only, contact your admin to get rights.</B></CENTER>"
	Response.End
End If
%>
<% 
printHeader

Response.Write(Request.Form)
If (Request.Form("howMany") = "") Then
	addDeleteButton = True
	printSeminars 
Else
	For xi = 1 to Request.Form("howMany")
		If (Request.Form("cb" & xi) = "on") Then
			Call hideSeminar(Request.Form("removeid" & xi), True)
		Else
			Call hideSeminar(Request.Form("removeid" & xi), False)
		End If
'		Response.Write(xi & ":" & Request.Form("cb" & xi) & "<BR>")
	Next
	Response.Redirect("seminars_list.asp")
End If

printBoxClose
printFooter 
%>
</FORM>