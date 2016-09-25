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

'Response.Write(Request.Form)
If (Request.Form("howMany") = "") Then
	addDeleteButton = True
	permanentlyRemove = True
	printSeminars 
Else
	For irem = 1 to Request.Form("howMany")
		If (Request.Form("cb" & irem) = "on") Then
			Call removeSeminar(Request.Form("removeid" & irem))
		End If
'		Response.Write(irem & ":" & Request.Form("cb" & irem) & "<BR>")
	Next
	Response.Redirect("seminars_edit.asp")
End If

printBoxClose
printFooter 
%>
</FORM>