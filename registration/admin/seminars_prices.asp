<!--#include file="functions.asp"-->
<% 
If Not (getUserInfo(getUserID,"accesslevel") = "Admin") Then 
	Response.Write "<CENTER><B>You do not have access to this section of the CMS, your access is Read-Only, contact your admin to get rights.</B></CENTER>"
	Response.End
End If
%>
<% 
printHeader

'Response.Write(Request.Form)
For Each frm In Request.Form
	If (InStr(frm,"remove") > 0) Then
		removeFound = true
		removeID = frm
		Exit For
	End If
Next
	
If (removeFound) Then
	printSeminarsPrices removeID
	Response.Redirect "seminars_prices.asp"
ElseIf (Request.Form("saveChanges") = "Update/Save") Then
	printSeminarsPrices "save"
	Response.Write("<script> alert('Saved Changes'); this.location.href = 'seminars_prices.asp'; </script>")
ElseIf (Request.Form("addPrice") = "Add New Price") Then
	printSeminarsPrices "add"
Else
	printSeminarsPrices "view"
End If

printBoxClose
printFooter 
%>