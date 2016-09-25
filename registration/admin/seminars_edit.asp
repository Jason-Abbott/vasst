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
If (Request.Form("saveChanges") = "Save") Then
'	Response.Write("<BR>")
	If (Request.Form("isVisible") = "on") Then
		isVisible = "True"
	Else
		isVisible = "False"
	End If

	If Not (isDate(Request.Form("dateSeminarDate"))) Then
		dateArray = Split(Request.Form("dateSeminarDate"),",")

		doDatesPass = True
		For X = 0 To Ubound(dateArray)
			dateArray(X) = Trim(dateArray(X))
			If Not (isDate(dateArray(X))) Then
				doDatesPass = False
			End If
		Next
		isDateArray = True
		strDates = Join(dateArray,",")
	Else
		doDatesPass = True
		isDateArray = False
		strDates = Request.Form("dateSeminarDate")
	End If


	If (doDatesPass) Then
		Call updateSeminar(Request.Form("numSeminarID"), Request.Form("strSeminarName"), strDates, Request.Form("strSeminarCity"), isVisible, isDateArray)
		Response.Redirect("seminars_edit.asp")
	Else
		addEditButton = True
		printSeminars
		Response.Write(errIncorrectDate)
	End If
ElseIf (Request.Form("edit") = "") Then
	addEditButton = True
	printSeminars
ElseIf (Request.Form("edit") > 0) Then
	addEditButton = True
	printSeminars
End If

printBoxClose
printFooter
%>