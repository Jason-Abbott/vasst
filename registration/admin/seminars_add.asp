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

If (Request.Form("strSeminarName") = "") OR (Request.Form("dateSeminarDate") = "") OR (Request.Form("strSeminarCity") = "") Then
	reloadValues = True
	addAddButton = True
	printSeminars
Else
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
		Call addSeminar(Request.Form("strSeminarName"), Request.Form("dateSeminarDate"), Request.Form("strSeminarCity"), isDateArray)
		Response.Redirect("seminars_add.asp")
	Else
		reloadValues = True
		addAddButton = True
		printSeminars
		Response.Write errIncorrectDate
	End If
End If
%>

</FORM>
<%
printBoxClose
printFooter
%>
<SCRIPT>
document.forms[0].dateSeminarDate.focus()
</SCRIPT>