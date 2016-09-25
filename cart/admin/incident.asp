<!--#include virtual="/cart/admin/includes.asp"-->
<%
	viewID = makeSafe(Request.QueryString("view"))
	action = makeSafe(Request.Form("action"))
	If (action = "Delete") Then
		openDB
		cartDB.Execute("DELETE FROM errors WHERE viewid = '" & viewID & "'")
		closeDB
		Response.Write("<script> /* alert('Deleted incident report.'); */ self.location.href = '/cart/admin/index.asp?function=ViewIncidents'; </script>")
		Response.Write("<noscript>Deleted incident report.</noscript>")
	ElseIf (Len(viewID) > 0) Then
		openDB
		Set dbError = cartDB.Execute("SELECT id, reported, message FROM errors WHERE viewid = '" & viewID & "'")
		If (dbError.EOF) Then
			Response.Write("Cannot find this incident.")
		Else
			Response.Write("<center><b>Incident Report</b></center>")
			Response.Write("<center><u>Incident ID#</u><br />" & dbError("id") & "</center>")
			Response.Write("<center><u>Reported</u><br />" & dbError("reported") & "</center>")
			Response.Write("<center><u>Report</u></center>")
			Response.Write(Replace(dbError("message"),vbNewline,"<br />"))
			Response.Write("<form method=""post"" action=""" & Request.ServerVariables("URL") & "?view=" & viewID & """>")
			Response.Write("<input type=""hidden"" name=""id"" value=""" & dbError("id") & """>")
			Response.Write("<center><input type=""submit"" name=""action"" value=""Delete""></center>")
			Response.Write("</form>")
		End If
		Set dbError = Nothing
		closeDB
	End If
%>
		