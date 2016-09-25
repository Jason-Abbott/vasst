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

If (Request.Form("function") = "Save") Then
	openDatabase
	
	For Each frmParam In Request.Form
		If (frmParam <> "function") Then
			Set dbInfo = dbConnection.Execute("SELECT infoID, infoURL FROM tblInfo WHERE effectsSeminar = '" & frmParam & "'")
			If (dbInfo.EOF) Then
				If (Trim(Request.Form(frmParam)) = "") Then
					'Response.Write("Skip " & frmParam & "=" & Request.Form(frmParam) & "<br />")
				Else
					dbConnection.Execute("INSERT INTO tblInfo ( effectsSeminar, infoURL ) VALUES ( '" & frmParam & "', '" & Trim(Request.Form(frmParam)) & "' )")
					'Response.Write("Insert " & frmParam & "=" & Request.Form(frmParam) & "<br />")
				End If
			Else
				'Response.Write("Update " & frmParam & "=" & Request.Form(frmParam) & "<br />")
				dbConnection.Execute("UPDATE tblInfo SET infoURL = '" & Trim(Request.Form(frmParam)) & "' WHERE effectsSeminar = '" & frmParam & "'")
			End If
		End If
	Next
	
	Response.Write("<br>&nbsp;&nbsp;Status: <b><font color=""green"">Updated info URLs.</font></b><br /><br />")
	closeDatabase
End If

openDatabase
Set dbTours = dbConnection.Execute("SELECT strSeminarName FROM tblSeminars GROUP BY strSeminarName ORDER BY strSeminarName")
If (dbTours.EOF) Then
	Response.Write("There are no tours available.")
Else
	Response.Write("<table border=""0"" cellpadding=""2"" cellspacing=""0"">")
	Response.Write("<form method=""post"">")
	Response.Write("<tr>")
	Response.Write("<td>Seminar Name</td>")
	Response.Write("<td>Info URL</td>")
	Response.Write("</tr>")
	Do Until dbTours.EOF
		Set dbInfo = dbConnection.Execute("SELECT infoURL FROM tblInfo WHERE effectsSeminar = '" & dbTours("strSeminarName") & "'")
		If (dbInfo.EOF) Then
			thisURL = ""
		Else
			thisURL = dbInfo("infoURL")
		End If
		
		Response.Write("<tr>")
		Response.Write("<td>" & dbTours("strSeminarName") & "</td>")
		Response.Write("<td><input type=""text"" size=""50"" name=""" & dbTours("strSeminarName") & """ value=""" & thisURL & """></td>")
		Response.Write("</tr>")
		
		dbTours.MoveNext
	Loop
	
	Response.Write("<tr>")
	Response.Write("<td><!--Blank--></td>")
	Response.Write("<td><input type=""submit"" name=""function"" value=""Save""></td>")
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write("<td colspan=""2"">A seminar must be added for it to be listed here.<br /><br />Leave the URL blank if you do not have a URL yet.<br /><br />You can enter a url like http://www.vasst.com/infopage.htm, or /infopage.htm, which means http://www.vasst.com/infopage.htm.</td>")
	Response.Write("</tr>")
	
	Response.Write("</form>")
	Response.Write("</table>")
End If
closeDatabase

printBoxClose
printFooter 
%>