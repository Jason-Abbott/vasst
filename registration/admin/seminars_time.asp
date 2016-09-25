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
	Set dbTime = Server.CreateObject("ADODB.RecordSet")
	openDatabase
	
	For Each frmParam In Request.Form
		If (Left(frmParam, 8) = "seminar_") Then
			strID = Replace(frmParam, "seminar_", "")
			strSeminar = Request.Form(frmParam)
			thisStart = Trim(Request.Form("start_" & strID))
			thisEnd = Trim(Request.Form("end_" & strID))

'			Response.Write("Seminar: '" & strSeminar & "' Start: #" & thisStart & "# End: #" & thisEnd & "#<br />")
			Set dbInfo = dbConnection.Execute("SELECT id FROM tblSeminarTimes WHERE effectsSeminar = '" & strSeminar & "'")
			bDoIt = false
			If (dbInfo.EOF) Then
				
				If (thisStart = "") Or (thisEnd = "") Then
'					Response.Write("Skip " & frmParam & "=" & Request.Form(frmParam) & "<br />")
				Else
					dbTime.Open "SELECT * FROM tblSeminarTimes", dbConnection, 2, 3
					dbTime.AddNew
					bDoIt = True
'					dbConnection.Execute("INSERT INTO tblSeminarTimes ( effectsSeminar, startTime, endTime ) VALUES ( '" & strSeminar & "', #" & thisStart & "#, #" & thisEnd & "# )")
'					Response.Write("Insert " & frmParam & "=" & Request.Form(frmParam) & "<br />")
				End If
			Else
				dbTime.Open "SELECT * FROM tblSeminarTimes WHERE id = " & dbInfo("id") & "", dbConnection, 2, 3
				bDoIt = True
'				Response.Write("Update " & frmParam & "=" & Request.Form(frmParam) & "<br />")
'				dbConnection.Execute("UPDATE tblSeminarTimes SET startTime = #" & thisStart & "#, endTime = #" & thisEnd & "# WHERE id = '" & dbInfo("id") & "'")
			End If
			If (bDoIt) Then
				dbTime.Fields("effectsSeminar").Value = strSeminar
				dbTime.Fields("startTime").Value = thisStart
				dbTime.Fields("endTime").Value = thisEnd
				dbTime.Update
				dbTime.Close
			End If
		End If
	Next
	
	Response.Write("<br>&nbsp;&nbsp;Status: <b><font color=""green"">Updated info URLs.</font></b><br /><br />")
	closeDatabase
	Set dbTime = Nothing
End If

openDatabase
Set dbTours = dbConnection.Execute("SELECT strSeminarName FROM tblSeminars GROUP BY strSeminarName ORDER BY strSeminarName")
If (dbTours.EOF) Then
	Response.Write("There are no tours available.")
Else
	Response.Write("<table border=""0"" cellpadding=""2"" cellspacing=""0"">")
	Response.Write("<form method=""post"">")
	Response.Write("<tr>")
	Response.Write("<td><b>Seminar Name</b></td>")
	Response.Write("<td><b>Start Time</b></td>")
	Response.Write("<td><b>End Time</b></td>")
	Response.Write("</tr>")
	iID = 0
	Do Until dbTours.EOF
		Set dbInfo = dbConnection.Execute("SELECT id, startTime, endTime FROM tblSeminarTimes WHERE effectsSeminar = '" & dbTours("strSeminarName") & "'")
		If (dbInfo.EOF) Then
			thisStart = ""
			thisEnd = ""
		Else
			thisStart = dbInfo("startTime")
			thisEnd = dbInfo("endTime")
		End If
		
		Response.Write("<tr>")
		Response.Write("<td><input type=""hidden"" name=""seminar_" & iID & """ value=""" & dbTours("strSeminarName") & """ />" & dbTours("strSeminarName") & "</td>")
		Response.Write("<td><input type=""text"" size=""20"" name=""start_" & iID & """ value=""" & thisStart & """></td>")
		Response.Write("<td><input type=""text"" size=""20"" name=""end_" & iID & """ value=""" & thisEnd & """></td>")
		Response.Write("</tr>")

		dbTours.MoveNext
		iID = iID + 1
	Loop
	
	Response.Write("<tr>")
	Response.Write("<td><!--Blank--></td>")
	Response.Write("<td><input type=""submit"" name=""function"" value=""Save""></td>")
	Response.Write("<td><!--Blank--></td>")
	Response.Write("</tr>")
	
	Response.Write("</form>")
	Response.Write("</table>")
End If
closeDatabase

printBoxClose
printFooter 
%>