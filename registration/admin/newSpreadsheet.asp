<% authNotNeeded = True %>
<!--#include file="functions.asp"-->
<form method="post">
<%
openDatabase
If (Request.Form("editButton") = "Edit") Then
	spreadsheetMode = "edit"
Else
	spreadsheetMode = "view"
End If

		If ((spreadsheetMode = "view") Or (spreadsheetMode = "edit") Or (spreadsheetMode = "update")) Then
			Set objRS = Server.CreateObject("ADODB.Recordset")
			If (Request.QueryString("searchFor") = "") Or (Request.QueryString("searchFor") = "All") Then
				'Show everything in the tblCustomers
				strSQL = "SELECT * FROM tblCustomers"
			Else
				strSQL = "SELECT * FROM tblCustomers WHERE strSeminarName = '" & Request.QueryString("searchFor") & "'"
			End If
			objRS.open strSQL, dbConnection, 2, 3

'			dbHeadings = array(	  "count",	"numCustID",	"dateSignup",	"strFirstName",	"strLastName",	"strCompanyName",	"strTitle",	"strAddress1",	"strAddress2",	"strCity",	"strState",	"strZip",	"numPhone",	"strEmail",	"strSeminarName",	"dateSeminarDate",	"strSeminarCity",	"currencyCost",	"optPayment",	"isDiscount",	"isPaid",	"isDeleted",	"optStatus",	"strRetail1",	"strRetail2",	"strRetail3",	"strComments1",	"strComments2",	"strComments3")
''			recordSize = array(		4,			4,					20,				15,				15,				20,					10,		20,			20,			15,		10,		10,		12,		20,		20,			20,			20,			8,			10,		6,		6,		6,		15,		20,		20,		20,		20,			20,			20)					
'			recordHeadings = array(	"#",	"ID",	"Date of Signup",	"First Name",	"Last Name",	"Company Name",	"Title",	"Address1",	"Address2",	"City",	"State",	"Zip",	"Phone Number",	"Email Address",	"Seminar Name",	"Seminar Date",	"Seminar City",	"Seminar Cost",	"Payment Type",	"Discount?",	"Paid?",	"Deleted?",	"Status",	"Retail1",	"Retail2",	"Retail3",	"Comments1",	"Comments2",	"Comments3")
			recordSize = array(		4,		4,		20,					15,				15,				20,				10,			20,		20,		15,		10,		10,	12,			20,			20,			20,			20,			8,			10,			6,		6,		6,		15,		20,		20,		20,		20,		20,		20)
	
			If (spreadsheetMode = "update") And Not (countRecord = 0) Then
				Dim records
				numberOfRecords = 0
				numberOfFields = 0

				For Each objForm In Request.Form
					If (InStr(objForm,"field_") > 0) Then
						splitter = Split(objForm,"_")
						id = Cint(splitter(1))
						field = Cint(splitter(2))
						If (id > Cint(numberOfRecords)) Then
							numberOfRecords = id
						End If
						If (field > Cint(numberOfFields)) Then
							numberOfFields = field
						End If
					End If
				Next
				ReDim records(numberOfRecords, numberOfFields)
	
				For Each objForm In Request.Form
					tmp = Split(objForm,"=")
					If (InStr(objForm,"field_") > 0) Then
'						Response.Write(objForm & "=" & Request.Form(objForm) & "<BR>")
						splitter = Split(objForm,"_")
						id = Cint(splitter(1))
						field = Cint(splitter(2))
'						Response.Write("id:" & id & " field:" & field & "(" & recordHeadings(field+1) & ") =" & Request.Form(objForm) & "<BR>")
						records(id,field) = Request.Form(objForm)
					End If
				Next

				objRS.Move 0
				For x = 1 to numberOfRecords
					For y = 0 to objRS.Fields.Count-1
'						Response.Write("X:" & x & " Y:" & y & "(" & recordHeadings(y+1) & ") --- " & objRS.Fields(y) & "=" & records(x,y))
'						Response.Write("<BR>")
						If Not (y = 0) Then
							objRS.Fields(y) = records(x,y)
						End If
					Next
					objRS.MoveNext
				Next
			Else
				Pos = 0
				Response.Write(tabTo(0) & "<STYLE>")
				Response.Write(tabTo(1) & "TABLE.spreadsheet { Border-Left: 1px solid black; }")
				Response.Write(tabTo(1) & "TR.heading { Background: black; Color: white; Font-Weight: bold; }")
				Response.Write(tabTo(1) & "TD { Font-Size: 9pt; }")
				Response.Write(tabTo(1) & "INPUT.field { Font-Size: 9pt; }")
				Response.Write(tabTo(0) & "</STYLE>")
				Response.Write(tabTo(2) & "<TABLE BORDER=1 CELLPADDING=2 CELLSPACING=0 CLASS=spreadsheet STYLE=""Border-Collapse: collapse; Border-Color: black;"">")
				Response.Write(tabTo(3) & "<TR>")
				Response.Write(tabTo(4) & "<TD COLSPAN=" & objRS.Fields.Count+1 & " CLASS=spreadsheet>")
				Response.Write(tabTo(5) & "<p align=left><font size=+2><b>Customer Database</b></font></p>")
				Response.Write(tabTo(4) & "</TD>")
				Response.Write(tabTo(3) & "</TR>")
				While Not objRS.EOF 
					If ((Pos Mod 10) = 0) Then
						Response.Write("<TR>" & vbNewline)
						Response.Write("<TD BGCOLOR=aaaaaa nowrap><font color=333333><b>&nbsp#&nbsp;</b></font></TD>")
						For Each X In objRS.fields
							Response.Write("<TD BGCOLOR=aaaaaa nowrap><font color=333333><b>" & humanize(X.name) & "</b></font></TD>" & vbNewline)
						Next
						Response.Write("</TR>" & vbNewline)
					End If
					Pos = Pos + 1
					Response.Write(tabTo(3) & "<TR>")
					Response.Write(tabTo(4) & "<TD NOWRAP CLASS=spreadsheet>" & Pos & "</TD>")
					for a = 0 to objRS.Fields.Count-1
						If (spreadsheetMode = "view") or (a = 0) Then
							Response.Write(tabTo(4) & "<TD NOWRAP>" & objRS.Fields(a) & "&nbsp;</TD>")
						ElseIf (spreadsheetMode = "edit") Then
							Response.Write(tabTo(4) & "<TD NOWRAP><INPUT CLASS=field NAME=field_" & Pos & "_" & a & " VALUE='" & objRS.Fields(a) & "' SIZE=" & recordSize(a+1) & "></TD>")
						End If
					next
					Response.Write(tabTo(3) & "</TR>")
					objRS.MoveNext
				Wend
				Response.Write(tabTo(3) & "<TR>")
				Response.Write(tabTo(4) & "<TD COLSPAN=" & objRS.Fields.Count+1 & " CLASS=spreadsheet>")
				Response.Write(tabTo(5) & "<p align=left><font size=2>Total # of Records = " & countRecord & "</font></p>")
				Response.Write(tabTo(4) & "</TD>")
				Response.Write(tabTo(3) & "</TR>")
				Response.Write(tabTo(3) & "<TR>")
				Response.Write(tabTo(4) & "<TD COLSPAN=" & objRS.Fields.Count+1 & " CLASS=spreadsheet>")
				Response.Write(tabTo(5) & "<TABLE BORDER=0 WIDTH='100%' ALIGN=center>")
				Response.Write(tabTo(6) & "<TR>")
				For looptimes = 1 to 5
					Response.Write(tabTo(7) & "<TD ALIGN=center NOWRAP>")
					If (spreadsheetMode = "edit") Then
						Response.Write(tabTo(8) & "<INPUT TYPE=submit NAME='saveButton' VALUE='Save' STYLE='Width:49%'><INPUT TYPE=reset VALUE='Reset' STYLE='Width:49%'>")
					ElseIf (spreadsheetMode = "view") Then
						'If (getUserInfo(getUserID,"accesslevel") = "Admin") Then
							Response.Write(tabTo(8) & "<INPUT TYPE=submit NAME='editButton' VALUE='Edit' STYLE='Width:98%'>")
						'End If
					End If
					Response.Write(tabTo(7) & "</TD>")
				Next
				Response.Write(tabTo(6) & "</TR>")
				Response.Write(tabTo(5) & "</TABLE>")
				Response.Write(tabTo(4) & "</TD>")
				Response.Write(tabTo(3) & "</TR>")
				Response.Write(tabTo(2) & "</TABLE>")
				objRS.close
			End If
		End If
closeDatabase
%>
</form>
<%
splitForm
%>