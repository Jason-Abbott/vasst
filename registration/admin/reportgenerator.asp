<% subPageTitle = "Designed Query" %>
<!--#include file="functions.asp"-->
<%
Function runSQL(i,c,o,showDeleted,showUnpaid)
	If (lcase(showDeleted) = "no") Then
		hideDeleted = true
	End If
	If (lcase(showUnpaid) = "no") Then
		hideUnpaid = true
	End If
	
	openDatabase
	on error resume next
	set objRS = server.createObject("ADODB.Recordset")
	
	SQL = "SELECT numCustID, " & Replace(Replace(i,"numCustID, ",""),"numCustID","") & " FROM tblCustomers "

	If Not (c = "") Then
		SQL = SQL & " WHERE (" & c & ") "
		If (hideDeleted) Then
			isDeleted = " AND (isDeleted = false) "
			If (hideUnpaid) Then
				isPaid = " AND (isPaid = true) "
			End If
		End If
	Else
		If (hideDeleted) Then
			isDeleted = " WHERE (isDeleted = false) "
			If (hideUnpaid) Then
				isPaid = " AND (isPaid = true) "
			End If
		Else
			If (hideUnpaid) Then
				isPaid = " WHERE (isPaid = true) "
			End If
		End If
	End If

	SQL = SQL & isDeleted & isPaid
	
	If Not (o = "") Then
		SQL = SQL & " ORDER BY " & o
	End IF

	objRS.open SQL, dbConnection, 2, 3

	If Err <> 0 Then
		Response.Write "<BR><HR><BLOCKQUOTE><B>A Query Error has Occurred:</B><BR>"
		Response.Write "URL: " & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING") & "<BR>"
		Response.Write "Error with statement: " & SQL & "...<BR>"
		Response.Write "From: " & Err.Source & "<br>"
		Response.Write "Error: " & Err.description & "(" & Err.Number & ")<br>"
		Response.Write "</BLOCKQUOTE><HR><BR>"
		Response.Write "<CENTER><B>If you are seeing this, please copy the text above and email it to syntax@sisna.com.</B></CENTER><BR>"
	Else
		Response.Write(tabTo(0) & "<STYLE>")
		Response.Write(tabTo(1) & "TABLE.spreadsheet { Border-Left: 1px solid black; }")
		Response.Write(tabTo(1) & "TR.heading { Background: black; Color: white; Font-Weight: bold; }")
		Response.Write(tabTo(1) & "TD { Font-Size: 9pt; }")
		Response.Write(tabTo(1) & "INPUT.field { Font-Size: 9pt; }")
		Response.Write(tabTo(0) & "</STYLE>")
		Response.Write(tabTo(2) & "<TABLE BORDER=1 CELLPADDING=2 CELLSPACING=1 CLASS=spreadsheet STYLE=""Border-Collapse: collapse; Border-Color: black;"">")
		Response.Write(tabTo(3) & "<TR>")
		Response.Write(tabTo(4) & "<TD COLSPAN=" & objRS.Fields.Count & " CLASS=spreadsheet>")
		Response.Write(tabTo(5) & "<font size=+2><b>Designed Query</b></font><br>")
		Response.Write(tabTo(4) & "<FONT COLOR=#000000>NOTE: You can bookmark this page for viewing in the future.</FONT></TD>")
		Response.Write(tabTo(3) & "</TR>")
'		Response.Write("<TABLE BORDER=1 CELLPADDING=0 CELLSPACING=0>" & vbNewline)
		Response.Write("<TR>" & vbNewline)
		For Each X In objRS.fields
			Response.Write("<TD BGCOLOR=aaaaaa nowrap><font color=333333><b>" & humanize(X.name) & "</b></font></TD>" & vbNewline)
		Next
		Response.Write("</TR>" & vbNewline)
		
		countRecord	= 0
		Do Until objRS.EOF
			countRecord = countRecord + 1
			Response.Write("<TR>" & vbNewline)
			For Each X In objRS.fields
				Response.Write("<TD NOWRAP><a href=""index.asp?page=customer_list.asp&viewCustomer=" & objRS("numCustID") & """ target=""_blank"">" & X & "</a>&nbsp;</TD>" & vbNewline)
			Next
			objRS.MoveNext
			Response.Write("</TR>" & vbNewline)
		Loop
		Response.Write(tabTo(3) & "<TR>")
		Response.Write(tabTo(4) & "<TD COLSPAN=" & objRS.Fields.Count & " CLASS=spreadsheet>")
		Response.Write(tabTo(5) & "<p align=left><font size=2>Total # of Records = " & countRecord & "</font></p>")
		Response.Write(tabTo(4) & "</TD>")
		Response.Write(tabTo(3) & "</TR>")
		Response.Write("</table>" & vbNewline)
	End If
	Set objRS = Nothing
	closeDatabase
End Function
%>


<FORM METHOD=post>
<% 
printHeader 
If Not (dontShowListOrEdit) Then
%>
<INPUT TYPE="hidden" NAME="include" VALUE="<%=Request.QueryString("i")%>">
<INPUT TYPE="hidden" NAME="include" VALUE="<%=Request.QueryString("c")%>">
<INPUT TYPE="hidden" NAME="include" VALUE="<%=Request.QueryString("o")%>">
<INPUT TYPE="hidden" NAME="submitted" VALUE="submitted">
<% If Not (Request.Form("submitted") = "submitted") Then %>
	<SCRIPT>
		this.focus()
		document.forms[0].submit()
	</SCRIPT>
<% Else %>
	<%
		runSQL Request.QueryString("i"),Request.QueryString("c"),Request.QueryString("o"),Request.QueryString("showDeleted"),Request.QueryString("showUnpaid")
	%>
<% End If %>

<%
Else
%>
<br>
<br>
<br>
<center>There are no customers in the database, please register a customer, or add one manually.</center>
<%
End If
printFooter
%>
</FORM>