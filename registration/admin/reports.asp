<% subPageTitle = "Customer Spreadsheet" %>
<!--#include file="functions.asp"-->
<%
Function printCustomerFields
	openDatabase
	set objRS=server.createObject("ADODB.Recordset")
	SQL = "SELECT * FROM tblCustomers WHERE false"
	objRS.open SQL, dbConnection, 2, 3

	For X = 0 to objRS.fields.count - 1
		For Each item In split(Request.Form("fieldSelect"),", ")
			If (item = objRS.fields(X).name) Then
				selected = "SELECTED"
				Exit For
			Else
				selected = ""
			End If
		Next
		Response.Write("<option value=""" & objRS.fields(X).name & """ " & selected & ">" & humanize(objRS.fields(X).name) & "</option>"& vbNewline)
	Next

	objRS.close
	Set objRS = Nothing
	closeDatabase
End Function

Function printNumberOfFields
	openDatabase
	set objRS=server.createObject("ADODB.Recordset")
	SQL = "SELECT * FROM tblCustomers WHERE false"
	objRS.open SQL, dbConnection, 2, 3

	Response.Write(objRS.fields.count)

	objRS.close
	Set objRS = Nothing
	closeDatabase
End Function

Function ArrayAddLast(theArray, newData)
	If isArray(theArray) Then
		ReDim Preserve theArray(Ubound(theArray)+1)
		theArray(Ubound(theArray)) = newData
	End If
	ArrayAddLast = theArray
End Function 

Function ArrayDeleteLast(theArray)
	If isArray(theArray) Then
		ReDim Preserve theArray(Ubound(theArray)-1)
	End If
	ArrayDeleteLast = theArray
End Function

Function ArrayGetLast(theArray)
	If isArray(theArray) Then
		ArrayGetLast = theArray(Ubound(theArray))
	Else
		ArrayGetLast = theArray
	End If
End Function

Function ArrayClear(theArray)
End Function

Function ArraySwitch(theArray, fromPos, toPos)
	If isArray(theArray) Then
		Dim tempVal
		temp = theArray(toPos) 
		theArray(toPos) = theArray(fromPos)
		theArray(fromPos) = temp
	End If
	ArraySwitch = theArray
End Function

Function CombineArray(theArray, newData)
	Dim tempArray, addTo
	
	If isArray(newData) Then
		addTo = Ubound(newData)+1
	Else
		addTo = 1
	End If
	tempArray = theArray
	ReDim Preserve tempArray(Ubound(theArray)+addTo)
	For pushArray = Ubound(theArray)+1 To Ubound(tempArray)
		If isArray(newData) Then
			tempArray(pushArray) = newData(pushArray-Ubound(theArray)-1)
		Else
			tempArray(pushArray) = newData
		End If
	Next	
	CombineArray = tempArray
End Function

Dim SQLStatement, hiddenForm, includeOptions, conditionOptions, orderOptions
includeSaved = Request.Form("includeSaved")
conditionSaved = Request.Form("conditionSaved")
orderSaved = Request.Form("orderSaved")

'fieldString = Request.Form("fieldSelect")

'SQLStatement = "SELECT " & fieldString & " FROM tblCustomers "

If (Request.Form("function") = "Add") AND (Request.Form("whatOption") = "include") AND Not (Request.Form("fieldSelect") = "") Then
	If (Request.Form("includeSaved") = "") Then
		includeToSave = Request.Form("fieldSelect")
	Else
		newIncludeTemp = ""
		For Each item In split(Request.Form("fieldSelect"),", ")
			For Each item2 In split(Request.Form("includeSaved"),", ")
				If (item = item2) Then
					doNotInclude = true
					Exit For
				Else
					doNotInclude = false
				End If
			Next
			If Not (doNotInclude) Then
				newIncludeTemp = newIncludeTemp & ", " & item
			End If
		Next
		includeToSave = Request.Form("includeSaved") & newIncludeTemp
	End If
ElseIf (Request.Form("function") = "Remove") AND Not (Request.Form("queryDesign") = "") Then
	newIncludeTemp = Request.Form("includeSaved")
	For Each item In split(Request.Form("queryDesign"),", ")
		For Each item2 In split(newIncludeTemp,", ")
			If (item = item2) Then
				newIncludeTemp = Replace(newIncludeTemp,item & ", ","")
				newIncludeTemp = Replace(newIncludeTemp,", " & item,"")
				newIncludeTemp = Replace(newIncludeTemp,item,"")
				Exit For
			End If
		Next
	Next
	includeToSave = newIncludeTemp
Else
	includeToSave = Request.Form("includeSaved")
End If	

If (Request.Form("function") = "Add") AND (Request.Form("whatOption") = "condition") AND Not (Request.Form("fieldSelect") = "") Then
	If (InStr(Request.Form("conditionStatement"),"*") > 0) Then
		conditionToSave = Request.Form("fieldSelect") & " LIKE '" & Replace(Request.Form("conditionStatement"),"*","%") & "'"
	Else
		conditionToSave = Request.Form("fieldSelect") & " = '" & Request.Form("conditionStatement") & "'"
	End If
'	conditionToSave = Replace(Request.Form("fieldSelect"),", ", Request.Form("conditionOperator") & " '" & Request.Form("conditionStatement") & "', ")
'	conditionToSave = conditionToSave & " " & Request.Form("conditionOperator") & " '" & Request.Form("conditionStatement") & "'"
'	If Not (Request.Form("conditionSaved") = "") Then
'		conditionToSave = Request.Form("conditionSaved") & ", " & conditionToSave
'	End If
ElseIf (Request.Form("function") = "Remove") AND Not (Request.Form("queryDesign") = "") Then
	newConditionTemp = Request.Form("conditionSaved")
	For Each item In split(Request.Form("queryDesign"),", ")
		For Each item2 In split(newConditionTemp,", ")
			If (item = item2) Then
				conditionToSave = ""
				newConditionTemp = ""
				removedCondition = true
'				newConditionTemp = Replace(newConditionTemp,item & ", ","")
'				newConditionTemp = Replace(newConditionTemp,", " & item,"")
'				newConditionTemp = Replace(newConditionTemp,item,"")
				Exit For
			End If
		Next
	Next
'	conditionToSave = newConditionTemp
Else
	conditionToSave = Request.Form("conditionSaved")
End If	

If (Request.Form("function") = "Add") AND (Request.Form("whatOption") = "order") AND Not (Request.Form("fieldSelect") = "") Then
	orderToSave = Replace(Request.Form("fieldSelect"),", ", " " & Request.Form("orderDirection") & ", ")
	orderToSave = orderToSave & " " & Request.Form("orderDirection")
	If Not (Request.Form("orderSaved") = "") Then
		orderToSave = Request.Form("orderSaved") & ", " & orderToSave
	End If
ElseIf (Request.Form("function") = "Remove") AND Not (Request.Form("queryDesign") = "") Then
	newOrderTemp = Request.Form("orderSaved")
	For Each item In split(Request.Form("queryDesign"),", ")
		For Each item2 In split(newOrderTemp,", ")
			If (item = item2) Then
				newOrderTemp = Replace(newOrderTemp,item & ", ","")
				newOrderTemp = Replace(newOrderTemp,", " & item,"")
				newOrderTemp = Replace(newOrderTemp,item,"")
				Exit For
			End If
		Next
	Next
	orderToSave = newOrderTemp
Else
	orderToSave = Request.Form("orderSaved")
End If	

hiddenForm = hiddenForm & "<INPUT TYPE=""hidden"" NAME=""includeSaved"" VALUE=""" & includeToSave & """>"
If (Ubound(split(includeToSave,", ")) = -1) Then
	includeToSave = "*"
Else
	For Each item In split(includeToSave,", ")
		includeOptions = includeOptions & "<OPTION VALUE=""" & item & """ CLASS=""included"">&nbsp;&nbsp;&nbsp;" & humanize(item)& "</OPTION>" & vbNewline
	Next
End If

hiddenForm = hiddenForm & "<INPUT TYPE=""hidden"" NAME=""conditionSaved"" VALUE=""" & conditionToSave & """>"
For Each item In split(conditionToSave,", ")
	conditionOptions = conditionOptions & "<OPTION VALUE=""" & item & """ CLASS=""conditions"">&nbsp;&nbsp;&nbsp;" & humanize(item)& "</OPTION>" & vbNewline
Next

hiddenForm = hiddenForm & "<INPUT TYPE=""hidden"" NAME=""orderSaved"" VALUE=""" & orderToSave & """>"
For Each item In split(orderToSave,", ")
	orderOptions = orderOptions & "<OPTION VALUE=""" & item & """ CLASS=""order"">&nbsp;&nbsp;&nbsp;" & humanize(item) & "</OPTION>" & vbNewline
Next

SQLStatement = "SELECT "
If (Ubound(split(includeToSave,", ")) = -1) Then
	SQLStatement = SQLStatement & "<SPAN CLASS=included>*</SPAN>"
Else
	SQLStatement = SQLStatement & "<SPAN CLASS=included>" & includeToSave & "</SPAN>"
End If

SQLStatement = SQLStatement & " <BR>FROM tblCustomers "
'If (Ubound(split(conditionToSave,", ")) > -1) Then
	SQLStatement = SQLStatement & " <BR>WHERE <SPAN CLASS=conditions>" & conditionToSave & "</SPAN>"
'End If

If (Ubound(split(orderToSave,", ")) > -1) Then
	SQLStatement = SQLStatement & " <BR>ORDER BY <SPAN CLASS=order>" & orderToSave & "</SPAN>"
End If
%>

<FORM METHOD=post>
<%=hiddenForm%>
<% 
printHeader 
If Not (dontShowListOrEdit) Then

'If (Request.Form("editButton") = "") Then
'	printCustomers "view"
'Else
'	Response.Redirect("customer_editspreadsheet.asp")
'End If
%>

<head>
<style>
 .included { background: #CCCCFF; }
 .conditions { background: #CCFFCC; }
 .order { background: #FFCCCC; }
 .webdings { Font-Family: webdings; }
</style>
<meta name="Microsoft Theme" content="modified-powerplugs-web-templates-art3dblue 011, default">
<meta name="Microsoft Border" content="none, default">
</head>
<script>
 function disableMultiselect() {
  document.forms[0].fieldSelect.multiple = false;
 }
 
 function enableMultiselect() {
  document.forms[0].fieldSelect.multiple = true;
 }
</script>

<% If (Request.Form("function") = "Open Query") Then %>
<SCRIPT>
//x = window.open("reportgenerator.asp?i=<%=Server.URLEncode(includeToSave)%>&c=<%=Server.URLEncode(conditionToSave)%>&o=<%=Server.URLEncode(orderToSave)%>","Designed Query","resizable=yes,scrollbars=yes,menubar=yes,toolbar=no,status=no,location=no");
window.open("reportgenerator.asp?i=<%=Server.URLEncode(includeToSave)%>&c=<%=Server.URLEncode(conditionToSave)%>&o=<%=Server.URLEncode(orderToSave)%>&showDeleted=<%=Request("showDeleted")%>&showUnpaid=<%=Request("showUnpaid")%>");
//  x.focus();
</SCRIPT>
<% End If %>

</script>
<!--mstheme--></font><table border=0 cellpadding=0 cellspacing=0 width="100%">
	<tr>
		<td colspan=3 align="center"><!--mstheme--><font face="Arial, Arial, Helvetica">
			<big><big><b><u>Query Designer<u></b></big></big>
		<!--mstheme--></font></td>
	</tr>
	<tr>
		<td valign=top width="1%"><!--mstheme--><font face="Arial, Arial, Helvetica">
			<!--mstheme--></font><table border=0 cellpadding=0 cellspacing=0>
				<tr>
					<td align="center"><!--mstheme--><font face="Arial, Arial, Helvetica">
						<u><b>Field Select</b></u><br>
						<select name="fieldSelect" size="<% printNumberOfFields %>" multiple>
							<% printCustomerFields %>
						</select>
					<!--mstheme--></font></td>
				</tr>
			</table><!--mstheme--><font face="Arial, Arial, Helvetica">
		<!--mstheme--></font></td>
		<td valign=top><!--mstheme--><font face="Arial, Arial, Helvetica">
			<br>
			<!--mstheme--></font><table border=0 cellpadding=0 cellspacing=0 height="100%" width="100%">
				<tr>
					<td align=right valign="top" nowrap width="1%" height="1%"><!--mstheme--><font face="Arial, Arial, Helvetica">Include:<!--mstheme--></font></td>
					<td valign="top" height="1%"><!--mstheme--><font face="Arial, Arial, Helvetica"><input class="normal" type="radio" name="whatOption" value="include" <% If (Request.Form("whatOption") = "include") OR (Request.Form("whatOption") = "") Then Response.Write("checked") End If %> onClick="enableMultiselect()"><!--mstheme--></font></td>
					<td><!--mstheme--><font face="Arial, Arial, Helvetica">field<!--mstheme--></font></td>
				</tr>
				<% If Not ((Request.Form("conditionSaved") <> "") Or (conditionToSave <> "")) Or (removedCondition) Then %>
				<tr>
					<td align=right valign="top" nowrap width="1%" height="1%"><!--mstheme--><font face="Arial, Arial, Helvetica">Condition:<!--mstheme--></font></td>
					<td valign="top" height="1%"><!--mstheme--><font face="Arial, Arial, Helvetica"><input class="normal" type="radio" name="whatOption" value="condition" <% If (Request.Form("whatOption") = "condition") Then Response.Write("checked") End If %> onClick="disableMultiselect()"><!--mstheme--></font></td>
					<td><!--mstheme--><font face="Arial, Arial, Helvetica">
						field = <input name="conditionStatement" size="30"> <small><small> * = wildcard</small></small><br>
					<!--mstheme--></font></td>
				</tr>
				<% End If %>
				<tr>
					<td align=right valign="top" nowrap width="1%" height="1%"><!--mstheme--><font face="Arial, Arial, Helvetica">Order By:<!--mstheme--></font></td>
					<td valign="top" height="1%"><!--mstheme--><font face="Arial, Arial, Helvetica"><input class="normal" type="radio" name="whatOption" value="order" <% If (Request.Form("whatOption") = "order") Then Response.Write("checked") End If %> onClick="enableMultiselect()"><!--mstheme--></font></td>
					<td><!--mstheme--><font face="Arial, Arial, Helvetica">field
						<select name="orderDirection" size="1">
							<option value="ASC" <% If (Request.Form("orderDirection") = "ASC") Then Response.Write("selected") End If %>>ASC</option>
							<option value="DESC" <% If (Request.Form("orderDirection") = "DESC") Then Response.Write("selected") End If %>>DESC</option>
						</select>
					<!--mstheme--></font></td>
				</tr>
				<tr>
					<td align=right valign="top" nowrap width="1%" height="1%"><!--mstheme--><font face="Arial, Arial, Helvetica">Options:<!--mstheme--></font></td>
					<td valign="top" height="1%"><!--mstheme--><font face="Arial, Arial, Helvetica">&nbsp;<!--mstheme--></font></td>
					<td><!--mstheme--><font face="Arial, Arial, Helvetica">
						Show Deleted? <select size="1" name="showDeleted">
							<option value="no" <% If Request.Form("showDeleted") <> "no" Then Response.Write("") Else Response.Write(" selected") End If %>>no</option>
							<option value="yes" <% If Request.Form("showDeleted") = "yes" Then Response.Write(" selected") End If %>>yes</option>
						</select> <small><small>no - hides deleted</small></small>
						<br>
						Show Un-Paid? <select size="1" name="showUnpaid">
							<option value="no" <% If Request.Form("showUnpaid") <> "no" Then Response.Write("") Else Response.Write(" selected") End If %>>no</option>
							<option value="yes" <% If Request.Form("showUnpaid") = "yes" Then Response.Write(" selected") End If %>>yes</option>
						</select> <small><small>no - hides un-paid</small></small>
						
					<!--mstheme--></font></td>
				</tr>
				<tr>
					<td valign="bottom" align="right" colspan=3 nowrap><!--mstheme--><font face="Arial, Arial, Helvetica">
						<input type="submit" name="function" value="Clear Query" onClick="top.location.href='index.asp?page=reports.asp'">
						<input type="submit" name="function" value="Open Query">&nbsp;&nbsp;&nbsp;
						<input type="submit" name="function" value="Add">
						<input type="submit" name="function" value="Remove"><br>
						<br><br>
					<!--mstheme--></font></td>
				</tr>
				<tr>
					<td valign="top" align="right" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica">Stats:<!--mstheme--></font></td>
					<td><!--mstheme--><font face="Arial, Arial, Helvetica"><!--mstheme--></font></td>
					<td valign="top"><!--mstheme--><font face="Arial, Arial, Helvetica">
						# of Fields Included: <%=Ubound(split(includeToSave,", "))+1%><br>
						# of Conditions: <%=Ubound(split(conditionToSave,", "))+1%><br>
						# of Order Statements: <%=Ubound(split(orderToSave,", "))+1%><br><br>
					<!--mstheme--></font></td>
				</tr>
				<tr>
					<td valign="top" align="right" nowrap><!--mstheme--><font face="Arial, Arial, Helvetica">SQL:<!--mstheme--></font></td>
					<td><!--mstheme--><font face="Arial, Arial, Helvetica"><!--mstheme--></font></td>
					<td valign="top"><!--mstheme--><font face="Arial, Arial, Helvetica">
						<%=SQLStatement%>
					<!--mstheme--></font></td>
				</tr>
			</table><!--mstheme--><font face="Arial, Arial, Helvetica">
		<!--mstheme--></font></td>
		<td valign=top width="1%"><!--mstheme--><font face="Arial, Arial, Helvetica">
			<!--mstheme--></font><table border=0 cellpadding=0 cellspacing=0>
				<tr>
					<td width=2% rowspan=4 valign=top align=center><!--mstheme--><font face="Arial, Arial, Helvetica">
						<u><b>Query&nbsp;Design</b></u><br>
						<select name="queryDesign" size="<% printNumberOfFields %>" multiple>
							<option value="doNothing" class="included">Included:</option>
							<%=includeOptions%>
							<option value="doNothing"></option>
							<option value="doNothing" class="conditions">Conditions:</option>
							<%=conditionOptions%>
							<option value="doNothing"></option>
							<option value="doNothing" class="order">Order:</option>
							<%=orderOptions%>
							<option value="doNothing"></option>
						</select>
					<!--mstheme--></font></td>
				</tr>
			</table><!--mstheme--><font face="Arial, Arial, Helvetica">
		<!--mstheme--></font></td>
	</tr>
</table><!--mstheme--><font face="Arial, Arial, Helvetica">
<% Else %>
<br>
<br>
<br>
<center>There are no customers in the database, please register a customer, or add one manually.</center>
<%
End If
printFooter
%>
</FORM>