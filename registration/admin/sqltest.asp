<!--#include file="functions.asp"-->
<% 
If Not (lcase(getUserInfo(getUserID,"username")) = "jwalker") Then 
	Response.Write "<CENTER><B>You do not have access to this section of the CMS, your access is not allowed, this tool can cause serious damage to the CMS and is only to be used by the developer in case of emergency.  <BR><BR>Contact James Walker by email syntax@sisna.com for questions.</B></CENTER>"
	Response.End
End If
%>
<%
	Dim SQLConsoleotherReturn, SQLConsoleselectReturn, SQLConsoledatabaseOutput
	Dim SQLConsoledbRS
	Set SQLConsoledbRS = Server.CreateObject("ADODB.RecordSet")
	
	Function runOtherSQL(runSQL)
				On Error Resume Next
		openDatabase
		If (Len(runSQL) > 0) Then
			dbConnection.Execute(runSQL)
			If Err <> 0 Then
				SQLConsoleotherReturn = ("<font color=""red"">Failed</font><br />" & Err.Description)
				Err.Clear
			Else
				SQLConsoleotherReturn = ("<font color=""green"">Success</font><br />")
			End If
		End If
		closeDatabase
	End Function
	
	Function runSelectSQL(runSQL)
				On Error Resume Next
		openDatabase
		dbConnection.CommandTimeout = 0
		If (Len(runSQL) > 0) Then
			SQLConsoledbRS.Open runSQL, dbConnection, 2, 3
			If (Err <> 0) Then
				SQLConsoleselectReturn = ("<font color=""red"">Failed</font><br />" & Err.Description)
				Err.Clear
			Else
				SQLConsoleselectReturn = ("<font color=""green"">Success</font><br />")
				SQLConsoledatabaseOutput = ("<center><b><big>Database Query</big></b></center>")
				SQLConsoledatabaseOutput = SQLConsoledatabaseOutput & ("<table border=""1"" cellpadding=""2"" cellspacing=""0"" align=""center"">")
				SQLConsoledatabaseOutput = SQLConsoledatabaseOutput & ("<tr bgcolor=""black"">")
				For field = 0 To SQLConsoledbRS.Fields.Count-1
					SQLConsoledatabaseOutput = SQLConsoledatabaseOutput & ("<td><font color=""white"" size=""-2"">" & SQLConsoledbRS.Fields(field).Name & " </font></td>")
				Next
				SQLConsoledatabaseOutput = SQLConsoledatabaseOutput & ("</tr>")
				Do Until SQLConsoledbRS.EOF
					SQLConsoledatabaseOutput = SQLConsoledatabaseOutput & ("<tr>")
					For field = 0 To SQLConsoledbRS.Fields.Count-1
						SQLConsoledatabaseOutput = SQLConsoledatabaseOutput & ("<td><font size=""-2"">")
						If (isNull(SQLConsoledbRS.Fields(field))) Then
							SQLConsoledatabaseOutput = SQLConsoledatabaseOutput & ("&nbsp;")
						Else
							SQLConsoledatabaseOutput = SQLConsoledatabaseOutput & (Replace(SQLConsoledbRS.Fields(field),"&","&amp;"))
						End If
						SQLConsoledatabaseOutput = SQLConsoledatabaseOutput & ("</font></td>")
					Next
					SQLConsoledatabaseOutput = SQLConsoledatabaseOutput & ("</tr>")
					SQLConsoledbRS.MoveNext
				Loop
				SQLConsoledatabaseOutput = SQLConsoledatabaseOutput & ("</table>")
			End If
		End If
		closeDatabase
	End Function
	
	'adFldCacheDeferred 0x1000 Provider caches values and reads from cache 
	'adFldFixed 0x10 Fixed-length data 
	'adFldIsChapter 0x2000 Chapter value with specified child recordset 
	'adFldIsCollection 0x40000 Collection of resources 
	'adFldIsDefaultStream 0x20000 Contains default stream 
	'adFldIsNullable 0x20 Accepts null values 
	'adFldIsRowURL 0x10000 Contains URL to resource in data source 
	'adFldKeyColumn 0x8000 Primary key or part of primary key 
	'adFldLong 0x80 Long binary field and can use AppendChunk and GetChunk methods 
	'adFldMayBeNull 0x40 Can read null values 
	'adFldMayDefer 0x2 Values are not retrieved with whole record 
	'adFldNegativeScale 0x4000 Can support negative scale values 
	'adFldRowID 0x100 Contains a row identifier used only to ID the row 
	'adFldRowVersion 0x200 Uses time/date to track updates 
	'adFldUnknownUpdatable 0x8 Provider cannot determine if you can write to field 
	'adFldUnspecified -1 Does not specify attributes 
	'adFldUpdatable 0x4 Can write to field 
		
	Function getDataType(intDataType)
		Select Case (intDataType)
		Case 2
			getDataType = "smallint"
		Case 3
			getDataType = "int"
	'	Case 4 
	'		getDataType = "real"
		Case 4 
			getDataType = "float"
		Case 6 
			getDataType = "money"
	'	Case 6 
	'		getDataType = "smallmoney"
		Case 11 
			getDataType = "bit"
		Case 17 
			getDataType = "tinyint"
		Case 72 
			getDataType = "uniqueidentifier"
		Case 128 
			getDataType = "binary"
	'	Case 128 
	'		getDataType = "timestamp"
		Case 129 
			getDataType = "char"
		Case 130 
			getDataType = "nchar"
		Case 131 
			getDataType = "decimal"
	'	Case 131 
	'		getDataType = "numeric"
		Case 135 
			getDataType = "datetime"
	'	Case 135 
	'		getDataType = "smalldatetime"
		Case 200 
			getDataType = "varchar"
		Case 201 
			getDataType = "text"
		Case 202 
			getDataType = "nvarchar"
	'	Case 202 
	'		getDataType = "sysname"
		Case 203 
			getDataType = "ntext"
		Case 204 
			getDataType = "varbinary"
		Case 205 
			getDataType = "image"
		Case Else
			getDataType = "unknown"
		End Select
	End Function
	
	Function printAvailableTables
	'	perRow = 5
		openDatabase
		Response.Write("<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"">")
		Response.Write("<tr><td align=""center"" colspan=""2""><b><big>Database Tables</big></b></td></tr>")
		Set dbTest = dbConnection.Execute("SELECT * FROM MSysObjects")
		If (dbTest.EOF) Then
			dbType = "SQL"
			Set tableList = dbConnection.Execute("SELECT sysobjects.name, sysusers.name as owner FROM sysobjects, sysusers WHERE type = 'U' AND sysobjects.uid = sysusers.uid ORDER BY sysobjects.name")
		Else
			dbType = "Access"
			Set tableList = dbConnection.Execute("SELECT name FROM MSysObjects WHERE type = 1 AND name NOT LIKE 'MSys%' ORDER BY name")
		End If
		If Not tableList.EOF Then
			tableListHTML = ("<select size=""20"" onDblClick=""showTable(this)"">" & vbNewline)
			dataScript = ("<script>" & vbNewline)
			dataScript = dataScript & ("function showTable(id)" & vbNewline)
			dataScript = dataScript & ("{" & vbNewline)
			dataScript = dataScript & ("	document.getElementById(""tableData"").innerHTML = eval(id.value);" & vbNewline)
			dataScript = dataScript & ("}" & vbNewline)
			dataScript = dataScript & ("" & vbNewline)
			Do Until tableList.EOF
'						dataScript = dataScript & ("var " & tableList("owner") & " = new Object();" & vbNewline)
				If (dbType = "SQL") Then
					SQLConsoledbRS.Open "SELECT count(*) as recordCount FROM " & tableList("owner") & "." & tableList("name") & "", dbConnection, 2, 3
				Else
					SQLConsoledbRS.Open "SELECT count(*) as recordCount FROM " & tableList("name") & "", dbConnection, 2, 3
				End If
				tblRecordCount = SQLConsoledbRS("recordCount")
				SQLConsoledbRS.Close
				If (dbType = "SQL") Then
					SQLConsoledbRS.Open "SELECT TOP 0 * FROM " & tableList("owner") & "." & tableList("name") & "", dbConnection, 2, 3
					tableListHTML = tableListHTML & ("<option value=""" & tableList("owner") & "_" & tableList("name") & """>" & tableList("owner") & "." & tableList("name") & " (" & SQLConsoledbRS.Fields.Count & " columns, " & tblRecordCount & " rows)</option>" & vbNewline)
					dataScript = dataScript & (vbNewline & "var " & tableList("owner") & "_" & tableList("name") & " = '<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%""><tr>")
					dataScript = dataScript & ("<tr bgcolor=""#000000""><td colspan=""2"" align=""center""><b><font color=""#FFFFFF"">" & tableList("owner") & "." & tableList("name") & " (" & SQLConsoledbRS.Fields.Count & " columns, " & tblRecordCount & " rows)</font></b></td></tr>")
				Else
					SQLConsoledbRS.Open "SELECT * FROM " & tableList("name") & "", dbConnection, 2, 3
					tableListHTML = tableListHTML & ("<option value=""" & tableList("name") & """>" & tableList("name") & " (" & SQLConsoledbRS.Fields.Count & " columns, " & tblRecordCount & " rows)</option>" & vbNewline)
					dataScript = dataScript & (vbNewline & "var " & tableList("name") & " = '<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%""><tr>")
					dataScript = dataScript & ("<tr bgcolor=""#000000""><td colspan=""2"" align=""center""><b><font color=""#FFFFFF"">" & tableList("name") & " (" & SQLConsoledbRS.Fields.Count & " columns, " & tblRecordCount & " rows)</font></b></td></tr>")
				End If
				If Not tableList.EOF Then
					For field = 0 To SQLConsoledbRS.Fields.Count-1
						If (bgcolor = "#FFFFFF") Then bgcolor = "#DDDDDD" Else bgcolor = "#FFFFFF" End If
						dataScript = dataScript & ("<tr bgcolor=""" & bgcolor & """><td>" & SQLConsoledbRS.Fields(field).Name & "</td><td>" & getDataType(SQLConsoledbRS.Fields(field).Type) & "(" & SQLConsoledbRS.Fields(field).DefinedSize & ")" & "</td></tr>")
					Next
	'				Response.Write(SQLConsoledbRS.Fields(SQLConsoledbRS.Fields.Count-1).Name)
				End If
				dataScript = dataScript & ("</table>';" & vbNewline & vbNewline)
				tableList.MoveNext
				SQLConsoledbRS.Close
			Loop
			tableListHTML = tableListHTML & ("</select>" & vbNewline)
			dataScript = dataScript & ("</script>" & vbNewline)

			Response.Write("<tr>" & vbNewline)
			Response.Write("<td width=""1%"" valign=""top"">" & vbNewline)
			Response.Write(tableListHTML & vbNewline)
			Response.Write(dataScript & vbNewline)
			Response.Write("</td>" & vbNewline)
			Response.Write("<td width=""99%"" valign=""top"" id=""tableData"">" & vbNewline)
			Response.Write("</td>" & vbNewline)
			Response.Write("</tr>" & vbNewline)
		End If
		Response.Write("</table>")
		closeDatabase
	End Function

	Function RenderSQL
		SQLConsoleselectReturn = "None"
		SQLConsoleotherReturn = "None"
		
		selectStatement = Request.Form("selectStatement")
		otherStatement = Request.Form("otherStatement")
		
		%>	
		<table border="0" cellpadding="0" cellspacing="1" width="100%">
			<tr>
				<td>
					<table border="0" cellpadding="3" cellspacing="0" width="100%">
						<tr>
							<td class="whitebg">
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td align="center"><b><big>SQL Console</big></b></td>
									</tr>
								</table>
		<%					
'		If (Len(otherStatement) > 0) Then
			runOtherSQL(otherStatement)
'		End If
		
'		If (Len(selectStatement) > 0) Then
			runSelectSQL(selectStatement)
'		End If
		%>
								<form name="SQLConsole" action="<%=Request.ServerVariables("URL")%>" method="post">
									<table border="1" align="center" width="100%" style="border-collapse: collapse;">
										<tr>
											<td align="right" valign="top">
												Select:<br />
												Update:<br /> 
												Insert:<br /> 
												Delete:<br /> 
												Create:<br /> 
												Drop:<br /> 
												Alter:<br /> 
												<br />
												Grant:<br />
											</td>
											<td colspan="2" valign="top">
												SELECT field FROM table [WHERE criteria] [ORDER BY criteria]<br />
												UPDATE table SET field = value WHERE criteria<br />
												INSERT INTO table ( field ) VALUES ( value )<br />
												DELETE field FROM table [WHERE criteria]<br />
												CREATE TABLE table ( field type[(size)] [primary] [identity(1,1)] [[not] null] )<br />
												DROP TABLE table<br />
												ALTER TABLE table ADD field type<br />
												ALTER TABLE table DROP COLUMN field<br />
												<span title="The GRANT command gives permissions to users and roles.

Ex. GRANT SELECT, INSERT ON table1 TO Public
This will give SELECT and INSERT permissions to the Public on table table1.">GRANT</span> <span title="TABLE PERMISSIONS:
------------------
SELECT
INSERT
UPDATE
DELETE

DATABASE PERMISSIONS:
---------------------
CREATE TABLE
ALTER DATABASE
GRANT

NOTE: More than one permission can be added, however table and database permissions cannot be performed in same statement.">&lt;permissions&gt;</span> ON <span title="Any one (1) table.">&lt;table&gt;</span> TO <span title="Any user or role that exists on the SQL server.">&lt;user/role&gt;</span> <span title="Use WITH GRANT OPTION to allow the user to grant permissions of the same type to other users.">[WITH GRANT OPTION]</span>
											</td>
										</tr>
										<tr>
											<td width="1%" align="right">Select:</td>
											<td width="99%">
												<textarea class="special" cols="80" rows="5" name="selectStatement"><%=selectStatement%></textarea><br />
												<input type="button" class="special" name="clearSelect" value="Clear" onClick="document.forms.SQLConsole.selectStatement.value='';">
												<input type="submit" name="formFunction" class="special" value="Run SQL">
											</td>
										</tr>
										<tr>
											<td></td>
											<td>
												<b>Last Statement:</b> <%= SQLConsoleselectReturn %><br>
												<ul><small><%=Replace(Replace(Replace(selectStatement," ","&nbsp;"),vbNewline,"<br />"),vbTab,"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")%></small></ul>
											</td>
										</tr>
										<tr>
											<td width="1%" align="right">Other:</td>
											<td width="99%">
												<textarea class="special" cols="80" rows="5" name="otherStatement"><% If (Request.Form("saveQuery") <> "") Then Response.Write(otherStatement) End If %></textarea><br />
												<input type="button" class="special" name="clearOther" value="Clear" onClick="document.SQLConsole.otherStatement.value='';">
												<input type="submit" name="formFunction" class="special" value="Run SQL">
												<input type="checkbox" name="saveQuery" value="saveQuery" <% If (Request("saveQuery") <> "") Then Response.Write("CHECKED") End If %>><small>&nbsp;Save Query</small>
											</td>
										</tr>
										<tr>
											<td></td>
											<td>
												<b>Last Statement:</b> <%= SQLConsoleotherReturn %><br>
												<blockquote><small><%=Replace(Replace(Replace(otherStatement," ","&nbsp;"),vbNewline,"<br />"),vbTab,"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")%></small></blockquote>
											</td>
										</tr>
										<tr>
											<td colspan="2">
												<hr>
												<%= SQLConsoledatabaseOutput %>
												<hr>
												<% printAvailableTables %>
											</td>
										</tr>
									</table>
								</form>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<%
	End Function
	
	RenderSQL
%>