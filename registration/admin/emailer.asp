<% Server.ScriptTimeout = 600 %>
<% authNotNeeded = True %>
<!--#include file="functions.asp"-->
<% 
'If Not (getUserInfo(getUserID,"accesslevel") = "Admin") Then 
'	Response.Write "<CENTER><B>You do not have access to this section of the CMS, your access is Read-Only, contact your admin to get rights.</B></CENTER>"
'	Response.End
'End If
%>
<html>
  <head>
<%
Session("cookieTest") = "Session Cookie"
If (Session("cookieTest") <> "Session Cookie") Then
	Response.Write("<script> alert(""Cookies must be turned on.""); </script>")
	Session("cookieTest") = ""
	Response.End
End If
Session("cookieTest") = ""
%>
    <style>
	body {
		background: white;
	}
  body, td, input, select, textarea {
    font-family: "MS Sans Serif";
    font-size: 10pt
  }
  .bigfont {
    font-size: 14pt
  }
  .smallfont {
    font-size: 8pt
  }
  .pixel {
    font-family: "Arial";
    font-size: 1px
  }
  .bluebg {
    background: #0099ff;  /*#0066ff*/
    border: 1px solid black
  }
  .tablebg {
    padding-left: 10px;
    padding-right: 10px;
    background: #DDDDDD;
    border: 1px solid #999999;
    border-top: 3px solid #0099ff;
    border-bottom: 3px solid #0099ff;
    margin-bottom: 5px
  }
  .heading {
    font-size: 10pt;
    border-bottom: 2px outset white
  }
  .darkheading {
    background: #AAAAAA;
    border-bottom: 2px outset white
  }
  .whitebf {
    background: white;
    border: 1px solid white
  }
  .blue {
    color: #0099ff
  }
  .white {
    color: white
  }
  .black {
    color: black
  }
  .bold {
    font-weight: bold
  }
  .blackborder {
    border: 1px solid black
  }
  .blueborder {
    border: 1px solid #0000ff
  }
  td.menubar {
    padding-top: 4px;
    padding-bottom: 5px;
    padding-left: 5px;
    padding-right: 5px
  }
  .headerbar {
    background: black;
    color: white;
    font-weight: bold
  }
  a {
    color: black
  }
  a:hover {
    text-decoration: none
  }
  a.functions {
    font-size: 8pt;
    color: black;
    font-weight: bold;
    text-decoration: none
  }
  a.functions:hover {
    text-decoration: underline
  }
  a.selected {
    color: black;
    font-weight: bold
  }
  a.menu {
    color: white
  }
  a.menu:hover {
    text-decoration: none
  }
  a.menubar {
    color: blue
  }
  a.menubar:hover {
    text-decoration: none
  }
  .myInput {
    background: #CCCCCC
  }
  .tabox {
    background: #CCCCCC;
    font-size: 8pt;
    scrollbar-3dlight-color:;
    scrollbar-arrow-color:;
    scrollbar-base-color:;
    scrollbar-darkshadow-color:;
    scrollbar-face-color: #CCCCCC;
    scrollbar-highlight-color:;
    scrollbar-shadow-color:
  }
  .tafixed {
    height: 350px;
    background: #CCCCCC;
    font-family: "Courier New";
    font-size: 8pt;
    scrollbar-3dlight-color:;
    scrollbar-arrow-color:;
    scrollbar-base-color:;
    scrollbar-darkshadow-color:;
    scrollbar-face-color: #CCCCCC;
    scrollbar-highlight-color:;
    scrollbar-shadow-color:
  }
  .tanormal {
    height: 350px;
    background: #CCCCCC;
    font-family: "MS Sans Serif";
    font-size: 10pt;
    scrollbar-3dlight-color:;
    scrollbar-arrow-color:;
    scrollbar-base-color:;
    scrollbar-darkshadow-color:;
    scrollbar-face-color: #CCCCCC;
    scrollbar-highlight-color:;
    scrollbar-shadow-color:
  }
    </style>
  <meta name="Microsoft Theme" content="modified-powerplugs-web-templates-art3dblue 011, default">
</head>
  <body background="../../_themes/modified-powerplugs-web-templates-art3dblue/background.gif" bgcolor="#C0C0C0" text="#000000" link="#6A6A6A" vlink="#808080" alink="#FFFFFF"><!--mstheme--><font face="Arial, Arial, Helvetica">
<form name="frmMenu" id="frmMenu" method="post">
  <!--mstheme--></font><table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td><!--mstheme--><font face="Arial, Arial, Helvetica">
  <!--mstheme--></font><table border="0" cellpadding="0" cellspacing="0" width="100%" class="tablebg">
    <tr>
      <td align="center"><!--mstheme--><font face="Arial, Arial, Helvetica">
	<span class="bigfont"><b>List Mailer</b></span><br>
      <!--mstheme--></font></td>
    </tr>
  </table><!--mstheme--><font face="Arial, Arial, Helvetica">
      <!--mstheme--></font></td>
    </tr>
    <tr>
      <td><!--mstheme--><font face="Arial, Arial, Helvetica">
<script>
  var isChanged = false;
  function changed() {
    isChanged = true;
  }

  function initPage(page) {
    if (isChanged) {
      lsReturn = prompt("You have modified something here,\nAre you sure you want to continue?");
      if (prompt) {
        navTo(page);
      }
    }
    else {
      navTo(page);
    }
    return true;
  }

  function doFunction(func,id) {
    if (func == "addList") {
      navTo(func);
    }
    else if (func == "renameList") {
      var currentName = eval("frmMenu.listName_" + id + "").value;
      var newName = prompt("Enter a new list name:",currentName);
      if ((newName == currentName) || (newName == null) || (newName == undefined) || (newName == "")) {
        // do nothing
      }
      else {
        frmMenu.frmParam1.value = id;
        frmMenu.frmParam2.value = newName;
        navTo(func);
      }
    }
    else if (func == "renameNotes") {
      var currentNotes = eval("frmMenu.listNotes_" + id + "").value;
      var newNotes = prompt("Enter new notes:",currentNotes);
      if ((newNotes == currentNotes) || (newNotes == null) || (newNotes == undefined)) {
        // do nothing
      }
      else {
        if (newNotes == "") {
          newNotes = " ";
        }
        frmMenu.frmParam1.value = id;
        frmMenu.frmParam2.value = newNotes;
        navTo(func);
      }
	}
    else if (func == "deleteList") {
      var currentName = eval("frmMenu.listName_" + id + "").value;
      var deleteOkay = confirm("Are you sure you want to delete '" + currentName + "'?");
      if (deleteOkay) {
        frmMenu.frmParam1.value = id;
        navTo(func);
      }
    }
    else if (func == "memberAdmin") {
      frmMenu.frmParam1.value = id;
      navTo(func);
    }
	else if (func == "searchAdmin")
	{
		navTo(func);
	}
    else if (func == "addMember") {
      navTo(func);
    }
    else if (func == "renameMember") {
      var currentName = eval("frmMenu.memberName_" + id + "").value;
      var newName = prompt("Enter a new member name:",currentName);
      if ((newName == currentName) || (newName == null) || (newName == undefined) || (newName == "")) {
        // do nothing
      }
      else {
        frmMenu.frmParam1.value = id;
        frmMenu.frmParam2.value = newName;
        navTo(func);
      }
    }
    else if (func == "renameMemberEmail") {
      var currentEmail = eval("frmMenu.memberEmail_" + id + "").value;
      var newEmail = prompt("Enter a new member email address:",currentEmail);
      if ((newEmail == currentEmail) || (newEmail == null) || (newEmail == undefined) || (newEmail == "")) {
        // do nothing
      }
      else {
        frmMenu.frmParam1.value = id;
        frmMenu.frmParam2.value = newEmail;
        navTo(func);
      }
    }
    else if (func == "renameMemberNotes") {
      var currentNotes = eval("frmMenu.memberNotes_" + id + "").value;
      var newNotes = prompt("Enter a new member notes:",currentNotes);
      if ((newNotes == currentNotes) || (newNotes == null) || (newNotes == undefined)) {
        // do nothing
      }
      else {
        if (newNotes == "") {
          newNotes = " ";
        }
        frmMenu.frmParam1.value = id;
        frmMenu.frmParam2.value = newNotes;
        navTo(func);
      }
    }
    else if (func == "deleteMember") {
      var currentName = eval("frmMenu.memberName_" + id + "").value;
      var deleteOkay = confirm("Are you sure you want to delete '" + currentName + "'?");
      if (deleteOkay) {
        frmMenu.frmParam1.value = id;
        navTo(func);
      }
    }
    else if (func == "disableMember") {
      frmMenu.frmParam1.value = id;
      navTo(func);
    }
    else if (func == "enableMember") {
      frmMenu.frmParam1.value = id;
      navTo(func);
    }
    else if (func == "disableSubscribe") {
      frmMenu.frmParam1.value = id;
      navTo(func);
    }
    else if (func == "enableSubscribe") {
      frmMenu.frmParam1.value = id;
      navTo(func);
    }
    else if (func == "sendEmail") {
      frmMenu.frmParam1.value = id;
      navTo(func);
    }
    else if (func == "massDelete") {
      var checkedText = "";
      var checkedCount = 0;
      if (frmMenu.massChecked.length == undefined) {
      	 checkedCount++;
      	 checkedText = frmMenu.massChecked.value;
      }
      else {
        for (var xPos = 0; xPos < frmMenu.massChecked.length; xPos++) {
			if (frmMenu.massChecked[xPos].checked == true) {
  			  checkedCount++;
			  if (checkedCount == 1) {
			    checkedText = frmMenu.massChecked[xPos].value;
			  }
			  else {
			    checkedText += "," + frmMenu.massChecked[xPos].value;
			  }
		   }
        }
      }
      if (checkedCount > 0) {
        var deleteOkay = confirm("Are you sure you want to delete the checked lists?");
        if (deleteOkay) {
		   frmMenu.frmParam1.value = checkedText;
          navTo(func);
        }
      }
    }
    else if (func == "massMove") {
      var checkedText = "";
      var checkedCount = 0;
      if (frmMenu.massChecked.length == undefined) {
      	 checkedCount++;
        checkedText = frmMenu.massChecked.value;
      }
      else {
        for (var xPos = 0; xPos < frmMenu.massChecked.length; xPos++) {
		  	if (frmMenu.massChecked[xPos].checked == true) {
  			  checkedCount++;
			  if (checkedCount == 1) {
			    checkedText = frmMenu.massChecked[xPos].value;
			  }
			  else {
			    checkedText += "," + frmMenu.massChecked[xPos].value;
			  }
          }
        }
      }
      if (checkedCount > 0) {
        var moveOkay = confirm("Are you sure you want to move the checked lists?");
        if (moveOkay) {
		   frmMenu.frmParam1.value = checkedText;
		   frmMenu.frmParam2.value = frmMenu.massMoveTo.value;
          navTo(func);
        }
      }
    }
    else {
      window.status = "Action: " + func + " ID:" + id;
    }
  }

  function navTo(page) {
    if (page == "returnToCMS") {
      self.location.href = "/registration/admin/index.asp";
    }
    else {
      frmMenu.frmFunction.value = page;
      frmMenu.submit();
    }
  }

  function checkAllToggle(fromID) {
    var x = 0;
    var loopDone = false;
    do {
      x++;
      toCheck = document.getElementById("massChecked_" + x);
      if (toCheck == null) {
        loopDone = true;
      }
      else {
        toCheck.checked = fromID.checked;
      }
    } while(!loopDone);
  }
</script>
<%
Dim alertMessage
Function printMenu
	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" class=""tablebg"">")
	Response.Write("<tr>")
	Response.Write("<td class=""menubar"" style=""padding-bottom:5px;"">")
	Response.Write("&nbsp;&nbsp;|&nbsp;&nbsp;")
	printMenuItem "listAdmin","Manage Lists","Administer the mailing lists."
	printMenuItem "searchAdmin","Search Lists","Search for a user among all lists and perform removal."
	printMenuItem "memberAdmin","Manage Members","Administer the members of a list."
	printMenuItem "importAdmin","Import To List","Import a batch of names to a list."
	printMenuItem "sendEmail","Send Email","Send email to a list."
	printMenuItem "returnToCMS","Return to CMS","Return to the Customer Management System."
	Response.Write("</td>")
	Response.Write("</tr>")
	Response.Write("</table>")
End Function

Function printMenuItem(func, funcLabel, funcMouseover)
	Response.Write("<a class=""")
	If (formFunction = func) Then
		Response.Write("selected")
	Else
		Response.Write("menubar")
	End If
	Response.Write(""" href=""javascript:initPage('" & func & "');"" onMouseDown=""window.status='';return true;"" onMouseOver=""window.status='" & funcMouseover & "';return true;"" onMouseOut=""window.status='';return true;"">" & funcLabel & "</a>" & vbNewline)
	Response.Write("&nbsp;&nbsp;|&nbsp;&nbsp;" & vbNewline)
End Function

If (Request.QueryString("function") = "mergetest") Then
	MergeList
End If

Function MergeList
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Server.ScriptTimeout = 180
	Response.Buffer = True
	
	dbConnection.Execute("DELETE * FROM emailer_tblMembers WHERE memberOf = 52")
	' memberId,memberName,memberEmail,memberOf,memberDisabled,memberNotes
	Set emails = dbConnection.Execute("SELECT DISTINCT memberEmail FROM emailer_tblMembers WHERE memberOf <> 52")
	Do Until (emails.EOF)
		Set thisEmail = dbConnection.Execute("SELECT TOP 1 memberName, memberNotes FROM emailer_tblMembers WHERE memberEmail = '" & emails("memberEmail") & "'")
		If (thisEmail.EOF) Then
			thisName = "None"
			thisNotes = ""
		Else
			thisName = thisEmail("memberName")
			thisNotes = thisEmail("memberNotes")
		End If
		Set thisEmail = Nothing
		
		dbConnection.Execute("INSERT INTO emailer_tblMembers ( memberName, memberEmail, memberOf, memberDisabled, memberNotes ) VALUES ( '" & thisName & "', '" & emails("memberEmail") & "', 52, False, '" & thisNotes & "' )")
		emails.MoveNext
	Loop
	Set emails = Nothing

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printSearchAdmin
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	frmSearchIn = Request.Form("search")
	frmSearchFor = Request.Form("searchfor")
	frmParam1 = Request.Form("frmParam1")
	frmSelected = Request.Form("selected")

	If (Len(frmParam1) > 0) Then
		For Each frmSelect In Split(frmSelected, ",")
			frmSelect = Trim(frmSelect)
			If (frmParam1 = "Enable") Then
				dbConnection.Execute("UPDATE emailer_tblMembers SET memberDisabled = false WHERE memberID = " & frmSelect & "")
			ElseIf (frmParam1 = "Disable") Then
				dbConnection.Execute("UPDATE emailer_tblMembers SET memberDisabled = true WHERE memberID = " & frmSelect & "")
			ElseIf (frmParam1 = "Delete") Then
				dbConnection.Execute("DELETE FROM emailer_tblMembers WHERE memberID = " & frmSelect & "")
			End If
		Next
		closeDatabase
		openDatabase
	End If
	
	%>
	<!--mstheme--></font><table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse;" bordercolordark="#4F79A4" bordercolorlight="#4F79A4">
		<tr>
			<td><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Search In:</b><!--mstheme--></font></td>
			<td><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Search For:</b><!--mstheme--></font></td>
			<td><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Submit</b><!--mstheme--></font></td>
		</tr>
		<tr>
			<td><!--mstheme--><font face="Arial, Arial, Helvetica">
				<select size="1" name="search" />
					<option value="memberName" <% If (frmSearchIn = "memberName") Then %>selected<% End If %>>Name</option>
					<option value="memberEmail" <% If (frmSearchIn = "memberEmail") Then %>selected<% End If %>>Email Address</option>
				</select>
			<!--mstheme--></font></td>
			<td><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="text" size="40" name="searchfor" value="<%=frmSearchFor%>" /><!--mstheme--></font></td>
			<td><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="submit" value="Search" /><!--mstheme--></font></td>
		</tr>
	</table><!--mstheme--><font face="Arial, Arial, Helvetica"><br />
	<%
	If (Len(frmSearchIn) > 0) Then
		
		Set dbEmails = dbConnection.Execute("SELECT * FROM emailer_tblMembers, emailer_tblMailingLists WHERE memberOf = listID AND " & frmSearchIn & " LIKE '%" & frmSearchFor & "%'")
		If (dbEmails.EOF) Then
			%>
				<b>There are no emails matching the search.</b>
			<%
		Else
			%>
				<script>
					function __selectAll()
					{
						try { var elements = document.body.getElementsByTagName("input"); } catch(ex) { alert("Couldn't get input tags."); }
						try { var checkall = document.getElementById("checkall"); } catch(ex) { alert("Couldn't get checkall by id."); }

						for (var x = 0; x < elements.length; x++)
						{
							if (elements[x].type == "checkbox" && elements[x] != checkall)
							{
								elements[x].checked = checkall.checked;
							}
						}
					}
					
					function __checkCheckAll()
					{
						document.getElementById("checkall").checked = true;
					}
					
					function __doAction(src)
					{
						frmMenu.frmParam1.value = src.value;
						frmMenu.submit();
					}
				</script>
				<!--mstheme--></font><table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse;" bordercolordark="#4F79A4" bordercolorlight="#4F79A4">
					<tr>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Operations:</b><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="button" name="searchAction" value="Enable" onclick="__doAction(this);" /><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="button" name="searchAction" value="Disable" onclick="__doAction(this);" /><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="button" name="searchAction" value="Delete" onclick="__doAction(this);" /><!--mstheme--></font></td>
					</tr>
				</table><!--mstheme--><font face="Arial, Arial, Helvetica"><br />
				<!--mstheme--></font><table border="1" cellpadding="3" cellspacing="0" width="100%" style="border-collapse: collapse;" bordercolordark="#4F79A4" bordercolorlight="#4F79A4">
					<tr>
						<td width="22"><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="checkbox" name="checkall" onclick="__selectAll()" /><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><b>ID</b><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Name</b><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Email Address</b><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><b>List</b><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Disabled</b><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><b>Notes</b><!--mstheme--></font></td>
					</tr>
			<%
			notAllChecked = False
			Do Until (dbEmails.EOF)
				If (InStr(frmSelected, Fix(dbEmails("memberID"))) = "0") Then
					notAllChecked = True
				End If
				%>
					<tr>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><input type="checkbox" name="selected" value="<%=dbEmails("memberID")%>" <% If (InStr(frmSelected, Fix(dbEmails("memberID"))) > 0) Then %>checked<% End If %> /><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><%=dbEmails("memberID")%><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><%=dbEmails("memberName")%><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><%=dbEmails("memberEmail")%><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><%=dbEmails("listName")%><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><%=dbEmails("memberDisabled")%><!--mstheme--></font></td>
						<td><!--mstheme--><font face="Arial, Arial, Helvetica"><%=dbEmails("memberNotes")%><!--mstheme--></font></td>
					</tr>
				
				<%
				dbEmails.MoveNext
			Loop
			%>
				</table><!--mstheme--><font face="Arial, Arial, Helvetica">
			<%
			If Not (notAllChecked) Then
				%>
					<script>
						__checkCheckAll();
					</script>
				<%
			End If
		End If
	End If
	
	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printListAdmin
'	On Error Resume Next
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set mailingLists = dbConnection.Execute("SELECT * " & _
		"FROM emailer_tblMailingLists " & _
		"ORDER BY listName")
	If Err <> 0 Then
		Response.Write("Error with mailing list query, could not get handle on mailing list.<br>" & Err.Description)
		Exit Function
	Else
		Response.Write("<table border=""1"" cellpadding=""0"" cellspacing=""0"" width=""100%"" style=""border-collapse: collapse;"" border-color=""black""><tr><td bgcolor=""black"">" & vbNewline)
		Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%""><tr><td bgcolor=""#c0c0c0"">" & vbNewline)
		Response.Write("<table border=""0"" cellpadding=""3"" cellspacing=""0"" width=""100%"" style=""border-collapse: collapse;"" border-color=""black"">" & vbNewline)
		Response.Write("<tr class=""smallfont headerbar"">" & vbNewline)
		Response.Write("<td width=""2%"" align=""center"">#</td>" & vbNewline)
		Response.Write("<td width=""1%"" align=""center"">ID</td>" & vbNewline)
		Response.Write("<td width=""20%"" align=""left"">List Name</td>" & vbNewline)
		Response.Write("<td width=""10%"" align=""left"" title=""Allow/Disallow subscriptions to be taken from a form.  If it is not allowed, you can only add people to the list from this interface."">Can Sub.?</td>" & vbNewline)
		Response.Write("<td width=""10%"" align=""center"">Members</td>" & vbNewline)
'		Response.Write("<td width=""15%"" align=""center"">Date Added</td>" & vbNewline)
		Response.Write("<td width=""30%"" align=""left"">Notes</td>" & vbNewline)
		Response.Write("<td width=""20%"" align=""center"">Functions</td>" & vbNewline)
		Response.Write("</tr>" & vbNewline)
		listPos = 0
		Do Until mailingLists.EOF
			Set memberCountDB = dbConnection.Execute("SELECT count(members.memberID) as totalMemberCount " & _
				"FROM emailer_tblMembers members " & _
				"WHERE memberOf = " & mailingLists("listID") & "")
			totalMemberCount = memberCountDB("totalMemberCount")
			Set memberCountDB = Nothing
			Set memberCountDB = dbConnection.Execute("SELECT count(members.memberID) as activeMemberCount " & _
				"FROM emailer_tblMembers members " & _
				"WHERE memberOf = " & mailingLists("listID") & " AND memberDisabled = False")
			activeMemberCount = memberCountDB("activeMemberCount")
			Set memberCountDB = Nothing
			listPos = listPos + 1
			Response.Write("<tr class=""smallfont"">" & vbNewline)
			Response.Write("<td align=""center"">" & listPos & "</td>" & vbNewline)
			Response.Write("<td align=""center"">" & mailingLists("listID") & "</td>" & vbNewline)
			Response.Write("<td><input type=""hidden"" name=""listName_" & mailingLists("listID") & """ value=""" & mailingLists("listName") & """><a class=""functions"" href=""javascript:doFunction('renameList','" & mailingLists("listID") & "');"" onMouseOver=""window.status='Rename this list.';return true;"" onMouseOut=""window.status='';return true;"">" & mailingLists("listName") & "</a></td>" & vbNewline)
			If (mailingLists("listAllowSubscribe") = True) Then
				Response.Write("<td align=""center""><input type=""checkbox"" name=""allowSubscribe_" & mailingLists("listID") & """ checked onClick=""doFunction('disableSubscribe','" & mailingLists("listID") & "')"" onMouseOver=""window.status='Disallow form subscriptions for this list.';return true;"" onMouseOut=""window.status='';return true;""></td>" & vbNewline)
			Else
				Response.Write("<td align=""center""><input type=""checkbox"" name=""allowSubscribe_" & mailingLists("listID") & """ onClick=""doFunction('enableSubscribe','" & mailingLists("listID") & "')"" onMouseOver=""window.status='Allow form subscriptions for this list.';return true;"" onMouseOut=""window.status='';return true;""></td>" & vbNewline)
			End If
			Response.Write("<td align=""center"" title=""Active Members/Total Members"">" & activeMemberCount & "/" & totalMemberCount & "</td>" & vbNewline)
'			Response.Write("<td align=""center"" nowrap>" & mailingLists("listAdded") & "</td>" & vbNewline)
			If (Replace(mailingLists("listNotes")," ","") = "") Then
				addEdit = "&lt;...none...&gt;"
			Else
				addEdit = ""
			End If
			Response.Write("<td><input type=""hidden"" name=""listNotes_" & mailingLists("listID") & """ value=""" & mailingLists("listNotes") & """><a class=""functions"" href=""javascript:doFunction('renameNotes','" & mailingLists("listID") & "');"" onMouseOver=""window.status='Change this list\'s notes.';return true;"" onMouseOut=""window.status='';return true;"">" & mailingLists("listNotes") & addEdit & "</a></td>" & vbNewline)
			Response.Write("<td align=""right"" nowrap>" & vbNewline)
			Response.Write("&nbsp|&nbsp;" & vbNewline)
			Response.Write("<a class=""functions"" href=""javascript:doFunction('sendEmail','list_" & mailingLists("listID") & "');"" onMouseOver=""window.status='Send email to this list.';return true;"" onMouseOut=""window.status='';return true;"">Send Email</a>" & vbNewline)
			Response.Write("&nbsp|&nbsp;" & vbNewline)
			Response.Write("<a class=""functions"" href=""javascript:doFunction('memberAdmin','" & mailingLists("listID") & "');""  onMouseOver=""window.status='Manage the members of this list.';return true;"" onMouseOut=""window.status='';return true;"">Manage Members</a>" & vbNewline)
'			Response.Write("&nbsp|&nbsp;" & vbNewline)
'			Response.Write("<a class=""functions"" href=""javascript:doFunction('renameList','" & mailingLists("listID") & "');""  onMouseOver=""window.status='Rename this list.';return true;"" onMouseOut=""window.status='';return true;"">Rename</a>" & vbNewline)
'			Response.Write("&nbsp|&nbsp;" & vbNewline)
'			Response.Write("<a class=""functions"" href=""javascript:doFunction('renameNotes','" & mailingLists("listID") & "');""  onMouseOver=""window.status='Change this list\'s notes.';return true;"" onMouseOut=""window.status='';return true;"">Edit Notes</a>" & vbNewline)
			Response.Write("&nbsp|&nbsp;" & vbNewline)
			Response.Write("<a class=""functions"" href=""javascript:doFunction('deleteList','" & mailingLists("listID") & "');""  onMouseOver=""window.status='Delete this list.';return true;"" onMouseOut=""window.status='';return true;"">Delete List</a>" & vbNewline)
			Response.Write("&nbsp|&nbsp;" & vbNewline)
			Response.Write("</td>" & vbNewline)
			Response.Write("</tr>" & vbNewline)
			mailingLists.MoveNext
		Loop

		Response.Write("<tr class=""smallfont"">" & vbNewline)
		Response.Write("<td align=""center"">New</td>" & vbNewline)
		Response.Write("<td align=""center""></td>" & vbNewline)
		Response.Write("<td><input type=""text"" name=""frmNewListName"" class=""myInput smallfont"" style=""width:100%"" value=""" & Session("newListName") & """></td>" & vbNewline)
		Response.Write("<td align=""center""><input type=""checkbox"" name=""frmNewListSubscribeable"" value=""true""></td>" & vbNewline)
		Response.Write("<td align=""center""></td>" & vbNewline)
'		Response.Write("<td align=""center""></td>" & vbNewline)
		Response.Write("<td><input type=""text"" name=""frmNewListNotes"" class=""myInput smallfont"" style=""width:100%"" value=""" & Session("newListNotes") & """></td>" & vbNewline)
		Response.Write("<td align=""right"" nowrap>" & vbNewline)
		Response.Write("&nbsp|&nbsp;" & vbNewline)
		Response.Write("<a class=""functions"" href=""javascript:doFunction('addList','-1');""  onMouseOver=""window.status='Create this as a new mailing list.';return true;"" onMouseOut=""window.status='';return true;"">Add</a>" & vbNewline)
		Response.Write("&nbsp|&nbsp;" & vbNewline)
		Response.Write("</td>" & vbNewline)
		Response.Write("</tr>" & vbNewline)

		Response.Write("</td></tr></table>" & vbNewline)
		Response.Write("</td></tr></table>" & vbNewline)
		Response.Write("</td></tr></table>" & vbNewline)

		%>
		<br>
		<big><b>Help</b></big><br>
		<br>
		<u>Properties</u><br>
		<li># - Count of the list.
		<li>ID - The database ID, needed for the form.  (See below)
		<li>List Name - The name you would like to call this list.
		<li>Can Sub.? - When checked, allows subscriptions from the form (See below), when not checked, only manual additions from here are allowed. 
		<li>Members - First number is how many active members, Second number is how many total members.
		<li>Notes - You can add notes to the list that are not included in the list name.
		<br>
		<br>
		<u>Functions</u>
		<li>Send Email - Will take you to the send email screen with this list checked.
		<li>Manage Members - Allows you to add/remove/change members that are in the list.
		<li>Delete List - Will allow you to delete this list and all it's members.
		<br>
		<br>
		<u>Form Information</u><br>
		<b>To setup a location for you to do a mailing list, please paste this code anywhere in an HTML page.</b><br>
		<br>
		<b>The following can be customized.  </b><br>
		<li>The line <code>&lt;input type="hidden" name="list" value="15"></code> is what you'll change to the database ID for which list this will subscribe members to. (Change is required.)
		<li>The line <code>&lt;td align="center" colspan="2">Subscribe to the Ulead mailing list:&lt;/td></code> the fields. (Change is optional.)
		<li>The line <code>&lt;input type="submit" name="add" value="Subscribe"></code> can be changed to change the words on the button, only change the value="Subscribe" value. (Change is optional.)
		<br>
		<br>
		<b>HTML Code for form, copy and paste this into the page, and modify the above 3 things.</b><br>

<textarea cols="100" rows="22" wrap="off"><table border="0" cellpadding="0" cellspacing="0">
  <form name="subscribe" method="post" action="http://www.sundancemediagroup.com/registration/subscribe.asp">
  <input type="hidden" name="list" value="15">
  <tr>
    <td align="center" colspan="2">Subscribe to the Ulead mailing list:&lt;/td>
  </tr>
  <tr>
    <td align="right">Name:&lt;/td>
    <td>&lt;input name="name" size="20">&lt;/td>
  </tr>
  <tr>
    <td align="right">Email:&lt;/td>
    <td>&lt;input name="email" size="20">&lt;/td>
  </tr>
  <tr>
    <td align="center" colspan="2">
      <input type="submit" name="add" value="Subscribe">
    </td>
  </tr>
  </form>
</table></textarea><br>

		<br>
		<b>This is what you end up with:</b> (This form is disabled.)<br>
		<!--mstheme--></font><table border="0" cellpadding="0" cellspacing="0">
		  <input type="hidden" name="list" value="15">
		  <tr>
		    <td align="center" colspan="2"><!--mstheme--><font face="Arial, Arial, Helvetica">Subscribe to the Ulead mailing list:<!--mstheme--></font></td>
		  </tr>
		  <tr>
		    <td align="right"><!--mstheme--><font face="Arial, Arial, Helvetica">Name:<!--mstheme--></font></td>
		    <td><!--mstheme--><font face="Arial, Arial, Helvetica"><input name="name" size="20"><!--mstheme--></font></td>
		  </tr>
		  <tr>
		    <td align="right"><!--mstheme--><font face="Arial, Arial, Helvetica">Email:<!--mstheme--></font></td>
		    <td><!--mstheme--><font face="Arial, Arial, Helvetica"><input name="email" size="20"><!--mstheme--></font></td>
		  </tr>
		  <tr>
		    <td align="center" colspan="2"><!--mstheme--><font face="Arial, Arial, Helvetica">
		      <input type="submit" name="add" value="Subscribe">
		    <!--mstheme--></font></td>
		  </tr>
		</table><!--mstheme--><font face="Arial, Arial, Helvetica">
		<%

		Session("newListName") = ""
		Session("newListNotes") = ""
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function addList(newName, newNotes, allowSubscribe)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	If (newName = "") Then
		alertMessage = "Please enter a list name."
		Session("newListName") = newName
		Session("newListNotes") = newNotes
		Session("allowSubscribe") = allowSubscribe
		dontDoAdd = True
	End If
	If Not (dontDoAdd) Then
		If (newNotes = "") Then
			newNotes = " "
		End If
		If (allowSubscribe = "") Then
			allowSubscribe = False
		Else
			allowSubscribe = True
		End If
		dbConnection.Execute("INSERT INTO emailer_tblMailingLists " & _
			"(" & _
				"listName," & _
				"listNotes," & _
				"listAllowSubscribe" & _
			") VALUES (" & _
				"'" & Replace(newName,"'","`") & "'," & _
				"'" & Replace(newNotes,"'","`") & "'," & _
				"" & allowSubscribe & "" & _
			")")

		'alertMessage = "New list added."
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function deleteList(listID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("DELETE * " & _
		"FROM emailer_tblMailingLists " & _
		"WHERE listID = " & listID & "")
	dbConnection.Execute("DELETE * " & _
		"FROM emailer_tblMembers " & _
		"WHERE memberOf = " & listID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function renameList(listID,newName)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("UPDATE emailer_tblMailingLists " & _
		"SET listName = '" & newName & "' " & _
		"WHERE listID = " & listID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function renameNotes(listID,newNotes)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("UPDATE emailer_tblMailingLists " & _
		"SET listNotes = '" & newNotes & "' " & _
		"WHERE listID = " & listID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function enableSubscribe(listID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("UPDATE emailer_tblMailingLists " & _
		"SET listAllowSubscribe = True " & _
		"WHERE listID = " & listID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function disableSubscribe(listID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("UPDATE emailer_tblMailingLists " & _
		"SET listAllowSubscribe = False " & _
		"WHERE listID = " & listID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printListChooser
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set mailingLists = dbConnection.Execute("SELECT * FROM emailer_tblMailingLists ORDER BY listName")

	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""50%"" align=""center"" class=""tablebg"">")
	Response.Write("<tr><td class=""heading"">Select list to edit:</td></tr>")

	Do Until mailingLists.EOF
		Response.Write("<tr>")
		Response.Write("<td>")
		Response.Write("<a href=""javascript:doFunction('memberAdmin','" & mailingLists("listID") & "');""  onMouseOver=""window.status='Manage the members of this list.';return true;"" onMouseOut=""window.status='';return true;"">" & mailingLists("listName") & "</a>" & vbNewline)
		Response.Write("</td>")
		Response.Write("</tr>")
		mailingLists.MoveNext
	Loop
	Response.Write("</table>")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function printMemberAdmin(memberOf)
'	On Error Resume Next
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Set listMembers = dbConnection.Execute("SELECT * " & _
		"FROM emailer_tblMembers " & _
		"WHERE memberOf = " & memberOf & " " & _
		"ORDER BY memberName, memberEmail")

	If Err <> 0 Then
		Response.Write("Error with list member query, could not get handle on members list.<br>" & Err.Description)
		Exit Function
	Else
		Response.Write("<table border=""1"" cellpadding=""0"" cellspacing=""0"" width=""100%"" style=""border-collapse: collapse;"" border-color=""black""><tr><td bgcolor=""black"">" & vbNewline)
		Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%""><tr><td bgcolor=""#c0c0c0"">" & vbNewline)
		Response.Write("<table border=""0"" cellpadding=""3"" cellspacing=""0"" width=""100%"" style=""border-collapse: collapse;"" border-color=""black"">" & vbNewline)
		Response.Write("<tr class=""smallfont headerbar"">" & vbNewline)
		Response.Write("<td width=""1%"" align=""left""><input type=""checkbox"" name=""checkAll"" onClick=""checkAllToggle(this)""></td>" & vbNewline)
		Response.Write("<td width=""2%"" align=""center"">#</td>" & vbNewline)
		Response.Write("<td width=""20%"" align=""left"">Name</td>" & vbNewline)
		Response.Write("<td width=""1%"" align=""left"">Off?</td>" & vbNewline)
		Response.Write("<td width=""10%"" align=""center"">E-Mail</td>" & vbNewline)
'		Response.Write("<td width=""15%"" align=""center"">Date Added</td>" & vbNewline)
		Response.Write("<td width=""30%"" align=""left"">Notes</td>" & vbNewline)
		Response.Write("<td width=""20%"" align=""center"">Functions</td>" & vbNewline)
		Response.Write("</tr>" & vbNewline)
		listPos = 0
		Do Until listMembers.EOF
			listPos = listPos + 1
			Response.Write("<tr class=""smallfont"">" & vbNewline)
			Response.Write("<td align=""center""><input type=""checkbox"" id=""massChecked_" & listPos & """ name=""massChecked"" value=""" & listMembers("memberID") & """></td>" & vbNewline)
			Response.Write("<td align=""center"">" & listPos & "</td>" & vbNewline)
			Response.Write("<td><input type=""hidden"" name=""memberName_" & listMembers("memberID") & """ value=""" & listMembers("memberName") & """><a class=""functions"" href=""javascript:doFunction('renameMember','" & listMembers("memberID") & "');""  onMouseOver=""window.status='Rename this member.';return true;"" onMouseOut=""window.status='';return true;"">" & listMembers("memberName") & "</a></td>" & vbNewline)
			If (listMembers("memberDisabled") = True) Then
				Response.Write("<td align=""center""><input type=""checkbox"" name=""memberDisabled_" & listMembers("memberID") & """ checked onClick=""doFunction('enableMember','" & listMembers("memberID") & "')"" onMouseOver=""window.status='Enable this member.';return true;"" onMouseOut=""window.status='';return true;""></td>" & vbNewline)
			Else
				Response.Write("<td align=""center""><input type=""checkbox"" name=""memberDisabled_" & listMembers("memberID") & """ onClick=""doFunction('disableMember','" & listMembers("memberID") & "')"" onMouseOver=""window.status='Disable this member.';return true;"" onMouseOut=""window.status='';return true;""></td>" & vbNewline)
			End If
			Response.Write("<td><input type=""hidden"" name=""memberEmail_" & listMembers("memberID") & """ value=""" & listMembers("memberEmail") & """><a class=""functions"" href=""javascript:doFunction('renameMemberEmail','" & listMembers("memberID") & "');""  onMouseOver=""window.status='Change this members\'s email address.';return true;"" onMouseOut=""window.status='';return true;"">" & listMembers("memberEmail") & "</a></td>" & vbNewline)
'			Response.Write("<td align=""center"" nowrap>" & listMembers("memberAdded") & "</td>" & vbNewline)
			If (listMembers("memberNotes") <> null) Then
				tempNotes = Replace(listMembers("memberNotes")," ","")
			Else
				tempNotes = ""
			End If
			If (tempNotes = "") Then
				addEdit = "&lt;...none...&gt;"
			Else
				addEdit = ""
			End If
			Response.Write("<td><input type=""hidden"" name=""memberNotes_" & listMembers("memberID") & """ value=""" & listMembers("memberNotes") & """><a class=""functions"" href=""javascript:doFunction('renameMemberNotes','" & listMembers("memberID") & "');""  onMouseOver=""window.status='Change this members\'s notes.';return true;"" onMouseOut=""window.status='';return true;"">" & listMembers("memberNotes") & addEdit & "</a></td>" & vbNewline)
			Response.Write("<td align=""right"" nowrap>" & vbNewline)
			If (listMembers("memberDisabled") = False) Then
				Response.Write("&nbsp|&nbsp;" & vbNewline)
				Response.Write("<a class=""functions"" href=""javascript:doFunction('sendEmail','member_" & listMembers("memberID") & "');""  onMouseOver=""window.status='Send email to this member.';return true;"" onMouseOut=""window.status='';return true;"">Send Email</a>" & vbNewline)
			End If
'			Response.Write("&nbsp|&nbsp;" & vbNewline)
'			Response.Write("<a class=""functions"" href=""javascript:doFunction('renameMember','" & listMembers("memberID") & "');""  onMouseOver=""window.status='Rename this member.';return true;"" onMouseOut=""window.status='';return true;"">Rename</a>" & vbNewline)
'			Response.Write("&nbsp|&nbsp;" & vbNewline)
'			Response.Write("<a class=""functions"" href=""javascript:doFunction('renameMemberEmail','" & listMembers("memberID") & "');""  onMouseOver=""window.status='Change this list\'s notes.';return true;"" onMouseOut=""window.status='';return true;"">Edit Address</a>" & vbNewline)
'			Response.Write("&nbsp|&nbsp;" & vbNewline)
'			Response.Write("<a class=""functions"" href=""javascript:doFunction('renameMemberNotes','" & listMembers("memberID") & "');""  onMouseOver=""window.status='Change this list\'s notes.';return true;"" onMouseOut=""window.status='';return true;"">Edit Notes</a>" & vbNewline)
'			Response.Write("&nbsp|&nbsp;" & vbNewline)
'			If (isDisabled) Then
'				Response.Write("<a class=""functions"" href=""javascript:doFunction('enableMember','" & listMembers("memberID") & "');""  onMouseOver=""window.status='Change this list\'s notes.';return true;"" onMouseOut=""window.status='';return true;"">Enable</a>" & vbNewline)
'			Else
'				Response.Write("<a class=""functions"" href=""javascript:doFunction('disableMember','" & listMembers("memberID") & "');""  onMouseOver=""window.status='Change this list\'s notes.';return true;"" onMouseOut=""window.status='';return true;"">Disable</a>" & vbNewline)
'			End If
			Response.Write("&nbsp|&nbsp;" & vbNewline)
			Response.Write("<a class=""functions"" href=""javascript:doFunction('deleteMember','" & listMembers("memberID") & "');""  onMouseOver=""window.status='Delete this member.';return true;"" onMouseOut=""window.status='';return true;"">Delete</a>" & vbNewline)
			Response.Write("&nbsp|&nbsp;" & vbNewline)
			Response.Write("</td>" & vbNewline)
			Response.Write("</tr>" & vbNewline)
			listMembers.MoveNext
		Loop

		Response.Write("<tr class=""smallfont"">" & vbNewline)
		Response.Write("<td align=""center""></td>" & vbNewline)
		Response.Write("<td align=""center"">New</td>" & vbNewline)
		Response.Write("<td><input type=""text"" name=""frmNewMemberName"" class=""myInput smallfont"" style=""width:100%"" value=""" & Session("newMemberName") & """></td>" & vbNewline)
		If (Session("newMemeberName") = True) Then
			thisChecked = "CHECKED"
		Else
			thisChecked = ""
		End If
		Response.Write("<td align=""center""><input type=""checkbox"" " & thisChecked & " name=""frmNewMemberDisabled""></td>" & vbNewline)
		Response.Write("<td align=""center""><input type=""text"" name=""frmNewMemberEmail"" class=""myInput smallfont"" style=""width:100%"" value=""" & Session("newMemeberEmail") & """></td>" & vbNewline)
'		Response.Write("<td align=""center"">n/a</td>" & vbNewline)
		Response.Write("<td><input type=""text"" name=""frmNewMemberNotes"" class=""myInput smallfont"" style=""width:100%"" value=""" & Session("newMemberNotes") & """></td>" & vbNewline)
		Response.Write("<td align=""right"" nowrap>" & vbNewline)
		Response.Write("&nbsp|&nbsp;" & vbNewline)
		Response.Write("<a class=""functions"" href=""javascript:doFunction('addMember','-1');""  onMouseOver=""window.status='Add this to the member list.';return true;"" onMouseOut=""window.status='';return true;"">Add</a>" & vbNewline)
		Response.Write("&nbsp|&nbsp;" & vbNewline)
		Response.Write("</td>" & vbNewline)
		Response.Write("</tr>" & vbNewline)

		Response.Write("<tr>")
		Response.Write("<td colspan=""8"">")
		Response.Write("Mass Options: ")
		Response.Write("&nbsp|&nbsp;" & vbNewline)
		Response.Write("<a class=""functions"" href=""javascript:doFunction('massDelete','-1');""  onMouseOver=""window.status='Deletes all checked members.';return true;"" onMouseOut=""window.status='';return true;"">Mass Delete</a>" & vbNewline)
		Set otherLists = dbConnection.Execute("SELECT * FROM emailer_tblMailingLists WHERE listID <> " & memberOf & " ORDER BY listName")
		If Not (otherLists.EOF) Then
			Response.Write("&nbsp|&nbsp;" & vbNewline)
			Response.Write("Move to: <select size=""1"" name=""massMoveTo"">")
			Do Until (otherLists.EOF)
				Response.Write("<option value=""" & otherLists("listID") & """>" & otherLists("listName") & "</option>")
				otherLists.MoveNext
			Loop
			Response.Write("</select>")
			Response.Write("&nbsp|&nbsp;" & vbNewline)
			Response.Write("<a class=""functions"" href=""javascript:doFunction('massMove','-1');""  onMouseOver=""window.status='Moves all checked members to the selected list.';return true;"" onMouseOut=""window.status='';return true;"">Mass Move</a>" & vbNewline)
		End If
		Set otherLists = Nothing
		Response.Write("&nbsp|&nbsp;" & vbNewline)
		Response.Write("</td>")
		Response.Write("</tr>")
		
		Response.Write("</td></tr></table>" & vbNewline)
		Response.Write("</td></tr></table>" & vbNewline)
		Response.Write("</td></tr></table>" & vbNewline)

		Session("newMemberName") = ""
		Session("newMemberDisabled") = ""
		Session("newMemberEmail") = ""
		Session("newMemberNotes") = ""
	End If

		%>
		<br>
		<big><b>Help</b></big><br>
		<br>
		<u>Properties</u><br>
		<li>[ ] - Allows you to select list for mass options.  (See Below)
		<li># - List count.
		<li>Name - Name of person, will be displayed on To field of email.
		<li>Off? - When checked, prevents this person from receiving emails when message are sent to this list.  Can be sent to individually.
		<li>E-Mail - Member's email address.
		<li>Notes - You can add notes to the list that are not included in the list name.
		<br>
		<br>
		<u>Functions</u>
		<li>Send Email - Will take you to the send email screen with this user checked.
		<li>Delete - Will allow you to delete this user.
		<br>
		<br>
		<u>Mass Options</u>
		<li>Mass Delete - This will delete all checked members.
		<li>Mass Move - This will move the checked members to the selected list.
		<%
	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Dim newMemberName, newMemberDisabled, newMemberEmail, newMemberNotes

Function addMember(whichList, newName, newEmail, newNotes, newDisabled)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	If (newName = "") Then
		alertMessage = "Please enter a member name."
		Session("newMemberName") = newName
		Session("newMemberDisabled") = newDisabled
		Session("newMemberEmail") = newEmail
		Session("newMemberNotes") = newNotes
		dontDoAdd = True
	ElseIf (newEmail = "") Then
		alertMessage = "Please enter an email address."
		Session("newMemberName") = newName
		Session("newMemberDisabled") = newDisabled
		Session("newMemberEmail") = newEmail
		Session("newMemberNotes") = newNotes
		dontDoAdd = True
	End If
	Set dbCheckMember = dbConnection.Execute("SELECT memberName FROM emailer_tblMembers WHERE memberOf = " & whichList & " AND memberEmail = '" & newEmail & "'")
	If Not (dbCheckMember.EOF) Then
		dontDoAdd = True
	End If
	If Not (dontDoAdd) Then
		If (newNotes = "") Then
			newNotes = " "
		End If
		If (newDisabled = "") Then
			newDisabled = False
		End If
		dbConnection.Execute("INSERT INTO emailer_tblMembers " & _
			"(" & _
				"memberOf," & _
				"memberName," & _
				"memberEmail," & _
				"memberDisabled," & _
				"memberNotes" & _
			") VALUES (" & _
				"" & whichList & "," & _
				"'" & Replace(newName,"'","`") & "'," & _
				"'" & Replace(newEmail,"'","`") & "'," & _
				"" & newDisabled & "," & _
				"'" & Replace(newNotes,"'","`") & "'" & _
			")")

		'alertMessage = "New list added."
	End If

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function deleteMember(memberID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("DELETE * " & _
		"FROM emailer_tblMembers " & _
		"WHERE memberID = " & memberID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function massDelete(memberList)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If
	
	memberList = Split(memberList,",")
	For Each member In memberList
		dbConnection.Execute("DELETE * " & _
			"FROM emailer_tblMembers " & _
			"WHERE memberID = " & member & "")
	Next
	
	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function massMove(memberList,toList)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If
	
	memberList = Split(memberList,",")
	For Each member In memberList
		dbConnection.Execute("UPDATE emailer_tblMembers " & _
			"SET memberOf = " & toList & " " & _
			"WHERE memberID = " & member & "")
	Next
	
	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function renameMember(memberID,newName)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	If (newName = "") Then
		newName = " "
	End If
	dbConnection.Execute("UPDATE emailer_tblMembers " & _
		"SET memberName = '" & newName & "' " & _
		"WHERE memberID = " & memberID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function renameMemberEmail(memberID,newEmail)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	If (newEmail = "") Then
		newEmail = " "
	End If
	dbConnection.Execute("UPDATE emailer_tblMembers " & _
		"SET memberEmail = '" & newEmail & "' " & _
		"WHERE memberID = " & memberID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function renameMemberNotes(memberID,newNotes)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	If (newNotes = "") Then
		newNotes = " "
	End If
	dbConnection.Execute("UPDATE emailer_tblMembers " & _
		"SET memberNotes = '" & newNotes & "' " & _
		"WHERE memberID = " & memberID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function enableMember(memberID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("UPDATE emailer_tblMembers " & _
		"SET memberDisabled = False " & _
		"WHERE memberID = " & memberID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function disableMember(memberID)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	dbConnection.Execute("UPDATE emailer_tblMembers " & _
		"SET memberDisabled = True " & _
		"WHERE memberID = " & memberID & "")

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function sendEmailPrompt(gatherFromParam)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Dim emailNames()
	Dim emailAddresses()

	sendParams = Split(gatherFromParam,"_")

	pullFromWhere = sendParams(0)
	whichID = sendParams(1)

	If (pullFromWhere = "list") Then
		Set emailListCount = dbConnection.Execute("SELECT count(memberID) as emailCount FROM emailer_tblMembers WHERE memberOf = " & whichID & " AND memberDisabled = False")
		emailCount = emailListCount("emailCount")
		Set emailListCount = Nothing
		Set emailList = dbConnection.Execute("SELECT memberName, memberEmail FROM emailer_tblMembers WHERE memberOf = " & whichID & " AND memberDisabled = False")
	ElseIf (pullFromWhere = "member") Then
		Response.Write("Send to member: " & whichID)
	ElseIf (pullFromWhere = "reg") Then
		Response.Write("Send to reg: " & whichID)
	ElseIf (pullFromWhere = "member-group") Then
		Response.Write("Send to group of members: " & whichID)
	ElseIf (pullFromWhere = "list-group") Then
		Response.Write("Send to group of lists: " & whichID)
	End If

	If (emailCount = 0) Then
		Exit Function
	End If

	ReDim emailNames(emailCount)
	ReDim emailAddresses(emailCount)

	listPos = 0
	Do Until emailList.EOF
		emailNames(listPos) = emailList("memberName")
		emailAddresses(listPos) = emailList("memberEmail")
		emailList.MoveNext
		listPos = listPos + 1
	Loop

	massSend emailNames, emailAddresses, "My Test", "My Message"

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function getEmail(addressData)
	Dim emailParts(2)
	addressData = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(addressData,"'","&#39;"),vbTab," "),">",""),"<",""),"&gt;",""),"&lt;",""),"""",""))
	If (addressData = "") Then
		emailParts(0) = "skip"
		emailParts(1) = "skip"
		getEmail = emailParts
		Exit Function
	End If
	If (InStr(addressData,"@") <> InStrRev(addressData,"@")) Then
		emailParts(0) = "error"
		emailParts(1) = "error"
		getEmail = emailParts
		Exit Function
	End If
	If ((InStr(addressData,"@") > 0) And (InStr(addressData," ") > 0)) Then
		'Name without quotes and email address.
		addressData = Trim(Replace(Replace(addressData,">",""),"<",""))
		splitEmail = split(addressData," ")
		If (InStrRev(addressData,"@") > InStrRev(addressData," ")) Then
			'Name before email.
			For X = 0 To Ubound(splitEmail)-1
				emailParts(0) = emailParts(0) & splitEmail(X) & " "
			Next
			emailParts(1) = splitEmail(Ubound(splitEmail))
		ElseIf (InStr(addressData,"@") < InStr(addressData," ")) Then
			'Name after email.
			For X = 1 To Ubound(splitEmail)
				emailParts(0) = emailParts(0) & splitEmail(X) & " "
			Next
			emailParts(1) = splitEmail(0)
		Else
			'Names on both sides...?
			For Each emailPart In splitEmail
				If (InStr(emailPart,"@") > 0) Then
					emailParts(1) = emailPart
				Else
					emailParts(0) = emailParts(0) & emailPart & " "
				End If
			Next
			emailParts(0) = Left(emailParts(0),Len(emailParts(0))-1)
		End If
		'emailParts(0) = Left(addressData,InStr(addressData," "))
		'emailParts(1) = addressData
	ElseIf ((InStr(addressData,"@") > 0) And (InStr(addressData," ") = 0)) Then
		'Address Only
		emailParts(0) = ""
		emailParts(1) = addressData
	Else
		'Return Error
		emailParts(0) = "error"
		emailParts(1) = "error"
		getEmail = emailParts
		Exit Function
	End If
	emailParts(0) = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(emailParts(0),"'","&#39;"),vbTab," "),">",""),"<",""),"&gt;",""),"&lt;",""),"""",""))
	emailParts(1) = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(emailParts(1),"'","&#39;"),vbTab," "),">",""),"<",""),"&gt;",""),"&lt;",""),"""",""))
	getEmail = emailParts
End Function

outgoingServers = array("mail.sisna.com","maila.myexcel.com","mailb.myexcel.com")

Function sendMassEmail(emailFrom, emailAddresses, messageSubject, messageBody)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	addyParts = getEmail(emailFrom)
	fromName = addyParts(0)
	fromAddress = addyParts(1)
	
	'Licence Key To Remove Demo Status
'	ezMailer.LicenseKey = "LIDONG/S1dI500R1AX60C0Rb100"	ver 5
'	ezMailer.LicenseKey = "Louise Priest (Single Developer)/0010630410721500AB30"  ver 6
	'Response.Buffer = True
	Response.Write("<script>" & vbNewline)
	Response.Write("var timer = null;" & vbNewline)
	Response.Write("timer = setTimeout('scrollpage()', 10);" & vbNewline)
	Response.Write("" & vbNewline)
	Response.Write("function stop(){" & vbNewline)
	Response.Write("	window.scrollBy(0,600);" & vbNewline)
	Response.Write("	clearTimeout(timer);" & vbNewline)
	Response.Write("}" & vbNewline)
	Response.Write("" & vbNewline)
	Response.Write("function scrollpage() {" & vbNewline)
	Response.Write("	window.scrollBy(0,600);" & vbNewline)
	Response.Write("	timer = setTimeout('scrollpage()', 10);" & vbNewline)
	Response.Write("}" & vbNewline)
	Response.Write("" & vbNewline)
	Response.Write("window.onload = stop;" & vbNewline)
	Response.Write("</script>" & vbNewline)
	
	Response.Write("Sending to...<br />")
	Response.Flush
	
'	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""600"" align=""center"" class=""tablebg"">")
'	Response.Write("<tr><td>")

'	sendHTML = sendHTML & ("<big><b>Sent to...</b></big><br><br>")
'	sendHTML = sendHTML & ("<table border=""1"" cellpadding=""0"" cellspacing=""0"" width=""100%"">")

	maxPerBatch = 25
	Dim sendThisBatch() : ReDim sendThisBatch(maxPerBatch)

	emailCount = 0
	thisBatch = 0
	sentBatch = 0
	
	failedCount = 0
	
	sentCount = 0
	skipCount = 0
	emailAddresses = split(Trim(emailAddresses), vbNewline)
	
	totalToSend = Ubound(emailAddresses) + 1
	totalBatches = Fix(totalToSend / maxPerBatch)
	totalInLastBatch = (totalToSend Mod maxPerBatch)
	If (totalInLastBatch > 0) Then
		totalBatches = totalBatches + 1
	End If
	
	Response.Write("Sending " & maxPerBatch & " per batch.")
	Response.Write("There are " & totalToSend & " addresses to send to.<br />")
	Response.Write("There are " & totalBatches & " batches.<br />")
	Response.Write("There are " & totalInLastBatch & " addresses in the last batch.<br />")
	Response.Flush
	
	For Each addy In emailAddresses
		If (Len(addy) > 0) Then
			If (thisBatch = maxPerBatch) Then
				Response.Write("This batch is being sent...")
				Response.Flush
				
				lsReturn = SendBatchMessage(outgoingMailServer, fromName, fromAddress, sendThisBatch, messageSubject, messageBody)
								
				If (lsReturn = 0) Then
					Response.Write("<font color=""green"">Success</font><br />")
				Else
					Response.Write("<font color=""red"">Failed</font><br />")
					Response.Write("&nbsp;&nbsp;&nbsp;Error: " & lsReturn & "<br />")
				End If
	
				thisBatch = 0
				For X = 0 To maxPerBatch-1
					dbConnection.Execute("INSERT INTO emailer_tblSent ( dSent, sEmail, iReturn ) VALUES ( #" & date & " " & time & "#, '" & sendThisBatch(X) & "', " & lsReturn & " )")
					sendThisBatch(X) = ""
				Next
				sentBatch = sentBatch + 1
				
				Response.Write("<br />Creating a new batch.<br />")
			End If
	
			thisBatch = thisBatch + 1
			emailCount = emailCount + 1
			
			Response.Write("" & addy & " (" & sentBatch & "/" & thisBatch & "/" & emailCount & ")")
			If (thisBatch < maxPerBatch) And (Fix(totalToSend) > Fix(emailCount)) Then
				Response.Write(", ")
			Else
				Response.Write("<br />")
			End If
			sendThisBatch(thisBatch-1) = addy
			Response.Flush
		End If
	Next

	Response.Write("This batch is being sent...")
	Response.Flush
	
	lsReturn = SendBatchMessage(outgoingMailServer, fromName, fromAddress, sendThisBatch, messageSubject, messageBody)
					
	If (lsReturn = 0) Then
		Response.Write("<font color=""green"">Success</font><br />")
	Else
		Response.Write("<font color=""red"">Failed</font><br />")
		Response.Write("&nbsp;&nbsp;&nbsp;Error: " & lsReturn & "<br />")
	End If

	thisBatch = 0
	For X = 0 To maxPerBatch-1
'		Response.Write("<br />INSERT INTO emailer_tblSent ( dSent, sEmail, iReturn ) VALUES ( #" & date & " " & time & "#, '" & sendThisBatch(X) & "', " & lsReturn & " )")
		dbConnection.Execute("INSERT INTO emailer_tblSent ( dSent, sEmail, iReturn ) VALUES ( #" & date & " " & time & "#, '" & sendThisBatch(X) & "', " & lsReturn & " )")
		sendThisBatch(X) = ""
	Next
	sentBatch = sentBatch + 1
	
	Response.Write("<br />Sent with " & sentBatch & " batches.<br />")
	Response.Write("Sent to " & emailCount & " addresses.<br />")
	Response.Flush
	
	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function SendBatchMessageOld2(outgoingMailServer, fromName, fromAddress, sendThisBatch, messageSubject, messageBody)
	Set ezMailer = Server.CreateObject("EasyMail.SMTP.5")

	ezMailer.LicenseKey = "LIDONG/S1dI500R1AX60C0Rb100"

	ezMailer.MailServer = outgoingMailServer

	ezMailer.From = fromName
	ezMailer.FromAddr = fromAddress
	ezMailer.AddCustomHeader "Return-Path", "<" & fromAddress & ">"

	ezMailer.AddRecipient "Sundance Media Group Members", fromAddress, 1 ' 1 - To, 2 - CC, 3 - BCC
	ezMailer.AddRecipient "List Mailer", "syntax-cart@sisna.com", 3 ' 1 - To, 2 - CC, 3 - BCC
'	ezMailer.AddRecipient "Sundance Media Group Members", "smg@modprobe.com", 1 ' 1 - To, 2 - CC, 3 - BCC

	For Each recipient In sendThisBatch
		If (Len(recipient) > 0) Then
			'Response.Write("Adding Recipient: " & recipient & "<br />")
			ezMailer.AddRecipient "", recipient, 3
		End If
	Next

	ezMailer.Subject = messageSubject

	ezMailer.AutoWrap = 0
			
	ezMailer.BodyFormat = 1

	ezMailer.BodyText = messageBody
		
	SendBatchMessage = ezMailer.Send()
'	SendBatchMessage = 0
	
	Set ezMailer = Nothing
End Function

Function SendBatchMessage(outgoingMailServer, fromName, fromAddress, sendThisBatch, messageSubject, messageBody)
	Call CDOMailer("""" & fromName & """ <" & fromAddress & ">", """Sundance Media Group Members"" <" & fromAddress & ">", "", Join(sendThisBatch,",") & ",syntax-cart@sisna.com", messageSubject, messageBody, "html")
	SendBatchMessage = 0
End Function

	Function CDOMailer(mailFrom, mailTo, mailCC, mailBCC, mailSubject, mailBody, mailType)
		Set objSendMail = CreateObject("CDO.Message") 
		With objSendMail
			.Subject = mailSubject 
			.From = mailFrom
			.To = Replace(mailTo,";",",")
			.CC = Replace(mailCC,";",",")
			.BCC = Replace(mailBCC,";",",")
			If (mailType = "html") Then
				.HTMLBody = mailBody
			Else
				.TextBody = mailBody
			End If
			.Send()
		End With
		Set objSendMail = Nothing
	End Function
	
Function sendMassEmailOld(emailFrom, emailAddresses, messageSubject, messageBody)
	addyParts = getEmail(emailFrom)
	fromName = addyParts(0)
	fromAddress = addyParts(1)
	
	'Licence Key To Remove Demo Status
'	ezMailer.LicenseKey = "LIDONG/S1dI500R1AX60C0Rb100"	ver 5
'	ezMailer.LicenseKey = "Louise Priest (Single Developer)/0010630410721500AB30"  ver 6

	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""600"" align=""center"" class=""tablebg"">")
	Response.Write("<tr><td>")

	sendHTML = sendHTML & ("<big><b>Sent to...</b></big><br><br>")
	sendHTML = sendHTML & ("<table border=""1"" cellpadding=""0"" cellspacing=""0"" width=""100%"">")

	emailCount = 0
	failedCount = 0
	sentCount = 0
	skipCount = 0
	emailAddresses = split(emailAddresses, vbNewline)
	For Each addy In emailAddresses
		thisAddy = getEmail(addy)
		If (thisAddy(0) <> "skip") And (thisAddy(0) <> "error") Then
			Set ezMailer = Server.CreateObject("EasyMail.SMTP.5")
			ezMailer.LicenseKey = "LIDONG/S1dI500R1AX60C0Rb100"
		
			ezMailer.MailServer = outgoingMailServer

			ezMailer.From = fromName
			ezMailer.FromAddr = fromAddress
			ezMailer.AddCustomHeader "Return-Path", "support@vasst.com"

'			ezMailer.AddRecipient "Sundance Media Group Members", fromAddress, 1 ' 1 - To, 2 - CC, 3 - BCC
'			ezMailer.AddRecipient "Sundance Media Group Members", "smg@modprobe.com", 1 ' 1 - To, 2 - CC, 3 - BCC

			ezMailer.Subject = messageSubject

			ezMailer.AutoWrap = 0
			
			ezMailer.BodyFormat = 1

			ezMailer.BodyText = messageBody

			emailCount = emailCount + 1

			sendHTML = sendHTML & ("<tr>")
'			sendHTML = sendHTML & ("<td><small>" & encodeData(addy) & " -- " & thisAddy(0) & " -- " & thisAddy(1) & "</small></td>")
		End If
		If (thisAddy(0) = "skip") Then
			skipCount = skipCount + 1
		ElseIf (thisAddy(0) = "error") Then
			'Error, Don't send.
			failedCount = failedCount + 1
			sendHTML = sendHTML & ("<tr><td>" & addy & "</td><td><font color=""red"">Failed</font></td></tr>" & vbNewline)
			sendHTML = sendHTML & ("<tr><td colspan=""2"">Error: Not a valid address.</td></tr>")
		ElseIf (thisAddy(0) = "") And (thisAddy(1) <> "") Then
			'Email Address Only, Send.
			sentCount = sentCount + 1
			sendHTML = sendHTML & ("<td>&lt;" & thisAddy(1) & "&gt;</td><td><font color=""green"">Sent</font></td>" & vbNewline)
			ezMailer.AddRecipient "", thisAddy(1), 1 ' BCC
		Else
			'Email and Name, Send.
			sentCount = sentCount + 1
			sendHTML = sendHTML & ("<td>""" & thisAddy(0) & """ &lt;" & thisAddy(1) & "&gt;</td>" & vbNewline)
			ezMailer.AddRecipient thisAddy(0), thisAddy(1), 1 ' BCC
		End If
		If (thisAddy(0) <> "skip") And (thisAddy(0) <> "error") Then
			lsReturn = ezMailer.Send()
				
			If (lsReturn = 0) Then
				sendHTML = sendHTML & ("<td><font color=""green"">Sent</font></td>")
				allOkay = true
			Else
				sendHTML = sendHTML & ("<td><font color=""red"">Failed</font>")
				sendHTML = sendHTML & ("<tr><td colspan=""2"">Error: " & lsReturn & "</td></tr>")
				allOkay = false
			End If

			Set ezMailer = Nothing

			sendHTML = sendHTML & ("</tr>")
		End If
'		If ((emailCount Mod 10) = 0) Or (emailCount = Ubound(emailAddresses)+1) Then
'		End If
	Next
	sendHTML = sendHTML & ("</table>")

	sendHTML = sendHTML & ("<br>Of the <b>" & emailCount & "</b> reciepiants, <font color=""green""><b>" & sentCount & "</b></font> were sent, <font color=""red""><b>" & failedCount & "</b></font> failed to send.<br>" & vbNewline)

	If (allOkay) Then
		Response.Write(sendHTML)
	Else
		Response.Write("<center>There was an error in the sending batch.</center>" & vbNewline)
	End If
	Response.Write("</td>")
	Response.Write("</tr>")
	Response.Write("</table>")
End Function

%>
<script>
  function toggleList(section, whichList) {
  	theList = document.getElementById(section + "_" + whichList);
  	if (theList.style.display == "none") {
  		theList.style.display = "block";
  	}
  	else {
  		theList.style.display = "none";
  	}
  }

  function toggleChecks(fromID, section, whichList) {
		var x = 0;
		var loopDone = false;
		do {
			x++;
			toCheck = document.getElementById(section + "_" + whichList + "_" + x);
			if (toCheck == null) {
				loopDone = true;
			}
			else {
				toCheck.checked = fromID.checked;
				//if (fromID.checked) {
				//	toCheck.checked = true;
				//}
				//else {
				//	toCheck.checked = true;
				//}
			}
		} while(!loopDone);
  }
        </script>
<%

Function printSendChooser
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Response.Write("<center>Please select the lists you want to send email to.  Click on that list's name to see the people in that list.</center>" & vbNewline)
	Response.Write("<br>" & vbNewline)

	Response.Write("<center><input type=""submit"" name=""clearEmails"" value=""Clear""><input type=""submit"" name=""sendToEmail"" value=""Next ->""></center><br>" & vbNewline)

	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbNewline)
	Response.Write("<tr><td width=""50%"" valign=""top"">" & vbNewline)

	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" class=""tablebg"">" & vbNewline)
	Response.Write("<tr><td class=""heading"">Mailing Lists</td></tr>" & vbNewline)

	Set mailingLists = dbConnection.Execute("SELECT listID, listName FROM emailer_tblMailingLists ORDER BY listName" & vbNewline)
	Do Until mailingLists.EOF
		Set memberCount = dbConnection.Execute("SELECT count(memberID) as memberCount FROM emailer_tblMembers WHERE memberDisabled = False AND memberOf = " & mailingLists("listID") & "")
		memberCount = memberCount("memberCount")
		Response.Write("<tr>" & vbNewline)
		Response.Write("<td>" & vbNewline)

		leaveOpen = false
		checkList = ""
		listType = ""
		listID = ""
		If (Request.Form("clearEmails") <> "Clear") Then
			If (InStr(Request.Form("emailTo"),"[" & mailingLists("listID") & "]") > 0) Then
				checkList = " CHECKED"
				leaveOpen = true
			End If
		End If

 		Set mailingMembers = dbConnection.Execute("SELECT memberID, memberName, memberEmail FROM emailer_tblMembers WHERE memberOf = " & mailingLists("listID") & " AND memberDisabled = False ORDER BY memberEmail")
		strEmails = "[" & mailingLists("listID") & "]"
		Do Until mailingMembers.EOF
			strEmails = strEmails & mailingMembers("memberEmail") & ", "
			mailingMembers.MoveNext
		Loop
		strEmails = Left(strEmails,Len(strEmails)-2)
		Set mailingMembers = Nothing

		Response.Write("<input type=""checkbox"" name=""emailTo"" value=""" & strEmails & """ " & checkList & ">")
'		Response.Write("<a href=""javascript:void(1);"" onClick=""toggleList('list','" & mailingLists("listID") & "')"" onMouseOver=""window.status='Expand/Collapse this list.';return true;"" onMouseOut=""window.status='';return true;"">")
		Response.Write(mailingLists("listName") & " (" & memberCount & ")")
'		Response.Write("</a>" & vbNewline)
		Response.Write("</td>" & vbNewline)
		Response.Write("</tr>" & vbNewline)

'		Response.Write("<tr>" & vbNewline)
'		Response.Write("<td>" & vbNewline)
'		hideStyle = ""
'		If Not (leaveOpen) Then
'			hideStyle = "style=""display: none;"""
'		End If
'		Response.Write("<span id=""list_" & mailingLists("listID") & """ " & hideStyle & ">" & vbNewline)
'
' 		Set mailingMembers = dbConnection.Execute("SELECT memberID, memberName, memberEmail FROM emailer_tblMembers WHERE memberOf = " & mailingLists("listID") & " ORDER BY memberName")
'		listPos = 0
'		Do Until mailingMembers.EOF
'			listPos = listPos + 1
'			checkThis = ""
'			If (Request.Form("clearEmails") <> "Clear") Then
'				If ((checkList = " CHECKED") Or (Cstr(mailingMembers("memberID")) = Cstr(listID)) Or (InStr(Request.Form("emailTo"),"[{" & mailingMembers("memberName") & "}{" &  mailingMembers("memberEmail") & "}]") > 0)) Then
'					checkThis = " CHECKED"
'				Else
'					checkThis = ""
'				End If
'			End If
'			Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""checkbox"" name=""emailTo"" id=""list_" & mailingLists("listID") & "_" & listPos & """ value=""[{" & mailingMembers("memberName") & "}{" &  mailingMembers("memberEmail") & "}]""" & checkThis & ">")
'			Response.Write("<span class=""smallfont"">" & mailingMembers("memberName") & " &lt;" &  mailingMembers("memberEmail") & "&gt;</span><br>" & vbNewline)
'			mailingMembers.MoveNext
'		Loop
'		Set mailingMembers = Nothing
'
'		Response.Write("</span>" & vbNewline)
'		Response.Write("</td>" & vbNewline)
'		Response.Write("</tr>" & vbNewline)
		mailingLists.MoveNext
	Loop
	Set mailingLists = Nothing

	Response.Write("</table>" & vbNewline)

	Response.Write("</td>" & vbNewline)
	Response.Write("<td>&nbsp;&nbsp;</td>" & vbNewline)
	Response.Write("<td width=""50%"" valign=""top"">" & vbNewline)

	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" class=""tablebg"">" & vbNewline)
	Response.Write("<tr><td class=""heading"">Registration Lists</td></tr>" & vbNewline)

	Response.Write("<tr><td class=""heading"">" & vbNewline)
	Response.Write("Include Unpaid Accounts: " & vbNewline)
	Response.Write("<select size=""1"" name=""includeUnPaid"">" & vbNewline)
	If (Request.Form("includeUnPaid") = "false") Then
		trueSelected = ""
		falseSelected = " SELECTED"
		unpaidQuery = "isPaid = True AND"
	Else
		trueSelected = " SELECTED"
		falseSelected = ""
		unpaidQuery = ""
	End If
	Response.Write("<option value=""true""" & trueSelected & ">Yes</option>" & vbNewline)
	Response.Write("<option value=""false""" & falseSelected & ">No</option>" & vbNewline)
	Response.Write("</select>" & vbNewline)
	Response.Write("<input type=""submit"" name=""applyFilter"" value=""Filter"">" & vbNewline)
	Response.Write("</td></tr>" & vbNewline)

	Response.Write("<tr><td class=""heading"" align=""center""><img src=""paid.gif""> - Paid &nbsp;&nbsp; <img src=""unpaid.gif""> - Unpaid</td></tr>" & vbNewline)

	Set seminarLists = dbConnection.Execute("SELECT DISTINCT strSeminarName FROM tblCustomers WHERE " & unpaidQuery & " isDeleted = False")
	Do Until seminarLists.EOF
		Set lastDateOnTour = dbConnection.Execute("SELECT dateSeminarDate FROM tblCustomers WHERE " & unpaidQuery & " isDeleted = False AND strSeminarName = '" & seminarLists("strSeminarName") & "' ORDER BY dateSeminarDate DESC")
		closeThis = ""
		If Not (lastDateOnTour.EOF) Then
			If (CDate(lastDateOnTour("dateSeminarDate")) < Now) Then
				closeThis = "style=""display: none;"""
			End If
		End If
		Set lastDateOnTour = Nothing

		Set memberCountTour = dbConnection.Execute("SELECT count(numCustID) as memberCountTour FROM tblCustomers WHERE " & unpaidQuery & " isDeleted = False AND strSeminarName = '" & seminarLists("strSeminarName") & "'")
		memberCountTour = memberCountTour("memberCountTour")
		Response.Write("<tr>" & vbNewline)
		Response.Write("<td>" & vbNewline)

		checkTour = ""
		If (Request.Form("clearEmails") <> "Clear") Then
			If (InStr(Request.Form("tourChecked"),"tour_" & seminarLists("strSeminarName") & "") > 0) Then
				checkTour = " CHECKED"
			End If
		End If

		Response.Write("<input type=""checkbox"" name=""tourChecked"" value=""tour_" & seminarLists("strSeminarName") & """ onpropertychange=""toggleChecks(this, 'tour','" & seminarLists("strSeminarName") & "');""" & checkTour & ">" & vbNewline)
		Response.Write("<a href=""javascript:void(1);"" onClick=""toggleList('tour','" & seminarLists("strSeminarName") & "')"" onMouseOver=""window.status='Expand/Collapse this list.';return true;"" onMouseOut=""window.status='';return true;"">" & seminarLists("strSeminarName") & " (" & memberCountTour & ")</a>" & vbNewline)
		Response.Write("</td>" & vbNewline)
		Response.Write("</tr>" & vbNewline)
		Response.Write("<tr>" & vbNewline)
		Response.Write("<td>" & vbNewline)
		Response.Write("<span id=""tour_" & seminarLists("strSeminarName") & """" & closeThis & ">" & vbNewline)

		Set toursInList = dbConnection.Execute("SELECT strSeminarCity, dateSeminarDate, count(numCustID) as memberCountSeminar FROM tblCustomers WHERE " & unpaidQuery & " isDeleted = False AND strSeminarName = '" & seminarLists("strSeminarName") & "' GROUP BY dateSeminarDate, strSeminarCity ORDER BY dateSeminarDate, strSeminarCity")
		seminarPos = 0
		Do Until toursInList.EOF
			seminarPos = seminarPos + 1

			checkSeminar = ""
			leaveOpen = false
			If (Request.Form("clearEmails") <> "Clear") Then
				If (InStr(Request.Form("seminarChecked"),"seminar_" & seminarLists("strSeminarName") & "_" & toursInList("strSeminarCity") & "_" & toursInList("dateSeminarDate") & "") > 0) Then
					checkSeminar = " CHECKED"
					leaveOpen = true
				End If
			End If
			
			Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""checkbox"" name=""seminarChecked"" value=""seminar_" & seminarLists("strSeminarName") & "_" & toursInList("strSeminarCity") & "_" & toursInList("dateSeminarDate") & """ id=""tour_" & seminarLists("strSeminarName") & "_" & seminarPos & """ onpropertychange=""toggleChecks(this, 'seminar','" & seminarLists("strSeminarName") & "_" & toursInList("strSeminarCity") & "_" & toursInList("dateSeminarDate") & "');""" & checkSeminar & ">" & vbNewline)
			Response.Write("<a href=""javascript:void(1);"" onClick=""toggleList('seminar','" & seminarLists("strSeminarName") & "_" & toursInList("strSeminarCity") & "_" & toursInList("dateSeminarDate") & "')"" onMouseOver=""window.status='Expand/Collapse this list.';return true;"" onMouseOut=""window.status='';return true;"">" & toursInList("strSeminarCity") & " " & toursInList("dateSeminarDate") & " (" & toursInList("memberCountSeminar") & ")" & "</a><br>" & vbNewline)

			hideStyle = ""
			If Not (leaveOpen) Then
				hideStyle = "style=""display: none;"""
			End If

			Response.Write("<span id=""seminar_" & seminarLists("strSeminarName") & "_" & toursInList("strSeminarCity") & "_" & toursInList("dateSeminarDate") & """ " & hideStyle & ">" & vbNewline)

   			Set membersInList = dbConnection.Execute("SELECT strFirstName, strLastName, strEmail, isPaid FROM tblCustomers WHERE " & unpaidQuery & " isDeleted = False AND strSeminarName = '" & seminarLists("strSeminarName") & "' AND dateSeminarDate = #" & toursInList("dateSeminarDate") & "# AND strSeminarCity = '" & toursInList("strSeminarCity") & "' ORDER BY strEmail")
			listPos = 0
			Do Until membersInList.EOF
				listPos = listPos + 1
				thisChecked = ""
				If (Request.Form("clearEmails") <> "Clear") Then
					If (InStr(Request.Form("emailTo"),"" &  membersInList("strEmail") & "") > 0) Then
						thisChecked = " CHECKED"
					Else
						thisChecked = ""
					End If
				End If
				If (membersInList("isPaid") = "True") Then
					thisPaid = "&nbsp;<img src=""paid.gif"">"
				Else
					thisPaid = "&nbsp;<img src=""unpaid.gif"">"
				End If
'				Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input height=""5"" width=""5"" type=""checkbox"" name=""emailTo"" id=""seminar_" & seminarLists("strSeminarName") & "_" & toursInList("strSeminarCity") & "_" & toursInList("dateSeminarDate") & "_" & listPos & """ value=""[{" & membersInList("strFirstName") & " " & membersInList("strLastName") & "}{" &  membersInList("strEmail") & "}]"" " & thisChecked & ">" & vbNewline)
				Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input height=""5"" width=""5"" type=""checkbox"" name=""emailTo"" id=""seminar_" & seminarLists("strSeminarName") & "_" & toursInList("strSeminarCity") & "_" & toursInList("dateSeminarDate") & "_" & listPos & """ value=""" &  membersInList("strEmail") & """ " & thisChecked & ">" & vbNewline)
				Response.Write("<span class=""smallfont"">" & membersInList("strFirstName") & " " & membersInList("strLastName") & " &lt;" &  membersInList("strEmail") & "&gt;" & thisPaid & "</span><br>" & vbNewline)			
				membersInList.MoveNext
			Loop
			Set membersInList = Nothing

			Response.Write("</span>" & vbNewline)
			toursInList.MoveNext
		Loop
		Set toursInList = Nothing

		Response.Write("</span>" & vbNewline)
		Response.Write("</td>" & vbNewline)
		Response.Write("</tr>" & vbNewline)
		seminarLists.MoveNext
	Loop

	Set seminarLists = Nothing

	Response.Write("</table>" & vbNewline)

	Response.Write("</td>" & vbNewline)
	Response.Write("</tr>" & vbNewline)
	Response.Write("</table>" & vbNewline)

	Response.Write("<center><input type=""submit"" name=""clearEmails"" value=""Clear""><input type=""submit"" name=""sendToEmail"" value=""Next ->""></center>" & vbNewline)

	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function
%>
<script>
  function updateMailType() {
    if (document.getElementById("_mailType").value == "plain") {
      document.getElementById("_message").wrap = "hard";
      document.getElementById("_message").className = "tafixed";
    }
    else {
      document.getElementById("_message").wrap = "soft";
      document.getElementById("_message").className = "tanormal";
    }
  }

  function resetTarget(newTarget) {
    document.frmMenu.target = newTarget;
  }

  function resetWrap(newWrap) {
    document.frmMenu.wrap = newWrap;
  }
  
  function pushToForm() {
    document.frmMenu._message.value = window.richedit.getHTML();
  }
  
  function clearEditor() {
    window.richedit.setHTML("");
  }
  
        </script>
<%
Function printEmailEditor(emailAddresses)
	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbNewline)
	Response.Write("<tr><td width=""1"" valign=""top"">" & vbNewline)

	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" class=""tablebg"">" & vbNewline)
	Response.Write("<tr><td class=""heading"">To:</td></tr>" & vbNewline)
	Response.Write("<tr><td>" & vbNewline)
	Response.Write("<textarea cols=""45"" rows=""30"" name=""_to"" wrap=""off"" CLASS=""tabox"">")

	Set addr = Server.CreateObject("Scripting.Dictionary")

	emailAddresses = Split(emailAddresses,", ")
	For Each address In emailAddresses
		addyPos = addyPos + 1
		address = Lcase(address)
		If (InStr(address,"[") < InStr(address,"]")) Then
			address = Right(address,Len(address)-InStr(address,"]"))
		End If
		If (addr.Item(address) <> "1") Then
			Response.Write(address)
			If (Ubound(emailAddresses)+1 <> addyPos) Then
				Response.Write(vbNewline)
			End If
			addr.Item(address) = "1"
		End If
	Next

	Response.Write("</textarea>" & vbNewline)
	Response.Write("</td></tr>" & vbNewline)
	Response.Write("</table>" & vbNewline)
	
	Response.Write("</td>" & vbNewline)
	Response.Write("<td>&nbsp;&nbsp;</td>" & vbNewline)
	Response.Write("<td width=""100%"" valign=""top"">" & vbNewline)

	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" class=""tablebg"">" & vbNewline)
	
	Response.Write("<tr><td class=""heading"">")
	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr>")
	Response.Write("<td width=""75"" align=""right"">Name:&nbsp;&nbsp;</td>")
	Response.Write("<td width=""*""><input type=""input"" name=""_fromName"" size=""70"" class=""myInput"" value=""Sundance Media Group""></td>")
	Response.Write("</tr></table>")
	Response.Write("</td></tr>" & vbNewline)

	Response.Write("<tr><td class=""heading"">")
	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr>")
	Response.Write("<td width=""75"" align=""right"">From:&nbsp;&nbsp;</td>")
	Response.Write("<td width=""*""><input type=""input"" name=""_from"" size=""70"" class=""myInput"" value=""" & fromAddress & """></td>")
	Response.Write("</tr></table>")
	Response.Write("</td></tr>" & vbNewline)

	Response.Write("<tr><td class=""heading"">")
	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr>")
	Response.Write("<td width=""75"" align=""right"">Subject:&nbsp;&nbsp;</td>")
	Response.Write("<td width=""*""><input type=""input"" name=""_subject"" size=""70"" class=""myInput""></td>")
	Response.Write("</tr></table>")
	Response.Write("</td></tr>" & vbNewline)
	
'	Response.Write("<tr><td class=""heading"">")
'	Response.Write("<script> setTimeout(""updateMailType()"",1); </script>")
'	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr>")
'	Response.Write("<td width=""75"" align=""right"">Mail Type:&nbsp;&nbsp;</td>")
'	Response.Write("<td width=""*""><select name=""_mailType"" size=""1"" onChange=""updateMailType()"" class=""myInput"">")
'	Response.Write("<option value=""plain"">Plain Text</option>")
'	Response.Write("<option value=""html"">HTML</option")
'	Response.Write("</select></td>")
'	Response.Write("</tr></table>")
'	Response.Write("</td></tr>" & vbNewline)

	Response.Write("<tr><td class=""heading"">")
	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr>")
	Response.Write("<td width=""75"" align=""right"">Message:&nbsp;&nbsp;</td>")

	Response.Write("<td width=""*"">")
	%>
	<input type="hidden" name="_message" value="">
<div id="cdiv" style='position:relative; left:0px; top:0px; height:500px; width:603px;'>
<IFRAME SRC="richedit/richedit.asp" id='richedit' style='margin: 0px; visibility: visible; position: absolute; left: 0px; top: 0px; height=100%; width=100%'></IFRAME>
</div>
    <%
'	Response.Write("<textarea cols=""80"" rows=""20"" name=""_message"" wrap=""hard"" class=""tafixed""></textarea>")
	Response.Write("</td>")
	
	Response.Write("</tr></table>")
	Response.Write("</td></tr>" & vbNewline)

	Response.Write("<tr><td align=""center"">" & vbNewline)
	Response.Write("<input type=""submit"" name=""_function"" value=""Preview"" onclick=""resetTarget('_blank');pushToForm();"" tabindex=""1"">" & vbNewline)
	Response.Write("<input type=""submit"" name=""_function"" value=""Send"" onclick=""resetTarget('_self');pushToForm();"" tabindex=""2"">" & vbNewline)
	Response.Write("<input type=""reset"" name=""_function"" value=""Reset"" onclick=""resetTarget('_self');clearEditor();"" tabindex=""3"">" & vbNewline)
	Response.Write("<input type=""submit"" name=""_function"" value=""Cancel"" onclick=""resetTarget('_self');"" tabindex=""4"">" & vbNewline)
	Response.Write("</td></tr>" & vbNewline)
	Response.Write("</table>" & vbNewline)

	Response.Write("</td></tr></table>" & vbNewline)
End Function

Function previewMail(pvFrom,pvTo,pvSubject,pvMessage)
	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""600"" align=""center"" class=""tablebg"">")
	Response.Write("<tr><td>")
	Response.Write("<center><b>Message Preview</b></center>")

	myNewline = "<br>" & vbNewline & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	
	toFormatted = ""
	emailAddresses = split(pvTo, vbNewline)
	For Each addy In emailAddresses
		thisAddy = getEmail(addy)
		If (thisAddy(0) = "skip") Then
		ElseIf (thisAddy(0) = "error") Then
			'Error, Don't send.
			toFormatted = toFormatted & ("<font color=""red"">" & addy & " (Invalid Address)</font>" & myNewline)
		Else
			'Email and Name, Send.
			sentCount = sentCount + 1
			toFormatted = toFormatted & ("""" & thisAddy(0) & """ &lt;" & thisAddy(1) & "&gt;" & myNewline)
		End If
	Next
	toFormatted = Left(toFormatted,Len(toFormatted)-Len(myNewline))

	pvNewline = "<br>" & vbNewline
	messageFormatted = pvMessage
	
	fromAddress = getEmail(pvFrom)
	Response.Write("<b>From:</b>    """ & fromAddress(0) & """ &lt;" & fromAddress(1) & "&gt;" & pvNewline)
	Response.Write("<b>To:</b>      " & toFormatted & pvNewline)
	Response.Write("<b>Subject:</b> " & pvSubject & pvNewline)
	Response.Write("<b>Date:</b>    " & date & " " & time & pvNewline)
	
	Response.Write("<blockquote style=""margin:0px;padding:5px;background:white;border:1px solid gray;"">")

	Response.Write(messageFormatted)

	Response.Write("</blockquote>" & pvNewline)

	If (pvMessageType = "plain") Then
		Response.Write("</pre>")
	End If

	Response.Write("<center><input type=""button"" value=""<- Back"" onClick=""self.close();""></center>")

	Response.Write("</td>")
	Response.Write("</tr>")
	Response.Write("</table>")

'	Response.Write("<input type=""hidden"" name=""_from"" value=""" & encodeData(pvFrom) & """>")
'	Response.Write("<input type=""hidden"" name=""_to"" value=""" & encodeData(pvTo) & """>")
'	Response.Write("<input type=""hidden"" name=""_subject"" value=""" & encodeData(pvSubject) & """>")
'	Response.Write("<input type=""hidden"" name=""_message"" value=""" & encodeData(pvMessage) & """>")
'	Response.Write("<input type=""hidden"" name=""_mailType"" value=""" & pvMessageType & """>")
End Function

Function printImportInterface
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""600"" align=""center"" class=""tablebg"">")
	Response.Write("<tr><td>")
	Response.Write("<center><b>Import Lists</b></center>")

	Response.Write("<b>Acceptable Formats For Emails:</b>")
	Response.Write("<ul>")
	Response.Write("first last email@address.com -&gt; ""first last"" &lt;email@address.com&gt;<br>")
	Response.Write("email@address.com first last -&gt; ""first last"" &lt;email@address.com&gt;<br>")
	Response.Write("first email@address.com last -&gt; ""first last"" &lt;email@address.com&gt;<br>")
	Response.Write("</ul>")
	Response.Write("<b>Unacceptable Formats For Emails (Will be skipped):</b>")
	Response.Write("<ul>")
	Response.Write("email@address.com -- Needs a name.<br>")
	Response.Write("first last -- Needs an email.<br>")
	Response.Write("</ul>")
	
	Set otherLists = dbConnection.Execute("SELECT * FROM emailer_tblMailingLists ORDER BY listName")
	If Not (otherLists.EOF) Then
		Response.Write("Select list to import into: <select size=""1"" name=""toWhichList"">")
		Do Until (otherLists.EOF)
			Response.Write("<option value=""" & otherLists("listID") & """>" & otherLists("listName") & "</option>")
			otherLists.MoveNext
		Loop
		Response.Write("</select><br>")
	End If
	Set otherLists = Nothing

	Response.Write("<textarea cols=""80"" rows=""20"" name=""importedAddresses""></textarea>")
	Response.Write("<center><input type=""submit"" value=""Import To List""></center>")
	Response.Write("</td>")
	Response.Write("</tr>")
	Response.Write("</table>")
	
	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function importAddresses(toList,theAddresses)
	If (isDbOpen) Then
		wasDbOpen = True
	Else
		openDatabase
	End If

	theAddresses = split(theAddresses,vbNewline)
	
	For Each addy In theAddresses
		thisEmail = getEmail(addy)
		If ((thisEmail(0) <> "") And (thisEmail(1) <> "") And (thisEmail(0) <> "error") And (thisEmail(0) <> "skip")) Then
'			Set dbEmail = dbConnection.Execute("SELECT memberName FROM emailer_tblMembers WHERE memberOf = " & toList & " AND memberEmail = '" & thisEmail(1) & "'")
			Set dbEmail = dbConnection.Execute("SELECT memberName FROM emailer_tblMembers WHERE memberEmail = '" & thisEmail(1) & "'")
			If (dbEmail.EOF) Then
				dbConnection.Execute("INSERT INTO emailer_tblMembers ( memberName, memberEmail, memberOf, memberDisabled ) VALUES ( '" & thisEmail(0) & "', '" & thisEmail(1) & "', " & toList & ", False )")
			End If
'			Response.Write("addy: """ & thisEmail(0) & """ [" & thisEmail(1) & "] <br>")
		End If
	Next
	
	If Not (wasDbOpen) Then
		closeDatabase
	End If
End Function

Function encodeData(strData)
	encodeData = strData
	encodeData = Replace(encodeData,"<","&lt;")
	encodeData = Replace(encodeData,">","&gt;")
	encodeData = Replace(encodeData,"""","&#34;")
End Function

Function decodeData(strData)
	decodeData = strData
	decodeData = Replace(decodeData,"&lt;","<")
	decodeData = Replace(decodeData,"&gt;",">")
	decodeData = Replace(decodeData,"&#34;","""")
End Function

If (Session("currentFunction") <> "") Then
	formFunction = Session("currentFunction")
	Session("currentFunction") = ""
Else
	formFunction = Request.Form("frmFunction")
End If

newFormFunction = formFunction
printMenu

Select Case (formFunction)
Case "listAdmin"
	Session("currentList") = ""
	printListAdmin
Case "addList"
	addList Request.Form("frmNewListName"), Request.Form("frmNewListNotes"), Request.Form("frmNewListSubscribeable")
	goBackToListAdmin = True
Case "renameList"
	renameList Request.Form("frmParam1"), Request.Form("frmParam2")
	goBackToListAdmin = True
Case "renameNotes"
	renameNotes Request.Form("frmParam1"), Request.Form("frmParam2")
	goBackToListAdmin = True
Case "deleteList"
	deleteList Request.Form("frmParam1")
	goBackToListAdmin = True
Case "enableSubscribe"
	enableSubscribe Request.Form("frmParam1")
	goBackToListAdmin = True
Case "disableSubscribe"
	disableSubscribe Request.Form("frmParam1")
	goBackToListAdmin = True
Case "memberAdmin"
	If (Request.Form("frmNewMemberName") <> "") Then
		Session("newMemberName") = newName
	End If
	If (Request.Form("frmNewMemberDisabled") <> "") Then
		Session("newMemberDisabled") = newDisabled
	End If
	If (Request.Form("frmNewMemberEmail") <> "") Then
		Session("newMemberEmail") = newEmail
	End If
	If (Request.Form("frmNewMemberNotes") <> "") Then
		Session("newMemberNotes") = newNotes
	End If

	If (Session("currentList") = "") And (Request.Form("frmParam1") = "")Then
		printListChooser
	ElseIf (Session("currentList") = "") Then
		whichList = Request.Form("frmParam1")
		Session("currentList") = whichList
		printMemberAdmin whichList
	Else
		whichList = Session("currentList")
		printMemberAdmin whichList
	End If
Case "addMember"
	If (isNumeric(Session("currentList"))) Then
		If (Request.Form("frmNewMemberDisabled") = "on") Then
			thisDisabled = True
		Else
			thisDisabled = False
		End If
		addMember Session("currentList"), Request.Form("frmNewMemberName"), Request.Form("frmNewMemberEmail"), Request.Form("frmNewMemberNotes"), Request.Form("frmNewMemberDisabled")
		goBackToMemberAdmin = True
	Else
		alertMessage = "You've been idle for too long, please reselect the member list you are editing, and try adding again."
		goBackToListAdmin = True
	End If
Case "renameMember"
	renameMember Request.Form("frmParam1"), Request.Form("frmParam2")
	goBackToMemberAdmin = True
Case "renameMemberEmail"
	renameMemberEmail Request.Form("frmParam1"), Request.Form("frmParam2")
	goBackToMemberAdmin = True
Case "renameMemberNotes"
	renameMemberNotes Request.Form("frmParam1"), Request.Form("frmParam2")
	goBackToMemberAdmin = True
Case "deleteMember"
	deleteMember Request.Form("frmParam1")
	goBackToMemberAdmin = True
Case "disableMember"
	disableMember Request.Form("frmParam1")
	goBackToMemberAdmin = True
Case "enableMember"
	enableMember Request.Form("frmParam1")
	goBackToMemberAdmin = True
Case "massDelete"
	massDelete Request.Form("frmParam1")
	goBackToMemberAdmin = True
Case "massMove"
	massMove Request.Form("frmParam1"), Request.Form("frmParam2")
	goBackToMemberAdmin = True
Case "sendEmail"
	Session("currentList") = ""
	If ((Request.Form("sendToEmail") <> "") And (Request.Form("emailTo") <> "")) Or (InStr(Request.Form("listChecked"),"list_") > 0) Then
		emailList = Request.Form("emailTo")
		'emailList = Left(Request.Form("emailTo"),Len(Request.Form("emailTo"))-1)
		'emailList = Right(emailList,Len(emailList)-1)
		'emailList = Split(emailList,"], [")
'		Response.Write("Addy: ")
		printEmailEditor emailList
	ElseIf (Request.Form("_function") = "Preview") Then
		previewMail """" & Request.Form("_fromName") & """ &lt;" & Request.Form("_from") & "&gt;", Request.Form("_to"), Request.Form("_subject"), Request.Form("_message")
	ElseIf (Request.Form("_function") = "Send") Then
		sendMassEmail """" & Request.Form("_fromName") & """ <" & Request.Form("_from") & ">", Request.Form("_to"), Request.Form("_subject"), Request.Form("_message")
	Else
		printSendChooser
	End If
Case "importAdmin"
	If (Request.Form("importedAddresses") = "") Then
		printImportInterface
	Else
		importAddresses Request.Form("toWhichList"), Request.Form("importedAddresses")
		goBackToListAdmin = True
	End If
Case "searchAdmin"
	printSearchAdmin
End Select

If (goBackToListAdmin) Then
	newFormFunction = "listAdmin"
	Session("currentFunction") = "listAdmin"
	selfRedirect = "emailer.asp"
ElseIf (goBackToMemberAdmin) Then
	newFormFunction = "memberAdmin"
	Session("currentFunction") = "memberAdmin"
	selfRedirect = "emailer.asp"
ElseIf (goBackToEmailer) Then
	newFormFunction = "sendEmail"
	Session("currentFunction") = "sendEmail"
	selfRedirect = "emailer.asp"
End If

%>
    <!--mstheme--></font></td>
  </tr>
</table><!--mstheme--><font face="Arial, Arial, Helvetica">
<input type="hidden" name="frmFunction" value="<%=newFormFunction%>" />
<input type="hidden" name="frmParam1" value="" />
<input type="hidden" name="frmParam2" value="" />
<input type="hidden" name="frmParam3" value="" />
<input type="hidden" name="frmParam4" value="" />
</form>
<%
'Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""tablebg""><tr><td>")
'For Each frm In Request.Form
'	Response.Write("<b>" & frm & "</b>=""" & encodeData(Request.Form(frm)) & """<br>")
'Next
'Response.Write("</td></tr></title>")
%>
<%
If (Session("alertMessage") <> "") Then
	Response.Write("<script> alert(""" & Session("alertMessage") & """); </script>")
	Session("alertMessage") = ""
End If
If (alertMessage <> "") Then
	Session("alertMessage") = alertMessage
'	Response.Write("<script> alert(""" & alertMessage & """); </script>")
End If
If (selfRedirect <> "") Then
	Response.Redirect selfRedirect
'	Response.Write("<script> self.location.href = self.location.href; </script>")
'	Response.Write("<script>frmMenu.submit();</script>")
End If
%>
        <!--mstheme--></font></td>
      </tr>
    </table><!--mstheme--><font face="Arial, Arial, Helvetica">
  <!--mstheme--></font></body>
</html>