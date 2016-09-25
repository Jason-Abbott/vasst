<!--#include file="functions.asp"-->
<% printHeader %>
<%
If Not (Request.Form("function") = "") Then
	Select Case lcase(Request.Form("function"))
	Case "save account info"
		If (checkPassword(Request.Form("currentPassword"))) Then
			setUserInfo getUserID, "username", Request.Form("username")
			setUserInfo getUserID, "realname", Request.Form("realName")
			setUserInfo getUserID, "email", Request.Form("emailAddress")
			If Not (Request.Form("accesslevel") = "") Then
				setUserInfo getUserID, "accesslevel", Request.Form("accessLevel")
			End If
			accountUpdated = "Account Updated."
		Else
			doNothing = true
		End If
	Case "change password"
		If (checkPassword(Request.Form("currentPassword"))) Then
			If (Request.Form("newPassword") = Request.Form("confirmPassword")) Then
				setUserInfo getUserID, "password", Request.Form("confirmPassword")
			Else
				Response.Write("<SCRIPT>alert('Passwords do not match, please try again.')</SCRIPT>")
			End If
			passwordChanged = "Password Changed."
		Else
			doNothing = true
		End IF
	Case "new user"
		If (checkPassword(Request.Form("currentPassword"))) Then
			If (checkAvailUser("newuser")) Then
				createUser("newuser")
				Response.Write("<SCRIPT> alert('Created a new user \'newuser\' with password \'newuser\', please change accordingly.') </SCRIPT>")
			Else
				Response.Write("<SCRIPT> alert('User \'newuser\' already exists, please rename, and then try again.') </SCRIPT>")
			End If
		Else
			doNothing = true
		End If
	Case "delete user"
		If (checkPassword(Request.Form("currentPassword"))) Then
			removeUser Request.Form("userMgrUserID")
		Else
			doNothing = true
		End If
	Case "set password"
		If (checkPassword(Request.Form("currentPassword"))) Then
			setUserInfo Request.Form("userMgrUserID"), "password", Request.Form("resetNewPassword")
		Else
			doNothing = true
		End If
	Case "load user"
		If Not (Request.Form("userMgrSelect") = "") Then
			userMgrUserID = Fix(Request.Form("userMgrSelect"))
			userMgrUsername = getUserInfo(userMgrUserID, "username")
			userMgrRealname = getUserInfo(userMgrUserID, "realname")
			userMgrEmail = getUserInfo(userMgrUserID, "email")
			userMgrAccessLevel = getUserInfo(userMgrUserID, "accesslevel")
			showUserMGR = true
		End If
	Case "save user"
		If (checkPassword(Request.Form("currentPassword"))) Then
			setUserInfo Request.Form("userMgrUserID"), "username", Request.Form("userMgrUsername")
			setUserInfo Request.Form("userMgrUserID"), "realname", Request.Form("userMgrRealname")
			setUserInfo Request.Form("userMgrUserID"), "email", Request.Form("userMgrEmail")
			setUserInfo Request.Form("userMgrUserID"), "accesslevel", Request.Form("userMgrAccessLevel")
			If Not (Request.Form("resetNewPassword") = "") Then
				setUserInfo Request.Form("userMgrUserID"), "password", Request.Form("resetNewPassword")
			End If
		Else
			userMgrUserID = Fix(Request.Form("userMgrUserID"))
			userMgrUsername = getUserInfo(userMgrUserID, "username")
			userMgrRealname = getUserInfo(userMgrUserID, "realname")
			userMgrEmail = getUserInfo(userMgrUserID, "email")
			userMgrAccessLevel = getUserInfo(userMgrUserID, "accesslevel")
			showUserMGR = true
			doNothing = true
		End If
	End Select
	
	If (doNothing) Then
		makeRed = "<FONT COLOR=red>"
		makeRedEnd = "</FONT>"
	End If
End If
'If (Request.Form("save") = "") Then
'	loadOptions
'Else
'	formOkay = True

'	If (Request.Form("") = "") Then
'		formOkay = False
'		= "CLASS="error""
'	End If

'	If (formOkay) Then
'		splitForm
'		saveOptions
'		Response.Write("<SCRIPT>top.location.href = "index.asp";</SCRIPT>")
'	End If
'End If
%>
<FORM METHOD=post>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td width="50" class="bar"><img src="blank.gif" width=50 height=1></td>
		<td width="200" colspan="2" class="bartitle" nowrap align="center">Required Password</td>
		<td width="100%" class="bar"><img src="blank.gif" width=100% height=1></td>
	</tr>
	<tr>
		<td colspan="2" align="right" nowrap>Current Password:</td>
		<td colspan="2" valign="top">&nbsp;<input type="password" name="currentPassword" size="20"><font size="1"><b><i><%=makeRed%>* Required for changes.<%=makeRedEnd%></i></b></font></td>
	</tr>
	<tr>
		<td colspan="4"><br></td>
	</tr>
	<tr>
		<td width="50" class="bar"><img src="blank.gif" width=50 height=1></td>
		<td width="200" colspan="2" class="bartitle" nowrap align="center">Your Account</td>
		<td width="100%" class="bar">&nbsp;<%=accountUpdated%></td>
	</tr>
	<tr>
		<td colspan="2" align="right" nowrap>Username:</td>
		<td colspan="2">
			<% If (getUserInfo(getUserID,"accesslevel") = "Admin") Then %>
				&nbsp;<input name="username" size="50" value="<%=getUserInfo(getUserID,"username")%>"><font size="1"><b><i>* Change requires re-logon.</i></b></font>
			<% Else %>
				&nbsp;<%=getUserInfo(getUserID,"username")%>
			<% End If %>
		</td>
	</tr>
	<tr>
		<td colspan="2" align="right" nowrap>Real Name:</td>
		<td colspan="2">&nbsp;<input name="realName" size="50" value="<%=getUserInfo(getUserID,"realname")%>"></td>
	</tr>
	<tr>
		<td colspan="2" align="right" nowrap>Email Address:</td>
		<td colspan="2">&nbsp;<input name="emailAddress" size="50" value="<%=getUserInfo(getUserID,"email")%>"></td>
	</tr>
	<tr>
		<td colspan="2" align="right" nowrap>Access Level:</td>
		<td colspan="2">
			<% If (getUserInfo(getUserID,"accesslevel") = "Admin") Then %>
			&nbsp;<select name="accessLevel" size="1" disabled>
				<option value="Admin" <% If getUserInfo(getUserID,"accesslevel") = "Admin" Then Response.Write("SELECTED") End If %>>Admin</option>
				<option value="Allow" <% If getUserInfo(getUserID,"accesslevel") = "Allow" Then Response.Write("SELECTED") End If %>>Allow</option>
				<option value="Deny" <% If getUserInfo(getUserID,"accesslevel") = "Deny" Then Response.Write("SELECTED") End If %>>Deny</option>
			</select><small> - WARNING: <A HREF=# onClick="alert('If you change your access level to something other than Admin, you will be logged out, you will lose admin abilities, and will have to have another admin give them back to you.  *BE CAREFUL*');document.forms[0].accessLevel.disabled = false;this.style.display = 'none';">CLICK HERE TO ALLOW</A></small><font size="1"><b><i>* Change requires re-logon.</i></b></font>
			<% Else %>
				&nbsp;<%=getUserInfo(getUserID,"accesslevel")%>
			<% End If %>
			</td>
	</tr>
	<tr>
		<td width="50" class="bar"><img src="blank.gif" width=50 height=1></td>
		<td width="200" colspan="2" class="bartitle" nowrap align="center"><input type="submit" name="function" size="60" value="Save Account Info"></td>
		<td width="100%" class="bar"><img src="blank.gif" width=100% height=1></td>
	</tr>
	<tr>
		<td colspan="4"><br></td>
	</tr>
	<tr>
		<td width="50" class="bar"><img src="blank.gif" width=50 height=1></td>
		<td width="200" colspan="2" class="bartitle" nowrap align="center">Change 
        Your Password</td>
		<td width="100%" class="bar">&nbsp;<%=passwordChanged%></td>
	</tr>
	<tr>
		<td colspan="2" align="right" nowrap>New Password:</td>
		<td colspan="2">&nbsp;<input type="password" name="newPassword" size="20" value="<%=blank%>"><font size="1"><b><i>* Change requires re-logon.</i></b></font></td>
	</tr>
	<tr>
		<td colspan="2" align="right" nowrap>Confirm Password:</td>
		<td colspan="2">&nbsp;<input type="password" name="confirmPassword" size="20" value="<%=blank%>"></td>
	</tr>
	<tr>
		<td width="50" class="bar"><img src="blank.gif" width=50 height=1></td>
		<td width="200" colspan="2" class="bartitle" nowrap align="center"><input type="submit" name="function" size="60" value="Change Password"></td>
		<td width="100%" class="bar"><img src="blank.gif" width=100% height=1></td>
	</tr>
	<tr>
		<td colspan="4"><br></td>
	</tr>
	<% If (getUserInfo(getUserID,"accesslevel") = "Admin") Then %>
	<tr>
		<td width="50" class="bar"><img src="blank.gif" width=50 height=1></td>
		<td width="200" colspan="2" class="bartitle" nowrap align="center">User Manager</td>
		<td width="100%" class="bar">
			<img src="blank.gif" width="50" height="1"><input type="submit" name="function" value="New User"><% If (showUserMGR) Then %><input type="submit" name="function" value="Delete User"><% End If %>
		</td>
	</tr>
	<tr>
		<td colspan="2" align="right" valign="top" nowrap>
			<select name="userMgrSelect" size=5 <% If (showUserMGR) Then Response.Write("disabled") End If %> onChange="document.getElementById('loaduser').click();">
				<% printUsers Request.Form("userMgrSelect") %>
			</select>
		</td>
		<td colspan="2">
			<% If (showUserMGR) Then %>
			<table border="0" cellpadding="0" cellspacing="0">
				<input type="hidden" name="userMgrUserID" value="<%=userMgrUserID%>">
				<tr>
					<td width="100" align="right">Username:</td>
					<td>&nbsp;<input name="userMgrUsername" size="50" value="<%=userMgrUsername%>"></td>
				</tr>
				<tr>
					<td width="100" align="right">Real Name:</td>
					<td>&nbsp;<input name="userMgrRealname" size="50" value="<%=userMgrRealname%>"></td>
				</tr>
				<tr>
					<td width="100" align="right">Email Address:</td>
					<td>&nbsp;<input name="userMgrEmail" size="50" value="<%=userMgrEmail%>"></td>
				</tr>
				<tr>
					<td width="100" align="right">Access Level:</td>
					<td>
						&nbsp;<select name="userMgrAccessLevel" size="1">
							<option value="Admin" <% If userMgrAccessLevel = "Admin" Then Response.Write("SELECTED") End If %>>Admin</option>
							<option value="Allow" <% If userMgrAccessLevel = "Allow" Then Response.Write("SELECTED") End If %>>Allow</option>
							<option value="Deny" <% If userMgrAccessLevel = "Deny" Then Response.Write("SELECTED") End If %>>Deny</option>
						</select>
					</td>
				</tr>
				<tr>
					<td width="100" align="right">New Password:</td>
					<td>
						<span id="chngPasswordButton">&nbsp;<input type="button" onclick="document.getElementById('chngPassword').style.display = ''; document.getElementById('chngPasswordButton').style.display = 'none';" value="Change Password"></span>
						<span id="chngPassword" style="display: none;">&nbsp;<input type="password" name="resetNewPassword" size="20"><input name="function" type="submit" value="Set Password"></span>
					</td>
			</table>
			<% End If %>
		</td>
	</tr>
	<tr>
		<td width="50" class="bar"><img src="blank.gif" width=50 height=1></td>
		<td width="200" colspan="2" class="bartitle" nowrap align="center">
			<input id="loaduser" type="submit" name="function" value="Load User">
		</td>
		<td width="100%" class="bar">
			<img src="blank.gif" width="50" height="1"><% If (showUserMGR) Then %><input type="submit" name="function" value="Save User"><input type="submit" name="function" value="Cancel Change"><% End If %>
		</td>
	</tr>
	<% End If %>
</table>
</form>