<!--#include file="functions.asp"-->
<% 
If Not (getUserInfo(getUserID,"accesslevel") = "Admin") Then 
	Response.Write "<CENTER><B>You do not have access to this section of the CMS, your access is Read-Only, contact your admin to get rights.</B></CENTER>"
	Response.End
End If
%>
<% printHeader %>
<FORM METHOD=post>
<%
'Response.Write("Form<BR>" & SplitForm & "Cookies<BR>" & SplitCookies)
If (Request.Form("save") = "") Then
	loadOptions
Else
	formOkay = True

'	If Not isDate(Request.Form("editSignupDate")) Then
'		formOkay = False
'		errorSignupDate = "CLASS='error'"
'	End If

'	If (Request.Form("editFirstName") = "") Or (Request.Form("editLastName") = "") Then
'		formOkay = False
'		errorName = "class='error'"
'	End If

'	If (Request.Form("") = "") Then
'		formOkay = False
'		= "CLASS='error'"
'	End If

'	If (Request.Form("") = "") Then
'		formOkay = False
'		= "CLASS='error'"
'	End If

'	If (Request.Form("") = "") Then
'		formOkay = False
'		= "CLASS='error'"
'	End If

'	If (Request.Form("") = "") Then
'		formOkay = False
'		= "CLASS='error'"
'	End If

	If (formOkay) Then
'		splitForm
		saveOptions
		Response.Write("<SCRIPT>top.location.href = 'index.asp?page=options.asp';</SCRIPT>")
	End If
End If
%>
<table border='0' cellpadding='0' cellspacing='0' width='100%'>
	<tr>
		<td width='50' class='bar'><img src='blank.gif' width=50 height=1></td>
		<td width='200' colspan='2' class='bartitle' nowrap align='center'>E-Mail Options</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Outgoing Mail Server:</td>
		<td colspan='2'>&nbsp;<input name='optionOutgoingMailServer' size='60' value='<%=outgoingMailServer%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Mail From Address:</td>
		<td colspan='2'>&nbsp;<input type="text" name="optionFromAddress" size="60" value='<%=fromAddress%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Registration Notification:</td>
		<td colspan='2'>&nbsp;<input type="text" name="optionRegistrationConfirmationAddress" size="60" value='<%=registrationConfirmationAddress%>'></td>
	</tr>
	<tr>
		<td width='50' class='bar'><img src='blank.gif' width=50 height=1></td>
		<td width='200' colspan='2' class='bartitle' nowrap align='center'>Look and Feel Options</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Display Size:</td>
		<td colspan='2'>&nbsp;<input name='optionDisplayWidth' size='6' value='<%=displayWidth%>'>X<input name='optionDisplayHeight' size='6' value='<%=displayHeight%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Title for Count:</td>
		<td colspan='2'>&nbsp;<input name='optionTitlePos' size='60' value='<%=titlePos%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Title for ID:</td>
		<td colspan='2'>&nbsp;<input name='optionTitleID' size='60' value='<%=titleID%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Title for Name:</td>
		<td colspan='2'>&nbsp;<input name='optionTitleName' size='60' value='<%=titleName%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Title for Date:</td>
		<td colspan='2'>&nbsp;<input name='optionTitleDate' size='60' value='<%=titleDate%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Title for City:</td>
		<td colspan='2'>&nbsp;<input name='optionTitleCity' size='60' value='<%=titleCity%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Title for Visible?:</td>
		<td colspan='2'>&nbsp;<input name='optionTitleIsVisible' size='60' value='<%=titleIsVisible%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Title for Remove?:</td>
		<td colspan='2'>&nbsp;<input name='optionTitleRemove' size='60' value='<%=titleRemove%>'></td>
	</tr>
	<tr>
		<td width='50' class='bar'><img src='blank.gif' width=50 height=1></td>
		<td width='200' colspan='2' class='bartitle' nowrap align='center'>Errors Options</td>
		<td width='100%' class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>Incorrect Date:</td>
		<td colspan='2'>&nbsp;<input name='optionErrIncorrectDate' size='60' value='<%=errIncorrectDate%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>No Seminars:</td>
		<td colspan='2'>&nbsp;<input name='optionErrNoEvents' size='60' value='<%=errNoEvents%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>&nbsp;&nbsp;&nbsp;No Removeable Seminars:</td>
		<td colspan='2'>&nbsp;<input name='optionErrNoEventsRemove' size='60' value='<%=errNoEventsRemove%>'></td>
	</tr>
	<tr>
		<td colspan='2' align='right' nowrap>No Seminars Active:</td>
		<td colspan='2'>&nbsp;<input name='optionErrNoSeminarsActive' size='60' value='<%=errNoSeminarsActive%>'></td>
	</tr>
	<tr>
		<td width='50' class='bar'><img src='blank.gif' width=50 height=1></td>
		<td width='200' colspan='2' class='bartitle' nowrap align='center'>
			<p align='center'>
			<input type='submit' name='save' value='Save'>
			<input type='reset' value='Reset'>
		</td>
		<td class='bar'><img src='blank.gif' width=100% height=1></td>
	</tr>
</table>
</FORM>
<%
printFooter 
%>