<!--#include file="functions.asp"-->
<% 
If Not (getUserInfo(getUserID,"accesslevel") = "Admin") Then 
	Response.Write "<CENTER><B>You do not have access to this section of the CMS, your access is Read-Only, contact your admin to get rights.</B></CENTER>"
	Response.End
End If
%>
<head>
<style>
TD { Font-Family: Courier New; }
INPUT { Font-Family: Courier New; }
</style>
<FORM METHOD=post>
<script>
function clearForm() {
//	document.forms[0].from.value = ""
	document.forms[0].to.value = ""
	document.forms[0].subject.value = ""
	document.forms[0].body.value = ""
}
</script>
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td align="right" valign="top">From:</td>
		<td><input type="text" name="from" size="50" value="<%=Request.Form("from")%>" />
	</tr>
	<tr>
		<td align="right" valign="top">To:</td>
		<td><input type="text" name="to" size="50" value="<%=Request.Form("to")%>" />
	</tr>
	<tr>
		<td align="right" valign="top">Subject:</td>
		<td><input type="text" name="subject" size="50" value="<%=Request.Form("subject")%>" />
	</tr>
	<tr>
		<td align="right" valign="top">Body:</td>
		<td><textarea name="body" cols="50" rows="5"><%=Request.Form("body")%></textarea></td>
	</tr>
	<tr>
		<td colspan="2" align="center">
			<input type="Submit" name="send" value="Send">
			<input type="Reset" name="reset" value="Reset">
			<input type="Button" name="clear" value="Clear" onClick="clearForm()">
	</tr>
	<tr>
		<td></td>
		<td>
			<br>
			<b>Use which mailer?</b><br>
			<input type="radio" name="mailerType" <% If (Request.Form("mailerType") = "CDONTS") Then Response.Write("CHECKED") End If %> value="CDONTS"> CDONTS - <% checkMailer("CDONTS") %><br />
			<input type="radio" name="mailerType" <% If (Request.Form("mailerType") = "ASPMail") Then Response.Write("CHECKED") End If %> value="ASPMail"> ASPMail - <% checkMailer("ASPMail") %><br />
			<input type="radio" name="mailerType" <% If (Request.Form("mailerType") = "EasyMail") Then Response.Write("CHECKED") End If %> value="EasyMail"> EasyMail - <% checkMailer("EasyMail") %><br />
			<% 
			If (Request.Form("submitted") = "true") And (Request.Form("mailerType") = "") Then 
				Response.Write("<font color=""red""><b>Please select a mailer type.</b></font><br />")
				stopMailer = true
			End If
			%>
			
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center">
			<br />
			<font color="red">
<%
If Not ((Request.Form("from") = "") Or (Request.Form("to") = "") Or (Request.Form("subject") = "") Or (Request.Form("body") = "")) Then
	If Not (stopMailer) And (Request.Form("submitted") = "true") Then
		If (Request.Form("mailerType") = "CDONTS") Then
			lsMailResult = CDONTSMail(Request.Form("from"), Request.Form("to"), Request.Form("subject"), Request.Form("body"))
		ElseIf (Request.Form("mailerType") = "ASPMail") Then
			lsMailResult = ASPMail(Request.Form("from"), Request.Form("to"), Request.Form("subject"), Request.Form("body"))
		ElseIf (Request.Form("mailerType") = "EasyMail") Then
			lsMailResult = EasyMail(Request.Form("from"), Request.Form("to"), Request.Form("subject"), Request.Form("body"))
		Else
			lsMailResult = "Unknown Mailer Type"
		End If
		Response.Write(lsMailResult)
	End If
End If
%>
			</font>
		</td>
	</tr>
</table>
<input type="hidden" name="submitted" value="true">
</form>
<%
Function checkMailer(mailerType)
	On Error Resume Next
	If (mailerType = "CDONTS") Then
		Set testMailer = Server.CreateObject("CDONTS.NewMail")
	ElseIf (mailerType = "ASPMail") Then
		Set testMailer = Server.CreateObject("ASPMail.ASPMailCtrl.1")
	ElseIf (mailerType = "EasyMail") Then
		Set testMailer = Server.CreateObject("EasyMail.SMTP.5")
	Else
		Response.Write("Unknown - Not configured")
		Set testMailer = Nothing
		Exit Function
	End If
	
	If Err <> 0 Then
		Response.Write("<font color=""red"">Not Available</font>")
	Else
		Response.WritE("<font color=""green"">Available</font>")
	End If

	Set testMailer = Nothing
End Function

Function EasyMail(mailFrom,mailTo,mailSubject,mailBody)
	set ezMailer = CreateObject("EasyMail.SMTP.5")
	ezMailer.FromAddr = mailFrom
	ezMailer.AddRecipient "", mailTo, 1
	ezMailer.MailServer = outgoingMailServer
	ezMailer.Subject = mailSubject

	'Set body format to HTML
	ezMailer.BodyFormat = 1

	ezMailer.BodyText = "<html><head><title>" & mailSubject & "</title></head><body>" & _
		mailBody & _
		"<br>" & _
		"<hr>" & _
		"<small>" & _
		"Message from the Sundance Media Group<br>" & _
		"<a href=""mailto:mannie@sundancemediagroup.com"">Click here to email Mannie of the Sundance Media Group.</a>" & _
		"</small>" & _
		"</body></html>"

	'Import an HTML file in to the object.
	'ret = SMTP.ImportBodyTextEx(Request.ServerVariables("APPL_PHYSICAL_PATH") & "index.htm", 2 + 4)

	'If Not ret = 0 Then
	'	lsMailResult = "Error importing File " & ret
	'	Exit Sub
	'End If

	ret = ezMailer.Send()

	If Not ret = 0 Then
		EasyMail = "Error Sending Message " & ret
		Exit Function
	End If

	EasyMail = "Message sent successfully!"
End Function
%>
<meta name="Microsoft Theme" content="modified-powerplugs-web-templates-art3dblue 011, default">
</head>
