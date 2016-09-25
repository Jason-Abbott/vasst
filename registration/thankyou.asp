<% authNotNeeded = True %>
<!--#include file="admin/functions.asp"-->
<% 
	If (Session("cartorder") = "yes") Then
		returnurl = Session("returnurl")
		Session("cartorder") = ""
		Session("returnurl") = ""
		Session("savestep4") = "yes"
		Response.Redirect returnurl & "?function=checkout"
	ElseIf (Session("certorder") = "yes") Then
		returnurl = Session("returnurl")
		Session("returnurl") = ""
		Session("certorder") = ""
		Response.Redirect returnurl & "?function=menu"
	End If
%>
<% If Request.Cookies("verisign")("custID") = "" Then Response.Redirect "register.asp" End If %>
<html>
<head>
<title>Thank you for your payment.</title>
<meta name="Microsoft Theme" content="modified-powerplugs-web-templates-art3dblue 011, default">
</head>

<body bgcolor=#C0C0C0 background="../_themes/modified-powerplugs-web-templates-art3dblue/background.gif" text="#000000" link="#6A6A6A" vlink="#808080" alink="#FFFFFF"><!--mstheme--><font face="Arial, Arial, Helvetica">
<!--#include file="top.html"-->
			<!--mstheme--></font><table border=0 cellpadding=5 cellspacing=0 style="border: 2px dotted gray;">
				<tr>
					<td><!--mstheme--><font face="Arial, Arial, Helvetica">
						<font color=#800000>
							<b>Thank you for going through the payment process!</b><br>
							<br>
							<%
							printVerisignConfirmation Request.Cookies("verisign")("custID")
							Response.Cookies("verisign")("custID") = ""
							%>
							<br>
							Kind regards,<br>
							<br>
							VASST Tour Registration<br>
							<br>
							<br>
							<center><a href="/aboutvasst.htm">Click here to return.</a></center>
						</font>
					<!--mstheme--></font></td>
				</tr>
			</table><!--mstheme--><font face="Arial, Arial, Helvetica">
<!--#include file="bottom.html"-->

<!--mstheme--></font></body>