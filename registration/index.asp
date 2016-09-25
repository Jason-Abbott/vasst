<% Response.Buffer = true %>
<%
'If (Len(Request.QueryString) = 0) Then
'	Response.Redirect "/registration/?sony"
'End If
If (Request.QueryString("noPrice")) then
	Response.Write("<script> alert('Pricing for this tour has not been setup, please contact the administrator.'); this.location.href = 'index.asp'; </script>")
End If
%>
<!-- LEAVE THIS HERE, MAKE SURE IT ALWAYS STAYS AT THE TOP -->
<!--#include file="top.html"--><head>
<title>Seminar Registration</title>
</head>


<B><FONT face=Arial color=#000000 size=3>Welcome to the V.A.S.S.T. Registration
page. We're looking forward to you joining us during the
tour.<BR></FONT><I><FONT face=Arial color=#000000 size=2><BR>Problems
registering for a tour?  </FONT><FONT face=Arial color=#000000 size=3><A
href="mailto:register@sundancemediagroup.com?subject=I'm having problems registering."><FONT
size=2>Click here</FONT></A><FONT face=Arial color=#000000 size=2> for
help.</FONT></FONT></I></B>
<table border=0 cellpadding=5 cellspacing=0 style="border: 2px dotted gray;">
	<tr>
		<td>
			<font color=#000000>
				<b>Please choose the tour you would like to register for:</b><br>
				<br>
				<!--This is where the seminar list will show up.-->
				<!--#include file="seminarlist.asp"-->
				<br>
			</font>
		</td>
	</tr>
</table>
<!--#include file="bottom.html"-->