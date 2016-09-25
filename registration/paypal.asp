<% authNotNeeded = True %>
<!--#include file="admin/functions.asp"-->
<% If Request.Cookies("verisign")("custID") = "" Then Response.Redirect "register.asp" End If %>
<meta name="Microsoft Theme" content="modified-powerplugs-web-templates-art3dblue 011, default">
<body bgcolor=#C0C0C0 background="../_themes/modified-powerplugs-web-templates-art3dblue/background.gif" text="#000000" link="#6A6A6A" vlink="#808080" alink="#FFFFFF"><!--mstheme--><font face="Arial, Arial, Helvetica">
<!--#include file="top.html"-->
			<!--mstheme--></font><table border=0 cellpadding=5 cellspacing=0 style="border: 2px dotted gray;">
				<tr>
					<td><!--mstheme--><font face="Arial, Arial, Helvetica">
						<font color=#800000>
<%
if (loadPaypalForm(Request.Cookies("verisign")("custID"))) Then
'	Response.Cookies("verisign") = ""
%> 
<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
<input type="hidden" name="cmd" value="_xclick">
<input type="hidden" name="business" value="paypal@sundancemediagroup.com">
<input type="hidden" name="item_name" value="<%=paypalTour%> registration for <%=paypalName%>">
<input type="hidden" name="item_number" value="<%=paypalCustID%>">
<input type="hidden" name="amount" value="<%=paypalPaymentAmount%>">
<input type="hidden" name="no_note" value="1">
<input type="hidden" name="currency_code" value="USD">
<b><%=paypalName%>, Thank you for registering!</b><br>
<br>
What we need to do now is have you go to a secure Paypal site so that we can get your payment, and complete the registration process.  Please click on the button to
go to Paypal.<br>
<br>
<center><i>All transactions are in US Dollars.</i><br /><input type="submit" value="Click to goto Paypal for Secure Payment."></center>
</form>
<%
else
%>
There was a problem with your registration, please <a href="/registration/register.asp">try again.</a>
<%
end if
%>
						</font>
					<!--mstheme--></font></td>
				</tr>
			</table><!--mstheme--><font face="Arial, Arial, Helvetica">
<!--#include file="bottom.html"-->
<!--mstheme--></font></td></tr></table><!--mstheme--><font face="Arial, Arial, Helvetica"><!--mstheme--></font></body>