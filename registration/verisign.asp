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
if (loadVerisignForm(Request.Cookies("verisign")("custID"))) Then
'	Response.Cookies("verisign") = ""
%>
<form method="POST" action="https://payflowlink.verisign.com/payflowlink.cfm">
<input type="hidden" name="LOGIN" value="vrn227251412">
<input type="hidden" name="PARTNER" value="wfb">
<input type="hidden" name="AMOUNT" value="<%=verisignPaymentAmount%>">
<input type="hidden" name="TYPE" value="S">
<input type="hidden" name="DESCRIPTION" value="<%=verisignTour%> registration for <%=verisignName%>">
<input type="hidden" name="NAME" value="<%=verisignName%>">
<input type="hidden" name="ADDRESS" value="<%=verisignAddress%>">
<input type="hidden" name="CITY" value="<%=verisignCity%>">
<input type="hidden" name="STATE" value="<%=verisignState%>">
<input type="hidden" name="ZIP" value="<%=verisignZip%>">
<input type="hidden" name="COUNTRY" value="<%=verisignCountry%>">
<input type="hidden" name="PHONE" value="<%=verisignPhone%>">
<input type="hidden" name="EMAIL" value="<%=verisignEmail%>">
<input type="hidden" name="INVOICE" value="<%=verisignCustID%>">
<b><%=verisignName%>, Thank you for registering!</b><br>
<br>
What we need to do now is have you go to a secure Verisign site so that we can get your Credit Card information for the payment, and complete the registration process.  Please click on the button to
go to Verisign.<br>
<br>
<center><i>All transactions are in US Dollars.</i><br /><input type="submit" value="Click to goto Verisign for Secure Payment."></center>
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