<% 
pageTitle = "Need Authentication" 
authPage = True
%>
<!--#include file="functions.asp"-->
<% 
printHeader
%>
<meta name="Microsoft Theme" content="modified-powerplugs-web-templates-art3dblue 011, default">
<BODY onLoad="focusForm()" background="../../_themes/modified-powerplugs-web-templates-art3dblue/background.gif" bgcolor="#C0C0C0" text="#000000" link="#6A6A6A" vlink="#808080" alink="#FFFFFF"><!--mstheme--><font face="Arial, Arial, Helvetica">
<FORM METHOD=post>
<BR>
<!--msthemeseparator--><p align="center"><img src="../../_themes/modified-powerplugs-web-templates-art3dblue/hr.gif" width="466" height="5"></p>
<BR>
<CENTER>
<B>
Sundance Media Group
</B>
</CENTER>
<BR>
<!--msthemeseparator--><p align="center"><img src="../../_themes/modified-powerplugs-web-templates-art3dblue/hr.gif" width="466" height="5"></p>
<BR>
<BR>
			<!--mstheme--></font><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0 WIDTH=300 VALIGN=center ALIGN=center>
				<TR>
					<TD CLASS=orangebox ALIGN=center><!--mstheme--><font face="Arial, Arial, Helvetica">
						Login Box<!--msthemeseparator--><p align="center"><img src="../../_themes/modified-powerplugs-web-templates-art3dblue/hr.gif" width="466" height="5"></p>
						<!--mstheme--></font><TABLE WIDTH=100%>
							<TR>
								<TD CLASS=orange WIDTH=50><!--mstheme--><font face="Arial, Arial, Helvetica">Name:<!--mstheme--></font></TD>
								<TD CLASS=orange><!--mstheme--><font face="Arial, Arial, Helvetica"><INPUT NAME=username size="45"><!--mstheme--></font></TD>
							</TR>
							<TR>
								<TD CLASS=orange WIDTH=50><!--mstheme--><font face="Arial, Arial, Helvetica">Pass:<!--mstheme--></font></TD>
								<TD CLASS=orange><!--mstheme--><font face="Arial, Arial, Helvetica"><INPUT TYPE=password NAME=password size="45"><!--mstheme--></font></TD>
							</TR>
							<TR>
								<TD CLASS=orange COLSPAN=2><!--mstheme--><font face="Arial, Arial, Helvetica"><INPUT TYPE=submit NAME="Login" VALUE="Login" STYLE="Width:100%;"><!--mstheme--></font></TD>
							</TR>
							<TR>
								<TD CLASS=orange COLSPAN=3 ALIGN=center><!--mstheme--><font face="Arial, Arial, Helvetica"><FONT COLOR=red><%=loginMessage%></FONT><!--mstheme--></font></TD>
							</TR>
						</TABLE><!--mstheme--><font face="Arial, Arial, Helvetica">
					<!--mstheme--></font></TD>
				</TR>
			</TABLE><!--mstheme--><font face="Arial, Arial, Helvetica">
<% 
printFooter 
%>
<SCRIPT>
	href = new String(top.location.href)
	if (href.indexOf("login.asp") == -1) 
	{
		top.location.href = "login.asp";
	}
	
	function focusForm()
	{
		if (document.forms[0].username.value == "")
		{
			document.forms[0].username.focus();
		}
		else
		{
			document.forms[0].password.focus();
		}
	}
</SCRIPT>
</FORM>