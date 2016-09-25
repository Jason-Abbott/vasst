<!--#include file="functions.asp"-->
<% logout %>
<FORM METHOD=post>
	<TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0 WIDTH=300 ALIGN=center>
		<TR>
			<TD CLASS=orangebox>
				Logging out...
			</TD>
		</TR>
	</TABLE>
</FORM>
<script>
	href = new String(top.location.href)
	if (href.indexOf("logout.asp") == -1) 
	{
		top.location.href = "logout.asp";
	}
</script>
<% printFooter %>