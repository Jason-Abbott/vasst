<!--#include file="functions.asp"-->
<% printHeader %>
<BR>
<HR>
<CENTER>
<B>Sundance Media Group</B><BR>
Customer Management System<BR>
<HR>
<BR>
</CENTER>
</HR>
<BR>
<BR>
<BLOCKQUOTE>
<B>Last Succesful Login:</B><BR>
&nbsp;&nbsp;&nbsp;&nbsp;<% printLastLoginInfo %><BR>
<BR>
<B>Last Failed Login:</B><BR>
&nbsp;&nbsp;&nbsp;&nbsp;<% printLastFailedLoginInfo %><BR>
<BR>
<B>New Signups Since Last Login:</B><BR>
<% printNewSignups %><BR>
</BLOCKQUOTE>
<% printFooter %>