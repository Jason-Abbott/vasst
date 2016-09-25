<% pageTitle = "Customer Manager" %>
<!--#include file="functions.asp"-->
<% 
printHeader 
printBox 
printMenu 

'Response.Cookies("userAuth")("menuShown") = True
'SplitCookies
If Not (Request.QueryString("page") = "") Then
	Select Case Request.QueryString("page")
	Case "index.asp"
		Response.Redirect Request.QueryString("page")
	Case "customer_spreadsheet.asp"
		Response.Redirect Request.QueryString("page")
	Case "customer_editspreadsheet.asp"
		Response.Redirect Request.QueryString("page")
	Case "verisignlog.asp"
		Response.Redirect Request.QueryString("page")
	Case "sqltest.asp"
		Response.Redirect Request.QueryString("page")
	Case "emailer.asp"
		Response.Redirect Request.QueryString("page")
	Case "reportgenerator.asp"
		Response.Redirect Request.QueryString("page") & "?" & "i=" & Request.QueryString("i") & "&c=" & Request.QueryString("c") & "&o=" & Request.QueryString("o")
	Case "logout.asp"
		Response.Redirect Request.QueryString("page")
	Case Else
		insidePage = Request.QueryString("page")
		If Not (Request.QueryString("searchFor") = "") Then
			searchFor = "?searchFor=" & Server.URLEncode(Request.QueryString("searchFor"))
			If Not (Request.QueryString("viewCustomer") = "") Then
				Response.Cookies("customerPass")("editID") = Request.QueryString("viewCustomer")
				Response.Redirect "index.asp?page=" & Request.QueryString("page") & "&searchFor=" & Server.URLEncode(Request.QueryString("searchFor"))
			End If
		Else
			If Not (Request.QueryString("viewCustomer") = "") Then
				Response.Cookies("customerPass")("editID") = Request.QueryString("viewCustomer")
				Response.Redirect "index.asp?page=" & Request.QueryString("page")
			End If
		End If
	End Select
Else
	insidePage = "about.asp"
End If

%>
<IFRAME SRC="<%=insidePage%><%=searchFor%><%=viewCustomer%>" FRAMEBORDER=0 ID=srcFrame NAME=srcFrame WIDTH=<%=displayWidth%> HEIGHT=<%=displayHeight%>></IFRAME>
<%
'Response.Cookies("userAuth")("menuShown") = False
'SplitCookies

printBoxClose 
printFooter 
%>