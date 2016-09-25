<% authNotNeeded = True %>
<!--#include file="functions.asp"-->
<%
'<!--#include virtual="/cart/savetransaction.asp"-->
'<!--#include virtual="/certification/savetransaction.asp"-->
If (Request.ServerVariables("HTTP_REFERER") = "http://www.vasst.com/") Then
	'strFilePath = "D:\inetpub\vasst_com"
	
	verisignResult = Fix(Request.Form("RESULT"))
	verisignCustID = Request.Form("INVOICE")
	verisignRefID = Request.Form("PNREF")
	verisignResponseMSG = Request.Form("RESPMSG")
	verisignAVSData = Request.Form("AVSDATA")
	verisignAuthCode = Request.Form("AUTHCODE")
	verisignHostCode = Request.Form("HOSTCDOE")

	Session("verisignResult") = Fix(Request.Form("RESULT"))
	Session("verisignCustID") = Request.Form("INVOICE")
	Session("verisignRefID") = Request.Form("PNREF")
	Session("verisignResponseMSG") = Request.Form("RESPMSG")
	Session("verisignAVSData") = Request.Form("AVSDATA")
	Session("verisignAuthCode") = Request.Form("AUTHCODE")
	Session("verisignHostCode") = Request.Form("HOSTCDOE")
		
	If (verisignResult = "0") Then
		transactionSuccess = true
	Else
		transactionSuccess = false
	End If

	Session("transactionSuccess") = transactionSuccess

	SaveTransaction verisignResult, verisignCustID, verisignRefID, verisignResponseMSG, verisignAuthCode, transactionSuccess

	If (InStr(verisignCustID,"C") > 0) Then
		Response.Write("Cart Transaction")
		Server.Transfer("../../cart/savetransaction.asp")
		'SaveCartTransaction verisignResult, verisignCustID, verisignRefID, verisignResponseMSG, verisignAuthCode, transactionSuccess
	ElseIf (InStr(verisignCustID,"T") > 0) Then
		Response.Write("Test Transaction")
		Server.Transfer("../../certification/savetransaction.asp")
		'SaveCertTransaction verisignResult, verisignCustID, verisignRefID, verisignResponseMSG, verisignAuthCode, transactionSuccess
	Else
		Response.WritE("Seminar Transaction")
		saveVerisignInfo verisignResult, verisignCustID, verisignRefID, verisignResponseMSG, verisignAuthCode, transactionSuccess
	End If
	
End If
%>