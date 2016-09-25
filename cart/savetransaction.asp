<!--#include virtual="/cart/admin/includes.asp"-->
<%
'Function SaveCartTransaction(verisignResult, verisignCustID, verisignRefID, verisignResponseMSG, verisignAuthCode, transactionSuccess)
Function SaveCartTransaction
	openDB

	verisignResult = Session("verisignResult")
	verisignCustID = Session("verisignCustID")
	verisignRefID = Session("verisignRefID")
	verisignResponseMSG = Session("verisignResponseMSG")
	verisignAVSData = Session("verisignAVSData")
	verisignAuthCode = Session("verisignAuthCode")
	verisignHostCode = Session("verisignHostCode")
	transactionSuccess = Cbool(Session("transactionSuccess"))

	sVerisignData = "date=" & Server.URLEncode(date & " " & time) & "&RESULT=" & verisignResult & "&INVOICE=" & verisignCustID & "&PNREF=" & verisignRefID & "&RESPMSG=" & verisignResponseMSG & "&AUTHCODE=" & verisignAuthCode & "&success=" & transactionSuccess
	iOrderID = ParseInt(verisignCustID)
	
	sqlPaid = "UPDATE ordersession SET ispaid = true, paid = #" & date & " " & time & "#, transactiondata = '" & sVerisignData & "' WHERE id = " & iOrderID & ""
	sqlUnpaid = "UPDATE ordersession SET ispaid = false, transactiondata = '" & sVerisignData & "' WHERE id = " & iOrderID & ""
	
	If (transactionSuccess) Then
		cartDB.Execute(sqlPaid)
	Else
		cartDB.Execute(sqlUnpaid)
	End If
	
	CompleteOrder iOrderID

	closeDB
End Function

SaveCartTransaction

%>
		