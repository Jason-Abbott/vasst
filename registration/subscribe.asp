<%
  If (Request.Form("add") <> "") Then
    If (Request.Form("name") <> "") And (Request.Form("email") <> "") And (Request.Form("list") <> "") Then
    	authNotNeeded = true
    Else
      Response.Write("<script>")
      Response.Write(" alert(""There was a problem subscribing to this mailing list, please try again.\nIf the problem persists please contact the administrator.\n\nError: Insufficient input from user.""); ")
      Response.Write(" self.history.go(-1); ")
      Response.Write("</script>")
      Response.End
    End If
  End If
%>
<!--#include file="admin/functions.asp"-->
<%
  If (Request.Form("add") <> "") Then
    openDatabase
    Set checkList = dbConnection.Execute("SELECT * FROM emailer_tblMailingLists WHERE listID = " & Request.Form("list") & " AND listAllowSubscribe = True")
    If (checkList.EOF) Then
      Response.Write("<script>")
      Response.Write(" alert(""There was a problem subscribing to this mailing list, please try again.\nIf the problem persists please contact the administrator.\n\nError: List does not exist. (" & Request.Form("list") & ")""); ")
      Response.Write(" self.history.go(-1); ")
      Response.Write("</script>")
      Response.End
    Else
      dbConnection.Execute("INSERT INTO emailer_tblMembers ( memberName, memberEmail, memberOf ) VALUES ( '" & Request.Form("name") & "', '" & Request.Form("email") & "', " & Request.Form("list") & " )")
      Response.Write("<script>")
      Response.Write(" alert(""Thank you for subscribing.""); ")
      Response.Write(" self.history.go(-1); ")
      Response.Write("</script>")
      Response.End
    End If
    Response.End
    closeDatabase
  End If
%>