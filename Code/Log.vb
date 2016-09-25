Imports System.Text
Imports System.Threading
Imports System.Data.OleDb
Imports System.Configuration.ConfigurationSettings

Public Class Log

    Private Shared _clientIP As String = HttpContext.Current.Request.UserHostAddress
    Private Shared _machineName As String = Environment.MachineName
    Private Shared _process As String = AppDomain.CurrentDomain.FriendlyName
    Private Shared _recursed As Boolean = False

#Region " Enumerations "

    <Flags()> _
    Public Enum ErrorType
        SendEmail = FileSystem
        Unknown = &H0
        FileSystem = &H1
        Syntax = &H2
        Serialization = &H4
        Custom = &H8
        Url = &H10
    End Enum

#End Region

#Region " Visits "

    '---COMMENT---------------------------------------------------------------
    '   get and set visit count for the day
    ' 
    '   Date:       Name:   Description:
    '	2/18/05     JEA     Created
    '-------------------------------------------------------------------------
    Public Shared Function VisitsToday() As Integer
        Dim db As New AMP.Data.Jet(AppSettings("ActivityLogStore"))
        Dim visits As Integer
        With db
            .Command.CommandText = "SELECT [count] FROM visits WHERE [day] = ?"
            .AddParam("[day]", DateTime.Today, OleDbType.Date)
            Try
                visits = CInt(.GetSingleValue())
            Catch
                visits = 0
            End Try
        End With
        Return visits
    End Function

    Public Shared Sub AddVisit()
        Dim db As New AMP.Data.Jet(AppSettings("ActivityLogStore"))
        With db
            If (Not Global.DailyLog) OrElse Log.VisitsToday = 0 Then
                .Command.CommandText = "INSERT INTO visits ([day], [count]) VALUES (?, ?)"
                .AddParam("[day]", DateTime.Today, OleDbType.Date)
                .AddParam("[count]", 1, OleDbType.Numeric)
                Global.DailyLog = True
            Else
                .Command.CommandText = "UPDATE visits SET [count] = [count] + 1 WHERE [day] = ?"
                .AddParam("[day]", DateTime.Today, OleDbType.Date)
            End If
            Try
                .ExecuteOnly()
            Catch ex As Exception
                BugOut(ex.Message)
            End Try
        End With
    End Sub

#End Region

#Region " Errors "

    '---COMMENT---------------------------------------------------------------
    '   log errors
    ' 
    '   Date:       Name:   Description:
    '	1/18/05     JEA     Created
    '   2/11/05     JEA     Limit frequency of e-mailed errors
    '-------------------------------------------------------------------------
    Public Shared Sub [Error](ByVal ex As Exception, ByVal type As AMP.Log.ErrorType, _
        ByVal user As AMP.Person)

        Dim stack As String
        Dim email As Boolean = False

        If user Is Nothing Then
            user = New AMP.Person
            user.FirstName = "VASST"
            user.LastName = "Guest"
        End If

        If (type And ErrorType.SendEmail) > 0 Then
            Dim frequency As TimeSpan = TimeSpan.FromMinutes(CDbl(AppSettings("ErrorFrequency")))
            Dim key As String = ex.Message
            Dim occurrences As Stack

            If Global.ErrorEmail.ContainsKey(key) Then
                occurrences = DirectCast(Global.ErrorEmail.Item(key), Stack)
            Else
                occurrences = New Stack
            End If

            If occurrences.Count = 0 OrElse _
                TimeSpan.Compare(Now.Subtract(CDate(occurrences.Peek)), frequency) >= 0 Then

                ' at least frequency number of minutes have passed since last mail
                Dim mail As New AMP.Email
                Dim addresses As String() = AppSettings("ErrorMailTo").Split(","c)

                mail.Error(ex, user, _clientIP, _machineName, _process, addresses, occurrences.Count)
                email = True

                ' reset the stack
                If occurrences.Count > 0 Then occurrences.Clear()
            End If

            occurrences.Push(DateTime.Now)
            Global.ErrorEmail.Item(key) = occurrences
        End If

        If Not _recursed Then
            ' only attempt db write if not recursing
            Try
                Dim db As New AMP.Data.Jet(AppSettings("ErrorLogStore"))
                With db
                    If _process Is Nothing Then _process = "Unknown"
                    stack = ex.StackTrace
                    If stack = Nothing Then stack = "empty"

                    .Command.CommandText = "INSERT INTO log ([Date], Message, Emailed, ClientIP, " _
                        & "ServerName, Stack, Process) VALUES (?, ?, ?, ?, ?, ?, ?)"

                    .AddParam("[Date]", DateTime.Now, OleDbType.Date)
                    .AddParam("Message", ex.Message, OleDbType.VarChar)
                    .AddParam("Emailed", email, OleDbType.Boolean)
                    .AddParam("ClientIP", _clientIP, OleDbType.VarChar)
                    .AddParam("ServerName", _machineName, OleDbType.VarChar)
                    .AddParam("Stack", stack, OleDbType.VarChar)
                    .AddParam("Process", _process, OleDbType.VarChar)

                    .ExecuteOnly()
                End With
            Catch e As Exception
                _recursed = True
                Log.Error(e, ErrorType.FileSystem, user)
            End Try
        Else
            _recursed = False
        End If
    End Sub

    Public Shared Sub [Error](ByVal ex As Exception, ByVal type As AMP.Log.ErrorType)
        AMP.Log.Error(ex, type, Nothing)
    End Sub

    ' overload for custom error message
    Public Shared Sub [Error](ByVal message As String, ByVal type As AMP.Log.ErrorType, _
        ByVal user As AMP.Person)

        Dim fullMessage As String = String.Format("{0} (for {1})", message, _
            HttpContext.Current.Request.Url.PathAndQuery)
        Dim ex As New Exception(fullMessage)
        AMP.Log.Error(ex, type, user)
    End Sub

    Public Shared Sub [Error](ByVal message As String, ByVal type As AMP.Log.ErrorType)
        AMP.Log.Error(message, type, Nothing)
    End Sub

#End Region

#Region " Activity "

    Private Structure Entry
        Public Activity As AMP.Site.Activity
        Public UserID As String
        Public AssetID As String
        Public ProductID As String
        Public ClientIP As String
    End Structure

    '---COMMENT---------------------------------------------------------------
    '   log activity
    ' 
    '   Date:       Name:   Description:
    '	1/18/05     JEA     Created
    '   2/10/05     JEA     UserID also holds e-mail
    '-------------------------------------------------------------------------
    Private Sub SaveActivity(ByVal state As Object)
        Dim entry As Log.Entry = DirectCast(state, Log.Entry)
        Dim db As New AMP.Data.Jet(AppSettings("ActivityLogStore"))

        With db
            .Command.CommandText = "INSERT INTO log ([Date], ActivityID, UserID, AssetID, " _
                & "ProductID, ClientIP) VALUES (?, ?, ?, ?, ?, ?)"

            .AddParam("[Date]", DateTime.Now, OleDbType.Date)
            .AddParam("ActivityID", entry.Activity, OleDbType.Integer)
            .AddParam("UserID", entry.UserID, OleDbType.VarChar)
            .AddParam("AssetID", entry.AssetID, OleDbType.VarChar)
            .AddParam("ProductID", entry.ProductID, OleDbType.VarChar)
            .AddParam("ClientIP", entry.ClientIP, OleDbType.VarChar)

            Try
                .ExecuteOnly()
            Catch e As Exception
                If e.Message.IndexOf("updateable query") <> -1 Then
                    Dim person As AMP.Person
                    If entry.UserID <> String.Empty Then person = WebSite.Persons(entry.UserID)
                    Me.Error(e, ErrorType.FileSystem, person)
                Else
                    BugOut("Couldn't save activity because {0}", e.Message)
                End If
            End Try
        End With
    End Sub

    '---COMMENT---------------------------------------------------------------
    '   shared methods for easy use
    ' 
    '   Date:       Name:   Description:
    '	1/18/05     JEA     Created
    '-------------------------------------------------------------------------
    Public Shared Sub Activity(ByVal type As AMP.Site.Activity, ByVal userID As Guid, _
        ByVal assetID As Guid, ByVal productID As Guid, ByVal email As String)

        Dim log As New AMP.Log
        Dim entry As New AMP.Log.Entry

        With entry
            .Activity = type
            .UserID = log.CheckGuid(userID)
            .AssetID = log.CheckGuid(assetID)
            .ProductID = log.CheckGuid(productID)
            .ClientIP = _clientIP
            If .UserID = Nothing Then .UserID = log.CheckNull(email)
        End With
        ThreadPool.QueueUserWorkItem(AddressOf log.SaveActivity, entry)
    End Sub

    Public Shared Sub Activity(ByVal type As AMP.Site.Activity, ByVal userID As Guid, _
        ByVal assetID As Guid)

        AMP.Log.Activity(type, userID, assetID, Nothing, Nothing)
    End Sub

    Public Shared Sub Activity(ByVal type As AMP.Site.Activity, ByVal userID As Guid)
        AMP.Log.Activity(type, userID, Nothing, Nothing, Nothing)
    End Sub

    Public Shared Sub Activity(ByVal type As AMP.Site.Activity, ByVal email As String)
        AMP.Log.Activity(type, Nothing, Nothing, Nothing, email)
    End Sub

    Public Shared Sub activity(ByVal type As AMP.Site.Activity)
        AMP.Log.Activity(type, Nothing, Nothing, Nothing, Nothing)
    End Sub

#End Region

    Private Shared Function CheckGuid(ByVal value As Guid) As String
        Return IIf(value.Equals(Guid.Empty), "", value).ToString
    End Function

    Private Shared Function CheckNull(ByVal value As Object) As String
        Return IIf(value Is Nothing, "", value).ToString
    End Function

End Class
