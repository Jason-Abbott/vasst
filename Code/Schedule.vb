Option Strict On

Imports System.text
Imports AMP.Common
Imports System.data.OleDb
Imports System.Threading
Imports System.Configuration.ConfigurationSettings

Public Class Schedule
    Private _db As AMP.Data.Sql
    Public Thread As System.Threading.Thread
    Private _scheduleConfig As Hashtable
    Private _listings As OleDbDataReader
    Private _job As Schedule.Job
    Private _sleepTime As Integer
    Private _executeAgain As Boolean = True
    Private _machineName As String              ' logged value
    Private _process As String                  ' logged value
    Private Const _oneMinute As Integer = 60000

    Public Enum Activity
        Notify
        Expire
    End Enum

    Public Sub New()
        _scheduleConfig = CType(ConfigurationSettings.GetConfig("Schedule"), Hashtable)
        Dim repeatMinutes As Double = CDbl(_scheduleConfig("RepeatIntervalMinutes"))
        _sleepTime = CInt(repeatMinutes * _oneMinute)
        _db = New AMP.Data.Sql
        _db.ReuseConnection = True
        _job = New Schedule.Job
        _job.URL = AppSettings.Item("BaseURL")
        _machineName = Environment.MachineName
        _process = AppDomain.CurrentDomain.FriendlyName
        Me.Log("Scheduling thread initialized for " & repeatMinutes & " minute repeat")
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	Call functions indefinitely in loop, pausing thread before repeating
    '
    '	Date:		Name:	Description:
    '	10/15/04	JEA 	Creation
    '-------------------------------------------------------------------------
    Public Sub Start()
        While _executeAgain
            Try
                ' execute all jobs
                Me.NotifyAboutOldListings()
                Me.ExpireListings()
            Finally
                Me.Log("Thread going to sleep for " & _sleepTime & " milliseconds")
                Thread.Sleep(_sleepTime)
                Me.Log("Thread awake")
            End Try
        End While
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	Notify user two weeks before listing expires
    '
    '	Date:		Name:	Description:
    '	10/5/04	    JEA 	Creation
    '   10/14/04    JEA     Log activity
    '-------------------------------------------------------------------------
    Public Sub NotifyAboutOldListings()
        _job.Activity = Activity.Notify
        With _db
            .Command.CommandText = "procGetListingsToExpire"
            .AddParam("@lWeeks", CInt(_scheduleConfig("WarnAfterWeeksUnedited")), SqlDbType.Int)
            _listings = CType(.GetReader, OleDbDataReader)
        End With

        If _listings.HasRows Then
            Me.GenerateEmail()
        Else
            Me.Reset()
            Me.Log("No listings to warn about")
        End If
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	Expire listings and notify user
    '
    '	Date:		Name:	Description:
    '	10/5/04	    JEA 	Creation
    '   10/14/04    JEA     Log activity
    '-------------------------------------------------------------------------
    Public Sub ExpireListings()
        _job.Activity = Activity.Expire
        With _db
            .Command.CommandText = "procExpireListings"
            .AddParam("@lWeeks", CInt(_scheduleConfig("ExpireWeeksAfterWarning")), SqlDbType.Int)
            _listings = CType(.GetReader, OleDbDataReader)
        End With

        If _listings.HasRows Then
            Me.GenerateEmail()
        Else
            Me.Reset()
            Me.Log("No listings to expire")
        End If
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	Notify users of scheduled activity
    '
    '	Date:		Name:	Description:
    '	10/5/04	    JEA 	Creation
    '   10/14/04    JEA     String formatting and log activity
    '-------------------------------------------------------------------------
    Private Sub GenerateEmail()
        Dim mail As New AMP.Email
        Dim emailCount As Integer = 0

        
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	Clear command and reader for next iteration
    '
    '	Date:		Name:	Description:
    '	10/15/04	JEA 	Creation
    '-------------------------------------------------------------------------
    Private Sub Reset()
        _db.Command.Cancel()
        _listings.Close()
        _db.Finish()
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	Cancel the job loop and abort thread
    '
    '	Date:		Name:	Description:
    '	10/15/04	JEA 	Creation
    '-------------------------------------------------------------------------
    Public Sub Cancel()
        _executeAgain = False
        Me.Log("Scheduling thread cancelled")
        Me.Thread.Abort()
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	Log timer activity
    '
    '	Date:		Name:	Description:
    '	10/14/04	JEA 	Creation
    '-------------------------------------------------------------------------
    Private Sub Log(ByVal description As String)
        With _db
            .Command.CommandText = "procLogTimerEvent"
            .AddParam("@description", description, SqlDbType.VarChar)
            .AddParam("@machineName", _machineName, SqlDbType.VarChar)
            .AddParam("@process", _process, SqlDbType.VarChar)
            .ExecuteOnly()
        End With
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	Object representing single job (e-mail contact)
    '
    '	Date:		Name:	Description:
    '	10/14/04	JEA 	Creation
    '-------------------------------------------------------------------------
    Public Class Job
        Private _activity As Schedule.Activity
        Private _contactName As String
        Private _email As String
        Private _userID As Short = 0
        Private _url As String
        Private _listingCount As Integer = 0
        Public Body As New StringBuilder

        Public Function Matches(ByVal userID As Object) As Boolean
            Return (_userID = CShort(userID))
        End Function

        Public Function HasContent() As Boolean
            Return (_listingCount > 0)
        End Function

        '---COMMENT---------------------------------------------------------------
        '	Build HTML for single listing
        '
        '	Date:		Name:	Description:
        '	10/14/04	JEA 	Creation
        '-------------------------------------------------------------------------
        Public Sub Add(ByRef listing As SqlClient.SqlDataReader)
            _listingCount += 1

            With listing
                Body.Append("<li style=""padding-bottom: 9px;"">")
                If .Item("referenceNumber").ToString <> Nothing Then
                    Body.Append("<i>")
                    Body.Append(.Item("referenceNumber"))
                    Body.Append("</i><br>")
                End If
                Body.Append(.Item("address1"))
                If .Item("address2").ToString <> Nothing Then
                    Body.Append("<br>")
                    Body.Append(.Item("address2"))
                End If
                Body.Append("<br>")
                Body.Append(.Item("city"))
                Body.Append(", ")
                Body.Append(.Item("state"))
                Body.Append(" &nbsp;")
                Body.Append(Format.Zip(.Item("zip").ToString))
                Body.Append("<br>[ <a href=""")
                Body.Append(_url)
                Body.Append("renew.aspx?id=")
                Body.Append(.Item("renewCode"))
                Body.Append("""><RenewListing></a> ]")
                If _activity = Activity.Notify Then
                    ' if only notifying then listing can still be viewed
                    Body.Append("&nbsp;[ <a href=""")
                    Body.Append(_url)
                    Body.Append("Detail.aspx?id=")
                    Body.Append(.Item("listingID"))
                    Body.Append("""><ViewListing></a> ]")
                End If
                Body.Append("</li>")
            End With
        End Sub

#Region " Properties "

        Public ReadOnly Property ListingCount() As Integer
            Get
                Return _listingCount
            End Get
        End Property

        Public WriteOnly Property URL() As String
            Set(ByVal Value As String)
                _url = Value
            End Set
        End Property

        Public Property UserID() As Short
            Get
                Return _userID
            End Get
            Set(ByVal Value As Short)
                _userID = Value
            End Set
        End Property

        Public Property Email() As String
            Get
                Return _email
            End Get
            Set(ByVal Value As String)
                _email = Value
            End Set
        End Property

        Public Property ContactName() As String
            Get
                Return _contactName
            End Get
            Set(ByVal Value As String)
                _contactName = Value
            End Set
        End Property

        Public Property Activity() As Schedule.Activity
            Get
                Return _activity
            End Get
            Set(ByVal Value As Schedule.Activity)
                _activity = Value
            End Set
        End Property
#End Region
    End Class
End Class


