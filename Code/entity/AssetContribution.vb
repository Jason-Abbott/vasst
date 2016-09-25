Public Class AssetContribution
    Private _finishedStep As Integer = 0
    Private _startedOn As DateTime
    Private _asset As AMP.Asset

#Region " Properties "

    Public Property StartedOn() As DateTime
        Get
            Return _startedOn
        End Get
        Set(ByVal Value As DateTime)
            _startedOn = Value
        End Set
    End Property

    Public Property Asset() As AMP.Asset
        Get
            Return _asset
        End Get
        Set(ByVal Value As AMP.Asset)
            _asset = Value
        End Set
    End Property

    Public Property FinishedStep() As Integer
        Get
            Return _finishedStep
        End Get
        Set(ByVal Value As Integer)
            _finishedStep = Value
        End Set
    End Property

#End Region

    Public Sub New()
        _asset = New AMP.Asset
        _startedOn = Now
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	save session asset to global entity
    '
    '	Date:		Name:	Description:
    '	1/15/05     JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Save() As Boolean
        WebSite.Assets.Add(_asset.Clone)
        WebSite.Save()
        Return True
    End Function

    '---COMMENT---------------------------------------------------------------
    '	remove any data associated with this contribution
    '
    '	Date:		Name:	Description:
    '	1/15/05     JEA		Creation
    '   2/11/05     JEA     Move logic to file object
    '-------------------------------------------------------------------------
    Public Sub Cancel()
        If Not _asset.File Is Nothing Then
            _asset.File.Delete()
            _asset = New AMP.Asset
        End If
        _finishedStep = 0
    End Sub

End Class
