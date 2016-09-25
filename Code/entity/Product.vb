<Serializable()> _
Public Class Product
    Private _id As Guid
    Private _title As String
    Private _description As String
    Private _webPage As String
    Private _imageFile As String
    Private _editedOn As DateTime
    Private _editedBy As AMP.Person
    Private _showOn As DateTime
    Private _hideAfter As DateTime
    Private _section As Integer = 0     ' bitmask
    Private _categories As New AMP.CategoryCollection
    Private _formats As AMP.ProductFormatCollection
    Private _publisher As AMP.Publisher
    Private _vendor As AMP.Vendor

#Region " Properties "

    Public Property ImageFile() As String
        Get
            Return _imageFile
        End Get
        Set(ByVal Value As String)
            _imageFile = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	optionally specify custom detail page
    '
    '	Date:		Name:	Description:
    '	3/7/05  	JEA		Creation
    '-------------------------------------------------------------------------
    Public Property WebPage() As String
        Get
            Return _webPage
        End Get
        Set(ByVal Value As String)
            _webPage = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	bitmask of site.section values
    '
    '	Date:		Name:	Description:
    '	12/10/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Section() As Integer
        Get
            Return _section
        End Get
        Set(ByVal Value As Integer)
            _section = Value
        End Set
    End Property

    Public ReadOnly Property SectionName() As String
        Get
            If Common.IsFlag(_section) Then
                Dim section As AMP.Site.Section = CType(_section, AMP.Site.Section)
                Return section.ToString
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Property Categories() As AMP.CategoryCollection
        Get
            Return _categories
        End Get
        Set(ByVal Value As AMP.CategoryCollection)
            _categories = Value
        End Set
    End Property

    Public Property Vendor() As AMP.Vendor
        Get
            Return _vendor
        End Get
        Set(ByVal Value As AMP.Vendor)
            _vendor = Value
        End Set
    End Property

    Public Property Publisher() As AMP.Publisher
        Get
            Return _publisher
        End Get
        Set(ByVal Value As AMP.Publisher)
            _publisher = Value
        End Set
    End Property

    Public Property ShowOn() As DateTime
        Get
            Return _showOn
        End Get
        Set(ByVal Value As DateTime)
            _showOn = Value
        End Set
    End Property

    Public Property HideAfter() As DateTime
        Get
            Return _hideAfter
        End Get
        Set(ByVal Value As DateTime)
            _hideAfter = Value
        End Set
    End Property

    Public ReadOnly Property CreatedBy() As AMP.Person
        Get
            Return _editedBy
        End Get
    End Property

    Public ReadOnly Property CreatedOn() As DateTime
        Get
            Return _editedOn
        End Get
    End Property


    Public Property ID() As Guid
        Get
            Return _id
        End Get
        Set(ByVal Value As Guid)
            _id = Value
        End Set
    End Property

    Public Property Title() As String
        Get
            Return _title
        End Get
        Set(ByVal Value As String)
            _title = Security.SafeString(Value, 150)
        End Set
    End Property

    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal Value As String)
            _description = Security.SafeString(Value, 1500)
        End Set
    End Property

#End Region

    Public Sub New()
        _id = Guid.NewGuid
        _formats = New AMP.ProductFormatCollection
    End Sub

    Public Function Price() As Single

    End Function
End Class
