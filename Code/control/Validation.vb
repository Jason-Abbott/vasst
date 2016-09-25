Namespace Controls
    Public Class Validation
        Inherits HtmlControls.HtmlControl

        Private _type As String
        Private _required As Boolean = False
        Private _targetID As String
        Private _message As String
        Private _control As Control

#Region " Properties "

        '---COMMENT---------------------------------------------------------------
        '   the message that will be displayed via javascript if the field is invalid
        ' 
        '   Date:       Name:   Description:
        '	1/2/05      JEA     Created
        '-------------------------------------------------------------------------
        Public Property Message() As String
            Get
                Return _message
            End Get
            Set(ByVal Value As String)
                _message = Value
            End Set
        End Property

        '---COMMENT---------------------------------------------------------------
        '   the message that will be displayed via javascript if the field is invalid
        ' 
        '   Date:       Name:   Description:
        '	1/2/05      JEA     Created
        '-------------------------------------------------------------------------
        Public WriteOnly Property Resx() As String
            Set(ByVal Value As String)
                _message = DirectCast(Me.Page, AMP.Page).Say(Value)
            End Set
        End Property

        '---COMMENT---------------------------------------------------------------
        '   the type must match a function in validation.js with name Is[Type]()
        ' 
        '   Date:       Name:   Description:
        '	1/2/05      JEA     Created
        '-------------------------------------------------------------------------
        Public Property Type() As String
            Get
                Return _type
            End Get
            Set(ByVal Value As String)
                _type = Value
            End Set
        End Property

        Public Property Required() As Boolean
            Get
                Return _required
            End Get
            Set(ByVal Value As Boolean)
                _required = Value
            End Set
        End Property

        '---COMMENT---------------------------------------------------------------
        '   the server-side ID of the field this validation instance should check
        ' 
        '   Date:       Name:   Description:
        '	1/2/05      JEA     Created
        '-------------------------------------------------------------------------
        Public WriteOnly Property Target() As String
            Set(ByVal Value As String)
                _targetID = Value
            End Set
        End Property

        '---COMMENT---------------------------------------------------------------
        '   the control having the given server-side ID
        ' 
        '   Date:       Name:   Description:
        '	1/2/05      JEA     Created
        '   2/18/05     JEA     Allow value to be set
        '-------------------------------------------------------------------------
        Public Property Control() As Control
            Get
                If _control Is Nothing Then _control = Me.Parent.FindControl(_targetID)
                Return _control
            End Get
            Set(ByVal Value As Control)
                _control = Value
            End Set
        End Property

#End Region
       
        Private Sub Validation_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If Me.Visible Then Me.Register()
        End Sub

        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            ' no rendering
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	add reference to self in form control for easier rendering
        '
        '	Date:		Name:	Description:
        '	1/2/05  	JEA		Creation
        '   2/11/05     JEA     Handle traversal that doesn't find form
        '   2/24/05     JEA     Get form directly
        '-------------------------------------------------------------------------
        Public Sub Register(ByVal page As AMP.Page)
            page.Form.Validation.Add(Me)
        End Sub

        Public Sub Register()
            Me.Register(DirectCast(Me.Page, AMP.Page))
        End Sub
    End Class
End Namespace