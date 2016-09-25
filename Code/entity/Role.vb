<Serializable()> _
Public Class Role
    Private _permission As AMP.Site.Permission()
    Private _inheritFrom As New AMP.RoleCollection
    Private _name As String
    Private _id As AMP.Site.Role

#Region " Properties "

    Public Property ID() As AMP.Site.Role
        Get
            Return _id
        End Get
        Set(ByVal Value As AMP.Site.Role)
            _id = Value
        End Set
    End Property

    Public ReadOnly Property [Enum]() As Integer
        Get
            Return CInt(_id)
        End Get
    End Property

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal Value As String)
            _name = Security.SafeString(Value, 100)
        End Set
    End Property

    Public Property Permissions() As AMP.Site.Permission()
        Get
            Return _permission
        End Get
        Set(ByVal Value As AMP.Site.Permission())
            _permission = Value
        End Set
    End Property

    Public Property InheritFrom() As AMP.RoleCollection
        Get
            Return _inheritFrom
        End Get
        Set(ByVal Value As AMP.RoleCollection)
            _inheritFrom = Value
        End Set
    End Property

#End Region

#Region " Add / Remove methods "

    '---COMMENT---------------------------------------------------------------
    '	does role contain given permission
    '
    '	Date:		Name:	Description:
    '	2/2/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function HasPermission(ByVal permission As AMP.Site.Permission) As Boolean
        Return (Array.IndexOf(Me.EffectivePermissions, permission) <> -1)
    End Function

    Public Function HasPermission(ByVal permission As String) As Boolean
        Return Me.HasPermission(CType(permission, AMP.Site.Permission))
    End Function

    '---COMMENT---------------------------------------------------------------
    '	add permission to array
    '
    '	Date:		Name:	Description:
    '	2/2/05	    JEA		Creation
    '   2/14/05     JEA     Reset cached permissions
    '-------------------------------------------------------------------------
    Public Function AddPermission(ByVal permission As AMP.Site.Permission) As Boolean
        If Not Me.HasPermission(permission) Then
            Dim newLength As Integer = _permission.Length + 1
            ReDim Preserve _permission(newLength)
            _permission(newLength - 1) = permission
            ' force cached permissions to reset
            Profile.Permissions = Nothing
            Return True
        End If
        Return False
    End Function

    Public Function AddPermission(ByVal permission As String) As Boolean
        Return Me.AddPermission(CType(permission, AMP.Site.Permission))
    End Function

    '---COMMENT---------------------------------------------------------------
    '	add permission to array
    '
    '	Date:		Name:	Description:
    '	2/2/05	    JEA		Creation
    '   2/14/05     JEA     Reset cached permissions
    '-------------------------------------------------------------------------
    Public Function RemovePermission(ByVal permission As AMP.Site.Permission) As Boolean
        If Me.HasPermission(permission) Then
            Dim newLength As Integer = _permission.Length - 1
            Dim permissions(newLength) As AMP.Site.Permission
            Dim y As Integer = 0

            For x As Integer = 0 To newLength
                If _permission(x) <> permission Then
                    permissions(y) = _permission(x)
                    y += 1
                End If
            Next

            _permission = permissions
            ' force cached permissions to reset
            Profile.Permissions = Nothing
            Return True
        End If
        Return False
    End Function

    Public Function RemovePermission(ByVal permission As String) As Boolean
        Return Me.RemovePermission(CType(permission, AMP.Site.Permission))
    End Function

#End Region

    '---COMMENT---------------------------------------------------------------
    '	get all permissions that apply to this role
    '
    '	Date:		Name:	Description:
    '	12/7/04	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function EffectivePermissions() As AMP.Site.Permission()
        Dim pc As New PermissionCollection

        ' permissions applied directly to this role
        For x As Integer = 0 To Me.Permissions.Length - 1
            pc.Add(Me.Permissions(x))
        Next

        ' recurse through inherited permissions
        For Each role As AMP.Role In Me.InheritFrom
            Dim inherited As AMP.Site.Permission() = role.EffectivePermissions
            For x As Integer = 0 To inherited.Length - 1
                pc.Add(inherited(x))
            Next
        Next
        Return pc.All
    End Function

End Class
