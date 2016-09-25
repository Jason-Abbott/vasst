Imports System.Text
Imports System.Collections

<Serializable()> _
Public Class RoleCollection
    Inherits CollectionBase

    <NonSerialized()> Private _roles As String

#Region " Properties "

    Public Property Item(ByVal index As Integer) As AMP.Role
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.Role)
        End Get
        Set(ByVal Value As AMP.Role)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    Default Public ReadOnly Property ByRole(ByVal roleID As AMP.Site.Role) As AMP.Role
        Get
            For Each r As AMP.Role In Me.InnerList
                If r.ID = roleID Then
                    Return r
                    Exit For
                End If
            Next
        End Get
    End Property

    Default Public ReadOnly Property ByRole(ByVal roleID As String) As AMP.Role
        Get
            Return Me.ByRole(CType(roleID, AMP.Site.Role))
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	return delimited list of roles, used primarily for output cache variance
    '
    '	Date:		Name:	Description:
    '	1/19/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Overrides Function ToString() As String
        If _roles Is Nothing Then
            Dim roles As New StringBuilder
            For Each role As AMP.Role In MyBase.InnerList
                If roles.Length > 0 Then roles.Append(",")
                roles.Append(role.Name)
            Next
            _roles = roles.ToString
        End If
        Return _roles
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get permissions for all roles in this collection
    '
    '	Date:		Name:	Description:
    '	12/7/04	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Permissions() As AMP.Site.Permission()
        Dim pc As New PermissionCollection
        Dim effective As AMP.Site.Permission()
        For Each role As AMP.Role In MyBase.InnerList
            effective = role.EffectivePermissions
            For x As Integer = 0 To effective.Length - 1
                pc.Add(effective(x))
            Next
        Next
        Return pc.All
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	12/7/04	    JEA		Creation
    '   1/19/05     JEA     Clear string when roles changed
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Role) As Integer
        _roles = Nothing
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Role)
        _roles = Nothing
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Remove(ByVal roleID As AMP.Site.Role)
        Me.Remove(Me.ByRole(roleID))
    End Sub

    Public Function Write() As String
        Dim output As New StringBuilder

        For Each r As AMP.Role In MyBase.InnerList
            output.Append(r.Name)
        Next
        Return output.ToString
    End Function
End Class
