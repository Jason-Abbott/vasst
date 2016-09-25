Namespace Data

    '---COMMENT---------------------------------------------------------------
    '   custom attribute to define deserialization mapping
    '   example: <AMP.Data.Mapping(Mapping.ChangeType.Rename, "_name")> _
    '
    '	Date:		Name:	Description:
    '	1/18/05	    JEA		Creation
    '-------------------------------------------------------------------------
    <AttributeUsage(AttributeTargets.Field)> _
    Public Class Mapping
        Inherits System.Attribute

        Private _type As Mapping.ChangeType
        Private _oldName As String  ' for renamed fields

#Region " Properties "

        Public ReadOnly Property Type() As Mapping.ChangeType
            Get
                Return _type
            End Get
        End Property

        Public ReadOnly Property OldName() As String
            Get
                Return _oldName
            End Get
        End Property

#End Region

        Public Enum ChangeType
            Rename                  ' renamed field
            NewField                ' new field
        End Enum

        Public Sub New(ByVal type As Mapping.ChangeType, ByVal oldName As String)
            Me.New(type)
            _oldName = oldName
        End Sub

        Public Sub New(ByVal type As Mapping.ChangeType)
            _type = type
        End Sub

    End Class
End Namespace