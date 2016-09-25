Imports System
Imports System.Reflection
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.FormatterServices
Imports AMP.Data.Mapping

'---COMMENT---------------------------------------------------------------
'	Serialization surrogate and selector that allows any new version of a
'   serialiazable class to be loaded from an outdated serialized representation.
'   Any new class fields not present in the serialization data will be
'   left uninitialized. Removed fields are ignored (no exceptions).
'   Based on Bamboo Prevalence implemention (http://bbooprevalence.sourceforge.net/).
'
'	Date:		Name:	Description:
'	12/3/04     JEA		Creation
'-------------------------------------------------------------------------

Namespace Data
    Public Class Migration
        Implements ISerializationSurrogate
        Implements ISurrogateSelector

        Private _assemblyToMigrate As System.Reflection.Assembly

        Public Sub New(ByVal assemblyToMigrate As System.Reflection.Assembly)
            _assemblyToMigrate = assemblyToMigrate
        End Sub

#Region " ISerializationSurrogate "

        '---COMMENT---------------------------------------------------------------
        '	serialization not implemented
        '
        '	Date:		Name:	Description:
        '	12/3/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Sub GetObjectData(ByVal entity As Object, ByVal info As SerializationInfo, _
            ByVal context As StreamingContext) Implements ISerializationSurrogate.GetObjectData

            Throw New NotImplementedException
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	custom de-serialization
        '   entity is the object being serialized; info contains serialization data
        '   do custom mapping to migrate serialized data to new entity version
        '
        '	Date:		Name:	Description:
        '	12/3/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Public Function SetObjectData(ByVal entity As Object, ByVal info As SerializationInfo, _
            ByVal context As StreamingContext, ByVal selector As ISurrogateSelector) As Object _
            Implements ISerializationSurrogate.SetObjectData

            Dim members As MemberInfo() = GetSerializableMembers(entity.GetType)

            ' spin through members to do custom mapping
            For Each field As FieldInfo In members
                field.SetValue(entity, Me.GetValue(field, info))
            Next

            Return Nothing
        End Function

        '---COMMENT---------------------------------------------------------------
        '	initialize field with value of matching name in serialization info
        '   or according to custom mapping attribute
        '
        '	Date:		Name:	Description:
        '	12/17/04	JEA		Creation
        '   2/6/05      JEA     Changed order of checks for better logic
        '-------------------------------------------------------------------------
        Private Function GetValue(ByVal field As FieldInfo, ByVal info As SerializationInfo) As Object
            Dim value As Object
            Dim fieldType As Type = field.FieldType
            Dim create As Boolean = False

            If Me.HasField(field.Name, info) Then
                Try
                    value = info.GetValue(field.Name, fieldType)
                    If Not (value Is Nothing OrElse field.FieldType.IsInstanceOfType(value)) Then
                        ' convert type if changed in new member
                        value = Convert.ChangeType(value, fieldType)
                    End If
                    Return value
                Catch e As Exception
                    BugOut("Getting value for {0} resulted in ""{1}""", field.Name, e.Message)
                End Try
            End If

            ' if we made it here then check for mapping to new member
            For Each a As Attribute In field.GetCustomAttributes(False)
                If TypeOf a Is AMP.Data.Mapping Then
                    ' get mapping
                    Dim map As AMP.Data.Mapping = DirectCast(a, AMP.Data.Mapping)
                    Select Case map.Type
                        Case AMP.Data.Mapping.ChangeType.NewField
                            value = Activator.CreateInstance(fieldType)
                        Case AMP.Data.Mapping.ChangeType.Rename
                            value = info.GetValue(map.OldName, fieldType)
                    End Select
                    Return value
                End If
            Next

            ' if we're still here then simply try to create instance of missing field
            If Not fieldType.IsValueType Then
                'BugOut("Creating instance of {0}", fieldType.Name)
                Try
                    value = Activator.CreateInstance(fieldType)
                Catch e As Exception
                    'BugOut("Failed to create instance of {0} with {1}", fieldType, e.Message)
                End Try
            End If

            Return value
        End Function

        '---COMMENT---------------------------------------------------------------
        '	determine if serialized info has field of name
        '
        '	Date:		Name:	Description:
        '	12/17/04	JEA		Creation
        '-------------------------------------------------------------------------
        Private Function HasField(ByVal fieldName As String, ByVal info As SerializationInfo) As Boolean
            Dim infoEnum As SerializationInfoEnumerator = info.GetEnumerator

            While (infoEnum.MoveNext)
                If infoEnum.Name = fieldName Then Return True
            End While
            BugOut("No serialization data for field {0}", fieldName)
            Return False
        End Function

#End Region

#Region " ISurrogateSelector "

        Function GetNextSelector() As ISurrogateSelector Implements ISurrogateSelector.GetNextSelector
            Return Nothing
        End Function

        '---COMMENT---------------------------------------------------------------
        '	custom surrogate selection
        '   should return this custom surrogate only when type to convert is the
        '   given type
        '
        '	Date:		Name:	Description:
        '	12/3/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Function GetSurrogate(ByVal type As System.Type, ByVal context As StreamingContext, _
            ByRef selector As ISurrogateSelector) As ISerializationSurrogate _
            Implements ISurrogateSelector.GetSurrogate

            If type.Assembly Is _assemblyToMigrate Then
                selector = Me
                Return Me
            Else
                selector = Nothing
                Return Nothing
            End If
        End Function

        Sub ChainSelector(ByVal selector As System.Runtime.Serialization.ISurrogateSelector) _
            Implements ISurrogateSelector.ChainSelector

            Throw New NotImplementedException("ChainSelector not supported")
        End Sub

#End Region

    End Class
End Namespace