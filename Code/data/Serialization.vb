Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary

Namespace Data
    Public Class Serialization

        Public Shared Function Clone(ByVal obj As Object) As Object
            Dim stream As New MemoryStream
            Dim formatter As New BinaryFormatter

            formatter.Serialize(stream, obj)
            stream.Seek(0, SeekOrigin.Begin)
            Dim copy As Object = formatter.Deserialize(stream)
            stream.Close()
            Return copy
        End Function

        Public Shared Function Serialize(ByVal entity As Object) As Byte()
            Dim stream As New MemoryStream
            Dim bf As New BinaryFormatter
            Dim bytes As Byte()

            bf.AssemblyFormat = Formatters.FormatterAssemblyStyle.Simple
            bf.Serialize(stream, entity)
            stream.Seek(0, SeekOrigin.Begin)
            bytes = stream.ToArray
            stream.Close()
            Return bytes
        End Function
    End Class
End Namespace