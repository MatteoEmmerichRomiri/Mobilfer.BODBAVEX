Imports System.Text

Public Module StringBuilderExtensions

#Region "METODI APPEND LINE"

    <System.Runtime.CompilerServices.Extension()>
    Public Sub AppendLineFormat(ByRef sb As StringBuilder, format As String, ParamArray args() As Object)
        Try

            sb.AppendLine(String.Format(format, args) & " ")

        Catch ex As Exception
            Dim strErr As String = $"Errore generato nel modulo {NameOf(StringBuilderExtensions)} " &
                         $"nel Metodo {NameOf(AppendLineFormat)}. " &
                         $"Errore: " & ex.Message
            Throw New Exception(strErr)
        End Try

    End Sub

#End Region

End Module

