Imports System.Text

Public Module ChiaviDiRegistro

#Region "GET AND SET"

    Public Function setChiaveRegistro(Of T)(ByRef oMenu As CLE__MENU, ByRef DittaCorrente As String, ByRef progettoCDR As String,
                                             pathProgettoCDR As String, ByRef _strChiave As T, nomeChiave As String,
                                             ByRef mySettingChiave As String) As Boolean
        Try
            Dim valore As String = oMenu.GetSettingBusDitt(DittaCorrente, progettoCDR, pathProgettoCDR, ".", nomeChiave, " ", " ", mySettingChiave)

            ' Controllo per evitare errori di conversione
            valore = If(String.IsNullOrWhiteSpace(valore), mySettingChiave, valore)

            ' Conversione generica per tipi che implementano IConvertible
            _strChiave = CType(Convert.ChangeType(valore, GetType(T)), T)

            Return True

        Catch ex As Exception
            Dim strErr As String = $"Errore generato nel modulo {NameOf(ChiaviDiRegistro)} " &
                         $"nel Metodo {NameOf(setChiaveRegistro)}. " &
                         $"Errore: " & ex.Message
            Throw New Exception(strErr)
        End Try
    End Function


#End Region

End Module

