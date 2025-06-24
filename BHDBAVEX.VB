Imports System.Data
Imports System.Data.Common
Imports NTSInformatica.CLN__STD

Imports System
Imports System.Text

Public Class CLHDBAVEX
    Inherits CLDDBAVEX

#Region "GET CORPO DOCUMENTO"

    Friend Function GetCorpoDocumento(strDittaCorrente As String, doc As DocumentoDiMagazzino, ByRef dttOut As DataTable) As Boolean
        Try
            Dim sSql As New StringBuilder

            sSql.AppendLine("SELECT *")
            sSql.AppendLine(" FROM MOVMAG")
            sSql.AppendLine(String.Format("WHERE codditt   = '{0}'", strDittaCorrente))
            sSql.AppendLine(String.Format("  AND mm_tipork = '{0}'", doc.tipork))
            sSql.AppendLine(String.Format("  AND mm_anno  = {0}", doc.anno))
            sSql.AppendLine(String.Format("  AND mm_serie = '{0}'", doc.serie))
            sSql.AppendLine(String.Format("  AND mm_numdoc = {0}", doc.numdoc))

            dttOut = OpenRecordset(sSql.ToString(), CLE__APP.DBTIPO.DBAZI)

            If dttOut Is Nothing OrElse dttOut.Rows.Count = 0 Then
                Return False
            End If

            Return True

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, ex.Message)
        End Try
    End Function

#End Region

End Class