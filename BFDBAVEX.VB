Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System.Globalization


Imports System
Public Class CLFDBAVEX
    Inherits CLEDBAVEX

#Region "LIVELLI"

    Public _oClhAvex As CLHDBAVEX
    Public Property oClhAvex() As CLHDBAVEX
        Get
            If _oClhAvex Is Nothing Then _oClhAvex = CType(oCldAvex, CLHDBAVEX)
            Return _oClhAvex
        End Get
        Set(ByVal value As CLHDBAVEX)
            _oClhAvex = value
        End Set
    End Property

#End Region

#Region "VARIABILI"
    Public caricoProduzioneInternaCreato As New DocumentoDiMagazzino
#End Region

#Region "EVENTI"
    Public Event StampaCaricoDiPoduzione()
#End Region

#Region "STAMPA CARICO DI PRODUZIONE"
    Public Overrides Sub RecordVersa_AfterSalvaDoc(ByRef oCleBoll As CLEVEBOLL)
        Try
            MyBase.RecordVersa_AfterSalvaDoc(oCleBoll)

            Dim dtr As DataRow = oCleBoll.dttET.Rows(0)

            caricoProduzioneInternaCreato = New DocumentoDiMagazzino With {
                .tipork = NTSCStr(dtr!et_tipork),
                .anno = NTSCInt(dtr!et_anno),
                .serie = NTSCStr(dtr!et_serie),
                .numdoc = NTSCInt(dtr!et_numdoc)
            }

            RaiseEvent StampaCaricoDiPoduzione()

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, ex.Message)
        End Try
    End Sub
#End Region

#Region "GET CORPO DOCUMENTO"

    Friend Function GetCorpoCaricoDiProduzioneInterna(currentRow As DataRow, caricoDiProduzione As DocumentoDiMagazzino, ByRef dttCorpoRigaSelezionata As DataTable) As Boolean
        Try
            If currentRow Is Nothing Then Return False

            If caricoDiProduzione Is Nothing OrElse caricoDiProduzione.numdoc = 0 Then Return False

            If dttCorpoRigaSelezionata Is Nothing Then
                dttCorpoRigaSelezionata = New DataTable
            End If

            Return oClhAvex.GetCorpoDocumento(strDittaCorrente, caricoDiProduzione, dttCorpoRigaSelezionata)

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, ex.Message)
        End Try
    End Function

#End Region

End Class