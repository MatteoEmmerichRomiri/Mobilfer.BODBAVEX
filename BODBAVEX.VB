Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System.Text.RegularExpressions

Imports System
Imports DevExpress.XtraBars

Public Class FRODBAVEX
    Inherits FRMDBAVEX

#Region "LIVELLI"

    Public _oClfAvex As CLFDBAVEX
    Public Property oClfAvex() As CLFDBAVEX
        Get
            If _oClfAvex Is Nothing Then _oClfAvex = CType(oCleAvex, CLFDBAVEX)
            Return _oClfAvex
        End Get
        Set(ByVal value As CLFDBAVEX)
            _oClfAvex = value
        End Set
    End Property

#End Region

#Region "VARIABILI"

    'Public rigaDocumentoDaStampare As Integer = 0

#End Region

    Public Overrides Sub FRMDBAVEX_Load(sender As Object, e As EventArgs)
        Try

            MyBase.FRMDBAVEX_Load(sender, e)
            If oClfAvex IsNot Nothing Then
                AddHandler oClfAvex.StampaCaricoDiPoduzione, AddressOf StampaCaricoDiProduzione
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

#Region "STAMPA CARICO DI PRODUZIONE"
    Private Sub StampaCaricoDiProduzione()
        Try

            ' Inizializzazione della riga corrente
            Dim currentRow As DataRow = Nothing

            ' Verifica che ci siano righe selezionate e una riga dati corrente valida
            If grvAvex.SelectedRowsCount = 0 OrElse
               grvAvex.DataRowCount = 0 OrElse
               grvAvex.NTSGetCurrentDataRow() Is Nothing Then

                ' Uscita anticipata se non ci sono righe valide
                Return
            End If

            '' Recupera la riga dati corrente
            'currentRow = grvAvex.NTSGetCurrentDataRow()



            'Dim dttCorpoRigaSelezionata As New DataTable
            'If Not oClfAvex.GetCorpoCaricoDiProduzioneInterna(currentRow, caricoProduzioneInternaCreato, dttCorpoRigaSelezionata) Then Return

            'For Each row As DataRow In dttCorpoRigaSelezionata.Rows

            'rigaDocumentoDaStampare = NTSCInt(row!mm_riga)

            Stampa(0)
            'Next

            '' Esegue la stampa
            'Stampa(0)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub


    Public Overridable Sub Stampa(ByVal nDestin As Integer)
        Dim nPjob As Object
        Dim nRis As Integer = 0
        Dim strCrpe As String = ""
        Dim i As Integer
        Dim Modello As String

        Me.ValidaLastControl()


        Try

            strCrpe =
                "{movmag.codditt}   = " & CStrSQL(DittaCorrente) & " AND " &
                "{movmag.mm_tipork} = " & CStrSQL(oClfAvex.caricoProduzioneInternaCreato.tipork) & " AND " &
                "{movmag.mm_anno}   = " & oClfAvex.caricoProduzioneInternaCreato.anno & " AND " &
                "{movmag.mm_serie}  = " & CStrSQL(oClfAvex.caricoProduzioneInternaCreato.serie) & " AND " &
                "{movmag.mm_numdoc} = " & oClfAvex.caricoProduzioneInternaCreato.numdoc
            '"{movord.mo_riga}   = " & rigaDocumentoDaStampare  'se ti serve anche la riga

            Modello = "PAR5110.rpt"

            '--------------------------------------------------------------------------------------------------------------
            nPjob = oMenu.ReportPEInit(oApp.Ditta, Me, "BSDBAVEX", "PAR5110.rpt", " ", 0, nDestin, Modello, False, "Etichette", False)
            '--------------------------------------------------------------------------------------------------------------
            If nPjob Is Nothing Then Return
            '--------------------------------------------------------------------------------------------------------------
            For i = LBound(CType(nPjob, Array), 2) To UBound(CType(nPjob, Array), 2)
                nRis = oMenu.PESetSelectionFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), CrpeResolveFormula(Me, CStr(CType(nPjob, Array).GetValue(2, i)), strCrpe))
                nRis = oMenu.ReportPEVai(NTSCInt(CType(nPjob, Array).GetValue(0, i)))
            Next
            '--------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

#End Region

#Region "GESTIONE EVENTI ENTITY"

    Public Overrides Sub GestisciEventiEntity(ByVal sender As Object, ByRef e As NTSEventArgs)
        Try
            'If String.IsNullOrWhiteSpace(e.Message) Then Return

            'Select Case True

            '    Case Regex.IsMatch(e.Message, "Elaborazione completata\..*?carichi\s+da\s+produzione\s+creati.*?\(n[°º]?\s*\d+\)",
            '                       RegexOptions.IgnoreCase Or RegexOptions.Singleline)

            '        'Estrai l’ID fra parentesi tonde
            '        numeroCaricoProduzioneInternaCreato = CInt(Regex.Match(e.Message, "\(n[°º]?\s*(\d+)\)",
            '                                                               RegexOptions.IgnoreCase).Groups(1).Value)


            '    Case Regex.IsMatch(e.Message, "Carico di produzione generato.*Aprire il carico di produzione",
            '                       RegexOptions.IgnoreCase Or RegexOptions.Singleline)

            '        StampaCaricoDiProduzione()
            'End Select

            MyBase.GestisciEventiEntity(sender, e)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

#End Region

End Class

