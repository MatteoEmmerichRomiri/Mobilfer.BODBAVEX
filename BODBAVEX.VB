Imports System.Data
Imports NTSInformatica.CLN__STD

Imports System
Public Class FRODBAVEX
  Inherits FRMDBAVEX

  Public _oClfAvex As CLFDBAVEX
  Public Property oClfAvex() As CLFDBAVEX
    Get
      If _oClfAvex Is Nothing Then _oClfAvex = CType(ocleAvex, CLFDBAVEX)
      Return _oClfAvex
    End Get
    Set(ByVal value As CLFDBAVEX)
      _oClfAvex = value
    End Set
  End Property

End Class

