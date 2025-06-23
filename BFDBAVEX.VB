Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System.Globalization


Imports System
Public Class CLFDBAVEX
  Inherits CLEDBAVEX

  Public _oClhAvex As CLHDBAVEX
  Public Property oClhAvex() As CLHDBAVEX
    Get
      If _oClhAvex Is Nothing Then _oClhAvex = CType(ocldAvex, CLHDBAVEX)
      Return _oClhAvex
    End Get
    Set(ByVal value As CLHDBAVEX)
      _oClhAvex = value
    End Set
  End Property

End Class