VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFichasDevolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Articulo As Long
Public Fichas As Collection

Public Sub AddFicha(ByVal ficha As Long, ByVal Cantidad As Integer)
    Dim oArtF As New clsArtEnFicha
    oArtF.Cantidad = Cantidad
    oArtF.IDFicha = ficha
    Fichas.Add oArtF
End Sub

Public Function CantidadTotalEnFichas() As Integer
    CantidadTotalEnFichas = 0
    Dim oArtF As clsArtEnFicha
    For Each oArtF In Fichas
        CantidadTotalEnFichas = CantidadTotalEnFichas + oArtF.Cantidad
    Next
End Function

Private Sub Class_Initialize()
    Set Fichas = New Collection
End Sub

Private Sub Class_Terminate()
    Set Fichas = Nothing
End Sub
