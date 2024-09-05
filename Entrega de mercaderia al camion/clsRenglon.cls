VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRenglon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Articulo As Long
Public ARetirar As Integer
Public NumerosDeSerie As Collection

Private Sub Class_Initialize()
    Set NumerosDeSerie = New Collection
End Sub

Private Sub Class_Terminate()
    Set NumerosDeSerie = Nothing
End Sub

Public Function SerieDuplicada(ByVal nroSerie As String) As Boolean
    
    Dim serie As Variant
    For Each serie In NumerosDeSerie
        If serie = nroSerie Then
            SerieDuplicada = True
            Exit Function
        End If
    Next
    
End Function
