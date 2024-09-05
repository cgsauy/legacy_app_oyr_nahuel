VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LocalRepara As clsCodigoTexto
Public Vias As Byte
Public Aclaracion As String
Public Motivos As Collection
Public IDNuevoServicio As Long

Private Sub Class_Initialize()
    Set Motivos = New Collection
    Set LocalRepara = New clsCodigoTexto
End Sub

Private Sub Class_Terminate()
    Set Motivos = Nothing
    Set LocalRepara = Nothing
End Sub
