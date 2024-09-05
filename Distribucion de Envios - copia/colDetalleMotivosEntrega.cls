VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfMotEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Envio As Long
Public Fecha As Date
Public Camion As Integer
Public Comentario As String
Public Modificado As Date
Public ModificadoPor As Integer

Public Motivos As Collection

Private Sub Class_Initialize()
    Set Motivos = New Collection
End Sub
