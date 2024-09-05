VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public ID As Integer
Public Nombre As String
Public Abreviacion As String
Public Pais As clsPais
Public TipoDocIdDGI As Integer
Public TipoDeCliente As TipoCliente

Private Sub Class_Initialize()
    Set Pais = New clsPais
End Sub
