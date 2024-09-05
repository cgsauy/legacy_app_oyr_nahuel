VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eTipoVenta
    NoHay = 0
    Web = 1
    Telefonica
End Enum

Public idTransaccion As Long
Public idCliente As Long
Public TipoCliente As Long

Public TipoVenta As eTipoVenta
Public idVenta As Long

Public Envios As String
Public Articulos As Collection

Private Sub Class_Initialize()
    Set Articulos = New Collection
End Sub

Private Sub Class_Terminate()
    Set Articulos = Nothing
End Sub
