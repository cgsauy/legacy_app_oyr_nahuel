VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Public Tipo As Byte
Public Categoria As Long
Public Nombre1 As String
Public Nombre2 As String
Public Apellido1 As String
Public Apellido2 As String
Public Cedula As String
Public RUC As String
Public Cheque As Boolean


Public Function MostrarComo() As String

End Function
