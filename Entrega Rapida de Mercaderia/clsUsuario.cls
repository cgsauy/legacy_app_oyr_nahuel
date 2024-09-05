VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Codigo As Long
Public Digito As Integer
Public Identificacion As String
Public Habilitado As Boolean

Public Function BuscarDigito(ByVal iDigito As Long) As Boolean
    With Me
        .Codigo = 0
        .Digito = iDigito
        .Habilitado = False
        .Identificacion = ""
    End With
    BuscarDigito = modPersistencia.Usuario_Buscar(Me)
End Function

Public Function BuscarCodigo(ByVal iCodigo As Long) As Boolean
    With Me
        .Codigo = iCodigo
        .Digito = 0
        .Habilitado = False
        .Identificacion = ""
    End With
    BuscarCodigo = modPersistencia.Usuario_Buscar(Me)
End Function



