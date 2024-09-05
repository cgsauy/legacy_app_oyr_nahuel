VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPropFieldPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Width As String              'toma un string ya que en la tabla el width es el largo de c/col
Public FontSize As Single
Public Bold As Boolean
Public Align As Byte
Public TablaCampos As String
Public Posicion As clsRegPosXY

Private Sub Class_Initialize()
    Set Posicion = New clsRegPosXY
End Sub

Private Sub Class_Terminate()
    Set Posicion = Nothing
End Sub
