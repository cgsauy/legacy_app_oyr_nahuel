VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmbarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Public Agencia As Long
Public ArriboPrevisto As Date
Public Arribo As Date
Public Embarco As Date
Public Transporte As Long

Public LineaAsignada As Long

Private Sub Class_Initialize()
    ArriboPrevisto = DateMinValue
    Arribo = DateMinValue
    Embarco = DateMinValue
End Sub
