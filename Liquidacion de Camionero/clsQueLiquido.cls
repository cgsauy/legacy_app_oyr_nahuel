VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueLiquido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum TipoQueCobro
    EnviosMerc = 1
    ServicioEntRet = 2
    ParadasDeTraslados = 3
    LiquidacionesPendientes = 4
End Enum

Public ID As Long       'ID servicio, envío, etc.
Public TipoID As TipoQueCobro
Public Reclamar As Currency
Public Liquidar As Currency
Public PendienteCaja As Currency
Public Diferencia As Currency
Public SeCobra As Boolean
Public ImporteDocumentoPendiente As Currency
