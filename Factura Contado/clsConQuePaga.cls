VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConQuePaga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eVolcarSaldo
    Aporte = 1
    Pendiente
End Enum

'Public Enum eDocumentosPaga
'    TransaccionRP = 1   'Transacción redpagos.
'    AportesACuenta = 2
'    Pendientes = 3
'End Enum

Public Enum eDocQueSeCobra
    ContadooCuota = 1
    Aporte = 2
    PendienteNegativo = 3
    LiquidacionPendiente = 4
End Enum

Public Enum eDocConQuePaga
    AsignoAporteCta = 12
    LiquidaPendienteNeg = 13
    PendienteCajaNegativoNuevo = 14
    GiroRedPagos = 15
End Enum


Public TipoConQuePaga As eDocConQuePaga
'ID depende del tipo: 1 = TraID (comTransacciones), 2 = idCliente, 3 = IDPendiente.
Public IDDocumentoPaga As Long
Public importe As Currency



