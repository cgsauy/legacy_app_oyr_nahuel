VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLiquidacionCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Liquidacion As Long
Public Moneda As Integer
Public TotalLiquidacion As Currency
Public TotalCobrado As Currency



'SELECT LiqID, LCoMoneda, LiqTotal, SUM(LCoCobrado)
'From Liquidacion
'INNER JOIN LiquidacionCobro ON LCoLiquidacion = LiqID
'WHERE LiqTipo = 1 --AND LiqEnte = idCamion
'GROUP BY LiqID, LCoMoneda, LiqTotal
'Having (IsNull(Sum(LCoCobrado), 0) <> LiqTotal)
