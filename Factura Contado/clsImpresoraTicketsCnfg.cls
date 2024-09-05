VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImpresoraTicketsCnfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
Public ImpresoraTickets As String

'Private Const cnfgApp As String = "ImpresionDocumentos"
Private Const cnfgSeccion As String = "Configuracion"
'Private Const cnfgKey As String = "CuotasImpresora"

Public Sub CargarConfiguracion(ByVal cnfgApp As String, ByVal KeyConfiguracion As String)

    Dim sConfig As String
    'cnfgKey
    sConfig = GetSetting(cnfgApp, cnfgSeccion, KeyConfiguracion, "")
    
    Dim vCnfg() As String
    If sConfig <> "" Then
        
        vCnfg = Split(sConfig, "|")
        
        If UBound(vCnfg) = 1 Then
            Dim opt As Byte
            Dim impr As Byte
            
            Opcion = Val(vCnfg(0))
            If IsNumeric(vCnfg(1)) Then
                ImpresoraTickets = Val(vCnfg(1))
            Else
                ImpresoraTickets = vCnfg(1)
            End If
            
        End If
    End If
    
End Sub

Public Sub GuardarConfiguracion(ByVal cnfgApp As String, ByVal KeyConfiguracion As String)
Dim sConfig As String

    sConfig = Opcion & "|" & ImpresoraTickets
    SaveSetting cnfgApp, cnfgSeccion, KeyConfiguracion, sConfig

End Sub
