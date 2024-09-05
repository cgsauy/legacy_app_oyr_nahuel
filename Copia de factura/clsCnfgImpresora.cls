VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCnfgImpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const cnfgSeccion As String = "Configuracion"

Public Opcion As Byte
Public ImpresoraTickets As String

Public Sub CargarConfiguracion(ByVal cnfgApp As String, ByVal cnfgKey As String)

    Dim sConfig As String
    sConfig = GetSetting(cnfgApp, cnfgSeccion, cnfgKey, "")
    
    Dim vCnfg() As String
    If sConfig <> "" Then
        
        vCnfg = Split(sConfig, "|")
        
        If UBound(vCnfg) = 1 Then
            Dim opt As Byte
            Dim impr As Byte
            
            Opcion = Val(vCnfg(0))
            ImpresoraTickets = vCnfg(1)
            
        End If
    End If
    
End Sub

Public Sub GuardarConfiguracion(ByVal cnfgApp As String, ByVal cnfgKey As String)
Dim sConfig As String

    sConfig = Opcion & "|" & ImpresoraTickets
    SaveSetting cnfgApp, cnfgSeccion, cnfgKey, sConfig

End Sub

