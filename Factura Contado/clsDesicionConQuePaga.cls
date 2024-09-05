VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDesicionConQuePaga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public VolcarSaldoAFavor As eVolcarSaldo
Public SaldoAFavor As Currency
Public ConQuePaga As Collection    'Objetos de clsConQuePaga

'Documento al cual se vuelcan los aportes (ctdo, recibo o pendiente)
Public TipoDocQueSalda As eDocQueSeCobra
Public DocumentoQueSalda As Long

Public Sub GrabarAportes()

    '1) Inserto los aportes existentes.
    InsertoAportesAsignados
    
    '2) Si tengo saldo lo asigno.
    If SaldoAFavor <> 0 Then
        
    End If
    
End Sub

Sub InsertoAporteACuenta(ByVal CQP As clsConQuePaga)
Dim sQY As String
    
    sQY = "Insert Into CuentaDocumento (CDoTipo, CDoIDTipo, CDoIDDocumento, CDoAsignado) Values ("
    cBase.Execute sQY
    
    MovimientoDeCaja paMCSenias, gFechaServidor, m_Disponibilidad, cMoneda.ItemData(cMoneda.ListIndex), CCur(labSalida.Caption), strDocumento, True
    Cons = "Select Max(MDIid) from MovimientoDisponibilidad"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aMovimientoCaja = RsAux(0)
    RsAux.Close
    
End Sub

Public Sub InsertoAportesAsignados()
Dim oCQP As clsConQuePaga
    
    'recorro todos los documentos asignados y defino a que se lo doy.
    For Each oCQP In ConQuePaga
        Select Case oCQP.TipoConQuePaga
            
            Case eDocConQuePaga.AsignoAporteCta
                InsertoAporteACuenta oCQP
            
            Case eDocConQuePaga.GiroRedPagos
                InsertoTransaccionRP oCQP
        End Select
    Next
End Sub

Sub InsertoTransaccionRP(ByVal oCQP As clsConQuePaga)
Dim sQY As String
    
    '1) cambio el estado de la transacción.
    sQY = "UPDATE ComTransaccion SET TraEstado = 9 WHERE TraID = " & oCQP.IDDocumentoPaga
    cBase.Execute sQY
    
    '2) Inserto el conquepaga.
    sQY = "INSERT INTO [CGSA].[dbo].[ConQueCobra] ([CQCTipoQueCobra] ,[CQCIdQueCobra] ,[CQCTipoQuePaga] ,[CQCIdQuePaga])" _
        & " VALUES (" & TipoDocQueSalda & ", " & Me.DocumentoQueSalda & ", " & oCQP.TipoConQuePaga & ", " & oCQP.IDDocumentoPaga & ")"
    cBase.Execute sQY

End Sub
