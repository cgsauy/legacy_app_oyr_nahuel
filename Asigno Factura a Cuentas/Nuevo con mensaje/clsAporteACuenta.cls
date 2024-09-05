VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAporteACuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TipoCuentaAporte
    Personal = 1
    Colectivo = 2
End Enum


Public Enum TipoMovimientosAportes
    Aporte = 1
    Asignado
    Inhabilitado
    Rehabilitado
End Enum

Public MovimientoAporte As TipoMovimientosAportes
Public Fecha As Date
Public tipoCuenta As Cuenta
Public idCuenta As Long
Public Documento As Long
Public Importe As Long
Public Articulo As Long

Public Function InsertarAsignacion(ByVal connect As rdoConnection) As Long
Dim sQy As String

    sQy = "INSERT INTO AporteACuenta (AaCTipo, AaCFecha, AaCTipoCuenta, AaCIDCuenta, AAcDocumento, AAcImporte) " & _
        "VALUES (" & MovimientoAporte & ", GetDate(), " & tipoCuenta & ", " & idCuenta & ", " & Documento & ", " & Importe & ")"
    
    connect.Execute sQy

End Function

Public Sub EliminarAporte(ByVal connect As rdoConnection)
    
    Cons = "DELETE AporteACuenta " & _
        " WHERE AaCTipo = " & MovimientoAporte & _
        " AND AaCFecha = '" & Format(Me.Fecha, "yyyyMMdd HH:nn:ss") & "'" & _
        " AND AaCDocumento = " & Me.Documento & _
        " AND AaCIDCuenta = " & Me.idCuenta & _
        " AND AaCImporte = " & Me.Importe
    
    connect.Execute Cons
    
End Sub

Public Function SaldoCuentaPersonal(ByVal connect As rdoConnection, ByVal tipoCuenta As Byte, ByVal idCuenta As Long, ByVal IncluyoInhabilitado As Boolean) As Currency
    
    SaldoCuentaPersonal = 0
    Dim rsS As rdoResultset
    'prm = 1 retorno sólo los habilitados.
    'prm = 0 retorna el total de aportes (vencidos y no vencidos).
    Cons = "SELECT Sum(Saldo) FROM SaldoCtaPersonalRecibos(" & tipoCuenta & ", " & idCuenta & ", " & IIf(IncluyoInhabilitado, 2, 1) & ")"
    Set rsS = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsS.EOF Then
        If Not IsNull(rsS(0)) Then SaldoCuentaPersonal = rsS(0)
    End If
    rsS.Close
    
End Function








