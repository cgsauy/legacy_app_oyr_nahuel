VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTraspaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Codigo As Long
Public Fecha As Date
Public LocalOrigen As Long
Public LocalIntermedio As Long
Public LocalDestino As Long
Public FechaEntregado As Date
Public Comentario As String
Public UsuarioInicial As Integer
Public UsuarioFinal As Integer
Public FModificacion As Date
Public FImpreso As Date
Public UsuarioReceptor As Integer
Public Terminal As Integer
Public Sucursal As Integer
Public Serie As String
Public Numero As Long
Public Anulado As Date
Public Remito As Long
Public Modificado As Date
Public ModificadoPor As Integer

Public Renglones As Collection

Public Function InsertarNuevoTraslado() As Long
    
    Cons = "INSERT INTO Traspaso (TraFecha, TraFModificacion, TraLocalOrigen, TraLocalIntermedio,  TraLocalDestino, TraComentario, " _
        & "TraFechaEntregado, TraUsuarioInicial, TraUsuarioFinal, TraRemito, TraFImpreso, TraUsuarioReceptor, TraSucursal) " _
        & "VALUES ('" & Format(Fecha, "yyyyMMdd hh:mm:ss") & "', GetDate(), " & Me.LocalOrigen
    Cons = Cons & ", " & IIf(Me.LocalIntermedio > 0, Me.LocalIntermedio, "Null") & ", " & Me.LocalDestino
    Cons = Cons & ", " & IIf(Trim(Comentario) <> vbNullString, "'" & Trim(Comentario) & "'", "Null")
    Cons = Cons & ", " & IIf(FechaEntregado <> "01/01/1901", "'" & Format(FechaEntregado, "yyyyMMdd hh:mm") & "'", "Null")
    Cons = Cons & ", " & UsuarioInicial
    Cons = Cons & ", " & IIf(UsuarioFinal > 0, UsuarioFinal, "Null")
    Cons = Cons & ", " & IIf(Remito > 0, Remito, "Null")
    Cons = Cons & ", " & IIf(FImpreso <> "01/01/1901", "'" & Format(FImpreso, "yyyyMMdd hh:mm") & "'", "Null")
    Cons = Cons & ", " & IIf(UsuarioReceptor > 0, UsuarioReceptor, "Null")
    Cons = Cons & ", " & IIf(Sucursal > 0, Sucursal, "Null")
    Cons = Cons & ")"
    cBase.Execute (Cons)
    
    'Saco el código del insertado.
    Cons = "Select MAX(TraCodigo) From Traspaso Where TraCodigo > 0" _
        & " And TraLocalOrigen = " & Me.LocalOrigen _
        & " And TraLocalDestino = " & Me.LocalDestino _
        & " AND TraUsuarioInicial = " & Me.UsuarioInicial
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not IsNull(RsAux(0)) Then InsertarNuevoTraslado = RsAux(0)
    RsAux.Close

End Function

Public Sub InsertarRenglonesEnBD()
    
    Dim oRen As clsRenglonTraspaso
    For Each oRen In Renglones
        Cons = "INSERT INTO RenglonTraspaso (RTrTraspaso, RTrArticulo, RTrEstado, RTrCantidad, RTrPendiente) VALUES ("
        Cons = Cons & Me.Codigo & ", " & oRen.Articulo & ", " & oRen.Estado & ", " & oRen.Cantidad & ", " & oRen.Pendientes & ")"
        cBase.Execute Cons
    Next
    
End Sub

Private Sub Class_Initialize()
    Set Renglones = New Collection
    Anulado = "01/01/1901"
    FechaEntregado = Anulado
    FModificacion = Anulado
    FImpreso = Anulado
    FImpreso = Anulado
End Sub
