VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstadosIngMerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Estados As Collection

Public Function CargarEstados(ByVal rdoBase As rdoConnection)
    Set Estados = New Collection
    Dim rsEst As rdoResultset
    Dim oEstado As clsEstadoMercaderia
'    Cons = "SELECT CodID, rTrim(CodTexto), CodValor3" & _
'            " FROM Codigos WHERE CodCual = 45 AND CodValor1 = 1 ORDER BY CodValor2"

    Cons = "SELECT Codigo CodID, rtrim(Texto), Valor1 FROM CodigoTexto WHERE Tipo = 45 ORDER BY PUNTAJE"
    Set rsEst = rdoBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsEst.EOF
        Set oEstado = New clsEstadoMercaderia
        oEstado.Estado = rsEst(2)
        oEstado.ID = rsEst(0)
        oEstado.Nombre = rsEst(1)
        
        Estados.Add oEstado
        rsEst.MoveNext
    Loop
    rsEst.Close
    CargarEstados = True
End Function

Private Function EstadoSeleccionado(ByVal ID As Long, ByVal ids As String) As Boolean
Dim iQ As Integer
Dim vSelect() As String
    EstadoSeleccionado = False
    vSelect = Split(ids, ":")
    For iQ = 0 To UBound(vSelect)
        If UBound(vSelect) = 0 Then
            If Val(vSelect(iQ)) = ID Then
                EstadoSeleccionado = True
                Exit Function
            End If
        ElseIf LCase(Mid(vSelect(iQ), 1, 1)) = "s" And Val(Mid(vSelect(iQ), 2)) = ID Then
            EstadoSeleccionado = True
            Exit Function
        End If
    Next
End Function

Public Function ObtenerStringEstado(ByVal ids As String) As String
    If Estados.Count = 0 Then Exit Function
    Dim bSelect As Boolean
    If Len(ids) > 0 Then
        Dim oEstado As clsEstadoMercaderia
        For Each oEstado In Estados
            bSelect = EstadoSeleccionado(oEstado.ID, ids)
            ObtenerStringEstado = ObtenerStringEstado & _
                            IIf(ObtenerStringEstado = "", "", ":") & IIf(bSelect, "s", "n") & oEstado.ID
            
        Next
    End If
End Function

'Public Function ObtenerValorEstadoSeleccionado(ByVal idsseleccionados As String) As Long
''Recibe un string separado con coma de aquellos valores que fueron seleccionados.
''Tengo que validar c/u de ellos
'Dim vSelect() As String
'Dim iQ As Byte
'
'    ObtenerValorEstadoSeleccionado = 0
'    If idsseleccionados = "" Then Exit Function
'    vSelect = Split(idsseleccionados, ",")
'    For iQ = 0 To UBound(vSelect)
'        If Val(vSelect(iQ)) > 0 Then
'            ObtenerValorEstadoSeleccionado = ObtenerValorEstadoSeleccionado + (2 ^ (Val(vSelect(iQ)) - 1))
'        End If
'    Next
'    Exit Function
'End Function

'Public Function ObtenerCadenaEstadosSeleccionados(ByVal Valor As Long) As String
'Dim ids As String
'    If Estados.Count = 0 Then Exit Function
'    ids = cDec2Bin(Valor)
'    If Len(ids) > 0 Then
'        Dim oEstado As clsEstadoMercaderia
'        For Each oEstado In Estados
'            If oEstado.ID <= Len(ids) Then
'                If Mid(Right(ids, oEstado.ID), 1, 1) = "1" Then
'                    ObtenerCadenaEstadosSeleccionados = ObtenerCadenaEstadosSeleccionados & _
'                                IIf(ObtenerCadenaEstadosSeleccionados = "", "", ", ") & oEstado.Nombre
'                End If
'            End If
'        Next
'    End If
'End Function

Public Function ObtenerCadenaEstadosSeleccionados(ByVal ids As String) As String
    
    If Estados.Count = 0 Or ids = "" Then Exit Function
    
    Dim vSelect() As String
    vSelect = Split(ids, ",")
        
    Dim oEstado As clsEstadoMercaderia
    Dim bSelect As Boolean
    For Each oEstado In Estados
        bSelect = EstadoSeleccionado(oEstado.ID, ids)
        If bSelect Then
            ObtenerCadenaEstadosSeleccionados = ObtenerCadenaEstadosSeleccionados & _
                        IIf(ObtenerCadenaEstadosSeleccionados = "", "", ", ") & oEstado.Nombre
        End If
    Next
   
End Function


Public Function EstadoEsARecuperar(ByVal ids As String) As Boolean
    
    EstadoEsARecuperar = False
    If Estados.Count = 0 Or ids = "" Then Exit Function
    
    Dim vSelect() As String
    vSelect = Split(ids, ",")
    
    Dim iQ As Integer
    Dim bClick As Boolean
    Dim oEstado As clsEstadoMercaderia
    For Each oEstado In Estados
        bClick = False
        For iQ = 0 To UBound(vSelect)
            If oEstado.ID = Val(vSelect(iQ)) Then
                bClick = True
                'Está clickeado entonces me tengo que fijar el estado.
                If oEstado.Estado = 1 Then
                    EstadoEsARecuperar = True
                    Exit Function
                Else
                    Exit For
                End If
            End If
        Next
        If Not bClick Then
            'No está clickeado el estado si cumple la condición me salgo.
            If oEstado.Estado = -1 Then
                EstadoEsARecuperar = True
                Exit Function
            End If
        End If
    Next
    
End Function


'Public Function EstadoEsARecuperar(ByVal valor As Long) As Boolean
'Dim ids As String
'
'    EstadoEsARecuperar = False
'    If Estados.Count = 0 Then Exit Function
'
'    ids = cDec2Bin(valor)
'    If Len(ids) < ObtenerMayorID Then ids = String(ObtenerMayorID - Len(ids), "0") & ids
'
'    ids = StrReverse(ids)
'
'    If Len(ids) > 0 Then
'
'        Dim oEstado As clsEstadoMercaderia
'        For Each oEstado In Estados
'            Select Case oEstado.estado
'
'                Case -1
'                    'Si no está clickeado va a recuperar.
'                    If Mid(ids, oEstado.ID, 1) = 0 Then
'                        EstadoEsARecuperar = True
'                        Exit Function
'                    End If
'
'                Case 1
'                    'Si está clickeado y es 1 entonces va a recuperar.
'                    If Mid(ids, oEstado.ID, 1) = 1 Then
'                        EstadoEsARecuperar = True
'                        Exit Function
'                    End If
'
'            End Select
'        Next
'
'    End If
'End Function

Private Function cBin2Dec(ByVal value As String) As Long
Dim retorno As Long, iQ As Integer, elevado As Byte
    
    elevado = 0
    For iQ = Len(value) To 1 Step -1
        Select Case Asc(Mid(value, iQ, 1))
            Case 48      ' "0" no hago nada
            Case 49      ' "1", incremento elevado a su posición.
                retorno = retorno Or 2 ^ elevado
            Case Else
                Err.Raise 5      ' Invalid procedure call or argument
        End Select
        elevado = elevado + 1
    Next
    cBin2Dec = retorno
End Function

Private Function cDec2Bin(ByVal value As Long) As String
    Do While value > 0
        cDec2Bin = IIf(value Mod 2, "1", "0") & cDec2Bin
        value = value \ 2
    Loop
End Function

Private Function ObtenerMayorID() As Byte
    ObtenerMayorID = 0
    Dim oEstado As clsEstadoMercaderia
    For Each oEstado In Estados
        If ObtenerMayorID < oEstado.ID Then ObtenerMayorID = oEstado.ID
    Next
End Function

Private Sub Class_Initialize()
    Set Estados = New Collection
End Sub

Private Sub Class_Terminate()
    Set Estados = Nothing
End Sub
