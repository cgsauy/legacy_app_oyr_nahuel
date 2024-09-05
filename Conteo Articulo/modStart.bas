Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public paEstadoSano As Long
Public prmPathApp As String

Enum enuEConteo
    Agendado = 0
    OK = 1
    paraRecuento = 2
    ParaVerificar = 3
End Enum

Type typConteo
    flagData As Boolean
    IDConteoAnterior As Long
    IDNuevoConteo As Long
    
    QStockLocalHoy As Long
    QStockAnteriorBueno As Long
    QDifAnteriorBueno As Long
    
    EstadoConteo As Integer
    EstadoAnteriorConteo As Integer
    
    FechaInicialRecuento As Date
    
    cLocal As Long
    cArticulo As Long
    cEstado As Long
    
End Type

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    Dim bAccesoOK As Boolean
    bAccesoOK = False
    bAccesoOK = miConexion.AccesoAlMenu("ContadorVerificador")
    If Not bAccesoOK Then bAccesoOK = miConexion.AccesoAlMenu("ContadorComun")
    
    If bAccesoOK Then
    
        If Not InicioConexionBD(miConexion.TextoConexion("comercio"), 45) Then End
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        CargoParametrosLocales
        CargoParametrosEntrada Trim(Command())
        frmConteo.Show vbModeless
        
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then
            MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "Usuario sin Autorización"
        End If
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocales()
On Error GoTo errCP

    cons = "Select * from Parametro Where ParNombre IN ( 'EstadoArticuloEntrega', 'PathApp', 'PlBalance' )"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            
            Case "estadoarticuloentrega": paEstadoSano = rsAux!ParValor
            
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto)
            
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    cons = "Select * from Terminal Where TerNombre = '" & miConexion.NombreTerminal & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        paCodigoDeTerminal = rsAux!TerCodigo
        paCodigoDeSucursal = rsAux!TerSucursal
    End If
    rsAux.Close
    
    Exit Sub

errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub


Public Function EndMain()
    On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Function

Private Function CargoParametrosEntrada(mPrms As String)
    On Error GoTo errCPE
    'Conteo | Local | Articulo
    If Trim(mPrms) = "" Then Exit Function
    
    Dim I As Integer, mIDConteo As Long
    Dim arrPrms() As String
    arrPrms = Split(mPrms, "|")
    
    mIDConteo = 0
    For I = LBound(arrPrms) To UBound(arrPrms)
        Select Case I
            Case 0      'Id Conteo
                    If Val(arrPrms(I)) <> 0 Then
                        mIDConteo = Val(arrPrms(I))
                        Exit For
                    End If
            
            Case 1      'Id Local
                    If Val(arrPrms(I)) <> 0 Then frmConteo.prmIDLocal = Val(arrPrms(I))
                    
            Case 2      'Id Articulo
                    If Val(arrPrms(I)) <> 0 Then frmConteo.prmIDArticulo = Val(arrPrms(I))
        End Select
        
    Next
    
    If mIDConteo <> 0 Then
        cons = "Select * from ConteoArticulo Where CArID = " & mIDConteo
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            frmConteo.prmIDLocal = rsAux!CArLocal
            frmConteo.prmIDArticulo = rsAux!CArArticulo
        End If
        rsAux.Close
    End If
    Exit Function
    
errCPE:
    clsGeneral.OcurrioError "Error al cargar los parámetros de entrada: " & mPrms, Err.Description
End Function


Public Function ValidoAccesoMnu(Menu As String, idUsr As Long) As Boolean

    On Error Resume Next
    ValidoAccesoMnu = False
    
    If idUsr = 0 Then Exit Function
    If Menu = "" Then Exit Function
    
    cons = " Select * from logdb.dbo.NivelPermiso, logdb.dbo.Aplicacion " _
            & " Where NPeNivel IN (Select UNiNivel from UsuarioNivel Where UNiUsuario = " & idUsr & ")" _
            & " And NPeAplicacion = AplCodigo" _
            & " And AplNombre = '" & Trim(Menu) & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then ValidoAccesoMnu = True
    rsAux.Close

End Function


