Attribute VB_Name = "modStart"
Option Explicit

Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum

Public Enum TipoMovimientoEstado
    ARetirar = 1
    AEntregar = 2
    Reserva = 3
End Enum

Public prmTipoComentario As String

Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion
Public txtConexion As String

Public paEstadoARecuperar As Integer
Public paTipoCuotaContado As Long
Public paTipoFleteVentaTelefonica As Long
Public paCamionRetiroVisita As Long
Public paCobroEnEntrega As Boolean
Public paMCVtaTelefonica As Long
Public paClienteEmpresa As Long
Public paCoefFleteRetiro As Currency
Public paLocalCompañia As Long

Public paBD As String
Public gPathListados As String

Public Sub Main()
    
    On Error GoTo errMain
    Dim aValor As Long: aValor = 0
    
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.title) Then
        txtConexion = miConexion.TextoConexion("comercio")
        InicioConexionBD txtConexion
        
        CargoParametros
        CargoParametrosSucursal
        CargoParametrosImpresion paCodigoDeSucursal
        
        prj_GetPrinter False
        
        aValor = Val(Trim(Command()))
        frmSeguimiento.prmServicio = aValor
        frmSeguimiento.Show
        
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If

    Exit Sub

errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.title & Chr(vbKeyReturn) & _
                Err.Number & " - " & Err.Description
    End
End Sub

Private Sub CargoParametros()
    
    On Error Resume Next
    
    paBD = miConexion.RetornoPropiedad(bDB:=True)
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
    ChDir App.Path: ChDir ("..")
    ChDir (CurDir & "\REPORTES\")
    gPathListados = CurDir & "\"
        
    Cons = "Select * from Parametro " & _
            "Where ParNombre IN('estadoarticuloentrega', 'estadoarecuperar', 'monedapesos', 'tipocuotacontado', 'clienteempresa', 'serviciocoeffleteretiro', " & _
                                            "'LocalCompañia', 'TipComServSeguimiento', 'tipofleteventatelefonica', 'camionretirovisita', 'primerahoraenvio', 'ultimahoraenvio', 'mcvtatelefonica')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        Select Case (LCase(Trim(RsAux!ParNombre)))
            
            Case LCase("LocalCompañia"): paLocalCompañia = RsAux!ParValor
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "estadoarecuperar": paEstadoARecuperar = RsAux!ParValor
            
            Case "monedapesos": paMonedaPesos = RsAux!ParValor
            Case "tipocuotacontado": paTipoCuotaContado = RsAux!ParValor
            
            Case "tipofleteventatelefonica": paTipoFleteVentaTelefonica = RsAux!ParValor
            
            Case "camionretirovisita":
                paCamionRetiroVisita = RsAux!ParValor
                If Not IsNull(RsAux("ParTexto")) Then
                    paCobroEnEntrega = (Val(RsAux("ParTexto")) = 1)
                End If
            
            Case "primerahoraenvio": paPrimeraHoraEnvio = RsAux!ParValor
            Case "ultimahoraenvio": paUltimaHoraEnvio = RsAux!ParValor

            Case "mcvtatelefonica": paMCVtaTelefonica = RsAux!ParValor
            Case "clienteempresa": paClienteEmpresa = RsAux!ParValor
            
            Case "serviciocoeffleteretiro": paCoefFleteRetiro = RsAux!ParValor
            
            Case LCase("TipComServSeguimiento"): prmTipoComentario = Trim(RsAux!ParTexto)
            
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    If paCoefFleteRetiro = 0 Then paCoefFleteRetiro = 1
            
End Sub


Public Function f_GetEventos(ByVal sAux As String) As String
On Error Resume Next
    f_GetEventos = ""
    If InStr(1, sAux, "[", vbTextCompare) = 1 And InStr(1, sAux, "/", vbTextCompare) > 1 And InStr(1, sAux, ":", vbTextCompare) > 2 And InStr(1, sAux, "]", vbTextCompare) > 1 Then
        f_GetEventos = Mid(sAux, InStr(1, sAux, "[", vbTextCompare), InStr(InStr(1, sAux, "[", vbTextCompare) + 1, sAux, "]"))
    End If
End Function

Public Function f_QuitarClavesDelComentario(ByVal sComentario As String) As String
Dim sAux As String
    sAux = f_GetEventos(sComentario)
    If sAux <> "" Then
        f_QuitarClavesDelComentario = Replace(sComentario, sAux, "")
    Else
        f_QuitarClavesDelComentario = sComentario
    End If
End Function

Public Sub AddEvento(ByVal sKey As String, ByVal lIDServ As Long)
Dim rsEv As rdoResultset
Dim sAux As String, sMemo As String
    
    On Error GoTo errAE
    
    Cons = "Select * From Taller where TalServicio = " & lIDServ
    Set rsEv = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsEv.EOF Then
        MsgBox "El servicio no tiene ficha de taller, no se pueden agregar eventos.", vbCritical, "Atención"
    Else
        If Not IsNull(rsEv!TalComentario) Then
            If Len(Trim(rsEv!TalComentario)) + Len(sKey) + 10 > rsEv.rdoColumns("TalComentario").Size Then
                rsEv.Close
                MsgBox "El largo del comentario de taller no le permite agregar eventos.", vbExclamation, "Atención"
                Exit Sub
            End If
            sMemo = rsEv!TalComentario
        Else
            sMemo = ""
        End If
        If sMemo <> "" Then sAux = f_GetEventos(sMemo)
        sMemo = Replace(sMemo, sAux, "")
        sAux = Replace(sAux, ";;", ";")
        If sAux = "" Then
            sAux = "["
        Else
            sAux = Replace(Trim(sAux), "]", ";")
        End If
        sAux = sAux & sKey & ":" & Format(Date, "dd/mm/yy") & "]"
        
        rsEv.Edit
        rsEv!TalComentario = sAux & Trim(sMemo)
        rsEv.Update
    End If
    rsEv.Close
    Exit Sub
errAE:
    clsGeneral.OcurrioError "Error al agregar el evento.", Err.Description
End Sub

Public Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub


