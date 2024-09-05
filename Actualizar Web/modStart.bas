Attribute VB_Name = "modStart"
Option Explicit

Public Enum TUsuarioM
    Persona = 1
    Grupo = 2
    Terminal = 3
    Mail = 4
End Enum

'RDO para Access.
Public cAccess As rdoConnection       'Conexión a la Base de Datos
Public rsAcc As rdoResultset

'ActiveX
Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion

Public sFileErr As String
Public bCentinela As Boolean

'Parámetros
Public paFileInvoco As String
Public paWebStock As Boolean, paWebRelacion As Boolean
Public paIntraStock As Boolean, paIntraRelacion As Boolean
Public paFUltActualizacion As String
Public paDSNBD As String, paPathWeb As String, paPathIntra As String
Public paMonedaPesos As Long, paMonedaDolar As Long, paCuotaCtdo As Long
Public paTipoServ As Long, paEstSano As Long, paPorcFEmb As Currency
Public paCuotaMin As Currency
Public paStockParaXDias As Integer

Public Sub Main()
Dim aSucursal As String
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If Trim(Command()) = "" Then        'No es el Centinela.
        bCentinela = False
        If Not miConexion.AccesoAlMenu(App.Title) Then
            MsgBox "Ud. no posee permisos para la aplicación.", vbExclamation, "Centinela"
            Screen.MousePointer = 0
            End
            Exit Sub
        End If
    Else
        bCentinela = True
    End If
    
    If InicioConexionBD(miConexion.TextoConexion("comercio"), 45) Then
        CargoParametros
        frmActualizarWeb.Show
    End If
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub
Public Function InicioConexionAccess() As Boolean
On Error GoTo errICA
    InicioConexionAccess = False
    eBase.CursorDriver = rdUseOdbc
    Set cAccess = eBase.OpenConnection(paDSNBD)
    cAccess.QueryTimeout = 45
    InicioConexionAccess = True
    Exit Function
errICA:
    clsGeneral.OcurrioError "No se pudo abrir la base de datos de la web.", Err.Description, "Error de Conexión"
End Function
Private Sub CargoParametros()
    
    paWebStock = False: paWebRelacion = False
    paIntraStock = False: paIntraRelacion = False
    
    Cons = "Select * From Parametro Where ParNombre like 'Web%' Or " _
        & " ParNombre Like 'Moneda%' Or " _
        & " ParNombre = 'TipoArticuloServicio' Or " _
        & " ParNombre = 'estadoarticuloentrega' Or " _
        & " ParNombre Like 'articulopath%' Or " _
        & " ParNombre Like 'TipoCuota%' "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
                        
            Case "webdsn": paDSNBD = Trim(RsAux!ParTexto)
            Case "webfechaactualizada"
                paFUltActualizacion = Trim(RsAux("ParTexto"))
                paFUltActualizacion = Replace(paFUltActualizacion, "a.m.", "", , , vbTextCompare)
                paFUltActualizacion = Replace(paFUltActualizacion, "p.m.", "", , , vbTextCompare)
            
            Case "articulopathejecutararchivo": paFileInvoco = Trim(RsAux("ParTexto"))
            Case "articulopathpageweb"
                paPathWeb = Trim(RsAux!ParTexto)
                If Not IsNull(RsAux!ParValor) Then
                    Select Case RsAux!ParValor
                        Case 1
                            paWebStock = True
                        Case 1.1
                            paWebStock = True
                            paWebRelacion = True
                        Case 0.1
                            paWebRelacion = True
                    End Select
                End If
                
            Case "articulopathpageintra"
                paPathIntra = Trim(RsAux!ParTexto)
                If Not IsNull(RsAux!ParValor) Then
                    Select Case RsAux!ParValor
                        Case 1
                            paIntraStock = True
                        Case 1.1
                            paIntraStock = True
                            paIntraRelacion = True
                        Case 0.1
                            paIntraRelacion = True
                    End Select
                End If
            Case LCase("webStockParaXDias"): paStockParaXDias = RsAux("ParValor")
            
            Case "monedapesos": paMonedaPesos = RsAux("ParValor")
            Case "monedadolar": paMonedaDolar = RsAux("ParValor")
            Case "webminimportecuota": paCuotaMin = RsAux("ParValor")
            Case "tipocuotacontado": paCuotaCtdo = RsAux("ParValor")
            Case "tipoarticuloservicio": paTipoServ = RsAux("ParValor")
            Case "estadoarticuloentrega": paEstSano = RsAux("ParValor")
            Case "webporcarriboembarque": paPorcFEmb = RsAux("ParValor")
            
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Cons = "Select * From logdb.dbo.logdbparametro where parnombre = 'comMsgPathFileError'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        sFileErr = Trim(RsAux("ParValorTexto"))
    End If
    RsAux.Close
    If Trim(sFileErr) <> "" Then
        If Right(sFileErr, 1) <> "\" Then sFileErr = sFileErr & "\"
        sFileErr = sFileErr & "ActWebError.Txt"
    End If
    
End Sub
