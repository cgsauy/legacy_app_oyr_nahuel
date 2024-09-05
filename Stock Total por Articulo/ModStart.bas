Attribute VB_Name = "ModStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public UsuLogueado As Long
Public miConexion As New clsConexion

Public prmPlStockEstado As String
Public prmPlUbicoStockCamion As Long, prmPlUltimaCpa As Long, prmPlAjusteStock As Long

Public Sub Main()
On Error GoTo ErrMain
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miConexion.AccesoAlMenu(App.Title) Then
            InicioConexionBD miConexion.TextoConexion("comercio")
            UsuLogueado = miConexion.UsuarioLogueado(True)
            CargoParametrosLocales
            frmListado.Show vbModeless
            If Trim(Command()) <> "" Then frmListado.SetArticuloParmetro Val(Command())
            
        Else
            If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contraseña.
        miConexion.AccesoAlMenu (App.Title)
        InicioConexionBD miConexion.TextoConexion("comercio")
    End If
    Exit Sub
    
ErrMain:
    clsGeneral.OcurrioError "Ocurrio un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub CargoParametrosLocales()

    On Error GoTo errParam
    
    
    Cons = "Select * from Parametro Where ParNombre in('plStockTotalEstado', 'plantillasartstock', 'PlUbicoStockCamion', 'plUltimaCompra', 'plAjusteStock')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
        '    Case "plantillasartstock": If Not IsNull(RsAux!ParTexto) Then prmPlantillasArtStock = Trim(RsAux!ParTexto)
            Case "plstocktotalestado": If Not IsNull(RsAux!ParTexto) Then prmPlStockEstado = Trim(RsAux!ParTexto)
            Case LCase("PlUbicoStockCamion"): If Not IsNull(RsAux!ParValor) Then prmPlUbicoStockCamion = RsAux!ParValor
            Case LCase("plUltimaCompra"): If Not IsNull(RsAux!ParValor) Then prmPlUltimaCpa = RsAux!ParValor
            Case LCase("plAjusteStock"): If Not IsNull(RsAux!ParValor) Then prmPlAjusteStock = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Sub
    
errParam:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub
