Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmTipoMCompraME As Long

Public Sub Main()

Dim aSucursal As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion("comercio")
        
        aSucursal = CargoParametrosSucursal
        CargoParametrosLocales
        
        frmMovimientos.Status.Panels("sucursal") = "Sucursal: " & aSucursal
        frmMovimientos.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmMovimientos.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmMovimientos.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        frmMovimientos.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
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
    
    cons = "Select * from Parametro Where ParNombre In('monedapesos', 'MCCompraMonedaE')"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsAux.EOF
        Select Case Trim(LCase(rsAux!ParNombre))
            Case "monedapesos": paMonedaPesos = rsAux!ParValor
            
            Case "mccompramonedae": prmTipoMCompraME = rsAux!ParValor
            
        End Select
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
End Sub
