Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public paLocalZF As Long, paLocalPuerto As Long

Private Const prmKeyApp = "Administrador de Caja"

Public prmPathApp As String
Public prmDispCierreCheques As Long

Public prmPlaPendienteCaja As String
Public prmPlaPendienteCajaCam As String

Public Sub Main()

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(prmKeyApp) Then
        If Not InicioConexionBD(miConexion.TextoConexion("comercio"), 30) Then End
        
        CargoParametrosComercio
        CargoParametrosLocales
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        frmCierre.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocales()
On Error GoTo errCP

    cons = "Select * from Parametro " & _
                " Where ParNombre IN ( 'pathapp', 'DisponibilidadCierreCheques', 'plantillaspendientecaja', 'plantillaspendientecajacam')"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto) & "\"
            
            Case "disponibilidadcierrecheques": prmDispCierreCheques = rsAux!ParValor
            
            Case "plantillaspendientecaja": prmPlaPendienteCaja = Trim(rsAux!ParTexto)
            Case "plantillaspendientecajacam": prmPlaPendienteCajaCam = Trim(rsAux!ParTexto)
            
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Exit Sub
errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub

