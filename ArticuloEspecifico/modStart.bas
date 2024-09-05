Attribute VB_Name = "modStart"
Option Explicit

'Impresora
Public paPrintConfB As Integer
Public paPrintConfD As String
Public paPrintConfXDef As Boolean
Public paPrintConfPaperSize As Integer


Public clsGeneral As New clsorCGSA

Public Sub Main()
Dim miConexion As clsConexion
Dim sErr As String
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    sErr = "Inicio objUser"
    Set miConexion = New clsConexion
    sErr = "Permisos"
    
    'Permisos para la aplicaci�n para el usuario logueado. (Referencia a componente aaconexion)
    If miConexion.AccesoAlMenu("ArticuloEspecifico") Then
        
        sErr = "Inicio conexi�n"
        'Si da error la conexi�n la misma despliega el msg de error
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            Screen.MousePointer = 0
            End: Exit Sub
        End If
        
        CargoDatosSucursal miConexion.NombreTerminal
        prj_GetPrinter False
                
        'Guardo el usuario logueado
        paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
        sErr = "Show form"
        frmABM.Show
        sErr = "end"
        Screen.MousePointer = 0
    Else
        MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
        Screen.MousePointer = 0
        End
    End If
    Set miConexion = Nothing
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicaci�n " & App.Title & Chr(13) & "Error: " & Trim(Err.Description) & vbCr & sErr, vbCritical, "ATENCI�N"
    End
End Sub

Public Sub prj_GetPrinter(ByVal bShowP As Boolean)
On Error GoTo errImp
    
    paPrintConfD = ""
    paPrintConfB = -1
    Dim objP As New clslPrintConfig
    With objP
        If bShowP Then
            If Not .ShowPrinterSetup("6", paCodigoDeTerminal) Then
                GoTo errImp
            End If
        End If
        If .LoadPrinterConfig("6", paCodigoDeTerminal) Then
            .GetPrinterDoc 6, paPrintConfD, paPrintConfB, paPrintConfXDef, paPrintConfPaperSize
        End If
    End With
    If paPrintConfD = "" Then MsgBox "Por favor verifique la configuraci�n de impresi�n.", vbInformation, "Atenci�n"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
End Sub


