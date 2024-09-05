Attribute VB_Name = "modStart"
Option Explicit

Public Enum HorariosFijos
    No = 0
    Ingreso = 1
    AlmuerzoS = 2
    TransitoriaS = 3
    AlmuerzoR = 5
    TransitoriaR = 6
    Salida = 9
    LicenciaS = 21
    LicenciaR = 29
    FaltaConAviso = 31
    SaleAntesHora = 32
End Enum


Public usrCodigo As Integer
Public cBase As rdoConnection
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
    If miConexion.AccesoAlMenu(App.Title) Then
        
        sErr = "Inicio conexi�n"
        Dim oFnc As New clsFunciones
        If Not oFnc.GetBDConnect(cBase, "login") Then
            Set oFnc = Nothing
            Screen.MousePointer = 0
            End
        End If
        Set oFnc = Nothing
        
        usrCodigo = miConexion.UsuarioLogueado(True)
        
        sErr = "Show form"
        frmModificarHorario.Show
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

