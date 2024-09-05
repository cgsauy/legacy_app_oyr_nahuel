Attribute VB_Name = "modStart"
Option Explicit

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
    
    'Permisos para la aplicación para el usuario logueado. (Referencia a componente aaconexion)
    If miConexion.AccesoAlMenu(App.Title) Then
        
        sErr = "Inicio conexión"
        Dim oFnc As New clsFunciones
        If Not oFnc.GetBDConnect(cBase, "login") Then
            Set oFnc = Nothing
            Screen.MousePointer = 0
            End
        End If
        Set oFnc = Nothing
        
        
        sErr = "Show form"
        frmABM.Show
        sErr = "end"
        Screen.MousePointer = 0
    Else
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Set miConexion = Nothing
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description) & vbCr & sErr, vbCritical, "ATENCIÓN"
    End
End Sub

