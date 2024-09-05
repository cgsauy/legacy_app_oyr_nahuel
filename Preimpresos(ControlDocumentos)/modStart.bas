Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public mSQL As String

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    Dim bAccesoOK As Boolean
    bAccesoOK = False
    
    If Not bAccesoOK Then bAccesoOK = miConexion.AccesoAlMenu("Control")
    
    If bAccesoOK Then
    
        If Not InicioConexionBD(miConexion.TextoConexion("comercio"), 30) Then End
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        frmControl.Show vbModeless
        
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

Public Function EndMain()
    On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Function

