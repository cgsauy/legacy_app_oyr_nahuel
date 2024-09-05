Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA  'clsLibGeneral
Public txtConexion As String
Public pathAppGeneral As String

Public Sub Main()
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion(logImportaciones)
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        InicioConexionBD txtConexion
        CargoParametrosComercio
        CargoParametrosImportaciones

        'Saco  el path a C:\.....\Aplicaciones      (general)
        Dim MyPos, MyPosA As Integer: MyPos = 1: MyPosA = 1
        Do While InStr(MyPos + 1, App.Path, "\") <> 0
            MyPosA = MyPos
            MyPos = InStr(MyPos + 1, App.Path, "\") + 1
        Loop
        pathAppGeneral = Mid(App.Path, 1, MyPosA - 1) & "Aplicaciones"
        
        frmInGastos.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmInGastos.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmInGastos.Status.Panels("bd") = "BD: " & PropiedadesConnect(miConexion.TextoConexion(logImportaciones), Database:=True) & " "
                
        frmInGastos.Show
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub

errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & _
                Err.Number & " - " & Err.Description
    End
End Sub
