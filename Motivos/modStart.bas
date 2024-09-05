Attribute VB_Name = "modStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion

Public Sub Main()
Dim sPrm As String
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        If InicioConexionBD(miConexion.TextoConexion("comercio")) Then
            sPrm = Command()
            If sPrm <> "" And IsNumeric(sPrm) Then frmMotivo.prmIDTipoNew = Val(sPrm)
            frmMotivo.Show
        Else
            End
        End If
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
