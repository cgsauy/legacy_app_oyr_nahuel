Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public txtConexion As String
Public Sub Main()
Dim aSucursal As String
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        'V�lido que no haya otra instancia corriendo.
        txtConexion = miConexion.TextoConexion("comercio")
        If InicioConexionBD(txtConexion) Then
            frmReglas.Show
        End If
    Else
        End
    End If
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurri� un error al inicializar la aplicaci�n " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub
