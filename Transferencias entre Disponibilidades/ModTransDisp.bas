Attribute VB_Name = "ModTransDisp"
Option Explicit

Public clsGeneral As New clsLibGeneral
Public miConexion As New clsConexion
Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion("Comercio")
        CargoParametro
        frmTransferencia.Show vbModeless
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title
    End
End Sub


Private Sub CargoParametro()
    Cons = "Select * From Parametro Where ParNombre = 'monedapesos'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paMonedaPesos = RsAux!ParValor
    Else
        MsgBox "No se cargo el parámetro moneda pesos.", vbExclamation, "ATENCIÓN"
    End If
    RsAux.Close
    
End Sub
