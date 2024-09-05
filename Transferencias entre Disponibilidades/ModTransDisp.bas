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
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurri� un error al inicializar la aplicaci�n " & App.Title
    End
End Sub


Private Sub CargoParametro()
    Cons = "Select * From Parametro Where ParNombre = 'monedapesos'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paMonedaPesos = RsAux!ParValor
    Else
        MsgBox "No se cargo el par�metro moneda pesos.", vbExclamation, "ATENCI�N"
    End If
    RsAux.Close
    
End Sub
