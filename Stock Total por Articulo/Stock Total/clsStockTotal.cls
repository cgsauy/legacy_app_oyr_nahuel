VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStockTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ShowStockTotal(Optional lArticulo As Long = 0)

    Screen.MousePointer = 11
    If Not miConexion.AccesoAlMenu(App.Title) Then
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
        Exit Sub
    End If

    If cBase Is Nothing Then
        'No se inicio la app hasta ahora.
        InicioConexionBD miConexion.TextoConexion("comercio")
        CargoDatosSucursal miConexion.NombreTerminal
        CargoParametrosLocales
    End If
    frmListado.Show
    If lArticulo > 0 Then frmListado.SetArticuloParmetro lArticulo
    If frmListado.WindowState = vbMinimized Then frmListado.WindowState = vbNormal
    frmListado.SetFocus
    frmListado.ZOrder 0
    Screen.MousePointer = 0

End Sub
