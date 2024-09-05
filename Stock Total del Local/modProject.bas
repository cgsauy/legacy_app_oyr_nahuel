Attribute VB_Name = "modProject"
Option Explicit
Public paCodigoDeSucursal As Integer
Public rdBase As rdoConnection

Public Sub Main()
On Error GoTo errM
    Screen.MousePointer = 11
    
    Dim oPerm As New clsConexion
    If Not oPerm.AccesoAlMenu("StockTotalDelLocal") Then
        Set oPerm = Nothing
        MsgBox "Sin acceso a la aplicación.", vbExclamation, "Acceso"
        End
        Exit Sub
    End If
    Set oPerm = Nothing
    
    Dim oFnc As New clsFunciones
    If Not oFnc.GetBDConnect(rdBase, 15) Then
        Set oFnc = Nothing
        End
    End If
    ps_GetSucursal
    frmStockLocal.Show vbModal
    End
    Exit Sub
errM:
    Screen.MousePointer = 0
    MsgBox "Error al iniciar la aplicación, error: " & Err.Description, vbCritical, "Stock en local"
End Sub


Private Sub ps_GetSucursal()
Dim Cons As String, rsAux As rdoResultset
    Dim oUsers As New clsConexion
    Cons = "Select TerSucursal From Terminal, Local" _
            & " Where TerNombre = '" & oUsers.NombreTerminal & "'" _
            & " And TerSucursal = LocCodigo"
    Set rsAux = rdBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then paCodigoDeSucursal = rsAux!TerSucursal
    rsAux.Close
    Set oUsers = Nothing
End Sub

