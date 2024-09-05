Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public paempNoTrabajaMas As Integer
Public paempSeguroParo As Integer

Public Sub Main()

Dim aSucursal As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        
        InicioConexionBD miConexion.TextoConexion("comercio")
        CargoParametrosLocales
        frmFuncEmpresa.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
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

Public Sub CargoParametrosLocales()

    On Error GoTo errParametro
 
    cons = "Select * from Parametro Where ParNombre like 'emp%'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            
            Case "empnotrabajamas": paempNoTrabajaMas = rsAux!ParValor
            Case "empseguroparo": paempSeguroParo = rsAux!ParValor
            
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
errParametro:
End Sub
