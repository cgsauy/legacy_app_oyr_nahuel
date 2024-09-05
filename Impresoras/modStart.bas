Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral

Public paLocalColonia As Integer
Public idSucursal As Integer

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    paLocalColonia = 5
    
    If Not miConexion.AccesoAlMenu("Cambiar Impresoras") Then End
    
    If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then End
    
    CargoParametrosSucursal
    If idSucursal <> paLocalColonia Then
        MsgBox "La sucursal es distinta al local COLONIA." & vbCrLf & "Este programa no se puede ejecutar.", vbInformation, "Sucursal <> a Colonia"
        CierroConexion
        Set cBase = Nothing
        Set eBase = Nothing
        
        Set clsGeneral = Nothing
        Set miConexion = Nothing
        End
        Exit Sub
    End If
    frmImpresoras.Show vbModeless
    
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub


Private Sub CargoParametrosSucursal()

Dim aNombreTerminal As String

    aNombreTerminal = Trim(miConexion.NombreTerminal)
    idSucursal = 0
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & aNombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then idSucursal = rsAux!TerSucursal
    rsAux.Close
    
End Sub

