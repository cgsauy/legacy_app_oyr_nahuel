Attribute VB_Name = "ModStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public UsuLogueado As Long
Public miConexion As New clsConexion

Public prmPlantillasArtStock As String

Public Sub Main()
On Error GoTo ErrMain
Dim sParam As String

    If App.StartMode = vbSModeStandalone Then
        Dim objStock As New clsStockTotal
        sParam = Trim(Command())
        If IsNumeric(sParam) Then
            objStock.ShowStockTotal Val(sParam)
        Else
            objStock.ShowStockTotal
        End If
        Set objStock = Nothing
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Ocurrio un error al activar la aplicación stock total.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Public Sub CargoParametrosLocales()

    On Error GoTo errParam
    prmPlantillasArtStock = 0
    
    Cons = "Select * from Parametro Where ParNombre like 'Plantillas%'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "plantillasartstock": If Not IsNull(RsAux!ParTexto) Then prmPlantillasArtStock = Trim(RsAux!ParTexto)
        End Select
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Sub
    
errParam:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub
