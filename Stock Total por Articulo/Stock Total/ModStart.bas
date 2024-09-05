Attribute VB_Name = "ModStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion

Public prmPlStockEstado As String  ', prmPlantillasArtStock As String
Public prmPlUbicoStockCamion As Long, prmPlUltimaCpa As Long, prmPlAjusteStock As Long

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
    clsGeneral.OcurrioError "Ocurrió un error al activar la aplicación stock total.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Public Sub CargoParametrosLocales()

    On Error GoTo errParam
    
    '            Case "plantillasartstock": If Not IsNull(RsAux!ParTexto) Then prmPlantillasArtStock = Trim(RsAux!ParTexto)
    
    Cons = "Select * from Parametro Where ParNombre in('plStockTotalEstado', 'PlUbicoStockCamion', 'plUltimaCompra', 'plAjusteStock')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "plstocktotalestado": If Not IsNull(RsAux!ParTexto) Then prmPlStockEstado = Trim(RsAux!ParTexto)
            Case LCase("PlUbicoStockCamion"): If Not IsNull(RsAux!ParValor) Then prmPlUbicoStockCamion = RsAux!ParValor
            Case LCase("plUltimaCompra"): If Not IsNull(RsAux!ParValor) Then prmPlUltimaCpa = RsAux!ParValor
            Case LCase("plAjusteStock"): If Not IsNull(RsAux!ParValor) Then prmPlAjusteStock = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Sub
    
errParam:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub

