Attribute VB_Name = "modPrintDocumento"
'................................................................................
'Módulo para imprimir los documentos.
'
'Objetos Externos:  MsgError y VsView
'
'................................................................................
Option Explicit

Private Type tPaperSizePrint
    Numero As Integer
    Nombre As String
End Type
Private arrPrintPaperSize() As tPaperSizePrint
Private iCnfgPaper1 As Integer

Private Type DEVMODE
        dmDeviceName As String * 32
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * 32
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As DEVMODE) As Long

Public prmCopiesPrint As Integer

'           ATENCIÓN        ---------------------
'Manejo string debido a los formatos.
'-------------------------------------------------

Private arrCfgPrint() As String

Private Type tCnfg
    X As Long
    Y As Long
    W As Long
    Bold As Boolean
    Size As Long
    Align As String
    FontName As String
End Type

Private Type tCnfgDato
    Cnfg As tCnfg
    Dato As String
End Type

Private Type tFieldDocument
    ID As tCnfgDato                        'ID del artículo
    NombreArticulo As tCnfgDato
    Nombre As tCnfgDato
    NroSerie As tCnfgDato
    Tipo As tCnfgDato
    Estado As tCnfgDato
    Local As tCnfgDato
    Precio As tCnfgDato            'Precio vigente
    Variacion As tCnfgDato
    PrecioVenta As tCnfgDato
    Comentario As tCnfgDato
    NroSerieCB As tCnfgDato
End Type

Dim lFontSizeG As Long
Public regPrintCnfg As tFieldDocument
Private arrLabel() As tCnfgDato

Public prmPathApp As String

Public Sub CleanArray()
    Erase arrLabel
End Sub

Public Sub StarDocument()
    
    With regPrintCnfg
        .ID.Dato = ""
        .Comentario.Dato = ""
        .Estado.Dato = ""
        .Local.Dato = ""
        .Nombre.Dato = ""
        .NombreArticulo.Dato = ""
        .NroSerie.Dato = ""
        .NroSerieCB.Dato = ""
        .Precio.Dato = ""
        .PrecioVenta.Dato = ""
        .Tipo.Dato = ""
        .Variacion.Dato = ""
    End With
End Sub

Public Function PrintDocument(ByVal vsView As VSPrinter) As Boolean
Dim iCount As Integer, iPrint As Integer
Dim Y As Long
Dim sVia As String

On Error GoTo errPrint
    
    PrintDocument = False
    With vsView
        .StartDoc
        If .Error <> 0 Then
            .EndDoc
            MsgBox "Error al iniciar el documento de impresión." & vbCrLf & "Error: " & Err.Description, vbCritical, "ATENCIÓN"
            Exit Function
        Else
            .TableBorder = tbNone
            .PageBorder = pbNone
            .FileName = "Articulo especifico"
            .FontItalic = False
            
                        'Imprimo las etiquetas.
            Dim iQ As Integer
            For iQ = 1 To UBound(arrLabel)
                If arrLabel(iQ).Cnfg.X <> 0 Or arrLabel(iQ).Cnfg.Y <> 0 Then
                    vsView.TextAlign = taLeftTop
                    vsView.FontName = arrLabel(iQ).Cnfg.FontName
                    If arrLabel(iQ).Cnfg.Size <> 0 Then .FontSize = arrLabel(iQ).Cnfg.Size
                    If arrLabel(iQ).Cnfg.Bold Then vsView.FontBold = True
                    Select Case UCase(arrLabel(iQ).Cnfg.Align)
                        Case "C": vsView.TextAlign = taCenterTop
                        Case "R": vsView.TextAlign = taRightTop
                    End Select
                    If arrLabel(iQ).Cnfg.W > 0 Then
                        .TextBox arrLabel(iQ).Dato, arrLabel(iQ).Cnfg.X, arrLabel(iQ).Cnfg.Y, arrLabel(iQ).Cnfg.W, .TextHeight(arrLabel(iQ).Dato) + 40
                    Else
                        .CurrentX = arrLabel(iQ).Cnfg.X
                        .CurrentY = arrLabel(iQ).Cnfg.Y
                        .Paragraph = arrLabel(iQ).Dato
                    End If
                    vsView.FontBold = False
                    vsView.FontSize = lFontSizeG
                End If
            
            Next
            
            vsView.TextAlign = taLeftTop
            If regPrintCnfg.ID.Cnfg.X <> 0 Or regPrintCnfg.ID.Cnfg.Y <> 0 Then
                If regPrintCnfg.ID.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.ID.Cnfg.Size
                If regPrintCnfg.ID.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.ID.Cnfg.FontName
                Select Case UCase(regPrintCnfg.ID.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.ID.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.ID.Dato, regPrintCnfg.ID.Cnfg.X, regPrintCnfg.ID.Cnfg.Y, regPrintCnfg.ID.Cnfg.W, .TextHeight(regPrintCnfg.ID.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.ID.Cnfg.X
                    .CurrentY = regPrintCnfg.ID.Cnfg.Y
                    .Paragraph = regPrintCnfg.ID.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            'Estado
            If regPrintCnfg.Estado.Cnfg.X <> 0 Or regPrintCnfg.Estado.Cnfg.Y <> 0 Then
                If regPrintCnfg.Estado.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.Estado.Cnfg.Size
                If regPrintCnfg.Estado.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.Estado.Cnfg.FontName
                Select Case UCase(regPrintCnfg.Estado.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.Estado.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.Estado.Dato, regPrintCnfg.Estado.Cnfg.X, regPrintCnfg.Estado.Cnfg.Y, regPrintCnfg.Estado.Cnfg.W, .TextHeight(regPrintCnfg.Estado.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.Estado.Cnfg.X
                    .CurrentY = regPrintCnfg.Estado.Cnfg.Y
                    .Paragraph = regPrintCnfg.Estado.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            'Nombre
            If regPrintCnfg.Nombre.Cnfg.X <> 0 Or regPrintCnfg.Nombre.Cnfg.Y <> 0 Then
                If regPrintCnfg.Nombre.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.Nombre.Cnfg.Size
                If regPrintCnfg.Nombre.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.Nombre.Cnfg.FontName
                
                Select Case UCase(regPrintCnfg.Nombre.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.Nombre.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.Nombre.Dato, regPrintCnfg.Nombre.Cnfg.X, regPrintCnfg.Nombre.Cnfg.Y, regPrintCnfg.Nombre.Cnfg.W, .TextHeight(regPrintCnfg.Nombre.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.Nombre.Cnfg.X
                    .CurrentY = regPrintCnfg.Nombre.Cnfg.Y
                    .Paragraph = regPrintCnfg.Nombre.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            'NOm artículo
            If regPrintCnfg.NombreArticulo.Cnfg.X <> 0 Or regPrintCnfg.NombreArticulo.Cnfg.Y <> 0 Then
                If regPrintCnfg.NombreArticulo.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.NombreArticulo.Cnfg.Size
                If regPrintCnfg.NombreArticulo.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.NombreArticulo.Cnfg.FontName
                Select Case UCase(regPrintCnfg.NombreArticulo.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.NombreArticulo.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.NombreArticulo.Dato, regPrintCnfg.NombreArticulo.Cnfg.X, regPrintCnfg.NombreArticulo.Cnfg.Y, regPrintCnfg.NombreArticulo.Cnfg.W, .TextHeight(regPrintCnfg.NombreArticulo.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.NombreArticulo.Cnfg.X
                    .CurrentY = regPrintCnfg.NombreArticulo.Cnfg.Y
                    .Paragraph = regPrintCnfg.NombreArticulo.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            
            'Nro serie
            If regPrintCnfg.NroSerie.Cnfg.X <> 0 Or regPrintCnfg.NroSerie.Cnfg.Y <> 0 Then
                If regPrintCnfg.NroSerie.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.NroSerie.Cnfg.Size
                If regPrintCnfg.NroSerie.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.NroSerie.Cnfg.FontName
                Select Case UCase(regPrintCnfg.NroSerie.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.NroSerie.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.NroSerie.Dato, regPrintCnfg.NroSerie.Cnfg.X, regPrintCnfg.NroSerie.Cnfg.Y, regPrintCnfg.NroSerie.Cnfg.W, .TextHeight(regPrintCnfg.NroSerie.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.NroSerie.Cnfg.X
                    .CurrentY = regPrintCnfg.NroSerie.Cnfg.Y
                    .Paragraph = regPrintCnfg.NroSerie.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            If regPrintCnfg.NroSerieCB.Cnfg.X <> 0 Or regPrintCnfg.NroSerieCB.Cnfg.Y <> 0 Then
                If regPrintCnfg.NroSerieCB.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.NroSerieCB.Cnfg.Size
                If regPrintCnfg.NroSerieCB.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.NroSerieCB.Cnfg.FontName
                Select Case UCase(regPrintCnfg.NroSerieCB.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.NroSerieCB.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.NroSerieCB.Dato, regPrintCnfg.NroSerieCB.Cnfg.X, regPrintCnfg.NroSerieCB.Cnfg.Y, regPrintCnfg.NroSerieCB.Cnfg.W, .TextHeight(regPrintCnfg.NroSerieCB.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.NroSerieCB.Cnfg.X
                    .CurrentY = regPrintCnfg.NroSerieCB.Cnfg.Y
                    .Paragraph = regPrintCnfg.NroSerieCB.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            'Precio
            If regPrintCnfg.Precio.Cnfg.X <> 0 Or regPrintCnfg.Precio.Cnfg.Y <> 0 Then
                If regPrintCnfg.Precio.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.Precio.Cnfg.Size
                If regPrintCnfg.Precio.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.Precio.Cnfg.FontName
                Select Case UCase(regPrintCnfg.Precio.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.Precio.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.Precio.Dato, regPrintCnfg.Precio.Cnfg.X, regPrintCnfg.Precio.Cnfg.Y, regPrintCnfg.Precio.Cnfg.W, .TextHeight(regPrintCnfg.Precio.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.Precio.Cnfg.X
                    .CurrentY = regPrintCnfg.Precio.Cnfg.Y
                    .Paragraph = regPrintCnfg.Precio.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            'Precio vta
            If regPrintCnfg.PrecioVenta.Cnfg.X <> 0 Or regPrintCnfg.PrecioVenta.Cnfg.Y <> 0 Then
                If regPrintCnfg.PrecioVenta.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.PrecioVenta.Cnfg.Size
                If regPrintCnfg.PrecioVenta.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.PrecioVenta.Cnfg.FontName
                Select Case UCase(regPrintCnfg.PrecioVenta.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.PrecioVenta.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.PrecioVenta.Dato, regPrintCnfg.PrecioVenta.Cnfg.X, regPrintCnfg.PrecioVenta.Cnfg.Y, regPrintCnfg.PrecioVenta.Cnfg.W, .TextHeight(regPrintCnfg.PrecioVenta.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.PrecioVenta.Cnfg.X
                    .CurrentY = regPrintCnfg.PrecioVenta.Cnfg.Y
                    .Paragraph = regPrintCnfg.PrecioVenta.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            'Tipo
            If regPrintCnfg.Tipo.Cnfg.X <> 0 Or regPrintCnfg.Tipo.Cnfg.Y <> 0 Then
                If regPrintCnfg.Tipo.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.Tipo.Cnfg.Size
                If regPrintCnfg.Tipo.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.Tipo.Cnfg.FontName
                Select Case UCase(regPrintCnfg.Tipo.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.Tipo.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.Tipo.Dato, regPrintCnfg.Tipo.Cnfg.X, regPrintCnfg.Tipo.Cnfg.Y, regPrintCnfg.Tipo.Cnfg.W, .TextHeight(regPrintCnfg.Tipo.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.Tipo.Cnfg.X
                    .CurrentY = regPrintCnfg.Tipo.Cnfg.Y
                    .Paragraph = regPrintCnfg.Tipo.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            'Variacion
            If regPrintCnfg.Variacion.Cnfg.X <> 0 Or regPrintCnfg.Variacion.Cnfg.Y <> 0 Then
                If regPrintCnfg.Variacion.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.Variacion.Cnfg.Size
                If regPrintCnfg.Variacion.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.Variacion.Cnfg.FontName
                Select Case UCase(regPrintCnfg.Variacion.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.Variacion.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.Variacion.Dato, regPrintCnfg.Variacion.Cnfg.X, regPrintCnfg.Variacion.Cnfg.Y, regPrintCnfg.Variacion.Cnfg.W, .TextHeight(regPrintCnfg.Variacion.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.Variacion.Cnfg.X
                    .CurrentY = regPrintCnfg.Variacion.Cnfg.Y
                    .Paragraph = regPrintCnfg.Variacion.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            'Local
            If regPrintCnfg.Local.Cnfg.X <> 0 Or regPrintCnfg.Local.Cnfg.Y <> 0 Then
                If regPrintCnfg.Local.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.Local.Cnfg.Size
                If regPrintCnfg.Local.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.Local.Cnfg.FontName
                Select Case UCase(regPrintCnfg.Local.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                
                If regPrintCnfg.Local.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.Local.Dato, regPrintCnfg.Local.Cnfg.X, regPrintCnfg.Local.Cnfg.Y, regPrintCnfg.Local.Cnfg.W, .TextHeight(regPrintCnfg.Local.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.Local.Cnfg.X
                    .CurrentY = regPrintCnfg.Local.Cnfg.Y
                    .Paragraph = regPrintCnfg.Local.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            
            'COMENTARIO
            If regPrintCnfg.Comentario.Cnfg.X <> 0 Or regPrintCnfg.Comentario.Cnfg.Y <> 0 Then
                vsView.TextAlign = taLeftTop
                If regPrintCnfg.Comentario.Cnfg.Size <> 0 Then .FontSize = regPrintCnfg.Comentario.Cnfg.Size
                If regPrintCnfg.Comentario.Cnfg.Bold Then vsView.FontBold = True
                vsView.FontName = regPrintCnfg.Comentario.Cnfg.FontName
                Select Case UCase(regPrintCnfg.Comentario.Cnfg.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.Comentario.Cnfg.W > 0 Then
                    .TextBox regPrintCnfg.Comentario.Dato, regPrintCnfg.Comentario.Cnfg.X, regPrintCnfg.Comentario.Cnfg.Y, regPrintCnfg.Comentario.Cnfg.W, .TextHeight(regPrintCnfg.Comentario.Dato) + 40
                Else
                    .CurrentX = regPrintCnfg.Comentario.Cnfg.X
                    .CurrentY = regPrintCnfg.Comentario.Cnfg.Y
                    .Paragraph = regPrintCnfg.Comentario.Dato
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            
            .EndDoc
        End If
        .PrintDoc False
        
    End With
        
    'Si llegue aqui sin errores marco como impreso.
    PrintDocument = True
        
    Exit Function
errPrint:
    MsgBox "Error al imprimir." & vbCr & Err.Number & "-" & Err.Description, vbCritical, "Imprimir ficha"
End Function

Public Function InitDevicePrinter(ByVal vsView As Object) As String
Dim iFile As Integer, iCont As Integer
Dim sAux As String
Dim sPaper1 As String, sPaper2 As String, sPaper3 As String
Dim iPaperSize As Integer

    Screen.MousePointer = 11
    prmCopiesPrint = 1
    
    With vsView
        .PrintQuality = -1
        .AbortWindow = False
    End With
    iCnfgPaper1 = 0     'Nro papersize hoja 1
    
        
    On Error GoTo errOpen
    iCont = 0
    Erase arrCfgPrint
    ReDim arrCfgPrint(iCont)
    iFile = FreeFile
    Open prmPathApp & "rptartespecifico.txt" For Input As iFile
    Do While Not EOF(iFile)
        Line Input #iFile, sAux
        'Proceso la línea
        ReDim Preserve arrCfgPrint(iCont)
        arrCfgPrint(iCont) = sAux
        iCont = iCont + 1
    Loop
    Close iFile
    
    On Error GoTo errOFC
    Dim sPaso As String
    For iCont = 0 To UBound(arrCfgPrint)
        If arrCfgPrint(iCont) <> "" Then
            If InStr(1, arrCfgPrint(iCont), "]") > 0 Then
                sAux = Mid(arrCfgPrint(iCont), 1, InStr(1, arrCfgPrint(iCont), "]"))
                Select Case LCase(sAux)
                    Case "[device]"
                        sPaso = "Dev: " & Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                        vsView.Device = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                        arrCfgPrint(iCont) = ""
                    Case "[mleft]"
                        vsView.MarginLeft = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                        arrCfgPrint(iCont) = ""
                    Case "[mright]"
                        vsView.MarginRight = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                        arrCfgPrint(iCont) = ""
                    Case "[mbottom]"
                        vsView.MarginBottom = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                        arrCfgPrint(iCont) = ""
                    Case "[mtop]"
                        vsView.MarginTop = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                        arrCfgPrint(iCont) = ""
                    Case "[font]"
                        vsView.Font = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                        If LCase(vsView.Font) = "arial" Then vsView.Font = "MS Sans Serif"
                    Case "[fontsize]"
                        vsView.FontSize = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
'                    Case "[papersize]"
'                        If IsNumeric(Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)) Then
'                            iPaperSize = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
'                        End If
                        
                    Case "[printquality]"
                        vsView.PrintQuality = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                        arrCfgPrint(iCont) = ""
                        
                    Case "[papersizename1]"
                        sPaper1 = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                    
                End Select
            End If
        End If
    Next
    lFontSizeG = vsView.FontSize
    loc_GetPosXY
        
    
    iPaperSize = paPrintConfPaperSize
    sPaso = "PaperSize: " & vsView.PaperSize & " seteo " & paPrintConfPaperSize
    vsView.PaperSize = paPrintConfPaperSize
    sPaso = "Device: " & vsView.Device & " seteo " & paPrintConfD
    vsView.Device = paPrintConfD
    sPaso = "PaperBin: " & vsView.PaperBin & " seteo " & paPrintConfB
    vsView.PaperBin = paPrintConfB
    
    'Cargo los <> papeles que tiene la impresora
    ReDim arrPrintPaperSize(0)
    iFile = 1
    For iCont = 1 To 256
        If vsView.PaperSizes(iCont) Then
            ReDim Preserve arrPrintPaperSize(iFile)
            arrPrintPaperSize(iFile).Numero = iCont
            iFile = iFile + 1
        End If
    Next
    
    loc_GetPaperSizeName vsView.Device
    If sPaper1 = "" And sPaper2 = "" And sPaper3 = "" Then
        iCnfgPaper1 = iPaperSize
    Else
        'Ahora recorro el array en busca del nombre del papel.
        For iCont = 1 To UBound(arrPrintPaperSize)
            With arrPrintPaperSize(iCont)
                If LCase(sPaper1) = LCase(.Nombre) Then iCnfgPaper1 = .Numero
            End With
        Next iCont
    End If
    Erase arrPrintPaperSize
    Screen.MousePointer = 0
    Exit Function
errOpen:
    Screen.MousePointer = 0
    MsgBox "Error al leer el archivo de configuración." & vbCr & sPaso & vbCr & Err.Number & "-" & Err.Description & " Archivo: " & prmPathApp & "rptartespecifico.txt", vbCritical, "Iniciar impresora"
    Exit Function
errOFC:
    Screen.MousePointer = 0
    MsgBox "Error al setear la configuración de la impresora." & vbCr & sPaso & vbCr & Err.Number & "-" & Err.Description, vbCritical, "Iniciar impresora"
End Function

Private Sub loc_GetPosXY()
Dim iCont As Integer
Dim sAux As String, sAlign As String, sfuente As String
Dim X As Long, Y As Long, W As Long, lFS As Long
Dim bBold As Boolean


    ReDim arrLabel(0)
    
    For iCont = 0 To UBound(arrCfgPrint)
        If arrCfgPrint(iCont) <> "" Then
            If InStr(1, arrCfgPrint(iCont), "]") > 0 Then
                sAux = loc_GetKeyCnfg(arrCfgPrint(iCont), X, Y, W, lFS, bBold, sAlign, sfuente)
                Select Case LCase(sAux)
                    Case LCase("[ID]")
                        With regPrintCnfg.ID.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With
                    
                    Case LCase("[Comentario]")
                        With regPrintCnfg.Comentario.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With

                    Case LCase("[Estado]")
                        With regPrintCnfg.Estado.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With
                    
                    Case LCase("[Local]")
                        With regPrintCnfg.Local.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With
                    
                    Case LCase("[Nombre]")
                        With regPrintCnfg.Nombre.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With
                    
                    Case LCase("[NombreArticulo]"):
                        With regPrintCnfg.NombreArticulo.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With
                        
                    Case LCase("[NroSerie]")
                        With regPrintCnfg.NroSerie.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With
                        
                    Case LCase("[NroSerieCB]")
                        With regPrintCnfg.NroSerieCB.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With
                    Case LCase("[Precio]")
                        With regPrintCnfg.Precio.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With
                        
                    Case LCase("[PrecioFinal]")
                        With regPrintCnfg.PrecioVenta.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With
                        
                    Case LCase("[Tipo]")
                        With regPrintCnfg.Tipo.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With
                    
                    Case LCase("[Variacion]")
                        With regPrintCnfg.Variacion.Cnfg
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                            .FontName = sfuente
                        End With
                        
                    Case Else
                        If InStr(1, sAux, "[label.", vbTextCompare) = 1 Then
                            sAux = Mid(sAux, 8, InStr(1, sAux, "]") - 8)
                            ReDim Preserve arrLabel(UBound(arrLabel) + 1)
                            With arrLabel(UBound(arrLabel)).Cnfg
                                .X = X: .Y = Y: .W = W
                                .Bold = bBold
                                If lFS <> 0 Then .Size = lFS
                                If sAlign <> "" Then .Align = sAlign
                                .FontName = sfuente
                            End With
                            arrLabel(UBound(arrLabel)).Dato = sAux
                        End If
                End Select
            End If
        End If
    Next
    
    
End Sub

Private Function loc_GetKeyCnfg(ByVal sLine As String, X As Long, Y As Long, W As Long, _
                    lFS As Long, bBold As Boolean, sAlign As String, fuente As String) As String
Dim vDatos() As String
    loc_GetKeyCnfg = ""
    X = 0
    Y = 0
    W = 0
    lFS = 0
    bBold = False
    sAlign = ""
    fuente = "tahoma"
    
    If sLine <> "" Then
        If InStr(1, sLine, "]") > 0 Then
            loc_GetKeyCnfg = Mid(sLine, 1, InStr(1, sLine, "]"))
            sLine = Mid(sLine, InStr(1, sLine, "]") + 1)
            If InStr(1, sLine, ":") > 0 Then
                vDatos = Split(sLine, ":")
                If UBound(vDatos) >= 0 Then
                    If IsNumeric(vDatos(0)) Then X = vDatos(0)
                
                    If UBound(vDatos) >= 1 Then
                        If IsNumeric(vDatos(1)) Then Y = vDatos(1)
                    End If
                    
                    If UBound(vDatos) >= 2 Then
                        If IsNumeric(vDatos(2)) Then W = vDatos(2)
                    End If
                    
                    If UBound(vDatos) >= 3 Then
                        If UCase(vDatos(3)) = "B" Then bBold = True
                    End If
                    
                    If UBound(vDatos) >= 4 Then
                        If IsNumeric(vDatos(4)) Then lFS = vDatos(4)
                    End If
                    
                    If UBound(vDatos) >= 5 Then sAlign = UCase(vDatos(5))
                        
                    If UBound(vDatos) >= 6 Then
                        fuente = vDatos(6)
                    End If
                End If
            End If
        End If
    End If
End Function

Private Sub loc_GetPaperSizeName(ByVal sDeviceName As String)
Dim sPaperAux As String, sNextP As String
Dim iCount As Integer, iCountP As Integer
Dim objDM As DEVMODE

    iCountP = DeviceCapabilities(sDeviceName, "", 16, ByVal vbNullString, objDM)
    sPaperAux = String(64 * iCountP, 0)

    ' Get paper names supported by active printer.
    iCountP = DeviceCapabilities(sDeviceName, "", 16, ByVal sPaperAux, objDM)
    
    For iCount = 1 To iCountP
        sNextP = Mid(sPaperAux, 64 * (iCount - 1) + 1, 64)
        sNextP = Trim(Left(sNextP, InStr(1, sNextP, Chr(0)) - 1))
        If iCount > UBound(arrPrintPaperSize) Then ReDim Preserve arrPrintPaperSize(iCount)
        arrPrintPaperSize(iCount).Nombre = sNextP
    Next
    
End Sub

