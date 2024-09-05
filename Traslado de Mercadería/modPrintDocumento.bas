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
Private Type tRowDocument
    Codigo As String
    Cantidad As String
    Articulo As String
End Type

Private Type tFieldDocument
    DocumentName As String               'Contado, Crédito, etc.
    DocumentNumero As String            'A 002566
    DocumentFecha As String
    DocumentoCodigo As String
    DocumentoComentario As String
    DocumentoUsuario As String
    
    ClienteNombre As String
    ClienteDireccion As String
    
    arrRow() As tRowDocument
End Type

Private regDocument As tFieldDocument
Private arrCfgPrint() As String

Private Type tCnfg
    X As Long
    Y As Long
    W As Long
    Bold As Boolean
    Size As Long
    Align As String
End Type

Private Type tPrintCnfg
    Documento As tCnfg
    DocumentoNumero As tCnfg
        
    TableFechaP As tCnfg               'Pos donde empieza la tabla.
    TableFechaD As String             'Tomo la dimensión de c/col
    TableFechaR As tCnfg
        
    ClienteNombre As tCnfg
    ClienteDireccion As tCnfg
            
    RenglonArtCodigo As tCnfg
    RenglonCantidad As tCnfg
    RenglonArtNombre As tCnfg
            
    TablaRenglonFirst As tCnfg
'    TablaRenglon As String
    DocumentoComentario As tCnfg
End Type

Dim lFontSizeG As Long
Private regPrintCnfg As tPrintCnfg
Private regPrintCnfg2 As tPrintCnfg    'Para la 2da copia
Private regPrintCnfg3 As tPrintCnfg    'Para la 2da copia


Private oFieldPrint As clsFieldsPrint

Public prmPathApp As String

Public Sub StarDocument()
    
    With regDocument
        ReDim .arrRow(0)
        .ClienteDireccion = ""
        .ClienteNombre = ""
        .DocumentName = ""
        .DocumentNumero = ""
        .DocumentoCodigo = ""
        .DocumentFecha = ""
        .DocumentoUsuario = ""
    End With
    
End Sub

Public Sub SetDataDocument(ByVal sFecha As String, ByVal sNombre As String, ByVal sNumero As String, ByVal sUser As String, ByVal iCodigo As Long, ByVal sMemo As String)
    With regDocument
        .DocumentFecha = sFecha
        .DocumentName = sNombre
        .DocumentNumero = sNumero
        .DocumentoUsuario = sUser
        .DocumentoCodigo = iCodigo
        .DocumentoComentario = sMemo
    End With
End Sub

Public Sub SetClienteDocument(ByVal sNombre As String, ByVal sDireccion As String)
    With regDocument
        .ClienteDireccion = sDireccion
        .ClienteNombre = sNombre
    End With
End Sub

Public Sub SetNewArticuloDocument(ByVal sCodigo As String, ByVal sCant As String, ByVal sArticulo As String)
'Agrega un artículo a la lista.
    
    ReDim Preserve regDocument.arrRow(UBound(regDocument.arrRow) + 1)
    
    With regDocument.arrRow(UBound(regDocument.arrRow))
        .Articulo = sArticulo: .Cantidad = sCant: .Codigo = sCodigo
    End With
    
End Sub

Public Function PrintDocument(ByVal vsView As vsPrinter) As Boolean
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
            .FileName = regDocument.DocumentName & "_" & regDocument.DocumentNumero
            .FontItalic = False
            
            vsView.TextAlign = taLeftTop
            
            '1er campo a presentar Documento
            If regPrintCnfg.Documento.X <> 0 Or regPrintCnfg.Documento.Y <> 0 Then
                If regPrintCnfg.Documento.Size <> 0 Then .FontSize = regPrintCnfg.Documento.Size
                If regPrintCnfg.Documento.Bold Then vsView.FontBold = True
                Select Case UCase(regPrintCnfg.Documento.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.Documento.W > 0 Then
                    .TextBox regDocument.DocumentName & " " & regDocument.DocumentNumero, regPrintCnfg.Documento.X, regPrintCnfg.Documento.Y, regPrintCnfg.Documento.W, .TextHeight(regDocument.DocumentName & " " & regDocument.DocumentNumero) + 40
                Else
                    .CurrentX = regPrintCnfg.Documento.X
                    .CurrentY = regPrintCnfg.Documento.Y
                    .Paragraph = regDocument.DocumentName
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            If regPrintCnfg.DocumentoNumero.X <> 0 Or regPrintCnfg.DocumentoNumero.Y <> 0 Then
                If regPrintCnfg.DocumentoNumero.Size <> 0 Then .FontSize = regPrintCnfg.DocumentoNumero.Size
                If regPrintCnfg.DocumentoNumero.Bold Then vsView.FontBold = True
                Select Case UCase(regPrintCnfg.DocumentoNumero.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.DocumentoNumero.W > 0 Then
                    .TextBox regDocument.DocumentNumero, regPrintCnfg.DocumentoNumero.X, regPrintCnfg.DocumentoNumero.Y, regPrintCnfg.DocumentoNumero.W, .TextHeight(regDocument.DocumentNumero) + 40
                Else
                    .CurrentX = regPrintCnfg.DocumentoNumero.X
                    .CurrentY = regPrintCnfg.DocumentoNumero.Y
                    .Paragraph = regDocument.DocumentNumero
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            
            If regPrintCnfg.TableFechaP.X <> 0 Or regPrintCnfg.TableFechaP.Y <> 0 Then
                If regPrintCnfg.TableFechaP.Size <> 0 Then .FontSize = regPrintCnfg.TableFechaP.Size
                If regPrintCnfg.TableFechaP.Bold Then vsView.FontBold = True
                
                vsView.TextAlign = taRightTop
                .CurrentX = regPrintCnfg.TableFechaP.X
                .CurrentY = regPrintCnfg.TableFechaP.Y
                .AddTable regPrintCnfg.TableFechaD, "", "Fecha|Usuario|Código"
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
                
                vsView.TextAlign = taRightTop
                'Ahora inserto los datos.
                .CurrentX = regPrintCnfg.TableFechaR.X
                .CurrentY = regPrintCnfg.TableFechaR.Y
                .AddTable regPrintCnfg.TableFechaD, "", regDocument.DocumentFecha & "|" & regDocument.DocumentoUsuario & "|" & regDocument.DocumentoCodigo
            End If
            
            vsView.TextAlign = taLeftTop
            'NOMBRE CLIENTE
            If regPrintCnfg.ClienteNombre.X <> 0 Or regPrintCnfg.ClienteNombre.Y <> 0 Then
                If regPrintCnfg.ClienteNombre.Size <> 0 Then .FontSize = regPrintCnfg.ClienteNombre.Size
                If regPrintCnfg.ClienteNombre.Bold Then vsView.FontBold = True
                Select Case UCase(regPrintCnfg.ClienteNombre.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                
                If regPrintCnfg.ClienteNombre.W > 0 Then
                    .TextBox regDocument.ClienteNombre, regPrintCnfg.ClienteNombre.X, regPrintCnfg.ClienteNombre.Y, regPrintCnfg.ClienteNombre.W, .TextHeight(regDocument.ClienteNombre) + 40
                Else
                    .CurrentX = regPrintCnfg.ClienteNombre.X
                    .CurrentY = regPrintCnfg.ClienteNombre.Y
                    .Paragraph = regDocument.ClienteNombre
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            'DIRECCION CLIENTE
            If regPrintCnfg.ClienteDireccion.X <> 0 Or regPrintCnfg.ClienteDireccion.Y <> 0 Then
                If regPrintCnfg.ClienteDireccion.Size <> 0 Then .FontSize = regPrintCnfg.ClienteDireccion.Size
                If regPrintCnfg.ClienteDireccion.Bold Then vsView.FontBold = True
                Select Case UCase(regPrintCnfg.ClienteDireccion.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                
                If regPrintCnfg.ClienteDireccion.W > 0 Then
                    .TextBox regDocument.ClienteDireccion, regPrintCnfg.ClienteDireccion.X, regPrintCnfg.ClienteDireccion.Y, regPrintCnfg.ClienteDireccion.W, .TextHeight(regDocument.ClienteDireccion) + 40
                Else
                    .CurrentX = regPrintCnfg.ClienteDireccion.X
                    .CurrentY = regPrintCnfg.ClienteDireccion.Y
                    .Paragraph = regDocument.ClienteDireccion
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            
            vsView.TextAlign = taLeftTop
            .FontBold = False
            .FontSize = lFontSizeG
            
            'RENGLONES
            'Posiciono el primer renglon y me guardo la y siguiente.
            .CurrentY = regPrintCnfg.TablaRenglonFirst.Y
            
            For iCount = 1 To UBound(regDocument.arrRow)
                Y = .CurrentY
                
                .FontItalic = False
                
                .CurrentX = regPrintCnfg.RenglonArtCodigo.X
                If regPrintCnfg.RenglonArtCodigo.Size <> 0 Then .FontSize = regPrintCnfg.RenglonArtCodigo.Size
                If regPrintCnfg.RenglonArtCodigo.Bold Then .FontBold = True
                Select Case UCase(regPrintCnfg.RenglonArtCodigo.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.RenglonArtCodigo.W > 0 Then
                    .TextBox regDocument.arrRow(iCount).Codigo, regPrintCnfg.RenglonArtCodigo.X, Y, regPrintCnfg.RenglonArtCodigo.W, .TextHeight(regDocument.arrRow(iCount).Codigo) + 40
                Else
                    .CurrentX = regPrintCnfg.RenglonArtCodigo.X
                    .CurrentY = .Y
                    .Paragraph = regDocument.arrRow(iCount).Codigo
                End If
                .FontBold = False
                .FontSize = lFontSizeG
                .TextAlign = taLeftTop
                    
                'CANTIDAD
                .CurrentX = regPrintCnfg.RenglonArtNombre.X
                If regPrintCnfg.RenglonArtNombre.Size <> 0 Then .FontSize = regPrintCnfg.RenglonArtNombre.Size
                If regPrintCnfg.RenglonArtNombre.Bold Then .FontBold = True
                Select Case UCase(regPrintCnfg.RenglonArtNombre.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.RenglonArtNombre.W > 0 Then
                    .TextBox regDocument.arrRow(iCount).Articulo, regPrintCnfg.RenglonArtNombre.X, Y, regPrintCnfg.RenglonArtNombre.W, .TextHeight(regDocument.arrRow(iCount).Articulo) + 40
                Else
                    .CurrentX = regPrintCnfg.RenglonArtNombre.X
                    .CurrentY = .Y
                    .Paragraph = regDocument.arrRow(iCount).Articulo
                End If
                .FontBold = False
                .FontSize = lFontSizeG
                .TextAlign = taLeftTop
                    
                .CurrentX = regPrintCnfg.RenglonCantidad.X
                If regPrintCnfg.RenglonCantidad.Size <> 0 Then .FontSize = regPrintCnfg.RenglonCantidad.Size
                If regPrintCnfg.RenglonCantidad.Bold Then .FontBold = True
                Select Case UCase(regPrintCnfg.RenglonCantidad.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg.RenglonCantidad.W > 0 Then
                    .TextBox regDocument.arrRow(iCount).Cantidad, regPrintCnfg.RenglonCantidad.X, Y, regPrintCnfg.RenglonCantidad.W, .TextHeight(regDocument.arrRow(iCount).Cantidad) + 40
                Else
                    .CurrentX = regPrintCnfg.RenglonCantidad.X
                    .CurrentY = .Y
                    .Paragraph = regDocument.arrRow(iCount).Cantidad
                End If
                .FontBold = False
                .FontSize = lFontSizeG
                .TextAlign = taLeftTop
                    
                'Hago el renglon siguiente
                .CurrentY = Y: .Paragraph = ""

            Next iCount
            '..................................................................................................
            
            'Comentario del documento
            If regDocument.DocumentoComentario <> "" And (regPrintCnfg.DocumentoComentario.X <> 0 Or regPrintCnfg.DocumentoComentario.Y <> 0) Then
                If regPrintCnfg.DocumentoComentario.Size <> 0 Then .FontSize = regPrintCnfg.DocumentoComentario.Size
                If regPrintCnfg.DocumentoComentario.Bold Then vsView.FontBold = True
                Select Case UCase(regPrintCnfg.DocumentoComentario.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                
                If regPrintCnfg.DocumentoComentario.W > 0 Then
                    .TextBox regDocument.DocumentoComentario, regPrintCnfg.DocumentoComentario.X, regPrintCnfg.DocumentoComentario.Y, regPrintCnfg.DocumentoComentario.W, .TextHeight("DOCUMENTO EN PESOS") + 40
                Else
                    .CurrentX = regPrintCnfg.DocumentoComentario.X
                    .CurrentY = regPrintCnfg.DocumentoComentario.Y
                    .Paragraph = regDocument.DocumentoComentario
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            '..................................................................................................
            
            vsView.TextAlign = taLeftTop
'2da via
            '1er campo a presentar Documento
            If regPrintCnfg2.Documento.X <> 0 Or regPrintCnfg2.Documento.Y <> 0 Then
                If regPrintCnfg2.Documento.Size <> 0 Then .FontSize = regPrintCnfg2.Documento.Size
                If regPrintCnfg2.Documento.Bold Then vsView.FontBold = True
                Select Case UCase(regPrintCnfg2.Documento.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg2.Documento.W > 0 Then
                    .TextBox regDocument.DocumentName & " " & regDocument.DocumentNumero, regPrintCnfg.Documento.X, regPrintCnfg.Documento.Y, regPrintCnfg.Documento.W, .TextHeight(regDocument.DocumentName & " " & regDocument.DocumentNumero) + 40
                Else
                    .CurrentX = regPrintCnfg2.Documento.X
                    .CurrentY = regPrintCnfg2.Documento.Y
                    .Paragraph = regDocument.DocumentName
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            If regPrintCnfg2.DocumentoNumero.X <> 0 Or regPrintCnfg2.DocumentoNumero.Y <> 0 Then
                If regPrintCnfg2.DocumentoNumero.Size <> 0 Then .FontSize = regPrintCnfg2.DocumentoNumero.Size
                If regPrintCnfg2.DocumentoNumero.Bold Then vsView.FontBold = True
                Select Case UCase(regPrintCnfg2.DocumentoNumero.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg2.DocumentoNumero.W > 0 Then
                    .TextBox regDocument.DocumentNumero, regPrintCnfg.DocumentoNumero.X, regPrintCnfg.DocumentoNumero.Y, regPrintCnfg.DocumentoNumero.W, .TextHeight(regDocument.DocumentNumero) + 40
                Else
                    .CurrentX = regPrintCnfg2.DocumentoNumero.X
                    .CurrentY = regPrintCnfg2.DocumentoNumero.Y
                    .Paragraph = regDocument.DocumentNumero
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            
            If regPrintCnfg2.TableFechaP.X <> 0 Or regPrintCnfg2.TableFechaP.Y <> 0 Then
                If regPrintCnfg2.TableFechaP.Size <> 0 Then .FontSize = regPrintCnfg2.TableFechaP.Size
                If regPrintCnfg2.TableFechaP.Bold Then vsView.FontBold = True
                
                vsView.TextAlign = taRightTop
                .CurrentX = regPrintCnfg2.TableFechaP.X
                .CurrentY = regPrintCnfg2.TableFechaP.Y
                .AddTable regPrintCnfg2.TableFechaD, "", "Fecha|Usuario|Código"
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
                
                vsView.TextAlign = taRightTop
                'Ahora inserto los datos.
                .CurrentX = regPrintCnfg2.TableFechaR.X
                .CurrentY = regPrintCnfg2.TableFechaR.Y
                .AddTable regPrintCnfg2.TableFechaD, "", regDocument.DocumentFecha & "|" & regDocument.DocumentoUsuario & "|" & regDocument.DocumentoCodigo
            End If
            
            vsView.TextAlign = taLeftTop
            'NOMBRE CLIENTE
            If regPrintCnfg2.ClienteNombre.X <> 0 Or regPrintCnfg2.ClienteNombre.Y <> 0 Then
                If regPrintCnfg2.ClienteNombre.Size <> 0 Then .FontSize = regPrintCnfg2.ClienteNombre.Size
                If regPrintCnfg2.ClienteNombre.Bold Then vsView.FontBold = True
                Select Case UCase(regPrintCnfg2.ClienteNombre.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                
                If regPrintCnfg2.ClienteNombre.W > 0 Then
                    .TextBox regDocument.ClienteNombre, regPrintCnfg.ClienteNombre.X, regPrintCnfg.ClienteNombre.Y, regPrintCnfg.ClienteNombre.W, .TextHeight(regDocument.ClienteNombre) + 40
                Else
                    .CurrentX = regPrintCnfg2.ClienteNombre.X
                    .CurrentY = regPrintCnfg2.ClienteNombre.Y
                    .Paragraph = regDocument.ClienteNombre
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            vsView.TextAlign = taLeftTop
            
            'DIRECCION CLIENTE
            If regPrintCnfg2.ClienteDireccion.X <> 0 Or regPrintCnfg2.ClienteDireccion.Y <> 0 Then
                If regPrintCnfg2.ClienteDireccion.Size <> 0 Then .FontSize = regPrintCnfg2.ClienteDireccion.Size
                If regPrintCnfg2.ClienteDireccion.Bold Then vsView.FontBold = True
                Select Case UCase(regPrintCnfg2.ClienteDireccion.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                
                If regPrintCnfg2.ClienteDireccion.W > 0 Then
                    .TextBox regDocument.ClienteDireccion, regPrintCnfg.ClienteDireccion.X, regPrintCnfg.ClienteDireccion.Y, regPrintCnfg.ClienteDireccion.W, .TextHeight(regDocument.ClienteDireccion) + 40
                Else
                    .CurrentX = regPrintCnfg2.ClienteDireccion.X
                    .CurrentY = regPrintCnfg2.ClienteDireccion.Y
                    .Paragraph = regDocument.ClienteDireccion
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            
            vsView.TextAlign = taLeftTop
            .FontBold = False
            .FontSize = lFontSizeG
            
            'RENGLONES
            'Posiciono el primer renglon y me guardo la y siguiente.
            .CurrentY = regPrintCnfg2.TablaRenglonFirst.Y
            
            For iCount = 1 To UBound(regDocument.arrRow)
                Y = .CurrentY
                
                .FontItalic = False
                
                .CurrentX = regPrintCnfg2.RenglonArtCodigo.X
                If regPrintCnfg2.RenglonArtCodigo.Size <> 0 Then .FontSize = regPrintCnfg2.RenglonArtCodigo.Size
                If regPrintCnfg2.RenglonArtCodigo.Bold Then .FontBold = True
                Select Case UCase(regPrintCnfg2.RenglonArtCodigo.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg2.RenglonArtCodigo.W > 0 Then
                    .TextBox regDocument.arrRow(iCount).Codigo, regPrintCnfg.RenglonArtCodigo.X, Y, regPrintCnfg.RenglonArtCodigo.W, .TextHeight(regDocument.arrRow(iCount).Codigo) + 40
                Else
                    .CurrentX = regPrintCnfg2.RenglonArtCodigo.X
                    .CurrentY = .Y
                    .Paragraph = regDocument.arrRow(iCount).Codigo
                End If
                .FontBold = False
                .FontSize = lFontSizeG
                .TextAlign = taLeftTop
                    
                'CANTIDAD
                .CurrentX = regPrintCnfg2.RenglonArtNombre.X
                If regPrintCnfg2.RenglonArtNombre.Size <> 0 Then .FontSize = regPrintCnfg2.RenglonArtNombre.Size
                If regPrintCnfg2.RenglonArtNombre.Bold Then .FontBold = True
                Select Case UCase(regPrintCnfg2.RenglonArtNombre.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg2.RenglonArtNombre.W > 0 Then
                    .TextBox regDocument.arrRow(iCount).Articulo, regPrintCnfg.RenglonArtNombre.X, Y, regPrintCnfg.RenglonArtNombre.W, .TextHeight(regDocument.arrRow(iCount).Articulo) + 40
                Else
                    .CurrentX = regPrintCnfg2.RenglonArtNombre.X
                    .CurrentY = .Y
                    .Paragraph = regDocument.arrRow(iCount).Articulo
                End If
                .FontBold = False
                .FontSize = lFontSizeG
                .TextAlign = taLeftTop
                    
                .CurrentX = regPrintCnfg2.RenglonCantidad.X
                If regPrintCnfg2.RenglonCantidad.Size <> 0 Then .FontSize = regPrintCnfg2.RenglonCantidad.Size
                If regPrintCnfg2.RenglonCantidad.Bold Then .FontBold = True
                Select Case UCase(regPrintCnfg2.RenglonCantidad.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                If regPrintCnfg2.RenglonCantidad.W > 0 Then
                    .TextBox regDocument.arrRow(iCount).Cantidad, regPrintCnfg.RenglonCantidad.X, Y, regPrintCnfg.RenglonCantidad.W, .TextHeight(regDocument.arrRow(iCount).Cantidad) + 40
                Else
                    .CurrentX = regPrintCnfg2.RenglonCantidad.X
                    .CurrentY = .Y
                    .Paragraph = regDocument.arrRow(iCount).Cantidad
                End If
                .FontBold = False
                .FontSize = lFontSizeG
                .TextAlign = taLeftTop
                    
                'Hago el renglon siguiente
                .CurrentY = Y: .Paragraph = ""

            Next iCount
            '..................................................................................................
            
            'Comentario del documento
            If regDocument.DocumentoComentario <> "" And (regPrintCnfg2.DocumentoComentario.X <> 0 Or regPrintCnfg2.DocumentoComentario.Y <> 0) Then
                If regPrintCnfg2.DocumentoComentario.Size <> 0 Then .FontSize = regPrintCnfg2.DocumentoComentario.Size
                If regPrintCnfg2.DocumentoComentario.Bold Then vsView.FontBold = True
                Select Case UCase(regPrintCnfg2.DocumentoComentario.Align)
                    Case "C": vsView.TextAlign = taCenterTop
                    Case "R": vsView.TextAlign = taRightTop
                End Select
                
                If regPrintCnfg2.DocumentoComentario.W > 0 Then
                    .TextBox regDocument.DocumentoComentario, regPrintCnfg.DocumentoComentario.X, regPrintCnfg.DocumentoComentario.Y, regPrintCnfg.DocumentoComentario.W, .TextHeight("DOCUMENTO EN PESOS") + 40
                Else
                    .CurrentX = regPrintCnfg2.DocumentoComentario.X
                    .CurrentY = regPrintCnfg2.DocumentoComentario.Y
                    .Paragraph = regDocument.DocumentoComentario
                End If
                vsView.FontBold = False
                vsView.FontSize = lFontSizeG
            End If
            
            
            
            
            
'3era via
            If regPrintCnfg3.Documento.X <> 0 Or regPrintCnfg3.Documento.Y <> 0 Then

            '1er campo a presentar Documento
'                If regPrintCnfg3.Documento.X <> 0 Or regPrintCnfg3.Documento.Y <> 0 Then
                    If regPrintCnfg3.Documento.Size <> 0 Then .FontSize = regPrintCnfg3.Documento.Size
                    If regPrintCnfg3.Documento.Bold Then vsView.FontBold = True
                    Select Case UCase(regPrintCnfg3.Documento.Align)
                        Case "C": vsView.TextAlign = taCenterTop
                        Case "R": vsView.TextAlign = taRightTop
                    End Select
                    If regPrintCnfg3.Documento.W > 0 Then
                        .TextBox regDocument.DocumentName & " " & regDocument.DocumentNumero, regPrintCnfg.Documento.X, regPrintCnfg.Documento.Y, regPrintCnfg.Documento.W, .TextHeight(regDocument.DocumentName & " " & regDocument.DocumentNumero) + 40
                    Else
                        .CurrentX = regPrintCnfg3.Documento.X
                        .CurrentY = regPrintCnfg3.Documento.Y
                        .Paragraph = regDocument.DocumentName
                    End If
                    vsView.FontBold = False
                    vsView.FontSize = lFontSizeG
 '               End If
                vsView.TextAlign = taLeftTop
                
                If regPrintCnfg3.DocumentoNumero.X <> 0 Or regPrintCnfg3.DocumentoNumero.Y <> 0 Then
                    If regPrintCnfg3.DocumentoNumero.Size <> 0 Then .FontSize = regPrintCnfg3.DocumentoNumero.Size
                    If regPrintCnfg3.DocumentoNumero.Bold Then vsView.FontBold = True
                    Select Case UCase(regPrintCnfg3.DocumentoNumero.Align)
                        Case "C": vsView.TextAlign = taCenterTop
                        Case "R": vsView.TextAlign = taRightTop
                    End Select
                    If regPrintCnfg3.DocumentoNumero.W > 0 Then
                        .TextBox regDocument.DocumentNumero, regPrintCnfg.DocumentoNumero.X, regPrintCnfg.DocumentoNumero.Y, regPrintCnfg.DocumentoNumero.W, .TextHeight(regDocument.DocumentNumero) + 40
                    Else
                        .CurrentX = regPrintCnfg3.DocumentoNumero.X
                        .CurrentY = regPrintCnfg3.DocumentoNumero.Y
                        .Paragraph = regDocument.DocumentNumero
                    End If
                    vsView.FontBold = False
                    vsView.FontSize = lFontSizeG
                End If
                vsView.TextAlign = taLeftTop
                
                
                If regPrintCnfg3.TableFechaP.X <> 0 Or regPrintCnfg3.TableFechaP.Y <> 0 Then
                    If regPrintCnfg3.TableFechaP.Size <> 0 Then .FontSize = regPrintCnfg3.TableFechaP.Size
                    If regPrintCnfg3.TableFechaP.Bold Then vsView.FontBold = True
                    
                    vsView.TextAlign = taRightTop
                    .CurrentX = regPrintCnfg3.TableFechaP.X
                    .CurrentY = regPrintCnfg3.TableFechaP.Y
                    .AddTable regPrintCnfg3.TableFechaD, "", "Fecha|Usuario|Código"
                    vsView.FontBold = False
                    vsView.FontSize = lFontSizeG
                    
                    vsView.TextAlign = taRightTop
                    'Ahora inserto los datos.
                    .CurrentX = regPrintCnfg3.TableFechaR.X
                    .CurrentY = regPrintCnfg3.TableFechaR.Y
                    .AddTable regPrintCnfg3.TableFechaD, "", regDocument.DocumentFecha & "|" & regDocument.DocumentoUsuario & "|" & regDocument.DocumentoCodigo
                End If
                
                vsView.TextAlign = taLeftTop
                'NOMBRE CLIENTE
                If regPrintCnfg3.ClienteNombre.X <> 0 Or regPrintCnfg3.ClienteNombre.Y <> 0 Then
                    If regPrintCnfg3.ClienteNombre.Size <> 0 Then .FontSize = regPrintCnfg3.ClienteNombre.Size
                    If regPrintCnfg3.ClienteNombre.Bold Then vsView.FontBold = True
                    Select Case UCase(regPrintCnfg3.ClienteNombre.Align)
                        Case "C": vsView.TextAlign = taCenterTop
                        Case "R": vsView.TextAlign = taRightTop
                    End Select
                    
                    If regPrintCnfg3.ClienteNombre.W > 0 Then
                        .TextBox regDocument.ClienteNombre, regPrintCnfg.ClienteNombre.X, regPrintCnfg.ClienteNombre.Y, regPrintCnfg.ClienteNombre.W, .TextHeight(regDocument.ClienteNombre) + 40
                    Else
                        .CurrentX = regPrintCnfg3.ClienteNombre.X
                        .CurrentY = regPrintCnfg3.ClienteNombre.Y
                        .Paragraph = regDocument.ClienteNombre
                    End If
                    vsView.FontBold = False
                    vsView.FontSize = lFontSizeG
                End If
                vsView.TextAlign = taLeftTop
                
                'DIRECCION CLIENTE
                If regPrintCnfg3.ClienteDireccion.X <> 0 Or regPrintCnfg3.ClienteDireccion.Y <> 0 Then
                    If regPrintCnfg3.ClienteDireccion.Size <> 0 Then .FontSize = regPrintCnfg3.ClienteDireccion.Size
                    If regPrintCnfg3.ClienteDireccion.Bold Then vsView.FontBold = True
                    Select Case UCase(regPrintCnfg3.ClienteDireccion.Align)
                        Case "C": vsView.TextAlign = taCenterTop
                        Case "R": vsView.TextAlign = taRightTop
                    End Select
                    
                    If regPrintCnfg3.ClienteDireccion.W > 0 Then
                        .TextBox regDocument.ClienteDireccion, regPrintCnfg.ClienteDireccion.X, regPrintCnfg.ClienteDireccion.Y, regPrintCnfg.ClienteDireccion.W, .TextHeight(regDocument.ClienteDireccion) + 40
                    Else
                        .CurrentX = regPrintCnfg3.ClienteDireccion.X
                        .CurrentY = regPrintCnfg3.ClienteDireccion.Y
                        .Paragraph = regDocument.ClienteDireccion
                    End If
                    vsView.FontBold = False
                    vsView.FontSize = lFontSizeG
                End If
                
                vsView.TextAlign = taLeftTop
                .FontBold = False
                .FontSize = lFontSizeG
                
                'RENGLONES
                'Posiciono el primer renglon y me guardo la y siguiente.
                .CurrentY = regPrintCnfg3.TablaRenglonFirst.Y
                
                For iCount = 1 To UBound(regDocument.arrRow)
                    Y = .CurrentY
                    
                    .FontItalic = False
                    
                    .CurrentX = regPrintCnfg3.RenglonArtCodigo.X
                    If regPrintCnfg3.RenglonArtCodigo.Size <> 0 Then .FontSize = regPrintCnfg3.RenglonArtCodigo.Size
                    If regPrintCnfg3.RenglonArtCodigo.Bold Then .FontBold = True
                    Select Case UCase(regPrintCnfg3.RenglonArtCodigo.Align)
                        Case "C": vsView.TextAlign = taCenterTop
                        Case "R": vsView.TextAlign = taRightTop
                    End Select
                    If regPrintCnfg3.RenglonArtCodigo.W > 0 Then
                        .TextBox regDocument.arrRow(iCount).Codigo, regPrintCnfg.RenglonArtCodigo.X, Y, regPrintCnfg.RenglonArtCodigo.W, .TextHeight(regDocument.arrRow(iCount).Codigo) + 40
                    Else
                        .CurrentX = regPrintCnfg3.RenglonArtCodigo.X
                        .CurrentY = Y
                        .Paragraph = regDocument.arrRow(iCount).Codigo
                    End If
                    .FontBold = False
                    .FontSize = lFontSizeG
                    .TextAlign = taLeftTop
                        
                    'CANTIDAD
                    .CurrentX = regPrintCnfg3.RenglonArtNombre.X
                    If regPrintCnfg3.RenglonArtNombre.Size <> 0 Then .FontSize = regPrintCnfg3.RenglonArtNombre.Size
                    If regPrintCnfg3.RenglonArtNombre.Bold Then .FontBold = True
                    Select Case UCase(regPrintCnfg3.RenglonArtNombre.Align)
                        Case "C": vsView.TextAlign = taCenterTop
                        Case "R": vsView.TextAlign = taRightTop
                    End Select
                    If regPrintCnfg3.RenglonArtNombre.W > 0 Then
                        .TextBox regDocument.arrRow(iCount).Articulo, regPrintCnfg.RenglonArtNombre.X, Y, regPrintCnfg.RenglonArtNombre.W, .TextHeight(regDocument.arrRow(iCount).Articulo) + 40
                    Else
                        .CurrentX = regPrintCnfg3.RenglonArtNombre.X
                        .CurrentY = Y
                        .Paragraph = regDocument.arrRow(iCount).Articulo
                    End If
                    .FontBold = False
                    .FontSize = lFontSizeG
                    .TextAlign = taLeftTop
                        
                    .CurrentX = regPrintCnfg3.RenglonCantidad.X
                    If regPrintCnfg3.RenglonCantidad.Size <> 0 Then .FontSize = regPrintCnfg3.RenglonCantidad.Size
                    If regPrintCnfg3.RenglonCantidad.Bold Then .FontBold = True
                    Select Case UCase(regPrintCnfg3.RenglonCantidad.Align)
                        Case "C": vsView.TextAlign = taCenterTop
                        Case "R": vsView.TextAlign = taRightTop
                    End Select
                    If regPrintCnfg3.RenglonCantidad.W > 0 Then
                        .TextBox regDocument.arrRow(iCount).Cantidad, regPrintCnfg.RenglonCantidad.X, Y, regPrintCnfg.RenglonCantidad.W, .TextHeight(regDocument.arrRow(iCount).Cantidad) + 40
                    Else
                        .CurrentX = regPrintCnfg3.RenglonCantidad.X
                        .CurrentY = Y
                        .Paragraph = regDocument.arrRow(iCount).Cantidad
                    End If
                    .FontBold = False
                    .FontSize = lFontSizeG
                    .TextAlign = taLeftTop
                        
                    'Hago el renglon siguiente
                    .CurrentY = Y: .Paragraph = ""
    
                Next iCount
                '..................................................................................................
                
                'Comentario del documento
                If regDocument.DocumentoComentario <> "" And (regPrintCnfg3.DocumentoComentario.X <> 0 Or regPrintCnfg3.DocumentoComentario.Y <> 0) Then
                    If regPrintCnfg3.DocumentoComentario.Size <> 0 Then .FontSize = regPrintCnfg3.DocumentoComentario.Size
                    If regPrintCnfg3.DocumentoComentario.Bold Then vsView.FontBold = True
                    Select Case UCase(regPrintCnfg3.DocumentoComentario.Align)
                        Case "C": vsView.TextAlign = taCenterTop
                        Case "R": vsView.TextAlign = taRightTop
                    End Select
                    
                    If regPrintCnfg3.DocumentoComentario.W > 0 Then
                        .TextBox regDocument.DocumentoComentario, regPrintCnfg.DocumentoComentario.X, regPrintCnfg.DocumentoComentario.Y, regPrintCnfg.DocumentoComentario.W, .TextHeight("DOCUMENTO EN PESOS") + 40
                    Else
                        .CurrentX = regPrintCnfg3.DocumentoComentario.X
                        .CurrentY = regPrintCnfg3.DocumentoComentario.Y
                        .Paragraph = regDocument.DocumentoComentario
                    End If
                    vsView.FontBold = False
                    vsView.FontSize = lFontSizeG
                End If
            End If
        End If
        .EndDoc
        .PrintDoc False
        
    End With
        
    'Si llegue aqui sin errores marco como impreso.
    PrintDocument = True
        
    Exit Function
errPrint:
    MsgBox "Error al imprimir." & vbCr & Err.Number & "-" & Err.Description, vbCritical, "Imprimir documento"
End Function

Public Function InitDevicePrinter(ByVal vsView As vsPrinter) As String
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
    Open prmPathApp & "\rpttraslado.txt" For Input As iFile
    Do While Not EOF(iFile)
        Line Input #iFile, sAux
        'Proceso la línea
        ReDim Preserve arrCfgPrint(iCont)
        arrCfgPrint(iCont) = sAux
        iCont = iCont + 1
    Loop
    Close iFile
    
    On Error GoTo errOFC
    For iCont = 0 To UBound(arrCfgPrint)
        If arrCfgPrint(iCont) <> "" Then
            If InStr(1, arrCfgPrint(iCont), "]") > 0 Then
                sAux = Mid(arrCfgPrint(iCont), 1, InStr(1, arrCfgPrint(iCont), "]"))
                Select Case LCase(sAux)
                    Case "[device]"
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
                    Case "[papersize]"
                        If IsNumeric(Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)) Then
                            iPaperSize = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                        End If
                        
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
    MsgBox "Error al leer el archivo de configuración." & vbCr & Err.Number & "-" & Err.Description, vbCritical, "Iniciar impresora"
    Exit Function
errOFC:
    Screen.MousePointer = 0
    MsgBox "Error al setear la configuración de la impresora." & vbCr & Err.Number & "-" & Err.Description, vbCritical, "Iniciar impresora"
End Function

Private Sub loc_GetPosXY()
Dim iCont As Integer
Dim sAux As String, sAlign As String
Dim X As Long, Y As Long, W As Long, lFS As Long
Dim bBold As Boolean

    For iCont = 0 To UBound(arrCfgPrint)
        If arrCfgPrint(iCont) <> "" Then
            If InStr(1, arrCfgPrint(iCont), "]") > 0 Then
                
                sAux = loc_GetKeyCnfg(arrCfgPrint(iCont), X, Y, W, lFS, bBold, sAlign)
                
                Select Case LCase(sAux)
                    Case LCase("[Documento]")
                        With regPrintCnfg.Documento
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[DocumentoNumero]")
                        With regPrintCnfg.DocumentoNumero
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With

                    Case LCase("[TableFechaP]")
                        With regPrintCnfg.TableFechaP
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[TableFechaD]")
                        regPrintCnfg.TableFechaD = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                    
                    Case LCase("[TableFechaR]")
                        With regPrintCnfg.TableFechaR
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[ClienteNombre]"):
                        With regPrintCnfg.ClienteNombre
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                        
                    Case LCase("[ClienteDireccion]")
                        With regPrintCnfg.ClienteDireccion
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                        
                    Case LCase("[RenglonArtCodigo]")
                        With regPrintCnfg.RenglonArtCodigo
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                        End With
                    Case LCase("[RenglonArtNombre]")
                        With regPrintCnfg.RenglonArtNombre
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                        End With
                    Case LCase("[RenglonCantidad]")
                        With regPrintCnfg.RenglonCantidad
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                        End With
                    
                    Case LCase("[RenglonFirstTable]")
                        With regPrintCnfg.TablaRenglonFirst
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[DocumentoComentario]")
                        With regPrintCnfg.DocumentoComentario
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
'                    Case LCase("[RenglonTable]"): regPrintCnfg.TablaRenglon = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
'2da vía
                    Case LCase("[Documento2]")
                        With regPrintCnfg2.Documento
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[DocumentoNumero2]")
                        With regPrintCnfg2.DocumentoNumero
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[TableFechaP2]")
                        With regPrintCnfg2.TableFechaP
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[TableFechaD2]")
                        regPrintCnfg2.TableFechaD = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                    
                    Case LCase("[TableFechaR2]")
                        With regPrintCnfg2.TableFechaR
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[ClienteNombre2]"):
                        With regPrintCnfg2.ClienteNombre
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                        
                    Case LCase("[ClienteDireccion2]")
                        With regPrintCnfg2.ClienteDireccion
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                        
                    Case LCase("[RenglonArtCodigo2]")
                        With regPrintCnfg2.RenglonArtCodigo
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                        End With
                    Case LCase("[RenglonArtNombre2]")
                        With regPrintCnfg2.RenglonArtNombre
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                        End With
                    Case LCase("[RenglonCantidad2]")
                        With regPrintCnfg2.RenglonCantidad
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                        End With
                    Case LCase("[RenglonFirstTable2]")
                        With regPrintCnfg2.TablaRenglonFirst
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    Case LCase("[DocumentoComentario2]")
                        With regPrintCnfg2.DocumentoComentario
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                        
                        
'3ERA vía---------------------------------------------------------------------------------------
                    Case LCase("[Documento3]")
                        With regPrintCnfg3.Documento
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[DocumentoNumero3]")
                        With regPrintCnfg3.DocumentoNumero
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[TableFechaP3]")
                        With regPrintCnfg3.TableFechaP
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[TableFechaD3]")
                        regPrintCnfg3.TableFechaD = Mid(arrCfgPrint(iCont), InStr(1, arrCfgPrint(iCont), "]") + 1)
                    
                    Case LCase("[TableFechaR3]")
                        With regPrintCnfg3.TableFechaR
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    
                    Case LCase("[ClienteNombre3]"):
                        With regPrintCnfg3.ClienteNombre
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                        
                    Case LCase("[ClienteDireccion3]")
                        With regPrintCnfg3.ClienteDireccion
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                        
                    Case LCase("[RenglonArtCodigo3]")
                        With regPrintCnfg3.RenglonArtCodigo
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                        End With
                    Case LCase("[RenglonArtNombre3]")
                        With regPrintCnfg3.RenglonArtNombre
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                        End With
                    Case LCase("[RenglonCantidad3]")
                        With regPrintCnfg3.RenglonCantidad
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                            If sAlign <> "" Then .Align = sAlign
                        End With
                    Case LCase("[RenglonFirstTable3]")
                        With regPrintCnfg3.TablaRenglonFirst
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                    Case LCase("[DocumentoComentario3]")
                        With regPrintCnfg3.DocumentoComentario
                            .X = X: .Y = Y: .W = W
                            .Bold = bBold
                            If lFS <> 0 Then .Size = lFS
                        End With
                        
                End Select
            End If
        End If
    Next
    
    
End Sub

Private Function loc_GetKeyCnfg(ByVal sLine As String, X As Long, Y As Long, W As Long, _
                    lFS As Long, bBold As Boolean, sAlign As String) As String
Dim vDatos() As String
    loc_GetKeyCnfg = ""
    X = 0
    Y = 0
    W = 0
    lFS = 0
    bBold = False
    sAlign = ""
    
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
                    
                    If UBound(vDatos) >= 5 Then
                        sAlign = UCase(vDatos(5))
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


Private Sub loc_PrintFieldLabel(ByVal oFP As clsRegFieldPrint, ByVal vsView As vsPrinter, ByVal vValue As Variant)
    Dim iP As Integer
    Select Case oFP.PropsFP.Align
        Case 0: vsView.TextAlign = taLeftTop
        Case taCenterTop: vsView.Text = taCenterTop
        Case taRightTop: vsView.TextAlign = taRightTop
    End Select
    If oFP.PropsFP.FontSize > 0 Then vsView.FontSize = oFP.PropsFP.FontSize
    vsView.FontBold = oFP.PropsFP.Bold
    'Para c/u de las posiciones las imprimo
    For iP = 1 To oFP.PropsFP.Posicion.Count
        If oFP.PropsFP.Width > 0 Then
            vsView.TextBox vValue, oFP.PropsFP.Posicion.GetPosicion(iP).posX, oFP.PropsFP.Posicion.GetPosicion(iP).posY, oFP.PropsFP.Width, vsView.TextHeight(CStr(vValue)) + 40
        Else
            vsView.CurrentX = oFP.PropsFP.Posicion.GetPosicion(iP).posX
            vsView.CurrentY = oFP.PropsFP.Posicion.GetPosicion(iP).posY
            vsView.Paragraph = IIf(IsNull(vValue), "", vValue)
        End If
    Next
End Sub

Private Sub loc_PrintEtiquetas(ByVal vsView As vsPrinter)
    Dim iQ As Integer
    Dim oFP As clsRegFieldPrint
    For iQ = 1 To oFieldPrint.Count
        Set oFP = oFieldPrint.GetField(iQ)
        If oFP.Tipo = Label Then loc_PrintFieldLabel oFP, vsView, oFP.Nombre
    Next
End Sub

Private Sub loc_PrintField(ByVal vsView As vsPrinter, ByVal sNameField As String, ByVal vValue As Variant)
    Dim iQ As Integer
    Dim oFP As clsRegFieldPrint
    For iQ = 1 To oFieldPrint.Count
        Set oFP = oFieldPrint.GetField(iQ)
        If oFP.Tipo = Campo And LCase(oFP.Nombre) = LCase(sNameField) Then loc_PrintFieldLabel oFP, vsView, vValue
    Next
End Sub

