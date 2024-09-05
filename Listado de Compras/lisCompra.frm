VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Begin VB.Form LisCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compra de Mercadería"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "lisCompra.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   9240
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   120
      TabIndex        =   12
      Top             =   900
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   14
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "configurarI"
            Object.ToolTipText     =   "Configurar impresora"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "configurarP"
            Object.ToolTipText     =   "Preparar página"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   600
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "primero"
            Object.ToolTipText     =   "Primero"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "anterior"
            Object.ToolTipText     =   "Anterior"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "siguiente"
            Object.ToolTipText     =   "Siguiente"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ultimo"
            Object.ToolTipText     =   "Último"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "zoomout"
            Object.ToolTipText     =   "Zoom out"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "zoomin"
            Object.ToolTipText     =   "Zoom in"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.ComboBox cZoom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   30
         Width           =   840
      End
   End
   Begin VB.ComboBox cTipo 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox tHasta 
      Height          =   285
      Left            =   2640
      MaxLength       =   12
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton bConsultar 
      Caption         =   "&Consultar"
      Height          =   340
      Left            =   8040
      TabIndex        =   10
      Top             =   450
      Width           =   975
   End
   Begin VB.ComboBox cMoneda 
      Height          =   315
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox tDesde 
      Height          =   285
      Left            =   840
      MaxLength       =   12
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cProveedor 
      Height          =   315
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   120
      Width           =   3015
   End
   Begin vsViewLib.vsPrinter vsPrinter 
      Height          =   4335
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   9015
      _Version        =   196608
      _ExtentX        =   15901
      _ExtentY        =   7646
      _StockProps     =   229
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Zoom            =   70
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Hasta:"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Moneda:"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   525
      Width           =   735
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   -120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisCompra.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisCompra.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisCompra.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisCompra.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisCompra.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisCompra.frx":13C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisCompra.frx":1C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisCompra.frx":1FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisCompra.frx":22D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisCompra.frx":25EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Desde:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Proveedor:"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   795
      Left            =   120
      Top             =   45
      Width           =   9015
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuOpConsultar 
         Caption         =   "&Consultar"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuOpImprimir 
         Caption         =   "&Imprimr"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuOpLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpConfigurar 
         Caption         =   "C&onfigurar Impresora"
      End
      Begin VB.Menu MnuOpPagina 
         Caption         =   "&Preparar Página"
      End
   End
   Begin VB.Menu MnuAcciones 
      Caption         =   "&Acciones"
      Begin VB.Menu MnuAcPrimera 
         Caption         =   "&Primera Página"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuAcAnterior 
         Caption         =   "Página &Anterior"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuAcSiguiente 
         Caption         =   "Página &Siguiente"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuAcUltima 
         Caption         =   "&Última Página"
         Shortcut        =   ^U
      End
      Begin VB.Menu MnuAcLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAcZoomIn 
         Caption         =   "Zoom &In"
      End
      Begin VB.Menu MnuAcZoomOut 
         Caption         =   "Zoom &out"
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalFormulario 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "LisCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFormato As String
Dim strEncabezado As String

Private Sub AccionConsultar()
On Error GoTo ErrBC

    Screen.MousePointer = vbHourglass
    If Not ValidoDatos Then Screen.MousePointer = vbDefault: Exit Sub
    
    'Inicio el documento.
    vsPrinter.StartDoc
    If vsPrinter.Error Then MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault: Exit Sub
    
    EncabezadoListado vsPrinter, "Compras " & Trim(cTipo.Text) & " desde " & Format(tDesde.Text, "d/mm/yy") & " al " & Format(tHasta.Text, "d/mm/yy"), True

    CargoReporte cProveedor.ItemData(cProveedor.ListIndex), cMoneda.ItemData(cMoneda.ListIndex), cTipo.ItemData(cTipo.ListIndex)
    
    vsPrinter.EndDoc

    Screen.MousePointer = vbDefault
    Exit Sub

ErrBC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al consultar.", Err.Description
End Sub

Private Sub PropiedadesImpresion()

  With vsPrinter
        .PreviewPage = 1
        .FontName = "Tahoma"
        .FontSize = 10
        .FontBold = False
        .FontItalic = False
        .TextAlign = 0  'Left
        .PageBorder = 3 ' pbTop
        .PenStyle = 0
        .BrushStyle = 0
        .PenWidth = 2
        .PenColor = 0
        .BrushColor = 0
        .TextColor = 0
        .Columns = 1
        .TableBorder = tbBoxRows
        .MarginRight = 400
        .Orientation = orLandscape
        .ColorMode = 1
    End With

End Sub

Private Sub CargoReporte(Proveedor As Long, Moneda As Integer, Tipo As Integer)
    
Dim aCompra As Long, Contador As Integer
Dim aTotal As Currency, aIva As Currency, aTTotal As Currency, aTIva As Currency
Dim aUnitario As Currency
    
    Cons = ArmoConsulta(Proveedor, Moneda, Tipo)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    If Not RsAux.EOF Then
        ArmoEncabezado
        aCompra = 0: Contador = 0: aTTotal = 0: aTIva = 0
        
        Do While Not RsAux.EOF
            If aCompra <> RsAux!ComCodigo Then
                If aCompra <> 0 Then    'Inserto Total Factura
                    aTexto = "||||||" _
                              & Format(aTotal - aIva, FormatoMonedaP) & "|" _
                              & Format(aIva, FormatoMonedaP) & "|" _
                              & Format(aTotal, FormatoMonedaP) & "|"
                    
                    vsPrinter.AddTable strFormato, "", aTexto, , vbWhite, True
                    Contador = 0
                End If
                aCompra = RsAux!ComCodigo
                aTotal = RsAux!ComImporte * RsAux!ComTC
                aIva = RsAux!ComIva * RsAux!ComTC
                aTTotal = aTTotal + aTotal
                aTIva = aTIva + aIva
            End If
            
            Contador = Contador + 1
            If Contador = 1 Then      'Cargo la línea de la factura------------------------------------------
                If RsAux!ComTipoDocumento = TipoDocumento.CompraContado Then
                    aTexto = "* " & Format(RsAux!ComFecha, "dd/mm/yy") & "|"
                Else
                    aTexto = Format(RsAux!ComFecha, "dd/mm/yy") & "|"
                End If
                aTexto = aTexto & Trim(RsAux!PClNombre) & "|"
                If Not IsNull(RsAux!ComSerie) Then aTexto = aTexto & Trim(RsAux!ComSerie) & " "
                If Not IsNull(RsAux!ComNumero) Then aTexto = aTexto & Trim(RsAux!ComNumero)
                aTexto = aTexto & "|"
                
            Else
                aTexto = "|||"
            End If
            
            aUnitario = RsAux!CRePrecioU * RsAux!ComTC
            
            aTexto = aTexto & Trim(RsAux!ArtCodigo) & " " & Trim(RsAux!ArtNombre) & "|" _
                                    & RsAux!CReCantidad & "|" _
                                    & Format(aUnitario, FormatoMonedaP) & "|" _
                                    & Format(aUnitario * RsAux!CReCantidad, FormatoMonedaP)
            
            vsPrinter.AddTable strFormato, "", aTexto, , vbWhite, True  '--------------------------------------

            RsAux.MoveNext
        Loop
        
        'Inserto el Último Total ------------------------------------------------------
        aTexto = "||||||" _
                  & Format(aTotal - aIva, FormatoMonedaP) & "|" _
                  & Format(aIva, FormatoMonedaP) & "|" _
                  & Format(aTotal, FormatoMonedaP)
        vsPrinter.AddTable strFormato, "", aTexto, , vbWhite, True
        '--------------------------------------------------------------------------------
        
        'Inserto el Total de las Compras--------------------------------------------
        aTexto = "|Resumen de Compras|||||" _
                  & Format(aTTotal - aTIva, FormatoMonedaP) & "|" _
                  & Format(aTIva, FormatoMonedaP) & "|" _
                  & Format(aTTotal, FormatoMonedaP)
        'vsPrinter.FontBold = True
        vsPrinter.TextColor = vbWhite
        vsPrinter.AddTable strFormato, "", aTexto, vbWhite, vbBlue, True
        '--------------------------------------------------------------------------------
    End If
    RsAux.Close
    
End Sub

Private Sub ArmoEncabezado()

    strFormato = "+<950|+<2450|+<950|+<3900|+>900|+>1000|+>1400|+>1100|+>1400"
    strEncabezado = "Fecha|Proveedor|Factura|Artículo|Cantidad|Unitario|Neto|I.V.A.|Total"
    
    With vsPrinter
        .FontSize = 8
        .FontBold = True
        .TextColor = vbWhite
        .AddTable strFormato, strEncabezado, "", vbBlue
        
        .FontBold = False
        .FontItalic = False
        .TextColor = vbBlack
    End With

End Sub

Private Function ArmoConsulta(Proveedor As Long, Moneda As Integer, Tipo As Integer) As String

    Cons = " Select * from Compra, CompraRenglon, ProveedorCliente, Articulo" _
            & " Where ComFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'"
            
    If Tipo <> 0 Then
        Cons = Cons & " And ComTipoDocumento = " & Tipo
    Else
        Cons = Cons & " And ComTipoDocumento In (" & TipoDocumento.CompraContado & ", " & TipoDocumento.CompraCredito & ")"
    End If
    
    If Moneda <> 0 Then Cons = Cons & " And ComMoneda = " & Moneda
            
    If Proveedor <> 0 Then Cons = Cons & " And ComProveedor = " & Proveedor
    
    Cons = Cons _
        & " And ComCodigo = CReCompra " _
        & " And ComProveedor = PClCodigo " _
        & " And CReArticulo = ArtId " _
        & " Order by ComFecha"
        
    ArmoConsulta = Cons
    
End Function

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub cProveedor_Change()
    Selecciono cProveedor, cProveedor.Text, gTecla
End Sub
Private Sub cProveedor_GotFocus()
    cProveedor.SelStart = 0
    cProveedor.SelLength = Len(cProveedor.Text)
End Sub

Private Sub cProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cProveedor.ListIndex
End Sub

Private Sub cProveedor_KeyPress(KeyAscii As Integer)
    cProveedor.ListIndex = gIndice
    If KeyAscii = vbKeyReturn And cProveedor.ListIndex > -1 Then Foco cTipo
End Sub

Private Sub cProveedor_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cProveedor
End Sub

Private Sub cProveedor_LostFocus()
    gIndice = -1
    cProveedor.SelLength = 0
End Sub

Private Sub cTipo_Change()
    Selecciono cTipo, cTipo.Text, gTecla
End Sub
Private Sub cTipo_GotFocus()
    cTipo.SelStart = 0
    cTipo.SelLength = Len(cTipo.Text)
End Sub

Private Sub cTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cTipo.ListIndex
End Sub

Private Sub cTipo_KeyPress(KeyAscii As Integer)
    cTipo.ListIndex = gIndice
    If KeyAscii = vbKeyReturn And cTipo.ListIndex > -1 Then Foco cMoneda
End Sub

Private Sub cTipo_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cTipo
End Sub

Private Sub cTipo_LostFocus()
    gIndice = -1
    cTipo.SelLength = 0
End Sub

Private Sub cZoom_Click()
    Zoom vsPrinter, CInt(cZoom.Text)
End Sub

Private Sub Form_Activate()
    Me.Refresh
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad

    SetToolBarFlat Toolbar1, True
    
    CargoDocumento
    CargoProveedor
    CargoMoneda
    CargoDatosZoom cZoom
    cZoom.ListIndex = 6
    
    'Cargo la fecha del servidor.
    FechaDelServidor
    tDesde.Text = Format(PrimerDia(gFechaServidor), "d-Mmm-yyyy")
    tHasta.Text = Format(UltimoDia(gFechaServidor), "d-Mmm-yyyy")
    PropiedadesImpresion
    Exit Sub
    
ErrLoad:
    Screen.MousePointer = vbDefault
    ErrorInesperado Err.Description
End Sub

Private Sub CargoDocumento()

    cTipo.Clear
  
    cTipo.AddItem "(Todos)"
    cTipo.ItemData(cTipo.NewIndex) = 0

    cTipo.AddItem DocContado
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraContado
    
    cTipo.AddItem DocCredito
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraCredito
    
    'Por defecto pongo todos.
    For I = 0 To cTipo.ListCount - 1
        If cTipo.ItemData(I) = 0 Then cTipo.ListIndex = I: Exit For
    Next
    
End Sub

Private Sub CargoProveedor()

    'Cargo los PROVEEDORES DE MERCADERIA
    Cons = "Select PMeCodigo, PMeNombre From ProveedorMercaderia Order by PMENombre"
    CargoCombo Cons, cProveedor, ""
    
    cProveedor.AddItem "(Todos)"
    cProveedor.ItemData(cProveedor.NewIndex) = 0
    
    'Por defecto pongo todos.
    For I = 0 To cProveedor.ListCount - 1
        If cProveedor.ItemData(I) = 0 Then cProveedor.ListIndex = I: Exit For
    Next
    
End Sub

Private Function ValidoDatos() As Boolean

    ValidoDatos = False
    
    If Not IsDate(tDesde.Text) Then
        MsgBox "La fecha desde ingresada no es correcta, verifique.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    If Not IsDate(tHasta.Text) Then
        MsgBox "La fecha hasta ingresada no es correcta, verifique.", vbExclamation, "ATENCIÓN"
        Foco tHasta: Exit Function
    End If
    
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then
        MsgBox "El período de fechas ingresado no es correcta, verifique.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    If cProveedor.ListIndex = -1 Then
        MsgBox "Debe seleccionar un proveedor para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco cProveedor: Exit Function
    End If
    
    If cTipo.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de venta para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco cTipo: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Seleccione una moneda para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    ValidoDatos = True
    
End Function


Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then Exit Sub
    vsPrinter.Left = 80
    vsPrinter.Width = Me.Width - 240
    Toolbar1.Width = vsPrinter.Width
    Shape1.Width = vsPrinter.Width
    vsPrinter.Height = Me.Height - 2050

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Forms(Forms.Count - 2).SetFocus
End Sub

Private Sub Label1_Click()
    Foco cProveedor
End Sub

Private Sub Label2_Click()
    Foco tDesde
End Sub

Private Sub Label3_Click()
    Foco tHasta
End Sub

Private Sub Label4_Click()
    Foco cMoneda
End Sub

Private Sub Label5_Click()
    Foco cTipo
End Sub

Private Sub MnuAcAnterior_Click()
    IrAPagina vsPrinter, vsPrinter.PreviewPage - 1
End Sub

Private Sub MnuAcPrimera_Click()
    IrAPagina vsPrinter, 1
End Sub

Private Sub MnuAcSiguiente_Click()
    IrAPagina vsPrinter, vsPrinter.PreviewPage + 1
End Sub

Private Sub MnuAcUltima_Click()
    IrAPagina vsPrinter, vsPrinter.PageCount
End Sub

Private Sub MnuAcZoomIn_Click()
    ZoomIn vsPrinter
End Sub

Private Sub MnuAcZoomOut_Click()
    ZoomOut vsPrinter
End Sub

Private Sub MnuOpConfigurar_Click()
    vsPrinter.PrintDialog (pdPrinterSetup)
End Sub

Private Sub MnuOpConsultar_Click()
    vsPrinter.Preview = True:    AccionConsultar
End Sub

Private Sub MnuOpImprimir_Click()
    AccionImprimir
End Sub

Private Sub MnuOpPagina_Click()
    vsPrinter.PrintDialog (pdPageSetup)
End Sub

Private Sub MnuSalFormulario_Click()
    Unload Me
End Sub

Private Sub tDesde_GotFocus()
    tDesde.SelStart = 0
    tDesde.SelLength = Len(tDesde.Text)
End Sub

Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tHasta
End Sub

Private Sub tDesde_LostFocus()
    If Not IsDate(tDesde.Text) Then
        MsgBox "No se ingreso una fecha válida.", vbExclamation, "ATENCIÓN"
    Else
        tDesde.Text = Format(tDesde.Text, "d-Mmm-yyyy")
    End If
End Sub


Private Sub tHasta_GotFocus()
    tHasta.SelStart = 0
    tHasta.SelLength = Len(tHasta.Text)
End Sub

Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cProveedor
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Toolbar1.Refresh
    
    Select Case Button.Key
        Case "zoomin": ZoomIn vsPrinter
        Case "zoomout": ZoomOut vsPrinter
                
        Case "imprimir": AccionImprimir
                
        Case "configurarI": vsPrinter.PrintDialog (pdPrinterSetup)
        Case "configurarP": vsPrinter.PrintDialog (pdPageSetup)
  
        Case "primero": IrAPagina vsPrinter, 1
        Case "anterior": IrAPagina vsPrinter, vsPrinter.PreviewPage - 1
        Case "siguiente": IrAPagina vsPrinter, vsPrinter.PreviewPage + 1
        Case "ultimo": IrAPagina vsPrinter, vsPrinter.PageCount
        
        Case "consultar": vsPrinter.Preview = True: AccionConsultar
    End Select

End Sub

Private Sub vsPrinter_EndDoc()
    
    EnumeroPiedePagina vsPrinter

End Sub

Private Sub CargoMoneda()
On Error GoTo ErrCM

    Cons = "Select MonCodigo, MonSigno From Moneda"
    CargoCombo Cons, cMoneda, ""
    
    cMoneda.AddItem "(Todos)"
    cMoneda.ItemData(cMoneda.NewIndex) = 0
    
    'Por defecto pongo todos.
    For I = 0 To cMoneda.ListCount - 1
        If cMoneda.ItemData(I) = 0 Then cMoneda.ListIndex = I: Exit For
    Next
    Exit Sub

ErrCM:
    Screen.MousePointer = vbDefault
    clsError.MuestroError "Ocurrió un error al cargar las monedas."
End Sub

Private Sub cMoneda_Change()
    Selecciono cMoneda, cMoneda.Text, gTecla
End Sub

Private Sub cMoneda_GotFocus()
    cMoneda.SelStart = 0
    cMoneda.SelLength = Len(cMoneda.Text)
End Sub

Private Sub cMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cMoneda.ListIndex
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)

    cMoneda.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus

End Sub

Private Sub cMoneda_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cMoneda
End Sub

Private Sub cMoneda_LostFocus()
    gIndice = -1
    cMoneda.SelLength = 0
End Sub

Private Sub AccionImprimir()

    On Error GoTo errPrint
    vsPrinter.filename = "Compras de Mercadería"
    vsPrinter.PrintDoc
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al impirmir. " & Trim(Err.Description)
End Sub

