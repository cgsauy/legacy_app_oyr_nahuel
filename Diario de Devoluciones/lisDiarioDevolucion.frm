VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Begin VB.Form LisDiarioDevolucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diario de Devoluciones"
   ClientHeight    =   5835
   ClientLeft      =   705
   ClientTop       =   1890
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "lisDiarioDevolucion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   9165
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   120
      TabIndex        =   7
      Top             =   660
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
         TabIndex        =   8
         Top             =   30
         Width           =   840
      End
   End
   Begin VB.ComboBox cMoneda 
      Height          =   315
      Left            =   8220
      Sorted          =   -1  'True
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   160
      Width           =   855
   End
   Begin VB.TextBox tFecha 
      Height          =   285
      Left            =   3600
      MaxLength       =   12
      TabIndex        =   3
      Top             =   160
      Width           =   1095
   End
   Begin VB.ComboBox cDocumento 
      Height          =   315
      Left            =   5760
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   160
      Width           =   1695
   End
   Begin VB.ComboBox cSucursal 
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   160
      Width           =   1935
   End
   Begin vsViewLib.vsPrinter vsPrinter 
      Height          =   4575
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   9015
      _Version        =   196608
      _ExtentX        =   15901
      _ExtentY        =   8070
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
      Zoom            =   80
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Moneda:"
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   160
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
            Picture         =   "lisDiarioDevolucion.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisDiarioDevolucion.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisDiarioDevolucion.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisDiarioDevolucion.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisDiarioDevolucion.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisDiarioDevolucion.frx":13C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisDiarioDevolucion.frx":1C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisDiarioDevolucion.frx":1FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisDiarioDevolucion.frx":22D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisDiarioDevolucion.frx":25EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha:"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   160
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Documento:"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Sucursal:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   160
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   550
      Left            =   120
      Top             =   50
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
Attribute VB_Name = "LisDiarioDevolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFormato As String
Dim strEncabezado As String

Dim cParcialNota As Currency
Dim cParcialIvaNota As Currency

Dim cTotalNota As Currency
Dim cIvaNota As Currency

Dim cParcialContado As Currency
Dim cParcialIva As Currency

Dim cTotalContado As Currency
Dim cTotalIva As Currency
Private strTexto As String

'Dim fBar As New AAProgress, aValor As Long
Private Sub AccionConsultar()
On Error GoTo ErrBC

    Screen.MousePointer = vbHourglass
    'Inicio el documento.
    vsPrinter.StartDoc
    If vsPrinter.Error Then MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault: Exit Sub
    
    'Controlo si ingreso los parámetros.
    If Not ValidoDatos Then Screen.MousePointer = vbDefault: Exit Sub
    
    'Pongo propiedades por defecto.
    PropiedadesImpresion
    
    'Título y Pie de impresión.
    CreoTitulo

    'Creo la tabla en base a mi formato
    If cDocumento.ItemData(cDocumento.ListIndex) = 0 Then
        HagoTablaConsultaDevoluciones
    ElseIf cDocumento.ItemData(cDocumento.ListIndex) = TipoDocumento.NotaCredito Then
        HagoTablaConsultaNotaCredito
    ElseIf cDocumento.ItemData(cDocumento.ListIndex) = TipoDocumento.NotaDevolucion Then
        HagoTablaConsultaNotaDevolucion
    ElseIf cDocumento.ItemData(cDocumento.ListIndex) = TipoDocumento.NotaEspecial Then
        HagoTablaConsultaNotaEspecial
    End If
    
    'Cierro el documento.
    vsPrinter.EndDoc

    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrBC:
    On Error Resume Next
    fBar.Drop
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al consultar."
    cBase.QueryTimeout = 15
End Sub
Private Sub PropiedadesImpresion()

  With vsPrinter
        .PreviewPage = 1
        .FontName = "Tahoma"
        .FontSize = 9
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
    End With

End Sub
Private Sub HagoTablaConsultaDevoluciones()

    'Por temor a que demore la consulta aumentamos el QueryTimeOut
    cBase.QueryTimeout = 120
    cTotalContado = 0
    cTotalIva = 0
    cIvaNota = 0
    cTotalNota = 0
    
    If cSucursal.ItemData(cSucursal.ListIndex) <> 0 Then
        CargoTablaNotaDevolucion cSucursal.ItemData(cSucursal.ListIndex), cSucursal.Text, cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text, TipoDocumento.NotaDevolucion
        InsertoSubTotalSucursal cSucursal.Text
        CargoTablaNotaCredito cSucursal.ItemData(cSucursal.ListIndex), cSucursal.Text, cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text
        InsertoSubTotalSucursal cSucursal.Text
        CargoTablaNotaDevolucion cSucursal.ItemData(cSucursal.ListIndex), cSucursal.Text, cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text, TipoDocumento.NotaEspecial
        InsertoSubTotalSucursal cSucursal.Text
    Else
        For I = 0 To cSucursal.ListCount - 1
            If cSucursal.ItemData(I) <> 0 Then
                CargoTablaNotaDevolucion cSucursal.ItemData(I), cSucursal.List(I), cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text, TipoDocumento.NotaDevolucion
                InsertoSubTotalSucursal cSucursal.List(I)
                CargoTablaNotaCredito cSucursal.ItemData(I), cSucursal.List(I), cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text
                InsertoSubTotalSucursal cSucursal.List(I)
                CargoTablaNotaDevolucion cSucursal.ItemData(I), cSucursal.List(I), cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text, TipoDocumento.NotaEspecial
                InsertoSubTotalSucursal cSucursal.List(I)
            End If
        Next I
        TotalMonedaSucursalCredito cMoneda.Text
    End If
    cBase.QueryTimeout = 15

End Sub
Private Sub HagoTablaConsultaNotaCredito()

    'Por temor a que demore la consulta aumentamos el QueryTimeOut
    cBase.QueryTimeout = 120
    cTotalContado = 0
    cTotalIva = 0
    cIvaNota = 0
    cTotalNota = 0
    
    If cSucursal.ItemData(cSucursal.ListIndex) <> 0 Then
        CargoTablaNotaCredito cSucursal.ItemData(cSucursal.ListIndex), cSucursal.Text, cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text
        InsertoSubTotalSucursal cSucursal.Text
    Else
        For I = 0 To cSucursal.ListCount - 1
            If cSucursal.ItemData(I) <> 0 Then
                CargoTablaNotaCredito cSucursal.ItemData(I), cSucursal.List(I), cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text
                InsertoSubTotalSucursal cSucursal.List(I)
            End If
        Next I
        TotalMonedaSucursalCredito cMoneda.Text
    End If
    cBase.QueryTimeout = 15

End Sub
Private Sub HagoTablaConsultaNotaDevolucion()

    'Por temor a que demore la consulta aumentamos el QueryTimeOut
    cBase.QueryTimeout = 120
    cTotalContado = 0
    cTotalIva = 0
    cIvaNota = 0
    cTotalNota = 0
    
    If cSucursal.ItemData(cSucursal.ListIndex) <> 0 Then
        CargoTablaNotaDevolucion cSucursal.ItemData(cSucursal.ListIndex), cSucursal.Text, cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text, TipoDocumento.NotaDevolucion
        InsertoSubTotalSucursal cSucursal.Text, True
    Else
        For I = 0 To cSucursal.ListCount - 1
            If cSucursal.ItemData(I) <> 0 Then
                CargoTablaNotaDevolucion cSucursal.ItemData(I), cSucursal.List(I), cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text, TipoDocumento.NotaDevolucion
                InsertoSubTotalSucursal cSucursal.List(I), True
            End If
        Next I
        TotalMonedaSucursal cMoneda.Text
    End If
    cBase.QueryTimeout = 15

End Sub
Private Sub HagoTablaConsultaNotaEspecial()

    'Por temor a que demore la consulta aumentamos el QueryTimeOut
    cBase.QueryTimeout = 120
    cTotalContado = 0
    cTotalIva = 0
    cIvaNota = 0
    cTotalNota = 0
    
    If cSucursal.ItemData(cSucursal.ListIndex) <> 0 Then
        CargoTablaNotaDevolucion cSucursal.ItemData(cSucursal.ListIndex), cSucursal.Text, cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text, TipoDocumento.NotaEspecial
        InsertoSubTotalSucursal cSucursal.Text, True
    Else
        For I = 0 To cSucursal.ListCount - 1
            If cSucursal.ItemData(I) <> 0 Then
                CargoTablaNotaDevolucion cSucursal.ItemData(I), cSucursal.List(I), cMoneda.ItemData(cMoneda.ListIndex), cMoneda.Text, TipoDocumento.NotaEspecial
                InsertoSubTotalSucursal cSucursal.List(I), True
            End If
        Next I
        TotalMonedaSucursal cMoneda.Text
    End If
    cBase.QueryTimeout = 15

End Sub

Private Sub TotalMonedaSucursal(strSignoMoneda As String)
Dim cSubTotal As Currency
Dim cIvaSubTotal  As Currency

    cSubTotal = cTotalContado - cTotalNota
    cIvaSubTotal = cTotalIva - cIvaNota
    
    If cSubTotal <> 0 Then
        vsPrinter = vbNullString
        strFormato = "+4400|+>400|+>1450|+>1400|+>1400|+>1450"
        strEncabezado = ""
        If cSubTotal <> 0 Then
            strTexto = "Total General" & "|" & Trim(strSignoMoneda) & "|" & Format(cSubTotal, FormatoMonedaP) & "|" & Format(cIvaSubTotal, FormatoMonedaP) & "|" & Format(cSubTotal + cIvaSubTotal, FormatoMonedaP) & "|" & Format(cSubTotal + cIvaSubTotal, FormatoMonedaP)
        Else
            strTexto = "Total General" & "|" & Trim(strSignoMoneda) & "|" & Format(Abs(cSubTotal), FormatoMonedaP) & "|" & Format(Abs(cIvaSubTotal), FormatoMonedaP) & "|" & Format(Abs(cSubTotal + cIvaSubTotal), FormatoMonedaP) & "|" & Format(Abs(cSubTotal + cIvaSubTotal), FormatoMonedaP)
        End If
        vsPrinter.TextColor = vbWhite
        'vsPrinter.FontSize = 10
        vsPrinter.AddTable strFormato, strEncabezado, strTexto, 0, vbBlue, True
        vsPrinter.TextColor = vbBlack
    End If
    
End Sub
Private Sub TotalMonedaSucursalCredito(strSignoMoneda As String)
Dim cSubTotal As Currency
Dim cIvaSubTotal  As Currency

    cSubTotal = cTotalContado - cTotalNota
    cIvaSubTotal = cTotalIva - cIvaNota

    If cSubTotal <> 0 Then
        strFormato = "+1800|+<2500|+>600|+>1400|+>1400|+>1400|+>1400"
        strEncabezado = ""
        If cSubTotal > 0 Then
            strTexto = "TOTALES" & "||" & Trim(strSignoMoneda) & "||" & Format(cIvaSubTotal, FormatoMonedaP) & "|" & Format(cSubTotal, FormatoMonedaP) & "|" & Format(cSubTotal + cIvaSubTotal, FormatoMonedaP)
        Else
            strTexto = "TOTALES" & "||" & Trim(strSignoMoneda) & "||" & Format(Abs(cIvaSubTotal), FormatoMonedaP) & "|" & Format(Abs(cSubTotal), FormatoMonedaP) & "|" & Format(Abs(cSubTotal + cIvaSubTotal), FormatoMonedaP)
        End If
        vsPrinter.AddTable strFormato, strEncabezado, strTexto, 0, vbYellow, True
    End If
    
End Sub

Private Sub CargoTablaNotaCredito(lnSucursal As Long, strSucursal As String, iMoneda As Integer, strSignoMoneda As String)
Dim lnAnterior As Long
Dim cTotal As Currency
Dim cSumaIva As Currency
Dim cSumaSubTotal As Currency
Dim sAnulado As Boolean
Dim cIva As Currency

    cParcialNota = 0
    cParcialIvaNota = 0
    
    'Consulta para inicializar la barra-----------------------------------------------------------------
    aValor = 0
    Cons = ArmoConsulta(lnSucursal, iMoneda, TipoDocumento.NotaCredito, 2)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then aValor = RsAux(0)
    RsAux.Close
    If aValor <> 0 Then
        fBar.Valor = 0
        fBar.ValorMaximo = aValor
        fBar.ValorMinimo = 0
        fBar.Update 0
        fBar.Show
        Me.Refresh
    Else
        Exit Sub
    End If
    '-----------------------------------------------------------------------------------------------------
    
    Cons = ArmoConsulta(lnSucursal, iMoneda, TipoDocumento.NotaCredito)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        NombreSucursal strSucursal, strSignoMoneda, DocNCredito
        ArmoEncabezadoCredito
    End If
    
    cSumaIva = 0
    cSumaSubTotal = 0
    lnAnterior = 0
    
    Do While Not RsAux.EOF
        
        If lnAnterior <> RsAux!DocCodigo Then
            strTexto = Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
            fBar.Update fBar.Valor + 1
        Else
            strTexto = ""
        End If

        lnAnterior = RsAux!DocCodigo
        If Not RsAux!DocAnulado Then
            cTotal = Format(RsAux!DocTotal, "#,##0.00")
            cIva = Format(RsAux!DocIva, "#,##0.00")
            'OJO FALTA CARGAR EL TOTAL DE LA FACTURA
            strTexto = strTexto & CargoRenglonCredito
        Else
            If Not sAnulado And strTexto <> vbNullString Then
                sAnulado = True
                strTexto = strTexto & "|| Anulado "
            Else
                sAnulado = True
            End If
        End If
        RsAux.MoveNext
        If Not sAnulado Then
            If RsAux.EOF Then
                strTexto = strTexto & "|" & Format(cIva, "#,##0.00") & "|" & Format(cTotal - cIva, "#,##0.00") & "|" & Format(cTotal, "#,##0.00")
                cSumaIva = CCur(Format(cSumaIva + cIva, "#,##0.00"))
                cSumaSubTotal = CCur(Format(cSumaSubTotal + (cTotal - cIva), "#,##0.00"))
            Else
                If lnAnterior = RsAux!DocCodigo Then
                    strTexto = strTexto & "| "
                Else
                    strTexto = strTexto & "|" & Format(cIva, "#,##0.00") & "|" & Format(cTotal - cIva, "#,##0.00") & "|" & Format(cTotal, "#,##0.00")
                    cSumaIva = CCur(Format(cSumaIva + cIva, "#,##0.00"))
                    cSumaSubTotal = CCur(Format(cSumaSubTotal + (cTotal - cIva), "#,##0.00"))
                End If
            End If
            vsPrinter.AddTable strFormato, strEncabezado, strTexto, 0, 0, True
        Else
            If strTexto <> vbNullString Then vsPrinter.AddTable strFormato, strEncabezado, strTexto, 0, &HE0E0E0, True
            sAnulado = False
        End If
    Loop
    RsAux.Close
    
    If cSumaSubTotal > 0 Then
        InsertoTotalSucursal DocNCredito, strSignoMoneda, cSumaIva, cSumaSubTotal
        cParcialNota = cSumaSubTotal
        cParcialIvaNota = cSumaIva
        cTotalNota = cTotalNota + cSumaSubTotal
        cIvaNota = cIvaNota + cSumaIva
    End If

End Sub
Private Sub InsertoTotalSucursal(strDocumento As String, strSignoMoneda As String, cSumaIva As Currency, cSumaSubTotal As Currency, Optional Contado As Boolean = False)
    
    strEncabezado = ""
    If Contado Then
        strFormato = "+2800|+<1450|+>550|+>1450|+>1400|+>1400|+>1450"
        strTexto = "Total " & strDocumento & "||" & strSignoMoneda & "|" & Format(cSumaSubTotal, FormatoMonedaP) & "|" & Format(cSumaIva, FormatoMonedaP) & "|" & Format(cSumaSubTotal + cSumaIva, FormatoMonedaP) & "|" & Format(cSumaSubTotal + cSumaIva, FormatoMonedaP) & "|" & Format(cSumaSubTotal + cSumaIva, FormatoMonedaP)
    Else
        strFormato = "+2800|+<1500|+>600|+>1400|+>1400|+>1400|+>1400"
        strTexto = "Total " & strDocumento & "||" & strSignoMoneda & "||" & Format(cSumaIva, FormatoMonedaP) & "|" & Format(cSumaSubTotal, FormatoMonedaP) & "|" & Format(cSumaSubTotal + cSumaIva, FormatoMonedaP) & "|" & Format(cSumaSubTotal + cSumaIva, FormatoMonedaP)
    End If
    vsPrinter.AddTable strFormato, strEncabezado, strTexto, 0, vbYellow, True
    vsPrinter = vbNullString
    
End Sub
Private Sub InsertoSubTotalSucursal(strSucursal As String, Optional Contado As Boolean = False)
Dim cSubTotal As Currency
Dim cIvaSubTotal  As Currency

    cSubTotal = cParcialNota
    If cParcialNota = 0 Then Exit Sub
    cIvaSubTotal = cParcialIvaNota
    strEncabezado = ""
    
    If Not Contado Then
        strFormato = "+2800|+<1500|+>600|+>1400|+>1400|+>1400|+>1400"
        If cSubTotal > 0 Then
            strTexto = "Total " & strSucursal & "||" & cMoneda.Text & "||" & Format(cIvaSubTotal, FormatoMonedaP) & "|" & Format(cSubTotal, FormatoMonedaP) & "|" & Format(cSubTotal + cIvaSubTotal, FormatoMonedaP) & "|" & Format(cSubTotal + cIvaSubTotal, FormatoMonedaP)
        Else
            strTexto = "Total " & strSucursal & "||" & cMoneda.Text & "||" & Format(cIvaSubTotal, FormatoMonedaP) & "|" & Format(Abs(cSubTotal), FormatoMonedaP) & "|" & Format(Abs(cSubTotal + cIvaSubTotal), FormatoMonedaP) & "|" & Format(Abs(cSubTotal + cIvaSubTotal), FormatoMonedaP)
        End If
    Else
        strFormato = "+2800|+<1400|+>600|+>1450|+>1400|+>1400|+>1450"
        If cSubTotal > 0 Then
            strTexto = "Total " & strSucursal & "||" & cMoneda.Text & "|" & Format(cSubTotal, FormatoMonedaP) & "|" & Format(cIvaSubTotal, FormatoMonedaP) & "|" & Format(cSubTotal + cIvaSubTotal, FormatoMonedaP) & "|" & Format(cSubTotal + cIvaSubTotal, FormatoMonedaP) & "|" & Format(cSubTotal + cIvaSubTotal, FormatoMonedaP)
        Else
            strTexto = "Total " & strSucursal & "||" & cMoneda.Text & "|" & Format(Abs(cSubTotal), FormatoMonedaP) & "|" & Format(Abs(cIvaSubTotal), FormatoMonedaP) & "|" & Format(Abs(cSubTotal + cIvaSubTotal), FormatoMonedaP) & "|" & Format(Abs(cSubTotal + cIvaSubTotal), FormatoMonedaP) & "|" & Format(Abs(cSubTotal + cIvaSubTotal), FormatoMonedaP)
        End If
    End If
    vsPrinter.AddTable strFormato, strEncabezado, strTexto, 0, vbYellow, True
    vsPrinter = ""
    
End Sub
Private Sub CargoTablaNotaDevolucion(lnSucursal As Long, strSucursal As String, iMoneda As Integer, strSignoMoneda As String, lnTipoDocumento As Long)

Dim lnAnterior As Long
Dim cTotal As Currency
Dim cSumaIva As Currency
Dim cSumaContado As Currency
Dim sAnulado As Boolean
    
    cParcialNota = 0
    cParcialIvaNota = 0

    'Consulta para inicializar la barra-----------------------------------------------------------------
    aValor = 0
    Cons = ArmoConsulta(lnSucursal, iMoneda, lnTipoDocumento, 2)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then aValor = RsAux(0)
    RsAux.Close
    If aValor <> 0 Then
        fBar.Valor = 0
        fBar.ValorMaximo = aValor
        fBar.ValorMinimo = 0
        fBar.Update 0
        fBar.Show
        Me.Refresh
    Else
        Exit Sub
    End If
    '-----------------------------------------------------------------------------------------------------
    
    Cons = ArmoConsulta(lnSucursal, iMoneda, lnTipoDocumento)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    If Not RsAux.EOF Then
        If lnTipoDocumento = TipoDocumento.NotaDevolucion Then
            NombreSucursal strSucursal, strSignoMoneda, DocNDevolucion
        Else
            NombreSucursal strSucursal, strSignoMoneda, DocNEspecial
        End If
        ArmoEncabezadoContado
    End If
    cSumaIva = 0
    cSumaContado = 0
    lnAnterior = 0
    sAnulado = False
    
    Do While Not RsAux.EOF
        
        If lnAnterior <> RsAux!DocCodigo Then
            strTexto = Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
            fBar.Update fBar.Valor + 1
        Else
            strTexto = ""
        End If
        lnAnterior = RsAux!DocCodigo
        If Not RsAux!DocAnulado Then
            cTotal = RsAux!DocTotal
        
            'OJO FALTA CARGAR EL TOTAL DE LA FACTURA
            strTexto = strTexto & CargoRenglonContado
        
            'Sumo los totales
            cSumaIva = CCur(Format(cSumaIva + (RsAux!RenIva * RsAux!RenCantidad), "#,##0.00"))
            cSumaContado = CCur(Format(cSumaContado + ((RsAux!RenPrecio - RsAux!RenIva) * RsAux!RenCantidad), "#,##0.00"))
        Else
            If Not sAnulado And strTexto <> vbNullString Then
                sAnulado = True
                strTexto = strTexto & "|| Anulado "
            Else
                sAnulado = True
            End If
        End If
        RsAux.MoveNext
        If Not sAnulado Then
            If RsAux.EOF Then
                strTexto = strTexto & "|" & Format(cTotal, "#,##0.00")
            Else
                If lnAnterior = RsAux!DocCodigo Then
                    strTexto = strTexto & "| "
                Else
                    strTexto = strTexto & "|" & Format(cTotal, "#,##0.00")
                End If
            End If
            vsPrinter.AddTable strFormato, strEncabezado, strTexto, 0, 0, True
        Else
            If strTexto <> vbNullString Then vsPrinter.AddTable strFormato, strEncabezado, strTexto, 0, &HE0E0E0, True
            sAnulado = False
        End If
    Loop
    RsAux.Close
    If cSumaContado > 0 Then
        InsertoTotalSucursal DocNDevolucion, strSignoMoneda, cSumaIva, cSumaContado, True
        cParcialNota = cSumaContado
        cParcialIvaNota = cSumaIva
        cTotalNota = cTotalNota + cSumaContado
        cIvaNota = cIvaNota + cSumaIva
    End If

End Sub


Private Function CargoRenglonCredito() As String
    CargoRenglonCredito = "|" & Trim(RsAux!ArtCodigo) & "|" & FormateoString(vsPrinter, 2300, Trim(RsAux!ArtNombre)) _
        & "|" & Trim(RsAux!RenCantidad) & "|" & Format(RsAux!RenPrecio - RsAux!RenIva, "#,##0.00")
End Function
Private Function CargoRenglonContado() As String
    
    CargoRenglonContado = "|" & Trim(RsAux!ArtCodigo) & "|" & FormateoString(vsPrinter, 2200, Trim(RsAux!ArtNombre)) _
        & "|" & Trim(RsAux!RenCantidad) & "|" & Format(RsAux!RenCantidad * (RsAux!RenPrecio - RsAux!RenIva), "#,##0.00") _
        & "|" & Format(RsAux!RenIva * RsAux!RenCantidad, "#,##0.00") & "|" & Format(RsAux!RenPrecio * RsAux!RenCantidad, "#,##0.00")
        
    'CargoRenglonContado = "|" & Trim(RsAux!ArtCodigo) & "|" & Trim(RsAux!ArtNombre) _
        & "|" & Trim(RsAux!RenCantidad) & "|" & Format(RsAux!RenCantidad * (RsAux!RenPrecio - RsAux!RenIva), "#,##0.00") _
        & "|" & Format(RsAux!RenIva * RsAux!RenCantidad, "#,##0.00") & "|" & Format(RsAux!RenPrecio * RsAux!RenCantidad, "#,##0.00")
        
        
End Function
Private Sub ArmoEncabezadoContado()

    strFormato = "+>950|+<850|+<2400|+>600|+>1450|+>1350|+>1450|+>1450"
    strEncabezado = "Factura|Código|Nombre|Cant|Contado|Iva|SubTotal|Total"
    
    With vsPrinter
        .FontSize = 9
        .FontBold = True
        .TextColor = vbWhite
        .AddTable strFormato, strEncabezado, "", vbBlue
    End With

    vsPrinter.FontSize = 9
    vsPrinter.FontBold = False
    vsPrinter.FontItalic = False
    vsPrinter.TextColor = vbBlack
        
End Sub
Private Sub ArmoEncabezadoCredito()

    strFormato = "+>950|+<850|+<2500|+>600|+>1400|+>1400|+>1400|+>1400"
    strEncabezado = "Factura|Código|Nombre|Cant|Unitario S/I|Iva|SubTotal|Total"
    
    vsPrinter.FontSize = 9
    vsPrinter.FontBold = True
    vsPrinter.TextColor = vbWhite
    With vsPrinter
        .AddTable strFormato, strEncabezado, "", vbBlue
    End With
    
    vsPrinter.FontSize = 9
    vsPrinter.FontBold = False
    vsPrinter.FontItalic = False
    vsPrinter.TextColor = vbBlack

End Sub

'------------------------------------------------------------------------------------------------------------------------------
'   El Tipo indica la clase de consulta a Armar
'       Tipo = 1 devuelve la consulta con filas de dicumentos
'       Tipo = 2 devuelve la consulta con Count (cantidad de documentos)
'------------------------------------------------------------------------------------------------------------------------------
Private Function ArmoConsulta(lnSucursal As Long, iMoneda As Integer, lnTipoDocumento As Long, Optional Tipo As Integer = 1) As String

    Select Case Tipo
        Case 1  'Consulta de datos
            Cons = "Select * From Documento, Renglon, Articulo" _
                    & " Where DocTipo = " & lnTipoDocumento _
                    & " And DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tFecha.Text, "mm/dd/yyyy 23:59:59") & "'" _
        
            If lnSucursal > 0 Then Cons = Cons & " And DocSucursal = " & lnSucursal
            If iMoneda > 0 Then Cons = Cons & " And DocMoneda = " & iMoneda
            
            Cons = Cons & " And DocCodigo = RenDocumento And RenArticulo = ArtId"
            
        Case 2 'Consulta de cantidad de operaciones (para inicializar barra)
            Cons = "Select Count(*) From Documento" _
                        & " Where DocTipo = " & lnTipoDocumento _
                        & " And DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tFecha.Text, "mm/dd/yyyy 23:59:59") & "'" _
            
            If lnSucursal > 0 Then Cons = Cons & " And DocSucursal = " & lnSucursal
            If iMoneda > 0 Then Cons = Cons & " And DocMoneda = " & iMoneda
            
    End Select
    
    ArmoConsulta = Cons
    
End Function

Private Sub cDocumento_Change()
    Selecciono cDocumento, cDocumento.Text, gTecla
End Sub

Private Sub cDocumento_GotFocus()
    cDocumento.SelStart = 0
    cDocumento.SelLength = Len(cDocumento.Text)
End Sub

Private Sub cDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cDocumento.ListIndex
End Sub

Private Sub cDocumento_KeyPress(KeyAscii As Integer)
    cDocumento.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then cMoneda.SetFocus
End Sub

Private Sub cDocumento_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cDocumento
End Sub

Private Sub cDocumento_LostFocus()
    gIndice = -1
    cDocumento.SelLength = 0
End Sub

Private Sub cSucursal_Change()
    Selecciono cSucursal, cSucursal.Text, gTecla
End Sub
Private Sub cSucursal_GotFocus()
    cSucursal.SelStart = 0
    cSucursal.SelLength = Len(cSucursal.Text)
End Sub

Private Sub cSucursal_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cSucursal.ListIndex
End Sub

Private Sub cSucursal_KeyPress(KeyAscii As Integer)
    cSucursal.ListIndex = gIndice
    If KeyAscii = vbKeyReturn And cSucursal.ListIndex > -1 Then tFecha.SetFocus
End Sub

Private Sub cSucursal_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cSucursal
End Sub

Private Sub cSucursal_LostFocus()
    gIndice = -1
    cSucursal.SelLength = 0
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
    CargoSucursal
    CargoDocumento
    CargoMoneda
    CargoDatosZoom cZoom
    cZoom.ListIndex = 7
    
    'Cargo la fecha del servidor.
    FechaDelServidor
    tFecha.Text = Format(gFechaServidor, "d-Mmm-yyyy")
    
    Exit Sub
    
ErrLoad:
    Screen.MousePointer = vbDefault
    ErrorInesperado Err.Description
End Sub
Private Sub CargoSucursal()

    Cons = "Select SucCodigo, SucAbreviacion " _
        & " From Sucursal " _
        & " Where SucDcontado <> null Or SucDCredito <> Null"
        
    CargoCombo Cons, cSucursal, ""
    
    cSucursal.AddItem "Todos"
    cSucursal.ItemData(cSucursal.NewIndex) = 0
    
    'Por defecto pongo todos.
    For I = 0 To cSucursal.ListCount - 1
        If cSucursal.ItemData(I) = 0 Then cSucursal.ListIndex = I: Exit For
    Next
    
End Sub

Private Sub CreoTitulo()

    If cDocumento.ItemData(cDocumento.ListIndex) = TipoDocumento.NotaDevolucion Then
        EncabezadoListado vsPrinter, "Diario de Devoluciones Contado al " & Format(tFecha.Text, "d/mm/yy"), True
    ElseIf cDocumento.ItemData(cDocumento.ListIndex) = TipoDocumento.NotaCredito Then
        EncabezadoListado vsPrinter, "Diario de Devoluciones Crédito al " & Format(tFecha.Text, "d/mm/yy"), True
    ElseIf cDocumento.ItemData(cDocumento.ListIndex) = TipoDocumento.NotaEspecial Then
        EncabezadoListado vsPrinter, "Diario de Devoluciones Especiales al " & Format(tFecha.Text, "d/mm/yy"), True
    Else
        EncabezadoListado vsPrinter, "Diario de Devoluciones al " & Format(tFecha.Text, "d/mm/yy"), True
    End If

End Sub

Private Sub CargoDocumento()

    cDocumento.Clear
    cDocumento.AddItem DocNDevolucion
    cDocumento.ItemData(cDocumento.NewIndex) = TipoDocumento.NotaDevolucion
    
    cDocumento.AddItem DocNCredito
    cDocumento.ItemData(cDocumento.NewIndex) = TipoDocumento.NotaCredito
    
    cDocumento.AddItem DocNEspecial
    cDocumento.ItemData(cDocumento.NewIndex) = TipoDocumento.NotaEspecial
    
    cDocumento.AddItem "Todos"
    cDocumento.ItemData(cDocumento.NewIndex) = 0
    
End Sub

Private Function ValidoDatos() As Boolean

    If cSucursal.ListIndex = -1 Then
        MsgBox "No selecciono una sucursal válida.", vbExclamation, "ATENCIÓN"
        cSucursal.SetFocus
        ValidoDatos = False
        Exit Function
    End If
    If Not IsDate(tFecha.Text) Then
        MsgBox "No ingreso una fecha válida, verifique.", vbExclamation, "ATENCIÓN"
        tFecha.SetFocus
        ValidoDatos = False
        Exit Function
    End If
    If cDocumento.ListIndex = -1 Then
        MsgBox "No selecciono un tipo de documento válido.", vbExclamation, "ATENCIÓN"
        cDocumento.SetFocus
        ValidoDatos = False
        Exit Function
    End If
    If cMoneda.ListIndex = -1 Then
        MsgBox "Seleccione una moneda o Todas.", vbExclamation, "ATENCIÓN"
        cMoneda.SetFocus
        ValidoDatos = False
        Exit Function
    End If
    ValidoDatos = True
End Function

Private Sub NombreSucursal(strSucursal As String, strSignoMoneda As String, strDocumento As String)
    
    With vsPrinter
        .FontSize = 10
        .FontBold = True
        .Text = " Sucursal : "
        .FontItalic = True
        .Text = strSucursal
        .FontItalic = False
        .TextAlign = taCenterTop
        .Text = strDocumento
        .TextAlign = taRightTop
        .Text = " Moneda: " & strSignoMoneda
    End With
    vsPrinter = ""
    vsPrinter.TextAlign = 0
    
End Sub

Private Sub Form_Resize()
    AjustoObjetos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Forms(Forms.Count - 2).SetFocus
End Sub

Private Sub Label1_Click()
    Foco cSucursal
End Sub

Private Sub Label2_Click()
    Foco tFecha
End Sub

Private Sub Label3_Click()
    Foco cDocumento
End Sub

Private Sub Label4_Click()
    Foco cMoneda
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

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0
    tFecha.SelLength = Len(tFecha.Text)
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cDocumento.SetFocus
End Sub

Private Sub tFecha_LostFocus()
    If Not IsDate(tFecha.Text) Then
        MsgBox "No se ingreso una fecha válida.", vbExclamation, "ATENCIÓN"
    Else
        tFecha.Text = Format(tFecha.Text, "d-Mmm-yyyy")
    End If
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
    BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
    Exit Sub

ErrCM:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al cargar las monedas."
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

End Sub

Private Sub cMoneda_KeyUp(KeyCode As Integer, Shift As Integer)

    ComboKeyUp cMoneda

End Sub

Private Sub cMoneda_LostFocus()
    gIndice = -1
    cMoneda.SelLength = 0
End Sub

Private Sub TotalSucursal(cSumaContado As Currency, cSumaIva As Currency, strSignoMoneda As String)

    'strFormato = "+>950|+>850|+<1900|+>500|+>1200|+>1300|+>1100|+>1300|+>1400"
    strFormato = "+4400|+>400|+>1450|+>1400|+>1500|+>1350"
    strEncabezado = ""
    strTexto = "  TOTAL SUCURSAL" & "|" & Trim(strSignoMoneda) & "|" & Format(cSumaContado, FormatoMonedaP) & "|" & Format(cSumaIva, FormatoMonedaP) & "|" & Format(cSumaContado + cSumaIva, FormatoMonedaP) & "|" & Format(cSumaContado + cSumaIva, FormatoMonedaP)
    vsPrinter.AddTable strFormato, strEncabezado, strTexto, 0, vbYellow, True

End Sub

Private Sub AjustoObjetos()

    If Me.WindowState = vbMinimized Then Exit Sub
    vsPrinter.Left = 80
    vsPrinter.Width = Me.Width - 240
    Toolbar1.Width = vsPrinter.Width
    Shape1.Width = vsPrinter.Width
    vsPrinter.Height = Me.Height - 1950
    
End Sub

Private Sub AccionImprimir()

    On Error GoTo errPrint
    vsPrinter.filename = "Diario de Devoluciones"
    vsPrinter.PrintDoc
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al impirmir. " & Trim(Err.Description)
End Sub

Private Sub vsPrinter_NewPage()

    If vsPrinter.CurrentPage > 1 Then
        Select Case cDocumento.ItemData(cDocumento.ListIndex)
            Case TipoDocumento.Contado: ArmoEncabezadoContado
            Case TipoDocumento.Credito: ArmoEncabezadoCredito
        End Select
    End If
    
End Sub
