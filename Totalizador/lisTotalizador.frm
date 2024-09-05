VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Begin VB.Form LisTotalizador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Totalizador de Operaciones"
   ClientHeight    =   5835
   ClientLeft      =   2535
   ClientTop       =   2670
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
   Icon            =   "lisTotalizador.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   9165
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   0
      TabIndex        =   6
      Top             =   660
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
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
         TabIndex        =   7
         Top             =   30
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de consulta"
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   30
      Width           =   9015
      Begin VB.ComboBox cMonedaT 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   230
         Width           =   855
      End
      Begin VB.ComboBox cMonedaN 
         Height          =   315
         Left            =   6120
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   230
         Width           =   855
      End
      Begin VB.CommandButton bConsultar 
         Caption         =   "&Consultar"
         Height          =   340
         Left            =   7800
         TabIndex        =   4
         Top             =   210
         Width           =   975
      End
      Begin VB.CheckBox cConvertir 
         Caption         =   "Convertir moneda extranjera a..."
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Totalizar operaciones en:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   260
         Width           =   1935
      End
   End
   Begin vsViewLib.vsPrinter vsPrinter 
      Height          =   4575
      Left            =   120
      TabIndex        =   5
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
      Zoom            =   70
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
            Picture         =   "lisTotalizador.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisTotalizador.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisTotalizador.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisTotalizador.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisTotalizador.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisTotalizador.frx":13C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisTotalizador.frx":1C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisTotalizador.frx":1FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisTotalizador.frx":22D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "lisTotalizador.frx":25EC
            Key             =   ""
         EndProperty
      EndProperty
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
      Begin VB.Menu MnuSL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCambiarBase 
         Caption         =   "Cambiar BD"
      End
   End
End
Attribute VB_Name = "LisTotalizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFormato As String
Dim aCantidadT As Long, aImporteT As Currency
Dim aTC As Currency

Private Sub AccionConsultar()

    On Error GoTo ErrBC
    If Not ValidoDatos Then Screen.MousePointer = vbDefault: Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    aCantidadT = 0: aImporteT = 0
    vsPrinter.StartDoc
    If vsPrinter.Error Then MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault: Exit Sub
    
    aTexto = BuscoNombreMoneda(cMonedaT.ItemData(cMonedaT.ListIndex))
    EncabezadoListado vsPrinter, "Totalizador de Operaciones (" & aTexto & ")", True
    
    cBase.QueryTimeout = 90
    
    'TOTALIZO CGSA----------------------------------------------------------------------------------------------------
    vsPrinter.FontSize = 8
    vsPrinter = ""
    vsPrinter.Paragraph = "Operaciones Totalizadas al: " & Format(Date, "Ddd d Mmm yyyy") & " - CGSA"
    'Veo si hay que sacar la TC a la moneda nacional
    If cMonedaN.Enabled Then
        aTC = TasadeCambio(CLng(cMonedaT.ItemData(cMonedaT.ListIndex)), CLng(cMonedaN.ItemData(cMonedaN.ListIndex)), Date)
        vsPrinter.Paragraph = "Tasa de Cambio " & cMonedaT.Text & " -> " & cMonedaN.Text & " al " & Format(Date, "d Mmm yyyy") & ": " & Format(aTC, "#,##0.000")
    End If
    
    'Tipos de Operaciones   a) Normales y En Gestor     b) A Perdida
    CargoCreditosNormalesYGestor
    CargoCreditosAPerdida
    
    'Total General de las operaciones
    vsPrinter = "": vsPrinter = ""
    vsPrinter.DrawLine 1800, vsPrinter.CurrentY, 8200, vsPrinter.CurrentY
    vsPrinter = ""
    CargoTabla "Total General", aCantidadT, aImporteT, True
    
    'TOTALIZO MEGA----------------------------------------------------------------------------------------------------
     aCantidadT = 0: aImporteT = 0
    vsPrinter.NewPage: vsPrinter = "": vsPrinter.FontSize = 8
    vsPrinter.Paragraph = "Operaciones Totalizadas al: " & Format(Date, "Ddd d Mmm yyyy") & " - MEGA"
    'Veo si hay que sacar la TC a la moneda nacional
    If cMonedaN.Enabled Then
        aTC = TasadeCambio(CLng(cMonedaT.ItemData(cMonedaT.ListIndex)), CLng(cMonedaN.ItemData(cMonedaN.ListIndex)), Date)
        vsPrinter.Paragraph = "Tasa de Cambio " & cMonedaT.Text & " -> " & cMonedaN.Text & " al " & Format(Date, "d Mmm yyyy") & ": " & Format(aTC, "#,##0.000")
    End If
    
    'Tipos de Operaciones   a) Normales y En Gestor     b) A Perdida
    CargoCreditosNormalesYGestor False
    CargoCreditosAPerdida False
    
    'Total General de las operaciones
    vsPrinter = "": vsPrinter = ""
    vsPrinter.DrawLine 1800, vsPrinter.CurrentY, 8200, vsPrinter.CurrentY
    vsPrinter = ""
    CargoTabla "Total General", aCantidadT, aImporteT, True
        
    vsPrinter.EndDoc
    
    cBase.QueryTimeout = 15
    Screen.MousePointer = vbDefault
    Exit Sub

ErrBC:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al consultar. " & Trim(Err.Description)
    cBase.QueryTimeout = 15
End Sub

Private Sub PropiedadesImpresion()

  With vsPrinter
        .PreviewPage = 1
        .FontName = "Tahoma"
        .PaperSize = 1      'Letter
        .FontSize = 9: .FontBold = False: .FontItalic = False
        .TextAlign = 0  'Left
        .PageBorder = 3 ' pbTop
        .PenStyle = 0: .BrushStyle = 0: .PenWidth = 2: .PenColor = 0
        .BrushColor = 0
        .TextColor = 0: .Columns = 1
        .Zoom = 100
        .TableBorder = tbBoxRows
        .MarginRight = 400
    End With

End Sub

Private Sub CargoCreditosNormalesYGestor(Optional CGSA As Boolean = True)

Dim aCantidad As Long, aImporte As Currency
Dim aCantidadP As Long, aImporteP As Currency
    
    aCantidadP = 0: aImporteP = 0
    vsPrinter = ""
    'Creditos Normales---------------------------------------------------------------------------------------------
    Cons = ArmoConsulta(TipoCredito.Normal, CLng(cMonedaT.ItemData(cMonedaT.ListIndex)), CGSA)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    aCantidad = 0: aImporte = 0
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aCantidad = RsAux!Cantidad
        If Not IsNull(RsAux!Importe) Then aImporte = RsAux!Importe
    End If
    RsAux.Close
    
    CargoTabla "Operaciones Normales", aCantidad, aImporte
    aImporteP = aImporteP + aImporte
    aCantidadP = aCantidadP + aCantidad
    '------------------------------------------------------------------------------------------------------------------
     vsPrinter = ""
     
    'Creditos en Gestor---------------------------------------------------------------------------------------------
    Cons = ArmoConsulta(TipoCredito.Gestor, CLng(cMonedaT.ItemData(cMonedaT.ListIndex)), CGSA)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    aCantidad = 0: aImporte = 0
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aCantidad = RsAux!Cantidad
        If Not IsNull(RsAux!Importe) Then aImporte = RsAux!Importe
    End If
    RsAux.Close
    
    CargoTabla "Operaciones en Gestor", aCantidad, aImporte
    aImporteP = aImporteP + aImporte
    aCantidadP = aCantidadP + aCantidad
    '------------------------------------------------------------------------------------------------------------------
    
    'Total de las operaciones
     vsPrinter = ""
     vsPrinter.DrawLine 1800, vsPrinter.CurrentY, 8200, vsPrinter.CurrentY
     vsPrinter = ""
     CargoTabla "Total Normales y Gestor", aCantidadP, aImporteP, True
     
    aImporteT = aImporteT + aImporteP
    aCantidadT = aCantidadT + aCantidadP
        
End Sub

Private Sub CargoCreditosAPerdida(Optional CGSA As Boolean = True)

Dim aCantidad As Long, aImporte As Currency
    
    vsPrinter = "": vsPrinter = ""
    
    'Creditos A Perdida---------------------------------------------------------------------------------------------
    Cons = ArmoConsulta(TipoCredito.Incobrable, CLng(cMonedaT.ItemData(cMonedaT.ListIndex)), CGSA)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    aCantidad = 0: aImporte = 0
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aCantidad = RsAux!Cantidad
        If Not IsNull(RsAux!Importe) Then aImporte = RsAux!Importe
    End If
    RsAux.Close
    
    CargoTabla "Operaciones A Pérdida", aCantidad, aImporte
    
    aImporteT = aImporteT + aImporte
    aCantidadT = aCantidadT + aCantidad
    
End Sub

Private Sub CargoTabla(Titulo As String, Cantidad As Long, Importe As Currency, Optional EsTotal As Boolean = False)

    vsPrinter.FontSize = 9
    vsPrinter.MarginLeft = 2000
    
    aTexto = Titulo
    strFormato = "+<5000"
    If EsTotal Then strFormato = strFormato & "|1000"
    
    vsPrinter.FontBold = True
    If EsTotal Then vsPrinter.TextColor = vbWhite Else vsPrinter.TextColor = vbBlack
    If EsTotal Then vsPrinter.AddTable strFormato, aTexto, "", vbBlue Else vsPrinter.AddTable strFormato, aTexto, "", Gris
    vsPrinter.FontBold = False
    vsPrinter.TextColor = vbBlack

    If Not cMonedaN.Enabled Then strFormato = "+<3000|+>2000" Else strFormato = "+<1000|+>2000|+>2000"
    If EsTotal Then strFormato = strFormato & "|1000"
    
    'Cantidad de Operaciones
    aTexto = "Cantidad|" & Format(Cantidad, "#,##0")
    vsPrinter.AddTable strFormato, "", aTexto, , , True
    
    'Importe que deben las operaciones
    aTexto = "Saldo|" & Trim(cMonedaT.Text) & " " & Format(Importe, FormatoMonedaP)
       
    If cMonedaN.Enabled Then
        aTexto = aTexto & "|" & Trim(cMonedaN.Text) & " " & Format(Importe * aTC, FormatoMonedaP)
    End If
    
    vsPrinter.AddTable strFormato, "", aTexto, , , True
    
    vsPrinter.MarginLeft = 720
    
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub


Private Sub cConvertir_Click()
    
    If cConvertir.Value = vbChecked Then
        cMonedaN.Enabled = True
        cMonedaN.BackColor = Obligatorio
        BuscoCodigoEnCombo cMonedaN, paMonedaFacturacion
    Else
        cMonedaN.Enabled = False
        cMonedaN.BackColor = Inactivo
        cMonedaN.Text = ""
    End If
    
End Sub

Private Sub cConvertir_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If cMonedaN.Enabled Then Foco cMonedaN Else bConsultar.SetFocus
End Sub

Private Sub cMonedaN_Change()
    Selecciono cMonedaN, cMonedaN.Text, gTecla
End Sub
Private Sub cMonedaN_GotFocus()
    cMonedaN.SelStart = 0
    cMonedaN.SelLength = Len(cMonedaN.Text)
End Sub

Private Sub cMonedaN_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cMonedaN.ListIndex
End Sub

Private Sub cMonedaN_KeyPress(KeyAscii As Integer)
    cMonedaN.ListIndex = gIndice
    If KeyAscii = vbKeyReturn And cMonedaN.ListIndex > -1 Then bConsultar.SetFocus
End Sub

Private Sub cMonedaN_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cMonedaN
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

    CargoMoneda
    
    cConvertir.Value = vbUnchecked
    cMonedaN.Enabled = False: cMonedaN.BackColor = Inactivo
    
    CargoDatosZoom cZoom
    cZoom.ListIndex = 6
       
    PropiedadesImpresion
    
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValidoDatos() As Boolean

    ValidoDatos = False
    
    If cMonedaT.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda para totalizar las operaciones.", vbExclamation, "ATENCIÓN"
        Foco cMonedaT: Exit Function
    End If
    
    If cConvertir.Value = vbChecked And cMonedaN.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda nacional para buscar la tasa de cambio.", vbExclamation, "ATENCIÓN"
        Foco cMonedaN: Exit Function
    End If
    
    ValidoDatos = True
    
End Function

Private Sub Form_Resize()
    AjustoObjetos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Forms(Forms.Count - 2).SetFocus
End Sub

Private Sub Label1_Click()
    Foco cMonedaT
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

Private Sub MnuCambiarBase_Click()
    Dim newB As String
    On Error GoTo errCh
    If MsgBox("Ud. sabe lo que está haciendo !!?", vbQuestion + vbYesNo + vbDefaultButton2, "Realmente desea cambiar la base") = vbNo Then Exit Sub
    
    newB = InputBox("Ingrese el texto del login para la nueva conexión" & vbCrLf & _
                "Id de aplicación en archivo de conexiones.", "Cambio de Base de Datos")
    
    If Trim(newB) = "" Then Exit Sub
    If MsgBox("Está seguro de cambiar la base de datos al login " & newB, vbQuestion + vbYesNo + vbDefaultButton2, "Cambiar Base") = vbNo Then Exit Sub
    
    newB = miConexion.TextoConexion(newB)
    If Trim(newB) = "" Then Exit Sub
    
    Screen.MousePointer = 11
    On Error Resume Next
    cBase.Close
    On Error GoTo errCh
    Set cBase = Nothing
    Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , newB)
    
    Screen.MousePointer = 0
    
    MsgBox "Ahora está trabajanbo en la nueva base de datos.", vbExclamation, "Base Cambiada OK"
    Exit Sub
    
errCh:
    clsGeneral.OcurrioError "Error de Conexión", Err.Description
    Screen.MousePointer = 0
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
    CargoCombo Cons, cMonedaT, ""
    CargoCombo Cons, cMonedaN, ""
    BuscoCodigoEnCombo cMonedaT, paMonedaFacturacion
    Exit Sub

ErrCM:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al cargar las monedas."
End Sub

Private Sub cMonedaT_Change()
    Selecciono cMonedaT, cMonedaT.Text, gTecla
End Sub

Private Sub cMonedaT_GotFocus()
    cMonedaT.SelStart = 0
    cMonedaT.SelLength = Len(cMonedaT.Text)
End Sub

Private Sub cMonedaT_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cMonedaT.ListIndex
End Sub

Private Sub cMonedaT_KeyPress(KeyAscii As Integer)

    cMonedaT.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then cConvertir.SetFocus

End Sub

Private Sub cMonedaT_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cMonedaT
End Sub

Private Sub cMonedaT_LostFocus()
    gIndice = -1
    cMonedaT.SelLength = 0
End Sub

Private Sub AjustoObjetos()

    If Me.WindowState = vbMinimized Then Exit Sub
    vsPrinter.Left = 80
    vsPrinter.Width = Me.Width - 240
    Toolbar1.Width = vsPrinter.Width
    Frame1.Width = vsPrinter.Width
    vsPrinter.Height = Me.Height - 1950
    
End Sub

Private Sub AccionImprimir()

    On Error GoTo errPrint
    vsPrinter.FileName = "Diario de Cobranza"
    vsPrinter.PrintDoc
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al impirmir. " & Trim(Err.Description)
End Sub

Private Sub vsPrinter_NewPage()
    'If vsPrinter.PageCount > 1 Then ArmoEncabezado
End Sub

Private Function ArmoConsulta(TipoDeCredito As Integer, Moneda As Long, Optional CGSA As Boolean = True) As String

    Cons = "Select Cantidad = Count(*), Importe = Sum(CreSaldoFactura) from Credito, Documento" _
           & " Where CreTipo = " & TipoDeCredito _
           & " And CreSaldoFactura > 0 " _
           & " And CreFactura = DocCodigo" _
           & " And DocMoneda = " & Moneda _
           & " And DocAnulado = 0"
    
    'O CGSA o MEGA
    If CGSA Then Cons = Cons & " And CreMega = 0" Else Cons = Cons & " And CreMega = 1"
    
    ArmoConsulta = Cons
    
End Function


