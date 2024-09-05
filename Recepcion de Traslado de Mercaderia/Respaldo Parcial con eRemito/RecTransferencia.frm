VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form RecTransferencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Traslado"
   ClientHeight    =   6210
   ClientLeft      =   30
   ClientTop       =   615
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RecTransferencia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6495
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "recepcion"
            Object.ToolTipText     =   "Recepcionar Mercadería"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprimir entrega"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   5100
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox tDoc 
      Height          =   285
      Left            =   3300
      MaxLength       =   8
      TabIndex        =   3
      Top             =   540
      Width           =   1395
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   2835
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5001
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0   'False
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin AACombo99.AACombo cDestino 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      ForeColor       =   12582912
      ListIndex       =   -1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AACombo99.AACombo cIntermediario 
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   1320
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      ForeColor       =   12582912
      ListIndex       =   -1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox tUsuReceptor 
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   13
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox tUsuario 
      Height          =   285
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   15
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox tComentario 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   6255
   End
   Begin VB.TextBox tCodigo 
      Height          =   285
      Left            =   900
      MaxLength       =   7
      TabIndex        =   1
      Top             =   540
      Width           =   915
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   5955
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   2055
      Left            =   60
      TabIndex        =   27
      Top             =   2760
      Visible         =   0   'False
      Width           =   6015
      _Version        =   196608
      _ExtentX        =   10610
      _ExtentY        =   3625
      _StockProps     =   229
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Documento:"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "&Label5"
      Height          =   195
      Left            =   4740
      TabIndex        =   8
      Top             =   2880
      Width           =   795
   End
   Begin VB.Label lbUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5625
      TabIndex        =   26
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label labUsuReceptor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2235
      TabIndex        =   25
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario que recibio:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario que ingresa:"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label labUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   24
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Destino:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentario:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label labFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   1020
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DETALLE TRANSFERENCIA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1020
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Transporte:"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label labOrigen 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Origen:"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   1185
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   900
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Código:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   855
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RecTransferencia.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RecTransferencia.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RecTransferencia.frx":0736
            Key             =   "Total"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RecTransferencia.frx":0A50
            Key             =   "Parcial"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RecTransferencia.frx":0D6A
            Key             =   "No"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RecTransferencia.frx":1084
            Key             =   "NoDoy"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RecTransferencia.frx":139E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuRecepcionar 
         Caption         =   "&Recepcionar"
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "RecTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 25-7-00 Dejo entrar stock negativo, Si me ponen + articulos de los originales modifico la cantidad de la tabla renglon y le sumo los otros articulos ademas mando mensaje a la terminal que lo hizo.
'27-3-2006 Elimine rs abierto
'                Cambios para documento
'29-6-2006 no dejo más entregar parcial.
'31-8-2006 si el camión me entrega parcial --> ahí hago un paso mostrando lo que está incompleto y luego hago mov de stock.

Option Explicit
Private EmpresaEmisora As clsClienteCFE
Private bMarqueControl As Boolean
Private TasaBasica As Currency, TasaMinima As Currency

Private Sub CargoValoresIVA()
Dim RsIva As rdoResultset
Dim sQy As String
    sQy = "SELECT IvaCodigo, IvaPorcentaje FROM TipoIva WHERE IvaCodigo IN (1,2)"
    Set RsIva = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsIva.EOF
        Select Case RsIva("IvaCodigo")
            Case 1: TasaBasica = RsIva("IvaPorcentaje")
            Case 2: TasaMinima = RsIva("IvaPorcentaje")
        End Select
        RsIva.MoveNext
    Loop
    RsIva.Close
End Sub

Private Function EmitirCFE(ByVal Documento As clsDocumentoCGSA, ByVal CAE As clsCAEDocumento) As String
On Error GoTo errEC
    If (TasaBasica = 0) Then CargoValoresIVA
    
    With New clsCGSAEFactura
        .URLAFirmar = prmURLFirmaEFactura
        .TasaBasica = TasaBasica
        .TasaMinima = TasaMinima
        .ImporteConInfoDeCliente = prmImporteConInfoCliente
        Set .Connect = cBase
        If Not .GenerarEComprobante(CAE, Documento, EmpresaEmisora, paCodigoDGI) Then
            EmitirCFE = .XMLRespuesta
        End If
    End With
    Exit Function
errEC:
    EmitirCFE = "Error en firma: " & Err.Description
End Function


Private Sub loc_FindHelp(ByVal sCons As String)
    Dim iCod As Long
    Dim objLista As New clsListadeAyuda
    If objLista.ActivarAyuda(cBase, sCons, 5000, 1, "Traslados") > 0 Then
        iCod = objLista.RetornoDatoSeleccionado(0)
    End If
    Me.Refresh
    Set objLista = Nothing
    If iCod > 0 Then BuscoTransferencia (iCod)
End Sub
Private Sub loc_FindTrasladoByDoc()
On Error GoTo errFTD
Dim sSerie As String, sNro As String
    
    If InStr(tDoc.Text, "-") <> 0 Then
        sSerie = Mid(tDoc.Text, 1, InStr(tDoc.Text, "-") - 1)
        sNro = Val(Mid(tDoc.Text, InStr(tDoc.Text, "-") + 1))
    Else
        sSerie = Mid(tDoc.Text, 1, 1)
        sNro = Val(Mid(tDoc.Text, 2))
    End If
    tDoc.Text = UCase(sSerie) & "-" & sNro
    
    Cons = "Select TraCodigo, TraSerie as Serie, TraNumero as Numero, SucAbreviacion as Sucursal " & _
                " From Traspaso, Sucursal Where TraSerie = '" & sSerie & "' And TraNumero = " & CLng(sNro) & _
                " And TraSucursal = SucCodigo And TraLocalDestino = " & paCodigoDeSucursal
    Screen.MousePointer = 11
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        sNro = RsAux("TraCodigo")
        RsAux.MoveNext
        If Not RsAux.EOF Then
            'Presento lista de ayuda hay más de uno.
            sNro = ""
            Dim objLista As New clsListadeAyuda
            If objLista.ActivarAyuda(cBase, Cons, 6000, 1, "Traslados de mercadería") > 0 Then
                sNro = objLista.RetornoDatoSeleccionado(0)
            End If
            Me.Refresh
            Set objLista = Nothing
        End If
    Else
        sNro = ""
        Screen.MousePointer = 0
        MsgBox "No existe un traslado que corresponda a los datos ingresados que tenga por destino su sucursal.", vbExclamation, "Atención"
    End If
    RsAux.Close
    
    If sNro <> "" Then BuscoTransferencia CLng(sNro)
Exit Sub
errFTD:
    Screen.MousePointer = 0
    MsgBox "Error al buscar el traslado por documento.", vbExclamation, "Atención"
End Sub

Private Sub cDestino_GotFocus()
    cDestino.SelStart = 0
    cDestino.SelLength = Len(cDestino.Text)
    Status.SimpleText = " Seleccione un Local. - [ F1] Pendientes, [ F2] Realizados -"
End Sub

Private Sub cDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1: BuscoTraspasoXDestino
    End Select
End Sub

Private Sub cDestino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If vsConsulta.Enabled Then vsConsulta.SetFocus
End Sub

Private Sub cDestino_LostFocus()
    cDestino.SelLength = 0
    Status.SimpleText = vbNullString
End Sub

Private Sub cIntermediario_GotFocus()
    cIntermediario.SelStart = 0
    cIntermediario.SelLength = Len(cIntermediario.Text)
    Status.SimpleText = " Seleccione un camión intermediario. - [ F1] Pendientes -"
End Sub

Private Sub cIntermediario_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1: BuscoTraspasoXIntermediario
    End Select
End Sub

Private Sub cIntermediario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cDestino.SetFocus
End Sub

Private Sub cIntermediario_LostFocus()
    cIntermediario.SelStart = 0
End Sub

Private Sub Form_Load()
On Error GoTo ErrL
    
    With vsConsulta
        .Redraw = False
        .Editable = False: .Rows = 1: .Cols = 1: .ExtendLastCol = True
        .FormatString = "Cantidad|<Estado|<Artículo|Cantidad que se recibe|Total"
        .ColWidth(1) = 1000: .ColWidth(2) = 3500: .ColWidth(3) = 4100: .ColWidth(4) = 1000
        .ColHidden(3) = True: .ColHidden(4) = True
        .Redraw = True
    End With

    If (EmpresaEmisora Is Nothing) Then
        Set EmpresaEmisora = New clsClienteCFE
        EmpresaEmisora.CargoClienteCarlosGutierrez paCodigoDeSucursal
    End If

    FechaDelServidor

    DeshabilitoIngreso
    
    Cons = "Select CamCodigo, CamNombre From Camion Order by CamNombre"
    CargoCombo Cons, cIntermediario, ""

    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cDestino, ""
    
    PrueboBandejaImpresora
    Screen.MousePointer = 0
    Exit Sub
    
ErrL:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al cargar el formulario.", Err.Description
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    CierroConexion
    cbMen.Close
    End
End Sub

Private Sub Label1_Click()
    Foco tCodigo
End Sub

Private Sub Label8_Click()
    Foco tComentario
End Sub

Private Sub MnuImprimir_Click()
    AccionImprimir
End Sub

Private Sub MnuRecepcionar_Click()
    AccionGrabar
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub tCodigo_Change()
    If Val(tCodigo.Tag) > 0 Then DeshabilitoIngreso: tDoc.Text = ""
End Sub
Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Ingrese el código de transferencia."
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCodigo.Text) And Val(tCodigo.Tag) = 0 Then
            BuscoTransferencia CLng(tCodigo.Text)
        ElseIf Val(tCodigo.Tag) = 0 Then
            MsgBox "El formato no es numérico.", vbExclamation, "ATENCIÓN"
            tCodigo.SetFocus
        End If
    End If

End Sub

Private Sub tCodigo_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Ingrese o modifique el comentario de traspaso."
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tUsuReceptor.SetFocus
End Sub

Private Sub tComentario_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub tDoc_Change()
    If Val(tCodigo.Tag) > 0 Then DeshabilitoIngreso
End Sub

Private Sub tDoc_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Val(tCodigo.Tag) = 0 Then loc_FindTrasladoByDoc
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "salir": Unload Me
        Case "recepcion": AccionGrabar
        Case "imprimir": AccionImprimir
    End Select
End Sub
Private Sub AccionImprimir()
    If MsgBox("Desea imprimir la hoja de detalle del traslado?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    Imprimo
End Sub
Private Sub DeshabilitoIngreso()
    tCodigo.Tag = ""
    tUsuario.BackColor = Inactivo: tUsuario.Enabled = False:  tUsuario.Text = vbNullString
    tUsuReceptor.BackColor = Inactivo: tUsuReceptor.Enabled = False: tUsuReceptor.Text = ""
    vsConsulta.Enabled = False: vsConsulta.Rows = 1
    tComentario.Enabled = False: tComentario.BackColor = Inactivo: tComentario.Text = vbNullString
    MnuImprimir.Enabled = False: Toolbar1.Buttons("imprimir").Enabled = False
    MnuRecepcionar.Enabled = False: Toolbar1.Buttons("recepcion").Enabled = False
    labFecha.Caption = vbNullString: labOrigen.Caption = vbNullString
    labUsuario.Caption = vbNullString: labUsuReceptor.Caption = vbNullString
    lbUsuario.Caption = vbNullString
    
    labOrigen.Caption = ""
    cIntermediario.Text = ""
    cDestino.Text = ""
End Sub
Private Sub HabilitoIngreso()
    vsConsulta.Enabled = True
    tComentario.Enabled = True: tComentario.BackColor = Blanco
    tUsuario.BackColor = Obligatorio: tUsuario.Enabled = True
    tUsuReceptor.BackColor = Obligatorio: tUsuReceptor.Enabled = True
    MnuImprimir.Enabled = True: Toolbar1.Buttons("imprimir").Enabled = True
    MnuRecepcionar.Enabled = True: Toolbar1.Buttons("recepcion").Enabled = True
End Sub

Private Sub BuscoTransferencia(ByVal lCod As Long)
On Error GoTo ErrBT
Dim rs As rdoResultset
    
    Screen.MousePointer = vbHourglass
    Cons = "Select * From Traspaso Where TraCodigo = " & lCod
    Set rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    DeshabilitoIngreso
    If rs.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "No existe un traspaso con ese código, verifique.", vbExclamation, "ATENCIÓN"
    Else
        If Not IsNull(rs("TraAnulado")) Then
            Screen.MousePointer = 0
            MsgBox "El traslado que selecciono está anulado.", vbExclamation, "Atención"
        Else
            CargoLista rs("TraCodigo")
            If Not IsNull(rs!TraLocalIntermedio) Then
                'Ojo acá dejo fimpresión x sistema anterior es lo mismo preguntar x serie y #
                If IsNull(rs!TraFechaEntregado) And Not IsNull(rs!TraFImpreso) Then
                    HabilitoIngreso
                ElseIf IsNull(rs!TraFImpreso) Then
                    MsgBox "Al intermediario no se le entregó la mercadería, verifique.", vbExclamation, "ATENCIÓN"
                    MnuImprimir.Enabled = True: Toolbar1.Buttons("imprimir").Enabled = True
                End If
            ElseIf IsNull(rs!TraFechaEntregado) And Not IsNull(rs!TraFImpreso) And IsNull(rs!TraLocalIntermedio) Then
                HabilitoIngreso
            End If
            
            labOrigen.Caption = BuscoNombreSucursal(rs!TraLocalOrigen)
            If Not IsNull(rs!TraLocalIntermedio) Then BuscoCodigoEnCombo cIntermediario, rs!TraLocalIntermedio
            BuscoCodigoEnCombo cDestino, rs!TraLocalDestino
            
            If rs!TraLocalDestino <> paCodigoDeSucursal And IsNull(rs!TraFechaEntregado) Then
                Screen.MousePointer = vbDefault
                If rs!TraLocalOrigen = paCodigoDeSucursal And IsNull(rs!TraFechaEntregado) Then
                    MsgBox "El local destino fijado no es este, si quiere devolver la mercadería al local de origen elimine el traslado.", vbExclamation, "ATENCIÓN"
                    MnuRecepcionar.Enabled = False
                    Toolbar1.Buttons("recepcion").Enabled = False
                    vsConsulta.Enabled = False
                Else
                    MsgBox "El local fijado como destino no es este, verifique.", vbExclamation, "ATENCIÓN"
                    DeshabilitoIngreso
                End If
            End If
            
            If Not IsNull(rs!TraUsuarioInicial) Then labUsuario.Caption = " " & BuscoUsuario(rs!TraUsuarioInicial, False, False, True)
            If Not IsNull(rs!TraUsuarioFinal) Then lbUsuario.Caption = " " & BuscoUsuario(rs!TraUsuarioFinal, False, False, True)
            If Not IsNull(rs!TraUsuarioReceptor) Then labUsuReceptor.Caption = " " & BuscoUsuario(rs!TraUsuarioReceptor, False, False, True)
            labFecha.Caption = Format(rs!TraFecha, "d-Mmm-yyyy")
            If Not IsNull(rs!TraComentario) Then tComentario.Text = Trim(rs!TraComentario)
            If Not IsNull(rs("TraSerie")) Then tDoc.Text = Trim(rs("TraSerie")) & "-" & Trim(rs("TraNumero"))
            tCodigo.Text = rs("TraCodigo")
            tCodigo.Tag = rs("TraCodigo")
        End If
    End If
    rs.Close
    Screen.MousePointer = vbDefault
    Exit Sub
ErrBT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al buscar la transferencia."
End Sub

Private Sub CargoLista(ByVal idTra As Long)
Dim aValor As Long

    Cons = "Select RenglonTraspaso.*, ArtNombre, ArtCodigo, EsMAbreviacion From RenglonTraspaso, Articulo, EstadoMErcaderia" _
        & " Where RTrTraspaso = " & idTra _
        & " And RTrArticulo = ArtID" _
        & " And RTrEstado = EsMCodigo"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    With vsConsulta
        Do While Not RsAux.EOF
            If RsAux!RTrPendiente = RsAux!RTrCantidad Then
                .AddItem "0"
            Else
                .AddItem RsAux!RTrCantidad - RsAux!RTrPendiente
            End If
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ESMAbreviacion)
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ArtCodigo, "(#,000,000) ") & Trim(RsAux!ArtNombre)
            
            aValor = RsAux!RTrArticulo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            aValor = RsAux!RTrEstado: .Cell(flexcpData, .Rows - 1, 1) = aValor
            aValor = RsAux!RTrCantidad: .Cell(flexcpData, .Rows - 1, 2) = aValor
            RsAux.MoveNext
        Loop
        RsAux.Close
    End With
    
End Sub

Private Function BuscoNombreSucursal(lnCod As Long) As String

    Cons = "Select  SucAbreviacion From Sucursal Where SucCodigo = " & lnCod
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    BuscoNombreSucursal = " " & Trim(RsAux!SucAbreviacion)
    RsAux.Close
    
End Function

Private Sub tUsuario_GotFocus()
    tUsuario.SelStart = 0
    tUsuario.SelLength = Len(tUsuario.Text)
    Status.SimpleText = " Ingrese su código de Usuario."
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    tUsuario.Tag = vbNullString
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = BuscoUsuarioDigito(CInt(tUsuario.Text), True)
            If CInt(tUsuario.Tag) > 0 Then
                AccionGrabar
            Else
                tUsuario.Tag = vbNullString
            End If
        Else
            MsgBox "El formato del código no es numérico.", vbExclamation, "ATENCIÓN"
            tUsuario.SetFocus
        End If
    End If
End Sub

Private Sub AccionGrabar()
Dim RsMerc As rdoResultset
'Dim Msg As String
Dim iCodEstado As Integer, idTerminal As Long
Dim iCant As Long
Dim sSuceso As Boolean
Dim strDefensa As String
Dim strSuceso As String

    If Trim(tUsuReceptor.Tag) = vbNullString Then
        MsgBox " Ingrese el código de usuario Receptor.", vbExclamation, "ATENCIÓN"
        tUsuReceptor.SetFocus
        Exit Sub
    End If
    If Trim(tUsuario.Tag) = vbNullString Then
        MsgBox " Ingrese su código de usuario.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    iCant = 0
    
    'Controlo las cantidades ingresadas.
    'En el tag de cada itmx tengo la cantidad real a entregar.
    
    Dim bRemitoMas As Boolean
    Dim bRemitoDeMenos As Boolean
    
    
    sSuceso = False
    With vsConsulta
        For I = 1 To .Rows - 1
'            If CLng(.Cell(flexcpText, I, 0)) > 0 Then
                iCant = 1
                If CLng(.Cell(flexcpText, I, 0)) <> CLng(.Cell(flexcpData, I, 2)) Then
                    sSuceso = True
                    If Not bMarqueControl Then .Cell(flexcpBackColor, I, 0, , .Cols - 1) = &H80&: .Cell(flexcpForeColor, I, 0, , .Cols - 1) = vbWhite
                    If CLng(.Cell(flexcpText, I, 0)) > CLng(.Cell(flexcpData, I, 2)) Then bRemitoMas = True
                    If CLng(.Cell(flexcpText, I, 0)) < CLng(.Cell(flexcpData, I, 2)) Then bRemitoDeMenos = True
                End If
'            End If
        Next
    End With
    
    
    If bRemitoMas Or bRemitoDeMenos Then
    
        Dim oInfoCAE As New clsCAEGenerador
        If Not oInfoCAE.SucursalTieneCae(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI) Then
            MsgBox "No hay un CAE disponible para emitir el eRemito, por favor comuniquese con administración." & vbCrLf & vbCrLf & "No podrá recepcionar", vbCritical, "eFactura"
            Screen.MousePointer = 0
            Exit Sub
        End If
    
    End If
    
    
    'Esto es si no le indique los errores y hay entonces me voy para que corrija
    If Not bMarqueControl And sSuceso Then
        MsgBox "La cantidad ingresada en algunos artículos es diferente a la asignada, verifique los datos ingresados y corrija los posibles errores.", vbExclamation, "Posible error"
        bMarqueControl = True
        Exit Sub
    Else
        If sSuceso Then
            If MsgBox("La cantidad ingresada en algunos artículos es diferente a la asignada." & vbCr & vbCr _
            & "¿Confirma grabar la información ingresada?" & vbCr & vbCr & "Se generaran los eRemitos necesarois para corregir los movimientos de stock.", vbExclamation + vbYesNo + vbDefaultButton2, "Posible error") = vbNo Then Exit Sub
        End If
    End If
    
    If iCant = 0 Then MsgBox "No hay datos a almacenar, verifique.", vbExclamation, "ATENCIÓN": Exit Sub
    
    If MsgBox("¿Confirma almacenar la recepción del traspaso.", vbQuestion + vbYesNo, "IMPRIMIR") = vbNo Then Exit Sub
    
    Dim aUsuario As Long
    aUsuario = 0: strDefensa = ""
    If sSuceso Then
        Dim objSuceso As New clsSuceso
        objSuceso.ActivoFormulario tUsuario.Tag, "Recepción de Traslado", cBase
        aUsuario = objSuceso.RetornoValor(True)
        strDefensa = objSuceso.RetornoValor(False, True)
        Set objSuceso = Nothing
        If aUsuario = 0 Then Exit Sub
    End If
    
    FechaDelServidor
    Dim rs As rdoResultset
    
    Screen.MousePointer = vbHourglass
'    Msg = vbNullString
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRelajo
    
    Cons = "Select * From Traspaso Where TraCodigo = " & Val(tCodigo.Tag)
    Set rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rs.EOF Then
        rs.Close
        cBase.RollbackTrans
        Screen.MousePointer = 0
        MsgBox "El traspaso fue eliminado por otra terminal, verifique.", vbExclamation, "Atención"
        Exit Sub
    Else
        If Not IsNull(rs!TraFechaEntregado) Then
            rs.Close
            cBase.RollbackTrans
            Screen.MousePointer = 0
            MsgBox "El traspaso ya fue entregado, verifique.", vbExclamation, "Atención"
            Exit Sub
        ElseIf Not IsNull(rs("TraAnulado")) Then
            rs.Close
            cBase.RollbackTrans
            Screen.MousePointer = 0
            MsgBox "El traslado fue anulado.", vbExclamation, "Atención"
            Exit Sub
        End If
    End If
    If Not IsNull(rs!TraTerminal) Then idTerminal = rs!TraTerminal Else idTerminal = 0
    
    
    Dim CAEDeMas As clsCAEDocumento
    Dim CAEDeMenos As clsCAEDocumento
    
    Dim oTraslDeMas As clsTraspaso
    Dim oTraslDeMenos As clsTraspaso
    
    Dim caeG As New clsCAEGenerador
    Dim docDeMas As clsDocumentoCGSA
    Dim docDeMenos As clsDocumentoCGSA
    
    If bRemitoDeMenos Then
        'Es cuando llega menos mercadería que la que generó el remito.
        'Inserto traslado con destino --> origen.
    
        Set CAEDeMenos = caeG.ObtenerNumeroCAEDocumento(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI)
        Set docDeMenos = New clsDocumentoCGSA
        
        With docDeMenos
            Set .Cliente = EmpresaEmisora
            .Emision = gFechaServidor
            .Tipo = TD_TrasladosInternos
            .Numero = CAEDeMas.Numero
            .Serie = CAEDeMas.Serie
            .Moneda.Codigo = 1
            .Total = 0
            .IVA = 0
            .Sucursal = paCodigoDeSucursal
            .Digitador = CInt(tUsuario.Tag)
            .Comentario = "Traslado " & CodTraslado
            .Vendedor = CInt(tUsuario.Tag)
            .Adenda = "Traslado de mercadería: " & CodTraslado & "<BR/>" & "Origen: " & cDestino.Text & ", Destino: " & labOrigen.Caption & "<BR/>" & "Memo: correción mercadería de menos."
        End With
        Set docDeMenos.Conexion = cBase
        docDeMenos.Codigo = docDeMenos.InsertoDocumentoBD(0)
        
        Set oTraslDeMenos = New clsTraspaso
        With oTraslDeMenos
            .Comentario "Corrección traslado " & Val(tCodigo.Tag) & " artículos que llegaron de más."
            .Fecha = gFechaServidor
            .FechaEntregado = gFechaServidor
            .FImpreso = gFechaServidor
            'No hago participar al camión.
            .LocalDestino = rs("TraLocalOrigen")
            .LocalOrigen = rs("TraLocalDestino")
            .Remito = docDeMenos.Codigo
            .UsuarioInicial = CInt(tUsuario.Tag)
            .UsuarioReceptor = .UsuarioInicial
            .UsuarioFinal = .UsuarioInicial
            .Codigo = .InsertarNuevoTraslado()
        End With
    End If
    
    
    If bRemitoMas Then
        'Es cuando llega más mercadería
        'Inserto traslado con origen --> destino.
    
        Set CAEDeMas = caeG.ObtenerNumeroCAEDocumento(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI)
        Set docDeMas = New clsDocumentoCGSA
        
        With docDeMas
            Set .Cliente = EmpresaEmisora
            .Emision = gFechaServidor
            .Tipo = TD_TrasladosInternos
            .Numero = CAEDeMas.Numero
            .Serie = CAEDeMas.Serie
            .Moneda.Codigo = 1
            .Total = 0
            .IVA = 0
            .Sucursal = paCodigoDeSucursal
            .Digitador = CInt(tUsuario.Tag)
            .Vendedor = CInt(tUsuario.Tag)
            .Comentario "Corrección traslado " & Val(tCodigo.Tag) & " artículos que llegaron de más."
            '.Adenda = "Traslado de mercadería: " & CodTraslado & "<BR/>" & "Origen: " & cOrigen.Text & IIf(cIntermediario.Text <> "", ", Camión: " & cIntermediario.Text, "") & ", Destino: " & cDestino.Text & "<BR/>" & "Memo: " & tComentario.Text
        End With
        Set docDeMas.Conexion = cBase
        docDeMas.Codigo = docDeMas.InsertoDocumentoBD(0)
        
        Set oTraslDeMas = New clsTraspaso
        With oTraslDeMas
            .Comentario "Corrección traslado " & Val(tCodigo.Tag) & " artículos que llegaron de más."
            .Fecha = gFechaServidor
            .FechaEntregado = gFechaServidor
            .FImpreso = gFechaServidor
            .LocalDestino = rs("TraLocalDestino")
            .LocalOrigen = rs("TraLocalOrigen")
            .Remito = docDeMenos.Codigo
            .UsuarioInicial = CInt(tUsuario.Tag)
            .UsuarioReceptor = .UsuarioInicial
            .UsuarioFinal = .UsuarioInicial
            .Codigo = .InsertarNuevoTraslado()
        End With
        
    End If
        
    Set caeG = Nothing
    
    Dim oRenTras As clsRenglonTraspaso
    
    With vsConsulta
        For I = 1 To .Rows - 1
            
            'Levanto primero los renglones para ver si se pasa en la cantidad.
            Cons = "Select * From RenglonTraspaso Where RTrTraspaso = " & Val(tCodigo.Tag) _
                & " And RTrArticulo = " & CLng(.Cell(flexcpData, I, 0)) & " And RTrEstado = " & CLng(.Cell(flexcpData, I, 1))
            Set RsMerc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    
'                If CLng(.Cell(flexcpText, I, 0)) <> RsMerc!RTrCantidad Then
'                    Msg = "Traslado código: " & rs!TraCodigo & Chr(13) & Chr(10) & "Existió diferencia para el artículo " & Trim(.Cell(flexcpText, I, 2)) & Chr(13) & Chr(10) _
                    & "Se debía recibir la cantidad de " & RsMerc!RTrCantidad & " arts. y se recibieron " & .Cell(flexcpText, I, 0) & Chr(13) & Chr(10)
                
'                    If CLng(.Cell(flexcpText, I, 0)) > RsMerc!RTrCantidad Then
'                        Msg = Msg & "Debe hacer un nuevo traslado del ORIGEN AL DESTINO por la diferencia para corregir el stock."
'                    Else
'                        Msg = Msg & "Debe hacer un nuevo traslado del DESTINO AL ORIGEN por la diferencia para corregir el stock."
'                    End If
                
'                    EnvioMensaje idTerminal, Msg
'                End If
'------------------------------------------------------------------------------CORRIJO STOCK X DIF (ORIGEN --> INTERMEDIARIO)
'31/8/2006 cambio
            'Cómo dejo ingresar parcial o superior --> o le doy más mercadería al camión desde el origen o le devuelvo.
            If Not IsNull(rs!TraLocalIntermedio) Then
                If CLng(.Cell(flexcpText, I, 0)) <> CLng(.Cell(flexcpData, I, 2)) Then
                    
                    If CLng(.Cell(flexcpText, I, 0)) > CLng(.Cell(flexcpData, I, 2)) Then
                        Set oRenTras = New clsRenglonTraspaso
                        With oRenTras
                            .Articulo = CLng(.Cell(flexcpData, I, 0))
                            .Cantidad = CLng(.Cell(flexcpText, I, 0)) - CLng(.Cell(flexcpData, I, 2))
                            .Estado = CLng(.Cell(flexcpData, I, 1))
                            .Pendientes = 0     'YA LO PONGO ENTREGADO.
                        End With
                    
                    
                        'Acá le quito al origen y se la doy al camión
                        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, rs!TraLocalOrigen, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpText, I, 0)) - CLng(.Cell(flexcpData, I, 2)), CLng(.Cell(flexcpData, I, 1)), -1
                        MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, rs!TraLocalOrigen, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpText, I, 0)) - CLng(.Cell(flexcpData, I, 2)), CLng(.Cell(flexcpData, I, 1)), -1, TipoDocumento.Traslados, rs!TraCodigo
                        'camión
                        MarcoMovimientoStockFisicoEnLocal TipoLocal.Camion, rs!TraLocalIntermedio, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpText, I, 0)) - CLng(.Cell(flexcpData, I, 2)), CLng(.Cell(flexcpData, I, 1)), 1
                        MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Camion, rs!TraLocalIntermedio, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpText, I, 0)) - CLng(.Cell(flexcpData, I, 2)), CLng(.Cell(flexcpData, I, 1)), 1, TipoDocumento.Traslados, rs!TraCodigo
                    Else
                        'Acá le quito al camión y se la doy al origen
                        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, rs!TraLocalOrigen, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpData, I, 2)) - CLng(.Cell(flexcpText, I, 0)), CLng(.Cell(flexcpData, I, 1)), 1
                        MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, rs!TraLocalOrigen, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpData, I, 2)) - CLng(.Cell(flexcpText, I, 0)), CLng(.Cell(flexcpData, I, 1)), 1, TipoDocumento.Traslados, rs!TraCodigo
                        'camión
                        MarcoMovimientoStockFisicoEnLocal TipoLocal.Camion, rs!TraLocalIntermedio, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpData, I, 2)) - CLng(.Cell(flexcpText, I, 0)), CLng(.Cell(flexcpData, I, 1)), -1
                        MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Camion, rs!TraLocalIntermedio, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpData, I, 2)) - CLng(.Cell(flexcpText, I, 0)), CLng(.Cell(flexcpData, I, 1)), -1, TipoDocumento.Traslados, rs!TraCodigo
                    End If
                    
                End If
                
            End If
'.................................................................................31/8/06
'FIN------------------------------------------------------------------------------CORRIJO STOCK X DIF (ORIGEN --> INTERMEDIARIO)

'------------------------------------------------------------------------------CORRIJO STOCK X DIF (iNTERMEDIARIO (SI HAY SI NO ES LOCAL ORIGEN) --> DESTINO)
            If CLng(.Cell(flexcpText, I, 0)) > 0 Then

                If Not IsNull(rs!TraLocalIntermedio) Then
                    'Quito mercadería del camión
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Camion, rs!TraLocalIntermedio, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpText, I, 0)), CLng(.Cell(flexcpData, I, 1)), -1
                    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Camion, rs!TraLocalIntermedio, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpText, I, 0)), CLng(.Cell(flexcpData, I, 1)), -1, TipoDocumento.Traslados, rs!TraCodigo
                Else
                    'Este caso es local origen. al no haber camión --> no hubo mov de stock
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, rs!TraLocalOrigen, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpText, I, 0)), CLng(.Cell(flexcpData, I, 1)), -1
                    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, rs!TraLocalOrigen, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpText, I, 0)), CLng(.Cell(flexcpData, I, 1)), -1, TipoDocumento.Traslados, rs!TraCodigo
                End If
            
           
                'Hago los movimientos para el destino.
                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpText, I, 0)), CLng(.Cell(flexcpData, I, 1)), 1
                MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), CLng(.Cell(flexcpText, I, 0)), CLng(.Cell(flexcpData, I, 1)), 1, TipoDocumento.Traslados, rs!TraCodigo
            End If
            
            RsMerc.Edit
            RsMerc!RTrPendiente = 0
            RsMerc.Update
            RsMerc.Close
            
            If CLng(.Cell(flexcpText, I, 0)) <> CLng(.Cell(flexcpData, I, 2)) And sSuceso Then
                strSuceso = "Traslado " & rs("TraCodigo") & "."
                If CLng(.Cell(flexcpText, I, 0)) > CLng(.Cell(flexcpData, I, 2)) Then
                    strSuceso = strSuceso & " Entregó más artículos"
                Else
                    strSuceso = " Entregó menos artículos."
                End If
                clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.RecepcionDeTraslados, paCodigoDeTerminal, aUsuario, 0, CLng(.Cell(flexcpData, I, 0)), strSuceso, strDefensa, CLng(.Cell(flexcpText, I, 0)) - CLng(.Cell(flexcpData, I, 2))
            End If

        Next
    End With
    
    rs.Edit
    rs!TraFModificacion = Format(gFechaServidor, sqlFormatoFH)
    rs!TraFechaEntregado = Format(gFechaServidor, sqlFormatoFH)
    rs!TraComentario = IIf(Trim(tComentario.Text) <> "", Trim(tComentario.Text), Null)
    If paCodigoDeSucursal <> rs!TraLocalDestino Then rs!TraLocalDestino = paCodigoDeSucursal
    rs!TraUsuarioReceptor = tUsuReceptor.Tag
    rs!TraUsuarioFinal = tUsuario.Tag
    rs.Update
    rs.Close
    
    cBase.CommitTrans
    DeshabilitoIngreso
    tCodigo.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrBT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al intentar iniciar la transaccion.", Err.Description
    Exit Sub
    
ErrRelajo:
    Resume Resumo
    Exit Sub
    
Resumo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al intentar grabar.", Err.Description

End Sub
Private Sub BuscoTraspasoXDestino()
On Error GoTo ErrBTXD

    cIntermediario.ListIndex = -1
    vsConsulta.Rows = 1
    tComentario.Text = vbNullString
        
    Cons = "Select Distinct(TraCodigo), Código = TraCodigo, Fecha = TraFecha, Origen = LocNombre, Intermediario = CamNombre From Traspaso " _
            & " Left Outer Join Camion ON TraLocalIntermedio = CamCodigo " _
            & ", Local, RenglonTraspaso " _
            & " Where TraLocalDestino = [idLocal] And TraAnulado Is Null" _
            & " And TraCodigo = RTrTraspaso And RTrPendiente > 0 " _
            & " And TraLocalOrigen = LocCodigo"
    If cDestino.ListIndex <> -1 Then
        Cons = Replace(Cons, "[idLocal]", cDestino.ItemData(cDestino.ListIndex))
    Else
        Cons = Replace(Cons, "[idLocal]", paCodigoDeSucursal)
    End If
    loc_FindHelp Cons
    Exit Sub
    
ErrBTXD:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al buscar la información."
End Sub

Private Sub BuscoTraspasoXIntermediario()
On Error GoTo ErrBTXI
    
    vsConsulta.Rows = 1
    cDestino.ListIndex = -1
    tComentario.Text = vbNullString
    
    If cIntermediario.ListIndex <> -1 Then
        Cons = "Select Distinct(TraCodigo), Código = TraCodigo, Fecha = TraFecha,  Origen = O.LocNombre, Destino =  D.LocNombre From Traspaso " _
                & " ,Camion, Local D, Local O, RenglonTraspaso " _
                & " Where TraLocalIntermedio = " & cIntermediario.ItemData(cIntermediario.ListIndex) _
                & " And TraCodigo = RTrTraspaso And RTrPendiente > 0 And TraAnulado Is Null" _
                & " And TraLocalDestino = D.LocCodigo" _
                & " And TraLocalOrigen = O.LocCodigo"
        loc_FindHelp Cons
    Else
        MsgBox "Debe seleccionar un local intermediario para poder consultar.", vbExclamation, "ATENCIÓN"
    End If
    Exit Sub
    
ErrBTXI:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al buscar la información."
End Sub
Private Sub tUsuReceptor_GotFocus()
    tUsuReceptor.SelStart = 0
    tUsuReceptor.SelLength = Len(tUsuReceptor.Text)
    Status.SimpleText = " Ingrese el código de Usuario Receptor."
End Sub

Private Sub tUsuReceptor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuReceptor.Text) Then
            tUsuReceptor.Tag = BuscoUsuarioDigito(CInt(tUsuReceptor.Text), True)
            If Val(tUsuReceptor.Tag) > 0 Then
                tUsuario.SetFocus
            Else
                tUsuReceptor.Tag = vbNullString
            End If
        Else
            MsgBox "El formato del código no es numérico.", vbExclamation, "ATENCIÓN"
            tUsuReceptor.Tag = vbNullString
        End If
    End If
End Sub

Private Sub vsConsulta_GotFocus()
    Status.SimpleText = " Ingrese la cantidad de artículos que recibe. [ + , - ] Agrega o resta."
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsConsulta.Rows > 1 Then
        Select Case KeyCode
            Case vbKeyReturn: Foco tComentario
            Case vbKeyAdd
                With vsConsulta
                    .Cell(flexcpText, .Row, 0) = CInt(.Cell(flexcpText, .Row, 0)) + 1
                    .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = vbWindowBackground
                    .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbButtonText
                End With
                bMarqueControl = False
            Case vbKeySubtract
                With vsConsulta
                    If CInt(.Cell(flexcpText, .Row, 0)) > 0 Then
                        .Cell(flexcpText, .Row, 0) = CInt(.Cell(flexcpText, .Row, 0)) - 1: bMarqueControl = False
                        .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = vbWindowBackground
                        .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbButtonText
                    End If
                End With
        End Select
    End If
End Sub

Private Sub vsConsulta_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub Imprimo()
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub

    Dim sPDef As String
    sPDef = Printer.DeviceName
    If paICartaN <> "" Then SeteoImpresoraPorDefecto paICartaN
    
    With vsListado
        .Device = paICartaN
        .Orientation = orPortrait
        .PaperSize = 1                     'Hoja carta
        .PaperBin = paICartaB         'Bandeja por defecto.
        
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    vsListado.FileName = "Traslado de Mercaderia"
    
    EncabezadoListado vsListado, "Traslado de Mercadería", False
    vsListado.FontBold = True
    vsListado.Paragraph = "Código de Traslado = " & Val(tCodigo.Tag)
    vsListado.Paragraph = "Origen: " & Trim(labOrigen.Caption) & Space(15) & "Camión:" & Trim(cIntermediario.Text) & Space(20) & "Destino: " & Trim(cDestino.Text)
    vsListado.Paragraph = "Terminal:  " & miConexion.NombreTerminal
    vsListado.Paragraph = ""
    vsListado.Paragraph = ""
    vsListado.Paragraph = "Recepcionado:" & Space(80) & "Camión:"
    vsListado.Paragraph = ""
    vsListado.FontBold = False
    
    With vsConsulta
        For I = 1 To .Rows - 1
            .RowHeight(I) = 400
        Next I
    End With
    vsConsulta.ColHidden(0) = True
    vsConsulta.ColHidden(3) = False
    vsConsulta.ColHidden(4) = False
    vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
    vsConsulta.ColHidden(3) = True
    vsConsulta.ColHidden(4) = True
    vsConsulta.ColHidden(0) = False
    With vsConsulta
        For I = 1 To .Rows - 1
            .RowHeight(I) = 240
        Next I
    End With
    vsListado.Paragraph = ""
    
    With vsListado
        .EndDoc
        .Device = paICartaN
        .PaperBin = paICartaB
        .PrintDoc
    End With
    SeteoImpresoraPorDefecto sPDef

    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    SeteoImpresoraPorDefecto sPDef
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
    vsConsulta.ColHidden(3) = True
End Sub
Private Sub PrueboBandejaImpresora()
On Error GoTo ErrPBI
    Exit Sub
    With vsListado
        .PaperSize = 1  'Hoja carta
        .Orientation = orPortrait
        .Device = paICartaN
        If .Device <> paICartaN Then MsgBox "Ud no tiene instalada la impresora para hoja blanca. Avise al administrador.", vbExclamation, "ATENCIÒN"
        If .PaperBins(paICartaB) Then .PaperBin = paICartaB Else MsgBox "Esta mal definida la bandeja de hoja blanca para grillas en su sucursal, comuniquele al administrador.", vbInformation, "ATENCIÓN": paICartaB = .PaperBin
    End With
    Exit Sub
ErrPBI:
    clsGeneral.OcurrioError "Error al setear la impresora, consulte con el administrador de impresión este problema.", Err.Description
End Sub

Private Sub EnvioMensaje(idTerm As Long, TextoMensaje As String)
Dim RsMensaje As rdoResultset
Dim aValor As Long, strHora As String
   
    'Cargo datos Tabla Mensaje----------------------------------
    Cons = "Select * From Mensaje Where MenID = 0"
    Set RsMensaje = cbMen.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsMensaje.AddNew
    RsMensaje!MenAsunto = Trim("Recepción de Traslado con Diferencias")
    RsMensaje!MenDe = tUsuario.Tag
    RsMensaje!MenEnviado = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsMensaje!MenCategoria = paCategoriaMensajeStock
    RsMensaje!MenPublico = 0
    RsMensaje!MenFechaHora = Format(DateAdd("s", 10, gFechaServidor), "mm/dd/yyyy hh:mm:ss")
    RsMensaje!MenTexto = Trim(TextoMensaje)
    RsMensaje.Update: RsMensaje.Close
    '------------------------------------------------------------------------------------------------------
    
    Cons = "Select Max(MenID) From Mensaje " _
        & " Where MenAsunto = '" & Trim("Recepción de Traslado con Diferencias") & "'"
    Set RsMensaje = cbMen.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aValor = RsMensaje(0)
    RsMensaje.Close
    
    'Cargo para Quien va el mensaje--------------------------------------------------------------------
'Public Enum TUsuarioM
'    Persona = 1
'    Grupo = 2
'    Terminal = 3
'End Enum

    Cons = "Insert Into MensajeUsuario (MUsIdMensaje, MUsIdUsuario, MUsTipoUsr) " & _
                "Values ( " & aValor & ", " & idTerm & ", 3)"
    cbMen.Execute Cons
    
    'Le envío el mensaje a los usuarios verificadores
    If paUserVerif <> "" Then
        Dim vUser() As String
        vUser = Split(paUserVerif, ",")
        Dim iQ As Integer
        For iQ = 0 To UBound(vUser)
            If Val(vUser(iQ)) > 0 Then
                Cons = "Insert Into MensajeUsuario (MUsIdMensaje, MUsIdUsuario, MUsTipoUsr) " & _
                    "Values ( " & aValor & ", " & vUser(iQ) & ", 1)"
                cbMen.Execute Cons
            End If
        Next
    End If
    '------------------------------------------------------------------------------------------------------
    
End Sub

Private Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

Private Sub EncabezadoListado(vsPrint As Control, strTitulo As String, sNombreEmpresa As Boolean)
    
    With vsPrint
        .HdrFont = "Arial"
        .HdrFontSize = 10
        .HdrFontBold = False
    End With
    
    If sNombreEmpresa Then
        vsPrint.Header = strTitulo + "||Carlos Gutiérrez S.A."
    Else
        vsPrint.Header = strTitulo
    End If
    vsPrint.HdrFontBold = False: vsPrint.FontBold = False
    vsPrint.HdrFontSize = 10: vsPrint.Footer = Format(Now, "dd/mm/yy hh:mm")
    
End Sub

Private Function BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
Dim RsUsr As rdoResultset
Dim aRetorno As String: aRetorno = ""
    
    On Error Resume Next
    
    Cons = "Select * from Usuario Where UsuCodigo = " & Codigo
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Digito Then aRetorno = Trim(RsUsr!UsuDigito)
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    
    BuscoUsuario = aRetorno
    
End Function

Private Function BuscoUsuarioDigito(Digito As Long, Optional Codigo As Boolean = False, Optional Identificacion As Boolean = False, Optional Iniciales As Boolean = False) As Variant
Dim RsUsr As rdoResultset
Dim aRetorno As Variant
On Error GoTo ErrBUD

    Cons = "Select * from Usuario Where UsuDigito = " & Digito
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Codigo Then aRetorno = RsUsr!UsuCodigo
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    BuscoUsuarioDigito = aRetorno
    Exit Function
    
ErrBUD:
    MsgBox "Error inesperado al buscar el usuario.", vbCritical, "ATENCIÓN"
End Function

