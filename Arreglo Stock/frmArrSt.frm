VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmArrST 
   Caption         =   "Arreglo Stock"
   ClientHeight    =   4215
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmArrSt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tArticulo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin ComctlLib.ProgressBar bProgreso 
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   2580
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
      _Version        =   196608
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   229
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
   Begin VB.PictureBox picBotones 
      BorderStyle     =   0  'None
      Height          =   425
      Left            =   240
      ScaleHeight     =   420
      ScaleWidth      =   2295
      TabIndex        =   8
      Top             =   2520
      Width           =   2295
      Begin VB.CommandButton bArreglar 
         Height          =   310
         Left            =   480
         Picture         =   "frmArrSt.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Arreglar Stock. [Ctrl+A]"
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmArrSt.frx":068C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   1440
         Picture         =   "frmArrSt.frx":098E
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   1080
         Picture         =   "frmArrSt.frx":0A90
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   50
         Width           =   310
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrilla 
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2355
      _ConvInfo       =   1
      Appearance      =   0
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
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3960
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmArrST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsSt As rdoResultset

Private Sub bArreglar_Click()
    AccionArregloStock
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    CargoGrilla
End Sub

Private Sub bImprimir_Click()
    AccionImprimir
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    DoEvents
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyE: CargoGrilla
            Case vbKeyI: AccionImprimir
            Case vbKeyX: Unload Me
        End Select
    End If
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    ObtengoSeteoForm Me, 1000, 500, 3840, 4230
    LimpioGrilla
    vsListado.Visible = False
    Exit Sub
ErrLoad:
    clGeneral.OcurrioError "Ocurrio un error al iniciar el formulario.", Trim(Err.Description)
End Sub
Private Sub Form_Resize()
On Error Resume Next
    picBotones.Top = Me.ScaleHeight - (Status.Height + 40 + picBotones.Height)
    bProgreso.Top = picBotones.Top + 80
    vsGrilla.Height = Me.ScaleHeight - (190 + Status.Height + picBotones.Height + vsGrilla.Top)
    vsGrilla.Width = Me.ScaleWidth - (vsGrilla.Left * 2)
    bProgreso.Width = Me.ScaleWidth - (bProgreso.Left + vsGrilla.Left)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clGeneral = Nothing
    Set miconexion = Nothing
    End
End Sub
Private Sub ArregloTablaStockTotal(QCantidad As Long, IDTipoEstado As Integer, IDEstado As Long, IdArticulo As Long)
    Cons = "Select * From StockTotal Where StTArticulo = " & IdArticulo _
        & " And SttTipoEstado = " & IDTipoEstado _
        & " And stTEstado = " & IDEstado
    Set RsSt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsSt.EOF Then
        RsSt.AddNew
        RsSt!StTArticulo = IdArticulo
        RsSt!StTTipoEstado = IDTipoEstado
        RsSt!StTEstado = IDEstado
        RsSt!StTCantidad = QCantidad
        RsSt.Update
    Else
        If RsSt!StTCantidad + QCantidad = 0 Then
            RsSt.Delete
        Else
            RsSt.Edit
            RsSt!StTCantidad = RsSt!StTCantidad + QCantidad
            RsSt.Update
        End If
    End If
    RsSt.Close
    If IDTipoEstado = TipoEstadoMercaderia.Virtual Then MarcoMovimientoStockEstado paCodigoDeUsuario, IdArticulo, CCur(QCantidad), CInt(IDEstado), 1, TipoDocumento.ArregloStock, 1
    
End Sub
Private Sub ArregloTablaStockLocal(QCantidad As Long, IDEstado As Long, IdArticulo As Long, IdLocal As Long, iTipoLocal As Integer)
    
    Cons = "Select * From StockLocal Where StLArticulo = " & IdArticulo _
        & " And stLEstado = " & IDEstado _
        & " And StLLocal = " & IdLocal & " And StLTipoLocal = " & iTipoLocal
        
    Set RsSt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsSt.EOF Then
        RsSt.AddNew
        RsSt!StLArticulo = IdArticulo
        RsSt!StLTipoLocal = iTipoLocal
        RsSt!StlLocal = IdLocal
        RsSt!StLEstado = IDEstado
        RsSt!StLCantidad = QCantidad
        RsSt.Update
    Else
        If RsSt!StLCantidad + QCantidad = 0 Then
            RsSt.Delete
        Else
            RsSt.Edit
            RsSt!StLCantidad = RsSt!StLCantidad + QCantidad
            RsSt.Update
        End If
    End If
    RsSt.Close
    MarcoMovimientoStockFisico paCodigoDeUsuario, iTipoLocal, IdLocal, IdArticulo, CCur(QCantidad), CInt(IDEstado), 1, TipoDocumento.ArregloStock, 1
    
End Sub

Private Sub AccionArregloStock()
On Error GoTo ErrAAS
Dim CodAux As Long
Dim Rs As rdoResultset, RsArt As rdoResultset
Dim Qtotal As Long, Q1 As Long, Q2 As Long, QDifCamion As Long
    
    If MsgBox("¿Confirma iniciar el proceso que corrige el stock.?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    LimpioGrilla
    Cons = "Select Count(*) From Articulo "
    If Val(tArticulo.Tag) > 0 Then Cons = Cons & " Where ArtID = " & tArticulo.Tag
    Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsArt(0) = 0 Then MsgBox "No hay datos a listar.", vbExclamation, "ATENCIÓN": Screen.MousePointer = 0: Exit Sub
    bProgreso.Value = 0
    bProgreso.Max = RsArt(0)
    RsArt.Close
    FechaDelServidor
    
    Cons = "Select * From Articulo "
    If Val(tArticulo.Tag) > 0 Then Cons = Cons & " Where ArtID = " & tArticulo.Tag
    Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsArt.EOF
        If RsArt!ArtTipo = paTipoArticuloServicio Then
            'Borro todo porque son de servicio.
            Cons = "Delete StockTotal Where StTArticulo = " & RsArt!ArtID
            cBase.Execute (Cons)
            Cons = "Delete StockLocal Where StLArticulo = " & RsArt!ArtID
            cBase.Execute (Cons)
        Else
            'Cargo la lista de Estados.------------------------------------------------------
            Cons = "Select StTEstado, Sum(StTCantidad) From StockTotal Where StTArticulo = " & RsArt!ArtID _
                & " And SttTipoEstado = " & TipoEstadoMercaderia.Virtual _
                & " Group by StTEstado "
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            Do While Not RsAux.EOF
                Qtotal = RsAux(1)
                If TipoMovimientoEstado.Reserva <> RsAux!StTEstado Then
                    If TipoMovimientoEstado.AEntregar = RsAux!StTEstado Then
                        Q1 = StockAFacturarEnvia(RsArt!ArtID)
                        Q2 = StockEnEnvio(RsArt!ArtID)
                    ElseIf TipoMovimientoEstado.ARetirar = RsAux!StTEstado Then
                        'Retira
                        Q1 = StockAFacturarRetira(RsArt!ArtID)
                        Q2 = StockARetirar(RsArt!ArtID)
                    End If
                    If Qtotal <> 0 And Q1 + Q2 = 0 Then
                        ArregloTablaStockTotal Qtotal, TipoEstadoMercaderia.Fisico, CLng(paEstadoArticuloEntrega), RsArt!ArtID
                        ArregloTablaStockTotal Qtotal * -1, TipoEstadoMercaderia.Virtual, RsAux!StTEstado, RsArt!ArtID
                    ElseIf Qtotal - (Q1 + Q2) <> 0 Then
                        ArregloTablaStockTotal Qtotal - (Q1 + Q2), TipoEstadoMercaderia.Fisico, CLng(paEstadoArticuloEntrega), RsArt!ArtID
                        ArregloTablaStockTotal (Qtotal - (Q1 + Q2)) * -1, TipoEstadoMercaderia.Virtual, RsAux!StTEstado, RsArt!ArtID
                    End If
                Else
                    ArregloTablaStockTotal Qtotal, TipoEstadoMercaderia.Fisico, CLng(paEstadoArticuloEntrega), RsArt!ArtID
                    ArregloTablaStockTotal Qtotal * -1, TipoEstadoMercaderia.Virtual, RsAux!StTEstado, RsArt!ArtID
                End If
                RsAux.MoveNext
            Loop
            RsAux.Close
            'Si no tiene local donde retirar la mercadería dejo como esta y lo incluyo en la lista indicando la diferencia.
            If Not IsNull(RsArt!ArtLocalRetira) Then
                
                Cons = "Select * From EstadoMercaderia Where EsMBajaStockTotal = 0"
                Set Rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                Do While Not Rs.EOF
                
                    'Cargo Stock de Camiones.
'                    Cons = "Select CamCodigo, CamNombre, Sum(STlCantidad) From StockLocal, Camion" _
                        & " Where StLArticulo = " & RsArt!ArtID _
                        & " And StLTipoLocal = " & TipoLocal.Camion _
                        & " And StLEstado = " & Rs!EsMCodigo & " And StLLocal = CamCodigo" _
                        & " Group By CamCodigo, CamNombre"
                    
                    Cons = "Select CamCodigo, CamNombre, Cantidad = Sum(STlCantidad) From Camion " _
                                & " Left Outer Join StockLocal On StLLocal = CamCodigo" _
                                                    & " And StLArticulo = " & RsArt!ArtID _
                                                    & " And StLTipoLocal = " & TipoLocal.Camion _
                                                    & " And StLEstado = " & Rs!EsMCodigo _
                                & " Left Outer Join  EstadoMercaderia On StLEstado = EsMCodigo And EsMBajaStockTotal = 0 " _
                        & " Group By CamCodigo, CamNombre"
                    
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    
                    Do While Not RsAux.EOF
                        If Not IsNull(RsAux!Cantidad) Then Qtotal = RsAux!Cantidad Else Qtotal = 0
                        Q1 = StockEnTrasladoXEstado(RsArt!ArtID, RsAux!CamCodigo, Rs!EsMCodigo)
                        Q2 = StockCamionEnvioEntregadoXEstado(RsArt!ArtID, RsAux!CamCodigo, Rs!EsMCodigo)
                        
                        If Qtotal - (Q1 + Q2) <> 0 Then
                            ArregloTablaStockLocal (Qtotal - (Q1 + Q2)) * -1, Rs!EsMCodigo, RsArt!ArtID, RsAux!CamCodigo, TipoLocal.Camion
                            ArregloTablaStockLocal Qtotal - (Q1 + Q2), Rs!EsMCodigo, RsArt!ArtID, RsArt!ArtLocalRetira, TipoLocal.Deposito
'                        ElseIf Qtotal - (Q1 + Q2) < 0 Then
'                            ArregloTablaStockLocal Qtotal - (Q1 + Q2), Rs!EsMCodigo, RsArt!ArtID, RsAux!CamCodigo, TipoLocal.Camion
'                            ArregloTablaStockLocal (Qtotal - (Q1 + Q2)) * -1, Rs!EsMCodigo, RsArt!ArtID, RsArt!ArtLocalRetira, TipoLocal.Deposito
                        End If
                        RsAux.MoveNext
                    Loop
                    RsAux.Close
                    Rs.MoveNext
                Loop
                Rs.Close
            Else
                'Cargo Stock de Camiones.
                'Cons = "Select CamCodigo, CamNombre, Sum(STlCantidad) From StockLocal, Camion, EstadoMercaderia " _
                    & " Where StLArticulo = " & RsArt!ArtID _
                    & " And StLTipoLocal = " & TipoLocal.Camion _
                    & " And StLEstado = EsMCodigo And EsMBajaStockTotal = 0 And StLLocal = CamCodigo" _
                    & " Group By CamCodigo, CamNombre"
                
                Cons = "Select CamCodigo, CamNombre, Cantidad = Sum(STlCantidad) From Camion " _
                                & " Left Outer Join StockLocal On StLLocal = CamCodigo" _
                                                    & " And StLArticulo = " & RsArt!ArtID _
                                                    & " And StLTipoLocal = " & TipoLocal.Camion _
                                & " Left Outer Join  EstadoMercaderia On StLEstado = EsMCodigo And EsMBajaStockTotal = 0 " _
                        & " Group By CamCodigo, CamNombre"
                
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                
                Do While Not RsAux.EOF
                    With vsGrilla
                        .AddItem ""
                        .Cell(flexcpText, vsGrilla.Rows - 1, 0) = Format(RsArt!ArtCodigo, "#,000,000") & " " & Trim(RsArt!ArtNombre)
                        CodAux = 2  'Me digo que es un camión
                        .Cell(flexcpData, vsGrilla.Rows - 1, 1) = CodAux
                        .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!CamNombre)
                        CodAux = RsAux!CamCodigo
                        .Cell(flexcpData, vsGrilla.Rows - 1, 2) = CodAux
                        If Not IsNull(RsAux!Cantidad) Then Qtotal = RsAux!Cantidad Else Qtotal = 0
                        .Cell(flexcpText, .Rows - 1, 2) = Qtotal: .Cell(flexcpBackColor, .Rows - 1, 2) = vbBlue: .Cell(flexcpForeColor, .Rows - 1, 2) = vbWhite: .Cell(flexcpFontBold, .Rows - 1, 2) = True
                        .Cell(flexcpText, .Rows - 1, 3) = StockEnTraslado(RsArt!ArtID, RsAux!CamCodigo) + StockEnTrasladoPendiente(RsArt!ArtID, RsAux!CamCodigo)
                        .Cell(flexcpText, .Rows - 1, 4) = StockCamionEnvioEntregado(RsArt!ArtID, RsAux!CamCodigo)
                        .Cell(flexcpText, .Rows - 1, 5) = Val(.Cell(flexcpText, .Rows - 1, 2)) - (Val(.Cell(flexcpText, .Rows - 1, 4)) + Val(.Cell(flexcpText, .Rows - 1, 3)))
                        If Val(.Cell(flexcpText, .Rows - 1, 5)) <> 0 Then
                            .Cell(flexcpForeColor, .Rows - 1, 5) = vbWhite: .Cell(flexcpBackColor, .Rows - 1, 5) = vbRed: .Cell(flexcpFontBold, .Rows - 1, 5) = True
                        Else
                            .RemoveItem .Rows - 1   'La borro porque no hay diferencia
                        End If
                    End With
                    RsAux.MoveNext
                Loop
                RsAux.Close
            End If
        End If
        RsArt.MoveNext
        bProgreso.Value = bProgreso.Value + 1
    Loop
    RsArt.Close
    If vsGrilla.Rows > 1 Then MsgBox "Los artículos que figuran en la grilla son aquellos que no tienen asignado un local de retiro.", vbExclamation, "ATENCIÓN"
    Screen.MousePointer = 0
    Exit Sub
ErrAAS:
    clGeneral.OcurrioError "Ocurrio un error al cargar los datos en la grilla.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub CargoGrilla()
On Error GoTo ErrCI
Dim CodAux As Long, aCant As Currency
Dim Rs As rdoResultset, RsArt As rdoResultset
    
    Screen.MousePointer = 11
    LimpioGrilla
    Cons = "Select count(*) From Articulo"
    If Val(tArticulo.Tag) > 0 Then Cons = Cons & " Where ArtID = " & tArticulo.Tag
    Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsArt(0) = 0 Then MsgBox "No hay datos a listar.", vbExclamation, "ATENCIÓN": Screen.MousePointer = 0: Exit Sub
    bProgreso.Value = 0
    bProgreso.Max = RsArt(0)
    RsArt.Close
    
    Cons = "Select * From Articulo "
    If Val(tArticulo.Tag) > 0 Then Cons = Cons & " Where ArtID = " & tArticulo.Tag
    Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsArt.EOF
        'Cargo la lista de Estados.------------------------------------------------------
        Cons = "Select StTEstado, Sum(StTCantidad) From StockTotal Where StTArticulo = " & RsArt!ArtID _
            & " And SttTipoEstado = " & TipoEstadoMercaderia.Virtual _
            & " Group by StTEstado "
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            With vsGrilla
                .AddItem ""
                .Cell(flexcpText, vsGrilla.Rows - 1, 0) = Format(RsArt!ArtCodigo, "#,000,000") & " " & Trim(RsArt!ArtNombre)
                CodAux = 1  'Me digo que es un estado
                .Cell(flexcpData, vsGrilla.Rows - 1, 1) = CodAux
                .Cell(flexcpText, vsGrilla.Rows - 1, 1) = RetornoNombreEstado(RsAux!StTEstado)
                CodAux = RsAux!StTEstado
                .Cell(flexcpData, vsGrilla.Rows - 1, 2) = CodAux
                .Cell(flexcpText, vsGrilla.Rows - 1, 2) = RsAux(1): .Cell(flexcpBackColor, .Rows - 1, 2) = vbBlue: .Cell(flexcpForeColor, .Rows - 1, 2) = vbWhite: .Cell(flexcpFontBold, .Rows - 1, 2) = True
                If TipoMovimientoEstado.AEntregar = RsAux!StTEstado Then
                    .Cell(flexcpText, vsGrilla.Rows - 1, 4) = StockAFacturarEnvia(RsArt!ArtID)
                    .Cell(flexcpText, vsGrilla.Rows - 1, 3) = StockEnEnvio(RsArt!ArtID)
                ElseIf TipoMovimientoEstado.ARetirar = RsAux!StTEstado Then
                    'Retira
                    .Cell(flexcpText, vsGrilla.Rows - 1, 4) = StockAFacturarRetira(RsArt!ArtID)
                    .Cell(flexcpText, vsGrilla.Rows - 1, 3) = StockARetirar(RsArt!ArtID)
                End If
                .Cell(flexcpText, vsGrilla.Rows - 1, 5) = Val(.Cell(flexcpText, vsGrilla.Rows - 1, 2)) - (Val(.Cell(flexcpText, vsGrilla.Rows - 1, 4)) + Val(.Cell(flexcpText, vsGrilla.Rows - 1, 3)))
                If Val(.Cell(flexcpText, vsGrilla.Rows - 1, 5)) <> 0 Then
                    .Cell(flexcpForeColor, .Rows - 1, 5) = vbWhite: .Cell(flexcpBackColor, .Rows - 1, 5) = vbRed: .Cell(flexcpFontBold, .Rows - 1, 5) = True
                Else
                    .RemoveItem .Rows - 1
                End If
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        'Cargo Stock de Camiones.
        'Cons = "Select CamCodigo, CamNombre, Sum(STlCantidad) From StockLocal, Camion, EstadoMercaderia " _
            & " Where StLArticulo = " & RsArt!ArtID _
            & " And StLTipoLocal = " & TipoLocal.Camion _
            & " And StLEstado = EsMCodigo And EsMBajaStockTotal = 0 And StLLocal = CamCodigo" _
            & " Group By CamCodigo, CamNombre"
        
        Cons = "Select CamCodigo, CamNombre, Cantidad = Sum(STlCantidad) From Camion " _
            & " Left Outer Join StockLocal On StLLocal = CamCodigo" _
                                & " And StLArticulo = " & RsArt!ArtID _
                                & " And StLTipoLocal = " & TipoLocal.Camion _
            & " Left Outer Join  EstadoMercaderia On StLEstado = EsMCodigo And EsMBajaStockTotal = 0 " _
        & " Group By CamCodigo, CamNombre"
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        Do While Not RsAux.EOF
            With vsGrilla
                .AddItem ""
                .Cell(flexcpText, vsGrilla.Rows - 1, 0) = Format(RsArt!ArtCodigo, "#,000,000") & " " & Trim(RsArt!ArtNombre)
                CodAux = 2  'Me digo que es un camión
                .Cell(flexcpData, vsGrilla.Rows - 1, 1) = CodAux
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!CamNombre)
                CodAux = RsAux!CamCodigo
                .Cell(flexcpData, vsGrilla.Rows - 1, 2) = CodAux
                If Not IsNull(RsAux!Cantidad) Then aCant = RsAux!Cantidad Else aCant = 0
                .Cell(flexcpText, .Rows - 1, 2) = aCant: .Cell(flexcpBackColor, .Rows - 1, 2) = vbBlue: .Cell(flexcpForeColor, .Rows - 1, 2) = vbWhite: .Cell(flexcpFontBold, .Rows - 1, 2) = True
                .Cell(flexcpText, .Rows - 1, 3) = StockEnTraslado(RsArt!ArtID, RsAux!CamCodigo) + StockEnTrasladoPendiente(RsArt!ArtID, RsAux!CamCodigo)
                .Cell(flexcpText, .Rows - 1, 4) = StockCamionEnvioEntregado(RsArt!ArtID, RsAux!CamCodigo)
                .Cell(flexcpText, .Rows - 1, 5) = Val(.Cell(flexcpText, .Rows - 1, 2)) - (Val(.Cell(flexcpText, .Rows - 1, 4)) + Val(.Cell(flexcpText, .Rows - 1, 3)))
                If Val(.Cell(flexcpText, .Rows - 1, 5)) <> 0 Then
                    .Cell(flexcpForeColor, .Rows - 1, 5) = vbWhite: .Cell(flexcpBackColor, .Rows - 1, 5) = vbRed: .Cell(flexcpFontBold, .Rows - 1, 5) = True
                Else
                    .RemoveItem .Rows - 1   'La borro porque no hay diferencia
                End If
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        RsArt.MoveNext
        bProgreso.Value = bProgreso.Value + 1
    Loop
    RsArt.Close
    Screen.MousePointer = 0
    Exit Sub
ErrCI:
    clGeneral.OcurrioError "Ocurrio un error al cargar los datos en la grilla.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub LimpioGrilla()
    With vsGrilla
        .Redraw = False
        .ExtendLastCol = True
        .Clear
        .Rows = 1
        .FixedCols = 0
        .Cols = 1
        .FormatString = "Artículo|<Estado/Camión|>Q Total|>Q Vta. Mostr./Traslados|>Q Vta. Telf./Envíos|>Q Diferencia|"
        .ColWidth(0) = 2000: .ColWidth(1) = 1600
        .ColWidth(2) = 1000: '.ColWidth(3) = 1900
        '.ColWidth(4) = 1800: .ColWidth(5) = 1000
        .ColWidth(6) = 14
        .AllowUserResizing = flexResizeColumns
        .MergeCells = flexMergeRestrictAll: .MergeCol(0) = True
        .Redraw = True
    End With
    
End Sub


Private Sub Label1_Click()
    Foco tArticulo
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(tArticulo.Text) Then
            BuscoArticulo 0, tArticulo.Text
        Else
            AyudaArticulo
        End If
    Else
        LimpioCampos
    End If
End Sub

Private Sub BuscoArticulo(IdArticulo As Long, Optional CodArticulo As Long = 0)
On Error GoTo ErrBA
    Screen.MousePointer = 11
    If CodArticulo > 0 Then
        Cons = "Select * From Articulo Where ArtCodigo = " & CodArticulo
    Else
        Cons = "Select * From Articulo Where ArtID = " & IdArticulo
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        LimpioCampos
        RsAux.Close
        MsgBox "No se encontró un artículo con esos datos.", vbExclamation, "ATENCIÓN"
    Else
        tArticulo.Tag = RsAux!ArtID
        tArticulo.Text = Trim(RsAux!ArtNombre)
        RsAux.Close
        Foco bConsultar
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBA:
    clGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub AyudaArticulo()
On Error GoTo ErrIA
Dim Resultado As Long
    Screen.MousePointer = 11
    Cons = "Select ID = ArtID,  Código = ArtCodigo, Nombre = RTRIM(ArtNombre) From Articulo Where ArtNombre Like '" & Trim(tArticulo.Text) & "%'"
    Dim LiAyuda As New clsListadeAyuda
    Screen.MousePointer = 0
    LiAyuda.ActivoListaAyuda Cons, False, cBase.Connect
    Me.Refresh
    Screen.MousePointer = 11
    Resultado = LiAyuda.ValorSeleccionado
    If Resultado > 0 Then
        BuscoArticulo Resultado
    Else
        LimpioCampos
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrIA:
    Screen.MousePointer = 0
    clGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
End Sub

Private Sub vsgrilla_Click()
    If vsGrilla.MouseRow = 0 Then
        vsGrilla.ColSel = vsGrilla.MouseCol
        If vsGrilla.ColSort(vsGrilla.MouseCol) = flexSortGenericAscending Then
            vsGrilla.ColSort(vsGrilla.MouseCol) = flexSortGenericDescending
        Else
            vsGrilla.ColSort(vsGrilla.MouseCol) = flexSortGenericAscending
        End If
        vsGrilla.Sort = flexSortUseColSort
    End If
End Sub
Private Sub ModificoRegistro()
On Error GoTo ErrNI
    Screen.MousePointer = 11
    Cons = "Select * From Where <> " & vsGrilla.Cell(flexcpText, vsGrilla.Row, 0) _
        & " And  = '"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Close
        MsgBox "Ya existe ............... con ese nombre, verifique.", vbExclamation, "ATENCIÓN"
    Else
        RsAux.Close
        'Cons = "Update CodigoTexto set Texto = '" & Trim(tNombre.Text) _
            & "' Where Codigo = " & vsGrilla.Cell(flexcpText, vsGrilla.Row, 0)
        'cBase.Execute (Cons)
    End If
    CargoGrilla
    vsGrilla.Enabled = True
    
    Screen.MousePointer = 0
    Exit Sub
ErrNI:
    clGeneral.OcurrioError "Ocurrio un error al modificar el registro.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Function StockARetirar(Articulo As Long) As Long
On Error GoTo ErrSAFR
    StockARetirar = 0
    Cons = "Select Sum(RenARetirar) From Documento, Renglon " _
        & " Where DocTipo IN (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ") " _
        & " And RenArticulo = " & Articulo _
        & " And RenDocumento = DocCodigo And DocAnulado = 0"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not IsNull(RsSt(0)) Then StockARetirar = RsSt(0)
    RsSt.Close
    'Tambien hay en Remitos.
    Cons = "Select Sum(RReAEntregar) From Remito, RenglonRemito " _
            & " Where RReArticulo = " & Articulo _
            & " And RReRemito = RemCodigo "
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not IsNull(RsSt(0)) Then StockARetirar = StockARetirar + RsSt(0)
    RsSt.Close
    Exit Function
ErrSAFR:
    msgError.MuestroError "Ocurrio un error al buscar el stock a retirar."
    StockARetirar = 0
End Function
Private Function StockAFacturarRetira(Articulo As Long) As Long
On Error GoTo ErrSAFR
    Cons = "Select Sum(RVTARetirar) From VentaTelefonica, RenglonVtaTelefonica " _
            & " Where VTeTipo = " & TipoDocumento.ContadoDomicilio _
            & " And VTeDocumento = Null And VTeAnulado = Null " _
            & " And RVTArticulo = " & Articulo _
            & "And VTeCodigo = RVTVentaTelefonica"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockAFacturarRetira = 0 Else StockAFacturarRetira = RsSt(0)
    RsSt.Close
    Exit Function
ErrSAFR:
    msgError.MuestroError "Ocurrio un error al buscar el stock a facturar que se retira."
    StockAFacturarRetira = 0
End Function
Private Function StockCamionEnvioEntregadoXEstado(Articulo As Long, Camion As Long, IDEstado As Long) As Long
On Error GoTo ErrSAFE
    Cons = "Select Sum(ReECantidadEntregada) From RenglonEntrega" _
        & " Where ReEArticulo = " & Articulo _
        & " And ReEEstado = " & IDEstado _
        & " And ReECodImpresion IN " _
        & "(Select REvCodImpresion From Envio, RenglonEnvio" _
            & " Where EnvTipo = " & TipoEnvio.Entrega _
            & " And EnvEstado = " & EstadoEnvio.Impreso _
            & " And EnvDocumento <> Null " _
            & " And EnvCamion = " & Camion _
            & " And REvArticulo = " & Articulo _
            & " And REvAEntregar > 0" _
            & "And EnvCodigo = REvEnvio)"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockCamionEnvioEntregadoXEstado = 0 Else StockCamionEnvioEntregadoXEstado = RsSt(0)
    RsSt.Close
    Exit Function
ErrSAFE:
    msgError.MuestroError "Ocurrio un error al buscar el stock en envío de camión."
    StockCamionEnvioEntregadoXEstado = 0
End Function
Private Function StockCamionEnvioEntregado(Articulo As Long, Camion As Long) As Long
On Error GoTo ErrSAFE
    Cons = "Select Sum(ReECantidadEntregada) From RenglonEntrega" _
        & " Where ReEArticulo = " & Articulo _
        & " And ReECodImpresion IN " _
        & "(Select REvCodImpresion From Envio, RenglonEnvio" _
            & " Where EnvTipo = " & TipoEnvio.Entrega _
            & " And EnvEstado = " & EstadoEnvio.Impreso _
            & " And EnvCamion = " & Camion _
            & " And EnvDocumento <> Null " _
            & " And REvArticulo = " & Articulo _
            & " And REvAEntregar > 0" _
            & "And EnvCodigo = REvEnvio)"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockCamionEnvioEntregado = 0 Else StockCamionEnvioEntregado = RsSt(0)
    
    RsSt.Close
    Exit Function
ErrSAFE:
    msgError.MuestroError "Ocurrio un error al buscar el stock en envío de camión."
    StockCamionEnvioEntregado = 0
End Function
Private Function StockCamionEnvio(Articulo As Long, Camion As Long) As Long
On Error GoTo ErrSAFE
    Cons = "Select Sum(REvAEntregar) From Envio, RenglonEnvio" _
        & " Where EnvTipo = " & TipoEnvio.Entrega _
        & " And EnvEstado = " & EstadoEnvio.Impreso _
        & " And EnvCamion = " & Camion _
        & " And EnvDocumento <> Null " _
        & " And REvArticulo = " & Articulo _
        & " And REvAEntregar > 0" _
        & "And EnvCodigo = REvEnvio"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockCamionEnvio = 0 Else StockCamionEnvio = RsSt(0)
    RsSt.Close
    Exit Function
ErrSAFE:
    msgError.MuestroError "Ocurrio un error al buscar el stock en envío de camión."
    StockCamionEnvio = 0
End Function
Private Function StockEnEnvio(Articulo As Long) As Long
On Error GoTo ErrSAFE
    Cons = "Select Sum(REvAEntregar) From Envio, RenglonEnvio" _
        & " Where EnvTipo = " & TipoEnvio.Entrega _
        & " And EnvEstado NOT IN (2,4,5)" _
        & " And EnvDocumento <> Null " _
        & " And REvArticulo = " & Articulo _
        & " And REvAEntregar <> 0" _
        & "And EnvCodigo = REvEnvio"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockEnEnvio = 0 Else StockEnEnvio = RsSt(0)
    RsSt.Close
    Exit Function
ErrSAFE:
    msgError.MuestroError "Ocurrio un error al buscar el stock en envío."
    StockEnEnvio = 0
End Function
Private Function StockEnTraslado(Articulo As Long, Camion As Long) As Long
On Error GoTo ErrSAFE
    Cons = "Select Sum(RTrPendiente) From Traspaso, RenglonTraspaso" _
        & " Where TraLocalIntermedio = " & Camion _
        & " And TraFechaEntregado = Null And TraFImpreso <> Null " _
        & " And RTrArticulo = " & Articulo _
        & "And TraCodigo = RTrTraspaso"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockEnTraslado = 0 Else StockEnTraslado = RsSt(0)
    RsSt.Close
    Exit Function
ErrSAFE:
    msgError.MuestroError "Ocurrio un error al buscar el stock de traslados."
    StockEnTraslado = 0
End Function
Private Function StockEnTrasladoXEstado(Articulo As Long, Camion As Long, IDEstado As Long) As Long
On Error GoTo ErrSAFE
    Cons = "Select Sum(RTrPendiente) From Traspaso, RenglonTraspaso" _
        & " Where TraLocalIntermedio = " & Camion _
        & " And TraFechaEntregado = Null And TraFImpreso <> Null " _
        & " And RTrArticulo = " & Articulo _
        & " And RTrEstado = " & IDEstado _
        & "And TraCodigo = RTrTraspaso"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockEnTrasladoXEstado = 0 Else StockEnTrasladoXEstado = RsSt(0)
    RsSt.Close
    
    Cons = "Select Sum(RTrPendiente) From Traspaso, RenglonTraspaso" _
        & " Where TraLocalIntermedio = " & Camion _
        & " And TraFechaEntregado <> Null And TraFImpreso <> Null " _
        & " And RTrArticulo = " & Articulo _
        & " And RTrEstado = " & IDEstado _
        & "And TraCodigo = RTrTraspaso"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not IsNull(RsSt(0)) Then StockEnTrasladoXEstado = StockEnTrasladoXEstado + RsSt(0)
    RsSt.Close
    Exit Function
ErrSAFE:
    msgError.MuestroError "Ocurrio un error al buscar el stock de traslados."
    StockEnTrasladoXEstado = 0
End Function
Private Function StockEnTrasladoPendiente(Articulo As Long, Camion As Long) As Long
On Error GoTo ErrSAFE
    Cons = "Select Sum(RTrPendiente) From Traspaso, RenglonTraspaso" _
        & " Where TraLocalIntermedio = " & Camion _
        & " And TraFechaEntregado <> Null And TraFImpreso <> Null " _
        & " And RTrArticulo = " & Articulo _
        & "And TraCodigo = RTrTraspaso"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockEnTrasladoPendiente = 0 Else StockEnTrasladoPendiente = RsSt(0)
    RsSt.Close
    Exit Function
ErrSAFE:
    msgError.MuestroError "Ocurrio un error al buscar el stock de traslados."
    StockEnTrasladoPendiente = 0
End Function
Private Function StockAFacturarEnvia(Articulo As Long) As Long
On Error GoTo ErrSAFE
    Cons = "Select Sum(REvAEntregar) From Envio, RenglonEnvio" _
        & " Where EnvTipo = " & TipoEnvio.Cobranza _
        & " And EnvEstado NOT IN (2,4,5)" _
        & " And EnvDocumento <> Null " _
        & " And REvArticulo = " & Articulo _
        & " And REvAEntregar > 0" _
        & "And EnvCodigo = REvEnvio"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockAFacturarEnvia = 0 Else StockAFacturarEnvia = RsSt(0)
    RsSt.Close
    Exit Function
ErrSAFE:
    msgError.MuestroError "Ocurrio un error al buscar el stock a facturar que se envía."
    StockAFacturarEnvia = 0
End Function

Private Function StockSano(Articulo As Long) As Long
On Error GoTo ErrSS
    Cons = " Select  StTCantidad From StockTotal" _
        & " Where StTTipoEstado = " & TipoEstadoMercaderia.Fisico & " And StTEstado = " & paEstadoArticuloEntrega & " And StTArticulo = " & Articulo
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsSt.EOF Then
        If IsNull(RsSt(0)) Then StockSano = 0 Else StockSano = RsSt(0)
    Else: StockSano = 0
    End If
    RsSt.Close
    Exit Function
ErrSS:
    msgError.MuestroError "Ocurrio un error al buscar el stock Sano."
    StockSano = 0
End Function

Private Sub LimpioCampos()
LimpioGrilla: tArticulo.Tag = "0"
End Sub

Private Function RetornoNombreEstado(IDEstado As Integer) As String
    Select Case IDEstado
        Case TipoMovimientoEstado.AEntregar: RetornoNombreEstado = "A Enviar"
        Case TipoMovimientoEstado.ARetirar: RetornoNombreEstado = "A Retirar"
        Case TipoMovimientoEstado.Reserva: RetornoNombreEstado = "Reservado"
    End Select
End Function
Private Sub AccionImprimir()
    
    If vsGrilla.Rows = 1 Then
        MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo errPrint
    Screen.MousePointer = 11
    
    With vsListado
        
        .Preview = True
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
    
        EncabezadoListado vsListado, "Arreglo de Stock", False
        
        .FileName = "Arreglo Stock"
        .FontSize = 8: .FontBold = False
        
        vsGrilla.ExtendLastCol = False
        .Paragraph = " Stock "
        .RenderControl = vsGrilla.hwnd
        vsGrilla.ExtendLastCol = True
        .EndDoc
        
        If Not .PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 0: Exit Sub
        
        .PrintDoc
    End With
    
    Screen.MousePointer = 0
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al realizar la impresión. " & Trim(Err.Description)
End Sub

