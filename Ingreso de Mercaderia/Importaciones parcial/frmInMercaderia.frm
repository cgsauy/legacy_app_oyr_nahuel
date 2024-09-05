VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.0#0"; "AACOMBO.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmInMercaderia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Mercadería"
   ClientHeight    =   5850
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInMercaderia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   2750
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Artículos arribados"
      ForeColor       =   &H00000080&
      Height          =   3615
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   5175
      Begin VB.TextBox tCantidad 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox tComentario 
         Height          =   285
         Left            =   120
         MaxLength       =   80
         TabIndex        =   17
         Top             =   2520
         Width           =   4935
      End
      Begin VB.TextBox tUsuario 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   21
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   60
         TabIndex        =   12
         Top             =   480
         Width           =   3975
      End
      Begin AACombo99.AACombo cLocal 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsArticulo 
         Height          =   1515
         Left            =   120
         TabIndex        =   15
         Top             =   760
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2672
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
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   4
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
         ExtendLastCol   =   -1  'True
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
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cantidad"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Co&mentarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario:"
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Local:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ficha de Arribo"
      ForeColor       =   &H00000080&
      Height          =   1335
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   5175
      Begin VB.CommandButton bLimpiar 
         Caption         =   "&Limpiar"
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox tFactura 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox tSerieF 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   7
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton bBuscar 
         Caption         =   "&Buscar..."
         Height          =   315
         Left            =   2520
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "A&rribó el:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lFactura 
         BackStyle       =   0  'Transparent
         Caption         =   "&Documento:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   5595
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   0
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
            Picture         =   "frmInMercaderia.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":10E2
            Key             =   "Alta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":13FC
            Key             =   "Baja"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuOpL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario "
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmInMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean
Dim Msg As String, aTexto As String
Dim gRemito As Long     'Para modfiicar/eliminar se carga es seleccionado de la lista

Dim aValor As Long

'Propiedades para el llamado desde la compra de mercadería-------------------
Dim gLlamado As Integer, gProveedor As Long
'-----------------------------------------------------------------------------------------

Private Sub bBuscar_Click()

    'Valido que los campos de busqueda esten cargados
    If Trim(tFecha.Text) <> "" And Not IsDate(tFecha.Text) Then
        MsgBox "La fecha ingresada no es correcta. Verifique.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Sub
    End If
    
    If Trim(tFecha.Text) = "" And cTipo.ListIndex = -1 And Val(tProveedor.Tag) = 0 And Trim(tSerieF.Text) = "" And Trim(tFactura.Text) = "" Then
        MsgBox "No se ingresaron filtros de búsqueda. Seleccione algún criterio y presione el botón buscar.", vbExclamation, "ATENCIÓN": Exit Sub
    End If
    
    Cons = "Select RCoCodigo, RCoCodigo Compra, RCoFecha Fecha, PMeNombre Proveedor, RCoSerie Serie, RCoNumero 'Número', RCoComentario Comentarios" _
            & " From RemitoCompra, ProveedorMercaderia" _
            & " Where RCoProveedor = PMeCodigo"
    
    If IsDate(tFecha.Text) Then Cons = Cons & " And RCoFecha = '" & Format(tFecha.Text, sqlFormatoFH) & "'"
    If cTipo.ListIndex <> -1 Then Cons = Cons & " And RCoTipo = " & cTipo.ItemData(cTipo.ListIndex)
    If Val(tProveedor.Tag) <> 0 Then Cons = Cons & " And RCoProveedor = " & Val(tProveedor.Tag)
    If Trim(tSerieF.Text) <> "" Then Cons = Cons & " And RCoSerie = '" & Trim(tSerieF.Text) & "'"
    If Trim(tFactura.Text) <> "" Then Cons = Cons & " And RCoNumero = " & Trim(tFactura.Text)
    
    Dim aLista As New clsListadeAyuda
    Dim aSeleccionado As Long
    aLista.ActivoListaAyuda Cons, False, miConexion.TextoConexion(logComercio), 8500
    
    aSeleccionado = aLista.ValorSeleccionado
    If aSeleccionado <> 0 Then gRemito = aSeleccionado Else gRemito = 0
    
    Set aLista = Nothing
    Screen.MousePointer = 0
    Me.Refresh
    
    If gRemito <> 0 Then CargoDatosBusqueda gRemito
        
End Sub

Private Sub CargoDatosBusqueda(Remito As Long)
    
    On Error GoTo errCargo
    Screen.MousePointer = 11
    Cons = "Select * from RemitoCompra, RemitoCompraRenglon, Articulo" _
           & " Where RCoCodigo = " & Remito _
           & " And RCoCodigo = RCRRemito" _
           & " And RCRArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        vsArticulo.Rows = 1
        tFecha.Text = Format(RsAux!RCoFecha, "d-Mmm-yyyy")
        BuscoCodigoEnCombo cTipo, RsAux!RCoTipo
        
        CargoDatosProveedor RsAux!RCoProveedor
        
        BuscoCodigoEnCombo cLocal, RsAux!RCoLocal
        If Not IsNull(RsAux!RCoSerie) Then tSerieF.Text = RsAux!RCoSerie Else: tSerieF.Text = ""
        tFactura.Text = RsAux!RCoNumero
        If Not IsNull(RsAux!RCoComentario) Then tComentario.Text = RsAux!RCoComentario Else: tComentario.Text = ""
        
        With vsArticulo
            Do While Not RsAux.EOF
                .AddItem Trim(RsAux!ArtNombre)
                aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                
                .Cell(flexcpText, .Rows - 1, 1) = RsAux!RCRCantidad
                RsAux.MoveNext
            Loop
        End With
        
        If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta Then
            Botones True, True, True, False, False, Toolbar1, Me
        Else
            Botones True, False, True, False, False, Toolbar1, Me
        End If
    End If
    RsAux.Close
    
    gRemito = Remito
    Screen.MousePointer = 0
    Exit Sub
    
errCargo:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al cargar los artículos. " & Trim(Err.Description)
End Sub


Private Sub bLimpiar_Click()
    tFecha.Text = ""
    cTipo.Text = ""
    tProveedor.Text = ""
    tSerieF.Text = "": tFactura.Text = ""
    Foco tFecha
    Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub cLocal_GotFocus()
    cLocal.SelStart = 0: cLocal.SelLength = Len(cLocal.Text)
End Sub

Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub cTipo_Change()
    If Not sNuevo And Not sModificar Then gRemito = 0
End Sub

Private Sub cTipo_Click()
    If Not sNuevo And Not sModificar Then gRemito = 0
End Sub

Private Sub cTipo_GotFocus()
    cTipo.SelStart = 0: cTipo.SelLength = Len(cTipo.Text)
End Sub

Private Sub cTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tProveedor
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrLoad
    ObtengoSeteoForm Me, Me.Left, Me.Top, Me.Width, Me.Height
    
    LimpioFicha
    InicializoGrilla
    sNuevo = False: sModificar = False
    
    CargoDocumentos         'Tipos de Documentos de Compra de Mercaderia
    
    'Cargo los LOCALES
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cLocal, ""

    DeshabilitoIngreso
    
    If Trim(Command()) <> "" Then
        gLlamado = TipoLlamado.IngresoNuevo
        AccionNuevo
        If Val(Command()) <> 0 Then
            gProveedor = Val(Command())
            CargoDatosProveedor gProveedor
        End If
    End If
    Exit Sub
    
ErrLoad:
    msgError.MuestroError "Ocurrió un error al inicializar el formulario.", Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next

    GuardoSeteoForm Me
    
    CierroConexion
    Set msgError = Nothing
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub Label1_Click()
    Foco cLocal
End Sub

Private Sub Label12_Click()
    Foco tProveedor
End Sub

Private Sub Label2_Click()
    Foco tFecha
End Sub
Private Sub Label3_Click()
    Foco tArticulo
End Sub
Private Sub Label4_Click()
    Foco tCantidad
End Sub

Private Sub Label5_Click()
    Foco cTipo
End Sub

Private Sub Label7_Click()
    Foco tComentario
End Sub
Private Sub Label8_Click()
    Foco tUsuario
End Sub

Private Sub lFactura_Click()
    Foco tSerieF
End Sub

Private Sub tUsuario_Change()
    tUsuario.Tag = ""
End Sub

Private Sub vsArticulo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If Col <> 1 Then Cancel = True: Exit Sub
    
End Sub

Private Sub vsArticulo_GotFocus()
    vsArticulo.Select vsArticulo.Row, 1
End Sub

Private Sub vsArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    With vsArticulo
    If .Rows = 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDelete
            If sModificar And Trim(.Cell(flexcpText, .Row, 2)) = "" Then MsgBox "No podrá eliminar los artículos arribados. Solo se permite agregar.", vbExclamation, "ATENCIÓN": Exit Sub
            If vsArticulo.Rows > 0 Then vsArticulo.RemoveItem .Row
    End Select
    
    End With
    
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub
Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub
Private Sub MnuVolver_Click()
    Unload Me
End Sub
Private Sub AccionNuevo()
    
    On Error Resume Next
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso
    LimpioFicha
    tFecha.Text = Format(Date, "d-Mmm-yyyy")
    If gLlamado = TipoLlamado.IngresoNuevo Then
        BuscoCodigoEnCombo cLocal, paLocalCompañia
    Else
        BuscoCodigoEnCombo cLocal, paCodigoDeSucursal
    End If
    sNuevo = True: gRemito = 0
    Foco tFecha
    
End Sub
Private Sub AccionGrabar()

    If ValidoDatos Then
        If MsgBox("Confirma almacenar los datos ingresados", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
            If sNuevo Then
                If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta Then GraboDatosImportacion Else GraboDatos
                
                If gLlamado = TipoLlamado.IngresoNuevo Then
                    If MsgBox("Desea volver al ingreso de facturas.", vbQuestion + vbYesNo, "SALIR") = vbYes Then Unload Me
                End If
            End If
            If sModificar Then GraboDatosModificaion
        End If
    End If
    
End Sub

Private Sub AccionCancelar()
    
    Screen.MousePointer = vbHourglass
    LimpioFicha
    DeshabilitoIngreso
    Botones True, False, False, False, False, Toolbar1, Me
    gRemito = 0
    sNuevo = False: sModificar = False
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub AccionEliminar()

    If gRemito = 0 Then
        MsgBox "Alguno de los datos del documento ha cambiado. Vuelva a realizar la selección de la lista de ayuda.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    If cTipo.ItemData(cTipo.ListIndex) <> TipoDocumento.CompraCarpeta Then EliminoDatosIngreso Else EliminoDatosImportacion
    
End Sub

Private Sub AccionModificar()
    On Error Resume Next
    If gRemito = 0 Then
        MsgBox "Alguno de los datos del documento ha cambiado. Vuelva a realizar la selección de la lista de ayuda.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    If cTipo.ItemData(cTipo.ListIndex) <> TipoDocumento.CompraCarpeta Then Exit Sub
    
    If MsgBox("Esta acción permite agregar artículos a un arribo de importaciones. " & Chr(vbKeyReturn) & "No peude cambiar las cantidades ya ingresadas." _
    & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Desea continuar.", vbQuestion + vbYesNo, "Agregar artículos.") = vbNo Then Exit Sub
    
    sModificar = True
    
    tFecha.Enabled = False: tFecha.BackColor = Inactivo
    cTipo.Enabled = False: cTipo.BackColor = Inactivo
    tProveedor.Enabled = False: tProveedor.BackColor = Inactivo
    tSerieF.Enabled = False: tSerieF.BackColor = Inactivo
    tFactura.Enabled = False: tFactura.BackColor = Inactivo
    
    tArticulo.Enabled = True: tArticulo.BackColor = Blanco
    tUsuario.Enabled = True: tUsuario.BackColor = Obligatorio
    tCantidad.Enabled = True: tCantidad.BackColor = Blanco
    
    cLocal.Enabled = True: cLocal.BackColor = Obligatorio
    
    vsArticulo.Editable = True
    bBuscar.Enabled = False: bLimpiar.Enabled = False
    
    Botones False, False, False, True, True, Toolbar1, Me

End Sub

Private Sub EliminoDatosImportacion()
Dim aUsuario As Long
Dim aTipo As Integer, aFolder As Integer
Dim aLocalAlta As Long

    Cons = "Select * from RemitoCompra Where RCoCodigo = " & gRemito
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not IsNull(RsAux!RCoTipoFolder) Then
        aTipo = RsAux!RCoTipoFolder: aFolder = RsAux!RCoIDFolder
    Else
        MsgBox "El arribo de la carpeta no tiene registrados los datos del embarque. No podrá eliminarlo." & Chr(vbKeyReturn) & "(*) No es compatible con nuevo link a importaciones.", vbExclamation, "ATENCIÓN"
        RsAux.Close: Exit Sub
    End If
    RsAux.Close
    
    'Verifico si los folders estan costeados-------------------------------------------------------------------------------------------
    If aTipo = Folder.cFSubCarpeta Then
        Cons = "Select * from Subcarpeta Where SubID = " & aFolder
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux!SubCosteada Then
            MsgBox "La subcarpeta está costeada. No podrá eliminar el arribo de mercadería.", vbExclamation, "Carpeta Costeada"
            RsAux.Close: Exit Sub
        End If
        RsAux.Close
    End If
    
    If aTipo = Folder.cFEmbarque Then
        Cons = "Select * from Embarque Where EmbID = " & aFolder
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux!EmbCosteado Then
            MsgBox "El embarque está costeado. No podrá eliminar el arribo de mercadería.", vbExclamation, "Carpeta Costeada"
            RsAux.Close: Exit Sub
        End If
        RsAux.Close
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------
    
    
    aTexto = "El sistema va a realizar los movimientos de stock para sacar la mercadería del local " & Trim(cLocal.Text) & Chr(vbKeyReturn)
    If aTipo = Folder.cFSubCarpeta Then
        aTexto = aTexto & "El sistema va a realizar los movimientos de stock para ingresar la mercadería a Zona Franca." & Chr(vbKeyReturn) & Chr(vbKeyReturn)
        aLocalAlta = paLocalZF
    End If
    If aTipo = Folder.cFEmbarque Then
        aTexto = aTexto & "El sistema va a realizar los movimientos de stock para ingresar la mercadería a Puerto." & Chr(vbKeyReturn) & Chr(vbKeyReturn)
        aLocalAlta = paLocalPuerto
    End If
    aTexto = aTexto & "Para eliminar el ingreso presione Aceptar"
    
    If MsgBox(aTexto, vbOKCancel + vbDefaultButton2, "ELIMINAR") = vbCancel Then Exit Sub
    
    'Llamo al registro del Suceso-------------------------------------------------------------
    frmInSuceso.pNombreSuceso = "Eliminación de Ingreso de Mercadería"
    frmInSuceso.Show vbModal, Me
    Me.Refresh
    aUsuario = frmInSuceso.pUsuario
    If aUsuario = 0 Then Screen.MousePointer = 0: Exit Sub
    
    On Error GoTo ErrGD
    FechaDelServidor
    
    cBase.BeginTrans            '--------------------------------------------------------------------------------
    On Error GoTo ErrResumo
    With vsArticulo
    For I = 1 To .Rows - 1
        'Doy Bajas al Local DESTINO--------------------------------------------------------------------------------------------------------
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1
        
        MarcoMovimientoStockFisico aUsuario, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1, cTipo.ItemData(cTipo.ListIndex)
        '----------------------------------------------------------------------------------------------------------------------------------------
        
        'Doy Altas al Local ORIGEN--------------------------------------------------------------------------------------------------------
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, aLocalAlta, _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1
        
        MarcoMovimientoStockFisico aUsuario, TipoLocal.Deposito, aLocalAlta, _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1, cTipo.ItemData(cTipo.ListIndex), gRemito
        '----------------------------------------------------------------------------------------------------------------------------------------
        
        'Elimino Remito de Compra y renglones
        Cons = "Delete RemitoCompraRenglon Where RCRRemito = " & gRemito
        cBase.Execute Cons
        
        Cons = "Delete RemitoCompra Where RCoCodigo = " & gRemito
        cBase.Execute Cons
            
    Next
    End With
    
    aTexto = Trim(tSerieF.Text) & Trim(tFactura.Text) & " - " & Trim(tProveedor.Text)
    RegistroSuceso gFechaServidor, TipoSuceso.AnulacionDeDocumentos, paCodigoDeTerminal, aUsuario, 0, _
                        Descripcion:="R/Compra " & aTexto, _
                        Defensa:=Trim(frmInSuceso.pDefensa)
    
    'Updateo el embarque o la subcarpeta (Fechas de arribo al local)
    If aTipo = Folder.cFSubCarpeta Then
        Cons = "Select * from Subcarpeta Where SubID = " & aFolder
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        RsAux!SubFArribo = Null
        RsAux!SubFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux.Update: RsAux.Close
    End If
    
    If aTipo = Folder.cFEmbarque Then
        Cons = "Select * from Embarque Where EmbID = " & aFolder
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        RsAux!EmbFLocal = Null
        RsAux!EmbFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux.Update: RsAux.Close
    End If
    
    cBase.CommitTrans         '------------------------------------------------------------------------------
    
    Botones True, False, False, False, False, Toolbar1, Me
    LimpioFicha
    Screen.MousePointer = 0
    Exit Sub

ErrGD:
    Screen.MousePointer = 0: msgError.MuestroError "Ocurrió un error al iniciar la transacción."
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    cBase.RollbackTrans
    Screen.MousePointer = 0: msgError.MuestroError Msg
End Sub

Private Sub EliminoDatosIngreso()
Dim aUsuario As Long

    aTexto = "El sistema va a realizar los movimientos de stock para sacar la mercadería del local " & Trim(cLocal.Text) & Chr(vbKeyReturn)
    aTexto = aTexto & "Si ud. sacó mercaderia del local Compañía deberá hacer un ingreso manual para reestablecer las cantidades del local Compañía." & Chr(vbKeyReturn) & Chr(vbKeyReturn)
    aTexto = aTexto & "Para eliminar el ingreso presione Aceptar"
    
    If MsgBox(aTexto, vbOKCancel + vbDefaultButton2, "ELIMINAR") = vbCancel Then Exit Sub
    
    'Llamo al registro del Suceso-------------------------------------------------------------
    frmInSuceso.pNombreSuceso = "Eliminación de Ingreso de Mercadería"
    frmInSuceso.Show vbModal, Me
    Me.Refresh
    aUsuario = frmInSuceso.pUsuario
    If aUsuario = 0 Then Screen.MousePointer = 0: Exit Sub
    
    On Error GoTo ErrGD
    FechaDelServidor
    
    cBase.BeginTrans            '--------------------------------------------------------------------------------
    On Error GoTo ErrResumo
    With vsArticulo
    For I = 1 To .Rows - 1
        'Doy Bajas al Local DESTINO--------------------------------------------------------------------------------------------------------
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1
        
        MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), -1
        
        MarcoMovimientoStockFisico aUsuario, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1, cTipo.ItemData(cTipo.ListIndex)
        '----------------------------------------------------------------------------------------------------------------------------------------
        
        'Elimino Remito de Compra y renglones
        Cons = "Delete RemitoCompraRenglon Where RCRRemito = " & gRemito
        cBase.Execute Cons
        
        Cons = "Delete RemitoCompra Where RCoCodigo = " & gRemito
        cBase.Execute Cons
    Next
    End With
    
    aTexto = Trim(tSerieF.Text) & Trim(tFactura.Text) & " - " & Trim(tProveedor.Text)
    RegistroSuceso gFechaServidor, TipoSuceso.AnulacionDeDocumentos, paCodigoDeTerminal, aUsuario, 0, _
                            Descripcion:="R/Compra " & aTexto, _
                            Defensa:=Trim(frmInSuceso.pDefensa)
    
    cBase.CommitTrans         '------------------------------------------------------------------------------
    
    Botones True, False, False, False, False, Toolbar1, Me
    LimpioFicha
    Screen.MousePointer = 0
    Exit Sub

ErrGD:
    Screen.MousePointer = 0: msgError.MuestroError "Ocurrió un error al iniciar la transacción."
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    cBase.RollbackTrans
    Screen.MousePointer = 0: msgError.MuestroError Msg
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = 0
End Sub

Private Sub tArticulo_GotFocus()
    tArticulo.SelStart = 0: tArticulo.SelLength = Len(tArticulo.Text)
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tArticulo.Text) <> vbNullString Then
            If Not IsNumeric(tArticulo.Text) Then BuscoArticulo Nombre:=tArticulo.Text Else: BuscoArticulo Codigo:=tArticulo.Text
            If tArticulo.Tag <> "0" Then Foco tCantidad
        Else
            vsArticulo.SetFocus
        End If
    End If

End Sub

Private Sub tCantidad_GotFocus()
    tCantidad.SelStart = 0: tCantidad.SelLength = Len(tCantidad.Text)
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        If Trim(tArticulo.Text) = "" Or tArticulo.Tag = "0" Then
            MsgBox "Debe ingresar el código del artículo.", vbExclamation, "ATENCIÓN"
            Foco tArticulo
            Exit Sub
        End If
        
        If Not IsNumeric(tCantidad.Text) Then
            MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tCantidad
            Exit Sub
        End If
        
        With vsArticulo
            For I = 1 To .Rows - 1
                If .Cell(flexcpData, I, 0) = tArticulo.Tag Then
                    MsgBox "El artículo ingresado ya está en la lista. Verifique.", vbExclamation, "ATENCIÓN": Exit Sub
                End If
            Next
        
            .AddItem Trim(tArticulo.Text)
            aValor = CLng(tArticulo.Tag): .Cell(flexcpData, .Rows - 1, 0) = aValor
        
            .Cell(flexcpText, .Rows - 1, 1) = tCantidad.Text
            If sModificar Then .Cell(flexcpText, .Rows - 1, 2) = "*"
            
        End With
            
        tArticulo.Text = ""
        tCantidad.Text = ""
        Foco tArticulo
    End If
    
End Sub

Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0: tComentario.SelLength = Len(tComentario.Text)
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If cLocal.Enabled Then Foco cLocal Else: Foco tUsuario
End Sub

Private Sub tFactura_GotFocus()
    tFactura.SelStart = 0: tFactura.SelLength = Len(tFactura.Text)
End Sub

Private Sub tFactura_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If sNuevo Then
            If cTipo.ListIndex = -1 Then MsgBox "Seleccione el tipo de documento para el ingreso de datos.", vbExclamation, "ATENCIÓN": Exit Sub
            
            If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta Then
                If Not IsNumeric(tFactura.Text) Then
                    If MsgBox("No se ingresó el número de carpeta." & Chr(vbKeyReturn) & "Desea ver todas las carpetas a arribar", vbQuestion + vbYesNo, "Carpetas Importación") = vbNo Then Exit Sub
                    BuscarImportacion
                Else
                    BuscarImportacion CLng(tFactura.Text)
                End If
            Else
                Foco tArticulo
            End If
        Else
            Foco tArticulo
        End If
    End If
End Sub

Private Sub BuscarImportacion(Optional Carpeta As Long = 0)
        
Dim aTxtSeleccionado As String
Dim aFolder As Long, aTipo As Integer

    'EN EL TAG DEL COMBO cTIPO GUARDO EL TIPO Y FOLDER SELECCIONADO  !!!!
    On Error GoTo errImportacion
    Cons = "Select Tipo = '" & Folder.cFEmbarque & "' + Convert(char(10), EmbID)," _
            & " CarCodigo 'Carpeta', EmbCodigo 'Embarque', Sub = '', EmbBL 'BL Embarque', EmbCarpetaDesp 'C.Despachante', CiuNombre 'Desde', LocNombre 'A....',  EmbFArribo 'Arribada'" _
            & " from Embarque, Carpeta, Ciudad, Local " _
            & " Where EmbLocal <> " & paLocalZF & " And EmbFLocal = Null " _
            & " And EmbCarpeta = CarID " _
            & " And EmbCiudadOrigen *= CiuCodigo " _
            & " And EmbLocal *= LocCodigo"
    If Carpeta <> 0 Then Cons = Cons & " And CarCodigo = " & Carpeta
    
    Cons = Cons & " Union All "
        
    Cons = Cons & "Select Tipo = '" & Folder.cFSubCarpeta & "' + Convert(char(10), SubID)," _
        & " CarCodigo 'Carpeta', EmbCodigo 'Embarque', Sub = SubCodigo, EmbBL 'BL Embarque', SubDespachante 'C.Despachante', Desde = 'Zona Franca', LocNombre 'A....', EmbFArribo 'Arribada' " _
        & " from SubCarpeta , Embarque, Carpeta, Local" _
        & " Where SubFArribo = Null" _
        & " And SubEmbarque = EmbID" _
        & " And EmbCarpeta = CarID" _
        & " And SubLocal *= LocCodigo"
    If Carpeta <> 0 Then Cons = Cons & " And CarCodigo = " & Carpeta
    
    Dim aLista As New clsListadeAyuda
    Dim aItem As String
    aLista.ActivoListaAyuda Cons, False, miConexion.TextoConexion(logComercio), 9000
    
    aTxtSeleccionado = CStr(aLista.ValorSeleccionado)
    aItem = aLista.ItemSeleccionado
    Set aLista = Nothing
    Screen.MousePointer = 0: Me.Refresh
    
    cTipo.Tag = ""
    If Val(aTxtSeleccionado) = 0 Then Exit Sub

    aTipo = Mid(aTxtSeleccionado, 1, 1): aFolder = Mid(aTxtSeleccionado, 2, Len(aTxtSeleccionado))
    cTipo.Tag = aTxtSeleccionado
    tFactura.Text = CLng(aItem)

    'Cargo los Articulos del Folder Seleccionado
    Cons = "Select * from ArticuloFolder, Articulo Where AFoTipo = " & aTipo & "And AFoCodigo = " & aFolder & " And AFoArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    With vsArticulo
        .Rows = 1
        Do While Not RsAux.EOF
            .AddItem Trim(RsAux!ArtNombre)
            aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            If sNuevo Then
                .Cell(flexcpText, .Rows - 1, 3) = RsAux!AFoCantidad         'Cantidad Original
            Else
                .Cell(flexcpText, .Rows - 1, 1) = RsAux!AFoCantidad
            End If
        
            RsAux.MoveNext
        Loop
    End With
    RsAux.Close
    
    If vsArticulo.Rows > 1 Then Foco tArticulo
    Screen.MousePointer = 0
    
    Exit Sub
errImportacion:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al procesar la información de importación", Err.Description
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0: tFecha.SelLength = Len(tFecha.Text)
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Not IsDate(tFecha.Text) And (sNuevo Or sModificar) Then
            MsgBox " La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tFecha
        Else
            Foco cTipo
        End If
    End If
    
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "d-Mmm-yyyy")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "grabar": AccionGrabar
        Case "modificar": AccionModificar
        Case "cancelar": AccionCancelar
        Case "eliminar": AccionEliminar
        Case "salir": Unload Me
    End Select

End Sub
Private Sub DeshabilitoIngreso()
    
    tFecha.Enabled = True: tFecha.BackColor = Blanco
    cTipo.Enabled = True: cTipo.BackColor = Blanco
    tProveedor.Enabled = True: tProveedor.BackColor = Blanco
    tSerieF.Enabled = True: tSerieF.BackColor = Blanco
    tFactura.Enabled = True: tFactura.BackColor = Blanco
    cLocal.Enabled = False: cLocal.BackColor = Inactivo
    tArticulo.Enabled = False: tArticulo.BackColor = Inactivo
    tUsuario.Enabled = False: tUsuario.BackColor = Inactivo
    tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
    tComentario.BackColor = Inactivo: tComentario.Enabled = False
    
    vsArticulo.Editable = False
    bBuscar.Enabled = True
    bLimpiar.Enabled = True
            
End Sub
Private Sub HabilitoIngreso()
    
    tFecha.Enabled = True: tFecha.BackColor = Obligatorio
    cTipo.Enabled = True: cTipo.BackColor = Obligatorio
    tProveedor.Enabled = True: tProveedor.BackColor = Obligatorio
    tSerieF.Enabled = True: tSerieF.BackColor = Blanco
    tFactura.Enabled = True: tFactura.BackColor = Obligatorio
    
    tArticulo.Enabled = True: tArticulo.BackColor = Blanco
    tUsuario.Enabled = True: tUsuario.BackColor = Obligatorio
    tCantidad.Enabled = True: tCantidad.BackColor = Blanco
    tComentario.BackColor = Blanco: tComentario.Enabled = True
    
    If gLlamado <> TipoLlamado.IngresoNuevo Then
        cLocal.Enabled = True
        cLocal.BackColor = Obligatorio
    End If
    
    vsArticulo.Editable = True
    bBuscar.Enabled = False
    bLimpiar.Enabled = False
End Sub

Private Sub LimpioFicha()
    
    tFecha.Text = ""
    cTipo.Text = ""
    tProveedor.Text = ""
    tSerieF.Text = ""
    tFactura.Text = ""
    cLocal.Text = ""
    tArticulo.Text = ""
    tUsuario.Text = ""
    tCantidad.Text = ""
    tComentario.Text = ""
    
    vsArticulo.Rows = 1
   
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = 0
End Sub

Private Sub tProveedor_GotFocus()
    tProveedor.SelStart = 0: tProveedor.SelLength = Len(tProveedor.Text)
End Sub

Private Sub tProveedor_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo errBuscar
    If KeyCode = vbKeyReturn Then
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then Foco tSerieF: Exit Sub
        
        'Busco el proveedor
        Screen.MousePointer = 11
        Cons = "Select PMeCodigo, PMeFantasia, PMeNombre from ProveedorMercaderia " _
                & " Where PMeNombre like '" & Trim(tProveedor.Text) & "%' Or PMeFantasia like '" & Trim(tProveedor.Text) & "%'"
        
        Dim aLista As New clsListadeAyuda
        aLista.ActivoListaAyuda Cons, False, miConexion.TextoConexion(logImportaciones), 5500
        If aLista.ValorSeleccionado <> 0 Then
            tProveedor.Text = Trim(aLista.ItemSeleccionado)
            tProveedor.Tag = aLista.ValorSeleccionado
            
            Me.Refresh
            HayStockCompañia
            Foco tSerieF
        Else
            tProveedor.Text = ""
        End If
        Set aLista = Nothing
        Screen.MousePointer = 0
    End If
    Exit Sub
    Screen.MousePointer = 0

errBuscar:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al procesar la lista de ayuda.", Err.Description
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) = 0 Then Exit Sub
        If vsArticulo.Rows > 1 Or (Not sNuevo And Not sModificar) Then Foco tSerieF: Exit Sub
        
        'IMPORTACION
        If cTipo.ListIndex <> -1 Then
            If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta Then Foco tSerieF: Exit Sub
        End If
        
        HayStockCompañia
        tSerieF.SetFocus
    End If
End Sub

Private Sub tSerieF_GotFocus()
    tSerieF.SelStart = 0: tSerieF.SelLength = Len(tSerieF.Text)
End Sub

Private Sub tSerieF_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then Foco tFactura
End Sub

Private Sub tUsuario_GotFocus()
    tUsuario.SelStart = 0: tUsuario.SelLength = Len(tUsuario.Text)
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    
    tUsuario.Tag = vbNullString
    If KeyAscii = vbKeyReturn And Trim(tUsuario.Text) <> "" Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = BuscoUsuarioDigito(CLng(tUsuario.Text), Codigo:=True)
            If Val(tUsuario.Tag) > 0 Then AccionGrabar: Exit Sub
            tUsuario.Tag = vbNullString
            MsgBox "No existe un usuario para el dígito ingresado", vbExclamation, "Dígito incorrecto."
        Else
            MsgBox "El formato dígito de usuario no es correcto.", vbExclamation, "ATENCIÓN"
            Foco tUsuario
        End If
    End If
    
End Sub

Private Function ValidoDatos() As Boolean

    ValidoDatos = False
    
    If vsArticulo.Rows = 1 Then
        MsgBox "No se han ingresado los artículos del documento.", vbExclamation, "ATENCIÓN"
        Foco tArticulo: Exit Function
    End If
    
    If tUsuario.Tag = "" Or tUsuario.Tag = "0" Then
        MsgBox "Debe ingresar eldígito  de usuario.", vbExclamation, "ATENCIÓN"
        Foco tUsuario: Exit Function
    End If
    
    If Val(tProveedor.Tag) = 0 Then
        MsgBox "Seleccione el proveedor de la mercadería.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Function
    End If
    
    If cLocal.ListIndex = -1 Then
        MsgBox "Seleccione el local para de ingreso de la mercadería.", vbExclamation, "ATENCIÓN"
        Foco cLocal: Exit Function
    End If
    
    If cTipo.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de documento ingresado.", vbExclamation, "ATENCIÓN"
        Foco cTipo: Exit Function
    End If
    
    If Trim(tFactura.Text) = "" Then
        MsgBox "Ingrese el número de documento asociado a la compra.", vbExclamation, "ATENCIÓN"
        Foco tFactura: Exit Function
    End If
    
    If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta And sNuevo Then
        If Val(cTipo.Tag) = 0 Then
            MsgBox "La carpeta de importación seleccionada no es correcta. Vuelva a cargar los datos.", vbExclamation, "ATENCIÓN"
            Foco tFactura: Exit Function
        End If
    End If
    
    With vsArticulo
        For I = 1 To .Rows - 1
            If .Cell(flexcpValue, I, 1) = 0 Or .Cell(flexcpText, I, 1) = "" Then
                MsgBox "Las cantidades arribadas no son correctas. Verifique la lista de arribo", vbExclamation, "ATENCIÓN"
                Exit Function
            End If
        Next
    End With
    
    ValidoDatos = True
    
End Function

Private Sub GraboDatos()

'   --> Para la mercadería que no figura en la compañia se inserta el remito y los renglones con esta mercadería
'   --> Para la mercadería que está en la compañia solamente se hace un traslado

Dim aCodigoRemito As Long
Dim aCodigoTraslado As Long

    Screen.MousePointer = 11
    
    FechaDelServidor
    On Error GoTo ErrGD
    cBase.BeginTrans            '--------------------------------------------------------------------------------
    On Error GoTo ErrResumo
    
    'Veo Si debo insertar en la tabla RemitoCompra (si hay mercadería q no pertence a la compañia)
    aCodigoRemito = GraboBDRemitoCompra
    aCodigoTraslado = GraboBDTraslado       'Si es q' hay mercaderia en la compañía para trasladar
    
    With vsArticulo
    For I = 1 To .Rows - 1
        'Doy Altas al Local DESTINO--------------------------------------------------------------------------------------------------------
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1
        
        MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), 1
        
        MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1, cTipo.ItemData(cTipo.ListIndex), aCodigoRemito
        '----------------------------------------------------------------------------------------------------------------------------------------
        
        If Trim(.Cell(flexcpText, I, 2)) <> "" Then            'Mercadería EN COMPAÑIA
            'Si la mercadería estaba en compañia, le doy la baja del LOCAL COMPAÑÍA---------------------------------------------------
            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paLocalCompañia, _
                CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1
            
            MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), -1
            
            MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paLocalCompañia, _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1, cTipo.ItemData(cTipo.ListIndex), aCodigoRemito
            '----------------------------------------------------------------------------------------------------------------------------------------
            
            'Tambien actualizo la cantidad "EnCompania" de la tabla RemitoCompraRenglon
            Cons = "Select * from RemitoCompraRenglon " _
                  & " Where RCRRemito = " & CLng(.Cell(flexcpText, I, 2)) _
                  & " And RCRArticulo = " & .Cell(flexcpData, I, 0)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            RsAux.Edit
            RsAux!RCREnCompania = RsAux!RCREnCompania - .Cell(flexcpValue, I, 1)
            RsAux.Update
            RsAux.Close
            '-------------------------------------------------------------------------------------------
            
            'Como Saque la mercaderia de la compañia realizo el Traslado de Mercadería
            Cons = "Insert Into RenglonTraspaso (RTrTraspaso, RTrArticulo, RTrEstado, RTrCantidad, RTrPendiente)" _
                & " Values (" _
                & aCodigoTraslado & ", " _
                & CLng(.Cell(flexcpData, I, 0)) & ", " _
                & paEstadoArticuloEntrega & ", " _
                & CCur(.Cell(flexcpValue, I, 1)) & ", " _
                & "0)"
            cBase.Execute (Cons)
            '-------------------------------------------------------------------------------------------
        
        Else            'LA MERCADERÍA NO ESTABA EN COMPAÑIA (Inserto tabla RenglonRemitoCompra)
            'Supuestamente ya inserte en la tabla RemitoCompra
            Cons = "Insert Into RemitoCompraRenglon (RCRRemito, RCRArticulo, RCRCantidad, RCREnCompania, RCRRemanente) " _
                    & " Values (" _
                    & aCodigoRemito & ", " _
                    & .Cell(flexcpData, I, 0) & ", " _
                    & CCur(.Cell(flexcpValue, I, 1)) & ", " _
                    & CCur(.Cell(flexcpValue, I, 1)) & ", " _
                    & CCur(.Cell(flexcpValue, I, 1)) & ")"
            cBase.Execute Cons
        End If
    Next
    End With
    
    cBase.CommitTrans           '------------------------------------------------------------------------------
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub

ErrGD:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al iniciar la transacción."
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    msgError.MuestroError Msg
    Exit Sub
End Sub
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Graba las bases de datos para las carpetas de importaciones y actualiza los datos en las capretas
'   (*) Si la mercadería está en el puerto hay que trasladarla del puerto al Local
'   (*) Si la mercadería está en el Zona hay que trasladarla de ZF al Local
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GraboDatosImportacion()

Dim aCodigoRemito As Long, aCodigoTraslado As Long
Dim aLocalBaja As Long      'Id de local para dar de baja el STOCK (Puerto o Zona Franca)
Dim aTipo As Integer, aFolder As Long
Dim RsFolder As rdoResultset
Dim aLineaSuceso As String

    aLocalBaja = 0: aLineaSuceso = ""
    Screen.MousePointer = 11
    On Error GoTo ErrGD
    
    'Busco los datos de los folders---------------------------------------------------------------------------------------------------
    aTipo = Mid(cTipo.Tag, 1, 1): aFolder = Mid(cTipo.Tag, 2, Len(cTipo.Tag))
    If aTipo = Folder.cFSubCarpeta Then
        If MsgBox("El sistema realizará un traslado de mercadería desde Zona Franca al local seleccionado." & Chr(vbKeyReturn) & "Tambien se actualizarán los datos en la subcarpeta." _
                & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Desea continuar con la operación.", vbQuestion + vbYesNo, "Arribo de Mercadería.") = vbNo Then Exit Sub
        aLocalBaja = paLocalZF
    End If
    If aTipo = Folder.cFEmbarque Then
        Cons = "Select * from Embarque Where EmbId = " & aFolder
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not IsNull(RsAux!EmbLocal) Then
            If RsAux!EmbLocal = paLocalPuerto Then
                If Not IsNull(RsAux!EmbFArribo) Then
                    If MsgBox("El sistema realizará un traslado de mercadería desde Puerto al local seleccionado." & Chr(vbKeyReturn) & "Tambien se actualizarán los datos en el embarque." _
                        & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Desea continuar con la operación.", vbQuestion + vbYesNo, "Arribo de Mercadería.") = vbNo Then RsAux.Close: Exit Sub
                    aLocalBaja = paLocalPuerto
                Else
                    If MsgBox("La mercadería estaba pendiente de arribo a Puerto. " & Chr(vbKeyReturn) & "El sistema actualizarán los datos en el embarque y se fijará como fecha de arribo a puerto hoy (no se realizarán traslados)." _
                        & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Desea continuar con la operación.", vbQuestion + vbYesNo, "Arribo de Mercadería.") = vbNo Then RsAux.Close: Exit Sub
                End If
            End If
        End If
        RsAux.Close
    End If
    '-------------------------------------------------------------------------------------------------------------------------------------
    
    FechaDelServidor
    cBase.BeginTrans            '--------------------------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo ErrResumo
    
    aCodigoRemito = GraboBDRemitoCompra(aTipo, aFolder)
    
    If aLocalBaja <> 0 Then     'Si hay que hacer un traslado desde ZF o Puerto
        Cons = "Insert Into Traspaso (TraFecha, TraLocalOrigen, TraLocalDestino, TraComentario, TraFechaEntregado, TraUsuarioInicial, TraUsuarioFinal) " _
               & " Values (" _
               & "'" & Format(gFechaServidor, sqlFormatoFH) & "', " _
               & aLocalBaja & ", " _
               & cLocal.ItemData(cLocal.ListIndex) & ", " _
               & "'Ingreso Mercadería: " & Trim(cTipo.Text) & " " & Trim(tSerieF.Text) & Trim(tFactura.Text) & "', " _
               & "'" & Format(gFechaServidor, sqlFormatoFH) & "', " & CLng(tUsuario.Tag) & ", " _
               & CLng(tUsuario.Tag) & ")"
        cBase.Execute (Cons)
        
        'Saco el código del insertado.--------------------------------------------
        Cons = "Select MAX(TraCodigo) From Traspaso"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        aCodigoTraslado = RsAux(0)
        RsAux.Close
    End If      '----------------------------------------------------------------------------------------------------
    
    With vsArticulo
    For I = 1 To .Rows - 1
        If .Cell(flexcpText, I, 3) <> "" Then
            If .Cell(flexcpValue, I, 3) <> .Cell(flexcpValue, I, 1) Then aLineaSuceso = aLineaSuceso & .Cell(flexcpText, I, 0) & " (" & .Cell(flexcpText, I, 1) & "/" & .Cell(flexcpText, I, 3) & "), "
        End If
        
        'Doy Altas al Local DESTINO--------------------------------------------------------------------------------------------------------
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1
        
        MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), 1
        
        MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1, cTipo.ItemData(cTipo.ListIndex), aCodigoRemito
        '----------------------------------------------------------------------------------------------------------------------------------------
        
        If aLocalBaja <> 0 Then            'Mercadería EN PUERTO O ZONA FRANCA
            'Si la mercadería estaba en compañia, le doy la baja del LOCAL COMPAÑÍA---------------------------------------------------
            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, aLocalBaja, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1
            
            MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), -1
            
            MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, aLocalBaja, _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1, cTipo.ItemData(cTipo.ListIndex), aCodigoRemito
            '----------------------------------------------------------------------------------------------------------------------------------------
                
            'Como Saque la mercaderia de la Puerto o ZF: realizo el Traslado de Mercadería
            Cons = "Insert Into RenglonTraspaso (RTrTraspaso, RTrArticulo, RTrEstado, RTrCantidad, RTrPendiente)" _
                & " Values (" _
                & aCodigoTraslado & ", " _
                & .Cell(flexcpData, I, 0) & ", " _
                & paEstadoArticuloEntrega & ", " _
                & .Cell(flexcpValue, I, 1) & ", " _
                & "0)"
            cBase.Execute (Cons)
            '-------------------------------------------------------------------------------------------
        End If
        
        'Supuestamente ya inserte en la tabla RemitoCompra
        Cons = "Insert Into RemitoCompraRenglon (RCRRemito, RCRArticulo, RCRCantidad, RCREnCompania, RCRRemanente) " _
                & " Values (" _
                & aCodigoRemito & ", " _
                & CLng(.Cell(flexcpData, I, 0)) & ", " _
                & CCur(.Cell(flexcpValue, I, 1)) & ", " _
                & 0 & ", " _
                & 0 & ")"
        cBase.Execute Cons
            
    Next
    End With
    
    'Updateo el embarque o la subcarpeta (Fechas de arribo al local)
    If aTipo = Folder.cFSubCarpeta Then
        Cons = "Select * from Subcarpeta Where SubID = " & aFolder
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        RsAux!SubFArribo = Format(gFechaServidor, sqlFormatoFH)
        RsAux!SubLocal = cLocal.ItemData(cLocal.ListIndex)
        RsAux!SubFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux.Update: RsAux.Close
    End If
    
    If aTipo = Folder.cFEmbarque Then
        Cons = "Select * from Embarque Where EmbID = " & aFolder
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        If IsNull(RsAux!EmbFArribo) Then RsAux!EmbFArribo = Format(gFechaServidor, sqlFormatoFH)
        RsAux!EmbFLocal = Format(gFechaServidor, sqlFormatoFH)
        RsAux!EmbFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux.Update: RsAux.Close
    End If
    
    If aLineaSuceso <> "" Then aLineaSuceso = Mid(aLineaSuceso, 1, Len(aLineaSuceso) - 2)
    If Trim(aLineaSuceso) <> "" Then
        RegistroSuceso gFechaServidor, TipoSuceso.DiferenciaDeArticulos, paCodigoDeTerminal, CLng(tUsuario.Tag), 0, _
                            Descripcion:="Carpeta " & tFactura.Text, _
                            Defensa:=Trim(aLineaSuceso)
    End If
    
    cBase.CommitTrans           '------------------------------------------------------------------------------
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub

ErrGD:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al iniciar la transacción."
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    cBase.RollbackTrans
    Screen.MousePointer = 0: msgError.MuestroError Msg
End Sub

Private Sub GraboDatosModificaion()
Dim aCodigoRemito As Long

    On Error GoTo ErrGD
    aCodigoRemito = gRemito
    
    FechaDelServidor
    cBase.BeginTrans            '--------------------------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo ErrResumo
    With vsArticulo
    For I = 1 To .Rows - 1
        If Trim(.Cell(flexcpText, I, 2)) <> "" Then        'Solo los que se agregaron
            'Doy Altas al Local DESTINO--------------------------------------------------------------------------------------------------------
            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
                CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1
            
            MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), 1
            
            MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
                CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1, cTipo.ItemData(cTipo.ListIndex), aCodigoRemito
            '----------------------------------------------------------------------------------------------------------------------------------------
            
            'Los agrego a las tablas de remitos de compra
            Cons = "Insert Into RemitoCompraRenglon (RCRRemito, RCRArticulo, RCRCantidad, RCREnCompania, RCRRemanente) " _
                    & " Values (" _
                    & aCodigoRemito & ", " _
                    & CLng(.Cell(flexcpData, I, 0)) & ", " _
                    & CCur(.Cell(flexcpValue, I, 1)) & ", " _
                    & 0 & ", " _
                    & 0 & ")"
            cBase.Execute Cons
        End If
    Next
    End With
    
    cBase.CommitTrans           '------------------------------------------------------------------------------
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub

ErrGD:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al iniciar la transacción."
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    cBase.RollbackTrans
    Screen.MousePointer = 0: msgError.MuestroError Msg
End Sub

Private Function GraboBDRemitoCompra(Optional TipoFolder As Integer = 0, Optional Folder As Long = 0) As Long

    'Veo Si debo insertar en la tabla RemitoCompra (si hay mercadería q no pertence a la compañia)
    GraboBDRemitoCompra = 0
    
    For I = 1 To vsArticulo.Rows - 1      'Recorro y se inserto grabo y Exit for
        If vsArticulo.Cell(flexcpText, I, 2) = "" Then
            
            Cons = "Insert into RemitoCompra (RCoProveedor, RCoTipo, RCoLocal, RCoSerie, RCoNumero, RCoFecha, RCoComentario, RCoTipoFolder, RCoIDFolder) " _
                    & " Values ( " _
                    & Val(tProveedor.Tag) & ", " _
                    & cTipo.ItemData(cTipo.ListIndex) & ", " _
                    & cLocal.ItemData(cLocal.ListIndex) & ", "
            If Trim(tSerieF.Text) <> "" Then Cons = Cons & "'" & Trim(tSerieF.Text) & "', " Else: Cons = Cons & "Null, "
            
            Cons = Cons & CLng(tFactura.Text) & ", " _
                & "'" & Format(tFecha.Text, sqlFormatoFH) & "', "
            
            If Trim(tComentario.Text) <> "" Then Cons = Cons & "'" & Trim(tComentario.Text) & "', " Else: Cons = Cons & "Null, "
            If TipoFolder = 0 Then Cons = Cons & " null, null)" Else Cons = Cons & TipoFolder & ", " & Folder & ")"
            
            cBase.Execute Cons
            
            'Saco el Máx. codigo de RemitoCompra---------------------------------------------
            Cons = "Select Max(RCoCodigo) from RemitoCompra"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            GraboBDRemitoCompra = RsAux(0)
            RsAux.Close
            '-----------------------------------------------------------------------------------------
            Exit For
        End If
    Next
    
End Function

Private Function GraboBDTraslado() As Integer

    GraboBDTraslado = 0
    For I = 1 To vsArticulo.Rows - 1
        If vsArticulo.Cell(flexcpText, I, 2) <> "" Then
    
            Cons = "Insert Into Traspaso (TraFecha, TraLocalOrigen, TraLocalDestino, TraComentario, TraFechaEntregado, TraUsuarioInicial, TraUsuarioFinal) " _
                & " Values (" _
                & "'" & Format(gFechaServidor, sqlFormatoFH) & "', " _
                & paLocalCompañia & ", " _
                & cLocal.ItemData(cLocal.ListIndex) & ", " _
                & "'Ingreso Mercadería: " & Trim(cTipo.Text) & " " & Trim(tSerieF.Text) & Trim(tFactura.Text) & "', " _
                & "'" & Format(gFechaServidor, sqlFormatoFH) & "', " & CLng(tUsuario.Tag) & ", " _
                & CLng(tUsuario.Tag) & ")"
            cBase.Execute (Cons)
            
            'Saco el código del insertado.--------------------------------------------
            Cons = "Select MAX(TraCodigo) From Traspaso"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            GraboBDTraslado = RsAux(0)
            RsAux.Close
            
            Exit For
        End If
    Next
    
End Function

Private Sub HayStockCompañia()

Dim sHay As Boolean

    On Error GoTo errConsultar
    sHay = False
    If Val(tProveedor.Tag) = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    Cons = "Select * from RemitoCompra, RemitoCompraRenglon" _
            & " Where RCoProveedor = " & Val(tProveedor.Tag) _
            & " And RCoLocal = " & paLocalCompañia _
            & " And RCoCodigo = RCRRemito" _
            & " And RCREnCompania > 0 "
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then sHay = True
    RsAux.Close
    
    If sHay Then
        frmDeStockCompañia.pProveedor = Val(tProveedor.Tag)
        frmDeStockCompañia.pProveedorNombre = Trim(tProveedor.Text)
        frmDeStockCompañia.pLista = vsArticulo
        frmDeStockCompañia.Show vbModal, Me
    End If
    Me.Refresh
    
    Screen.MousePointer = 0
    Exit Sub

errConsultar:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al consultar el stock en la compañia.", Err.Description
End Sub

Public Sub BuscoArticulo(Optional Codigo As String = "", Optional Nombre As String = "")

    On Error GoTo ErrBAC

    If Trim(Codigo) <> "" Then          'Articulo por Codigo--------------------------------------------------------
        Cons = "Select ArtID, ArtNombre From Articulo Where ArtCodigo = " & CLng(Codigo)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If RsAux.EOF Then
            MsgBox "No se encontró un artículo para el código ingresado.", vbExclamation, "ATENCIÓN"
            tArticulo.Text = "": tArticulo.Tag = 0
        Else
            tArticulo.Text = Trim(RsAux!ArtNombre)
            tArticulo.Tag = RsAux!ArtID
        End If
        RsAux.Close
        Exit Sub
    End If
    
    If Trim(Nombre) <> "" Then          'Articulo por Nombre------------------------------------------------------
        Cons = "Select ArtId, ArtCodigo, ArtNombre from Articulo" _
                & " Where ArtNombre LIKE '" & Nombre & "%'" _
                & " Order by ArtNombre"
        
        Dim aLista As New clsListadeAyuda
        Dim aSeleccionado As Long, aItem As String
        aLista.ActivoListaAyuda Cons, False, miConexion.TextoConexion(logComercio), 5500
        
        aSeleccionado = aLista.ValorSeleccionado
        aItem = aLista.ItemSeleccionado
        Set aLista = Nothing
        Screen.MousePointer = 0: Me.Refresh
        
        If aSeleccionado <> 0 Then
            tArticulo.Text = aItem
            BuscoArticulo Codigo:=aItem
        End If
        Exit Sub
    End If
    Exit Sub
    
ErrBAC:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al buscar el artículo.", Err.Description
End Sub

Private Sub CargoDocumentos()

    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraCarpeta)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraCarpeta
    
    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraCarta)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraCarta
    
    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.Compracontado)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.Compracontado
    
    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraCredito)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraCredito
    
    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaCredito)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraNotaCredito
    
    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaDevolucion)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraNotaDevolucion
    
    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraRemito)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraRemito
    
End Sub

Private Sub CargoDatosProveedor(Codigo As Long)
Dim Rs1 As rdoResultset

    'Cargo los datos del proveedor
    Cons = "Select * from ProveedorMercaderia Where PMeCodigo = " & Codigo
    Set Rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not Rs1.EOF Then
        tProveedor.Text = Trim(Rs1!PMeNombre)
        tProveedor.Tag = Codigo
    End If
    Rs1.Close
    
End Sub

Private Sub InicializoGrilla()

    On Error Resume Next
    With vsArticulo
        .Rows = 1: .Cols = 1
        .Editable = False
        .FormatString = "Artículo|>Cantidad|N. Remito|Q Original|"
        .WordWrap = False
        .ColHidden(2) = True: .ColHidden(3) = True
        .ColWidth(0) = 3600: .ColWidth(1) = 1000:
    End With
    
End Sub

Private Sub vsArticulo_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vsArticulo
    
    If Not IsNumeric(vsArticulo.EditText) Then
        Cancel = True: .EditText = .Cell(flexcpText, Row, 1)
        Exit Sub
    End If
        
    If sNuevo Then
        If .Cell(flexcpText, Row, 2) <> "" And .Cell(flexcpText, Row, 1) <> .EditText Then
            MsgBox "El artículo seleccionado pertence a un documento ya ingresado." & Chr(vbKeyReturn) & "No puede cambiar la cantidad en éste formulario.", vbInformation, "ATENCIÓN"
            Cancel = True: Exit Sub
        End If
    End If
    
    If sModificar Then
        If .Cell(flexcpText, Row, 2) = "" And .Cell(flexcpText, Row, 1) <> .EditText Then
            MsgBox "No podrá modificar las cantidades de los artículos arribados.", vbExclamation, "ATENCIÓN"
            Cancel = True: Exit Sub
        End If
    End If
    
    End With
    
End Sub
