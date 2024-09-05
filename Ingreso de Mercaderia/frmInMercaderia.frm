VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmInMercaderia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Mercadería"
   ClientHeight    =   7140
   ClientLeft      =   3915
   ClientTop       =   2985
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
   ScaleHeight     =   7140
   ScaleWidth      =   5430
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
      Height          =   4875
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
         Top             =   3780
         Width           =   4935
      End
      Begin VB.TextBox tUsuario 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   21
         Top             =   4320
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
         Width           =   3915
      End
      Begin AACombo99.AACombo cLocal 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   4320
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
         Height          =   2655
         Left            =   120
         TabIndex        =   15
         Top             =   825
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4683
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
      Begin VB.Label lSubEmbarques 
         Caption         =   " Ver Sub-Embarques "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1800
         MouseIcon       =   "frmInMercaderia.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   0
         Width           =   1470
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
         Top             =   3540
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario:"
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Local:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4080
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
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   855
      End
      Begin VB.Label lFactura 
         BackStyle       =   0  'Transparent
         Caption         =   "&Documento:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   300
         Width           =   495
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   6885
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
            Picture         =   "frmInMercaderia.frx":074C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":085E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":0970
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":0A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":0B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":0CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":0FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":10D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":13EC
            Key             =   "Alta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInMercaderia.frx":1706
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
Dim gTraslado As Long   'Para eliminar los trasalados que se hicieron de Compañia

Dim aValor As Long

'Propiedades para el llamado desde la compra de mercadería-------------------
Dim gLlamado As Integer, gProveedor As Long

Dim bGrabarSuceso As Boolean

Dim prmArticulosEmb As String       'Articulos del embarque, lo uso para checkear mov. stock al modificar imp.

Private Sub bBuscar_Click()

    'Valido que los campos de busqueda esten cargados
    Dim aFecha As String: aFecha = ""
    If Trim(tFecha.Text) <> "" Then
        
        If Not IsDate(tFecha.Text) Then
            aFecha = Trim(tFecha.Text)
            If (Mid(aFecha, 1, 1) = ">" Or Mid(aFecha, 1, 1) = "<") And IsDate(Mid(aFecha, 2)) Then
                aFecha = Mid(aFecha, 1, 1) & " '" & Format(Mid(aFecha, 2), sqlFormatoFH) & "'"
            Else
                MsgBox "La fecha ingresada no es correcta. Verifique.", vbExclamation, "Error de Ingreso"
                Foco tFecha: Exit Sub
            End If
        Else
            aFecha = "= '" & Format(tFecha.Text, sqlFormatoFH) & "'"
            aFecha = "Between '" & Format(tFecha.Text, "yyyy-mm-dd 00:00") & "' And '" & Format(tFecha.Text, "yyyy-mm-dd 23:59") & "'"
        End If
    End If
    
    If Trim(tFecha.Text) = "" And cTipo.ListIndex = -1 And Val(tProveedor.Tag) = 0 And Trim(tSerieF.Text) = "" And Trim(tFactura.Text) = "" Then
        MsgBox "No se ingresaron filtros de búsqueda. Seleccione algún criterio y presione el botón buscar.", vbExclamation, "ATENCIÓN": Exit Sub
    End If
    
    'Busco Remitos de Compras y Traslados de Mercaderia
'    cons = "Select 1 as IDTipo, TDoNombre as Tipo, RCoCodigo as Codigo, RCoFecha as Fecha, PMeNombre as Proveedor, RCoSerie as Serie, RCoNumero as 'Número', RCoComentario as Comentarios" _
        & " From RemitoCompra, ProveedorMercaderia, TipoDocumento" _
        & " Where RCoProveedor = PMeCodigo" & _
          " And RCoTipo *= TDoID"
    
    Cons = "Select 1 as IDTipo, TDoNombre as Tipo, RCoCodigo as Codigo, RCoFecha as Fecha, PMeNombre as Proveedor, RCoSerie as Serie, RCoNumero as 'Número', RCoComentario as Comentarios" & _
            " From (( RemitoCompra" & _
                " LEFT JOIN ProveedorMercaderia ON RemitoCompra.RCoProveedor = ProveedorMercaderia.PMeCodigo)" & _
                " LEFT JOIN TipoDocumento ON RemitoCompra.RCoTipo = TipoDocumento.TDoID)"
    
    Dim pstrAux As String
    If Trim(tFecha.Text) <> "" Then pstrAux = pstrAux & " And RCoFecha " & aFecha
    If cTipo.ListIndex <> -1 Then pstrAux = pstrAux & " And RCoTipo = " & cTipo.ItemData(cTipo.ListIndex)
    If Val(tProveedor.Tag) <> 0 Then pstrAux = pstrAux & " And RCoProveedor = " & Val(tProveedor.Tag)
    If Trim(tSerieF.Text) <> "" Then pstrAux = pstrAux & " And RCoSerie = '" & Trim(tSerieF.Text) & "'"
    If Trim(tFactura.Text) <> "" Then pstrAux = pstrAux & " And RCoNumero = " & Trim(tFactura.Text)
    
    If pstrAux <> "" Then
        pstrAux = " Where" & Mid(pstrAux, 5)
        Cons = Cons & pstrAux
    End If
    
    If Trim(tFecha.Text) <> "" Then
        Cons = Cons & " UNION ALL " & _
            "Select 2 as IDTipo, 'Traslado' as Tipo, TraCodigo as Codigo, TraFecha as Fecha, '' as Proveedor, '' as Serie, Null as 'Número', TraComentario as Comentarios" & _
            " From Traspaso " & _
            " Where TraLocalOrigen = " & paLocalCompañia & _
            " And TraAnulado Is NULL" & " And TraFecha " & aFecha
    End If
    
    Cons = Cons & " Order by Fecha DESC"
    
    
    Dim objLista As New clsListadeAyuda
    Dim plngCodigo  As Long, pbytTipo As Byte
        
    If objLista.ActivarAyuda(cBase, Cons, 8500, 1) <> 0 Then
        plngCodigo = objLista.RetornoDatoSeleccionado(2)
        pbytTipo = objLista.RetornoDatoSeleccionado(0)
    End If
    Set objLista = Nothing
    
    gRemito = 0: gTraslado = 0
    If plngCodigo <> 0 Then
        Call bLimpiar_Click
        
        Select Case pbytTipo
            Case 1: gRemito = plngCodigo
            Case 2: gTraslado = plngCodigo
        End Select
    End If
   
    Screen.MousePointer = 0
    Me.Refresh
    
    If gRemito <> 0 Then CargoDatosRemito gRemito
    If gTraslado <> 0 Then CargoDatosTraslado gTraslado
        
End Sub

Private Sub CargoDatosRemito(ByVal Remito As Long)
    
    On Error GoTo errCargo
    cTipo.Tag = ""
    
    Dim auxSeleccionado As Long
    auxSeleccionado = Remito
    Screen.MousePointer = 11
    Cons = "Select * from RemitoCompra, RemitoCompraRenglon, Articulo" _
           & " Where RCoCodigo = " & auxSeleccionado _
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
        If Not IsNull(RsAux!RCoComentario) Then tComentario.Text = Trim(RsAux!RCoComentario) Else: tComentario.Text = ""
        
        If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta Then
        
            Dim bCosteado As Boolean
            If Not IsNull(RsAux!RCoTipoFolder) And Not IsNull(RsAux!RCoIDFolder) Then
                Select Case RsAux!RCoTipoFolder
                    Case Folder.cFEmbarque: Cons = "Select EmbCosteado From Embarque Where EmbId = " & RsAux!RCoIDFolder
                    Case Folder.cFSubCarpeta: Cons = "Select SubCosteada From SubCarpeta Where SubId = " & RsAux!RCoIDFolder
                End Select
                Dim rsFol As rdoResultset
                bCosteado = False
                Set rsFol = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsFol.EOF Then If rsFol(0) Then bCosteado = True
                rsFol.Close
                
                cTipo.Tag = RsAux!RCoTipoFolder & RsAux!RCoIDFolder
                lSubEmbarques.Visible = True
            End If
            If bCosteado Then
                MsgBox "El embarque seleccionado ha sido costeado, no podrá modificar los datos del arribo.", vbExclamation, "Embarque Costeado"
                Botones True, False, False, False, False, Toolbar1, Me
            Else
                Botones True, True, True, False, False, Toolbar1, Me
            End If
        Else
            Botones True, False, True, False, False, Toolbar1, Me
        End If
    
        With vsArticulo
            Do While Not RsAux.EOF
                .AddItem Trim(RsAux!ArtNombre)
                aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                
                .Cell(flexcpText, .Rows - 1, 1) = RsAux!RCRCantidad
                RsAux.MoveNext
            Loop
        End With
    
    End If
    RsAux.Close
    
    gRemito = auxSeleccionado
    Screen.MousePointer = 0
    Exit Sub
    
errCargo:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos. " & Trim(Err.Description)
End Sub

Private Sub CargoDatosTraslado(ByVal IDTraslado As Long)
    
    On Error GoTo errCargo
    
    Screen.MousePointer = 11
    
    Cons = "Select TraFecha, TraLocalDestino, TraComentario, ArtID, ArtNombre, RTrCantidad " & _
        " From Traspaso, RenglonTraspaso, Articulo" & _
        " Where TRaCodigo = " & IDTraslado & _
        " And TRaCodigo = RTrTraspaso" & _
        " And RTrArticulo = ArtID"
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        
        Botones True, False, True, False, False, Toolbar1, Me
        
        vsArticulo.Rows = 1
        tFecha.Text = Format(RsAux!TraFecha, "d-Mmm-yyyy")
        'BuscoCodigoEnCombo cTipo, rsAux!RCoTipo
        cTipo.Text = "Traslado"
        tProveedor.Text = "Local Compañía"
        BuscoCodigoEnCombo cLocal, RsAux!TraLocalDestino
        If Not IsNull(RsAux!TraComentario) Then tComentario.Text = Trim(RsAux!TraComentario) Else: tComentario.Text = ""
        
    
        With vsArticulo
            Do While Not RsAux.EOF
                .AddItem Trim(RsAux!ArtNombre)
                aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                
                .Cell(flexcpText, .Rows - 1, 1) = RsAux!RTrCantidad
                RsAux.MoveNext
            Loop
        End With
    
    End If
    RsAux.Close
    
    gTraslado = IDTraslado
    Screen.MousePointer = 0
    Exit Sub
    
errCargo:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos. " & Trim(Err.Description)
End Sub

Private Sub bLimpiar_Click()
    tFecha.Text = ""
    cTipo.Text = ""
    cTipo.Tag = ""
    tProveedor.Text = ""
    tSerieF.Text = "": tFactura.Text = ""
    Foco tFecha
    Botones True, False, False, False, False, Toolbar1, Me
    vsArticulo.Rows = 1
    lSubEmbarques.Visible = False
    
End Sub

Private Sub cLocal_GotFocus()
    cLocal.SelStart = 0: cLocal.SelLength = Len(cLocal.Text)
End Sub

Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub cTipo_Change()
    If Not sNuevo And Not sModificar Then gRemito = 0: gTraslado = 0
End Sub

Private Sub cTipo_Click()
    If Not sNuevo And Not sModificar Then gRemito = 0: gTraslado = 0
End Sub

Private Sub cTipo_GotFocus()
    cTipo.SelStart = 0: cTipo.SelLength = Len(cTipo.Text)
End Sub

Private Sub cTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cTipo.ListIndex <> -1 Then
            If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraNotaCredito Or cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraNotaDevolucion Then
                MsgBox "Las notas de devolución de mercadería se deben ingresar por el sistema de administración.", vbExclamation, "Error de ingreso"
                cTipo.Text = "": Exit Sub
            End If
        End If
        Foco tProveedor
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrLoad
    'ObtengoSeteoForm Me, Me.Left, Me.Top  ', Me.Width, Me.Height
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
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
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next

    GuardoSeteoForm Me
    
    CierroConexion
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

Private Sub lSubEmbarques_Click()
On Error GoTo errVer

Dim mSQL As String
Dim mTipoF As Byte, mIDFolder As Long

    If cTipo.ItemData(cTipo.ListIndex) <> TipoDocumento.CompraCarpeta Then Exit Sub
    If Trim(cTipo.Tag) = "" Then Exit Sub
    Screen.MousePointer = 11
    
    mSQL = "Select MSFFecha as 'Arribo', ArtCodigo as 'Cód.', ArtNombre as 'Artículo', MSFCantidad as 'Q', SucAbreviacion as 'Local'" & _
                " From RemitoCompra, MovimientoStockfisico, Articulo, Sucursal " & _
                " Where RCoCodigo = " & gRemito & _
                " And MSFTipoDocumento = " & TipoDocumento.CompraCarpeta & _
                " And MSFDocumento = RCoCodigo" & _
                " And MSFArticulo = ArtID And MSFLocal = SucCodigo"

    Dim objLista As New clsListadeAyuda
    objLista.ActivoListaAyudaSQL cBase, mSQL
    Set objLista = Nothing
    Screen.MousePointer = 0
    Me.Refresh

    Exit Sub
errVer:
    clsGeneral.OcurrioError "Error al buscar los Sub-Embarques", Err.Description
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
    If Not sNuevo And Not sModificar Then Exit Sub
    
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
    sNuevo = True: gRemito = 0: gTraslado = 0
    Foco tFecha
    
End Sub
Private Sub AccionGrabar()

    bGrabarSuceso = False
    If Not ValidoDatos Then Exit Sub
        
    If MsgBox("Confirma almacenar los datos ingresados", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
        
    Dim aUsuario As Long, aDefensa As String
    aUsuario = -1
    
    'Llamo al registro del Suceso-------------------------------------------------------------
    If (bGrabarSuceso) Or (cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta And sModificar) Then
        Dim objSuceso As clsSuceso
        Set objSuceso = New clsSuceso
        aUsuario = 0
        objSuceso.TipoSuceso = TipoSuceso.DiferenciaDeArticulos
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "Modificar Arribo Importación", cBase
        Me.Refresh
        aUsuario = objSuceso.Usuario
        aDefensa = objSuceso.Defensa
        Set objSuceso = Nothing
        If aUsuario = 0 Then Screen.MousePointer = 0: Exit Sub
        
'        frmInSuceso.pNombreSuceso = "Diferencias en Arribo Importación"
'        frmInSuceso.Show vbModal, Me
'        Me.Refresh
'        aUsuario = frmInSuceso.pUsuario
'        aDefensa = Trim(frmInSuceso.pDefensa)
'        If aUsuario = 0 Then Screen.MousePointer = 0: Exit Sub
    
'        frmInSuceso.pNombreSuceso = "Modificar Arribo Importación"
'        frmInSuceso.Show vbModal, Me
'        Me.Refresh
'        aUsuario = frmInSuceso.pUsuario
'        aDefensa = Trim(frmInSuceso.pDefensa)
'        If aUsuario = 0 Then Screen.MousePointer = 0: Exit Sub
    End If
    If sNuevo Then
        If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta Then
            GraboDatosImportacion aUsuario, aDefensa
        Else
            GraboDatos
        End If
        
        If gLlamado = TipoLlamado.IngresoNuevo Then
            If MsgBox("Desea volver al ingreso de facturas.", vbQuestion + vbYesNo, "SALIR") = vbYes Then Unload Me
        End If
    End If
    If sModificar Then GraboDatosModificacion aUsuario, aDefensa
    
End Sub

Private Sub AccionCancelar()
    
    Screen.MousePointer = vbHourglass
    LimpioFicha
    DeshabilitoIngreso
    Botones True, False, False, False, False, Toolbar1, Me
    gRemito = 0: gTraslado = 0
    sNuevo = False: sModificar = False
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub AccionEliminar()

    If gRemito = 0 And gTraslado = 0 Then
        MsgBox "Alguno de los datos del documento ha cambiado. Vuelva a realizar la selección de la lista de ayuda.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If gTraslado <> 0 Then
        EliminoDatosTraslado
    Else
        If cTipo.ItemData(cTipo.ListIndex) <> TipoDocumento.CompraCarpeta Then
            EliminoDatosIngreso
        Else
            EliminoDatosImportacion
        End If
    End If
    
End Sub

Private Sub AccionModificar()
    On Error Resume Next
    If gRemito = 0 Then
        MsgBox "Alguno de los datos del documento ha cambiado. Vuelva a realizar la selección de la lista de ayuda.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    If cTipo.ItemData(cTipo.ListIndex) <> TipoDocumento.CompraCarpeta Then Exit Sub
    
    
    If MsgBox("Esta acción permite agregar artículos a un arribo de importaciones. " & vbCrLf & _
                    "No puede cambiar las cantidades ya ingresadas." & vbCrLf & vbCrLf & _
                    "¿Desea continuar?", vbQuestion + vbYesNo, "Agregar artículos.") = vbNo Then Exit Sub
    
    sModificar = True
    
    'tFecha.Enabled = False: tFecha.BackColor = Inactivo
    tFecha.Enabled = True: tFecha.BackColor = Colores.Obligatorio
    tFecha.Text = Format(Date, "dd/mm/yyyy")
    tComentario.Enabled = True: tComentario.BackColor = Colores.Blanco
    
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

    Foco tFecha
End Sub

Private Sub EliminoDatosImportacion()
Dim aUsuario As Long
Dim aTipo As Integer, aFolder As Integer
Dim aLocalAlta As Long

    aLocalAlta = 0
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
        If Not IsNull(RsAux!EmbLocal) Then
            If RsAux!EmbLocal <> paLocalPuerto And RsAux!EmbLocal <> paLocalZF Then aLocalAlta = -1
        End If
        If RsAux!EmbCosteado Then
            MsgBox "El embarque está costeado. No podrá eliminar el arribo de mercadería.", vbExclamation, "Carpeta Costeada"
            RsAux.Close: Exit Sub
        End If
        RsAux.Close
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------
    
    'Saco los artículo que están en el embarque     (para trasladar a Puerto), los demas desparecen
    Dim aArticulosEmb As String: aArticulosEmb = ""
    Cons = "Select * from ArticuloFolder " & _
                  " Where AFoTipo = " & aTipo & _
                  " And AFoCodigo = " & aFolder
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        aArticulosEmb = ","
        Do While Not RsAux.EOF
            aArticulosEmb = aArticulosEmb & RsAux!AFoArticulo & ","
            RsAux.MoveNext
        Loop
    End If
    RsAux.Close
    
    aTexto = "El sistema va a realizar los movimientos de stock para sacar la mercadería del local " & Trim(cLocal.Text) & Chr(vbKeyReturn)
    If aTipo = Folder.cFSubCarpeta And aLocalAlta = 0 Then
        aTexto = aTexto & "El sistema va a realizar los movimientos de stock para ingresar la mercadería a Zona Franca." & Chr(vbKeyReturn) & Chr(vbKeyReturn)
        aLocalAlta = paLocalZF
    End If
    If aTipo = Folder.cFEmbarque And aLocalAlta = 0 Then
        aTexto = aTexto & "El sistema va a realizar los movimientos de stock para ingresar la mercadería a Puerto." & Chr(vbKeyReturn) & Chr(vbKeyReturn)
        aLocalAlta = paLocalPuerto
    End If
    
    aTexto = aTexto & "(*) Solamente para la mercadería que está embarcada." & vbCrLf
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
    Dim bFueEmbarcado As Boolean
    
    cBase.BeginTrans            '--------------------------------------------------------------------------------
    On Error GoTo ErrResumo
    With vsArticulo
    For I = 1 To .Rows - 1
        
        bFueEmbarcado = InStr(aArticulosEmb, "," & .Cell(flexcpData, I, 0) & ",")
        
        'Doy Bajas al Local DESTINO--------------------------------------------------------------------------------------------------------
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1
        
        MarcoMovimientoStockFisico aUsuario, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
             CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1, cTipo.ItemData(cTipo.ListIndex), gRemito
            
        'Si no hay local de alta --> hay que bajarlos del stock TOTAL
        If aLocalAlta = -1 Or Not bFueEmbarcado Then
            MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), -1
        End If
        '----------------------------------------------------------------------------------------------------------------------------------------
        
        'Doy Altas al Local ORIGEN--------------------------------------------------------------------------------------------------------
        If aLocalAlta <> -1 And bFueEmbarcado Then
            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, aLocalAlta, _
                CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1
            
            MarcoMovimientoStockFisico aUsuario, TipoLocal.Deposito, aLocalAlta, _
                CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1, cTipo.ItemData(cTipo.ListIndex), gRemito
        
        '5/12/2007 Si es puerto le resto la cantidad pendiente
            If aLocalAlta = paLocalPuerto Then
                'Le resto al pendiente la cantidad.
                Cons = "UPDATE ArticuloFolder SET AFoPendiente = AFoPendiente + " & CCur(.Cell(flexcpValue, I, 1)) & _
                    " WHERE AFoTipo = 2 And AFoCodigo = " & aFolder & " And AFoArticulo = " & CLng(.Cell(flexcpData, I, 0))
                cBase.Execute (Cons)
            End If
        
        End If
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
        If aLocalAlta = -1 Then RsAux!EmbFArribo = Null
        RsAux!EmbFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux.Update: RsAux.Close
    End If
    
    cBase.CommitTrans         '------------------------------------------------------------------------------
    
    Botones True, False, False, False, False, Toolbar1, Me
    LimpioFicha
    Screen.MousePointer = 0
    Exit Sub

ErrGD:
    Screen.MousePointer = 0: clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción."
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    cBase.RollbackTrans
    Screen.MousePointer = 0: clsGeneral.OcurrioError Msg
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
    Screen.MousePointer = 0: clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción."
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    cBase.RollbackTrans
    Screen.MousePointer = 0: clsGeneral.OcurrioError Msg
End Sub

Private Sub EliminoDatosTraslado()
'Esta rutina debe hacer un Traslado desde el local a el local Compañía
'Debe agregar a los remitos de compañía la Q de artículos que se hacen en el traslado


    aTexto = "El sistema va a realizar los movimientos de stock para sacar la mercadería del local " & Trim(cLocal.Text) & " y pasarla al local Compañía." & Chr(vbKeyReturn)
    aTexto = aTexto & "Para eliminar el Traslado presione Aceptar"
    
    If MsgBox(aTexto, vbOKCancel + vbDefaultButton2 + vbInformation, "ELIMINAR") = vbCancel Then Exit Sub
    
    On Error GoTo ErrGD
    FechaDelServidor
    
    Dim pintTipoDoc As Integer, plngCodigoDoc As Long
    
    cBase.BeginTrans            '--------------------------------------------------------------------------------
    On Error GoTo ErrResumo
    
    plngCodigoDoc = GraboBDTraslado(DesdeCompania:=False)       'Trasladar Mercadería desde Local a Compañia
    pintTipoDoc = TipoDocumento.Traslados

    For I = 1 To vsArticulo.Rows - 1
    
    With vsArticulo
        'Doy Bajas al Local donde está la mercadería -------------------------------------------------------------------------------------------
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1
        
        MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), -1
        
        MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1, pintTipoDoc, plngCodigoDoc
        '----------------------------------------------------------------------------------------------------------------------------------------
        
        'Doy de alta la meracedería en LOCAL COMPAÑÍA---------------------------------------------------
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paLocalCompañia, _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1
        
        MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), 1
        
        MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, paLocalCompañia, _
        CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1, pintTipoDoc, plngCodigoDoc
        '----------------------------------------------------------------------------------------------------------------------------------------
            
        'Como Saque la mercaderia de la compañia realizo el Traslado de Mercadería
        Cons = "Insert Into RenglonTraspaso (RTrTraspaso, RTrArticulo, RTrEstado, RTrCantidad, RTrPendiente)" _
            & " Values (" _
            & plngCodigoDoc & ", " _
            & CLng(.Cell(flexcpData, I, 0)) & ", " _
            & paEstadoArticuloEntrega & ", " _
            & CCur(.Cell(flexcpValue, I, 1)) & ", " _
            & "0)"
        cBase.Execute (Cons)
        '-------------------------------------------------------------------------------------------
            
        'Busco un remito en compañía para subir la mercadería
        Dim pcurQ As Currency, pcurAux As Currency
        pcurQ = CCur(.Cell(flexcpValue, I, 1))
        
        '"And RCoProveedor = " & Val(tProveedor.Tag)
        Cons = "Select * from RemitoCompra, RemitoCompraRenglon" _
            & " Where RCoLocal = " & paLocalCompañia _
            & " And RCoCodigo = RCRRemito" _
            & " And RCREnCompania < RCRCantidad " _
            & " And RCRArticulo = " & CLng(.Cell(flexcpData, I, 0))
            
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            pcurAux = RsAux!RCRCantidad - RsAux!RCREnCompania
            If pcurAux > pcurQ Then pcurAux = pcurQ
            
            RsAux.Edit
            RsAux!RCREnCompania = RsAux!RCREnCompania + pcurAux
            RsAux.Update
            
            pcurQ = pcurQ - pcurAux
            If pcurQ = 0 Then Exit Do
            
            RsAux.MoveNext
        Loop
        RsAux.Close

    
    End With
    Next
    
    'Anulo el traslado Anterior
    Cons = "Update Traspaso Set TraAnulado = '" & Format(gFechaServidor, "yyyy-mm-dd hh:mm:ss") & "'" & _
           " Where TraCodigo = " & gTraslado
    cBase.Execute Cons
    
    cBase.CommitTrans         '------------------------------------------------------------------------------
    
    Botones True, False, False, False, False, Toolbar1, Me
    LimpioFicha
    Screen.MousePointer = 0
    Exit Sub

ErrGD:
    Screen.MousePointer = 0: clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción."
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    cBase.RollbackTrans
    Screen.MousePointer = 0: clsGeneral.OcurrioError Msg
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
            Foco tArticulo: Exit Sub
        End If
        
        If Not IsNumeric(tCantidad.Text) Then
            MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tCantidad: Exit Sub
        End If
        tCantidad.Text = Abs(tCantidad.Text)
        
        With vsArticulo
            For I = 1 To .Rows - 1
                If .Cell(flexcpData, I, 0) = Val(tArticulo.Tag) Then
                    If sModificar Then
                        If MsgBox("El artículo está ingresado en la lista y ya fue arribado. " & Chr(vbKeyReturn) & "Ud. desea agregar esta cantidad al arribo de mercadería.", vbQuestion + vbYesNo, "Artículo ya Arribado") = vbNo Then Exit Sub Else Exit For
                    Else
                        MsgBox "El artículo ingresado ya está en la lista. Verifique.", vbExclamation, "ATENCIÓN": Exit Sub
                    End If
                End If
            Next
        
            .AddItem Trim(tArticulo.Text)
            aValor = CLng(tArticulo.Tag): .Cell(flexcpData, .Rows - 1, 0) = aValor
        
            .Cell(flexcpText, .Rows - 1, 1) = tCantidad.Text
            If sModificar Then .Cell(flexcpText, .Rows - 1, 2) = "*"
        End With
            
        tArticulo.Text = "": tCantidad.Text = "": Foco tArticulo
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
    Screen.MousePointer = 11
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
    
    Dim objLista As New clsListadeAyuda
    Dim aItem As String
    
    If objLista.ActivarAyuda(cBase, Cons, 9000, 1) > 0 Then
        aTxtSeleccionado = CStr(objLista.RetornoDatoSeleccionado(0))
        aItem = objLista.RetornoDatoSeleccionado(1)
    End If
    Set objLista = Nothing
    
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
    clsGeneral.OcurrioError "Ocurrió un error al procesar la información de importación", Err.Description
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0: tFecha.SelLength = Len(tFecha.Text)
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = vbKeyReturn Then
        If Not IsDate(tFecha.Text) And (sNuevo Or sModificar) Then
            MsgBox " La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tFecha
        Else
            If cTipo.Enabled Then Foco cTipo Else Foco tArticulo
        End If
    End If
    
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "d-Mmm-yyyy")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComCtlLib.Button)
    
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
   
    lSubEmbarques.Visible = False
       
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
        
        'Busco el proveedor------------------------------------------------------------------------------------------------
        Dim aQ As Integer, aNombre As String, aId As Long
        aQ = 0
        Screen.MousePointer = 11
        Cons = "Select PMeCodigo, PMeFantasia as 'Nombre Fantasía', PMeNombre as 'Razón Social' from ProveedorMercaderia " _
                & " Where PMeNombre like '" & Trim(tProveedor.Text) & "%' Or PMeFantasia like '" & Trim(tProveedor.Text) & "%'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aQ = 1
            aNombre = Trim(RsAux(1)): aId = RsAux(0)
            RsAux.MoveNext: If Not RsAux.EOF Then aQ = 2
        End If
        RsAux.Close
        
        Select Case aQ
            Case 0: MsgBox "No existe un proveedor con el nombre ingresado.", vbInformation, "Proveedor Inexistente": Screen.MousePointer = 0: Exit Sub
            Case 1: tProveedor.Text = aNombre: tProveedor.Tag = aId
                        
            Case 2:
                Dim objLista As New clsListadeAyuda
                If objLista.ActivarAyuda(cBase, Cons, 5500, 1) > 0 Then
                    tProveedor.Text = Trim(objLista.RetornoDatoSeleccionado(1))
                    tProveedor.Tag = objLista.RetornoDatoSeleccionado(0)
                Else
                    tProveedor.Text = ""
                End If
                Set objLista = Nothing
                Me.Refresh
        End Select
        
        If Val(tProveedor.Tag) <> 0 Then
            If sNuevo Or sModificar Then HayStockCompañia
            Foco tSerieF
        End If
    End If
   
    Screen.MousePointer = 0
    Exit Sub
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar la lista de ayuda.", Err.Description
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) = 0 Then Exit Sub
        If vsArticulo.Rows > 1 Or (Not sNuevo And Not sModificar) Then Foco tSerieF: Exit Sub
        
        'IMPORTACION
        If cTipo.ListIndex <> -1 Then
            If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta Then Foco tSerieF: Exit Sub
        End If
        
        'HayStockCompañia
        'tSerieF.SetFocus
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
            tUsuario.Tag = z_BuscoUsuarioDigito(CLng(tUsuario.Text), Codigo:=True)
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
On Error GoTo errValido
    ValidoDatos = False
    Dim bTienePermiso As Boolean
    bTienePermiso = miConexion.AccesoAlMenu("Ingreso de MercaderiaE")
    
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
    If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraNotaCredito Or cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraNotaDevolucion Then
        MsgBox "Las notas de devolución de mercadería se deben ingresar por el sistema de administración.", vbExclamation, "Error de ingreso"
        Foco cTipo: Exit Function
    End If
    
    If Trim(tFactura.Text) = "" Then
        MsgBox "Ingrese el número de documento asociado a la compra.", vbExclamation, "ATENCIÓN"
        Foco tFactura: Exit Function
    End If
    
    With vsArticulo
        For I = 1 To .Rows - 1
            If .Cell(flexcpValue, I, 1) = 0 Or .Cell(flexcpText, I, 1) = "" Then
                MsgBox "Las cantidades arribadas no son correctas. Verifique la lista de arribo", vbExclamation, "ATENCIÓN"
                Exit Function
            End If
        Next
    End With
    
    If Not IsDate(tFecha.Text) Then
        MsgBox "Ingrese la fecha de arribo de la mercadería al local.", vbExclamation, "Faltan datos"
        Foco tFecha: Exit Function
    End If
    
        
    FechaDelServidor
    If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta Then
        If Format(tFecha.Text, "dd/mm/yyyy") <> Format(gFechaServidor, "dd/mm/yyyy") Then
            bGrabarSuceso = True
            If Abs(DateDiff("d", CDate(tFecha.Text), gFechaServidor)) > 5 Then
                MsgBox "Ud. no puede ingresar un arribo de importación com más de 5 días de atraso." & vbCrLf & _
                            "Consulte con importaciones.", vbExclamation, "Fecha de Arribo No Permitida."
                If Not bTienePermiso Then Exit Function
            End If
            
            If MsgBox("La fecha del Arribo no es la del día de hoy." & vbCrLf & vbCrLf & _
                            "Este embarque arribó hoy.", vbQuestion + vbYesNo, "Arribo del Embarque") = vbYes Then
                    MsgBox "Si el embarque arribó hoy, ingrese la fecha del día de hoy....", vbInformation, "Porque No Ingresa la Fecha de Hoy ?"
                    Exit Function
            End If
        End If
    End If
    
    'Nuevo e Importacion------------------------------------------------------------------------------------------------------------
    If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta And sNuevo Then
        If Val(cTipo.Tag) = 0 Then
            MsgBox "La carpeta de importación seleccionada no es correcta. Vuelva a cargar los datos.", vbExclamation, "ATENCIÓN"
            Foco tFactura: Exit Function
        End If
        
        'Controlo  Q original con la ingresada
        'If sNuevo Then .Cell(flexcpText, .Rows - 1, 3) = RsAux!AFoCantidad         'Cantidad Original
        'Else .Cell(flexcpText, .Rows - 1, 1) = RsAux!AFoCantidad
        
        With vsArticulo
            For I = 1 To .Rows - 1
                If Val(.Cell(flexcpText, I, 1)) <> Val(.Cell(flexcpText, I, 3)) Then
                    bGrabarSuceso = True
                    
                    If MsgBox("La cantidad que ud. ingresó no es la que se esperaba." & vbCrLf & vbCrLf & _
                                "Está seguro que contó " & .Cell(flexcpText, I, 1) & " " & .Cell(flexcpText, I, 0) & vbCrLf & vbCrLf & _
                                "Si está seguro presione si.", vbExclamation + vbYesNo + vbDefaultButton2, .Cell(flexcpText, I, 1) & " " & .Cell(flexcpText, I, 0) & " ?") = vbNo Then
                            Exit Function
                    End If
                    'Si la cantidad es mayor controlo que no se exeda en 50
                    If (.Cell(flexcpText, I, 1) > .Cell(flexcpValue, I, 3)) And (.Cell(flexcpText, I, 1) - .Cell(flexcpValue, I, 3) > 50) Then
                        MsgBox "La diferencia entre lo contado y la cantidad esperada supera el límite admitido." & vbCrLf & vbCrLf & _
                                    "Comuníquese con Importaciones para solucionar el problema.", vbExclamation, .Cell(flexcpText, I, 1) & " " & .Cell(flexcpText, I, 0) & " ?"
                        Exit Function       'Hay que salir si o si no se pueden ingresar componentes de pique
                    End If
                End If
            Next
        End With
        
    End If
    
    'Lo puse el 15/5/2002
    If sModificar And cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta Then
        prmArticulosEmb = ""
        'Valido que se hayan ingresado todos los articulos del embarque
        Dim aTipo As Integer, aFolder As Long, bSalir As Boolean, bArticuloAgregado As Boolean
        Dim aQEmb As Long
        bSalir = False
        
        aTipo = Mid(cTipo.Tag, 1, 1): aFolder = Mid(cTipo.Tag, 2, Len(cTipo.Tag))
        aTexto = ""
        Cons = "Select * from ArticuloFolder, Articulo " & _
                  " Where AFoTipo = " & aTipo & _
                  " And AFoCodigo = " & aFolder & _
                  " And AFoArticulo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF And Not bSalir
            aQEmb = 0: bArticuloAgregado = False
            
            With vsArticulo
                For I = 1 To .Rows - 1
                    If .Cell(flexcpData, I, 0) = RsAux!AFoArticulo Then
                        If Trim(.Cell(flexcpText, I, 2)) <> "" Then bArticuloAgregado = True
                        aQEmb = aQEmb + .Cell(flexcpValue, I, 1)
                    End If
                Next
                    
                If bArticuloAgregado Then prmArticulosEmb = prmArticulosEmb & RsAux!AFoArticulo & ","
                
                'Si la cantidad es mayor controlo que no se exeda en 50
                If (aQEmb > RsAux!AFoCantidad) And (aQEmb - RsAux!AFoCantidad > 50) And bArticuloAgregado Then
                    MsgBox "La diferencia entre lo contado y la cantidad esperada supera el límite admitido." & vbCrLf & vbCrLf & _
                                "Comuníquese con Importaciones para solucionar el problema.", vbExclamation, Trim(RsAux!ArtNombre) & " ?"
                    If Not bTienePermiso Then bSalir = True 'Exit Function
                End If
            End With
            RsAux.MoveNext
            
        Loop
        RsAux.Close
        If Trim(prmArticulosEmb) <> "" Then
            If Right(prmArticulosEmb, 1) = "," Then prmArticulosEmb = Mid(prmArticulosEmb, 1, Len(prmArticulosEmb) - 1)
        End If
        
        If bSalir Then Exit Function
    End If
    '-------------------------------------------------------------------------------------------------------------------------------------
        
    ValidoDatos = True
    Exit Function
    
errValido:
    clsGeneral.OcurrioError "Error al validar los datos.", Err.Description
    
End Function

Private Sub GraboDatos()

'   --> Para la mercadería que no figura en la compañia se inserta el remito y los renglones con esta mercadería
'   --> Para la mercadería que está en la compañia solamente se hace un traslado

Dim aCodigoRemito As Long
Dim aCodigoTraslado As Long

Dim pintTipoDoc As Integer, plngCodigoDoc As Long

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
        
        If Trim(.Cell(flexcpText, I, 2)) <> "" Then            'Mercadería EN COMPAÑIA
            pintTipoDoc = TipoDocumento.Traslados
            plngCodigoDoc = aCodigoTraslado
        Else
            pintTipoDoc = cTipo.ItemData(cTipo.ListIndex)
            plngCodigoDoc = aCodigoRemito
        End If
    
        'Doy Altas al Local DESTINO--------------------------------------------------------------------------------------------------------
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1
        
        MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), 1
        
        MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1, pintTipoDoc, plngCodigoDoc
        '----------------------------------------------------------------------------------------------------------------------------------------
        
        If Trim(.Cell(flexcpText, I, 2)) <> "" Then            'Mercadería EN COMPAÑIA
            'Si la mercadería estaba en compañia, le doy la baja del LOCAL COMPAÑÍA---------------------------------------------------
            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paLocalCompañia, _
                CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1
            
            MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), -1
            
            MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paLocalCompañia, _
            CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, -1, pintTipoDoc, plngCodigoDoc
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
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción."
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    clsGeneral.OcurrioError Msg
    Exit Sub
End Sub
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Graba las bases de datos para las carpetas de importaciones y actualiza los datos en las capretas
'   (*) Si la mercadería está en el puerto hay que trasladarla del puerto al Local
'   (*) Si la mercadería está en el Zona hay que trasladarla de ZF al Local
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GraboDatosImportacion(Optional sucUsuario As Long = -1, Optional sucDefensa As String = "")

Dim aCodigoRemito As Long, aCodigoTraslado As Long
Dim aLocalBaja As Long      'Id de local para dar de baja el STOCK (Puerto o Zona Franca)
Dim aTipo As Integer, aFolder As Long
Dim RsFolder As rdoResultset
Dim aLineaSuceso As String
Dim darAltasAPuerto As Boolean

    darAltasAPuerto = False
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
                    'If MsgBox("La mercadería estaba pendiente de arribo a Puerto. " & Chr(vbKeyReturn) & _
                                   "El sistema actualizará los datos en el embarque y se fijará como fecha de arribo a puerto hoy (no se realizarán traslados)." & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                                   "Desea continuar con la operación.", vbQuestion + vbYesNo, "Arribo de Mercadería.") = vbNo Then rsAux.Close: Exit Sub
                    If MsgBox("La mercadería estaba pendiente de arribo a Puerto. " & Chr(vbKeyReturn) & _
                                   "El sistema actualizará los datos en el embarque y se fijará como fecha de arribo a puerto hoy." & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                                   "Se va a realizar el Arribo a Puerto y el Traslado de la Mercadería al local" & Chr(vbKeyReturn) & _
                                   "¿Desea continuar con la operación?", vbQuestion + vbYesNo, "Arribo de Mercadería.") = vbNo Then RsAux.Close: Exit Sub
                    darAltasAPuerto = True
                    aLocalBaja = paLocalPuerto
                End If
            End If
        End If
        RsAux.Close
    End If
    '-------------------------------------------------------------------------------------------------------------------------------------
    
    FechaDelServidor
    ' *** Para hacer los movimientos con la fecha que ingresaron en el comprobante  ****
    gFechaServidor = Format(tFecha.Text, "dd/mm/yyyy") & " " & Format(gFechaServidor, "hh:mm:ss")
    
    cBase.BeginTrans            '--------------------------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo ErrResumo
    
    If darAltasAPuerto Then     '01/09/2006 - Cambio para arribar a puerto cunado no está la mercadería !!
        With vsArticulo
            For I = 1 To .Rows - 1
                 'Si es Nuevo la Q orig, va en la 3 ... si no en la 1
                 Dim mQOrigen As Currency
                 mQOrigen = CCur(.Cell(flexcpValue, I, 3))
                 If mQOrigen = 0 And CCur(.Cell(flexcpValue, I, 1)) <> 0 Then mQOrigen = CCur(.Cell(flexcpValue, I, 1))
                'Doy Altas al Local DESTINO PUERTO--------------------------------------------------------------------------------------------------------
                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, aLocalBaja, _
                     CLng(.Cell(flexcpData, I, 0)), mQOrigen, paEstadoArticuloEntrega, 1
                
                MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, mQOrigen, 1
                
                MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, aLocalBaja, _
                    CLng(.Cell(flexcpData, I, 0)), mQOrigen, paEstadoArticuloEntrega, 1, cTipo.ItemData(cTipo.ListIndex), aCodigoRemito
                '----------------------------------------------------------------------------------------------------------------------------------------
                
                
                '5/12/2007 --- updateo articulofolder y le pongo la cantidad pendiente al artículo.
                ' Si el embarque tiene 50 y me ponen 20 lo pendiente es = 20 ya que el alta de stock a puerto es de 20
                Cons = "UPDATE ArticuloFolder SET AFoPendiente = " & mQOrigen & _
                    " WHERE AFoTipo = 2 And AFoCodigo = " & aFolder & " And AFoArticulo = " & CLng(.Cell(flexcpData, I, 0))
                cBase.Execute (Cons)
                '5/12/2007 ---------------------------------------
            Next
        End With
    End If                              '01/09/2006 - Hasta acá
    
    
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
            
            '5/12/2007 Si es puerto le resto la cantidad pendiente
            If aLocalBaja = paLocalPuerto Then
                Dim RsF As rdoResultset
                'Le resto al pendiente la cantidad.
                Cons = "Select * From ArticuloFolder " & _
                     " WHERE AFoTipo = 2 And AFoCodigo = " & aFolder & " And AFoArticulo = " & CLng(.Cell(flexcpData, I, 0))
                Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsF.EOF Then
                    RsF.Edit
                    If Not IsNull(RsF("AFoPendiente")) Then
                        RsF("AFoPendiente") = RsF("AFoPendiente") - CCur(.Cell(flexcpValue, I, 1))
                    Else
                        RsF("AFoPendiente") = RsF("AFoCantidad") - CCur(.Cell(flexcpValue, I, 1))
                    End If
                    RsF.Update
                End If
                RsF.Close
            End If
            '5/12/07----------------------------------------------

            
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
    
    FechaDelServidor        'Reestablezco la fecha para registrar los sucesos *****
    
    If aLineaSuceso <> "" Then aLineaSuceso = Mid(aLineaSuceso, 1, Len(aLineaSuceso) - 2)
    If Trim(aLineaSuceso) <> "" Then
        RegistroSuceso gFechaServidor, TipoSuceso.DiferenciaDeArticulos, paCodigoDeTerminal, CLng(tUsuario.Tag), 0, _
                            Descripcion:="Carpeta " & tFactura.Text, _
                            Defensa:=Trim(aLineaSuceso)
    End If
    
    If sucUsuario > 0 Then      'Suceso con defensa     22/2/01
        RegistroSuceso gFechaServidor, TipoSuceso.DiferenciaDeArticulos, paCodigoDeTerminal, CLng(tUsuario.Tag), 0, _
                            Descripcion:="Carpeta " & tFactura.Text, _
                            Defensa:=Trim(sucDefensa)
    End If
   
    cBase.CommitTrans           '------------------------------------------------------------------------------
        
    On Error Resume Next
    If sucUsuario > 0 Then      'Envio mensaje 22/2/01
        Dim aMsg As String
        aMsg = "Carpeta " & Trim(tFactura.Text)
        If sNuevo Then aMsg = aMsg & "  (Nuevo Arribo)" Else aMsg = aMsg & "  (Modifica Arribo) "
        aMsg = aMsg & vbCrLf
        aMsg = aMsg & "Arribó el " & Trim(tFecha.Text) & " al local " & Trim(cLocal.Text) & vbCrLf
        If Trim(tComentario.Text) <> "" Then aMsg = aMsg & "Comentarios: " & Trim(tComentario.Text) & vbCrLf
        aMsg = aMsg & vbCrLf
        
        With vsArticulo
        For I = 1 To .Rows - 1
            aMsg = aMsg & .Cell(flexcpText, I, 0) & vbTab
            
            'Si es Nuevo la Q orig, va en la 3 ... si no en la 1
            If sNuevo Then
                aMsg = aMsg & "Q.Ing: " & .Cell(flexcpText, I, 1) & vbTab & " Q.Imp: " & .Cell(flexcpText, I, 3) & vbTab & "Q.Dif: " & .Cell(flexcpValue, I, 1) - .Cell(flexcpValue, I, 3) & vbCrLf
            
            End If
        Next
        End With
        aMsg = aMsg & vbCrLf & "Ingresado por el usuario " & miConexion.UsuarioLogueado(Nombre:=True)
        aMsg = aMsg & vbCrLf & "Defensa al suceso: " & Trim(sucDefensa)
        miConexion.EnviaMensaje prmMenUsuarioImportacion, "Diferencias Arribo Importación " & Trim(tFactura.Text), aMsg, DateAdd("s", 10, gFechaServidor), 0, prmMenUsuarioSistema
    End If
    
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub

ErrGD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción."
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    cBase.RollbackTrans
    Screen.MousePointer = 0: clsGeneral.OcurrioError Msg
End Sub

Private Sub GraboDatosModificacion(Optional sucUsuario As Long = -1, Optional sucDefensa As String = "")
Dim aCodigoRemito As Long
Dim rsRem As rdoResultset


Dim aCodigoTraslado As Long
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
        If MsgBox("El sistema realizará un traslado de mercadería desde Zona Franca al local seleccionado." & Chr(vbKeyReturn) & _
                        "Tambien se actualizarán los datos en la subcarpeta." & vbCrLf & _
                        "(*) Sólo para la mercadería embarcada." & vbCrLf & vbCrLf & _
                        "Desea continuar con la operación.", vbQuestion + vbYesNo, "Arribo de Mercadería (complemento)") = vbNo Then Exit Sub
        aLocalBaja = paLocalZF
    End If
    If aTipo = Folder.cFEmbarque Then
        Cons = "Select * from Embarque Where EmbId = " & aFolder
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not IsNull(RsAux!EmbLocal) Then
            If RsAux!EmbLocal = paLocalPuerto Then
                If Not IsNull(RsAux!EmbFArribo) Then
                    If MsgBox("El sistema realizará un traslado de mercadería desde Puerto al local seleccionado." & Chr(vbKeyReturn) & "Tambien se actualizarán los datos en el embarque." & vbCrLf & _
                                   "(*) Sólo para la mercadería embarcada." & vbCrLf & vbCrLf & _
                                   "Desea continuar con la operación.", vbQuestion + vbYesNo, "Arribo de Mercadería (complemento)") = vbNo Then RsAux.Close: Exit Sub
                    aLocalBaja = paLocalPuerto
                Else
                    If MsgBox("La mercadería estaba pendiente de arribo a Puerto. " & Chr(vbKeyReturn) & "El sistema actualizarán los datos en el embarque (no se realizarán traslados)." & vbCrLf & _
                                   "(*) Sólo para la mercadería embarcada." & vbCrLf & vbCrLf & _
                                   "Desea continuar con la operación.", vbQuestion + vbYesNo, "Arribo de Mercadería (complemento)") = vbNo Then RsAux.Close: Exit Sub
                End If
            End If
        End If
        RsAux.Close
    End If
    '-------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo ErrGD
    aCodigoRemito = gRemito
    
    FechaDelServidor
    ' *** Para hacer los movimientos con la fecha que ingresaron en el comprobante  ****
    gFechaServidor = Format(tFecha.Text, "dd/mm/yyyy") & " " & Format(gFechaServidor, "hh:mm:ss")
    
    cBase.BeginTrans            '--------------------------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo ErrResumo
    
    If aLocalBaja <> 0 And prmArticulosEmb <> "" Then     'Si hay que hacer un traslado desde ZF o Puerto
        'Y se agregaron articulos que estaban embarcados
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
        If Trim(.Cell(flexcpText, I, 2)) <> "" Then        'Solo los que se agregaron
        
            'Doy Altas al Local DESTINO--------------------------------------------------------------------------------------------------------
            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
                CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1
            
            MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpValue, I, 1)), 1
            
            MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), _
                CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpValue, I, 1)), paEstadoArticuloEntrega, 1, cTipo.ItemData(cTipo.ListIndex), aCodigoRemito
            '----------------------------------------------------------------------------------------------------------------------------------------
            Dim bEmbarco As Boolean
            bEmbarco = InStr("," & prmArticulosEmb & ",", "," & .Cell(flexcpData, I, 0) & ",")
            If aLocalBaja <> 0 And bEmbarco Then            'Mercadería EN PUERTO O ZONA FRANCA
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
            
            '5/12/2007 Si es puerto le resto la cantidad pendiente
            If aLocalBaja = paLocalPuerto Then
                Dim RsF As rdoResultset
                'Le resto al pendiente la cantidad.
                Cons = "Select * From ArticuloFolder " & _
                     " WHERE AFoTipo = 2 And AFoCodigo = " & aFolder & " And AFoArticulo = " & CLng(.Cell(flexcpData, I, 0))
                Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsF.EOF Then
                    RsF.Edit
                    If Not IsNull(RsF("AFoPendiente")) Then
                        RsF("AFoPendiente") = RsF("AFoPendiente") - CCur(.Cell(flexcpValue, I, 1))
                    Else
                        RsF("AFoPendiente") = RsF("AFoCantidad") - CCur(.Cell(flexcpValue, I, 1))
                    End If
                    RsF.Update
                End If
                RsF.Close
            End If
            '5/12/07----------------------------------------------
            
            'Los agrego a las tablas de remitos de compra
            'Ahora se pueden agregar artículos del mismo tipo 6/6/00
            Cons = "Select * from RemitoCompraRenglon " & _
                       " Where RCRRemito  = " & aCodigoRemito & _
                       " And RCRArticulo = " & .Cell(flexcpData, I, 0)
            Set rsRem = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If rsRem.EOF Then rsRem.AddNew Else rsRem.Edit
            rsRem!RCRRemito = aCodigoRemito
            rsRem!RCRArticulo = .Cell(flexcpData, I, 0)
            rsRem!RCRCantidad = rsRem!RCRCantidad + .Cell(flexcpValue, I, 1)
            If Not rsRem.EOF Then
                rsRem!RCREnCompania = rsRem!RCREnCompania + .Cell(flexcpValue, I, 1)
                rsRem!RCRRemanente = rsRem!RCRRemanente + .Cell(flexcpValue, I, 1)
            Else
                rsRem!RCREnCompania = 0
                rsRem!RCRRemanente = 0
            End If
            rsRem.Update: rsRem.Close
            
        End If
    Next
    
    Cons = "Select * from RemitoCompra Where RCoCodigo  = " & aCodigoRemito
    Set rsRem = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rsRem.Edit
    rsRem!RCoComentario = Trim(tComentario.Text)
    rsRem.Update
    rsRem.Close

    End With
    
    FechaDelServidor        'Reestablezco la fecha para registrar los sucesos *****
    
    If sucUsuario > 0 Then      'Suceso con defensa     22/2/01
        RegistroSuceso gFechaServidor, TipoSuceso.DiferenciaDeArticulos, paCodigoDeTerminal, CLng(tUsuario.Tag), 0, _
                            Descripcion:="Carpeta " & tFactura.Text & " (Modifica Arribo)", _
                            Defensa:=Trim(sucDefensa)
    End If
    cBase.CommitTrans           '------------------------------------------------------------------------------
    
    On Error Resume Next
    If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.CompraCarpeta Then      'Envio mensaje 22/2/01
        Dim aMsg As String
        aMsg = "Carpeta " & Trim(tFactura.Text) & "  (Modifica Arribo) " & vbCrLf
        aMsg = aMsg & "Arribó el " & Trim(tFecha.Text) & " al local " & Trim(cLocal.Text) & vbCrLf
        If Trim(tComentario.Text) <> "" Then aMsg = aMsg & "Comentarios: " & Trim(tComentario.Text) & vbCrLf
        
        With vsArticulo
            aMsg = aMsg & vbCrLf & "Artículos existentes..." & vbCrLf
            For I = 1 To .Rows - 1
                'If Trim(.Cell(flexcpText, I, 2)) <> "" Then        'Solo los que se agregaron
                If Trim(.Cell(flexcpText, I, 2)) = "" Then        'Solo los que Estaban
                    aMsg = aMsg & .Cell(flexcpText, I, 0) & vbTab
                    aMsg = aMsg & "Q: " & .Cell(flexcpText, I, 1) & vbCrLf
                End If
            Next
            
            aMsg = aMsg & vbCrLf & "Artículos agregados..." & vbCrLf
            For I = 1 To .Rows - 1
                If Trim(.Cell(flexcpText, I, 2)) <> "" Then        'Solo los que se agregaron
                    aMsg = aMsg & .Cell(flexcpText, I, 0) & vbTab
                    aMsg = aMsg & "Q: " & .Cell(flexcpText, I, 1) & vbCrLf
                End If
            Next
        End With
        
        aMsg = aMsg & vbCrLf & "Ingresado por el usuario " & miConexion.UsuarioLogueado(Nombre:=True)
        aMsg = aMsg & vbCrLf & "Defensa al suceso: " & Trim(sucDefensa)
        miConexion.EnviaMensaje prmMenUsuarioImportacion, "Diferencias Arribo Importación " & Trim(tFactura.Text), aMsg, DateAdd("s", 10, gFechaServidor), 0, prmMenUsuarioSistema
    End If
    
    
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub

ErrGD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción."
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    cBase.RollbackTrans
    Screen.MousePointer = 0: clsGeneral.OcurrioError Msg
End Sub

Private Function GraboBDRemitoCompra(Optional TipoFolder As Integer = 0, Optional Folder As Long = 0) As Long

    'Veo Si debo insertar en la tabla RemitoCompra (si hay mercadería q no pertence a la compañia) --> Todo lo que ingresa acá y mueva el stock
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

Private Function GraboBDTraslado(Optional DesdeCompania As Boolean = True) As Long
Dim pstrMemo As String, plngUsuario As Long
Dim pintOrigen As Integer, pintDestino As Integer

    If DesdeCompania Then
        pstrMemo = "Ingreso Mercadería: " & Trim(cTipo.Text) & " " & Trim(tSerieF.Text) & Trim(tFactura.Text)
        plngUsuario = CLng(tUsuario.Tag)
        pintOrigen = paLocalCompañia
        pintDestino = cLocal.ItemData(cLocal.ListIndex)
    Else
        pstrMemo = "Anulación Traslado " & gTraslado & " por " & Trim(tComentario.Text)
        plngUsuario = paCodigoDeUsuario
        pintDestino = paLocalCompañia
        pintOrigen = cLocal.ItemData(cLocal.ListIndex)
    End If
    
    GraboBDTraslado = 0
    For I = 1 To vsArticulo.Rows - 1
        If vsArticulo.Cell(flexcpText, I, 2) <> "" Or Not DesdeCompania Then
    
            Cons = "Insert Into Traspaso (TraFecha, TraLocalOrigen, TraLocalDestino, TraComentario, TraFechaEntregado, TraUsuarioInicial, TraUsuarioFinal) " _
                & " Values (" _
                & "'" & Format(gFechaServidor, sqlFormatoFH) & "', " _
                & pintOrigen & ", " _
                & pintDestino & ", " _
                & "'" & pstrMemo & "', " _
                & "'" & Format(gFechaServidor, sqlFormatoFH) & "', " & plngUsuario & ", " _
                & plngUsuario & ")"
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
    clsGeneral.OcurrioError "Ocurrió un error al consultar el stock en la compañia.", Err.Description
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
        Cons = "Select ArtId, ArtCodigo as Codigo, ArtNombre as Nombre from Articulo" _
                & " Where ArtNombre LIKE '" & Replace(Nombre, " ", "%") & "%'" _
                & " Order by ArtNombre"
        
        Dim objLista As New clsListadeAyuda
        Dim aSeleccionado As Long, aItem As String
        If objLista.ActivarAyuda(cBase, Cons, 5500, 1) > 0 Then
            aSeleccionado = objLista.RetornoDatoSeleccionado(0)
            aItem = objLista.RetornoDatoSeleccionado(1)
        End If
        Set objLista = Nothing
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
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
End Sub

Private Sub CargoDocumentos()

    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraCarpeta)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraCarpeta
    
    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraCarta)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraCarta
    
    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraContado)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraContado
    
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
Dim rs1 As rdoResultset

    'Cargo los datos del proveedor
    Cons = "Select * from ProveedorMercaderia Where PMeCodigo = " & Codigo
    Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rs1.EOF Then
        tProveedor.Text = Trim(rs1!PMeNombre)
        tProveedor.Tag = Codigo
    End If
    rs1.Close
    
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
    .EditText = Abs(.EditText)
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
