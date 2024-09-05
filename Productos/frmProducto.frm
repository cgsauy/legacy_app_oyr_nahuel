VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos"
   ClientHeight    =   5520
   ClientLeft      =   2490
   ClientTop       =   2520
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8205
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   250
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   800
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cliente"
            Object.ToolTipText     =   "Clientes."
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "historia"
            Object.ToolTipText     =   "Historia"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ceder"
            Object.ToolTipText     =   "Ceder producto."
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   60
      TabIndex        =   20
      Top             =   480
      Width           =   8055
      Begin MSMask.MaskEdBox tCi 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   12582912
         PromptInclude   =   0   'False
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#.###.###-#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tRuc 
         Height          =   285
         Left            =   2460
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   12582912
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99 999 999 9999"
         PromptChar      =   "_"
      End
      Begin VB.Label lCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Casa 9242557; Celular 099405236"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   6975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lNDireccion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   915
         Width           =   705
      End
      Begin VB.Label lDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Niagara 2345"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   915
         UseMnemonic     =   0   'False
         Width           =   6975
      End
      Begin VB.Label lTelefono 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Casa 9242557; Celular 099405236"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   1230
         UseMnemonic     =   0   'False
         Width           =   6975
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfonos:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "C.I./&RUC:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Producto "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1725
      Left            =   60
      TabIndex        =   18
      Top             =   3480
      Width           =   8055
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   14
         Top             =   960
         Width           =   6735
      End
      Begin VB.CommandButton bDireccion 
         BackColor       =   &H8000000E&
         Caption         =   "Dirección&..."
         Height          =   320
         Left            =   120
         Picture         =   "frmProducto.frx":0442
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox tDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Width           =   6735
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3300
         TabIndex        =   5
         Top             =   240
         Width           =   4635
      End
      Begin VB.TextBox tNroMaquina 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength       =   40
         TabIndex        =   12
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox tFacturaS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox tFacturaN 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox tFCompra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3300
         MaxLength       =   11
         TabIndex        =   10
         Text            =   "88/88/8888"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Co&mentarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lIdProducto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   28
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lEmpresa 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   2460
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lNroSerie 
         Caption         =   "Nº &Serie:"
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   4380
         TabIndex        =   11
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº Factura: "
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "F/&Compra:"
         Height          =   255
         Left            =   2460
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   5265
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9287
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   1335
      Left            =   60
      TabIndex        =   3
      Top             =   2100
      Width           =   8055
      _ExtentX        =   14208
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
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
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
   Begin ComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProducto.frx":0552
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProducto.frx":0664
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProducto.frx":0776
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProducto.frx":0888
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProducto.frx":099A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProducto.frx":0CB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProducto.frx":0DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProducto.frx":10E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProducto.frx":13FA
            Key             =   ""
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
      Begin VB.Menu MnuLinea1 
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
   Begin VB.Menu MnuIr 
      Caption         =   "&Ir a"
      Begin VB.Menu MnuIrHistoria 
         Caption         =   "&Historia de Servicios"
         Shortcut        =   ^H
      End
      Begin VB.Menu MnuIrCeder 
         Caption         =   "&Ceder Producto"
      End
      Begin VB.Menu MnuIrL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIrVOpe 
         Caption         =   "Visualización de Operaciones"
         Shortcut        =   {F12}
      End
      Begin VB.Menu MnuIrCliente 
         Caption         =   "Ficha de Cliente"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu MnuEmpresa 
      Caption         =   "Car&gar Lista"
      Begin VB.Menu MnuEmTodos 
         Caption         =   "Cargar &Todos los productos"
      End
      Begin VB.Menu MnuEmPorID 
         Caption         =   "Por ID de producto"
      End
      Begin VB.Menu MnuEmpPorCodigoArt 
         Caption         =   "Por Código de Artículo"
      End
      Begin VB.Menu MnuEmPorModelo 
         Caption         =   "Todos los"
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuMigrar 
      Caption         =   "mirgrar"
      Visible         =   0   'False
      Begin VB.Menu MnuMigrarEste 
         Caption         =   "Unificar otro producto en este"
      End
      Begin VB.Menu MnuMigrarCancel 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "frmProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'............................................................................................................................
'Modificacoines
'31-8-04    si viene a editar solo cargo ese producto no importa que sea cliente normal.
'10/3/2008  no borro productos vendidos.

Dim aTexto As String
Dim sNuevo As Boolean, sModificar As Boolean

Dim RsPro As rdoResultset

Public m_IDProducto As Long
Public m_IDArticulo As Long
Dim gCliente As Long, gTipoCliente As Integer
Dim gNuevo As Boolean

Public Property Get prmCliente() As Long
    prmCliente = gCliente
End Property
Public Property Let prmCliente(Codigo As Long)
    gCliente = Codigo
End Property

Public Property Get prmNuevo() As Boolean
    prmNuevo = gNuevo
End Property
Public Property Let prmNuevo(Estado As Boolean)
    gNuevo = Estado
End Property

Private Sub bDireccion_Click()

Dim aDirAnterior As Long, aRetorno As Long
    
    On Error GoTo errDirecccion
    If Val(lIdProducto.Caption) = 0 And Not sNuevo Then Exit Sub
    
    Screen.MousePointer = 11
    aDirAnterior = Val(bDireccion.Tag)
    
    Dim objDireccion As New clsDireccion
    objDireccion.ActivoFormularioDireccion cBase, bDireccion.Tag, gCliente, "Producto", "ProDireccion", "ProCodigo", Val(lIdProducto.Caption)
    Me.Refresh
    aRetorno = objDireccion.CodigoDeDireccion
    Set objDireccion = Nothing
    
    If aDirAnterior <> aRetorno And Not sNuevo Then
        If aRetorno <> 0 Then
            Cons = "Update Producto Set ProDireccion = " & aRetorno & " Where ProCodigo = " & Val(lIdProducto.Caption)
        Else
            Cons = "Update Producto Set ProDireccion = Null Where ProCodigo = " & Val(lIdProducto.Caption)
        End If
        cBase.Execute Cons
    End If
    
    CargoCamposDesdeBDDireccion aRetorno
    
    If tFCompra.Enabled Then Foco tFCompra
    Screen.MousePointer = 0
    Exit Sub
    
errDirecccion:
    clsGeneral.OcurrioError "Ocurrió un error al cargar la dirección.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoCamposDesdeBDDireccion(idDireccion As Long)

    If idDireccion <> 0 Then
        tDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, idDireccion, Departamento:=True, Localidad:=True, Zona:=True, EntreCalles:=True, Ampliacion:=True, ConfYVD:=True, ConEnter:=False)
    Else
        tDireccion.Text = ""
    End If
    tDireccion.Refresh
    bDireccion.Tag = idDireccion

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error GoTo errEliminar
    Screen.MousePointer = 11
    If sNuevo And Val(bDireccion.Tag) <> 0 Then
        Cons = "Delete Direccion Where DirCodigo = " & Val(bDireccion.Tag)
        cBase.Execute Cons
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errEliminar:
    clsGeneral.OcurrioError "Ocurrió un error al eliminar la dirección " & Val(lDireccion.Tag), Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End

End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error Resume Next
    FechaDelServidor

    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    
    sNuevo = False: sModificar = False
    
    DeshabilitoCampos
    bDireccion.Enabled = False
    
    InicializoGrillas
    LimpioCamposProducto
    LimpioCamposCliente
    
    If gCliente <> 0 Then
        CargoDatosCliente
    Else
        Botones False, False, False, False, False, Toolbar1, Me
    End If
    MnuEmPorModelo.Visible = False
    
    Cons = ""
    
    If m_IDProducto > 0 Then
        CargoLista False, , m_IDProducto
        Cons = "Select ArtNombre, ArtCodigo From Articulo, Producto Where ProCodigo = " & m_IDProducto & " And ProArticulo = ArtID"
    ElseIf m_IDArticulo > 0 Then
        Cons = "Select ArtNombre, ArtCodigo From Articulo Where ArtID = " & m_IDArticulo
    End If
    
    If Cons <> "" Then
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            With MnuEmPorModelo
                .Visible = True
                .Caption = Trim(RsAux(0))
                .Tag = RsAux(1)
            End With
        End If
    End If
    
    If gNuevo And vsLista.Rows = 1 Then AccionNuevo

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Panels(3).Text = ""
End Sub


Private Sub Label13_Click()
    Foco tFCompra
End Sub

Private Sub Label4_Click()
    Foco tCi
End Sub

Private Sub Label9_Click()
    Foco tFacturaS
End Sub

Private Sub lEmpresa_Click()
    Foco tArticulo
End Sub

Private Sub lNroSerie_Click()
    Foco tNroMaquina
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuEmPorID_Click()
Dim sCodArt As String
On Error Resume Next
    If gCliente = 0 Then Exit Sub
    sCodArt = InputBox("Ingrese el código del producto a buscar.", "Cargar lista de productos")
    If IsNumeric(sCodArt) Then CargoLista True, , Val(sCodArt)
End Sub

Private Sub MnuEmPorModelo_Click()
    If Val(MnuEmPorModelo.Tag) > 0 Then CargoLista , Val(MnuEmPorModelo.Tag)
End Sub

Private Sub MnuEmpPorCodigoArt_Click()
Dim sCodArt As String
On Error Resume Next
    If gCliente = 0 Then Exit Sub
    sCodArt = InputBox("Ingrese el código o nombre del artículo a buscar.", "Cargar lista de productos")
    If sCodArt = "" Then Exit Sub
    If Not IsNumeric(sCodArt) Then sCodArt = db_FindArtXNombre(sCodArt)
    If IsNumeric(sCodArt) Then CargoLista True, Val(sCodArt)
End Sub

Private Sub MnuEmTodos_Click()
    If gCliente <> 0 Then CargoLista ClienteEmpresa:=True
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuIrCeder_Click()
    IrACeder
End Sub

Private Sub MnuIrCliente_Click()
    IrACliente
End Sub

Private Sub MnuIrHistoria_Click()
    IrAHistoria
End Sub

Private Sub MnuIrVOpe_Click()
    If gCliente <> 0 Then
        EjecutarApp App.Path & "\visualizacion de operaciones.exe", CStr(gCliente)
    End If
End Sub

Private Sub MnuMigrarCancel_Click()
Dim iQ As Integer

    MnuMigrar.Tag = ""
    For iQ = vsLista.FixedRows To vsLista.Rows - 1
        With vsLista
            .Cell(flexcpBackColor, iQ, 0, , vsLista.Cols - 1) = vbWhite
            .Cell(flexcpForeColor, iQ, 0, , vsLista.Cols - 1) = vbBlack
        End With
    Next

End Sub

Private Sub MnuMigrarEste_Click()
    
    MnuMigrarCancel_Click
    
    With vsLista
        MnuMigrar.Tag = .Row
        .Cell(flexcpBackColor, vsLista.Row, 0, , vsLista.Cols - 1) = &H8000&
        .Cell(flexcpForeColor, vsLista.Row, 0, , vsLista.Cols - 1) = vbWhite
    End With
    MsgBox "Seleccione con doble click el producto que considera que duplica al seleccionado.", vbInformation, "Unificar productos"

End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub AccionNuevo()
    
    Screen.MousePointer = 11
    If MnuMigrar.Tag <> "" Then Call MnuMigrarCancel_Click
    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    
    LimpioCamposProducto
    HabilitoCampos
    vsLista.Rows = 1
    Foco tArticulo
    Screen.MousePointer = 0
    
End Sub

Private Sub AccionModificar()
    
    On Error GoTo ErrAM
    Dim aidProducto As Long
    
    If MnuMigrar.Tag <> "" Then Call MnuMigrarCancel_Click
    
    aidProducto = vsLista.Cell(flexcpData, vsLista.Row, 0)
    If vsLista.Rows = 1 Or aidProducto = 0 Then
        MsgBox "No hay seleccionado un producto para modificar.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    CargoDatosProducto aidProducto

    sModificar = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoCampos
    
    'Si el articulo tiene servicios no se puede modificar el tipo (solo q tenga permisos de Retoques)
    If tArticulo.Enabled Then
        If Not miConexion.AccesoAlMenu("Retoques") Then
            Cons = "Select * from Servicio Where SerProducto = " & aidProducto
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then tArticulo.Enabled = False: tArticulo.BackColor = Colores.Inactivo
            RsAux.Close
        End If
    End If
    
    If tArticulo.Enabled Then Foco tArticulo Else Foco tNroMaquina
    Exit Sub
    
ErrAM:
    clsGeneral.OcurrioError "Error al cargar la ficha para modificar.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionGrabar()
Dim aMsgError As String

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar los datos ingresados en la ficha", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    On Error GoTo errBegin
    FechaDelServidor
    If sNuevo Then
        cBase.BeginTrans            'Comienzo la transaccion------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!
        
        On Error GoTo errResumo
        
        'Si selecciono alguno con nro de serie, elimino tabla ProductosVendidos
        If Trim(tNroMaquina.Tag) <> "" And Val(tFacturaS.Tag) <> 0 And UCase(Trim(tNroMaquina.Tag)) = UCase(Trim(tNroMaquina.Text)) Then
            Cons = "Select * from ProductosVendidos " & _
                       " Where PVeDocumento = " & Val(tFacturaS.Tag) & _
                       " And PVeArticulo = " & Val(tArticulo.Tag) & _
                       " And PVeNSerie = '" & Replace(Trim(tNroMaquina.Text), "'", "''") & "'"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                '10/3/2008 no borramos más
                'RsAux.Delete
                lNroSerie.Tag = 1
            End If
            RsAux.Close
        End If
        
        
        Cons = "Select * From Producto Where ProCodigo = 0"
        Set RsPro = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsPro.AddNew
        GraboDatosBDProducto
        RsPro.Update: RsPro.Close
        
        cBase.CommitTrans         'Finalizo la transaccion------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!

    Else
        cBase.BeginTrans            'Comienzo la transaccion------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!
        
        On Error GoTo errResumo
        
        'Si selecciono alguno con nro de serie, elimino tabla ProductosVendidos
        If Trim(tNroMaquina.Tag) <> "" And Val(tFacturaS.Tag) <> 0 And UCase(Trim(tNroMaquina.Tag)) = UCase(Trim(tNroMaquina.Text)) Then
            Cons = "Select * from ProductosVendidos " & _
                       " Where PVeDocumento = " & Val(tFacturaS.Tag) & _
                       " And PVeArticulo = " & Val(tArticulo.Tag) & _
                       " And PVeNSerie = '" & Replace(Trim(tNroMaquina.Tag), "'", "''") & "'"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                'RsAux.Delete
                lNroSerie.Tag = 1
            End If
            RsAux.Close
        End If
        
        
        Cons = "Select * From Producto Where ProCodigo = " & vsLista.Cell(flexcpData, vsLista.Row, 0)
        Set RsPro = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        RsPro.Edit
        GraboDatosBDProducto
        RsPro.Update: RsPro.Close
        
        cBase.CommitTrans         'Finalizo la transaccion------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!
    End If
    
    sNuevo = False: sModificar = False
    DeshabilitoCampos
    LimpioCamposProducto
    bDireccion.Enabled = False
    CargoLista
    vsLista.SetFocus
    Screen.MousePointer = 0
    Exit Sub
    
errBegin:
    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
    
errResumo:
    Resume errRelajo
errRelajo:
    cBase.RollbackTrans
    If aMsgError = "" Then aMsgError = "Ocurrió un error al intentar realizar la transacción."
    clsGeneral.OcurrioError aMsgError, Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionEliminar()
    
    Dim aidProducto As Long, aTexto As String
    
    aidProducto = vsLista.Cell(flexcpData, vsLista.Row, 0)
    If vsLista.Rows = 1 Or aidProducto = 0 Then
        MsgBox "No hay seleccionado un producto para eliminar.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    With vsLista
        aTexto = "Id: " & .Cell(flexcpText, .Row, 0) & Chr(vbKeyReturn)
        aTexto = aTexto & "Tipo: " & .Cell(flexcpText, .Row, 1) & Chr(vbKeyReturn)
    End With
    
    'Hay que validar si tiene servicions !!!!!!!!-----------------------------------------------------------------------------
    Screen.MousePointer = 11
    Cons = "Select * from Servicio Where SerProducto = " & aidProducto
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        MsgBox "El producto seleccionado tiene datos de servicio. No podrá eliminarlo.", vbExclamation, "ATENCIÓN"
        RsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If
    RsAux.Close: Screen.MousePointer = 0
    '---------------------------------------------------------------------------------------------------------------------------
    
    If MsgBox("Confirma eliminar el producto seleccionado." & Chr(vbKeyReturn) & Chr(vbKeyReturn) & aTexto, vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo errResumo
    
    Cons = "Select * From Producto Where ProCodigo = " & vsLista.Cell(flexcpData, vsLista.Row, 0)
    Set RsPro = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(RsPro!ProDireccion) Then
        Cons = "Delete Direccion Where DirCodigo = " & RsPro!ProDireccion
        cBase.Execute Cons
    End If
    RsPro.Delete
    RsPro.Close
    
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub
    
errResumo:
    clsGeneral.OcurrioError "Ocurrió un error al intentar eliminar el producto.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionCancelar()
    
    Screen.MousePointer = 11
    On Error GoTo errCancelar
    
    If sNuevo And Val(bDireccion.Tag) <> 0 Then
        Cons = "Delete Direccion Where DirCodigo = " & Val(bDireccion.Tag)
        cBase.Execute Cons
    End If
    
    sNuevo = False: sModificar = False

    DeshabilitoCampos
    LimpioCamposProducto
    bDireccion.Enabled = False
    
    CargoLista
    
    If vsLista.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me Else: Botones True, False, False, False, False, Toolbar1, Me
    If vsLista.Rows > 1 Then vsLista.SetFocus
    Screen.MousePointer = 0
    Exit Sub

errCancelar:
    clsGeneral.OcurrioError "Error al cancelar la operación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tArticulo_Change()
    If Val(tArticulo.Tag) = 0 Then Exit Sub
    tArticulo.Tag = 0
    If sNuevo Then
        tFCompra.BackColor = Colores.Blanco: tFCompra.Enabled = True: tFCompra.Text = ""
        tFacturaS.BackColor = Colores.Blanco: tFacturaS.Enabled = True: tFacturaS.Text = ""
        tFacturaN.BackColor = Colores.Blanco: tFacturaN.Enabled = True: tFacturaN.Text = ""
        tNroMaquina.BackColor = Colores.Blanco: tNroMaquina.Enabled = True: tNroMaquina.Text = "": tNroMaquina.Tag = ""
    End If
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrTA
    
    If KeyCode = vbKeyReturn And Trim(tArticulo.Text) <> "" Then
        
        If Val(tArticulo.Tag) <> 0 Then
            If tFCompra.Enabled Then
                Foco tFacturaS
            Else
                If tNroMaquina.Enabled Then Foco tNroMaquina Else Foco tComentario
            End If
            Exit Sub
        End If
        
        Screen.MousePointer = 11
        If Not IsNumeric(tArticulo.Text) Then   'Busqueda por nombre
            Cons = "Select ArtID, 'Nombre' = ArtNombre, 'Código' = ArtCodigo From Articulo " _
                    & " Where ArtNombre Like '" & Replace(Trim(tArticulo.Text), " ", "%") & "%'" _
                    & " Order by ArtNombre"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
            If RsAux.EOF Then
                RsAux.Close
                MsgBox "No se encontró un artículo para el nombre ingresado.", vbInformation, "ATENCIÓN"
            Else
                RsAux.MoveNext
                If RsAux.EOF Then
                    RsAux.MoveFirst
                    tArticulo.Text = Trim(RsAux!Nombre)
                    tArticulo.Tag = RsAux!ArtId
                Else
                    
                    Dim objAyuda  As New clsListadeAyuda
                    If objAyuda.ActivarAyuda(cBase, Cons, 5000, 1, "Artículos") > 0 Then
                        tArticulo.Text = objAyuda.RetornoDatoSeleccionado(1)
                        tArticulo.Tag = objAyuda.RetornoDatoSeleccionado(0)
                    End If
                    Me.Refresh
                    Set objAyuda = Nothing
                End If
                RsAux.Close
            End If
        
        Else                                            'Busqueda por codigo
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & Val(tArticulo.Text)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
            If RsAux.EOF Then
                MsgBox "No se encontró un artículo para el código ingresado.", vbInformation, "ATENCIÓN"
            Else
                tArticulo.Text = Trim(RsAux!Nombre)
                tArticulo.Tag = RsAux!ArtId
            End If
            RsAux.Close
        End If
        
        If Val(tArticulo.Tag) <> 0 Then
            If paClienteEmpresa <> gCliente Then
                If ListaDeCompras Then
                    If tNroMaquina.Enabled Then Foco tNroMaquina Else Foco tComentario
                Else
                    Foco tFacturaS
                End If
            Else
                If tNroMaquina.Enabled Then Foco tNroMaquina Else Foco tComentario
            End If
        End If
        
        Screen.MousePointer = 0
    Else
        If KeyCode = vbKeyReturn Then Foco tFCompra
    End If
    Exit Sub

ErrTA:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ListaDeCompras() As Boolean
    
    On Error GoTo errCompras
    Dim aCompra As Long, aKey As String
    
    aCompra = 0: ListaDeCompras = False
    Screen.MousePointer = 11
    
    Cons = "Select DocCodigo, 'Fecha' = DocFecha, 'Documento' = RTrim(DocSerie) + ' ' + Convert(char(6), DocNumero), PVeNSerie as 'Nro. de Serie', 'Artículo' = ArtNombre " & _
               " From Documento Left Outer Join ProductosVendidos On DocCodigo = PVeDocumento And PVeArticulo = " & Val(tArticulo.Tag) & _
               " , Renglon, Articulo " & _
               " Where DocCodigo = RenDocumento" & _
               " And DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ")" & _
               " And RenArticulo = ArtID" & _
               " And DocCliente = " & gCliente & _
               " And DocAnulado = 0 " & _
               " And ArtID = " & Val(tArticulo.Tag)
    
    frmAyuda.ActivarAyuda Cons, 8500, False, Trim(tArticulo.Text) & " Comprados"
    Me.Refresh
    
    aCompra = frmAyuda.RetornoFilaSeleccionada
    If aCompra = 0 Then Unload frmAyuda: Screen.MousePointer = 0: Exit Function
    
    aCompra = frmAyuda.RetornoDatoSeleccionado(0)
    aKey = Trim(frmAyuda.RetornoDatoSeleccionado(3))
    Unload frmAyuda
    
    If aCompra <> 0 Then
        Dim QProductos As Long, QDocumentos As Long
        QProductos = 0: QDocumentos = 0
        '1) Valido la cantidad asignada a documentos
        Cons = "Select Count(*) from Producto " & _
                    " Where ProCliente = " & gCliente & _
                    " And ProArticulo = " & Val(tArticulo.Tag) & _
                    " And ProDocumento is not null "
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then QProductos = RsAux(0)
        RsAux.Close
        
        If QProductos > 0 Then
            Cons = "Select Sum(RenCantidad) From Documento, Renglon" & _
               " Where DocCodigo = RenDocumento" & _
               " And DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ")" & _
               " And RenArticulo = " & Val(tArticulo.Tag) & _
               " And DocCliente = " & gCliente & _
               " And DocAnulado = 0"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then QDocumentos = RsAux(0)
            RsAux.Close
        End If
        
        If QProductos < QDocumentos Or QProductos = 0 Then
            'Busco los datos del documento para cargar campos
            'Cons = "Select * from Documento Left Outer Join ProductosVendidos On DocCodigo = PVeDocumento And PVeArticulo = " & Val(tArticulo.Tag) & " Where DocCodigo = " & aCompra
            Cons = "Select * from Documento Where DocCodigo = " & aCompra
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                If Not IsNull(RsAux!DocFecha) Then tFCompra.Text = Format(RsAux!DocFecha, "dd/mm/yyyy")
                If Not IsNull(RsAux!DocSerie) Then tFacturaS.Text = Trim(RsAux!DocSerie)
                If Not IsNull(RsAux!DocNumero) Then tFacturaN.Text = RsAux!DocNumero
                
                tFacturaS.Tag = aCompra             'id Del Documento Seleccionado
                tFCompra.BackColor = Colores.Inactivo: tFCompra.Enabled = False
                tFacturaS.BackColor = Colores.Inactivo: tFacturaS.Enabled = False
                tFacturaN.BackColor = Colores.Inactivo: tFacturaN.Enabled = False
                tNroMaquina.Text = aKey: tNroMaquina.Tag = aKey
                If Trim(aKey) <> "" Then tNroMaquina.BackColor = Colores.Inactivo: tNroMaquina.Enabled = False
                
                ListaDeCompras = True
            End If
            RsAux.Close
        Else
            MsgBox "Todos los artículos " & Trim(tArticulo.Text) & " ya fueron asignados." & Chr(vbKeyReturn) & _
                        "No podrá asignar el documento seleccionado al nuevo producto. Verifique.", vbExclamation, "Todos los documento fueron Asignados"
        End If
    End If
    
    Screen.MousePointer = 0
    Exit Function
    
errCompras:
    clsGeneral.OcurrioError "Ocurrió un error al procesar las compras del cliente.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub tCi_GotFocus()
    tCi.SelStart = 0: tCi.SelLength = 11
    Status.Panels(3).Text = "Ingrese la C.I. del cliente.      [F4]- Buscar personas"
End Sub

Private Sub tCi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then AccionBuscarCliente TipoCliente.Cliente
End Sub

Private Sub tCi_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        On Error GoTo errBuscar
        If Trim(tCi.Text) = "" Then Foco tRuc: Exit Sub
        If Len(tCi.Text) = 7 Then tCi.Text = clsGeneral.AgregoDigitoControlCI(tCi.Text)
        
        'Busco el cliente----------------------------------------------------------------------------------
        Screen.MousePointer = 11
        Dim aIDEmpresa As Long
        Cons = "Select * from Cliente Where CliCiRuc = '" & Trim(tCi.Text) & "' And CliTipo = " & TipoCliente.Cliente
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then gCliente = RsAux!CliCodigo Else gCliente = 0
        RsAux.Close
        '---------------------------------------------------------------------------------------------------
        
        Screen.MousePointer = 0
        If gCliente = 0 Then
            MsgBox "No existe un cliente para la C.I. ingresada.", vbExclamation, "Empresa Inexistente"
        Else
            CargoDatosCliente
        End If
        
        If Trim(tCi.Text) <> "" And vsLista.Rows > 1 Then vsLista.SetFocus
        Screen.MousePointer = 0
    End If
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar la empresa.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0: tComentario.SelLength = Len(tComentario.Text)
    Status.Panels(3).Text = "Comentarios del producto."
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tDireccion_GotFocus()
    Status.Panels(3).Text = "Dirección del producto (para la realización de servicios)."
End Sub

Private Sub tFacturaN_GotFocus()
    tFacturaN.SelStart = 0: tFacturaN.SelLength = Len(tFacturaN.Text)
    Status.Panels(3).Text = "Ingrese el número de la factura de compra."
End Sub

Private Sub tFacturaN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'busco el documento para el artículo.
        If Val(tArticulo.Tag) > 0 And Val(tFacturaS.Tag) = 0 Then
            loc_FindDocumento
            If tFCompra.Enabled Then
                Foco tFCompra
            ElseIf tNroMaquina.Enabled Then
                Foco tNroMaquina
            Else
                Foco tComentario
            End If
        Else
            Foco tFCompra
        End If
    End If
End Sub

Private Sub tFacturaS_GotFocus()
    tFacturaS.SelStart = 0: tFacturaS.SelLength = Len(tFacturaS.Text)
    Status.Panels(3).Text = "Ingrese la serie de la factura de compra."
End Sub

Private Sub tFacturaS_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then Foco tFacturaN
End Sub

Private Sub tFCompra_GotFocus()
    tFCompra.SelStart = 0: tFCompra.SelLength = Len(tFCompra.Text)
    Status.Panels(3).Text = "Fecha de compra del producto."
End Sub

Private Sub tFCompra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tNroMaquina
End Sub

Private Sub tFCompra_LostFocus()
    If IsDate(tFCompra.Text) Then tFCompra.Text = Format(tFCompra.Text, "dd/mm/yyyy")
End Sub

Private Sub tNroMaquina_GotFocus()
    tNroMaquina.SelStart = 0: tNroMaquina.SelLength = Len(tNroMaquina.Text)
    Status.Panels(3).Text = "Ingrese el número de máquina del producto"
End Sub

Private Sub tNroMaquina_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentario
End Sub

Private Sub tRuc_GotFocus()
    tRuc.SelStart = 0: tRuc.SelLength = 15
    Status.Panels(3).Text = "Ingrese el número de R.U.C..      [F4]- Buscar Empresas"
End Sub

Private Sub tRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then AccionBuscarCliente TipoCliente.Empresa
End Sub

Private Sub AccionBuscarCliente(Tipo As Integer)
    
    On Error GoTo errorB
    Screen.MousePointer = 11
    Dim objCliente As New clsBuscarCliente
    If Tipo = TipoCliente.Cliente Then objCliente.ActivoFormularioBuscarClientes cBase, Persona:=True
    If Tipo = TipoCliente.Empresa Then objCliente.ActivoFormularioBuscarClientes cBase, Empresa:=True
    Me.Refresh
    gCliente = objCliente.BCClienteSeleccionado
    Set objCliente = Nothing
        
    If gCliente <> 0 Then CargoDatosCliente
    Screen.MousePointer = 0
    Exit Sub
    
errorB:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tRuc_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        On Error GoTo errBuscar
        If Trim(tRuc.Text) = "" Then Foco tCi: Exit Sub
        
        Screen.MousePointer = 11
        'Busco la empresa por numero de Ruc--------------------------------------------------------
        Screen.MousePointer = 11
        Dim aIDEmpresa As Long
        Cons = "Select * from Cliente Where CliCiRuc = '" & Trim(tRuc.Text) & "' And CliTipo = " & TipoCliente.Empresa
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then gCliente = RsAux!CliCodigo Else gCliente = 0
        RsAux.Close
        '---------------------------------------------------------------------------------------------------
        
        If gCliente = 0 Then
            Screen.MousePointer = 0
            MsgBox "No existe una empresa para el número de RUC ingresado.", vbExclamation, "Empresa Inexistente"
        Else
            CargoDatosCliente
        End If
        
        If Trim(tRuc.Text) <> "" And vsLista.Rows > 1 Then vsLista.SetFocus
        Screen.MousePointer = 0
    End If
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar la empresa.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        
        Case "historia": IrAHistoria
        Case "ceder": IrACeder
        Case "cliente": IrACliente
        
        Case "salir": Unload Me
            
    End Select

End Sub

Private Sub CargoDatosCliente()
    
    On Error GoTo errCliente
    Screen.MousePointer = 11
    LimpioCamposCliente
    
    'Ficha del Cliente----------------------------------------------------------------------------------------------------------------
     Cons = "Select Cliente.*, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
           & " From Cliente, CPersona " _
           & " Where CliCodigo = " & gCliente _
           & " And CliCodigo = CPeCliente " _
                                                & " UNION All" _
           & " Select Cliente.*, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
           & " From Cliente, CEmpresa " _
           & " Where CliCodigo = " & gCliente _
           & " And CliCodigo = CEmCliente"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        gTipoCliente = RsAux!CliTipo
        If Not IsNull(RsAux!CliCIRuc) Then
            If RsAux!CliTipo = TipoCliente.Cliente Then tCi.Text = RsAux!CliCIRuc
            If RsAux!CliTipo = TipoCliente.Empresa Then tRuc.Text = RsAux!CliCIRuc
        End If
    End If
    lCliente.Caption = " " & Trim(RsAux!Nombre)
    
    lDireccion.Tag = 0
    If Not IsNull(RsAux!CliDireccion) Then
        lDireccion.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, True, True, True)
        lDireccion.Tag = RsAux!CliDireccion
    End If
    
    RsAux.Close
    '----------------------------------------------------------------------------------------------------------------------------------
    If gCliente <> 0 Then
        lTelefono.Caption = " " & TelefonoATexto(gCliente)
        Me.Refresh
        CargoLista
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCliente:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosProducto(idProducto As Long)
    
    On Error GoTo ErrCE
    Screen.MousePointer = 11
    LimpioCamposProducto
    
    Cons = "Select * from Producto, Articulo " _
            & " Where ProCodigo = " & idProducto _
            & " And ProArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        bDireccion.Enabled = True
        
        lIdProducto.Caption = Format(idProducto, "000")
        tArticulo.Text = Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!ArtId
        
        If Not IsNull(RsAux!ProCompra) Then tFCompra.Text = Format(RsAux!ProCompra, "dd/mm/yyyy")
        If Not IsNull(RsAux!ProFacturaS) Then tFacturaS.Text = Trim(RsAux!ProFacturaS)
        If Not IsNull(RsAux!ProFacturaN) Then tFacturaN.Text = RsAux!ProFacturaN
        If Not IsNull(RsAux!ProNroSerie) Then tNroMaquina.Text = Trim(RsAux!ProNroSerie)
        If Not IsNull(RsAux!ProDireccion) Then CargoCamposDesdeBDDireccion RsAux!ProDireccion
        If Not IsNull(RsAux!ProComentario) Then tComentario.Text = Trim(RsAux!ProComentario)
        
        If Not IsNull(RsAux!ProDocumento) Then tFacturaS.Tag = RsAux!ProDocumento Else tFacturaS.Tag = 0
        
        If RsAux!ProAsignEnDiaVta Then
            With lNroSerie
                .Caption = "Nº &Serie x Vta.:"
                .Tag = 1
                .ForeColor = &HC0&
            End With
        End If
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub

ErrCE:
    clsGeneral.OcurrioError "Ocurrió un error al cargar la información del producto.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub LimpioCamposProducto()

    lIdProducto.Caption = ""
    tArticulo.Text = ""
    tFCompra.Text = ""
    tFacturaS.Text = "": tFacturaS.Tag = 0: tFacturaN.Text = ""
    tNroMaquina.Text = ""
    bDireccion.Tag = 0
    tDireccion.Text = ""
    tComentario.Text = ""
    With lNroSerie
        .Caption = "Nº &Serie:"
        .Tag = ""
        .ForeColor = &H80000012
    End With
        
End Sub

Private Sub LimpioCamposCliente()
    tCi.Text = "": tRuc.Text = ""
    lDireccion.Caption = ""
    lTelefono.Caption = ""
    lCliente.Caption = ""
End Sub


Private Sub DeshabilitoCampos()

    'Campos de cliente--------------------------------------------------------
    tRuc.BackColor = Colores.Blanco: tRuc.Enabled = True
    tCi.BackColor = Colores.Blanco: tCi.Enabled = True
    
    'Campos del producto.-------------------------------
    tArticulo.BackColor = Inactivo: tArticulo.Enabled = False
    tFCompra.BackColor = Inactivo: tFCompra.Enabled = False
    tFacturaS.BackColor = Inactivo: tFacturaS.Enabled = False
    tFacturaN.BackColor = Inactivo: tFacturaN.Enabled = False
    tNroMaquina.BackColor = Inactivo: tNroMaquina.Enabled = False
    tComentario.BackColor = Inactivo: tComentario.Enabled = False
    
    vsLista.Enabled = True: vsLista.BackColor = Colores.Blanco
    
End Sub

Private Sub HabilitoCampos()
    
    'Campos de cliente--------------------------------------------------------
    tRuc.BackColor = Inactivo: tRuc.Enabled = False
    tCi.BackColor = Inactivo: tCi.Enabled = False
    
    'Campos del producto.-------------------------------
    If (sModificar And Val(tFacturaS.Tag) = 0) Or sNuevo Then tArticulo.BackColor = Colores.Obligatorio: tArticulo.Enabled = True
    
    If Val(tFacturaS.Tag) = 0 And Val(lNroSerie.Tag) = 0 Then       'No tiene documento
        tFacturaS.BackColor = Colores.Blanco: tFacturaS.Enabled = True
        tFacturaN.BackColor = Colores.Blanco: tFacturaN.Enabled = True
    End If
    
    If Val(lNroSerie.Tag) = 0 Then
        tNroMaquina.BackColor = Colores.Blanco: tNroMaquina.Enabled = True
    End If
    
    If Not miConexion.AccesoAlMenu("productosunificar") Then
        tFCompra.BackColor = Colores.Blanco: tFCompra.Enabled = True
    End If
        
    tComentario.BackColor = Colores.Blanco: tComentario.Enabled = True
    
    vsLista.Enabled = False: vsLista.BackColor = Colores.Inactivo
    bDireccion.Enabled = True
    
End Sub

Private Function ValidoCampos() As Boolean
    
    On Error GoTo ErrVC
    ValidoCampos = False
    
    If gCliente = 0 Then
        MsgBox "no hay un cliente seleccionado para grabar la información.", vbExclamation, "ATENCION"
        Exit Function
    End If
    
    If Val(tArticulo.Tag) = 0 Then
        MsgBox "Se debe seleccionar el tipo de artículo para el producto.", vbExclamation, "ATENCION"
        Foco tArticulo: Exit Function
    End If
    
    If Trim(tFCompra.Text) <> "" Then
        If Not IsDate(tFCompra.Text) Then
            MsgBox "La fecha de compra ingresada no es correcta.", vbExclamation, "ATENCION"
            Foco tFCompra: Exit Function
        End If
    End If
    
    'Valido en numero de serie del producto, para el Tipo-----------------------------------------------
    If Trim(tNroMaquina.Text) <> "" Then
        Dim bHay As Boolean: bHay = False
        Screen.MousePointer = 11
        
        If Trim(tNroMaquina.Text) <> "0" And Trim(tNroMaquina.Text) <> "" And sNuevo Then
        
            Dim aValor As Long
            vsLista.Rows = 1
            'Si es nuevo cargo todos los productos que pueda tener el cliente con la misma serie o fecha de compra.
            Cons = "Select ProCodigo, ArtNombre, ProCompra, ProNroSerie From Producto, Articulo Where ProCliente = " & gCliente _
                & " And ProArticulo = " & Val(tArticulo.Tag) _
                & " And ProNroSerie = '" & Replace(Trim(tNroMaquina.Text), "'", "''") & "'" _
                & " And ProArticulo = ArtID"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            Do While Not RsAux.EOF
                With vsLista
                    .AddItem ""
                    .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!ProCodigo, "#,000")
                    aValor = RsAux!ProCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
                    .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
                
                    aValor = CalculoEstadoProducto(RsAux!ProCodigo)
                    .Cell(flexcpText, .Rows - 1, 2) = EstadoProducto(CInt(aValor), True)
                    If Not IsNull(RsAux!ProCompra) Then .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ProCompra, "dd/mm/yyyy")
                    If Not IsNull(RsAux!ProNroSerie) Then .Cell(flexcpText, .Rows - 1, 5) = RsAux!ProNroSerie
                End With
                RsAux.MoveNext
            Loop
            RsAux.Close
            
            If vsLista.Rows > 1 Then
                Screen.MousePointer = 0
                MsgBox "El cliente ya posee productos ingresados para el modelo y número de serie ingresados." & vbCr & vbCr & "Los mismos fueron cargados en la lista." & vbCr & "No podrá grabar.", vbInformation, "Posible Duplicación"
                Exit Function
            End If
            
        Else
            'Esta modificando.
            If UCase(Trim(tNroMaquina.Tag)) <> UCase(Trim(tNroMaquina.Text)) And Trim(tNroMaquina.Text) <> "0" And Trim(tNroMaquina.Text) <> "" Then
                Cons = "Select ProCodigo, ArtNombre, ProCompra, ProNroSerie From Producto, Articulo Where ProCliente = " & gCliente _
                    & " And ProArticulo = " & Val(tArticulo.Tag) & " And ProCodigo <> " & vsLista.Cell(flexcpData, vsLista.Row, 0) _
                    & " And ProNroSerie = '" & Replace(Trim(tNroMaquina.Text), "'", "''") & "'" _
                    & " And ProArticulo = ArtID"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    Screen.MousePointer = 0
                    MsgBox "El cliente tiene un producto ingresado con el mismo número de serie." & vbCr & vbCr & "El código del producto es :" & RsAux!ProCodigo & vbCr & vbCr & "No podrá grabar.", vbExclamation, "Posible duplicación"
                    RsAux.Close
                    Exit Function
                End If
                RsAux.Close
            End If
        End If
        Dim iCliCod As Long, iProCed As Long
        
        '11/9/2006 agregue cPersona y cEmpresa para sacar el nombre.
        Cons = "Select CliCodigo, CliCiRuc as 'CI/RUC Cliente', ProCodigo as 'Id_Producto', ProCompra as 'Compra' , RTrim(ProFacturaS) + ' ' + Convert(char(6), ProFacturaN) as 'Factura', ProComentario as 'Comentario Producto', CPeApellido1, CPeNombre1, CEmFantasia " _
                & " from Producto, Cliente" _
                    & " Left outer join CPersona On CliCodigo = CPeCliente" _
                    & " Left Outer Join CEmpresa On CliCodigo = CEmCliente " _
                & " Where ProArticulo = " & Val(tArticulo.Tag) _
                & " And ProNroSerie = '" & Replace(Trim(tNroMaquina.Text), "'", "''") & "'" _
                & " And ProCliente = CliCodigo"
        If sModificar Then Cons = Cons & " And ProCodigo <> " & Val(lIdProducto.Caption)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            iCliCod = RsAux("CliCodigo")
            iProCed = RsAux("Id_Producto")
            'bHay = True
            If Not IsNull(RsAux("CPeApellido1")) Then
                If Not IsNull(RsAux(1)) Then
                    Cons = "(" & clsGeneral.RetornoFormatoCedula(RsAux(1)) & ") "
                Else
                    Cons = ""
                End If
                Cons = Cons & Trim(RsAux("CPeApellido1")) & ", " & Trim(RsAux("CPeNombre1"))
            Else
                If Not IsNull(RsAux(1)) Then
                    Cons = "( " & clsGeneral.RetornoFormatoRuc(RsAux(1)) & ") "
                Else
                    Cons = ""
                End If
                Cons = Cons & Trim(RsAux("CEmFantasia"))
            End If
            Screen.MousePointer = 0
            
            If iCliCod = paClienteEmpresa Then
                RsAux.Close
                '10/3/2008 no pregunto más lo hago derecho.
                'If MsgBox("Un producto con ese número de serie está en stock." & vbCrLf & vbCrLf & "¿Desea cederlo al cliente?", vbQuestion + vbYesNo, "Ceder producto de stock") = vbYes Then
                    'cedo y limpio la ficha.
                    loc_CederProducto iProCed
                'End If
                Exit Function
            Else
                If MsgBox("Existe un producto ingresado con ese número de serie para el cliente:" & vbCrLf & vbCrLf & Cons & vbCrLf & vbCrLf & "¿Usted quiere ingresarlo sin número de serie?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                    tNroMaquina.Text = ""
                Else
                    RsAux.Close
                    Exit Function
                End If
            End If
'            MsgBox "Atención un producto con ese número de serie ya está asignado al cliente:" & vbCrLf & vbCrLf & Cons & vbCrLf & "Ingrese el servicio con los datos de este cliente ó en caso de estar completamente seguro que el producto pertenece a este cliente ceda el producto.", vbInformation, "Atención"
        End If
        RsAux.Close
        
        'Controlo Tabla Productos Vendidos
        Cons = " Select CEmNombre as 'Cliente', DocFecha, DocSerie, DocNumero From ProductosVendidos, Documento, CEmpresa" & _
                    " Where PVeDocumento = DocCodigo And DocCliente = CEmCliente " & _
                    " And PVeNSerie = '" & Replace(Trim(tNroMaquina.Text), "'", "''") & "' And DocCliente <> " & gCliente & _
                    " And PVeArticulo = " & Val(tArticulo.Tag) & _
                            " Union All" & _
                    " Select Rtrim(CPeApellido1) + ', ' + CPeNombre1  as 'Cliente', DocFecha, DocSerie, DocNumero From ProductosVendidos, Documento, CPersona" & _
                    " Where PVeDocumento = DocCodigo And DocCliente = CPeCliente" & _
                    " And PVeNSerie = '" & Replace(Trim(tNroMaquina.Text), "'", "''") & "'" & " And DocCliente <> " & gCliente & _
                    " And PVeArticulo = " & Val(tArticulo.Tag)

        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            Cons = ""
            Do While Not RsAux.EOF
                Cons = Cons & Trim(RsAux!Cliente) & " - Factura " & Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero) & ", el " & Format(RsAux!DocFecha, "dd/mm/yyyy hh:ss") & vbCrLf
                RsAux.MoveNext
            Loop
            MsgBox "ATENCIÓN, este producto fue entregado a..." & vbCrLf & vbCrLf & Cons, vbInformation, "Producto Entregado a el/los clientes..."
        End If
        RsAux.Close
        
        Screen.MousePointer = 0
    End If
    '-----------------------------------------------------------------------------------------------------------
    
    If Trim(tNroMaquina.Tag) <> "" And UCase(Trim(tNroMaquina.Tag)) <> UCase(Trim(tNroMaquina.Text)) Then
        If MsgBox("Ud. cambió el nro de serie del artículo. Desea continuar", vbQuestion + vbYesNo, "Validación de Nº de Serie") = vbNo Then Exit Function
    End If
    
    ValidoCampos = True
    Exit Function
ErrVC:
    clsGeneral.OcurrioError "Error al validar los datos.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub loc_CederProducto(ByVal iProdC As Long)
On Error GoTo errCP
Dim bSave As Boolean
    FechaDelServidor
    Cons = "Select * From Producto where ProCodigo = " & iProdC
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        RsAux!ProCliente = gCliente
        RsAux!ProFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux.Update
        bSave = True
    Else
        MsgBox "No se encontró la ficha, refresque la información.", vbExclamation, "Atención"
    End If
    RsAux.Close
    If bSave Then
        sNuevo = False: sModificar = False
        DeshabilitoCampos
        LimpioCamposProducto
        bDireccion.Enabled = False
        CargoLista
        vsLista.SetFocus
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errCP:
    clsGeneral.OcurrioError "Error al ceder el producto.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub GraboDatosBDProducto()

    RsPro!ProCliente = gCliente
    RsPro!ProArticulo = Val(tArticulo.Tag)
    
    If Trim(tFCompra.Text) <> "" Then RsPro!ProCompra = Format(tFCompra.Text, sqlFormatoF) Else RsPro!ProCompra = Null
    If Trim(tFacturaS.Text) <> "" Then RsPro!ProFacturaS = Trim(tFacturaS.Text) Else RsPro!ProFacturaS = Null
    If Trim(tFacturaN.Text) <> "" Then RsPro!ProFacturaN = Trim(tFacturaN.Text) Else RsPro!ProFacturaN = Null
    If Trim(tNroMaquina.Text) <> "" Then RsPro!ProNroSerie = Replace(Trim(tNroMaquina.Text), "'", "''") Else RsPro!ProNroSerie = Null
    If sNuevo Then
        If Val(bDireccion.Tag) <> 0 Then RsPro!ProDireccion = Val(bDireccion.Tag) Else RsPro!ProDireccion = Null
    End If
    RsPro!ProFModificacion = Format(gFechaServidor, sqlFormatoFH)
    
    If Trim(tComentario.Text) <> "" Then RsPro!ProComentario = Trim(tComentario.Text) Else RsPro!ProComentario = Null
    If Val(tFacturaS.Tag) <> 0 Then RsPro!ProDocumento = Val(tFacturaS.Tag) Else RsPro!ProDocumento = Null
    If lNroSerie.Tag = "1" Then RsPro!ProAsignEnDiaVta = 1 Else RsPro!ProAsignEnDiaVta = 0

End Sub

Private Sub InicializoGrillas()

    With vsLista
        .Rows = 1: .Cols = 1
        
        .FormatString = "<Id|<Tipo de Artículo|^Estado|^F.Compra|<Garantía|<Nº Serie"
        .ColWidth(0) = 700: .ColWidth(1) = 3800: .ColWidth(3) = 975: .ColWidth(5) = 1000 ': .ColWidth(5) = 1300
        
        .WordWrap = False
        .MergeCells = flexMergeSpill
        .ExtendLastCol = True
    End With

End Sub

Private Sub vsLista_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    vsLista.Row = vsLista.MouseRow
End Sub

Private Sub vsLista_DblClick()

    On Error Resume Next
    Dim aidProducto As Long
    aidProducto = vsLista.Cell(flexcpData, vsLista.Row, 0)
    If vsLista.Rows > 1 And aidProducto <> 0 Then CargoDatosProducto aidProducto
    
    If MnuMigrar.Tag <> "" And vsLista.Cell(flexcpBackColor, vsLista.Row, 0) = vbWhite Then
        
        'Veo si es el mismo id.
        If vsLista.Cell(flexcpData, Val(MnuMigrar.Tag), 1) = vsLista.Cell(flexcpData, vsLista.Row, 1) Then
            Select Case MsgBox("La información del artículo pintado en verde son considerados los datos correctos y los que se mantendrán." & vbCr & vbCr & "¿Confirma unificar el producto?", vbYesNoCancel, "Unificar productos")
                Case vbYes
                    db_SaveUnificacion
                Case vbCancel
                    MnuMigrarCancel_Click
            End Select
        Else
            MsgBox "Son dos tipos de artículos distintos.", vbExclamation, "Atención"
        End If
    End If
    
End Sub

Private Sub vsLista_GotFocus()
    Status.Panels(3).Text = "[Enter]- Cargar datos para del producto seleccionado."
End Sub

Private Sub vsLista_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyReturn Then
        Dim aidProducto As Long
        
        aidProducto = vsLista.Cell(flexcpData, vsLista.Row, 0)
        If vsLista.Rows > 1 And aidProducto <> 0 Then CargoDatosProducto aidProducto
    End If
    
End Sub

Private Sub vsLista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Shift = 0 And Button = 2 And vsLista.Row >= 1 Then
        PopupMenu MnuMigrar
    End If
    
End Sub


Private Function db_FindArtXNombre(ByVal sName As String) As Long
On Error GoTo errFAN
Dim rsFA As rdoResultset
    Screen.MousePointer = 11
    db_FindArtXNombre = 0
    Cons = "Select ArtCodigo as 'Código', rTrim(ArtNombre) as 'Nombre' From Artículo Where ArtNombre Like '" & Replace(sName, " ", "%") & "%'"
    Set rsFA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsFA.EOF Then
        rsFA.Close
        MsgBox "No existe un artículo con ese nombre.", vbExclamation, "Buscar"
    Else
        rsFA.MoveNext
        If rsFA.EOF Then
            rsFA.MoveFirst
            db_FindArtXNombre = rsFA(0)
            rsFA.Close
        Else
            rsFA.Close
            Dim objL As New clsListadeAyuda
            If objL.ActivarAyuda(cBase, Cons, Titulo:="Buscar artículos") > 0 Then
                db_FindArtXNombre = objL.RetornoDatoSeleccionado(0)
            End If
            Set objL = Nothing
        End If
    End If
    Screen.MousePointer = 0
    Exit Function
errFAN:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar artículo por nombre", Err.Description
End Function

Private Sub CargoLista(Optional ClienteEmpresa As Boolean = False, Optional lCodProd As Long = 0, Optional lIdProducto As Long)
Dim aRsPro As rdoResultset
Dim aValor As Long
Dim RsGar As rdoResultset

    If gCliente = 0 Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo errEmpleos
    LimpioCamposProducto
    vsLista.Rows = 1: Me.Refresh
    
    Cons = "Select * from Producto, Articulo" _
           & " Where ProCliente = " & gCliente & " And ProArticulo = ArtID"

    If lCodProd > 0 Then
        Cons = Cons & " And ArtCodigo = " & lCodProd
    ElseIf lIdProducto > 0 Then
        Cons = Cons & " And ProCodigo = " & lIdProducto
    End If
    
    'Se el paClienteEmpresa = gCliente y ClienteEmpresa = true ---> Cargo Todos, sino solo los modificados HOY
    If paClienteEmpresa = gCliente And lIdProducto = 0 Then
        If Not ClienteEmpresa Then Cons = Cons & " And ProFModificacion >= '" & Format(gFechaServidor, sqlFormatoF) & "'"
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        With vsLista
            
            .AddItem ""
            
            aValor = RsAux!ArtCodigo: .Cell(flexcpData, .Rows - 1, 1) = aValor
            
            .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!ProCodigo, "#,000")
            aValor = RsAux!ProCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
            
            aValor = CalculoEstadoProducto(RsAux!ProCodigo)
            .Cell(flexcpText, .Rows - 1, 2) = EstadoProducto(CInt(aValor), True)
            
            If Not IsNull(RsAux!ProCompra) Then .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ProCompra, "dd/mm/yyyy")
            
            'Saco garantia-----------------------------------------------------------------------------------------------------------------------
            Cons = "Select Garantia.* from ArticuloFacturacion, Garantia " _
                   & " Where AFaArticulo = " & RsAux!ArtId _
                   & " And AFaGarantia = GarCodigo"
            Set RsGar = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsGar.EOF Then .Cell(flexcpText, .Rows - 1, 4) = Trim(RsGar!GarNombre)
            RsGar.Close
            '--------------------------------------------------------------------------------------------------------------------------------------
            
            If Not IsNull(RsAux!ProNroSerie) Then .Cell(flexcpText, .Rows - 1, 5) = RsAux!ProNroSerie
            If .Rows > 15 And Not ClienteEmpresa Then
                MsgBox "Se cargaron los primeros 15 artículos del cliente." & vbCrLf _
                    & "Para visualizar todos acceda por el menú.", vbInformation, "ATENCIÓN"
                Exit Do
            End If
        End With
        RsAux.MoveNext
        
    Loop
    RsAux.Close
    
    If vsLista.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me Else: Botones True, False, False, False, False, Toolbar1, Me
    MnuMigrar.Tag = ""
    Screen.MousePointer = 0
    Exit Sub
    
errEmpleos:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los productos del cliente.", Err.Description
End Sub

Private Sub IrAHistoria()
    On Error GoTo errorH
    If vsLista.Rows = 1 Then Exit Sub
    EjecutarApp App.Path & "\Historia servicio", CStr(vsLista.Cell(flexcpValue, vsLista.Row, 0))
    Exit Sub

errorH:
    clsGeneral.OcurrioError "Ocurrió un error al ejecutar la aplicación", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub IrACeder()
    
    On Error GoTo errorC
    If vsLista.Rows = 1 Then Exit Sub
    If sNuevo Or sModificar Then
        MsgBox "Ud. está realizando cambios en los datos del producto." & Chr(vbKeyReturn) & "Termine la operación para luego acceder a al formulario.", vbInformation, "ATENCIÓN"
        Exit Sub
    End If
    
    If gCliente <> 0 Then EjecutarApp App.Path & "\Ceder producto", CStr(gCliente)
    Exit Sub
    
errorC:
    clsGeneral.OcurrioError "Ocurrió un error al ejecutar la aplicación", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub IrACliente()
    
    On Error GoTo errorC
    
    If sNuevo Or sModificar Then
        MsgBox "Ud. está realizando cambios en los datos del producto." & Chr(vbKeyReturn) & "Termine la operación para luego acceder a al formulario.", vbInformation, "ATENCIÓN"
        Exit Sub
    End If
    
    If gCliente <> 0 Then
        Screen.MousePointer = 11
        Dim objCliente As New clsCliente
        If gTipoCliente = TipoCliente.Cliente Then objCliente.Personas IdCliente:=gCliente
        If gTipoCliente = TipoCliente.Empresa Then objCliente.Empresas IdCliente:=gCliente
        Set objCliente = Nothing
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errorC:
    clsGeneral.OcurrioError "Ocurrió un error al ejecutar la aplicación", Err.Description
    Screen.MousePointer = 0
End Sub


Private Function ValidoNroMaquina(Tipo As Long, Numero As String, Optional idProducto As Long = 0) As Boolean

    On Error GoTo errValidoNM
    ValidoNroMaquina = False
    
    Dim rsSer As rdoResultset
    Cons = "Select * from Producto " & _
               " Where ProArticulo = " & Tipo & _
               " And ProNroSerie = '" & Trim(Numero) & "'" & _
               " And ProCodigo <> " & idProducto
    Set rsSer = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsSer.EOF Then ValidoNroMaquina = True
    rsSer.Close
    Exit Function
    
errValidoNM:
    clsGeneral.OcurrioError "Ocurrió un error al validar el número de serie.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub db_SaveUnificacion()
Dim rsQueda As rdoResultset
Dim lDir As Long

    Screen.MousePointer = 11
    On Error GoTo errBegin
    cBase.BeginTrans
    
    'Migro todos los servicios
    
    Cons = "Update Servicio Set SerProducto = " & vsLista.Cell(flexcpData, Val(MnuMigrar.Tag), 0) _
            & " Where SerProducto = " & vsLista.Cell(flexcpData, vsLista.Row, 0)
    cBase.Execute (Cons)
    
    Cons = "Select * From Producto Where ProCodigo = " & vsLista.Cell(flexcpData, vsLista.Row, 0)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Cons = "Select * From Producto Where ProCodigo = " & vsLista.Cell(flexcpData, Val(MnuMigrar.Tag), 0)
    Set rsQueda = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    rsQueda.Edit
    If IsNull(rsQueda!ProFacturaS) And Not IsNull(RsAux!ProFacturaS) Then
        rsQueda!ProFacturaS = RsAux!ProFacturaS
        rsQueda!ProFacturaN = RsAux!ProFacturaN
    End If
    If IsNull(rsQueda!ProCompra) And Not IsNull(RsAux!ProCompra) Then rsQueda!ProCompra = RsAux!ProCompra
    
    If IsNull(rsQueda!ProDireccion) And Not IsNull(RsAux!ProDireccion) Then
        rsQueda!ProDireccion = RsAux!ProDireccion
    ElseIf Not IsNull(RsAux!ProDireccion) Then
        lDir = RsAux!ProDireccion
    End If
    If IsNull(rsQueda!ProDocumento) And Not IsNull(RsAux!ProDocumento) Then rsQueda!ProDocumento = RsAux!ProDocumento
    If Not rsQueda!ProAsignEnDiaVta And RsAux!ProAsignEnDiaVta Then rsQueda!ProAsignEnDiaVta = 1
    rsQueda!ProFModificacion = Format(Now, "mm/dd/yyyy hh:nn:ss")
    rsQueda.Update
    rsQueda.Close
        
    RsAux.Delete
    
    If lDir > 0 Then
        Cons = "Delete Direccion Where DirCodigo = " & lDir
        cBase.Execute (Cons)
    End If
    
    cBase.CommitTrans
    Screen.MousePointer = 0
    CargoLista
    Exit Sub

errBegin:
    clsGeneral.OcurrioError "Error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
    
errResumo:
    Resume errRelajo
    Exit Sub
    
errRelajo:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al intentar realizar la transacción.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_FindDocumento()
On Error GoTo errFD
    Cons = "Select DocCodigo, DocFecha, DocSerie, DocNumero, PVeNSerie " & _
       " From Documento Left Outer Join ProductosVendidos On DocCodigo = PVeDocumento And PVeArticulo = " & Val(tArticulo.Tag) & _
        " , Renglon, Articulo " & _
       " Where DocCodigo = RenDocumento" & _
       " And DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ")" & _
       " And RenArticulo = ArtID" & _
       " And DocCliente = " & gCliente & _
       " And DocAnulado = 0 " & _
       " And ArtID = " & Val(tArticulo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!DocFecha) Then tFCompra.Text = Format(RsAux!DocFecha, "dd/mm/yyyy")
        If Not IsNull(RsAux!DocSerie) Then tFacturaS.Text = Trim(RsAux!DocSerie)
        If Not IsNull(RsAux!DocNumero) Then tFacturaN.Text = RsAux!DocNumero
        If Not IsNull(RsAux("PVeNSerie")) Then
            tNroMaquina.Text = RsAux("PVeNSerie")
            tNroMaquina.Tag = tNroMaquina.Text
            tNroMaquina.BackColor = Colores.Inactivo: tNroMaquina.Enabled = False
        End If
        tFacturaS.Tag = RsAux("DocCodigo")             'id Del Documento Seleccionado
        tFCompra.BackColor = Colores.Inactivo: tFCompra.Enabled = False
        tFacturaS.BackColor = Colores.Inactivo: tFacturaS.Enabled = False
        tFacturaN.BackColor = Colores.Inactivo: tFacturaN.Enabled = False
    End If
    RsAux.Close
    Exit Sub
errFD:
    clsGeneral.OcurrioError "Error al buscar un documento para el artículo ingresado.", Err.Description

End Sub
