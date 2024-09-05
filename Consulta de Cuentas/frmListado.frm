VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmListado 
   Caption         =   "Consulta de Cuentas"
   ClientHeight    =   7530
   ClientLeft      =   1770
   ClientTop       =   1680
   ClientWidth     =   10830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10830
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3375
      Left            =   3960
      TabIndex        =   8
      Top             =   1920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5953
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
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
      OutlineBar      =   1
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
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   11595
      TabIndex        =   13
      Top             =   6720
      Width           =   11655
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmListado.frx":0EF0
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmListado.frx":0FDA
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":10C4
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmListado.frx":12FE
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "frmListado.frx":1400
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "frmListado.frx":17C6
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":18C8
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":1BCA
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmListado.frx":1F0C
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":220E
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   7275
      Width           =   10830
      _ExtentX        =   19103
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
            Object.Width           =   10901
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1740
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton bEmail 
         Caption         =   "Email ..."
         Height          =   255
         Left            =   6480
         TabIndex        =   34
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cCuenta 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox tTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1320
         MaxLength       =   45
         TabIndex        =   3
         Top             =   600
         Width           =   5055
      End
      Begin MSMask.MaskEdBox tCCi 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox tCRuc 
         Height          =   285
         Left            =   3720
         TabIndex        =   7
         Top             =   960
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
      Begin VB.Label lEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7320
         TabIndex        =   35
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Comentario:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   32
         Top             =   1320
         Width           =   8895
      End
      Begin VB.Label lClave 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "APORTES REALIZADOS A LA CUENTA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   29
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo de &Cuenta:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "&Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lCCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5160
         TabIndex        =   26
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label7 
         Caption         =   "&Empresa:"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "&Persona:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   3735
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   6588
      _StockProps     =   229
      BorderStyle     =   1
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
      PreviewMode     =   1
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta1 
      Height          =   1695
      Left            =   1200
      TabIndex        =   9
      Top             =   3840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2990
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
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
      OutlineBar      =   1
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsSaldo 
      Height          =   555
      Left            =   120
      TabIndex        =   30
      Top             =   6120
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   979
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
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   0
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
      OutlineBar      =   1
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
   Begin VB.Label lSaldo 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALDOS PARA LA CUENTA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   31
      Top             =   5820
      Width           =   5055
   End
   Begin VB.Label lFactura 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FACTURAS EMITIDAS PARA LA CUENTA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6180
      TabIndex        =   28
      Top             =   5280
      Width           =   5055
   End
   Begin VB.Label lRecibo 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "APORTES REALIZADOS A LA CUENTA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   27
      Top             =   5460
      Width           =   5055
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Tipo As Integer
Private m_Id As Long

Private RsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

Public Property Let prmTipo(ByVal iTipo As Integer)
On Error Resume Next
    m_Tipo = iTipo
End Property

Public Property Let prmID(ByVal lID As Long)
On Error Resume Next
    m_Id = lID
End Property

Private Sub AccionLimpiar()
    tTitulo.Text = "": tCCi.Text = "": tCRuc.Text = "": lCCliente.Caption = ""
    vsConsulta.Rows = 1: vsConsulta1.Rows = 1: vsSaldo.Rows = 0
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bEmail_Click()
    If Val(bEmail.Tag) = 0 Then Exit Sub
    EjecutarApp App.Path & "\EMails.exe", Val(bEmail.Tag), True
    lEmail.Caption = CargoDireccionEMail(Val(bEmail.Tag))
    If Trim(lEmail.Caption) = "" Then
        lEmail.BackColor = vbRed
        lEmail.Caption = " Falta Ingresar EMail ..."
    Else
        lEmail.BackColor = &H800000
    End If
End Sub

Private Sub bImprimir_Click()
    AccionImprimir True
End Sub
Private Sub bNoFiltros_Click()
    AccionLimpiar
End Sub

Private Sub bPrimero_Click()
    IrAPagina vsListado, 1
End Sub

Private Sub bSiguiente_Click()
    IrAPagina vsListado, vsListado.PreviewPage + 1
End Sub

Private Sub bUltima_Click()
    IrAPagina vsListado, vsListado.PageCount
End Sub

Private Sub bZMas_Click()
    Zoom vsListado, vsListado.Zoom + 5
End Sub

Private Sub bZMenos_Click()
    Zoom vsListado, vsListado.Zoom - 5
End Sub

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub


Private Sub cCuenta_Click()

    On Error Resume Next
    If cCuenta.ItemData(cCuenta.ListIndex) = Cuenta.Colectivo Then
        tCCi.Text = "": tCRuc.Text = ""
        lCCliente.Caption = ""
        lEmail.Caption = "": lEmail.BackColor = &H800000
        lComentario.Caption = ""
        tCCi.Enabled = False: tCRuc.Enabled = False
        lCCliente.Enabled = False: lCCliente.BackColor = Colores.Gris
        bEmail.Tag = ""
        tTitulo.Enabled = True: tTitulo.BackColor = Colores.Blanco
    End If
    
    If cCuenta.ItemData(cCuenta.ListIndex) = Cuenta.Personal Then
        tCCi.Enabled = True: tCRuc.Enabled = True
        lCCliente.Enabled = True: lCCliente.BackColor = Colores.Azul
        lEmail.Caption = "": lEmail.BackColor = &H800000
        lComentario.Caption = ""
        tTitulo.Text = "": lCCliente.Caption = ""
        tTitulo.Enabled = False: tTitulo.BackColor = Colores.Gris
        bEmail.Tag = ""
    End If
    
End Sub

Private Sub cCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tTitulo.Enabled Then Foco tTitulo
        If tCCi.Enabled Then tCCi.SetFocus
    End If
End Sub

Private Sub chVista_Click()
    
    If chVista.Value = 0 Then
        vsListado.Visible = False
    Else
        AccionImprimir
        vsListado.Visible = True: vsListado.ZOrder 0
    End If
    Me.Refresh

End Sub

Private Sub Form_Activate()
    
    Me.Refresh
    If m_Tipo > 0 Then
        If m_Tipo = Cuenta.Colectivo Then
            cCuenta.ListIndex = 0
            CargoColectivoPorID m_Id
            AccionConsultar
        ElseIf m_Tipo = Cuenta.Personal Then
            cCuenta.ListIndex = 1
            BuscoClienteInicio m_Id
            AccionConsultar
        End If
        m_Tipo = 0
    End If
    Screen.MousePointer = 0
    
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad
    lClave.Visible = False
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas
    AccionLimpiar
    
    'Cargo datos en combo Cuenta-------------------------------------------------------
    cCuenta.AddItem "Colectivos"
    cCuenta.ItemData(cCuenta.NewIndex) = Cuenta.Colectivo
    cCuenta.AddItem "Cuenta Personal"
    cCuenta.ItemData(cCuenta.NewIndex) = Cuenta.Personal
    BuscoCodigoEnCombo cCuenta, Cuenta.Colectivo
    '------------------------------------------------------------------------------------------
    
    bCargarImpresion = True
    With vsListado
        .PaperSize = 1
        .Orientation = orPortrait
        .Zoom = 100
        .MarginBottom = 650: .MarginTop = 650
        .MarginRight = 350
    End With
    vsListado.Visible = False
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .OutlineBar = flexOutlineBarNone ' flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Moneda|<Fecha|<Recibo|<Cliente|<Destinado a...|>Aporte"
         .ColWidth(0) = 0: .ColWidth(1) = 950: .ColWidth(2) = 950: .ColWidth(3) = 3200: .ColWidth(4) = 3400: .ColWidth(5) = 1400
            
        .WordWrap = False
        .MergeCells = flexMergeSpill
        
    End With
    With vsConsulta1
        .OutlineBar = flexOutlineBarNone ' flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Moneda|<Fecha|<Boleta|<Cliente|<Articulo|>Aporte"
         .ColWidth(0) = 0: .ColWidth(1) = 950: .ColWidth(2) = 950: .ColWidth(3) = 3200: .ColWidth(4) = 3400: .ColWidth(5) = 1400
            
        .WordWrap = False
        .MergeCells = flexMergeSpill
    End With
    
    With vsSaldo
        .Cols = 2: .Rows = 0
        .ColWidth(0) = 3000: .ExtendLastCol = True
        .ColAlignment(1) = flexAlignRightCenter
        .FontBold = True
        .BackColor = Colores.Gris
    End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            
            Case vbKeyE: AccionConsultar
            
            Case vbKeyP: IrAPagina vsListado, 1
            Case vbKeyA: IrAPagina vsListado, vsListado.PreviewPage - 1
            Case vbKeyS: IrAPagina vsListado, vsListado.PreviewPage + 1
            Case vbKeyU: IrAPagina vsListado, vsListado.PageCount
            
            Case vbKeyAdd: Zoom vsListado, vsListado.Zoom + 5
            Case vbKeySubtract: Zoom vsListado, vsListado.Zoom - 5
            
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir True
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyC: AccionConfigurar
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11

    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    vsListado.Left = fFiltros.Left
    
    lRecibo.Width = vsListado.Width: lRecibo.Left = vsListado.Left
    lFactura.Width = vsListado.Width: lFactura.Left = vsListado.Left
    lSaldo.Width = vsListado.Width: lSaldo.Left = vsListado.Left
    
    lRecibo.Top = vsListado.Top
    vsConsulta.Top = lRecibo.Top + lRecibo.Height
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = ((vsListado.Height - (lRecibo.Height * 2)) / 2) - 500
    vsConsulta.Left = vsListado.Left
    
    lFactura.Top = vsListado.Top + lRecibo.Height + vsConsulta.Height + 100
    vsConsulta1.Top = lFactura.Top + lFactura.Height
    vsConsulta1.Width = vsListado.Width
    vsConsulta1.Height = vsConsulta.Height
    vsConsulta1.Left = vsListado.Left
    
    vsSaldo.Left = vsListado.Left: vsSaldo.Width = vsListado.Width
    lSaldo.Top = vsConsulta1.Top + vsConsulta1.Height + 100
    vsSaldo.Top = lSaldo.Top + lSaldo.Height
        
    picBotones.Width = vsListado.Width
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub AccionConsultar()

    On Error GoTo errConsultar
    If Not ValidoCampos Then Exit Sub
    
    Screen.MousePointer = 11
    vsListado.Visible = False: Me.Refresh
    bCargarImpresion = True
    vsConsulta.Rows = 1: vsConsulta1.Rows = 1: vsSaldo.Rows = 0
    Me.Refresh
    
    ConsultoAportes
    If vsConsulta.Rows > 1 Then ConsultoFacturas
    
    'Cargo los saldos------------------------------------------------------------------------------------------------------------------
    vsSaldo.Rows = 0
    With vsConsulta
        For I = 1 To .Rows - 1
            If .IsSubtotal(I) Then
                vsSaldo.AddItem .Cell(flexcpText, I, 0)
                vsSaldo.Cell(flexcpText, vsSaldo.Rows - 1, 1) = .Cell(flexcpText, I, 5)
            End If
        Next
    End With
    
    With vsConsulta1
    Dim J As Integer
        For I = 1 To .Rows - 1
            If .IsSubtotal(I) Then
                'Busco si esta inrgesada para descontar
                For J = 0 To vsSaldo.Rows - 1
                    If vsSaldo.Cell(flexcpText, J, 0) = .Cell(flexcpText, I, 0) Then
                        vsSaldo.Cell(flexcpText, J, 1) = Format(vsSaldo.Cell(flexcpValue, J, 1) - .Cell(flexcpValue, I, 5), FormatoMonedaP)
                        Exit For
                    End If
                Next
            End If
        Next
    End With
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub ConsultoAportes()

Dim aIDMoneda As Long, aTxtMoneda As String
Dim rs1 As rdoResultset

    aIDMoneda = 0
    
    'Armo consulta para sacar los documentos-------------------------------------------------------------
    Cons = "Select * From Documento, Cliente " _
                                & " Left Outer Join CPersona On CliCodigo = CPeCliente " _
                                & " Left Outer Join CEmpresa On CliCodigo = CEmCliente, " _
            & " CuentaDocumento " _
                & " Left Outer Join Articulo On ArtID = CDoIDArticulo " _
        & " Where CDoTipo = " & cCuenta.ItemData(cCuenta.ListIndex)
                
    Select Case cCuenta.ItemData(cCuenta.ListIndex)
        Case Cuenta.Colectivo: Cons = Cons & " And CDoIDTipo = " & Val(tTitulo.Tag)
        Case Cuenta.Personal: Cons = Cons & " And CDoIDTipo = " & Val(lCCliente.Tag)
    End Select
    
    Cons = Cons & " And CDoIDDocumento = DocCodigo " _
                        & " And CDoAsignado Is null" _
                        & " And DocAnulado = 0 " _
                        & " And DocTipo = " & TipoDocumento.ReciboDePago _
                        & " And DocCliente = CliCodigo " _
                        & " Order by DocMoneda, DocCodigo"
    '--------------------------------------------------------------------------------------------------------------
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        RsAux.Close: Screen.MousePointer = 0: vsConsulta.Rows = 1: Exit Sub
    End If
    
    vsConsulta.Rows = 1: vsConsulta.Refresh
    Do While Not RsAux.EOF
        
        If aIDMoneda <> RsAux!DocMoneda Then
            Cons = "Select * from Moneda Where MonCodigo = " & RsAux!DocMoneda
            Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            aTxtMoneda = ""
            If Not rs1.EOF Then
                aIDMoneda = RsAux!DocMoneda: aTxtMoneda = Trim(rs1!MonNombre)
            End If
            rs1.Close
        End If
        
        With vsConsulta
            .AddItem aTxtMoneda
                
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!DocFecha, "dd/mm/yyyy")
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!DocSerie) & " " & RsAux!DocNumero
            
            If RsAux!CliTipo = TipoCliente.Cliente Then
                If Not IsNull(RsAux!CliCiRuc) Then aTexto = "(" & clsGeneral.RetornoFormatoCedula(RsAux!CliCiRuc) & ") " Else aTexto = ""
                aTexto = aTexto & Trim(RsAux!CPeNombre1) & " " & Trim(RsAux!CPeApellido1)
            Else
                If Not IsNull(RsAux!CliCiRuc) Then aTexto = "(" & clsGeneral.RetornoFormatoRuc(RsAux!CliCiRuc) & ") " Else aTexto = ""
                aTexto = aTexto & Trim(RsAux!CEmFantasia)
            End If
            .Cell(flexcpText, .Rows - 1, 3) = aTexto
            
            If Not IsNull(RsAux!ArtNombre) Then .Cell(flexcpText, .Rows - 1, 4) = "(" & Format(RsAux!ArtCodigo, "#,000,000") & ") " & Trim(RsAux!ArtNombre)
            
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!DocTotal, FormatoMonedaP)
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close

    With vsConsulta
        .Subtotal flexSTSum, 0, 5, , Colores.Inactivo, Colores.Rojo, True, "%s"
    End With

End Sub

Private Sub ConsultoFacturas()

Dim aIDMoneda As Long, aTxtMoneda As String
Dim rs1 As rdoResultset
Dim aDocAnterior As Long

    aIDMoneda = 0: aDocAnterior = 0
    
    'Armo consulta para sacar los documentos-------------------------------------------------------------
    Cons = "Select * From CuentaDocumento, " _
                    & " Documento " _
                                & "Left Outer Join Renglon On DocCodigo = RenDocumento Left Outer Join Articulo on RenArticulo = ArtID, " _
                    & " Cliente " _
                                & " Left Outer Join CPersona On CliCodigo = CPeCliente " _
                                & " Left Outer Join CEmpresa On CliCodigo = CEmCliente " _
            & " Where CDoTipo = " & cCuenta.ItemData(cCuenta.ListIndex) _

    Select Case cCuenta.ItemData(cCuenta.ListIndex)
        Case Cuenta.Colectivo: Cons = Cons & " And CDoIDTipo = " & Val(tTitulo.Tag)
        Case Cuenta.Personal: Cons = Cons & " And CDoIDTipo = " & Val(lCCliente.Tag)
    End Select
    
    Cons = Cons & " And CDoIDDocumento = DocCodigo " _
                       & " And DocAnulado = 0 " _
                       & " And DocTipo In ( " & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ") " _
                       & " And DocCliente = CliCodigo " _
                       & " And CDoAsignado is not Null " _
                       & " Order by DocMoneda, DocCodigo"
    '--------------------------------------------------------------------------------------------------------------
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.Close: Screen.MousePointer = 0: vsConsulta1.Rows = 1: Exit Sub
    End If
    
    vsConsulta1.Rows = 1: vsConsulta1.Refresh
    Do While Not RsAux.EOF
        
        If aIDMoneda <> RsAux!DocMoneda Then
            Cons = "Select * from Moneda Where MonCodigo = " & RsAux!DocMoneda
            Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            aTxtMoneda = ""
            If Not rs1.EOF Then
                aIDMoneda = RsAux!DocMoneda: aTxtMoneda = Trim(rs1!MonNombre)
            End If
            rs1.Close
        End If
        
        With vsConsulta1
            .AddItem aTxtMoneda
                
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!DocFecha, "dd/mm/yyyy")
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!DocSerie) & " " & RsAux!DocNumero
            
            If RsAux!CliTipo = TipoCliente.Cliente Then
                If Not IsNull(RsAux!CliCiRuc) Then aTexto = "(" & clsGeneral.RetornoFormatoCedula(RsAux!CliCiRuc) & ") " Else aTexto = ""
                aTexto = aTexto & Trim(RsAux!CPeNombre1) & " " & Trim(RsAux!CPeApellido1)
            Else
                If Not IsNull(RsAux!CliCiRuc) Then aTexto = "(" & clsGeneral.RetornoFormatoRuc(RsAux!CliCiRuc) & ") " Else aTexto = ""
                aTexto = aTexto & Trim(RsAux!CEmFantasia)
            End If
            .Cell(flexcpText, .Rows - 1, 3) = aTexto
            
            If Not IsNull(RsAux!ArtNombre) Then
                .Cell(flexcpText, .Rows - 1, 4) = RsAux!RenCantidad & " (" & Format(RsAux!ArtCodigo, "#,000,000") & ") " & Trim(RsAux!ArtNombre)
            Else
                'Es un recibo, saco los articulos del credito
                Cons = "Select * from DocumentoPago, Renglon, Articulo Where DPaDocQSalda = " & RsAux!DocCodigo & " And DPaDocASaldar = RenDocumento And RenArticulo = ArtID"
                Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                aTexto = ""
                Do While Not rs1.EOF
                    aTexto = aTexto & rs1!RenCantidad & " (" & Format(rs1!ArtCodigo, "#,000,000") & ") " & Trim(rs1!ArtNombre) & ", "
                    rs1.MoveNext
                Loop
                rs1.Close
                If Len(aTexto) >= 2 Then aTexto = Mid(aTexto, 1, Len(aTexto) - 2)
                .Cell(flexcpText, .Rows - 1, 4) = aTexto
            End If
            
            If aDocAnterior <> RsAux!DocCodigo Then
                aDocAnterior = RsAux!DocCodigo
                If Not IsNull(RsAux!CDoAsignado) Then .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!CDoAsignado, FormatoMonedaP)
            Else
                .Cell(flexcpText, .Rows - 1, 1) = " ": .Cell(flexcpText, .Rows - 1, 2) = " ": .Cell(flexcpText, .Rows - 1, 3) = " "
            End If
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close

    With vsConsulta1
        .Subtotal flexSTSum, 0, 5, , Colores.Inactivo, Colores.Rojo, True, "%s"
    End With

End Sub

Private Sub Label12_Click()
    Foco tCCi
End Sub

Private Sub Label5_Click()
    Foco tTitulo
End Sub

Private Sub Label7_Click()
    tCRuc.SetFocus
End Sub

Private Sub Label8_Click()
    Foco cCuenta
End Sub

Private Sub tCCi_Change()
    lCCliente.Tag = 0
End Sub

Private Sub tCCi_GotFocus()
    tCCi.SelStart = 0: tCCi.SelLength = 11
End Sub

Private Sub tCCi_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo errorB
    If KeyCode = vbKeyF4 Then
        Dim objCliente As New clsBuscarCliente
        objCliente.ActivoFormularioBuscarClientes miConexion.TextoConexion(logComercio), True
        Me.Refresh
        If objCliente.BCClienteSeleccionado > 0 Then
            If objCliente.BCTipoClienteSeleccionado = TipoCliente.Cliente Then
                BuscoClienteID idPersona:=objCliente.BCClienteSeleccionado
            Else
                BuscoClienteID idEmpresa:=objCliente.BCClienteSeleccionado
            End If
        End If
        Set objCliente = Nothing
        If Val(lCCliente.Tag) <> 0 Then bConsultar.SetFocus
    End If
    Exit Sub
    
errorB:
    clsGeneral.OcurrioError "Ocurrió un error al procesar la información del cliente", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCCi_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(lCCliente.Tag) <> 0 Then bConsultar.SetFocus: Exit Sub
        If Trim(tCCi.Text) = "" Then tCRuc.SetFocus: Exit Sub
        If Len(tCCi.Text) = 7 Then tCCi.Text = clsGeneral.AgregoDigitoControlCI(tCCi.Text)
        
        'Valido la Cédula ingresada-------------------------------------------------------------------------------------------
        If Trim(tCCi.Text) <> "" Then
            If Len(tCCi.Text) <> 8 Then
                MsgBox "La cédula de identidad ingresada no es válida. Verifique", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            If Not clsGeneral.CedulaValida(tCCi.Text) Then
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        
        BuscoCliente Cedula:=tCCi.Text
        If Val(lCCliente.Tag) <> 0 Then bConsultar.SetFocus
    End If
    
End Sub

Private Sub tCRuc_Change()
    lCCliente.Tag = 0
End Sub

Private Sub tCRuc_GotFocus()
    tCRuc.SelStart = 0: tCRuc.SelLength = 15
End Sub

Private Sub tCRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo errorB
    If KeyCode = vbKeyF4 Then
        Dim objCliente As New clsBuscarCliente
        objCliente.ActivoFormularioBuscarClientes miConexion.TextoConexion(logComercio), Empresa:=True
        Me.Refresh
        If objCliente.BCTipoClienteSeleccionado = TipoCliente.Cliente Then
            BuscoClienteID idPersona:=objCliente.BCClienteSeleccionado
        Else
            BuscoClienteID idEmpresa:=objCliente.BCClienteSeleccionado
        End If
        Set objCliente = Nothing
        If Val(lCCliente.Tag) <> 0 Then bConsultar.SetFocus
    End If
    Exit Sub
    
errorB:
    clsGeneral.OcurrioError "Ocurrió un error al procesar la información del cliente", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCRuc_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(lCCliente.Tag) <> 0 Then bConsultar.SetFocus: Exit Sub
        If Trim(tCRuc.Text) = "" Then tCCi.SetFocus: Exit Sub
        
        If Len(tCRuc.Text) <> 12 Then
            MsgBox "La número de RUC ingresado no es correcto. Verifique", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        
        BuscoCliente Ruc:=tCRuc.Text
        bConsultar.SetFocus
    End If
    
End Sub

Private Sub tTitulo_Change()
    tTitulo.Tag = 0
    lClave.Visible = False
    lComentario.Caption = ""
    lEmail.Caption = "": lEmail.BackColor = &H800000
    vsConsulta.Rows = 1: vsConsulta1.Rows = 1: vsSaldo.Rows = 0
    bEmail.Tag = ""
End Sub

Private Sub tTitulo_GotFocus()
    tTitulo.SelStart = 0: tTitulo.SelLength = Len(tTitulo.Text)
End Sub

Private Sub tTitulo_KeyPress(KeyAscii As Integer)
Dim sCliente As String
Dim aSeleccionado As Long

    If KeyAscii = vbKeyReturn Then
        lComentario.Caption = ""
        If Val(tTitulo.Tag) <> 0 Then bConsultar.SetFocus: Exit Sub
        If Trim(tTitulo.Text) = "" Then Exit Sub
            
            On Error GoTo errLista
            Screen.MousePointer = 11
            Dim aLista As New clsListadeAyuda
            Cons = "Select ColCodigo, 'Título' = ColNombre, 'Fecha Civil' = ColFechaCivil, 'Fecha Iglesia' = ColFechaIglesia, " _
                            & " 'Cliente1' = (RTrim(P1.CPeNombre1) + ' ' + RTrim(P1.CPeApellido1)),  " _
                            & " 'Cliente2' = (RTrim(P2.CPeNombre1) + ' ' + RTrim(P2.CPeApellido1))  " _
                    & " From Colectivo left Outer Join CPersona P2 On ColCliente2 = P2.CPeCliente, CPersona P1" _
                    & " Where ColNombre Like '" & Trim(tTitulo.Text) & "%'" _
                    & " And ColCliente1 = P1.CPeCliente" _
                    & " Order by ColNombre"
            
            aLista.ActivoListaAyuda Cons, False, miConexion.TextoConexion(logComercio), 8500
            
            aSeleccionado = aLista.ValorSeleccionado
            If aSeleccionado <> 0 Then
                CargoColectivoPorID aSeleccionado
                bConsultar.SetFocus
            End If
            
            Set aLista = Nothing
            Screen.MousePointer = 0
    End If
    Exit Sub
    
errLista:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al activar la lista de ayuda. ", Err.Description

End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    If bCargarImpresion Then
        If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub
        With vsListado
            .StartDoc
            .MarginLeft = 750
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        Select Case cCuenta.ItemData(cCuenta.ListIndex)
            Case Cuenta.Personal:
                    aTexto = "Cta. Personal: "
                    If Trim(tCCi.Text) <> "" Then aTexto = aTexto & "(" & tCCi.FormattedText & ") "
                    If Trim(tCRuc.Text) <> "" Then aTexto = aTexto & "(" & tCRuc.FormattedText & ") "
                    aTexto = aTexto & Trim(lCCliente.Caption)
            Case Cuenta.Colectivo: aTexto = "Colectivo " & Trim(tTitulo.Text)
        End Select
        
        EncabezadoListado vsListado, "Consulta de Cuentas - " & aTexto, False
        vsListado.FileName = "Consulta de Cuentas"
        
        If vsConsulta.Rows > 1 Then
            vsListado.Paragraph = Trim(lRecibo.Caption)
            vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        End If
        If vsConsulta1.Rows > 1 Then
            vsListado.Paragraph = " ": vsListado.Paragraph = Trim(lFactura.Caption)
            vsConsulta1.ExtendLastCol = False: vsListado.RenderControl = vsConsulta1.hwnd: vsConsulta1.ExtendLastCol = True
        End If
        
        If vsSaldo.Rows > 0 Then
            vsListado.Paragraph = " ": vsListado.Paragraph = Trim(lSaldo.Caption)
            vsSaldo.ExtendLastCol = False: vsListado.RenderControl = vsSaldo.hwnd: vsSaldo.ExtendLastCol = True
        End If
        
        vsListado.EndDoc
        bCargarImpresion = False
    End If
    
    If Imprimir Then
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    Screen.MousePointer = 0
    
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Function ValidoCampos() As Boolean
    
    ValidoCampos = False
    If cCuenta.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de cuenta para realizar la consulta de datos.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    Select Case cCuenta.ItemData(cCuenta.ListIndex)
        Case Cuenta.Personal
                If Val(lCCliente.Tag) = 0 Then
                    MsgBox "Debe seleccionar el cliente de la cuenta personal para realizar la consulta.", vbExclamation, "ATENCIÓN"
                    tCCi.SetFocus: Exit Function
                End If
        Case Cuenta.Colectivo
                If Val(tTitulo.Tag) = 0 Then
                    MsgBox "Debe seleccionar el colectivo para realizar la consulta.", vbExclamation, "ATENCIÓN"
                    Foco tTitulo: Exit Function
                End If
    End Select
    
    ValidoCampos = True
    
End Function

Private Sub BuscoCliente(Optional Cedula As String = "", Optional Ruc As String = "")

    On Error GoTo errBuscar
    Screen.MousePointer = 11

    If Cedula <> "" Then
        Cons = "Select Cliente.*, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2)" _
                & " From Cliente, CPersona " _
                & " Where CliCiRuc = '" & Cedula & "'" _
                & " And CliCodigo = CPeCliente"
    End If
    
    If Ruc <> "" Then
        Cons = "Select Cliente.*, Nombre = (RTrim(CEmNombre) + RTrim(' (' + CEmFantasia) + ')')" _
                & " From Cliente, CEmpresa " _
                & " Where CliCiRuc = '" & Ruc & "'" _
                & " And CliCodigo = CEmCliente"
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Cedula <> "" Then tCRuc.Text = "" Else tCCi.Text = ""
        lCCliente.Tag = RsAux!CliCodigo
        lCCliente.Caption = " " & Trim(RsAux!Nombre)
        RsAux.Close
        bEmail.Tag = lCCliente.Tag
        lEmail.Caption = CargoDireccionEMail(lCCliente.Tag)
        If Trim(lEmail.Caption) = "" Then
            lEmail.BackColor = vbRed
            lEmail.Caption = " Falta Ingresar EMail ..."
        End If
    Else
        RsAux.Close
        MsgBox "No existe un cliente para la CI/Ruc, ingresado.", vbExclamation, "Cliente Inexistente"
        lCCliente.Caption = "": lCCliente.Tag = ""
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
End Sub

Private Sub BuscoClienteID(Optional idPersona As Long = 0, Optional idEmpresa As Long = 0)

    On Error GoTo errBuscar
    If idPersona = 0 And idEmpresa = 0 Then Exit Sub
    Screen.MousePointer = 11

    If idPersona <> 0 Then
        Cons = "Select Cliente.*, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2)" _
                & " From Cliente, CPersona " _
                & " Where CliCodigo = " & idPersona _
                & " And CliCodigo = CPeCliente"
    End If
    
    If idEmpresa <> 0 Then
        Cons = "Select Cliente.*, Nombre = (RTrim(CEmNombre) + RTrim(' (' + CEmFantasia) + ')')" _
                & " From Cliente, CEmpresa " _
                & " Where CliCodigo = " & idEmpresa _
                & " And CliCodigo = CEmCliente"
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        tCCi.Text = "": tCRuc.Text = ""
        If Not IsNull(RsAux!CliCiRuc) Then
            If idPersona <> 0 Then tCCi.Text = Trim(RsAux!CliCiRuc) Else tCRuc.Text = Trim(RsAux!CliCiRuc)
        End If
        
        lCCliente.Tag = RsAux!CliCodigo
        lCCliente.Caption = " " & Trim(RsAux!Nombre)
        RsAux.Close
        bEmail.Tag = lCCliente.Tag
        lEmail.Caption = CargoDireccionEMail(lCCliente.Tag)
        If Trim(lEmail.Caption) = "" Then
            lEmail.BackColor = vbRed
            lEmail.Caption = " Falta Ingresar EMail ..."
        End If
    Else
        RsAux.Close
        MsgBox "No existe un cliente para el código seleccionado.", vbExclamation, "Cliente Inexistente"
        lCCliente.Caption = "": lCCliente.Tag = "0"
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
End Sub

Private Sub BuscoClienteInicio(ByVal lID As Long)

    On Error GoTo errBuscar
    Screen.MousePointer = 11

    Cons = "Select * " _
        & " From Cliente " _
            & " Left Outer Join CPersona On CliCodigo = CPeCliente " _
            & " Left Outer Join CEmpresa On CliCodigo = CEmCliente " _
        & " Where CLiCodigo = " & lID
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
    
        tCCi.Text = "": tCRuc.Text = ""
        If Not IsNull(RsAux!CliCiRuc) Then
            If RsAux!CliTipo = TipoCliente.Cliente Then
                tCCi.Text = Trim(RsAux!CliCiRuc)
                lCCliente.Caption = " " & RTrim(RsAux!CPeApellido1)
                If Not IsNull(RsAux!CPeApellido2) Then
                    lCCliente.Caption = lCCliente.Caption & " " & RTrim(RsAux!CPeApellido2)
                End If
                lCCliente.Caption = lCCliente.Caption & ", " + RTrim(RsAux!CPeNombre1)
                If Not IsNull(RsAux!CPeNombre2) Then
                    lCCliente.Caption = lCCliente.Caption & " " & RTrim(RsAux!CPeNombre2)
                End If
            Else
                tCRuc.Text = Trim(RsAux!CliCiRuc)
                If Not IsNull(RsAux!CEmNombre) Then
                    lCCliente.Caption = RTrim(RsAux!CEmNombre)
                End If
                If Not IsNull(RsAux!CEmFantasia) Then
                    lCCliente.Caption = lCCliente.Caption & "(" & Trim(RsAux!CEmFantasia) & ")"
                End If
            End If
        End If
        
        lCCliente.Tag = RsAux!CliCodigo
        RsAux.Close
        
        bEmail.Tag = lCCliente.Tag
        lEmail.Caption = CargoDireccionEMail(lCCliente.Tag)
        If Trim(lEmail.Caption) = "" Then
            lEmail.BackColor = vbRed
            lEmail.Caption = " Falta Ingresar EMail ..."
        End If
    Else
        RsAux.Close
        MsgBox "No existe un cliente para el código seleccionado.", vbExclamation, "Cliente Inexistente"
        lCCliente.Caption = "": lCCliente.Tag = "0"
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
End Sub


Private Function CargoDireccionEMail(ByVal sCliente As String) As String
    
    CargoDireccionEMail = ""
    Cons = "Select * From EMailDireccion, EMailServer " _
        & " Where EMDIDCliente In (" & sCliente & ")" _
        & " And EMDServidor = EMSCodigo"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        If CargoDireccionEMail = "" Then
            CargoDireccionEMail = Trim(RsAux!EMDDireccion) & "@" & Trim(RsAux!EMSDireccion)
        Else
            CargoDireccionEMail = CargoDireccionEMail & "; ..."
            Exit Do
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
        
End Function

Private Sub CargoColectivoPorID(ByVal lCod As Long)
Dim sCliente As String

    Cons = "Select * from Colectivo Where ColCodigo = " & Val(lCod)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
    
        tTitulo.Text = Trim(RsAux!ColNombre)
        tTitulo.Tag = lCod
    
        bEmail.Tag = RsAux!ColCliente1
        sCliente = RsAux!ColCliente1 & "," & RsAux!ColCliente2
        If Not IsNull(RsAux!ColComentario) Then lComentario.Caption = Trim(RsAux!ColComentario)
        If Not IsNull(RsAux!ColClave) Then
            lClave.Caption = "Pedir Clave: '" & UCase(Trim(RsAux!ColClave)) & "'": lClave.Visible = True
        Else
            lClave.Caption = "": lClave.Visible = False
        End If
    End If
    RsAux.Close
    
    lEmail.Caption = CargoDireccionEMail(sCliente)
    If Trim(lEmail.Caption) = "" Then
        lEmail.BackColor = vbRed
        lEmail.Caption = " Falta Ingresar EMail ..."
    End If
    
    If Not lClave.Visible Then      'Hay que Sacar las cédulas
    
        Cons = "Select CI1 = P1.CliCiRuc, CI2 = P2.CliCiRuc " _
            & " From Colectivo left Outer Join Cliente P2 On ColCliente2 = P2.CliCodigo, Cliente P1" _
            & " Where ColCodigo = " & lCod _
            & " And ColCliente1 = P1.CliCodigo"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        lClave.Caption = ""
        If Not RsAux.EOF Then
            If Not IsNull(RsAux!CI1) Then lClave.Caption = "C.I.(1): " & clsGeneral.RetornoFormatoCedula(RsAux!CI1)
            If Not IsNull(RsAux!CI2) Then lClave.Caption = lClave.Caption & "               C.I.(2): " & clsGeneral.RetornoFormatoCedula(RsAux!CI2)
            lClave.Caption = Trim(lClave.Caption)
            If Trim(lClave.Caption) <> "" Then lClave.Visible = True
        End If
        RsAux.Close
    End If
    
End Sub
