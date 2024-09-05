VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmVerSolicitud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualización de Solicitudes"
   ClientHeight    =   5880
   ClientLeft      =   1590
   ClientTop       =   2805
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerSolicitud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9690
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
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
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   3
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
            ImageIndex      =   5
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
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Solicitud"
      ForeColor       =   &H00000080&
      Height          =   4035
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   9495
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   9
         Top             =   3660
         Width           =   8055
      End
      Begin VB.TextBox tSolicitud 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox tCi 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   900
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         ForeColor       =   12582912
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
      Begin MSMask.MaskEdBox tGarantia 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         ForeColor       =   12582912
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
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
         Height          =   1275
         Left            =   120
         TabIndex        =   10
         Top             =   2340
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   2249
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
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
      Begin VB.Label lFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10-Dic-1998"
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
         Left            =   3720
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lProceso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "La Solicitud se está Facturando !!!"
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
         Left            =   6240
         TabIndex        =   29
         Top             =   540
         Width           =   3135
      End
      Begin VB.Label lEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10-Dic-1998"
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
         Left            =   6240
         TabIndex        =   28
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Co&mentarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lPago 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10-Dic-1998"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de pago:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label lUsuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10-Dic-1998"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7800
         TabIndex        =   24
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lVendedor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10-Dic-1998"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5400
         TabIndex        =   23
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   7095
         TabIndex        =   22
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   255
         Left            =   4575
         TabIndex        =   21
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label lMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3720
         TabIndex        =   20
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label lDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Niagara 2345"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3720
         TabIndex        =   19
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   5655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3720
         TabIndex        =   17
         Top             =   900
         UseMnemonic     =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&R.U.C.:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&C.I. Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "C.I. &Garantía:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lGarantia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº de Solicitud:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   260
         Width           =   1215
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   260
         Width           =   615
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   1980
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   5625
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "sucursal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6773
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid lCondicion 
      Height          =   1035
      Left            =   120
      TabIndex        =   32
      Top             =   4560
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   1826
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
      HighLight       =   0
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8760
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
            Picture         =   "frmVerSolicitud.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerSolicitud.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerSolicitud.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerSolicitud.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerSolicitud.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerSolicitud.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerSolicitud.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerSolicitud.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerSolicitud.frx":10E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmVerSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sMoficicar As Boolean

Dim rs1 As rdoResultset
Dim RsAux As rdoResultset

Const FormatoCedula = "_.___.___-_"

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    InicializoGrilla
    LimpioFicha
    Ingreso False
    
    Botones False, False, False, False, False, Toolbar1, Me
    
    If Trim(Command()) <> "" Then
        If Val(Command()) <> 0 Then
            tSolicitud.Text = Trim(Command())
            BuscoSolicitud CLng(Command())
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    'GuardoSeteoForm Me
    
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End

End Sub

Private Sub AccionModificar()
    
    'Hay que validar si se puede
    On Error Resume Next
    If Not ValidoModEli Then Exit Sub
    
    Botones False, False, False, True, True, Toolbar1, Me
    Ingreso True
    tCi.SetFocus

End Sub

Private Sub AccionCancelar()
    
    Botones False, True, True, False, False, Toolbar1, Me
    Ingreso False
    LimpioFicha
    BuscoSolicitud CLng(tSolicitud.Text)
    Foco tSolicitud
    
End Sub

Private Sub AccionEliminar()
Dim aMsg As String: aMsg = ""

    If Not ValidoModEli Then Exit Sub
    
    If MsgBox("Confirma eliminar la solicitud Nº " & Trim(tSolicitud.Text), vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar") = vbNo Then Exit Sub
    If MsgBox("Está seguro que desea eliminarla.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirme") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    On Error GoTo ErrGFR
    cBase.BeginTrans
    
    On Error GoTo ErrResumo
    Cons = "Select * From Solicitud Where SolCodigo = " & CLng(tSolicitud.Text)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Verifico modificaciones de la  SOLICITUD-------------------------------------------------------------------------------------
    If RsAux!SolProceso <> Val(lProceso.Tag) Then
        aMsg = "Los datos de la solicitud se han modificado (proceso autorización). Vuelva a cargar los datos"
        GoTo ErrRelajo: Exit Sub
    End If
    
    If RsAux!SolEstado <> Val(lEstado.Tag) Then
        aMsg = "Los datos de la solicitud se han modificado (el estado de la solicitud). Vuelva a cargar los datos"
        GoTo ErrRelajo: Exit Sub
    End If
    
    If Not IsNull(RsAux!SolUsuarioR) Then
        aMsg = "Los datos de la solicitud se han modificado (está siendo analizada por otro usuario en decisión). Vuelva a cargar los datos"
        GoTo ErrRelajo: Exit Sub
    End If
    '---------------------------------------------------------------------------------------------------------------------------------------------
        
    Cons = "Delete RenglonSolicitud Where RSoSolicitud = " & CLng(tSolicitud.Text)
    cBase.Execute Cons
    
    Cons = "Delete Solicitud Where SolCodigo = " & CLng(tSolicitud.Text)
    cBase.Execute Cons
    
    cBase.CommitTrans                               'Fin TRANSACCION----------------------------------------------!!!!!!!!!!!!!!!!!!!
    
    Ingreso False
    LimpioFicha
    tSolicitud.Text = ""
    Botones False, False, False, False, False, Toolbar1, Me
    Screen.MousePointer = 0
    Exit Sub
    
ErrGFR:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción.", Err.Description
    Exit Sub
ErrResumo:
    Resume ErrRelajo
ErrRelajo:
    Screen.MousePointer = 0
    cBase.RollbackTrans
    If aMsg = "" Then aMsg = "Ocurrió un error al realizar la solicitud."
    clsGeneral.OcurrioError aMsg, Err.Description
End Sub

Private Sub Label16_Click()
    Foco tGarantia
End Sub

Private Sub Label2_Click()
    tRuc.SelStart = 0: tRuc.SelLength = 15: tRuc.SetFocus
End Sub

Private Sub Label26_Click()
    Foco tComentario
End Sub

Private Sub LimpioFicha()
    
    lFecha.Caption = ""
    lEstado.Caption = "": lProceso.Caption = ""
    
    lNombre.Caption = "": lNombre.Tag = 0
    lDireccion.Caption = "":
    lGarantia.Tag = 0: lGarantia.Caption = ""
    
    tRuc.Text = "": tCi.Text = FormatoCedula: tGarantia.Text = FormatoCedula
    
    lMoneda.Caption = ""
    lUsuario.Caption = ""
    lVendedor.Caption = ""
    lPago.Caption = ""
    
    tComentario.Text = ""
      
    vsLista.Rows = 1
    lCondicion.Rows = lCondicion.FixedRows

End Sub

Private Sub lProceso_DblClick()

On Error GoTo errFnc

    If MnuModificar.Enabled And IsNumeric(tSolicitud.Text) Then
        If Val(lProceso.Tag) = TipoResolucionSolicitud.Facturando Then
        
            If MsgBox("¿Cambia el proceso de la Solicitud a: 'Resolución Manual' ?", vbQuestion + vbYesNo + vbDefaultButton2, "Cambiar Proceso de Resolución") = vbYes Then
        
                Cons = "Select * from Solicitud Where SolCodigo = " & Val(tSolicitud.Text)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    If RsAux!SolProceso = Val(lProceso.Tag) Then
                        RsAux.Edit
                        RsAux!SolProceso = TipoResolucionSolicitud.Manual
                        RsAux.Update
                    End If
                End If
                RsAux.Close
                
               tSolicitud_KeyDown vbKeyReturn, 0
               
            End If
        End If
    End If
    Exit Sub
    
errFnc:
    clsGeneral.OcurrioError "Error al cambiar el proceso de resolución.", Err.Description
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

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub tCi_GotFocus()
    tCi.SelStart = 0: tCi.SelLength = Len(tCi.Text)
End Sub

Private Sub TCI_KeyPress(KeyAscii As Integer)

    On Error GoTo ErrTCK

    If KeyAscii = vbKeyReturn Then
        Dim aCi As String
        Screen.MousePointer = 11
        
        If Trim(tCi.Text) <> FormatoCedula Then       'Valido la Cédula ingresada----------
            aCi = clsGeneral.QuitoFormatoCedula(tCi.Text)
            If Len(aCi) <> 8 Or Not clsGeneral.CedulaValida(aCi) Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            
            BuscoClienteCI clsGeneral.QuitoFormatoCedula(tCi.Text), Titular:=True
            If Val(lNombre.Tag) <> 0 Then tGarantia.SetFocus
        Else
            tRuc.SetFocus
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
    
ErrTCK:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar la cédula de identidad."
End Sub

Private Sub BuscoClienteCI(Cedula As String, Optional Titular As Boolean = False, Optional Garantia As Boolean = False)
On Error GoTo errBuscar

    If Titular Then tRuc.Text = "": lDireccion.Caption = "": lNombre.Caption = ""
    If Cedula = "" Then Exit Sub
    Screen.MousePointer = 11

    Cons = " Select CliCodigo, CliCIRuc, CliDireccion, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
            & " From Cliente, CPersona " _
            & " Where CliCiRuc = '" & Cedula & "'" _
            & " And CliTipo = " & TipoCliente.Cliente _
            & " And CliCodigo = CPeCliente"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    If RsAux.EOF Then
        Screen.MousePointer = 0
        RsAux.Close
        MsgBox "No existe un cliente para la cédula de indentidad ingresada.", vbExclamation, "ATENCIÓN"
    Else        'El cliente ingresado existe-------------------
        If Titular Then
            If Not IsNull(RsAux!CliCIRuc) Then tCi.Text = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc)
            lNombre.Caption = " " & RsAux!Nombre
            lNombre.Tag = RsAux!CliCodigo
            
            If Not IsNull(RsAux!CliDireccion) Then lDireccion.Caption = " " & DireccionATexto(RsAux!CliDireccion)
        End If
        
        If Garantia Then
            If Not IsNull(RsAux!CliCIRuc) Then tGarantia.Text = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc)
            lGarantia.Caption = " " & RsAux!Nombre
            lGarantia.Tag = RsAux!CliCodigo
        End If
        
        RsAux.Close
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
End Sub



Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tGarantia_Change()
    tGarantia.Tag = "": lGarantia.Caption = ""
End Sub

Private Sub tGarantia_GotFocus()
    tGarantia.SelStart = 0: tGarantia.SelLength = Len(tGarantia.Text)
End Sub

Private Sub tGarantia_KeyPress(KeyAscii As Integer)

On Error GoTo errBuscar

    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = 11
        Dim aCi As String

        If Trim(tGarantia.Text) <> FormatoCedula Then           'Valido la Cédula ingresada----------
            aCi = clsGeneral.QuitoFormatoCedula(tGarantia.Text)
            If Len(aCi) <> 8 Or Not clsGeneral.CedulaValida(aCi) Then
                Screen.MousePointer = 0: lGarantia.Caption = ""
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            BuscoClienteCI clsGeneral.QuitoFormatoCedula(tGarantia.Text), , Garantia:=True
            If Val(lGarantia.Tag) <> 0 Then Foco tComentario
        Else
            Foco tComentario
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar la cédula de identidad."
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        
        Case "salir": Unload Me
    End Select

End Sub

Private Sub tRuc_GotFocus()
    tRuc.SelStart = 0: tRuc.SelLength = 15
End Sub

Private Sub tRuc_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tRuc.Text) <> "" Then
            tCi.Text = FormatoCedula
            BuscoEmpresaRuc Trim(tRuc.Text)
        Else
            tGarantia.SetFocus
        End If
    End If
    
End Sub

Private Sub BuscoEmpresaRuc(Codigo As String)

    On Error GoTo errBuscar

    Screen.MousePointer = 11
    tCi.Text = FormatoCedula: lDireccion.Caption = "": lNombre.Caption = ""
    Cons = "Select CliCodigo,CliCIRuc, CliCategoria, CEmFantasia, CEmNombre, CliDireccion from Cliente, CEmpresa" _
            & " Where CliCiRuc = '" & Codigo & "'" _
            & " And CliTipo = " & TipoCliente.Empresa _
            & " And CliCodigo = CEmCliente"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        Screen.MousePointer = 0: RsAux.Close
        MsgBox "No existe una empresa para el número de R.U.C. ingresado.", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else        'La empresa seleccionada Existe-----------
        
        If Not IsNull(RsAux!CliCIRuc) Then tRuc.Text = Trim(RsAux!CliCIRuc)
        lNombre.Caption = " " & Trim(RsAux!CEmFantasia)
        If Not IsNull(RsAux!CEmNombre) Then lNombre.Caption = lNombre.Caption & " (" & Trim(RsAux!CEmNombre) & ")"
        
        lNombre.Tag = RsAux!CliCodigo
        If Not IsNull(RsAux!CliDireccion) Then lDireccion.Caption = " " & DireccionATexto(RsAux!CliDireccion)
        
        tGarantia.SetFocus
        
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de la empresa.", Err.Description
End Sub

Private Sub AccionGrabar()

Dim aSolicitud As Long
Dim aMsg As String: aMsg = ""

    If Val(lNombre.Tag) = 0 Then
        MsgBox "Debe ingresar el titular de la solicitud de crédito.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("Confirma almacenar los datos ingresados en la solicitud.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    On Error GoTo ErrGFR
    cBase.BeginTrans
    
    On Error GoTo ErrResumo
    Cons = "Select * From Solicitud Where SolCodigo = " & CLng(tSolicitud.Text)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Verifico modificaciones de la  SOLICITUD-------------------------------------------------------------------------------------
    If RsAux!SolProceso <> Val(lProceso.Tag) Then
        aMsg = "Los datos de la solicitud se han modificado (proceso autorización). Vuelva a cargar los datos"
        GoTo ErrRelajo: Exit Sub
    End If
    
    If RsAux!SolEstado <> Val(lEstado.Tag) Then
        aMsg = "Los datos de la solicitud se han modificado (el estado de la solicitud). Vuelva a cargar los datos"
        GoTo ErrRelajo: Exit Sub
    End If
    
    If Not IsNull(RsAux!SolUsuarioR) Then
        aMsg = "Los datos de la solicitud se han modificado (está siendo analizada por otro usuario en decisión). Vuelva a cargar los datos"
        GoTo ErrRelajo: Exit Sub
    End If
    '---------------------------------------------------------------------------------------------------------------------------------------------
        
    RsAux.Edit
    RsAux!SolCliente = CLng(lNombre.Tag)
    
    If Val(lGarantia.Tag) <> 0 Then RsAux!SolGarantia = lGarantia.Tag Else RsAux!SolGarantia = Null
    If Trim(tComentario.Text) <> "" Then RsAux!SolComentarioS = Trim(tComentario.Text) Else RsAux!SolComentarioS = Null
    
    RsAux.Update: RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------
    
    cBase.CommitTrans                               'Fin TRANSACCION----------------------------------------------!!!!!!!!!!!!!!!!!!!
    
    Ingreso False
    BuscoSolicitud CLng(tSolicitud.Text)
    Screen.MousePointer = 0
    Exit Sub
    
ErrGFR:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción.", Err.Description
    Exit Sub

ErrResumo:
    Resume ErrRelajo
    
ErrRelajo:
    Screen.MousePointer = 0
    cBase.RollbackTrans
    If aMsg = "" Then aMsg = "Ocurrió un error al realizar la solicitud."
    clsGeneral.OcurrioError aMsg, Err.Description
End Sub

Private Sub tSolicitud_Change()
    Botones False, False, False, False, False, Toolbar1, Me
End Sub

Private Sub tSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Not IsNumeric(tSolicitud.Text) Then
            MsgBox "El número de solicitud ingresado no es correcto. Verifique.", vbExclamation, "ATENCIÓN"
            Foco tSolicitud: Exit Sub
        End If
        On Error GoTo errBusco
        LimpioFicha
        BuscoSolicitud CLng(tSolicitud.Text)
    End If
    Exit Sub

errBusco:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar la solicitud.", Err.Description
End Sub

Private Sub BuscoSolicitud(Codigo As Long)
    
    On Error GoTo errBuscar
    Screen.MousePointer = 11
    
    Cons = "Select * from Solicitud Where SolCodigo = " & Codigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        MsgBox "No existe una solicitud para el código: " & Codigo & ". Verifique.", vbExclamation, "Solicitud Inexistente"
        RsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If
    
    lFecha.Caption = Format(RsAux!SolFecha, "dd/mm/yyyy hh:mm")
    
    Select Case RsAux!SolEstado
        Case EstadoSolicitud.Aprovada: lEstado.Caption = "APROBADA"
        Case EstadoSolicitud.Condicional: lEstado.Caption = "CONDICIONAL"
        Case EstadoSolicitud.Rechazada: lEstado.Caption = "RECHAZADA"
        Case EstadoSolicitud.Pendiente: lEstado.Caption = "PENDIENTE"
        Case EstadoSolicitud.ParaRetomar: lEstado.Caption = "PARA RETOMAR"
        Case EstadoSolicitud.SinEfecto: lEstado.Caption = "SIN EFECTO"
    End Select
    lEstado.Tag = RsAux!SolEstado
    
    Select Case RsAux!SolFormaPago
        Case TipoPagoSolicitud.ChequeDiferido: lPago.Caption = " Con cheuqes dif."
        Case TipoPagoSolicitud.Efectivo: lPago.Caption = " En efectivo"
    End Select

    Select Case RsAux!SolProceso
        Case TipoResolucionSolicitud.Automatica:  lProceso.Caption = "Resolución Automática"
        Case TipoResolucionSolicitud.Manual: lProceso.Caption = "Resolución Manual"
        Case TipoResolucionSolicitud.Facturada: lProceso.Caption = "Solicitud Facturada"
        Case TipoResolucionSolicitud.Facturando: lProceso.Caption = "La Solicitud se está Facturando !!!"
    End Select
    lProceso.Tag = RsAux!SolProceso
    
    If Not IsNull(RsAux!SolComentarioS) Then tComentario.Text = Trim(RsAux!SolComentarioS)
    
    If Not IsNull(RsAux!SolUsuarioS) Then lUsuario.Caption = z_BuscoUsuario(RsAux!SolUsuarioS, Identificacion:=True)
    If Not IsNull(RsAux!SolVendedor) Then lVendedor.Caption = z_BuscoUsuario(RsAux!SolVendedor, Identificacion:=True)
    
    Cons = "Select * from Moneda Where MonCodigo = " & RsAux!SolMoneda
    Set rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rs1.EOF Then lMoneda.Caption = " " & Trim(rs1!MonSigno)
    rs1.Close
        
    CargoCliente RsAux!SolCliente, Titular:=True
    If Not IsNull(RsAux!SolGarantia) Then CargoCliente RsAux!SolGarantia, Garantia:=True
    
    RsAux.Close
    
    'Cargo datos de las resoluciones anteriores     ----------------------------------------------------------------------------
    Cons = "Select SolicitudResolucion.*, UsuIdentificacion, ConAbreviacion, dbo.ValorDentroDelResTexto(ResTexto) ValorResTexto " & _
                " From SolicitudResolucion " & _
                    " Left Outer Join Usuario ON ResUsuario = UsuCodigo " & _
                    " Left Outer Join CondicionResolucion ON ResCondicion = ConCodigo " & _
            " Where ResSolicitud = " & Codigo & _
            " Order by ResNumero DESC"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With lCondicion
            .AddItem "" ' |^Fecha|<Visto por ...|<Resolución
            .Cell(flexcpText, .Rows - 1, 0) = RsAux!ResNumero: .Cell(flexcpFontBold, .Rows - 1, 0) = True
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ResFecha, "d/mm hh:mm")
            If Not IsNull(RsAux!UsuIdentificacion) Then
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!UsuIdentificacion)
            Else
                .Cell(flexcpText, .Rows - 1, 2) = "Res.Autom."
            End If
            
            Dim sMemo As String, sAbrev As String, sValorRTxt As String
            Dim sTxt As String
            
            If Not IsNull(RsAux!ResComentario) Then sMemo = Trim(RsAux!ResComentario)
            If Not IsNull(RsAux!ConAbreviacion) Then sAbrev = Trim(RsAux!ConAbreviacion)
            If Not IsNull(RsAux!ValorResTexto) Then sValorRTxt = Trim(RsAux!ValorResTexto)
            
            If sMemo = sAbrev Then
                .Cell(flexcpText, .Rows - 1, 3) = sMemo & " " & sValorRTxt
            ElseIf sMemo <> "" Then
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ResComentario)
            ElseIf Not IsNull(RsAux!ConAbreviacion) Then
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ConAbreviacion)
            End If
            If (.Rows - .FixedRows) = 1 Then .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------
    
    
    CargoArticulos Codigo
    
    Botones False, True, True, False, False, Toolbar1, Me
    Screen.MousePointer = 0
    Exit Sub

errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar la solicitud.", Err.Description
End Sub

Private Sub CargoArticulos(Codigo As Long)
Dim aTotal As Currency

    On Error GoTo errRenglon
    aTotal = 0
    vsLista.Rows = 1
    Cons = "Select RenglonSolicitud.*, ArtNombre, TCuAbreviacion, TCuCantidad from RenglonSolicitud, Articulo, TipoCuota" _
            & " Where RSoSolicitud = " & Codigo _
            & " And RSoArticulo = ArtID " _
            & " And RSoTipoCuota = TCuCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    With vsLista
    
    Do While Not RsAux.EOF
        .AddItem CStr(RsAux!RSoCantidad)
        .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
        
        If Not IsNull(RsAux!RSoValorEntrega) Then .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!RSoValorEntrega, "#,##0.00")
        .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!TCuAbreviacion)
        .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!RSoValorCuota, "#,##0.00")
        
        If Not IsNull(RsAux!RSoValorEntrega) Then
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!RSoValorEntrega + RsAux!RSoValorCuota * RsAux!TCuCantidad, "#,##0.00")
        Else
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!RSoValorCuota * RsAux!TCuCantidad, "#,##0.00")
        End If
        
        aTotal = aTotal + .Cell(flexcpValue, .Rows - 1, 5)
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 2, , Colores.Obligatorio, , True, " "
    .Subtotal flexSTSum, -1, 4, , , , True, " "
    .Subtotal flexSTSum, -1, 5, , , , True, " "
    
    End With
    Exit Sub

errRenglon:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los articulos de la solicitud.", Err.Description
End Sub

Private Sub CargoCliente(Codigo As Long, Optional Titular As Boolean = False, Optional Garantia As Boolean = False)
Dim RsCli As rdoResultset

    If Titular Then     'Cargo los datos de Titular de la Operacion (cliente o empresa)
        
        Cons = "Select Cliente.*, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
               & " From Cliente, CPersona " _
               & " Where CliCodigo = " & Codigo _
               & " And CliCodigo = CPeCliente " _
                                    & " UNION ALL " _
               & " Select Cliente.*, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
               & " From Cliente, CEmpresa " _
               & " Where CliCodigo = " & Codigo _
               & " And CliCodigo = CEmCliente"
               
        Set RsCli = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsCli.EOF Then
            If RsCli!CliTipo = TipoCliente.Cliente Then
                If Not IsNull(RsCli!CliCIRuc) Then tCi.Text = clsGeneral.RetornoFormatoCedula(RsCli!CliCIRuc)
            Else
                If Not IsNull(RsCli!CliCIRuc) Then tRuc.Text = Trim(RsCli!CliCIRuc)
            End If
            
            lNombre.Caption = " " & Trim(RsCli!Nombre)
            lNombre.Tag = RsCli!CliCodigo
            
            If Not IsNull(RsCli!CliDireccion) Then lDireccion.Caption = " " & DireccionATexto(RsCli!CliDireccion)
        End If
        RsCli.Close
    
    End If
    
    If Garantia Then    'Cargo los datos de la garantia de la Operacion (solo cliente)
        
        Cons = "Select Cliente.*, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
               & " From Cliente, CPersona " _
               & " Where CliCodigo = " & Codigo _
               & " And CliCodigo = CPeCliente "
               
        Set RsCli = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsCli.EOF Then
            If Not IsNull(RsCli!CliCIRuc) Then tGarantia.Text = clsGeneral.RetornoFormatoCedula(RsCli!CliCIRuc)
            
            lGarantia.Caption = " " & Trim(RsCli!Nombre)
            lGarantia.Tag = RsCli!CliCodigo
        End If
        RsCli.Close
        
    End If
    
End Sub

Private Sub InicializoGrilla()
    
    On Error Resume Next
    With vsLista
        .Cols = 1: .Rows = 1:
        .FormatString = ">Q|Artículo|>Entrega|<Plan|>Cuota|>Total"
        .ExtendLastCol = True
        .WordWrap = True
        .ColWidth(0) = 550: .ColWidth(1) = 3200: .ColWidth(2) = 1200: .ColWidth(3) = 1200: .ColWidth(4) = 1300
    End With
      
    With lCondicion
        .Cols = 1: .Rows = 1:
        .ExtendLastCol = True: .WordWrap = False
        .FormatString = "|^Hora|<Visto por ...|<Resolución"
        .ColWidth(0) = 300: .ColWidth(1) = 1000: .ColWidth(2) = 1300: .ColWidth(3) = 6000
        .RowHeight(0) = 260
    End With

End Sub

Private Function Ingreso(Habilitado As Boolean)

    If Not Habilitado Then
        tSolicitud.Enabled = True
        
        tCi.Enabled = False: tRuc.Enabled = False: tGarantia.Enabled = False
        tComentario.Enabled = False: tComentario.BackColor = Colores.Inactivo
    Else
        tSolicitud.Enabled = False
        
        tCi.Enabled = True: tRuc.Enabled = True: tGarantia.Enabled = True
        tComentario.Enabled = True: tComentario.BackColor = Colores.Blanco
    End If
    
End Function

Private Function ValidoModEli() As Boolean
    
    ValidoModEli = False
    
    If Val(lProceso.Tag) = TipoResolucionSolicitud.Facturada Then
        MsgBox "La solicitud ya ha sido facturada. No podrá realizar esta acción.", vbExclamation, "Modificar/Eliminar"
        Exit Function
    End If
    
    If Val(lProceso.Tag) = TipoResolucionSolicitud.Facturando Then
        MsgBox "La solicitud se está facturando. No podrá realizar esta acción (vuelva a cargar los datos).", vbExclamation, "Modificar/Eliminar"
        Exit Function
    End If
    
    ValidoModEli = True
    
End Function
