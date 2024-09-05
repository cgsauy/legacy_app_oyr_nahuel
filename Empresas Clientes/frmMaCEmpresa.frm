VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmMaCEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresas Proveedores"
   ClientHeight    =   5850
   ClientLeft      =   1635
   ClientTop       =   3060
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaCEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9345
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   13
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
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
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cliente"
            Object.ToolTipText     =   "Ficha de Empresa."
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "help"
            Object.ToolTipText     =   "Ayuda (Ctrl+H)."
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "bBotones"
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   5700
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Otros Datos"
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
      Height          =   1605
      Left            =   120
      TabIndex        =   31
      Top             =   3960
      Width           =   9135
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   28
         Top             =   1260
         Width           =   7215
      End
      Begin VB.CheckBox cEstatal 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa E&statal:"
         Height          =   255
         Left            =   4440
         TabIndex        =   19
         Top             =   600
         Width           =   1635
      End
      Begin VB.CheckBox cTC 
         Alignment       =   1  'Right Justify
         Caption         =   "&T/C del día anterior:"
         Height          =   255
         Left            =   6360
         TabIndex        =   20
         Top             =   240
         Width           =   1875
      End
      Begin VB.CheckBox cCheque 
         Alignment       =   1  'Right Justify
         Caption         =   "O&pera c/Cheques:"
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox tAfiliado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   8040
         TabIndex        =   22
         Top             =   540
         Width           =   975
         _ExtentX        =   1720
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cCategoria 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cSubRubroC 
         Height          =   315
         Left            =   5880
         TabIndex        =   26
         Top             =   900
         Width           =   3135
         _ExtentX        =   5530
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
      Begin AACombo99.AACombo cRubroC 
         Height          =   315
         Left            =   1800
         TabIndex        =   24
         Top             =   900
         Width           =   2475
         _ExtentX        =   4366
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
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Rubro Conta&ble:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "S&ub Rubro:"
         Height          =   255
         Left            =   4485
         TabIndex        =   25
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Com&entarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1260
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda por &Defecto:"
         Height          =   375
         Left            =   6360
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ca&tegoría de Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lAfiliado 
         BackStyle       =   0  'Transparent
         Caption         =   "&Afiliado al Clearing Nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ficha"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   30
      Top             =   480
      Width           =   9135
      Begin VB.TextBox tCargoC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox tNombreC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1140
         Width           =   2895
      End
      Begin VB.TextBox tRazonSocial 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   3
         Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
         Top             =   540
         Width           =   4215
      End
      Begin VB.TextBox tNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   5
         Top             =   840
         Width           =   4215
      End
      Begin MSMask.MaskEdBox tRuc 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   6840
         TabIndex        =   7
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cRamo 
         Height          =   315
         Left            =   6840
         TabIndex        =   9
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
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
         Text            =   ""
      End
      Begin VB.Label lModificacion 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/2000 14:56"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7560
         TabIndex        =   42
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lAltaPor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alta:"
         Height          =   255
         Left            =   5580
         TabIndex        =   41
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label lModificadoPor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modificación:"
         Height          =   255
         Left            =   4860
         TabIndex        =   40
         Top             =   240
         Width           =   2595
      End
      Begin VB.Label lAlta 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/2000 14:56"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7560
         TabIndex        =   39
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "T&ipo Proveedor:"
         Height          =   255
         Left            =   5640
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Carg&o:"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Contacto:"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label lRuc 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº de R.U.C.:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lRazonSocial 
         BackStyle       =   0  'Transparent
         Caption         =   "&Razón Social:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label lRamo 
         BackStyle       =   0  'Transparent
         Caption         =   "&Giro / Ramo:"
         Height          =   255
         Left            =   5640
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lNombre 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom. &Fantasía:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   5595
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
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
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8361
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1740
      Left            =   120
      TabIndex        =   32
      Top             =   2220
      Width           =   9135
      Begin VB.ComboBox cEmail 
         Height          =   315
         Left            =   5100
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   180
         Width           =   3915
      End
      Begin VB.TextBox tDireccion 
         Appearance      =   0  'Flat
         Height          =   1125
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   540
         Width           =   4335
      End
      Begin VB.ComboBox cDireccion 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   180
         Width           =   4335
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsTelefono 
         Height          =   1130
         Left            =   4560
         TabIndex        =   35
         Top             =   540
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1993
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "E-&Mail:"
         Height          =   255
         Left            =   4560
         TabIndex        =   38
         Top             =   195
         Width           =   615
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":11D8
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
      Begin VB.Menu MnuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefrescar 
         Caption         =   "&Refrescar"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MnuBases 
      Caption         =   "&Bases"
      Begin VB.Menu MnuBx 
         Caption         =   "MnuBx"
         Index           =   0
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "&?"
      Begin VB.Menu MnuHelp 
         Caption         =   "&Ayuda"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmMaCEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim prmIDEmpresa As Long      'Propiedad para Setear el Cliente Seleccionado
Dim gFModificacion As Date  'Guardo la fecha para contrlar modificaciones

Dim sNuevo As Boolean, sModificar As Boolean
Dim aTexto As String

Dim RsCEm As rdoResultset       'BD Cliente
Dim RsEmp As rdoResultset       'BD CEmpresa

Dim aEmpresa As Long      'Guardo el id de cliente (empresa) p/almacenar en Tabla: CEmpresa

Private Sub cCategoria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cCheque.SetFocus
End Sub

Private Sub cCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cEstatal.SetFocus
End Sub

Private Sub cDireccion_Click()

    On Error GoTo errCargar
    If cDireccion.ListIndex <> -1 Then
        Dim miDir As Long
        miDir = cDireccion.ItemData(cDireccion.ListIndex)
        If miDir = 0 Then Exit Sub
        
        Screen.MousePointer = 11
        tDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, miDir, Departamento:=True, Localidad:=True, Zona:=True, EntreCalles:=True, Ampliacion:=True, ConfYVD:=True, ConEnter:=True)
        Screen.MousePointer = 0
    End If

errCargar:
    Screen.MousePointer = 0
End Sub


Private Sub cEmail_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo errEj
    If prmIDEmpresa = 0 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyA: If Shift = vbCtrlMask Then CargoCamposDesdeBDEMail prmIDEmpresa
        
        Case vbKeyF3: EjecutarApp prmPathApp & "\Emails", CStr(prmIDEmpresa)
    End Select
    
errEj:
End Sub

Private Sub cEstatal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cTC.SetFocus
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cRubroC
End Sub

Private Sub cRamo_GotFocus()
    cRamo.SelStart = 0: cRamo.SelLength = Len(cRamo.Text)
End Sub

Private Sub cRamo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tNombreC
End Sub

Private Sub cRubroC_Change()
    cSubRubroC.Clear
End Sub

Private Sub cRubroC_Click()
    cSubRubroC.Clear
End Sub

Private Sub cRubroC_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cRubroC.ListIndex = -1 Then Foco tComentario: Exit Sub
        If cSubRubroC.ListCount > 0 Then Foco cSubRubroC: Exit Sub
        
        On Error GoTo errCargar
        Screen.MousePointer = 11
        cons = "Select SRuID, SRunombre From SubRubro Where SRuRubro = " & cRubroC.ItemData(cRubroC.ListIndex) _
                & " Order by SRuNombre"
        CargoCombo cons, cSubRubroC
        Screen.MousePointer = 0
        Foco cSubRubroC
    End If
    Exit Sub

errCargar:
    clsGeneral.OcurrioError "Error al cargar los subrubros.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub cSubRubroC_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentario
End Sub


Private Sub cTC_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cMoneda
End Sub

Private Sub cTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cRamo
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error Resume Next
    sNuevo = False: sModificar = False
    LimpioFicha
    Botones True, False, False, False, False, Toolbar1, Me
    
    cons = "Select RamCodigo, RamNombre From Ramo Order by RamNombre"   'Cargo los RAMOS DE EMPRESAS
    CargoCombo cons, cRamo, ""
    cons = "Select CClCodigo, CClNombre From CategoriaCliente Order by CClNombre"    'Cargo las CategoriaCliente
    CargoCombo cons, cCategoria, ""
    cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"   'Cargo las monedas
    CargoCombo cons, cMoneda, ""

    'Cargo los tipos de proveedores de importacion
    cTipo.AddItem RetornoGiroEmpresa(GiroEmpresa.AgenciaDeTransporte)
    cTipo.ItemData(cTipo.NewIndex) = GiroEmpresa.AgenciaDeTransporte
    cTipo.AddItem RetornoGiroEmpresa(GiroEmpresa.Banco)
    cTipo.ItemData(cTipo.NewIndex) = GiroEmpresa.Banco
    cTipo.AddItem RetornoGiroEmpresa(GiroEmpresa.Aseguradoras)
    cTipo.ItemData(cTipo.NewIndex) = GiroEmpresa.Aseguradoras
    cTipo.AddItem RetornoGiroEmpresa(GiroEmpresa.Despachantes)
    cTipo.ItemData(cTipo.NewIndex) = GiroEmpresa.Despachantes
    cTipo.AddItem RetornoGiroEmpresa(GiroEmpresa.ProveedorDeMercaderia)
    cTipo.ItemData(cTipo.NewIndex) = GiroEmpresa.ProveedorDeMercaderia
    
    'Inicializo grilla Telefonos-----------------------------------------------------
    With vsTelefono
        .Rows = 1: .Cols = 1
        .FormatString = "<Tipo|<Teléfono|<Descripción"
        .ColWidth(0) = 1350: .ColWidth(1) = 1300
        .WordWrap = False: .MergeCells = flexMergeSpill: .ExtendLastCol = True
    End With
    '--------------------------------------------------------------------------------
    
    HabilitoIngreso False
    
    LoadME
    
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
       
        
End Sub

Private Sub LoadME()

    Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
    
    cons = "Select RubID, RubNombre From Rubro Order by RubNombre"
    CargoCombo cons, cRubroC
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If sNuevo Or sModificar Then
        If MsgBox("Ud. realizó modificaciones en la ficha y no ha grabado." & Chr(13) _
            & "Desea almacenar la información ingresada.", vbYesNo + vbExclamation, "ATENCIÓN") = vbYes Then
            
            AccionGrabar
            
            If sNuevo Or sModificar Then
                Cancel = True
                Exit Sub
            End If
        
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End
End Sub

Private Sub Label2_Click()
    Foco tComentario
End Sub

Private Sub Label3_Click()
    Foco cCategoria
End Sub

Private Sub lNombre_Click()
    Foco tNombre
End Sub

Private Sub lRamo_Click()
    Foco cRamo
End Sub

Private Sub lRazonSocial_Click()
    Foco tRazonSocial
End Sub

Private Sub lRuc_Click()
    tRuc.SelStart = 0: tRuc.SelLength = 15: tRuc.SetFocus
End Sub


Public Function prmColorBase(keyConn As String)
    On Error GoTo errColor
    For I = MnuBx.LBound To MnuBx.UBound
        If LCase(Trim(MnuBx(I).Tag)) = LCase(keyConn) Then
            Dim arrC() As String
            arrC = Split(MnuBases.Tag, "|")
            If arrC(I) <> "" Then Me.BackColor = arrC(I) Else Me.BackColor = vbButtonFace
            Exit For
        End If
    Next
    
    Frame2.BackColor = Me.BackColor
    Frame1.BackColor = Me.BackColor
    Frame3.BackColor = Me.BackColor
    cCheque.BackColor = Me.BackColor: cTC.BackColor = Me.BackColor: cEstatal.BackColor = Me.BackColor
errColor:
End Function

Private Sub MnuBx_Click(Index As Integer)
On Error Resume Next

    AccionCancelar
    prmIDEmpresa = 0
    LimpioFicha
    Botones True, False, False, False, False, Toolbar1, Me
    
    If Not AccionCambiarBase(MnuBx(Index).Tag, MnuBx(Index).Caption) Then Exit Sub
    Screen.MousePointer = 11
    
    CargoParametrosImportaciones
    LoadME
   
    'Cambio el Color del fondo de controles ----------------------------------------------------------------------------------------
    prmColorBase Trim(MnuBx(Index).Tag)
    '-------------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0

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

Private Sub MnuHelp_Click()
On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    cons = "Select * from Aplicacion Where AplNombre = '" & App.Title & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux!AplHelp) Then aFile = Trim(rsAux!AplHelp)
    rsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
    tRuc.SetFocus
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub tAfiliado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cCheque.SetFocus
End Sub

Private Sub tCargoC_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cCategoria.SetFocus
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tNombre_GotFocus()
    Foco tNombre
End Sub

Private Sub tNombre_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not sNuevo And Not sModificar Then
        If KeyCode = vbKeyF1 And Trim(tNombre.Text) <> "" Then
            If Not clsGeneral.TextoValido(tNombre.Text) Then MsgBox "Se han ingresado caracteres no válidos.", vbExclamation, "ATENCIÓN": Exit Sub
            
            cons = " Select ID_Cliente = CEmCliente, 'Razón Social' = CEmNombre, 'Nombre Fantasía' = CEmFantasia, Ramo = RamNombre " _
                    & " From CEmpresa, Ramo" _
                    & " Where CEmFantasia like '" & Trim(tNombre.Text) & "%'" _
                    & " And CEmRamo *= RamCodigo " _
                    & " Order by CEmFantasia"
            AyudaEmpresa cons
        End If
    End If

End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        tNombre.Text = NombreEmpresa(tNombre.Text, False)
        If cTipo.Enabled Then Foco cTipo Else Foco tRuc
    End If
    
End Sub

Private Sub tNombreC_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tCargoC
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

    Select Case Button.Key
        
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        
        Case "cliente": EjecutarApp prmPathApp & "\Clientes.exe", "2:" & CStr(prmIDEmpresa)
        Case "help": Call MnuHelp_Click
        Case "salir": Unload Me
            
    End Select

End Sub

Private Sub AccionGrabar()
    
Dim aError As String: aError = ""
    
    On Error GoTo errorBT
    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar los datos ingresados en la ficha.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    FechaDelServidor
    
    'SOLO se pueden modificar los datos en esta aplicación  !!!!
    

    cBase.BeginTrans    'COMIENZO TRANSACCION-------------------------------------------------------------------------
    On Error GoTo errorET
    
    cons = "Select * from Cliente Where CliCodigo = " & prmIDEmpresa                  'Tabla Cliente  ---------------------
    Set RsCEm = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    'Controlo Modificacion Multiusuario -----------------------------------------------------
    If gFModificacion <> RsCEm!CliModificacion Then
        aError = "La ficha ha sido modificada por otro usuario." & vbCrLf & _
                     "Verifique los datos antes de grabar."
        GoTo errorET: Exit Sub
    End If
    
    RsCEm.Edit
    RsCEm!CliModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsCEm!CliUsuario = paCodigoDeUsuario
    If cCheque.Value = 0 Then RsCEm!CliCheque = "N" Else: RsCEm!CliCheque = "S"
    If cCategoria.ListIndex <> -1 Then RsCEm!CliCategoria = cCategoria.ItemData(cCategoria.ListIndex) Else: RsCEm!CliCategoria = Null

    RsCEm.Update: RsCEm.Close
    '-------------------------------------------------------------------------------------------------------------------------------
    
    CargoCamposBDCEmpresa prmIDEmpresa
    CargoCamposBDEmpresaDato prmIDEmpresa
    
    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
    
    Botones True, True, True, False, False, Toolbar1, Me

    sNuevo = False: sModificar = False
    gFModificacion = gFechaServidor
    lModificacion.Caption = Format(gFechaServidor, "dd/mm/yyyy hh:mm")
        
    HabilitoIngreso False
    Screen.MousePointer = 0
    tRuc.SetFocus
    Exit Sub

errorBT:
    If aError = "" Then aError = "No se ha podido inicializar la transacción. Reintente la operación."
    clsGeneral.OcurrioError aError, Err.Description
    Screen.MousePointer = 0
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    If aError = "" Then aError = "No se ha podido inicializar la transacción. Reintente la operación."
    clsGeneral.OcurrioError aError, Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub AccionEliminar()

    'PREGUNTO PARA ELIIMINAR----------------------------------------------------------------------------------
    'Valido eliminación del cliente
    If Not ValidoEliminacion(prmIDEmpresa) Then Exit Sub
    
    If MsgBox("Confirma eliminar la empresa seleccionada." & vbCrLf & vbCrLf & _
                   "Sólo se eliminaran los datos de las tablas de proveedores.", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Datos") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    On Error GoTo errEliminar

    'Borro los datos de la tabla EmpresaDato
    cons = "Delete EmpresaDato Where EDaTipoEmpresa = " & TipoEmpresa.Cliente & " And EDaCodigo = " & prmIDEmpresa
    cBase.Execute (cons)
    
    BuscoDatosEmpresa prmIDEmpresa
    Screen.MousePointer = 0
    Exit Sub

errEliminar:
    clsGeneral.OcurrioError "Error al eliminar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValidoEliminacion(idCliente As Long) As Boolean
Dim bHay As Boolean: bHay = False

    Screen.MousePointer = 11
    On Error Resume Next
    ValidoEliminacion = False
    'Tabla Compra
    cons = "Select * from Compra where ComProveedor = " & idCliente
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True: rsAux.Close
    If bHay Then
        MsgBox "Hay compras o gastos que hacen referencia al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    'Tabla RemitoCompra
    cons = "Select * from RemitoCompra where RCoProveedor = " & idCliente
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True: rsAux.Close
    If bHay Then
        MsgBox "Hay compras o gastos que hacen referencia al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    'Tabla Sucursales y cheques
    cons = "Select * from SucursalDeBanco where SBaBanco = " & idCliente
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True: rsAux.Close
    If bHay Then
        MsgBox "Hay sucursales de banco referenciadas al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    cons = "Select * from ChequeDiferido where CDiBanco = " & idCliente
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True: rsAux.Close
    If bHay Then
        MsgBox "Hay cheques diferidos referenciados al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    'Informacion de carpetas
    cons = "Select * from Carpeta Where CarBcoEmisor = " & idCliente
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True: rsAux.Close
    If bHay Then
        MsgBox "Hay carpetas con bancos emisores que hacen referencia al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    cons = "Select * from Embarque Where EmbAgencia = " & idCliente
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True: rsAux.Close
    If bHay Then
        MsgBox "Hay embarques con agencias que hacen referencia al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    ValidoEliminacion = True
Salir:
    Screen.MousePointer = 0
End Function

Public Sub AccionNuevo()
On Error GoTo errNuevo
Dim idSel As Long

    Screen.MousePointer = 11
    Botones True, False, False, False, False, Toolbar1, Me
    prmIDEmpresa = 0
    LimpioFicha
    
    Dim objCliente As New clsCliente
    objCliente.Empresas 0, True
    Me.Refresh
    
    idSel = objCliente.IDIngresado
    Set objCliente = Nothing
    
    If idSel <> 0 Then
        BuscoDatosEmpresa idSel
        If prmIDEmpresa <> 0 Then AccionModificar
    End If
    
    Screen.MousePointer = 0
    Exit Sub

errNuevo:
    clsGeneral.OcurrioError "Error al ingresar el nuevo cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionModificar()

On Error GoTo errModificar

    Screen.MousePointer = 11
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    
    sModificar = True
    
    cTipo.SetFocus
    Screen.MousePointer = 0
    Exit Sub

errModificar:
    clsGeneral.OcurrioError "Error al cargar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionCancelar()

    On Error Resume Next
    Screen.MousePointer = 11
    HabilitoIngreso False

    LimpioFicha
    Botones True, False, False, False, False, Toolbar1, Me
    
    If sModificar Then BuscoDatosEmpresa prmIDEmpresa
    
    sNuevo = False: sModificar = False
    tRuc.SetFocus
    Screen.MousePointer = 0
    
End Sub


Private Sub CargoCamposBDCEmpresa(idEmpresa As Long)

    cons = "Select * from CEmpresa Where CEmCliente = " & idEmpresa
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    rsAux.Edit
    If cRamo.ListIndex <> -1 Then rsAux!CEmRamo = cRamo.ItemData(cRamo.ListIndex) Else rsAux!CEmRamo = Null
    If cEstatal.Value = 0 Then rsAux!CEmEstatal = False Else rsAux!CEmEstatal = True
    rsAux.Update
    
    rsAux.Close
    
End Sub

Private Sub CargoCamposBDEmpresaDato(idEmpresa As Long)

    cons = "Select * from EmpresaDato" _
           & " Where EDaTipoEmpresa = " & TipoEmpresa.Cliente _
           & " And EDaCodigo = " & idEmpresa
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then rsAux.AddNew Else rsAux.Edit
    
    rsAux!EDaTipoEmpresa = TipoEmpresa.Cliente
    rsAux!EDaCodigo = idEmpresa
    rsAux!EDaRubro = cTipo.ItemData(cTipo.ListIndex)
    
    If Trim(tNombreC.Text) <> "" Then rsAux!EDaContacto = Trim(tNombreC.Text) Else rsAux!EDaContacto = Null
    If Trim(tCargoC.Text) <> "" Then rsAux!EDaCargoContacto = Trim(tCargoC.Text) Else rsAux!EDaCargoContacto = Null
    
    rsAux!EDaMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    If cTC.Value = vbChecked Then rsAux!EDaTCAnterior = True Else rsAux!EDaTCAnterior = False
    If Trim(tComentario.Text) <> "" Then rsAux!EDaComentario = Trim(tComentario.Text) Else rsAux!EDaComentario = Null
    
    If cSubRubroC.ListIndex <> -1 Then rsAux!EDaSRubroContable = cSubRubroC.ItemData(cSubRubroC.ListIndex) Else rsAux!EDaSRubroContable = Null
    
    rsAux.Update: rsAux.Close
    
End Sub

Private Function ValidoCampos()

    ValidoCampos = False
    
    If (cRamo.ListIndex = -1 And Trim(cRamo.Text) <> "") Or (cCategoria.ListIndex = -1 And Trim(cCategoria.Text) <> "") Then
        MsgBox "Los datos ingresados no son correctos o la ficha está incompleta. Verifique", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If cTipo.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de proveedor de importaciones (campo obligatorio).", vbExclamation, "ATENCIÓN"
        Foco cTipo: Exit Function
    End If
    
    If cCategoria.ListIndex = -1 Then
        MsgBox "Debe seleccionar la categoría de cliente (campo obligatorio).", vbExclamation, "ATENCIÓN"
        Foco cCategoria: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda por defecto para los comprobantes de la empresa.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    
    ValidoCampos = True
    
End Function

Private Sub HabilitoIngreso(Optional bSi As Boolean = True)
    
    tNombre.Enabled = Not bSi
    tRazonSocial.Enabled = Not bSi
    tRuc.Enabled = Not bSi
    
    If bSi Then
        tRuc.BackColor = Colores.Inactivo
        tNombre.BackColor = Colores.Inactivo
        tRazonSocial.BackColor = Colores.Inactivo
        
        cRamo.BackColor = Blanco
        cTipo.BackColor = Colores.Obligatorio
        tNombreC.BackColor = Blanco
        tCargoC.BackColor = Blanco
        
        cCategoria.BackColor = Obligatorio
        tAfiliado.Enabled = False: tAfiliado.BackColor = Colores.Inactivo
        cMoneda.BackColor = Obligatorio
        tComentario.BackColor = Blanco
        
        cRubroC.BackColor = Colores.Blanco
        cSubRubroC.BackColor = Colores.Blanco
    Else
        tRuc.BackColor = Colores.Blanco
        tNombre.BackColor = Colores.Blanco
        tRazonSocial.BackColor = Colores.Blanco
        
        cRamo.BackColor = Inactivo
        cTipo.BackColor = Inactivo
        tNombreC.BackColor = Inactivo
        tCargoC.BackColor = Inactivo
    
        cCategoria.BackColor = Inactivo
        tAfiliado.Enabled = False: tAfiliado.BackColor = Inactivo
        cMoneda.BackColor = Inactivo
        tComentario.BackColor = Inactivo
        
        cRubroC.BackColor = Inactivo
        cSubRubroC.BackColor = Inactivo
    End If
    
    cCategoria.Enabled = bSi
    tNombreC.Enabled = bSi: tCargoC.Enabled = bSi
    cTipo.Enabled = bSi: cRamo.Enabled = bSi
    cCheque.Enabled = bSi
    cEstatal.Enabled = bSi
    cMoneda.Enabled = bSi
        
    tComentario.Enabled = bSi
    cRubroC.Enabled = bSi
    cSubRubroC.Enabled = bSi
    cTC.Enabled = bSi
    
End Sub

Private Sub LimpioFicha()

    tRuc.Text = ""
    tRazonSocial.Text = ""
    tNombre.Text = ""
    cRamo.Text = ""
    cTipo.Text = ""
    tNombreC.Text = ""
    tCargoC.Text = ""
    
    tDireccion.Text = ""
    cEmail.Clear
    
    vsTelefono.Rows = 1
    
    cCheque.Value = vbUnchecked
    cEstatal.Value = vbUnchecked
    cTC.Value = vbUnchecked
    cMoneda.Text = ""
    cCategoria.Text = ""
    tAfiliado.Text = ""
    tComentario.Text = ""
    
    lModificadoPor.Caption = "Modificación:": lAltaPor.Caption = "Alta:"
    lModificacion.Caption = "": lAlta.Caption = ""
    
    cRubroC.Text = "": cSubRubroC.Text = ""
        
End Sub

Private Sub tRazonSocial_GotFocus()
    tRazonSocial.SelStart = 0: tRazonSocial.SelLength = Len(tRazonSocial.Text)
End Sub

Private Sub tRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not sNuevo And Not sModificar Then
        If KeyCode = vbKeyF1 And Trim(tRazonSocial.Text) <> "" Then
            If Not clsGeneral.TextoValido(tRazonSocial.Text) Then MsgBox "Se han ingresado caracteres no válidos.", vbExclamation, "ATENCIÓN": Exit Sub
            
            cons = " Select ID_Cliente = CEmCliente, 'Razón Social' = CEmNombre, 'Nombre Fantasía' = CEmFantasia, Ramo = RamNombre " _
                    & " From CEmpresa, Ramo" _
                    & " Where CEmNombre like '" & Trim(tRazonSocial.Text) & "%'" _
                    & " And CEmRamo *= RamCodigo" _
                    & " Order by CEmNombre"
            AyudaEmpresa cons
        End If
    End If
    
End Sub

Private Sub tRazonSocial_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        On Error GoTo errX
        If Trim(tRazonSocial.Text) <> "" Then
            tRazonSocial.Text = NombreEmpresa(tRazonSocial.Text, True)
            tRazonSocial.Tag = NombreEmpresa(tRazonSocial.Text, False)
        End If
        tNombre.SetFocus
    End If
    Exit Sub
    
errX:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al procesar la información."
End Sub

Private Sub tRuc_GotFocus()
    tRuc.SelStart = 0: tRuc.SelLength = 15
End Sub

Private Sub tRuc_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Not sNuevo And Not sModificar Then
            'Busco la Empresa------------------------
            If Trim(tRuc.Text) <> "" Then BuscoDatosEmpresa 0, Trim(tRuc.Text) Else tRazonSocial.SetFocus
        Else
            'Valido que no exista------------------------
            If Trim(tRuc.Text) <> "" Then
                If sNuevo Then
                    ValidoEmpresaRuc tRuc.Text
                Else
                    If Trim(tRuc.Text) <> tRuc.Tag Then ValidoEmpresaRuc tRuc.Text Else tRazonSocial.SetFocus
                End If
            Else
                tRazonSocial.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub CargoCamposDesdeBDCliente()

    prmIDEmpresa = RsCEm!CliCodigo
    gFModificacion = RsCEm!CliModificacion
    
    If Not IsNull(RsCEm!CliCIRuc) Then tRuc.Text = Trim(RsCEm!CliCIRuc): tRuc.Tag = Trim(RsCEm!CliCIRuc)
    
    If Not IsNull(RsCEm!CliCheque) Then If RsCEm!CliCheque = "S" Then cCheque.Value = vbChecked
    If Not IsNull(RsCEm!CliCategoria) Then BuscoCodigoEnCombo cCategoria, RsCEm!CliCategoria
    
    
    lAlta.Caption = Format(RsCEm!CliAlta, "dd/mm/yyyy hh:mm")
    lModificacion.Caption = Format(RsCEm!CliModificacion, "dd/mm/yyyy hh:mm")
    If Not IsNull(RsCEm!UsuIdentificacion) Then lModificadoPor.Caption = "Mod. por " & Trim(RsCEm!UsuIdentificacion) & ":"
    
    If Not IsNull(RsCEm!CliUsuAlta) Then lAltaPor.Caption = "Ing. por " & Trim(BuscoUsuario(RsCEm!CliUsuAlta, Identificacion:=True)) & ":"
    
End Sub

Private Sub CargoCamposDesdeBDCEmpresa(idEmpresa As Long)

    cons = "Select * From CEmpresa Where CEmCliente = " & idEmpresa
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    tNombre.Text = Trim(rsAux!CEmFantasia)
    If Not IsNull(rsAux!CEmRamo) Then BuscoCodigoEnCombo cRamo, rsAux!CEmRamo
    If Not IsNull(rsAux!CEmNombre) Then tRazonSocial.Text = Trim(rsAux!CEmNombre)
    
    If rsAux!CEmEstatal Then cEstatal.Value = vbChecked
    If Not IsNull(rsAux!CEmAfiliado) Then tAfiliado.Text = Trim(rsAux!CEmAfiliado)
    
    rsAux.Close

End Sub

Private Sub CargoCamposDesdeBDEmpresaDato(idEmpresa As Long)

    cons = " Select * From EmpresaDato " _
            & " Where EDaCodigo = " & idEmpresa _
            & " And EDaTipoEmpresa = " & TipoEmpresa.Cliente
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!EDaRubro) Then BuscoCodigoEnCombo cTipo, rsAux!EDaRubro
        If Not IsNull(rsAux!EDaContacto) Then tNombreC.Text = Trim(rsAux!EDaContacto)
        If Not IsNull(rsAux!EDaCargoContacto) Then tCargoC.Text = Trim(rsAux!EDaCargoContacto)
        
        If rsAux!EDaTCAnterior Then cTC.Value = vbChecked
        If Not IsNull(rsAux!EDaMoneda) Then BuscoCodigoEnCombo cMoneda, rsAux!EDaMoneda
        If Not IsNull(rsAux!EDaComentario) Then tComentario.Text = Trim(rsAux!EDaComentario)
        
        If Not IsNull(rsAux!EDaSRubroContable) Then
            Dim rsRub As rdoResultset
            cons = "Select * from SubRubro Where SRuId = " & rsAux!EDaSRubroContable
            Set rsRub = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rsRub.EOF Then
                BuscoCodigoEnCombo cRubroC, rsRub!SRuRubro
                If cRubroC.ListIndex <> -1 Then
                    cons = "Select SRuID, SRuNombre From SubRubro Where SRuRubro = " & cRubroC.ItemData(cRubroC.ListIndex) _
                            & " Order by SRuNombre"
                    CargoCombo cons, cSubRubroC
                    BuscoCodigoEnCombo cSubRubroC, rsRub!SRuID
                End If
            End If
            rsRub.Close
        End If
    End If
    
    rsAux.Close

End Sub

Private Sub CargoCamposDesdeBDDireccion(idDireccionPpal As Long)
    
    cDireccion.Clear
    cDireccion.BackColor = Colores.Inactivo: tDireccion.BackColor = Colores.Inactivo
    'If gCliente <> 0 Then cDireccion.AddItem "(Agregar Nueva)": cDireccion.ItemData(cDireccion.NewIndex) = -2
    
    cDireccion.AddItem "Dirección Principal": cDireccion.ItemData(cDireccion.NewIndex) = idDireccionPpal
        
    tDireccion.Text = ""
        
    Dim rsDA As rdoResultset
    cons = "Select * from DireccionAuxiliar Where DAuCliente = " & prmIDEmpresa & _
               " Order by DAuNombre"
    Set rsDA = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsDA.EOF Then
        Do While Not rsDA.EOF
            cDireccion.AddItem Trim(rsDA!DAuNombre)
            cDireccion.ItemData(cDireccion.NewIndex) = rsDA!DAuDireccion
            rsDA.MoveNext
        Loop
        
        cDireccion.BackColor = Colores.Blanco: tDireccion.BackColor = Colores.Blanco
    End If
    rsDA.Close
    
    BuscoCodigoEnCombo cDireccion, idDireccionPpal
    
    tDireccion.Refresh
    
End Sub

Private Sub AyudaEmpresa(Consulta As String)
    
    On Error GoTo errBuscar
    Screen.MousePointer = 11
    Dim aIdSel As Long
    Dim aLista As New clsListadeAyuda
    
    aIdSel = 0
    aLista.ActivoListaAyudaSQL cBase, Consulta
    Me.Refresh
    If Trim(aLista.ItemSeleccionadoSQL) <> "" Then aIdSel = aLista.ItemSeleccionadoSQL
    
    Set aLista = Nothing
    
    If aIdSel > 0 Then
        BuscoDatosEmpresa aIdSel
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los datos de la empresa.", Err.Description
    If Not sNuevo And Not sModificar Then Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub BuscoDatosEmpresa(Codigo As Long, Optional miRuc As String = "")

    On Error GoTo errCargar
    Dim miIdDireccion As Long
    
    Screen.MousePointer = 11
    LimpioFicha
    
    cons = "Select * from Cliente Left Outer Join Usuario on CliUsuario = UsuCodigo " & _
              " Where CliTipo = " & TipoCliente.Empresa
    
    If Codigo > 0 Then
        cons = cons & " And CliCodigo = " & Codigo
    Else
        cons = cons & " And CliCiRuc = '" & miRuc & "'"
    End If
    
    Set RsCEm = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsCEm.EOF Then
        prmIDEmpresa = RsCEm!CliCodigo
        CargoCamposDesdeBDCliente
        If Not IsNull(RsCEm!CliDireccion) Then miIdDireccion = RsCEm!CliDireccion Else miIdDireccion = 0
        CargoCamposDesdeBDCEmpresa prmIDEmpresa
        CargoCamposDesdeBDTelefono prmIDEmpresa
        CargoCamposDesdeBDEmpresaDato prmIDEmpresa
        Botones True, True, True, False, False, Toolbar1, Me
    Else
        prmIDEmpresa = 0
        Screen.MousePointer = 0
        MsgBox "La empresa seleccionada ha sido eliminada. Verifique", vbExclamation, "ATENCIÓN"
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    RsCEm.Close
    
    If prmIDEmpresa <> 0 Then
        CargoCamposDesdeBDDireccion miIdDireccion
        CargoCamposDesdeBDEMail prmIDEmpresa
    End If
    
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los datos de la empresa.", Err.Description
End Sub

Private Sub CargoCamposDesdeBDTelefono(idCliente As Long)

    On Error GoTo ErrNT
    
    cons = "Select TelTipo, TelNumero, TelInterno, TTeCodigo, TTeNombre From Telefono, TipoTelefono " _
            & " Where TelCliente = " & idCliente _
            & " And TelTipo = TTeCodigo"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Dim aValor As Long
    
    With vsTelefono
        .Rows = 1
        Do While Not rsAux.EOF
            .AddItem Trim(rsAux!TTeNombre)
            aValor = rsAux!TTeCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!TelNumero)
            If Not IsNull(rsAux!TelInterno) Then .Cell(flexcpText, .Rows - 1, 2) = Trim(rsAux!TelInterno)
    
            rsAux.MoveNext
        Loop
    End With
    rsAux.Close
    Exit Sub
        
ErrNT:
    clsGeneral.OcurrioError "Error al cargar los números de teléfonos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoCamposDesdeBDEMail(Cliente As Long)

    On Error GoTo ErrNT
    Screen.MousePointer = 11
    
    cons = "Select EMDCodigo, RTrim(EMDDireccion) + '@' + EMSDireccion as Direccion" & _
                " From EMailDireccion, EMailServer" & _
                " Where EMDIdCliente = " & Cliente & _
                " And EMDServidor = EMSCodigo"

    CargoCombo cons, cEmail
    If cEmail.ListCount = 0 Then
        cEmail.AddItem " « « Oprima [F3] para agregar nuevo e-Mail » »"
        cEmail.ForeColor = Colores.Rojo
    Else
        cEmail.ForeColor = vbWindowText
    End If
    cEmail.ListIndex = 0
    
    cEmail.Enabled = True
    Exit Sub
        
ErrNT:
    clsGeneral.OcurrioError "Error al cargar la lista de e-mails.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub ValidoEmpresaRuc(Codigo As String)

On Error GoTo errBuscar

    Dim aCodCli As Long: aCodCli = 0
    
    If Codigo = "" Then Exit Sub
    Screen.MousePointer = 11

    cons = "Select * from Cliente " _
            & " Where CliCiRuc = '" & Codigo & "'" _
            & " And CliTipo = " & TipoCliente.Empresa
    If sModificar Then cons = cons & " And CliCodigo <> " & prmIDEmpresa
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then aCodCli = rsAux!CliCodigo
    rsAux.Close
    
    If aCodCli <> 0 Then
        sNuevo = False: sModificar = False
        HabilitoIngreso False
        BuscoDatosEmpresa aCodCli
    Else
        Foco tRazonSocial
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los datos."
    Botones True, False, False, False, False, Toolbar1, Me
End Sub

Public Function BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
Dim RsUsr As rdoResultset
Dim aRetorno As String: aRetorno = ""
    
    On Error Resume Next
    
    cons = "Select * from Usuario Where UsuCodigo = " & Codigo
    Set RsUsr = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Digito Then aRetorno = Trim(RsUsr!UsuDigito)
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    
    BuscoUsuario = aRetorno
    
End Function

