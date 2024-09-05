VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmMaUsuario 
   Caption         =   "Administrador de Usuarios"
   ClientHeight    =   5370
   ClientLeft      =   3135
   ClientTop       =   3060
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaUsuario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   6885
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   6885
      _ExtentX        =   12144
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
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   4200
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
   Begin VB.Frame fMEnsajeria 
      Caption         =   "Opciones de Mensajería"
      ForeColor       =   &H00800000&
      Height          =   3315
      Left            =   60
      TabIndex        =   39
      Top             =   840
      Width           =   915
      Begin AACombo99.AACombo cCategoria 
         Height          =   315
         Left            =   2040
         TabIndex        =   27
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.CheckBox oMensajeria 
         Caption         =   "Habilitar Mensajería."
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox tFLectura 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   25
         Top             =   300
         Width           =   915
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsGrupo 
         Height          =   2085
         Left            =   120
         TabIndex        =   28
         Top             =   1020
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   3678
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
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
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ca&tegoría Asociada:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   660
         Width           =   1755
      End
   End
   Begin VB.Frame fLogin 
      Caption         =   "Opciones de Login"
      ForeColor       =   &H00800000&
      Height          =   3255
      Left            =   120
      TabIndex        =   32
      Top             =   1980
      Width           =   6735
      Begin VB.TextBox tTrabajoAl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   45
         Top             =   2140
         Width           =   1095
      End
      Begin VB.TextBox tTrabajoD 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   43
         Top             =   1840
         Width           =   1095
      End
      Begin VB.TextBox tEraDigito 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   40
         Top             =   2400
         Width           =   495
      End
      Begin VB.CheckBox oHabilitado 
         Caption         =   "Usuario habilitado."
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   1875
      End
      Begin VB.TextBox tCaduca 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   23
         Top             =   2880
         Width           =   495
      End
      Begin VB.CheckBox oCaduca 
         Caption         =   "Cambiar contraseña cada ... días."
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   2775
      End
      Begin VB.CheckBox oSesion 
         Caption         =   "Cambiar al inicio de sesión."
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox tDigito 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6000
         MaxLength       =   3
         TabIndex        =   13
         Top             =   540
         Width           =   495
      End
      Begin VB.TextBox tNueva 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5400
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox tVerificacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5400
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   19
         Top             =   1500
         Width           =   1095
      End
      Begin VB.TextBox tInicial 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   11
         Top             =   540
         Width           =   495
      End
      Begin VB.TextBox tAnterior 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5400
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox tLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4680
         MaxLength       =   12
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
         Height          =   2085
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   3678
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
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
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Al:"
         Height          =   255
         Left            =   3840
         TabIndex        =   44
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Trabajó Desde:"
         Height          =   255
         Left            =   3840
         TabIndex        =   42
         Top             =   1900
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Era dígito"
         Height          =   255
         Left            =   2220
         TabIndex        =   41
         Top             =   2415
         Width           =   735
      End
      Begin VB.Label lCambio 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4560
         TabIndex        =   34
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   255
         Left            =   3840
         TabIndex        =   33
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dígito:"
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nueva Contraseña:"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Verificación:"
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Iniciales:"
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lPassOld 
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña &Actual:"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Login:"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin ComctlLib.TabStrip TabOp 
      Height          =   3675
      Left            =   60
      TabIndex        =   38
      Top             =   1680
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6482
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Lo&gin"
            Key             =   "login"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Mensa&jería"
            Key             =   "mensajeria"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton bUsuario 
      Caption         =   "&Usuarios..."
      Height          =   320
      Left            =   5640
      TabIndex        =   0
      Top             =   570
      Width           =   1095
   End
   Begin VB.TextBox tApellido1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   960
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox tApellido2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox tNombre1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   960
      MaxLength       =   20
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox tNombre2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   3
      Top             =   960
      Width           =   3015
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   5115
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4022
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   960
      TabIndex        =   36
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lCreado 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S/D"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2820
      TabIndex        =   31
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Creado:"
      Height          =   255
      Left            =   2100
      TabIndex        =   30
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lApellido 
      BackStyle       =   0  'Transparent
      Caption         =   "&Apellidos:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombres:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6240
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
            Picture         =   "frmMaUsuario.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaUsuario.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaUsuario.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaUsuario.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaUsuario.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaUsuario.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaUsuario.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaUsuario.frx":0DC8
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
         Caption         =   "&Del formulario"
      End
   End
End
Attribute VB_Name = "frmMaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean

Dim UnicoSA As Boolean
Dim aAdm As Boolean
Dim aAnterior As String

Private Sub bUsuario_Click()
    ListaDeUsuarios
End Sub


Private Sub cCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then vsGrupo.SetFocus
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error Resume Next
    Botones True, False, False, False, False, Toolbar1, Me
    
    Cons = "Select CMeCodigo, CMeNombre from CategoriaMensaje Order by CMeNombre"
    CargoCombo Cons, cCategoria
    
    LimpioFicha
    DeshabilitoIngreso
    
    InicializoGrilla
    
    fLogin.Top = TabOp.ClientTop: fLogin.Left = TabOp.ClientLeft: fLogin.Width = TabOp.ClientWidth: fLogin.Height = TabOp.ClientHeight
    fMEnsajeria.Top = TabOp.ClientTop: fMEnsajeria.Left = TabOp.ClientLeft: fMEnsajeria.Width = TabOp.ClientWidth: fMEnsajeria.Height = TabOp.ClientHeight
        
    fLogin.ZOrder 0
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    CierroConexion
    Set msgError = Nothing
    Set clsGeneral = Nothing
    
    End
    
    
End Sub

Private Sub Label1_Click()
    Foco tDigito
End Sub

Private Sub Label3_Click()
    Foco tNueva
End Sub

Private Sub Label4_Click()
    Foco tVerificacion
End Sub

Private Sub Label5_Click()
    Foco tInicial
End Sub

Private Sub Label6_Click()
    Foco cCategoria
End Sub

Private Sub Label8_Click()
    Foco tLogin
End Sub

Private Sub Label9_Click()
    Foco tNombre1
End Sub

Private Sub lApellido_Click()
    Foco tApellido1
End Sub

Private Sub lPassOld_Click()
    Foco tAnterior
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

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub oCaduca_Click()

    If sNuevo Or sModificar Then
        If oCaduca.Value = vbChecked Then
            tCaduca.Enabled = True
            tCaduca.BackColor = Obligatorio
        Else
            tCaduca.Enabled = False
            tCaduca.BackColor = Inactivo
        End If
    End If
    
End Sub

Private Sub oCaduca_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If tCaduca.Enabled Then Foco tCaduca Else: AccionGrabar
    End If
    
End Sub

Private Sub oHabilitado_Click()
    
    If oHabilitado.Value Then
        tEraDigito.Enabled = False: tEraDigito.Text = "": tEraDigito.BackColor = Colores.Inactivo
        tTrabajoAl.Enabled = False: tTrabajoAl.Text = "": tTrabajoAl.BackColor = Colores.Inactivo
    Else
        tEraDigito.Enabled = True: tEraDigito.Text = "": tEraDigito.BackColor = Colores.Blanco
        tTrabajoAl.Enabled = True: tTrabajoAl.Text = "": tTrabajoAl.BackColor = Colores.Blanco
    End If
    
End Sub

Private Sub oHabilitado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tEraDigito.Enabled Then Foco tEraDigito Else oSesion.SetFocus
    End If
End Sub

Private Sub oMensajeria_Click()
    
    If sNuevo Or sModificar Then
        If oMensajeria.Value = vbChecked Then
            If sModificar Then      'Verifico si tiene fecha de lectura
                Dim rsL As rdoResultset
                Cons = "Select * from MensajeLectura Where MLeIdUsuario = " & Val(lCodigo.Caption)
                Set rsL = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsL.EOF Then
                    tFLectura.Text = Format(rsL!MLeFecha, "dd/mm/yyyy")
                    'MsgBox "El usuario ya ha utilizado el correo, tiene asignada fecha para lectura de mesajes.", vbExclamation, "ATENCIÓN"
                Else
                    tFLectura.Enabled = True: tFLectura.BackColor = Obligatorio
                    tFLectura.Text = Format(Date, "dd/mm/yyyy")
                End If
                rsL.Close
            Else
                tFLectura.Enabled = True: tFLectura.BackColor = Obligatorio
                tFLectura.Text = Format(Date, "dd/mm/yyyy")
            End If
        Else
            tFLectura.Enabled = False: tFLectura.BackColor = Inactivo
            tFLectura.Text = ""
        End If
    End If
    
End Sub

Private Sub oMensajeria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If tFLectura.Enabled Then Foco tFLectura Else Foco cCategoria
End Sub

Private Sub oSesion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then oCaduca.SetFocus
End Sub

Private Sub TabOp_Click()

    Select Case LCase(TabOp.SelectedItem.Key)
        Case "login": fLogin.ZOrder 0
        Case "mensajeria": fMEnsajeria.ZOrder 0
    End Select
    Me.Refresh

End Sub

Private Sub tAnterior_GotFocus()
    tAnterior.SelStart = 0: tAnterior.SelLength = Len(tAnterior.Text)
End Sub

Private Sub tAnterior_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tNueva.SetFocus
End Sub

Private Sub tApellido1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tApellido2
End Sub

Private Sub tApellido2_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then vsLista.SetFocus
End Sub

Private Sub tCaduca_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tDigito_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If tAnterior.Enabled Then Foco tAnterior Else: Foco tNueva
    End If
    
End Sub

Private Sub tEraDigito_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then oSesion.SetFocus
End Sub

Private Sub tFLectura_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tFLectura.Text) Then tFLectura.Text = Format(tFLectura.Text, "dd/mm/yyyy")
        Foco cCategoria
    End If
End Sub

Private Sub tInicial_GotFocus()
    tInicial.SelStart = 0: tInicial.SelLength = Len(tInicial.Text)
End Sub

Private Sub tInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDigito
End Sub

Private Sub tLogin_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tInicial
End Sub

Private Sub tNombre1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tNombre2
End Sub

Private Sub tNombre2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tApellido1
End Sub

Private Sub tNueva_GotFocus()
    tNueva.SelStart = 0: tNueva.SelLength = Len(tNueva.Text)
End Sub

Private Sub tNueva_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tVerificacion.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "grabar": AccionGrabar
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select

End Sub

Private Sub AccionEliminar()

    If lCodigo.Caption = "" Then Exit Sub
    
    If MsgBox("ATENCIÓN: Esta acción puede causar inconsistencias en las tablas." & Chr(vbKeyReturn) & _
        "Confirma eliminar el usuario seleccionado.", vbQuestion + vbYesNo + vbDefaultButton2, "ELIMINAR") = vbNo Then Exit Sub
        
    On Error GoTo ErrBE
    Cons = "SELECT * FROM Usuario Where UsuCodigo = " & lCodigo.Caption
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        MsgBox "No se encontró el usuario seleccionado, verifique que no haya sido eliminado.", vbCritical, "ERROR"
        RsAux.Close
        Exit Sub
    End If
    RsAux.Delete: RsAux.Close
    
    Cons = "Select * from MensajeLectura Where MLeIdUsuario = " & lCodigo.Caption
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then RsAux.Delete
    RsAux.Close
    
    LimpioFicha
    Botones True, False, False, False, False, Toolbar1, Me
    Exit Sub
    
ErrBE:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al eliminar el usuario.", Err.Description
End Sub

Private Sub AccionNuevo()
    
    sNuevo = True
    
    LimpioFicha
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    
End Sub

Private Sub AccionModificar()
    
    HabilitoIngreso False
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    
    sModificar = True
    
End Sub

Private Sub AccionGrabar()

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma grabar los datos del usuario.", vbYesNo + vbQuestion, "GRABAR") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    On Error GoTo errGrabar
    If sNuevo Then
        Cons = "Select * from Usuario Where UsuCodigo = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.AddNew
        CargoDatosBD
        RsAux.Update: RsAux.Close
        
        Dim aUsuario As Long
        Cons = "Select Max(UsuCodigo) from Usuario"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        aUsuario = RsAux(0): RsAux.Close
        CargoDatosBDNiveles aUsuario
        
        GraboDatosMensajeria aUsuario
        
        
    Else
        Cons = "Select * from Usuario Where UsuCodigo = " & lCodigo.Caption
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        RsAux.Edit
        CargoDatosBD
        RsAux.Update: RsAux.Close
        
        CargoDatosBDNiveles CLng(lCodigo.Caption)
        
        GraboDatosMensajeria CLng(lCodigo.Caption)
        
    End If
    If Not sModificar Then
        LimpioFicha
        Botones True, False, False, False, False, Toolbar1, Me
    Else
        Botones True, True, True, False, False, Toolbar1, Me
    End If
    sNuevo = False: sModificar = False
    
    DeshabilitoIngreso
    
    Screen.MousePointer = 0
    Exit Sub

errGrabar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al grabar los datos."
    On Error Resume Next
    RsAux.Close
End Sub

Private Sub GraboDatosMensajeria(idUsr As Long)
    
    If oMensajeria.Value = vbChecked And tFLectura.Enabled Then
        Cons = "Select * from MensajeLectura Where MLeIdUsuario = " & idUsr
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then RsAux.AddNew Else RsAux.Edit
        RsAux!MLeIDUsuario = idUsr
        RsAux!MLeFecha = Format(tFLectura.Text, "mm/dd/yyyy")
        RsAux.Update: RsAux.Close
    End If
    
    'Grabo los grupos asignados al usuario
    If sModificar Then
        Cons = "Delete UsuarioGrupo Where UGrIdUsuario = " & idUsr
        cBase.Execute Cons
    End If
    
    With vsGrupo
    For I = 1 To .Rows - 1
        If .Cell(flexcpChecked, I, 0) = flexChecked Then
            Cons = "Insert Into UsuarioGrupo (UGrIdUsuario, UGrIdGrupo) Values (" & idUsr & ", " & .Cell(flexcpData, I, 0) & ")"
            cBase.Execute Cons
        End If
    Next I
    End With
        
End Sub

Private Sub AccionCancelar()
    
    If sModificar Then
        CargoUsuario CLng(lCodigo.Caption)
        Botones True, True, True, False, False, Toolbar1, Me
    
    Else
        LimpioFicha
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    
    sNuevo = False: sModificar = False
    DeshabilitoIngreso
    
End Sub

Private Sub DeshabilitoIngreso()
    
    bUsuario.Enabled = True
    
    tNombre1.Enabled = False: tNombre1.BackColor = Inactivo
    tNombre2.Enabled = False: tNombre2.BackColor = Inactivo
    tApellido1.Enabled = False: tApellido1.BackColor = Inactivo
    tApellido2.Enabled = False: tApellido2.BackColor = Inactivo
    
    vsLista.Editable = False: vsLista.BackColor = Inactivo
    
    tLogin.Enabled = False: tLogin.BackColor = Inactivo
    tInicial.Enabled = False: tInicial.BackColor = Inactivo
    tDigito.Enabled = False: tDigito.BackColor = Inactivo
    tAnterior.Enabled = False: tAnterior.BackColor = Inactivo
    tNueva.Enabled = False: tNueva.BackColor = Inactivo
    tVerificacion.Enabled = False: tVerificacion.BackColor = Inactivo
    
    tEraDigito.Enabled = False: tEraDigito.BackColor = Inactivo
    tTrabajoD.Enabled = False: tTrabajoD.BackColor = Inactivo
    tTrabajoAl.Enabled = False: tTrabajoAl.BackColor = Inactivo
        
    oHabilitado.Enabled = False
    oSesion.Enabled = False
    oCaduca.Enabled = False
    tCaduca.Enabled = False: tCaduca.BackColor = Inactivo
    
    oMensajeria.Enabled = False
    tFLectura.Enabled = False: tFLectura.BackColor = Inactivo
    cCategoria.Enabled = False: cCategoria.BackColor = Colores.Inactivo
    vsGrupo.Editable = False: vsGrupo.BackColor = Inactivo
    
End Sub

Private Sub HabilitoIngreso(Optional ParaNuevo As Boolean = True)

    bUsuario.Enabled = False
    
    tNombre1.Enabled = True: tNombre1.BackColor = Obligatorio
    tNombre2.Enabled = True: tNombre2.BackColor = Blanco
    tApellido1.Enabled = True: tApellido1.BackColor = Obligatorio
    tApellido2.Enabled = True: tApellido2.BackColor = Blanco
    
    vsLista.Editable = True: vsLista.BackColor = Blanco
    tLogin.Enabled = True: tLogin.BackColor = Obligatorio
    tInicial.Enabled = True: tInicial.BackColor = Obligatorio
    tDigito.Enabled = True: tDigito.BackColor = Obligatorio
    
    tNueva.Enabled = True: tNueva.BackColor = Blanco
    tVerificacion.Enabled = True: tVerificacion.BackColor = Blanco
    
    tTrabajoD.Enabled = True: tTrabajoD.BackColor = Blanco
    
    oHabilitado.Enabled = True
    If oHabilitado.Value Then
        tTrabajoAl.Enabled = False: tTrabajoAl.BackColor = Inactivo
        tEraDigito.Enabled = False: tEraDigito.BackColor = Inactivo
    Else
        tTrabajoAl.Enabled = True: tTrabajoAl.BackColor = Blanco
        tEraDigito.Enabled = True: tEraDigito.BackColor = Blanco
    End If
    
    oSesion.Enabled = True
    oCaduca.Enabled = True
    If oCaduca.Value = vbChecked Then
        tCaduca.BackColor = Obligatorio
        tCaduca.Enabled = True
    End If
    
    oMensajeria.Enabled = True
    cCategoria.Enabled = True: cCategoria.BackColor = Colores.Blanco
    vsGrupo.Editable = True: vsGrupo.BackColor = Blanco
    
    If Not ParaNuevo Then
        tAnterior.Enabled = True
        tAnterior.BackColor = Blanco
    End If
    
End Sub

Private Sub LimpioFicha()

    lCodigo.Caption = ""
    tNombre1.Text = ""
    tNombre2.Text = ""
    tApellido1.Text = ""
    tApellido2.Text = ""
    lCreado.Caption = "S/D"
    
    With vsLista
        For I = 1 To .Rows - 1: .Cell(flexcpChecked, I, 0) = flexUnchecked: Next I
    End With
    
    tLogin.Text = ""
    tInicial.Text = ""
    tDigito.Text = ""
    tNueva.Text = ""
    tVerificacion.Text = ""
    tAnterior.Text = ""
    tEraDigito.Text = ""
    tTrabajoD.Text = "": tTrabajoAl.Text = ""
    
    oHabilitado.Value = vbUnchecked
    oSesion.Value = vbUnchecked
    oCaduca.Value = vbUnchecked
    tCaduca.Text = ""
    lCambio.Caption = "S/D"
    
    oMensajeria.Value = vbUnchecked: tFLectura.Text = ""
    cCategoria.Text = ""
    With vsGrupo
        For I = 1 To .Rows - 1: .Cell(flexcpChecked, I, 0) = flexUnchecked: Next I
    End With
    
End Sub

Private Sub CargoUsuario(Codigo As Long)
    
    LimpioFicha
    Cons = "SELECT * FROM Usuario Where UsuCodigo = " & Codigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        lCodigo.Caption = RsAux!UsuCodigo
        tNombre1.Text = Trim(RsAux!UsuNombre1)
        If Not IsNull(RsAux!UsuNombre2) Then tNombre2.Text = Trim(RsAux!UsuNombre2)
        tApellido1.Text = Trim(RsAux!UsuApellido1)
        If Not IsNull(RsAux!UsuApellido2) Then tApellido2.Text = Trim(RsAux!UsuApellido2)
        
        tLogin.Text = Trim(RsAux!UsuIdentificacion)
        tInicial.Text = Trim(RsAux!UsuInicial)
        tDigito.Text = Trim(RsAux!UsuDigito)
                
        If Not IsNull(RsAux!UsuCaduca) Then
            oCaduca.Value = vbChecked
            tCaduca.Text = Trim(RsAux!UsuCaduca)
        End If
        
        If RsAux!UsuCambio Then oSesion.Value = vbChecked
        If RsAux!UsuHabilitado Then oHabilitado.Value = vbChecked
        If RsAux!UsuMensajeria Then oMensajeria.Value = vbChecked
        If Not IsNull(RsAux!UsuCategoriaM) Then BuscoCodigoEnCombo cCategoria, RsAux!UsuCategoriaM
        
        If Not IsNull(RsAux!UsuFContraseña) Then lCambio.Caption = Format(RsAux!UsuFContraseña, "Ddd d-Mmm yyyy hh:mm")
        lCreado.Caption = Format(RsAux!UsuFCreacion, "Ddd d-Mmm yyyy hh:mm")
        
        'Contraseñas
        tAnterior.Text = EncryptoString(Trim(RsAux!UsuContraseña))
        tAnterior.Tag = Trim(tAnterior.Text)
        tVerificacion.Text = Trim(tAnterior.Text)
        
        If Not IsNull(RsAux!UsuContraseñaAnterior) Then lPassOld.Tag = EncryptoString(Trim(RsAux!UsuContraseñaAnterior)) Else: lPassOld.Tag = ""
        
        If Not IsNull(RsAux!UsuEraDigito) Then tEraDigito.Text = Trim(RsAux!UsuEraDigito)
        If Not IsNull(RsAux!UsuDesde) Then tTrabajoD.Text = Format(RsAux!UsuDesde, "dd/mm/yyyy")
        If Not IsNull(RsAux!UsuHasta) Then tTrabajoAl.Text = Format(RsAux!UsuHasta, "dd/mm/yyyy")
    
        'Cargo los Niveles asignados------------------------------------------------------------
        Dim RsNiv As rdoResultset
        With vsLista
            Cons = "Select * from UsuarioNivel Where UNiUsuario = " & RsAux!UsuCodigo
            Set RsNiv = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            Do While Not RsNiv.EOF
                For I = 1 To .Rows - 1
                    If .Cell(flexcpData, I, 0) = RsNiv!UNiNivel Then .Cell(flexcpChecked, I, 0) = flexChecked: Exit For
                Next I
                RsNiv.MoveNext
            Loop
            RsNiv.Close
        End With
        '------------------------------------------------------------------------------------------
        
        'Cargo los Grupos de Mensajeria asignados------------------------------------------------------------
        With vsGrupo
            Cons = "Select * from UsuarioGrupo Where UGrIdUsuario = " & RsAux!UsuCodigo
            Set RsNiv = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            Do While Not RsNiv.EOF
                For I = 1 To .Rows - 1
                    If .Cell(flexcpData, I, 0) = RsNiv!UGrIdGrupo Then .Cell(flexcpChecked, I, 0) = flexChecked: Exit For
                Next I
                RsNiv.MoveNext
            Loop
            RsNiv.Close
        End With
        '------------------------------------------------------------------------------------------
        
        Botones True, True, True, False, False, Toolbar1, Me
        
    End If
    RsAux.Close

End Sub

Private Function ValidoCampos()

    ValidoCampos = False
    If Trim(tNombre1.Text) = "" Or Trim(tApellido1.Text) = "" Then
        MsgBox "Se debe ingresar el nombre y apellido del usuario.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If Trim(tLogin.Text) = "" Then
        MsgBox "Debe ingresar una descripcion de login.", vbExclamation, "ATENCIÓN"
        Foco tLogin: Exit Function
    End If
    
    If Trim(tInicial.Text) = "" Then
        MsgBox "Debe ingresar las iniciales del usuario.", vbExclamation, "ATENCIÓN"
        Foco tInicial: Exit Function
    End If
    
    If Not IsNumeric(tDigito.Text) Then
        MsgBox "El dígito ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tDigito: Exit Function
    End If
    
    If oCaduca.Value = vbChecked Then
        If Not IsNumeric(tCaduca.Text) Then
            MsgBox "Debe ingresar los días en los que va a caducar la contraseña.", vbExclamation, "ATENCIÓN"
            Foco tCaduca: Exit Function
        End If
    End If
    
    If oMensajeria.Value = vbChecked And tFLectura.Enabled Then
        If Not IsDate(tFLectura.Text) Then
            MsgBox "Debe ingresar la fecha inicial para la lectura de mensajes.", vbExclamation, "ATENCIÓN"
            Foco tFLectura: Exit Function
        End If
    End If
    
    'La nueva debe coincidir con la verificaicon---------------------------------------------------
    If sNuevo Then
        If Trim(tNueva.Text) = "" Then
            MsgBox "Se debe ingresar una contraseña.", vbCritical, "ATENCIÓN"
            Exit Function
        End If
        
        If EncryptoString(EncryptoString(Trim(tNueva.Text))) <> tNueva.Text Then
            MsgBox "La contraseña ingresada no se pudo encryptar. Por favor ingrese otra.", vbExclamation, "Error de Encryptación"
            Foco tNueva: Exit Function
        End If
    
        If Trim(tVerificacion.Text) <> Trim(tNueva.Text) Then
            MsgBox "La contraseña no coincide con la verificación.", vbCritical, "ATENCIÓN"
            Exit Function
        End If
    End If
    
    'La anterior debe coincidir con la anterior ingresada--------------------------------------------
    If sModificar Then
        If Trim(tAnterior.Text) <> Trim(tAnterior.Tag) Then
            MsgBox "La contraseña anterior no coincide con la ingresada.", vbCritical, "ATENCIÓN"
            Exit Function
        End If
        
        If Trim(tNueva.Text) <> "" Then
            'Modifico la contraseña
            If Trim(tVerificacion.Text) <> Trim(tNueva.Text) Then
                MsgBox "La contraseña no coincide con la verificación.", vbCritical, "ATENCIÓN"
                Exit Function
            End If
            'No puede coincidir con la Anterior a la Acutal
            If Trim(lPassOld.Tag) <> "" And Trim(tVerificacion.Text) = Trim(lPassOld.Tag) Then
                MsgBox "Ud. ha utilizado esta contraseña anteriormente. Por su seguridad cámbiela.", vbCritical, "ATENCIÓN"
                Exit Function
            End If
            
            tNueva.Text = UCase(tNueva.Text)
            If EncryptoString(EncryptoString(Trim(tNueva.Text))) <> tNueva.Text Then
                MsgBox "La contraseña ingresada no se pudo encryptar. Por favor ingrese una nueva.", vbExclamation, "Error de Encryptación"
                Foco tNueva: Exit Function
            End If
            
        Else
            If Trim(tVerificacion.Text) <> Trim(tAnterior.Text) Then
                MsgBox "La contraseña no coincide con la verificación.", vbCritical, "ATENCIÓN"
                Exit Function
            End If
        End If
    End If
    
    'Valido si no hay otro usuario habilitado con ese dígito-----------------------------------------------
    If oHabilitado.Value = vbChecked Then
        Cons = "Select * from Usuario " _
                & " Where UsuDigito = " & tDigito.Text _
                & " And UsuHabilitado = 1 "
        If sModificar Then Cons = Cons & " And UsuCodigo <> " & lCodigo.Caption
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not RsAux.EOF Then
            MsgBox "El dígito ingresado está siendo usado por: " & Trim(RsAux!UsuApellido1) & ", " & Trim(RsAux!UsuNombre1), vbExclamation, "ATENCIÓN"
            RsAux.Close
            Exit Function
        End If
        RsAux.Close
        
        'Valido si no hay otro usuario habilitado con ese login-----------------------------------------------
        Cons = "Select * from Usuario " _
                & " Where UsuIdentificacion = '" & Trim(tLogin.Text) & "'" _
                & " And UsuHabilitado = 1 "
        If sModificar Then Cons = Cons & " And UsuCodigo <> " & lCodigo.Caption
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not RsAux.EOF Then
            MsgBox "El login ingresado está siendo usado por: " & Trim(RsAux!UsuApellido1) & ", " & Trim(RsAux!UsuNombre1), vbExclamation, "ATENCIÓN"
            RsAux.Close
            Exit Function
        End If
        RsAux.Close
    End If
    
    If Trim(tTrabajoD.Text) <> "" Then
        If Not IsDate(tTrabajoD.Text) Then
            MsgBox "La fecha 'Trabajó Desde' no es correcta. Verifique.", vbExclamation, "Dato Incorrecto"
            Foco tTrabajoD: Exit Function
        Else
            tTrabajoD.Text = Format(tTrabajoD, "dd/mm/yyyy")
        End If
    End If
    If Trim(tTrabajoAl.Text) <> "" Then
        If Not IsDate(tTrabajoAl.Text) Then
            MsgBox "La fecha 'Trabajó Al' no es correcta. Verifique.", vbExclamation, "Dato Incorrecto"
            Foco tTrabajoAl: Exit Function
        Else
            tTrabajoAl.Text = Format(tTrabajoAl, "dd/mm/yyyy")
        End If
    End If
    If Trim(tEraDigito.Text) <> "" Then
        If Not IsNumeric(tEraDigito.Text) Then
            MsgBox "El valor del campo Era Dígito no es correcto. Verifique.", vbExclamation, "Dato Incorrecto"
            Foco tEraDigito: Exit Function
        End If
    End If
    
    ValidoCampos = True
    
End Function


Private Sub tTrabajoAl_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsDate(tTrabajoAl.Text) Then tTrabajoAl.Text = Format(tTrabajoAl.Text, "dd/mm/yyyy")
        oHabilitado.SetFocus
    End If
    
End Sub

Private Sub tTrabajoD_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tTrabajoD.Text) Then tTrabajoD.Text = Format(tTrabajoD.Text, "dd/mm/yyyy")
        If tTrabajoAl.Enabled Then Foco tTrabajoAl Else oHabilitado.SetFocus
    End If
End Sub

Private Sub tVerificacion_GotFocus()
    tVerificacion.SelStart = 0: tVerificacion.SelLength = Len(tVerificacion.Text)
End Sub

Private Sub tVerificacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tTrabajoD.Enabled Then Foco tTrabajoD Else oHabilitado.SetFocus
    End If
End Sub

Private Sub ListaDeUsuarios()

Dim aSeleccionado As Long

    On Error GoTo errBuscar
    Screen.MousePointer = 11
    Dim aLista As New clsListadeAyuda
    
    Cons = "Select UsuCodigo, Nombre = (RTrim(UsuApellido1) + RTrim(' ' + UsuApellido2)+', ' + RTrim(UsuNombre1)) + RTrim(' ' + UsuNombre2)" _
            & " From Usuario Order By Nombre"
                              
    aLista.ActivoListaAyuda Cons, False, cBase.Connect, 5000
    
    aSeleccionado = aLista.ValorSeleccionado
    If aSeleccionado <> 0 Then
        CargoUsuario aSeleccionado
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    
    Set aLista = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos.", Err.Description
    Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub CargoDatosBD()

    RsAux!UsuNombre1 = Trim(tNombre1.Text)
    RsAux!UsuApellido1 = Trim(tApellido1.Text)
    If Trim(tNombre2.Text) <> "" Then RsAux!UsuNombre2 = Trim(tNombre2.Text) Else: RsAux!UsuNombre2 = Null
    If Trim(tApellido2.Text) <> "" Then RsAux!UsuApellido2 = Trim(tApellido2.Text) Else: RsAux!UsuApellido2 = Null
    
    RsAux!UsuIdentificacion = Trim(tLogin.Text)
    RsAux!UsuInicial = Trim(tInicial.Text)
    RsAux!UsuDigito = Trim(tDigito.Text)
        
    If oHabilitado.Value = vbChecked Then RsAux!UsuHabilitado = 1 Else: RsAux!UsuHabilitado = 0
    If oCaduca.Value Then RsAux!UsuCaduca = tCaduca.Text Else RsAux!UsuCaduca = Null
    If oSesion.Value Then RsAux!UsuCambio = 1 Else: RsAux!UsuCambio = 0
    If oMensajeria.Value Then RsAux!UsuMensajeria = 1 Else: RsAux!UsuMensajeria = 0
    If cCategoria.ListIndex <> -1 Then RsAux!UsuCategoriaM = cCategoria.ItemData(cCategoria.ListIndex) Else RsAux!UsuCategoriaM = Null
    
    If sNuevo Then
        RsAux!UsuFCreacion = Format(Now, sqlFormatoFH)
        RsAux!UsuFContraseña = Format(Now, sqlFormatoFH)
        RsAux!UsuContraseña = EncryptoString(Trim(tNueva.Text))
    End If
    
    'Contraseñas----------------------------------------------------------------------------
    If Trim(tAnterior.Text) <> "" And Trim(tAnterior.Text) <> Trim(tVerificacion.Text) Then          'Cambio la Contraseña
        RsAux!UsuContraseña = EncryptoString(Trim(tNueva.Text))
        RsAux!UsuContraseñaAnterior = EncryptoString(Trim(tAnterior.Text))
        RsAux!UsuFContraseña = Format(Now, sqlFormatoFH)
    End If
    
    If Trim(tEraDigito.Text) <> "" Then RsAux!UsuEraDigito = Val(tEraDigito.Text) Else RsAux!UsuEraDigito = Null
    If Trim(tTrabajoD.Text) <> "" Then RsAux!UsuDesde = Format(tTrabajoD.Text, "mm/dd/yyyy") Else RsAux!UsuDesde = Null
    If Trim(tTrabajoAl.Text) <> "" Then RsAux!UsuHasta = Format(tTrabajoAl.Text, "mm/dd/yyyy") Else RsAux!UsuHasta = Null
    
End Sub

Private Sub CargoDatosBDNiveles(Usuario As Long)

    'Grabo los niveles de acceso--------------------------------------------------------------
    If sModificar Then
        Cons = "Delete UsuarioNivel Where UNiUsuario = " & Usuario
        cBase.Execute Cons
    End If
    
    With vsLista
    For I = 1 To .Rows - 1
        If .Cell(flexcpChecked, I, 0) = flexChecked Then
            Cons = "Insert Into UsuarioNivel (UNiUsuario, UNiNivel) Values (" & Usuario & ", " & .Cell(flexcpData, I, 0) & ")"
            cBase.Execute Cons
        End If
    Next I
    End With

End Sub

Private Sub InicializoGrilla()
Dim aValor As Long

    On Error Resume Next
    With vsLista
        .Cols = 1: .Rows = 1:
        .FormatString = "|Niveles de Acceso"
            
        .WordWrap = True
        .ColDataType(0) = flexDTBoolean
        .ColWidth(0) = 315
        Cons = "Select NSiNivel, NSiNombre from NivelSistema order by NSiNombre"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            .AddItem ""
            aValor = RsAux!NSiNivel: .Cell(flexcpData, .Rows - 1, 0) = aValor
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!NSiNombre)
            RsAux.MoveNext
        Loop
        RsAux.Close
    End With
      
    With vsGrupo
        .Cols = 1: .Rows = 1:
        .FormatString = "|Grupos de Mensajería"
            
        .WordWrap = True
        .ColDataType(0) = flexDTBoolean
        .ColWidth(0) = 315
        Cons = "Select GMeCodigo, GMeNombre from GrupoMensaje order by GMeNombre"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            .AddItem ""
            aValor = RsAux!GMeCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!GMeNombre)
            RsAux.MoveNext
        Loop
        RsAux.Close
    End With
      
End Sub

Private Sub vsGrupo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsLista_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub
