VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTaller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Taller"
   ClientHeight    =   6165
   ClientLeft      =   3180
   ClientTop       =   2655
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTaller.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7260
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   7260
      _ExtentX        =   12806
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
            Key             =   "salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
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
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Taller"
      ForeColor       =   &H00800000&
      Height          =   1395
      Left            =   120
      TabIndex        =   26
      Top             =   480
      Width           =   7035
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   5775
      End
      Begin VB.TextBox tDigito 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   11
         TabIndex        =   1
         Top             =   540
         Width           =   495
      End
      Begin VB.Label lServicio 
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
         Left            =   1140
         TabIndex        =   39
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Servicio:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lEstado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         TabIndex        =   37
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "&Comentarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   860
         Width           =   1095
      End
      Begin VB.Label lTecnico 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1680
         TabIndex        =   29
         Top             =   540
         Width           =   1665
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Técnico:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Reparado:"
         Height          =   255
         Left            =   4440
         TabIndex        =   28
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lReparado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/09/00 23:55"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5280
         TabIndex        =   27
         Top             =   540
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Evolución del Servicio"
      ForeColor       =   &H00800000&
      Height          =   3900
      Left            =   120
      TabIndex        =   25
      Top             =   1980
      Width           =   7035
      Begin VB.CommandButton butDisposicion 
         Caption         =   "Disponer"
         Height          =   255
         Left            =   6000
         TabIndex        =   48
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton butMemo 
         Caption         =   "Comentario contacto"
         Height          =   255
         Left            =   3480
         TabIndex        =   47
         Top             =   2430
         Width           =   2055
      End
      Begin VB.TextBox txtUbicacion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   43
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox chAceptado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "&SI"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1260
         TabIndex        =   14
         Top             =   2085
         Width           =   495
      End
      Begin VB.TextBox tAceptado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1860
         TabIndex        =   15
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox tUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   22
         Top             =   3540
         Width           =   975
      End
      Begin VB.TextBox tCosto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   12
         Top             =   1695
         Width           =   1095
      End
      Begin AACombo99.AACombo cCamionI 
         Height          =   315
         Left            =   2640
         TabIndex        =   7
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
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
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   1695
         Width           =   675
         _ExtentX        =   1191
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
      Begin AACombo99.AACombo cCamionE 
         Height          =   315
         Left            =   2640
         TabIndex        =   20
         Top             =   3135
         Width           =   1935
         _ExtentX        =   3413
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
      Begin AACombo99.AACombo cLocalR 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   1035
         Width           =   2295
         _ExtentX        =   4048
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
      Begin AACombo99.AACombo cLocalE 
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   2775
         Width           =   1815
         _ExtentX        =   3201
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
      Begin AACombo99.AACombo cLocalI 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsMotivo 
         Height          =   1335
         Left            =   3900
         TabIndex        =   42
         Top             =   1020
         Width           =   3015
         _ExtentX        =   5318
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
         BackColor       =   14737632
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14737632
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
      Begin VB.Label lblContacto 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contacto telefónico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   3480
         TabIndex        =   46
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label lblDispocicion 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "Disposición WEB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4680
         TabIndex        =   45
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ubicación:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   2430
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Modificado:"
         Height          =   195
         Left            =   2580
         TabIndex        =   41
         Top             =   3570
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   3570
         Width           =   735
      End
      Begin VB.Label lTModificado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/09/00 23:55"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   40
         Top             =   3540
         Width           =   1365
      End
      Begin VB.Label lPresupuesto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         TabIndex        =   36
         Top             =   1395
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Presupuestado:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   1420
         Width           =   1275
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Local de &Entrega:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2835
         Width           =   1335
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "&Aceptado:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2085
         Width           =   855
      End
      Begin VB.Label lRecepcionE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5880
         TabIndex        =   34
         Top             =   3510
         Width           =   1005
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Recepción:"
         Height          =   195
         Left            =   5040
         TabIndex        =   33
         Top             =   3510
         Width           =   855
      End
      Begin VB.Label lTrasladoE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30/09 14:22"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   3150
         Width           =   1005
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Traslado:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   3150
         Width           =   1395
      End
      Begin VB.Label lRecepcionI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5880
         TabIndex        =   32
         Top             =   615
         Width           =   1005
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Recepción:"
         Height          =   195
         Left            =   5040
         TabIndex        =   31
         Top             =   615
         Width           =   855
      End
      Begin VB.Label lTrasladoI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30/09 14:22"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Top             =   615
         Width           =   1005
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Traslado:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   615
         Width           =   1395
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Local &Reparación:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1095
         Width           =   1515
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Costo Reparación:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1755
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Local de Ingreso:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1395
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   5910
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12753
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6060
      Top             =   -60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTaller.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTaller.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTaller.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTaller.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTaller.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTaller.frx":0BA4
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
      Begin VB.Menu MnuLinea 
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
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmTaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim sNuevo As Boolean, sModificar As Boolean
Dim gServicio As Long
Dim gAceptaPto As Boolean

Public Property Get prmServicio() As Long
    prmServicio = gServicio
End Property
Public Property Let prmServicio(Codigo As Long)
    gServicio = Codigo
End Property

Public Property Get prmAceptaPto() As Boolean
    prmAceptaPto = gAceptaPto
End Property
Public Property Let prmAceptaPto(Codigo As Boolean)
    gAceptaPto = Codigo
End Property

Private Sub butDisposicion_Click()
    Dim frmD As New frmDisposicion
    frmD.gServicio = gServicio
    frmD.Show (vbModal)
    
    If (frmD.boolGrabo) Then
        MsgBox "Vera los cambios cuando cargue la ficha de taller.", vbInformation, "ATENCIÓN"
    End If
End Sub

Private Sub butMemo_Click()
Dim smsg As String
    smsg = InputBox("Ingrese el comentario de contacto", "Comentario")
    If smsg <> "" Then
        Cons = "UPDATE Taller SET TalContactoTel = '" & smsg & "' WHERE TalServicio = " & gServicio
        cBase.Execute Cons
        lblContacto.Caption = smsg
        lblContacto.Visible = True
    End If
End Sub

Private Sub chAceptado_Click()
    If Trim(tAceptado.Text) = "" And tAceptado.Enabled Then tAceptado.Text = Now
End Sub

Private Sub chAceptado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tAceptado
End Sub

Private Sub cLocalE_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub cLocalR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

    On Error Resume Next
    ObtengoSeteoForm Me
    
    If Me.Height <> 6885 Then Me.Height = 6885
    InicializoGrilla
    sNuevo = False: sModificar = False
    CargoCombos
    LimpioFicha
    
    DeshabilitoIngreso
    If gServicio <> 0 Then
        CargoDatosTaller
        If gAceptaPto Then
            If Not IsNumeric(tCosto.Text) Then Exit Sub
            AccionModificar
            If Trim(tAceptado.Text) = "" Then
                tAceptado.Text = Format(Now, "dd/mm/yyyy hh:mm:ss")
                chAceptado.Value = vbChecked
            End If
            Foco tUsuario
        End If
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me

End Sub

Private Sub Label11_Click()
    Foco tUsuario
End Sub

Private Sub Label20_Click()
    Foco cLocalE
End Sub

Private Sub Label33_Click()
    Foco tComentario
End Sub

Private Sub Label34_Click()
    If cLocalR.Enabled Then Foco cLocalR
End Sub

Private Sub Label45_Click()
    Foco tAceptado
End Sub

Private Sub lblContacto_Click()
    MsgBox lblContacto.Caption, vbInformation, "Contacto telefónico"
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

Private Sub AccionNuevo()
   
    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    
    HabilitoIngreso True
    lRecepcionI.Caption = Format(gFechaServidor, "dd/mm hh:mm")
    
    Foco tComentario
  
End Sub

Private Sub AccionModificar()
    
    On Error Resume Next
    sModificar = True
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    
    Cons = "Select TalFIngresoRecepcion, TalFSalidaRealizado, TalIngresoCamion,SerEstadoServicio from taller inner join servicio on sercodigo = talservicio where talservicio = " & gServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        cLocalR.Enabled = True: cLocalR.BackColor = vbWindowBackground
    Else
        cLocalR.Enabled = (IsNull(RsAux("TalFIngresoRecepcion")) And IsNull(RsAux("TalFSalidaRealizado")) And _
                IsNull(RsAux("TalIngresoCamion")) And RsAux("SerEstadoservicio") = 3)
        If cLocalR.Enabled = True Then cLocalR.BackColor = vbWindowBackground
    End If
    RsAux.Close
'    cLocalI.Enabled = True
'    cLocalI.BackColor = vbWindowBackground
    tUsuario.Text = ""
    Foco tComentario
        
End Sub

Private Sub AccionGrabar()

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo errGrabar
    
    If sNuevo Then
        Cons = "Select * from Taller Where TalServicio = " & gServicio
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.AddNew
        CargoCamposBD Nuevo:=True
        RsAux.Update: RsAux.Close
        
    Else                                    'Modificar----
    
        If Not cLocalR.Enabled Then
            
            On Error GoTo errGrabar
            Cons = "Select * from Taller Where TalServicio = " & gServicio
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                If (Not IsNull(RsAux("TalFAceptacion")) And RsAux("TalAceptado") And Not IsNull(RsAux("TalFReparado"))) Then
                
                    If tAceptado.Enabled And Trim(tAceptado.Text) <> "" And Not chAceptado.Value = vbChecked Then
                                        
                        Dim gSucesoUsr As Long, gSucesoDef As String
                        Dim objSuceso As New clsSuceso
                        objSuceso.ActivoFormulario paCodigoDeUsuario, "Anulación de Documentos en Servicio", cBase
                        Me.Refresh
                        gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
                        gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
                        Set objSuceso = Nothing
                        If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub 'Abortó el ingreso del suceso
                        
                    End If
                End If
            
                RsAux.Edit
                CargoCamposBD
                RsAux.Update
                
                If (gSucesoUsr > 0) Then
                    clsGeneral.RegistroSuceso cBase, gFechaServidor, 32, paCodigoDeTerminal, gSucesoUsr, 0, _
                        Descripcion:="Servicio: " & gServicio & " de aceptado a NO", Defensa:=Trim(gSucesoDef)
                End If
            End If
            RsAux.Close
            
        Else
        
            Screen.MousePointer = 0
            If Val(cLocalR.Tag) > 0 And Val(cLocalR.Tag) <> cLocalR.ItemData(cLocalR.ListIndex) Then
                If MsgBox("Ha cambiado el local de reparación, se eliminaran los datos de recepción y quedará en condiciones de trasladar." & vbCrLf & vbCrLf & "¿Confirma continuar?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
            End If
            
            Screen.MousePointer = 11
            'ACA solo entra si no hay datos en la tabla taller.
            If Not AccionGraboCambioSucursal Then Exit Sub
        End If
    End If
    
    sNuevo = False: sModificar = False
    DeshabilitoIngreso
    LimpioFicha
    CargoDatosTaller
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    clsGeneral.OcurrioError "Error al realizar la operación.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function AccionGraboCambioSucursal() As Boolean
    
    Dim idLocalIngreso As Long, idCliente As Long, idProducto As Long, IDArticulo As Long
    AccionGraboCambioSucursal = False
    FechaDelServidor
    
    On Error GoTo errBT
    cBase.BeginTrans
    On Error GoTo errRB
    
    Cons = "Select * From Servicio Where SerCodigo = " & gServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        idLocalIngreso = RsAux!SerLocalIngreso
        idProducto = RsAux!SerProducto
        RsAux.Edit
        RsAux!SerLocalReparacion = cLocalR.ItemData(cLocalR.ListIndex)
        RsAux("SerLocalIngreso") = cLocalI.ItemData(cLocalI.ListIndex)
        RsAux("SerModificacion") = Format(gFechaServidor, "yyyy/MM/dd HH:nn:ss")
        RsAux.Update
        RsAux.Close
    Else
        RsAux.Close
        MsgBox "No se encontró la ficha del servicio, verifique si no lo eliminaron.", vbInformation, "ATENCIÓN"
        Exit Function
    End If
    
    'Me cambio el local para el que estoy parado entonces inserto en la tabla taller (no hay traslado).
    If paCodigoDeSucursal = cLocalR.ItemData(cLocalR.ListIndex) Then
        Cons = "Select * From Taller Where TalServicio = " & gServicio
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            RsAux.AddNew
            RsAux!TalServicio = gServicio
            RsAux!TalFIngresoRealizado = Format(gFechaServidor, sqlFormatoFH)
            RsAux!TalFIngresoRecepcion = Format(gFechaServidor, sqlFormatoFH)
        Else
            RsAux.Edit
        End If
        RsAux!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux!TalUsuario = tUsuario.Tag
        If Not IsNull(RsAux!TalComentario) Then tComentario.Tag = f_GetEventos(RsAux!TalComentario)
        If Trim(tComentario.Text) <> "" Or Trim(tComentario.Tag) <> "" Then
            RsAux!TalComentario = tComentario.Tag & Trim(tComentario.Text)
        Else
            RsAux!TalComentario = Null
        End If
        If tAceptado.Enabled Then
            If Trim(tAceptado.Text) <> "" Then RsAux!TalFAceptacion = Format(tAceptado.Text, sqlFormatoFH) Else RsAux!TalFAceptacion = Null
            If chAceptado.Value = vbChecked Then RsAux!TalAceptado = 1 Else RsAux!TalAceptado = 0
        End If
        If cLocalE.ListIndex <> -1 Then RsAux!TalLocalAlCliente = cLocalE.ItemData(cLocalE.ListIndex) Else RsAux!TalLocalAlCliente = Null
        RsAux.Update
        RsAux.Close
        
    Else
    
        Cons = "Select * from Taller Where TalServicio = " & gServicio
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If Val(cLocalR.Tag) > 0 And Val(cLocalR.Tag) <> cLocalR.ItemData(cLocalR.ListIndex) Then
                RsAux.Delete
            Else
                RsAux.Edit
                CargoCamposBD
                RsAux.Update
            End If
        End If
        RsAux.Close
    End If
    
    Cons = "Select * from Producto Where ProCodigo = " & idProducto
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    idCliente = RsAux!ProCliente
    IDArticulo = RsAux!ProArticulo
    RsAux.Close
   
    
    'ESTA CONDICION LA CREE AL INICIO DE AT.Cliente luego no tiene por que estar.
'    If idLocalIngreso <> cLocalI.ItemData(cLocalI.ListIndex) And idCliente = paClienteEmpresa Then
'        HagoTraslado IDArticulo, gServicio, idLocalIngreso, cLocalI.ItemData(cLocalI.ListIndex)
'    End If
    
    cBase.CommitTrans
    AccionGraboCambioSucursal = True
    Exit Function
    
errBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al intentar iniciar la transacción.", Err.Description
    Exit Function
errRB:
    Resume errB
errB:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al intentar almacenar la información.", Err.Description
    
End Function

Private Sub AccionEliminar()

    Screen.MousePointer = 11
    On Error GoTo Error
    Screen.MousePointer = 0
    Exit Sub
    
Error:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
End Sub

Private Sub AccionCancelar()

    On Error Resume Next
    DeshabilitoIngreso
    LimpioFicha
    CargoDatosTaller
    
    sNuevo = False: sModificar = False
    
End Sub


Private Sub tAceptado_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsDate(tAceptado.Text) Then If cLocalE.Enabled Then Foco cLocalE Else Foco tUsuario
    End If
    
End Sub

Private Sub tComentario_GotFocus()
    tComentario.SelLength = 0: tComentario.SelStart = Len(tComentario.Text)
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If sNuevo Then Foco cLocalR: Exit Sub
        If tAceptado.Enabled Then Foco tAceptado: Exit Sub
        If cLocalE.Enabled Then Foco cLocalE: Exit Sub
        Foco tUsuario
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select

End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    If tAceptado.Enabled Then
        If Not IsDate(tAceptado.Text) And Trim(tAceptado.Text) <> "" Then
            MsgBox "La fecha de aceptado el presupuesto, no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tAceptado: Exit Function
        End If
    End If
    
    If Val(tUsuario.Tag) = 0 Then
        MsgBox "Para grabar debe ingresar el dígito de usuario.", vbExclamation, "ATENCIÓN"
        Foco tUsuario: Exit Function
    End If
    
    If sNuevo Then
        If cLocalR.ListIndex = -1 Then
            MsgBox "Seleccione el local en dónde se va a reparar el producto.", vbExclamation, "ATENCIÓN"
            Foco cLocalR: Exit Function
        End If
    Else
        If cLocalR.Enabled Then
            If cLocalR.ListIndex = -1 Then
                MsgBox "Es obligatorio ingresar el local de reparación.", vbExclamation, "ATENCIÓN"
                cLocalR.SetFocus: Exit Function
            End If
        End If
    End If
    If cLocalI.ListIndex = -1 Then
        MsgBox "Es necesario indicar el local donde ingresó el producto.", vbExclamation, "ATENCIÓN"
        cLocalI.SetFocus
        Exit Function
    End If
    ValidoCampos = True
    
End Function

Private Sub DeshabilitoIngreso()
        
    tDigito.Enabled = False: tDigito.BackColor = Inactivo
    tComentario.Enabled = False: tComentario.BackColor = Inactivo
    
    cLocalI.Enabled = False: cLocalI.BackColor = Colores.Inactivo
    cCamionI.Enabled = False: cCamionI.BackColor = Colores.Inactivo
    
    cLocalR.Enabled = False: cLocalR.BackColor = Colores.Inactivo
    cMoneda.Enabled = False: cMoneda.BackColor = Colores.Inactivo
    tCosto.Enabled = False: tCosto.BackColor = Colores.Inactivo
    tAceptado.Enabled = False: tAceptado.BackColor = Colores.Inactivo
    chAceptado.Enabled = False
    cLocalE.Enabled = False: cLocalE.BackColor = Colores.Inactivo
    cCamionE.Enabled = False: cCamionE.BackColor = Colores.Inactivo
        
    tUsuario.Enabled = False: tUsuario.BackColor = Colores.Inactivo
    txtUbicacion.Enabled = False: txtUbicacion.BackColor = Colores.Inactivo
    butMemo.Enabled = False
    butDisposicion.Enabled = False
    
End Sub

Private Sub HabilitoIngreso(Optional Nuevo As Boolean = False)
    
    tComentario.Enabled = True: tComentario.BackColor = Colores.Blanco
    If Not Nuevo Then
        If IsDate(lPresupuesto.Caption) Then tAceptado.Enabled = True: tAceptado.BackColor = Colores.Blanco: chAceptado.Enabled = True
        'If IsDate(tAceptado.Text) And Trim(lReparado.Caption) <> "" Then tAceptado.Enabled = False: tAceptado.BackColor = Colores.Gris: chAceptado.Enabled = False
        If Not IsNumeric(tCosto.Text) Then tAceptado.Enabled = False: tAceptado.BackColor = Colores.Gris: chAceptado.Enabled = False: chAceptado.Enabled = False
        
        'Deshabilito esta condición para validar bien los casos de movs de stock.
        
'        If lReparado.Caption = "" And lTrasladoI.Caption = "" Then
'            cLocalR.Enabled = True: cLocalR.BackColor = Colores.Blanco
'        End If
        
    Else
        cLocalR.Enabled = True: cLocalR.BackColor = Colores.Blanco
    End If
    tUsuario.Enabled = True: tUsuario.BackColor = Colores.Blanco
    butMemo.Enabled = True
    butDisposicion.Enabled = True
    
    If lTrasladoE.Caption = "" Then cLocalE.Enabled = True: cLocalE.BackColor = Colores.Blanco
    txtUbicacion.Enabled = True: txtUbicacion.BackColor = Colores.Blanco
    
End Sub

Private Sub CargoCamposBD(Optional Nuevo As Boolean = False)
    
    If Not IsNull(RsAux!TalComentario) Then tComentario.Tag = f_GetEventos(RsAux!TalComentario)
    If Trim(tComentario.Text) <> "" Or Trim(tComentario.Tag) <> "" Then
        RsAux!TalComentario = tComentario.Tag & Trim(tComentario.Text)
    Else
        RsAux!TalComentario = Null
    End If
    
    If tAceptado.Enabled Then
        If Trim(tAceptado.Text) <> "" Then RsAux!TalFAceptacion = Format(tAceptado.Text, sqlFormatoFH) Else RsAux!TalFAceptacion = Null
        If chAceptado.Value = vbChecked Then RsAux!TalAceptado = 1 Else RsAux!TalAceptado = 0
    End If
    
    RsAux!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsAux!TalUsuario = Val(tUsuario.Tag)
    If cLocalE.ListIndex <> -1 Then RsAux!TalLocalAlCliente = cLocalE.ItemData(cLocalE.ListIndex) Else RsAux!TalLocalAlCliente = Null
    
    If Val(txtUbicacion.Tag) > 0 Then RsAux("TalUbicacionSucursal") = Val(txtUbicacion.Tag) Else RsAux("TalUbicacionSucursal") = Null
    
    If Nuevo Then
        RsAux!TalServicio = gServicio
        RsAux!TalFIngresoRealizado = Format(lRecepcionI.Caption, sqlFormatoFH)
        RsAux!TalFIngresoRecepcion = Format(lRecepcionI.Caption, sqlFormatoFH)
                
'        Cons = "Update Servicio Set SerEstadoServicio = " & EstadoS.Taller & "," & _
'                    " SerLocalReparacion = " & cLocalR.ItemData(cLocalR.ListIndex) & ", " & _
'                    " SerModificacion = '" & Format(gFechaServidor, sqlFormatoFH) & "'" & _
'                    ", SerLocalIngreso = " & cLocalI.ItemData(cLocalI.ListIndex) & _
'                    " Where SerCodigo = " & gServicio
'        cBase.Execute Cons
    Else
        If Val(cLocalR.Tag) > 0 And Val(cLocalR.Tag) <> cLocalR.ItemData(cLocalR.ListIndex) Then
            RsAux!TalFIngresoRealizado = Null
            RsAux!TalFIngresoRecepcion = Null
            RsAux("TalIngresoCamion") = Null
        End If
'        Cons = "Update Servicio Set SerModificacion = '" & Format(gFechaServidor, sqlFormatoFH) & "'" & _
'                ", SerLocalIngreso = " & cLocalI.ItemData(cLocalI.ListIndex) & _
'                " Where SerCodigo = " & gServicio
'        cBase.Execute Cons
    End If
    
    Dim idLocalIngreso As Long, idProducto As Long
    Dim rsServ As rdoResultset
    Cons = "SELECT * FROM Servicio WHERE SerCodigo = " & gServicio
    Set rsServ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    idLocalIngreso = rsServ!SerLocalIngreso
    idProducto = rsServ!SerProducto
    rsServ.Edit
    rsServ("SerModificacion") = Format(gFechaServidor, sqlFormatoFH)
    rsServ("SerLocalIngreso") = cLocalI.ItemData(cLocalI.ListIndex)
    If sNuevo Then
        rsServ("SerLocalReparacion") = cLocalR.ItemData(cLocalR.ListIndex)
        rsServ("SerEstadoServicio") = EstadoS.Taller
    End If
    rsServ.Update
    rsServ.Close
    
    If idLocalIngreso <> cLocalI.ItemData(cLocalI.ListIndex) Then
        Cons = "SELECT ProCliente, ProArticulo FROM Producto WHERE ProCodigo = " & idProducto
        Set rsServ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If rsServ("ProCliente") = paClienteEmpresa Then
            HagoTraslado rsServ("ProArticulo"), gServicio, idLocalIngreso, cLocalI.ItemData(cLocalI.ListIndex)
        End If
        rsServ.Close
    End If
    
End Sub

Private Sub LimpioFicha()

    tDigito.Text = "": lTecnico.Caption = ""
    lReparado.Caption = ""
    tComentario.Text = "": tComentario.Tag = ""
    
    cLocalI.Text = "": lTrasladoI.Caption = "": cCamionI.Text = "": lRecepcionI.Caption = ""
    cLocalR.Text = "": cMoneda.Text = "": tCosto.Text = "": lPresupuesto.Caption = "": tAceptado.Text = ""
    cLocalE.Text = "": lTrasladoE.Caption = "": cCamionE.Text = "": lRecepcionE.Caption = ""
    
    tUsuario.Text = "": lTModificado.Caption = ""
    vsMotivo.Rows = 1
    
    cLocalR.Tag = ""
    
End Sub

Private Sub CargoDatosTaller()

    lblDispocicion.Visible = False
    lblContacto.Visible = False
    
    Cons = "Select * from Servicio Left Outer Join Taller On SerCodigo = TalServicio " & _
            "LEFT OUTER JOIN UbicacionSucursal ON USuID = TalUbicacionSucursal " & _
               " Where SerCodigo = " & gServicio '& _
               " And SerLocalReparacion Is Not Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        lServicio.Caption = " " & RsAux!SerCodigo
        lEstado.Caption = UCase(EstadoServicio(RsAux!SerEstadoServicio))
        lEstado.Tag = RsAux!SerEstadoServicio
        
        If Not IsNull(RsAux!TalTecnico) Then
            tDigito.Text = BuscoUsuario(Codigo:=RsAux!TalTecnico, Digito:=True)
            lTecnico.Caption = " " & BuscoUsuario(Codigo:=RsAux!TalTecnico, Identificacion:=True)
        End If
        
        If Not IsNull(RsAux!TalFReparado) Then lReparado.Caption = Format(RsAux!TalFReparado, "Ddd d/mm/yy hh:mm")
        
        If Not IsNull(RsAux!TalComentario) Then
            tComentario.Text = f_QuitarClavesDelComentario(RsAux!TalComentario)
            tComentario.Tag = f_GetEventos(RsAux!TalComentario)
        End If
        
        If Not IsNull(RsAux("USuID")) Then
            txtUbicacion.Text = RsAux("USuPasillo") & "-" & RsAux("USuGondola")
            txtUbicacion.Tag = RsAux("USuID")
        End If
        
        If Not IsNull(RsAux!SerLocalIngreso) Then BuscoCodigoEnCombo cLocalI, RsAux!SerLocalIngreso
        
        If Not IsNull(RsAux!TalDisposicionProducto) Then
            
            lblDispocicion.Visible = True
            lblDispocicion.Caption = "Disposición: " & Format(RsAux("TalDisposicionProducto"), "dd/MM/yy")
            lblDispocicion.BackColor = &H4000&
            If Not IsNull(RsAux("TalNoDispuso")) Then
                If (RsAux("TalNoDispuso")) Then
                    lblDispocicion.Caption = "Retira hasta: " & Format(RsAux("TalDisposicionProducto"), "dd/MM/yy")
                    lblDispocicion.BackColor = vbRed
                End If
            End If
            
        End If
        
        If Not IsNull(RsAux("TalContactoTel")) Then
            lblContacto.Visible = True
            lblContacto.Caption = Trim(RsAux("TalContactoTel"))
        End If
        
        If Not IsNull(RsAux!TalFIngresoRealizado) Then lTrasladoI.Caption = Format(RsAux!TalFIngresoRealizado, "dd/mm hh:mm")
        If Not IsNull(RsAux!TalFIngresoRecepcion) Then lRecepcionI.Caption = Format(RsAux!TalFIngresoRecepcion, "dd/mm hh:mm")
        If lTrasladoI.Caption = lRecepcionI.Caption Then lTrasladoI.Caption = "" 'Porque no hay traslado, entro directo x recepcion
        
        If Not IsNull(RsAux!TalIngresoCamion) Then BuscoCodigoEnCombo cCamionI, RsAux!TalIngresoCamion
        
        If Not IsNull(RsAux!SerLocalReparacion) Then
            BuscoCodigoEnCombo cLocalR, RsAux!SerLocalReparacion
            cLocalR.Tag = RsAux("SerLocalReparacion")
        Else
            cLocalR.Tag = ""
        End If
        If Not IsNull(RsAux!SerMoneda) Then BuscoCodigoEnCombo cMoneda, RsAux!SerMoneda
        If Not IsNull(RsAux!SerCostoFinal) Then tCosto.Text = Format(RsAux!SerCostoFinal, FormatoMonedaP)
        If Not IsNull(RsAux!TalFAceptacion) Then tAceptado.Text = Format(RsAux!TalFAceptacion, "dd/mm/yyyy hh:mm")
        
        chAceptado.Value = vbUnchecked
        If Not IsNull(RsAux!TalAceptado) Then If RsAux!TalAceptado Then chAceptado.Value = vbChecked
        If Not IsNull(RsAux!TalFPresupuesto) Then lPresupuesto.Caption = Format(RsAux!TalFPresupuesto, "dd/mm hh:mm")
        
        If Not IsNull(RsAux!TalLocalAlCliente) Then BuscoCodigoEnCombo cLocalE, RsAux!TalLocalAlCliente
        If Not IsNull(RsAux!TalFSalidaRealizado) Then lTrasladoE.Caption = Format(RsAux!TalFSalidaRealizado, "dd/mm hh:mm")
        If Not IsNull(RsAux!TalFSalidaRecepcion) Then lRecepcionE.Caption = Format(RsAux!TalFSalidaRecepcion, "dd/mm hh:mm")
        If Not IsNull(RsAux!TalSalidaCamion) Then BuscoCodigoEnCombo cCamionE, RsAux!TalSalidaCamion
        
        If Not IsNull(RsAux!TalModificacion) Then lTModificado.Caption = " " & Format(RsAux!TalModificacion, "dd/mm/yy hh:mm")
        If Not IsNull(RsAux!TalUsuario) Then tUsuario.Text = BuscoUsuario(RsAux!TalUsuario, Identificacion:=True)
        
    End If
    RsAux.Close
    
    'Cargo los renglones de reparación
    vsMotivo.Rows = 1
    Cons = "Select * From ServicioRenglon, Articulo" & _
                " Where SReServicio = " & gServicio & _
                " And SReTipoRenglon = " & TipoRenglonS.Cumplido & _
                " And SReMotivo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        With vsMotivo
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = RsAux!SReCantidad & " " & Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Actualizo los botones----------------------------------------------
    Select Case Val(lEstado.Tag)
        Case EstadoS.Taller: Botones False, True, False, False, False, Toolbar1, Me: Exit Sub
        Case EstadoS.Visita: Botones True, False, False, False, False, Toolbar1, Me: Exit Sub
        Case Else: Botones False, False, False, False, False, Toolbar1, Me: Exit Sub
    End Select
   

End Sub

Private Sub CargoCombos()
    
    Cons = "Select * from Moneda Where MonFactura = 1 Order by MonSigno"
    CargoCombo Cons, cMoneda
    
    Cons = "Select * from Camion order by CamNombre"
    CargoCombo Cons, cCamionI
    CargoCombo Cons, cCamionE
    
    Cons = "Select * from Sucursal order by SucAbreviacion " 'LocNombre"
    CargoCombo Cons, cLocalI
    CargoCombo Cons, cLocalR
    CargoCombo Cons, cLocalE
    
End Sub

Private Sub tUsuario_Change()
    tUsuario.Tag = 0
End Sub

Private Sub tUsuario_GotFocus()
    With tUsuario: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Val(tUsuario.Tag) <> 0 Then AccionGrabar: Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(tUsuario.Text) Then Exit Sub
        Dim aId As Long
        aId = BuscoUsuarioDigito(CLng(tUsuario.Text), Codigo:=True)
        tUsuario.Text = BuscoUsuario(aId, Identificacion:=True)
        tUsuario.Tag = aId
        If Val(tUsuario.Tag) <> 0 Then AccionGrabar
    End If
    
End Sub

Private Sub InicializoGrilla()

    With vsMotivo
        .Rows = 1: .Cols = 1
        .FormatString = "<Repuestos Utilizados"
        .ColWidth(0) = 1400
        .WordWrap = False
        .ExtendLastCol = True
    End With

End Sub

Private Sub HagoTraslado(IDArticulo As Long, IdServicio As Long, idLocalInicial As Long, ByVal idLocalFinal As Long)
    
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), 2, idLocalFinal, IDArticulo, 1, paEstadoARecuperar, 1, TipoDocumento.ServicioCambioEstado, IdServicio
    MarcoMovimientoStockFisicoEnLocal 2, idLocalFinal, IDArticulo, 1, paEstadoARecuperar, 1
    
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), 2, idLocalInicial, IDArticulo, 1, paEstadoARecuperar, -1, TipoDocumento.ServicioCambioEstado, IdServicio
    MarcoMovimientoStockFisicoEnLocal 2, idLocalInicial, IDArticulo, 1, paEstadoARecuperar, -1

End Sub

Private Sub HagoCambioDeEstado(IDArticulo As Long, IdServicio As Long, idLocalInicial As Long)
Dim Camion As Integer: Camion = 1
Dim Deposito As Integer: Deposito = 2

    'Hago los cambios para esta sucursal.
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, 1, TipoDocumento.ServicioCambioEstado, IdServicio
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, -1, TipoDocumento.ServicioCambioEstado, IdServicio
        
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoARecuperar, 1, 1
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, 1, -1
    
    MarcoMovimientoStockFisicoEnLocal Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, 1
    MarcoMovimientoStockFisicoEnLocal Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, -1
    '.......................................................................................................................................
    
    'Retorno el estado del local inicial.
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), Deposito, idLocalInicial, IDArticulo, 1, paEstadoARecuperar, -1, TipoDocumento.ServicioCambioEstado, IdServicio
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), Deposito, idLocalInicial, IDArticulo, 1, paEstadoArticuloEntrega, 1, TipoDocumento.ServicioCambioEstado, IdServicio
        
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoARecuperar, 1, -1
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, 1, 1
    
    MarcoMovimientoStockFisicoEnLocal Deposito, idLocalInicial, IDArticulo, 1, paEstadoARecuperar, -1
    MarcoMovimientoStockFisicoEnLocal Deposito, idLocalInicial, IDArticulo, 1, paEstadoArticuloEntrega, 1
    '.......................................................................................................................................
    
End Sub

Private Function CargoUbicacionEnSucursal() As Boolean
On Error GoTo errCUS
    Dim sQy As String
    Dim rsU As rdoResultset
    
    If cLocalR.ListIndex = -1 Then
        MsgBox "Debe indicar el local de reparación.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    Dim pasillo As String
    Dim gondola As Long
    pasillo = txtUbicacion.Text
    pasillo = Replace(Replace(pasillo, "-", ""), " ", "")
    If Len(pasillo) < 2 Or Not IsNumeric(Mid(pasillo, 2)) Then
        MsgBox "Ubicacion incorrecta.", vbExclamation, "Atención"
        Exit Function
    End If
    gondola = Val(Mid(pasillo, 2))
    pasillo = Mid(pasillo, 1, 1)
    sQy = "SELECT USuID, USuPasillo, USuGondola FROM UbicacionSucursal " & _
        "WHERE USuPasillo = '" & pasillo & "' AND USuGondola = " & gondola & " AND USuSucursal = " & cLocalR.ItemData(cLocalR.ListIndex)
    Set rsU = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurReadOnly)
    If Not rsU.EOF Then
        txtUbicacion.Text = Trim(rsU("UsuPasillo")) & "-" & rsU("USuGondola")
        txtUbicacion.Tag = rsU("USuID")
        CargoUbicacionEnSucursal = True
    Else
        MsgBox "No hay datos para pasillo '" & UCase(pasillo) & "' y góndola " & gondola, vbInformation, "Atención"
    End If
    rsU.Close
    
    Exit Function
errCUS:
    clsGeneral.OcurrioError "Error al cargar la ubicación.", Err.Description
End Function

Private Sub txtUbicacion_Change()
    txtUbicacion.Tag = ""
End Sub

Private Sub txtUbicacion_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Val(txtUbicacion.Tag) > 0 Or txtUbicacion.Text = "" Then
            cLocalE.SetFocus
        Else
            CargoUbicacionEnSucursal
        End If
    End If
End Sub

Private Sub txtUbicacion_Validate(Cancel As Boolean)
On Error Resume Next
    If Val(txtUbicacion.Tag) = 0 And txtUbicacion.Text <> "" Then
        Cancel = Not CargoUbicacionEnSucursal
    End If
End Sub
