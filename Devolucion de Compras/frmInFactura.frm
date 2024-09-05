VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmInFactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolución de Compra"
   ClientHeight    =   6510
   ClientLeft      =   3960
   ClientTop       =   3120
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInFactura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8625
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
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
            Key             =   "salir"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
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
            ImageIndex      =   6
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
            Style           =   4
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "paga"
            Object.ToolTipText     =   "Con qué paga"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "nota"
            Object.ToolTipText     =   "Asignar nota"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle del Documento"
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
      Height          =   2715
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Width           =   8415
      Begin VB.TextBox tFactura 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   6900
         MaxLength       =   7
         TabIndex        =   27
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox tTCDolar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6900
         MaxLength       =   6
         TabIndex        =   25
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox tNeto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5460
         MaxLength       =   13
         TabIndex        =   24
         Text            =   "1,000,000.00"
         Top             =   600
         Width           =   1035
      End
      Begin VB.TextBox tUnitario 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   23
         Text            =   "1,000,000.00"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   60
         TabIndex        =   22
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox tCantidad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6360
         MaxLength       =   5
         TabIndex        =   21
         Text            =   "00000"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton bRemito 
         Caption         =   "En Compañía"
         Height          =   315
         Left            =   7080
         TabIndex        =   20
         Top             =   1170
         Width           =   1215
      End
      Begin VB.TextBox tCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   4620
         TabIndex        =   26
         Top             =   240
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
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   4620
         TabIndex        =   28
         Top             =   585
         Width           =   795
         _ExtentX        =   1402
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
         Height          =   1035
         Left            =   120
         TabIndex        =   29
         Top             =   1560
         Width           =   8175
         _ExtentX        =   14420
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
         BackStyle       =   0  'Transparent
         Caption         =   "Imp. Total:"
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Top             =   630
         Width           =   975
      End
      Begin VB.Label lFactura 
         BackStyle       =   0  'Transparent
         Caption         =   "&Documento:"
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "P&roveedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   2100
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
      Begin VB.Label lTC 
         BackStyle       =   0  'Transparent
         Caption         =   "21/08/2000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7560
         TabIndex        =   34
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T/&C:"
         Height          =   255
         Left            =   6480
         TabIndex        =   33
         Top             =   645
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Co&sto:"
         Height          =   255
         Left            =   4200
         TabIndex        =   32
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Q:"
         Height          =   255
         Left            =   5940
         TabIndex        =   31
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Artículos a Devolver"
      ForeColor       =   &H00000080&
      Height          =   2535
      Left            =   120
      TabIndex        =   18
      Top             =   3300
      Width           =   8415
      Begin VB.TextBox tLArticulo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         MaxLength       =   60
         TabIndex        =   10
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox tLCantidad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7440
         MaxLength       =   5
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsRemito 
         Height          =   1035
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
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
         FocusRect       =   2
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsLocal 
         Height          =   795
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1402
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
         FocusRect       =   2
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
      Begin AACombo99.AACombo cLLocal 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   1320
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
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Del &Local:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1365
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Q:"
         Height          =   255
         Left            =   6960
         TabIndex        =   11
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   80
      TabIndex        =   15
      Top             =   5880
      Width           =   7335
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   6255
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "sucursal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4895
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
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
            Picture         =   "frmInFactura.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":0A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":0D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":0E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":11B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentarios:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5940
      Width           =   975
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
         Visible         =   0   'False
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
Attribute VB_Name = "frmInFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------
'   --> Los remitos que el Local <> Compañia: La cantidad en compañia es la que queda por
'         sacar del remito (la que no está asignada a las facturas).
'
'   --> Los remitos que el Local = Compañia: La cantidad en compañia es la que queda para
'         pasar a los locales.
'
'   En este formulario solo se manejan los remitos con local <> de compañia
'----------------------------------------------------------------------------------------------------------
Option Explicit

Dim I As Long, gFactura As Long
Dim Msg As String, aTexto As String

Dim aValor As Long
Dim RsAux As rdoResultset, RsCom As rdoResultset

Private Sub bRemito_Click()
Dim bOK As Boolean

    'Valido los datos para cargar el stock de los remitos (en local compañía)-----------------------------------------------
    If Val(tProveedor.Tag) = 0 Then
        MsgBox "Debe seleccionar el proveedor del documento, para buscar el stock.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Sub
    End If
    
    If vsArticulo.Rows = 1 Then
        MsgBox "Debe ingresar los artículos de la nota para buscar en el stock los posibles remitos.", vbExclamation, "ATENCIÓN"
        Foco tArticulo: Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    Cons = "Select * from RemitoCompra, RemitoCompraRenglon, Articulo" _
            & " Where RCoProveedor = " & Val(tProveedor.Tag) _
            & " And RCREnCompania > 0" _
            & " And RCoLocal = " & paLocalCompañia _
            & " And RCoCodigo = RCRRemito" & " And RCRArticulo = ArtID"
            
     Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
     With vsRemito
     .Rows = 1
     Do While Not RsAux.EOF
        'Tipo|<Remito|Fecha|Articulo|>Q|Devuelve
        
        bOK = False
        For I = 1 To vsArticulo.Rows - 1
            If vsArticulo.Cell(flexcpData, I, 0) = RsAux!ArtID Then bOK = True: Exit For
        Next
        
        If bOK Then
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = RetornoNombreDocumento(RsAux!RCoTipo)
            
            If Not IsNull(RsAux!RCoSerie) Then aTexto = Trim(RsAux!RCoSerie) & " " Else aTexto = ""
            If Not IsNull(RsAux!RCoNumero) Then aTexto = aTexto & RsAux!RCoNumero
            .Cell(flexcpText, .Rows - 1, 1) = aTexto
            aValor = RsAux!RCoCodigo: .Cell(flexcpData, .Rows - 1, 1) = aValor
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!RCoFecha, "dd/mm/yy")
            
            .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ArtNombre)
            aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 3) = aValor
            
            .Cell(flexcpText, .Rows - 1, 4) = RsAux!RCREnCompania
            .Cell(flexcpText, .Rows - 1, 5) = 0
            .Cell(flexcpFontBold, .Rows - 1, 5) = True: .Cell(flexcpForeColor, .Rows - 1, 5) = Colores.Blanco:: .Cell(flexcpBackColor, .Rows - 1, 5) = Colores.Azul
            
        End If
        RsAux.MoveNext
     Loop
     End With
     
     RsAux.Close
     Screen.MousePointer = 0
     
     If vsRemito.Rows = 1 Then
        MsgBox "No hay artículos en en local compañía que coincidan con los devueltos.", vbExclamation, "ATENCIÓN"
    Else
        vsRemito.SetFocus
    End If
    '---------------------------------------------------------------------------------------------------------------------------------
    
End Sub

Private Sub cLLocal_GotFocus()
    With cLLocal: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub cLLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cLLocal.ListIndex <> -1 Then Foco tLArticulo Else Foco tComentario
    End If
End Sub

Private Sub Label8_Click()
    Foco cLLocal
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
            If Not IsNumeric(tArticulo.Text) Then BuscoArticulo tArticulo, Nombre:=tArticulo.Text Else: BuscoArticulo tArticulo, Codigo:=tArticulo.Text
            If tArticulo.Tag <> "0" Then Foco tUnitario
        Else
            bRemito.SetFocus
        End If
    End If

End Sub

Private Sub tCantidad_GotFocus()
    tCantidad.SelStart = 0: tCantidad.SelLength = Len(tCantidad.Text)
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)

Dim aValor As Long

    If KeyAscii = vbKeyReturn Then
        
        If Trim(tArticulo.Text) = "" Or tArticulo.Tag = "0" Then
            MsgBox "Debe ingresar el código del artículo.", vbExclamation, "ATENCIÓN"
            Foco tArticulo: Exit Sub
        End If
        If Not IsNumeric(tUnitario.Text) Then
            MsgBox "El costo unitario del artículo no es correcto.", vbExclamation, "ATENCIÓN"
            Foco tUnitario: Exit Sub
        End If
        
        If Not IsNumeric(tCantidad.Text) Then
            MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tCantidad: Exit Sub
        End If
        
        With vsArticulo
            For I = 1 To .Rows - 1
                If .Cell(flexcpData, I, 0) = tArticulo.Tag Then
                    MsgBox "El artículo ingresado ya está en la lista. Verifique.", vbExclamation, "ATENCIÓN": Exit Sub
                End If
            Next
        
            .AddItem Trim(tArticulo.Text)
            aValor = CLng(tArticulo.Tag): .Cell(flexcpData, .Rows - 1, 0) = aValor
        
            .Cell(flexcpText, .Rows - 1, 1) = tUnitario.Text
            .Cell(flexcpText, .Rows - 1, 2) = tCantidad.Text
            .Cell(flexcpText, .Rows - 1, 3) = Format(.Cell(flexcpText, .Rows - 1, 1) * .Cell(flexcpText, .Rows - 1, 2), FormatoMonedaP)
            
        End With
            
        tArticulo.Text = "": tCantidad.Text = "": tUnitario.Text = ""
        Foco tArticulo
    End If
    
End Sub

Private Sub tCodigo_Change()
    gFactura = 0
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCodigo.Text) Then CargoDatosCompra Val(tCodigo.Text)
    End If
End Sub

Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then BuscarCompras
End Sub

Private Sub tLArticulo_Change()
    tLArticulo.Tag = 0
End Sub

Private Sub tLArticulo_GotFocus()
    With tLArticulo: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tLArticulo_KeyPress(KeyAscii As Integer)
Dim sOk As Boolean: sOk = False

    If KeyAscii = vbKeyReturn Then
    
        If Trim(tLArticulo.Text) <> vbNullString Then
            If Not IsNumeric(tLArticulo.Text) Then BuscoArticulo tLArticulo, Nombre:=tLArticulo.Text Else: BuscoArticulo tLArticulo, Codigo:=tLArticulo.Text
            
            'Valido si está en la lista de los a Devolver
            If Val(tLArticulo.Tag) <> 0 Then
                For I = 1 To vsArticulo.Rows - 1
                    If vsArticulo.Cell(flexcpData, I, 0) = Val(tLArticulo.Tag) Then sOk = True: Exit For
                Next I
                
                If sOk Then Foco tLCantidad: Exit Sub
                                
                MsgBox "El artículo ingresado no está en la lista de los artículos a devolver.", vbExclamation, "ATENCIÓN"
                tLArticulo.Tag = 0
            End If
            
        Else
            vsLocal.SetFocus
        End If
    End If
    
End Sub

Private Sub tLCantidad_GotFocus()
    With tLCantidad: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tLCantidad_KeyPress(KeyAscii As Integer)
Dim aValor As Long

    If KeyAscii = vbKeyReturn Then
    
        If cLLocal.ListIndex = -1 Then
            MsgBox "Debe seleccionar un local para sacar la mercadería.", vbExclamation, "ATENCIÓN"
            Foco cLLocal: Exit Sub
        End If
        
        If Trim(tLArticulo.Text) = "" Or tLArticulo.Tag = "0" Then
            MsgBox "Debe ingresar el código del artículo.", vbExclamation, "ATENCIÓN"
            Foco tLArticulo: Exit Sub
        End If
        
        If Not IsNumeric(tLCantidad.Text) Then
            MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tLCantidad: Exit Sub
        End If
        
        'Si el local es compañia veo que no tenga servicios pendientes.
        Dim idServCpia As String
        If CLng(cLLocal.ItemData(cLLocal.ListIndex)) = 14 Then
            idServCpia = HayServiciosEnCompañia(Val(tLArticulo.Tag))
        End If
'        If bServCpia Then
'            If MsgBox("Hay servicios en compañia para el artículo, desea descartar el/los servicios que están asignados para este artículo.", vbQuestion + vbYesNo, "Servicios") = vbYes Then
'                bServCpia = True
'            Else
'                bServCpia = False
'            End If
'        End If
        With vsLocal
            For I = 1 To .Rows - 1
                If .Cell(flexcpData, I, 1) = tLArticulo.Tag And .Cell(flexcpData, I, 0) = cLLocal.ItemData(cLLocal.ListIndex) Then
                    MsgBox "El artículo ingresado ya está en la lista. Verifique local y artículo.", vbExclamation, "ATENCIÓN": Exit Sub
                End If
            Next
        
            .AddItem Trim(cLLocal.Text)
            aValor = CLng(cLLocal.ItemData(cLLocal.ListIndex)): .Cell(flexcpData, .Rows - 1, 0) = aValor
        
            .Cell(flexcpText, .Rows - 1, 1) = Trim(tLArticulo.Text)
            aValor = CLng(tLArticulo.Tag): .Cell(flexcpData, .Rows - 1, 1) = aValor
            
            .Cell(flexcpText, .Rows - 1, 2) = tLCantidad.Text
            .Cell(flexcpData, .Rows - 1, 2) = idServCpia
            .Cell(flexcpFontBold, .Rows - 1, 2) = True: .Cell(flexcpForeColor, .Rows - 1, 2) = Colores.Blanco:: .Cell(flexcpBackColor, .Rows - 1, 2) = Colores.Azul
        End With
            
        tLArticulo.Text = "": tLCantidad.Text = "": cLLocal.Text = ""
        Foco cLLocal
    End If
    
End Sub

Private Sub tProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then BuscarCompras
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then
            If cMoneda.Enabled Then Foco cMoneda Else Foco tFecha
            Exit Sub
        End If
        
        Screen.MousePointer = 11
        Cons = "Select PMeCodigo, PMeFantasia as 'Nombre Fantasía', PMeNombre as 'Razón Social' from ProveedorMercaderia " _
                & " Where PMeNombre like '" & Trim(tProveedor.Text) & "%' Or PMeFantasia like '" & Trim(tProveedor.Text) & "%'"
        
        Dim miHelp As New clsListadeAyuda, aIDSel As Long
        aIDSel = miHelp.ActivarAyuda(cBase, Cons, 5500, 1, "Lista de Proveedores")
        Me.Refresh
        If aIDSel <> 0 Then
            tProveedor.Text = Trim(miHelp.RetornoDatoSeleccionado(1))
            tProveedor.Tag = miHelp.RetornoDatoSeleccionado(0)
        End If
        Set miHelp = Nothing
        Screen.MousePointer = 0
        
        If Val(tProveedor.Tag) <> 0 Then If Not cMoneda.Enabled Then Foco cMoneda
        
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    LimpioFicha
    
    CargoDocumentos         'Tipos de Documentos de Compra de Mercaderia
    
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cLLocal
    
    DeshabilitoIngreso
    InicializoGrillas
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Err.Description
End Sub

Private Sub InicializoGrillas()

    With vsArticulo
        .Rows = 1: .Cols = 1
        .Editable = False
        .ExtendLastCol = True
        .FormatString = "Artículo|>Costo Unitario|>Q|>Subtotal"
        
        .WordWrap = False
        .ColWidth(0) = 4110: .ColWidth(2) = 900:
    End With
    
    With vsRemito
        .Rows = 1: .Cols = 1
        .Editable = False
        .ExtendLastCol = True
        .FormatString = "Tipo|<Remito|Fecha|Articulo|>Q|>Devuelve"
        
        .WordWrap = False
        .ColWidth(0) = 1600: .ColWidth(1) = 1000: .ColWidth(2) = 1100: .ColWidth(3) = 2500: .ColWidth(4) = 800
    End With
    
    With vsLocal
        .Rows = 1: .Cols = 1
        .Editable = False
        .ExtendLastCol = True
        .FormatString = "Local|Artículo|>Q"
        
        .WordWrap = False
        .ColWidth(0) = 3000: .ColWidth(1) = 4000
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
        
End Sub

Private Sub Label12_Click()
    Foco tProveedor
End Sub

Private Sub Label2_Click()
    Foco tFecha
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub
Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub AccionGrabar()

Dim aCodigoCompra As Long
    
    If Not ValidoDatos Then Exit Sub
    If MsgBox("Confirma almacenar los datos ingresados.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub

    aCodigoCompra = gFactura
    FechaDelServidor
    On Error GoTo ErrGD
    cBase.BeginTrans            '--------------------------------------------------------------------------------
    On Error GoTo ErrResumo
    
    'Grabo los renglones de compra --------------------------------------------------------------------------------
    Cons = "Select * from CompraRenglon Where CReCompra = " & aCodigoCompra
    Set RsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    With vsArticulo
    For I = 1 To .Rows - 1
        RsCom.AddNew
        RsCom!CReCompra = aCodigoCompra
        RsCom!CReArticulo = .Cell(flexcpData, I, 0)
        RsCom!CReCantidad = .Cell(flexcpValue, I, 2)
        RsCom!CRePrecioU = Abs(.Cell(flexcpValue, I, 1))
        RsCom!CRePrecioReal = Abs(.Cell(flexcpValue, I, 1))
        RsCom!CReACostear = .Cell(flexcpValue, I, 2)
        RsCom.Update
    Next
    End With
    RsCom.Close
    '----------------------------------------------------------------------------------------------------------------------
    
    'Actualizo las cantidades en los remitos de compra-------------------------------------------------------------
    Dim pcurCantidad As Currency, plngArticulo As Long
    With vsRemito
    For I = 1 To .Rows - 1
        If .Cell(flexcpValue, I, 5) > 0 Then    'Si la cantidad <> 0 --> GRABO
            
            pcurCantidad = .Cell(flexcpValue, I, 5)
            plngArticulo = .Cell(flexcpData, I, 3)
            
            Cons = "Select * from RemitoCompraRenglon " _
                  & " Where RCRRemito = " & .Cell(flexcpData, I, 1) _
                  & " And RCRArticulo = " & plngArticulo
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            RsAux.Edit
            RsAux!RCREnCompania = RsAux!RCREnCompania - .Cell(flexcpValue, I, 5)
            RsAux.Update
            
            RsAux.Close
            
            'Doy BAJAS al Local COMPAÑÍA--------------------------------------------------------------------------------------------------------
            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paLocalCompañia, _
                 plngArticulo, pcurCantidad, paEstadoArticuloEntrega, -1
            
            MarcoMovimientoStockTotal plngArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, pcurCantidad, -1
            
            MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, paLocalCompañia, _
                 plngArticulo, pcurCantidad, paEstadoArticuloEntrega, -1, cTipo.ItemData(cTipo.ListIndex), aCodigoCompra
            '----------------------------------------------------------------------------------------------------------------------------------------
        End If
    Next
    End With
    '-----------------------------------------------------------------------------------------------------------------------
    Dim qServ As Long
    Dim rsS As rdoResultset
    Dim sQ As String
    
    'Actualizo el stock en los locales-----------------------------------------------------------------------------------
    With vsLocal
    For I = 1 To .Rows - 1
                
        qServ = 0
        If CStr(.Cell(flexcpData, I, 2)) <> "" And Val(.Cell(flexcpData, I, 0)) = 14 Then
            qServ = CumploServiciosCompañia(.Cell(flexcpData, I, 2), .Cell(flexcpValue, I, 2), aCodigoCompra, .Cell(flexcpData, I, 1))
        End If
        
        If Val(.Cell(flexcpValue, I, 2)) - qServ > 0 Then
            
            'Doy BAJAS al Local DESTINO--------------------------------------------------------------------------------------------------------
            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, .Cell(flexcpData, I, 0), _
                 .Cell(flexcpData, I, 1), .Cell(flexcpValue, I, 2) - qServ, paEstadoArticuloEntrega, -1
            
            MarcoMovimientoStockTotal .Cell(flexcpData, I, 1), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, .Cell(flexcpValue, I, 2) - qServ, -1
            
            MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, .Cell(flexcpData, I, 0), _
                .Cell(flexcpData, I, 1), .Cell(flexcpValue, I, 2) - qServ, paEstadoArticuloEntrega, -1, cTipo.ItemData(cTipo.ListIndex), aCodigoCompra
            '----------------------------------------------------------------------------------------------------------------------------------------
        End If
    Next
    End With
    '-----------------------------------------------------------------------------------------------------------------------
        
    cBase.CommitTrans           '------------------------------------------------------------------------------
    
    AccionCancelar
    Exit Sub

ErrGD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción.", Err.Description
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    Screen.MousePointer = 0
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción.", Err.Description
    Exit Sub
End Sub

Private Sub AccionModificar()
    
    If gFactura = 0 Then Exit Sub
    
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    
    Screen.MousePointer = 0
    
End Sub

Private Sub AccionCancelar()
    
    Screen.MousePointer = vbHourglass
    
    Dim mAuxID As Long
    mAuxID = gFactura
    
    LimpioFicha
    
    DeshabilitoIngreso
    Botones True, False, False, False, False, Toolbar1, Me
    
    
    If mAuxID > 0 Then CargoDatosCompra mAuxID
    On Error Resume Next
    tCodigo.SetFocus
    Screen.MousePointer = vbDefault
    
    
    
End Sub


Private Sub tTCDolar_GotFocus()
    tTCDolar.SelStart = 0: tTCDolar.SelLength = Len(tTCDolar.Text)
End Sub

Private Sub tTCDolar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tTCDolar.Text) Then tTCDolar.Text = Format(tTCDolar.Text, "#.000")
        If tArticulo.Enabled Then Foco tArticulo Else Foco tComentario
    End If
End Sub

Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0: tComentario.SelLength = Len(tComentario.Text)
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tFactura_GotFocus()
    tFactura.SelStart = 0: tFactura.SelLength = Len(tFactura.Text)
End Sub

Private Sub tFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tProveedor
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0: tFecha.SelLength = Len(tFecha.Text)
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tFecha.Text) = "" Then Foco cTipo: Exit Sub
        
        If Not IsDate(tFecha.Text) Then
            MsgBox " La fecha ingresada no es correcta.", vbExclamation, "Posible Error"
            Foco tFecha
        Else
            Foco cTipo
        End If
    End If
    
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": 'AccionNuevo
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        
        Case "modificar": AccionModificar
        
        Case "salir": Unload Me
    End Select

End Sub

Private Sub DeshabilitoIngreso()
    
    tFecha.Enabled = True: tFecha.BackColor = Colores.Blanco
    tProveedor.Enabled = True: tProveedor.BackColor = Colores.Blanco
    tCodigo.Enabled = True: tCodigo.BackColor = Colores.Blanco
    
    cTipo.Enabled = False: cTipo.BackColor = Inactivo
    tFactura.Enabled = False: tFactura.BackColor = Inactivo
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tNeto.Enabled = False: tNeto.BackColor = Inactivo
    tTCDolar.Enabled = False: tTCDolar.BackColor = Inactivo
    tComentario.Enabled = False: tComentario.BackColor = Inactivo
    
    tArticulo.Enabled = False: tArticulo.BackColor = Inactivo
    tUnitario.Enabled = False: tUnitario.BackColor = Inactivo
    tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
    
    vsArticulo.Enabled = False
    vsArticulo.Editable = False: vsRemito.Editable = False
    
    bRemito.Enabled = False
    cLLocal.Enabled = False: cLLocal.BackColor = Colores.Inactivo
    tLArticulo.Enabled = False: tLArticulo.BackColor = Colores.Inactivo
    tLCantidad.Enabled = False: tLCantidad.BackColor = Colores.Inactivo

End Sub

Private Sub HabilitoIngreso()
    
    tCodigo.Enabled = False: tCodigo.BackColor = Colores.Gris
    tFecha.Enabled = False: tFecha.BackColor = Colores.Gris
    tProveedor.Enabled = False: tProveedor.BackColor = Colores.Gris
            
    tArticulo.Enabled = True: tArticulo.BackColor = Colores.Blanco
    tUnitario.Enabled = True: tUnitario.BackColor = Colores.Blanco
    tCantidad.Enabled = True: tCantidad.BackColor = Colores.Blanco
    
    vsArticulo.Enabled = True
    vsRemito.Editable = True
    
    bRemito.Enabled = True
    cLLocal.Enabled = True: cLLocal.BackColor = Colores.Blanco
    tLArticulo.Enabled = True: tLArticulo.BackColor = Colores.Blanco
    tLCantidad.Enabled = True: tLCantidad.BackColor = Colores.Blanco
    
End Sub

Private Sub LimpioFicha()
    
    tCodigo.Text = "": tFecha.Text = ""
    cTipo.Text = "": tFactura.Text = ""
    tProveedor.Text = ""

    cMoneda.Text = "": tNeto.Text = ""
    tTCDolar.Text = ""
    lTC.Caption = ""
    
    tComentario.Text = ""
    tArticulo.Text = "": tCantidad.Text = "": tUnitario.Text = ""
    vsArticulo.Rows = 1: vsRemito.Rows = 1
    
    vsLocal.Rows = 1
    cLLocal.Text = "": tLArticulo.Text = "": tLCantidad.Text = ""
    
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = 0
End Sub

Private Sub tProveedor_GotFocus()
    tProveedor.SelStart = 0: tProveedor.SelLength = Len(tProveedor.Text)
End Sub

Private Function ValidoDatos() As Boolean

    ValidoDatos = False
    
    If vsArticulo.Rows < 2 Then
        MsgBox "No se han ingresado los artículos del documento.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Function
    End If
    
    If Not IsDate(tFecha.Text) Then
        MsgBox "Ingrese la fecha del documento.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    
    If cTipo.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de documento ingresado.", vbExclamation, "ATENCIÓN"
        Foco cTipo: Exit Function
    End If
    
    If Val(tProveedor.Tag) = 0 Then
        MsgBox "Seleccione el proveedor de la mercadería.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Seleccione la moneda del documento.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
      
    'Valido si todos los articulos fueron ingresados
    Dim J As Integer
    Dim aCantidad As Currency, aCosto As Currency
    aCosto = 0
    For I = 1 To vsArticulo.Rows - 1
        aCantidad = 0
        aCosto = aCosto + vsArticulo.Cell(flexcpValue, I, 3)
        With vsRemito
            For J = 1 To .Rows - 1
                If .Cell(flexcpData, J, 3) = vsArticulo.Cell(flexcpData, I, 0) Then aCantidad = aCantidad + .Cell(flexcpValue, J, 5)
            Next J
        End With
        With vsLocal
            For J = 1 To .Rows - 1
                If .Cell(flexcpData, J, 1) = vsArticulo.Cell(flexcpData, I, 0) Then aCantidad = aCantidad + .Cell(flexcpValue, J, 2)
            Next J
        End With
        
        'If sNuevo Then      'Solo para nuevos ingresos
            If aCantidad <> vsArticulo.Cell(flexcpValue, I, 2) Then
                MsgBox "La cantidad a devolver para el artículo <<" & Trim(vsArticulo.Cell(flexcpText, I, 0)) & ">> no coincide con la de la nota." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                        & "En Nota: " & vsArticulo.Cell(flexcpValue, I, 2) & Chr(vbKeyReturn) & "Remitos + Locales: " & aCantidad, vbExclamation, "ATENCIÓN"
                Exit Function
            End If
        'End If
    Next I
    
    If Abs(aCosto) <> CCur(Abs(tNeto.Text)) Then
        If MsgBox("El Importe neto de la nota es de " & Format(Abs(tNeto.Text), FormatoMonedaP) & " y la suma de los costos es de " & Format(aCosto, FormatoMonedaP) & Chr(vbKeyReturn) _
                     & "Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Costos/Importe nota") = vbNo Then Exit Function
    End If
    ValidoDatos = True
    
End Function

Private Sub CargoDocumentos()

    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaCredito)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraNotaCredito
    
    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaDevolucion)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraNotaDevolucion
    
End Sub

Public Sub BuscoArticulo(aControl As TextBox, Optional Codigo As String = "", Optional Nombre As String = "")

    On Error GoTo ErrBAC
    Screen.MousePointer = 11
    If Trim(Codigo) <> "" Then          'Articulo por Codigo--------------------------------------------------------
        Cons = "Select ArtID, ArtNombre From Articulo Where ArtCodigo = " & CLng(Codigo)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If RsAux.EOF Then
            MsgBox "No se encontró un artículo para el código ingresado.", vbExclamation, "ATENCIÓN"
            aControl.Text = "": aControl.Tag = 0
        Else
            aControl.Text = Trim(RsAux!ArtNombre)
            aControl.Tag = RsAux!ArtID
        End If
        RsAux.Close
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If Trim(Nombre) <> "" Then          'Articulo por Nombre------------------------------------------------------
        Cons = "Select ArtId, ArtCodigo as 'Código', ArtNombre as 'Nombre' from Articulo" _
                & " Where ArtNombre LIKE '" & Nombre & "%'" _
                & " Order by ArtNombre"
        
        Dim miHelp As New clsListadeAyuda, mIDSel As Long, aItem As String
        mIDSel = miHelp.ActivarAyuda(cBase, Cons, 5500, 1, "Lista de Artículos")
        Me.Refresh
        If mIDSel <> 0 Then
            mIDSel = miHelp.RetornoDatoSeleccionado(0)
            aItem = Trim(miHelp.RetornoDatoSeleccionado(1))
        End If
        Set miHelp = Nothing
                    
        'Screen.MousePointer = 0: Me.Refresh
        
        If mIDSel <> 0 Then
            aControl.Text = aItem
            BuscoArticulo aControl, Codigo:=aItem
        End If
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrBAC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
End Sub

Private Sub tUnitario_GotFocus()
    tUnitario.SelStart = 0: tUnitario.SelLength = Len(tUnitario.Text)
End Sub

Private Sub tUnitario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUnitario.Text) Then Foco tCantidad
    End If
    
End Sub

Private Sub tUnitario_LostFocus()

    If IsNumeric(tUnitario.Text) Then
        tUnitario.Text = Format(tUnitario.Text, FormatoMonedaP)
        Foco tCantidad
    Else
        tUnitario.Text = ""
    End If
    
End Sub

Private Sub vsArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyDelete        'Elimino los articulos (remito y locales)--------------------------------------------
            If vsArticulo.Rows = 1 Then Exit Sub
            Dim aRow As Integer
            aRow = 1
            With vsRemito
                For I = 1 To .Rows - 1
                    If vsArticulo.Cell(flexcpData, .Row, 0) = .Cell(flexcpData, aRow, 3) Then .RemoveItem aRow Else aRow = aRow + 1
                Next I
            End With
            
            aRow = 1
            With vsLocal
                For I = 1 To .Rows - 1
                    If vsArticulo.Cell(flexcpData, .Row, 0) = .Cell(flexcpData, aRow, 3) Then .RemoveItem aRow Else aRow = aRow + 1
                Next I
            End With
            
            vsArticulo.RemoveItem vsArticulo.Row
        '--------------------------------------------------------------------------------------------------------------------
    End Select
            
End Sub

Private Sub vsLocal_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 2 Then Cancel = True
End Sub

Private Sub vsLocal_GotFocus()
    If vsLocal.Rows > 1 Then vsLocal.Select 1, 2
End Sub

Private Sub vsLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If vsLocal.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyDelete: vsLocal.RemoveItem vsLocal.Row
    End Select
    
End Sub

Private Sub vsRemito_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 5 Then Cancel = True
End Sub

Private Sub vsRemito_GotFocus()
    If vsRemito.Rows > 1 Then vsRemito.Select 1, 5
End Sub

Private Sub vsRemito_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vsRemito
        If CCur(.EditText) < 0 Then
            Cancel = True
            MsgBox "La cantidad a devolver no puede ser negativa.", vbExclamation, "ATENCIÓN"
        End If
        
        If .Cell(flexcpValue, Row, 4) < CLng(.EditText) Then
            Cancel = True
            MsgBox "La cantidad a devolver no puede ser mayor a la que figura en el remito.", vbExclamation, "ATENCIÓN"
        End If
        
    End With
    
End Sub

Private Sub BuscarCompras()
    On Error GoTo errBuscar
    gFactura = 0
    Botones True, False, False, False, False, Toolbar1, Me
    
    If Not IsDate(tFecha.Text) And Val(tProveedor.Tag) = 0 Then Exit Sub
    
    Cons = "SELECT ComCodigo, ComCodigo Compra, ComFecha Fecha, ComSerie Serie, ComNumero 'Número', MonSigno Moneda, ComImporte Importe, ComComentario Comentarios" & _
        " From Compra, Moneda " & _
        " Where ComMoneda = MonCodigo" & _
        " And ComTipoDocumento in (" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ")"
    
    If IsDate(tFecha.Text) Then Cons = Cons & " And ComFecha >= " & Format(tFecha.Text, "'mm/dd/yyyy'")
    If Val(tProveedor.Tag) <> 0 Then Cons = Cons & " And ComProveedor = " & Val(tProveedor.Tag)
    
    Cons = Cons & " And ComCodigo IN (Select GSrIDCompra from GastoSubRubro " & _
                                     " Where GSrIDSubRubro = " & paSubrubroCompraMercaderia & ")"
    
    Cons = Cons & " Order by ComFecha asc"
    
    Dim aIDCompra As Long: aIDCompra = 0
    Dim aLista As New clsListadeAyuda
    
    aIDCompra = aLista.ActivarAyuda(cBase, Cons, 8000, 1)
    Me.Refresh
    If aIDCompra <> 0 Then aIDCompra = aLista.RetornoDatoSeleccionado(0)
    Set aLista = Nothing
    
    If aIDCompra <> 0 Then CargoDatosCompra aIDCompra
            
    Screen.MousePointer = 0
    Exit Sub
errBuscar:
    clsGeneral.OcurrioError "Error al buscar las devoluciones.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosCompra(aIDCompra As Long)
On Error GoTo errCargar
    
    gFactura = 0
    Botones False, False, False, False, False, Toolbar1, Me
    
    Cons = "Select * from Compra, ProveedorMercaderia " & _
        " Where ComProveedor = PMeCodigo " & _
        " And ComCodigo = " & aIDCompra & _
        " And ComTipoDocumento in (" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        tCodigo.Text = RsAux!ComCodigo
        tFecha.Text = Format(RsAux!ComFecha, "dd/mm/yyyy")
        
        BuscoCodigoEnCombo cTipo, RsAux!ComTipoDocumento
        BuscoCodigoEnCombo cMoneda, RsAux!ComMoneda
        
        If Not IsNull(RsAux!ComNumero) Then tFactura.Text = RsAux!ComNumero Else tFactura.Text = ""
        
        tProveedor.Text = Trim(RsAux!PMeFantasia)
        tProveedor.Tag = RsAux!PMeCodigo
        
        tNeto.Text = Format(RsAux!ComImporte, FormatoMonedaP)         'ComImporte = NETO

        tTCDolar.Text = Format(RsAux!ComTC, "0.000")
        If Not IsNull(RsAux!ComComentario) Then tComentario.Text = Trim(RsAux!ComComentario) Else tComentario.Text = ""
        
        gFactura = aIDCompra
    End If
    RsAux.Close
    
    vsArticulo.Rows = 1
    
    If gFactura <> 0 Then
        
        Cons = "Select * from CompraRenglon, Articulo " & _
                   " Where CReCompra = " & aIDCompra & _
                   " And CReArticulo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        Do While Not RsAux.EOF
            With vsArticulo
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!ArtNombre)
                aValor = RsAux!ArtCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
                .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!CRePrecioU, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 2) = RsAux!CReCantidad
                .Cell(flexcpText, .Rows - 1, 3) = Format(.Cell(flexcpValue, .Rows - 1, 1) * RsAux!CReCantidad, "#,##0.00")
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        Dim bEsCompraM As Boolean
        bEsCompraM = True
        If vsArticulo.Rows = vsArticulo.FixedRows Then
            Cons = "Select * from GastoSubRubro " & _
                    " Where GSrIDCompra = " & gFactura & _
                    " And GSrIDSubRubro = " & paSubrubroCompraMercaderia
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            bEsCompraM = Not RsAux.EOF
            RsAux.Close
        End If
    
        If bEsCompraM And Not (vsArticulo.Rows > vsArticulo.FixedRows) Then
            Botones True, True, False, False, False, Toolbar1, Me
        End If
    
    
    End If
    
    
    If gFactura = 0 Then
        MsgBox "Posiblemente la compra seleccionada no es una Compra de Mercadería." & vbCrLf & _
               "Causa: No existen valores asignados al rubro Mercadería.", vbInformation, "Es compra de Mercadería ?"
    End If
    
    
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos de la devolución.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function HayServiciosEnCompañia(ByVal idArt As Long) As String
Dim errBC
Dim sQ As String
    HayServiciosEnCompañia = ""

Dim IdsSelect As String
Dim iArtsServ As Integer

    sQ = "SELECT SerCodigo Servicio, SerFecha Fecha " & _
        "FROM CGSA.dbo.Servicio " & _
        "INNER JOIN CGSA.dbo.Producto ON SerProducto = ProCodigo AND ProArticulo = " & idArt & _
        " WHERE SerLocalReparacion = 14 AND SerEstadoServicio = 3 AND SerCliente = 1"
    Dim rsS As rdoResultset
    Set rsS = cBase.OpenResultset(sQ, rdOpenDynamic, rdConcurValues)
    If Not rsS.EOF Then
        rsS.Close
        Dim sHelp As String
        Do While iArtsServ < CLng(tLCantidad.Text)
            sHelp = sQ
            If IdsSelect <> "" Then
                sHelp = sHelp & " AND SerCodigo NOT IN (" & IdsSelect & ")"
            End If
            Dim oAyuda As clsListadeAyuda
            Set oAyuda = New clsListadeAyuda
            If oAyuda.ActivarAyuda(cBase, sHelp, 4000, 0, "Servicios en compañia") > 0 Then
                iArtsServ = iArtsServ + 1
                If IdsSelect <> "" Then IdsSelect = IdsSelect & ","
                IdsSelect = IdsSelect & oAyuda.RetornoDatoSeleccionado(0)
            Else
                Exit Do
            End If
            Set oAyuda = Nothing
        Loop
    Else
        rsS.Close
    End If
    HayServiciosEnCompañia = IdsSelect
    Exit Function
errBC:
    clsGeneral.OcurrioError "Error al buscar los servicios de compañia", Err.Description, "Error"
End Function

Private Function CumploServiciosCompañia(ByVal idServicios As String, ByVal CantTotal As Integer, ByVal aCodigoCompra As Long, ByVal idArticulo As Long) As Integer
Dim idServ As Long
Dim rsS As rdoResultset
Dim sQ As String

    'Retorno cantidad de bajas por servicio.
    CumploServiciosCompañia = 0

'Cumplo el menor id de servicio para el artículo.
    sQ = "SELECT * " & _
        "FROM CGSA.dbo.Servicio " & _
        " WHERE SerCodigo IN (" & idServicios & ")"
'        " WHERE SerLocalReparacion = 14 AND SerEstadoServicio = 3 AND SerCliente = 1 " & _
'        " AND SerProducto IN (SELECT ProCodigo FROM CGSA.dbo.Producto WHERE ProArticulo = " & idArticulo & " AND ProCliente = 1) ORDER BY SerCodigo"
    Set rsS = cBase.OpenResultset(sQ, rdOpenDynamic, rdConcurValues)
    Do While Not rsS.EOF And CumploServiciosCompañia < CantTotal
        CumploServiciosCompañia = CumploServiciosCompañia + 1
        rsS.Edit
        rsS!SerFCumplido = Format(gFechaServidor, sqlFormatoF)
        rsS!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
        rsS!SerUsuario = paCodigoDeUsuario
        rsS!SerEstadoServicio = 0   'Anulo.
        If IsNull(rsS("SerComentario")) Then
            rsS("SerComentario") = "Baja x Nota."
        Else
            rsS("SerComentario") = Trim(rsS("SerComentario")) & ", Baja x Nota."
        End If
        rsS.Update
        rsS.MoveNext
    Loop
    rsS.Close
    
    If CumploServiciosCompañia > 0 Then
    'Tengo que restar del stock con estado a recuperar.
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, 14, _
          idArticulo, CCur(CumploServiciosCompañia), 275, -1
        
        MarcoMovimientoStockTotal idArticulo, TipoEstadoMercaderia.Fisico, 275, CCur(CumploServiciosCompañia), -1
        
        MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, 14, _
             idArticulo, CCur(CumploServiciosCompañia), 275, -1, cTipo.ItemData(cTipo.ListIndex), aCodigoCompra
            
    End If

End Function
