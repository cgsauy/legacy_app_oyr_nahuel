VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmCosteo 
   Caption         =   "Costeo de Carpetas"
   ClientHeight    =   8310
   ClientLeft      =   2100
   ClientTop       =   1725
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCosteo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   12345
   Begin VB.Frame frGasto 
      Caption         =   "Gastos asignados / Distribución"
      ForeColor       =   &H00000080&
      Height          =   3495
      Left            =   360
      TabIndex        =   16
      Top             =   2160
      Width           =   8295
      Begin VSFlex6DAOCtl.vsFlexGrid vsGasto 
         Height          =   1005
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1773
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
         Rows            =   20
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
      Begin ComctlLib.TabStrip tbCosteo 
         Height          =   1095
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1931
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "&Asignación de Gastos"
               Key             =   "asignacion"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Costeo por Unidad  "
               Key             =   "totales"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsTotales 
         Height          =   1005
         Left            =   1920
         TabIndex        =   13
         Top             =   2280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1773
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsCosteo 
         Height          =   1005
         Left            =   4680
         TabIndex        =   14
         Top             =   2280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1773
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
   End
   Begin ComctlLib.TabStrip tbArticulo 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1508
      TabWidthStyle   =   2
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Carpeta &Madre"
            Key             =   "madre"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Embarque"
            Key             =   "embarque"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Sub Carpeta"
            Key             =   "sub"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   8055
      Width           =   12345
      _ExtentX        =   21775
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
            Object.Width           =   13547
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsArticulo 
      Height          =   1005
      Left            =   3120
      TabIndex        =   8
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1773
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsArticuloE 
      Height          =   1005
      Left            =   5400
      TabIndex        =   9
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1773
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsArticuloS 
      Height          =   1005
      Left            =   7800
      TabIndex        =   10
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1773
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
   Begin VB.Frame frDatos 
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   9735
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   2520
         ScaleHeight     =   405
         ScaleWidth      =   2535
         TabIndex        =   18
         Top             =   120
         Width           =   2535
         Begin VB.CommandButton bConsultar 
            Height          =   310
            Left            =   120
            Picture         =   "frmCosteo.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Ejecutar."
            Top             =   90
            Width           =   310
         End
         Begin VB.CommandButton bCancelar 
            Height          =   310
            Left            =   2040
            Picture         =   "frmCosteo.frx":060C
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Salir."
            Top             =   90
            Width           =   310
         End
         Begin VB.CommandButton bNoFiltros 
            Height          =   310
            Left            =   1560
            Picture         =   "frmCosteo.frx":070E
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Limpiar datos."
            Top             =   90
            Width           =   310
         End
         Begin VB.CommandButton bImprimir 
            Height          =   310
            Left            =   1200
            Picture         =   "frmCosteo.frx":0AD4
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Imprimir."
            Top             =   90
            Width           =   310
         End
         Begin VB.CommandButton bCostear 
            Height          =   310
            Left            =   480
            Picture         =   "frmCosteo.frx":0BD6
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Costear."
            Top             =   90
            Width           =   310
         End
      End
      Begin AACombo99.AACombo cFolder 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label1 
         Caption         =   "&Carpeta:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Apertura:"
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   255
         Width           =   735
      End
      Begin VB.Label lApertura 
         Caption         =   "N/D"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6240
         TabIndex        =   21
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Arribo Merc.:"
         Height          =   255
         Left            =   7440
         TabIndex        =   20
         Top             =   255
         Width           =   975
      End
      Begin VB.Label lArribo 
         Caption         =   "N/D"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8520
         TabIndex        =   19
         Top             =   255
         Width           =   975
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   6135
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   11355
      _Version        =   196608
      _ExtentX        =   20029
      _ExtentY        =   10821
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
      TableBorder     =   3
      PreviewMode     =   1
   End
   Begin VB.Menu MnuAcceso 
      Caption         =   "MnuAcceso"
      Visible         =   0   'False
      Begin VB.Menu MnuVerGasto 
         Caption         =   "Ver Gasto"
      End
      Begin VB.Menu MnuVerEmbarque 
         Caption         =   "Ver Embarque"
      End
      Begin VB.Menu MnuL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBuGastos 
         Caption         =   "Buscar Gastos"
      End
      Begin VB.Menu MnuBuEmbarques 
         Caption         =   "Buscar Embarques"
      End
      Begin VB.Menu MnuBuSub 
         Caption         =   "Buscar Subcarpetas"
      End
   End
End
Attribute VB_Name = "frmCosteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
'   Parametos para inicializar: TipoFolder + IdFolder
'------------------------------------------------------------------------------------------
Option Explicit

Dim aValor As Long
Dim aCarpeta As Long, aEmbarque As Long, aSub As Long
Dim aFolder As Long, aTipoFolder As Integer

Dim aPrecioTotal As Currency     'Guardo la Suma de $ (por si no coincide con divisa)
Dim aCantidadTotal As Long
Dim aVolumenTotal As Currency

Dim aPrecioTotalE As Currency     'Nivel Embarque
Dim aCantidadTotalE As Long
Dim aVolumenTotalE As Currency

Dim aPrecioTotalS As Currency     'Nivel Subcarpeta
Dim aCantidadTotalS As Long
Dim aVolumenTotalS As Currency

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bCostear_Click()
    AccionCostear
End Sub

Private Sub bImprimir_Click()
    AccionImprimir
End Sub

Private Sub bNoFiltros_Click()
    AccionLimpiar
End Sub

Private Sub cFolder_Change()
    AccionLimpiar True
End Sub

Private Sub cFolder_Click()
    AccionLimpiar True
End Sub

Private Sub cFolder_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionConsultar
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub AccionConsultar()
    If cFolder.ListIndex <> -1 Then BuscoFolder
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    With vsListado
        .Visible = False
        .Orientation = orPortrait
        .MarginRight = 350: .MarginTop = 750: .MarginBottom = 750: .MarginLeft = 550
        .PaperSize = 1
    End With
    
    CargoComboFolder
    InicializoGrillas
    vsArticulo.ZOrder 0
    vsTotales.ZOrder 0
    
    ObtengoSeteoForm Me, , , 10200, 6500
    
    If Trim(Command()) <> "" Then
        BuscoCodigoEnCombo cFolder, Val(Trim(Command()))
        If cFolder.ListIndex <> -1 Then BuscoFolder
    End If
    
    Screen.MousePointer = 0
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.Width < 10200 Or Me.Height < 6500 Then Exit Sub
    
    frDatos.Width = Me.Width - 320
    
    'Tab y Listas de articulos------------------------------------------------------------------------------------------------------
    tbArticulo.Top = 700
    tbArticulo.Height = 1700
    tbArticulo.Left = 120
    tbArticulo.Width = Me.Width - 320
    
    vsArticulo.Left = tbArticulo.ClientLeft
    vsArticulo.Top = tbArticulo.ClientTop
    vsArticulo.Height = tbArticulo.Height - (tbArticulo.ClientTop - tbArticulo.Top) - 100
    vsArticulo.Width = tbArticulo.ClientWidth
    
    vsArticuloE.Left = vsArticulo.Left: vsArticuloE.Top = vsArticulo.Top: vsArticuloE.Height = vsArticulo.Height: vsArticuloE.Width = vsArticulo.Width
    vsArticuloS.Left = vsArticulo.Left: vsArticuloS.Top = vsArticulo.Top: vsArticuloS.Height = vsArticulo.Height: vsArticuloS.Width = vsArticulo.Width
    '------------------------------------------------------------------------------------------------------------------------------------
    
    'Lista de gastos y distribucion----------------------------------------------------------------------------------------------------
    frGasto.Top = tbArticulo.Top + tbArticulo.Height + 100
    frGasto.Left = tbArticulo.Left
    frGasto.Width = tbArticulo.Width
    frGasto.Height = Me.Height - frGasto.Top - 800
    
    tbCosteo.Left = 120
    tbCosteo.Width = frGasto.Width - 240
    tbCosteo.Height = 1695
    tbCosteo.Top = frGasto.Height - tbCosteo.Height - 120
    
    vsGasto.Left = 120: vsGasto.Top = 250
    vsGasto.Width = frGasto.Width - 240
    vsGasto.Height = frGasto.Height - tbCosteo.Height - 400
    
    vsTotales.Left = tbCosteo.ClientLeft
    vsTotales.Top = tbCosteo.ClientTop
    vsTotales.Height = tbCosteo.Height - (tbCosteo.ClientTop - tbCosteo.Top) - 100
    vsTotales.Width = tbCosteo.ClientWidth
    
    vsCosteo.Left = vsTotales.Left: vsCosteo.Top = vsTotales.Top: vsCosteo.Height = vsTotales.Height: vsCosteo.Width = vsTotales.Width
    
    vsListado.Top = tbArticulo.Top: vsListado.Left = tbArticulo.Left
    vsListado.Width = tbArticulo.Width: vsListado.Height = Me.ScaleHeight - (tbArticulo.Top + Status.Height + 40)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    cBase.Close
    'Set msgError = Nothing
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    
    If Me.Width > 10200 And Me.Height > 6500 Then GuardoSeteoForm Me
    
End Sub

Private Sub BuscoFolder()

    On Error GoTo errBusco
    Screen.MousePointer = 11
    aCarpeta = 0: aEmbarque = 0: aSub = 0
    AccionLimpiar True
    
    aTipoFolder = Mid(cFolder.ItemData(cFolder.ListIndex), 1, 1)
    aFolder = Mid(cFolder.ItemData(cFolder.ListIndex), 2, Len(CStr(cFolder.ItemData(cFolder.ListIndex))))
    
    Select Case aTipoFolder
        Case Folder.cFEmbarque
            Cons = "Select * from Embarque, Carpeta Where EmbId = " & aFolder & " And EmbCarpeta = CarID"
            'Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If ObtenerResultSet(cBase, RsAux, Cons, logImportaciones) = RAQ_SinError Then
                If Not IsNull(RsAux!CarFApertura) Then lApertura.Caption = Format(RsAux!CarFApertura, FormatoFP)
                If Not IsNull(RsAux!EmbFLocal) Then lArribo.Caption = Format(RsAux!EmbFLocal, FormatoFP)
                aCarpeta = RsAux!CarID
                aEmbarque = aFolder
            End If
            RsAux.Close
            
        Case Folder.cFSubCarpeta
            Cons = " Select * from SubCarpeta, Embarque, Carpeta " _
                    & " Where SubId = " & aFolder & " And SubEmbarque = EmbID And EmbCarpeta = CarID"
            'Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If ObtenerResultSet(cBase, RsAux, Cons, logImportaciones) = RAQ_SinError Then
                If Not IsNull(RsAux!CarFApertura) Then lApertura.Caption = Format(RsAux!CarFApertura, FormatoFP)
                If Not IsNull(RsAux!EmbFLocal) Then lArribo.Caption = Format(RsAux!EmbFLocal, FormatoFP)
                aCarpeta = RsAux!CarID
                aEmbarque = RsAux!EmbID
            End If
            RsAux.Close
            aSub = aFolder

    End Select
    
    CargoArticulosCarpeta aCarpeta
    CargoArticulosEmbarque aEmbarque
    If aSub <> 0 Then CargoArticulosSubcarpeta aSub
    
    CargoGastos
        
    Screen.MousePointer = 0
    
    Exit Sub
errBusco:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar los datos del folder.", Err.Description
End Sub

Private Sub CargoArticulosCarpeta(idCarpeta As Long)

    On Error GoTo errCargar
    vsArticulo.Rows = 1
    aPrecioTotal = 0: aCantidadTotal = 0: aVolumenTotal = 0
    
    'Cargo la lista de articulos para la carpeta---------------------------------------------------------------------
    Cons = "Select ArtID, ArtCodigo, ArtNombre, ArtVolumen, AFoPUnitario, Cantidad = Sum(AFoCantidad) From ArticuloFolder, Articulo" _
            & " Where AFoTipo = " & Folder.cFEmbarque _
            & " And AFoCodigo IN (Select EmbId from Embarque Where EmbCarpeta = " & idCarpeta & ")" _
            & " And AFoArticulo = ArtID" _
            & " Group by ArtID, ArtCodigo, ArtNombre, ArtVolumen, AFoPUnitario"
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    With vsArticulo
        If ObtenerResultSet(cBase, RsAux, Cons, logImportaciones) = RAQ_SinError Then
            Do While Not RsAux.EOF
                .AddItem Format(RsAux!ArtCodigo, "000,000"), .Rows
                aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 2) = RsAux!Cantidad
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!AFoPUnitario, "##,##0.00")
                
                .Cell(flexcpText, .Rows - 1, 4) = Format(.Cell(flexcpValue, .Rows - 1, 3) * .Cell(flexcpText, .Rows - 1, 2), "##,##0.00")   'Precio T
                
                If Not IsNull(RsAux!ArtVolumen) Then .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!ArtVolumen, "##,##0.00") Else: .Cell(flexcpText, .Rows - 1, 5) = "1.00"
                .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpValue, .Rows - 1, 5) * .Cell(flexcpText, .Rows - 1, 2), "##,##0.00")
                
                aCantidadTotal = aCantidadTotal + RsAux!Cantidad
                aVolumenTotal = aVolumenTotal + .Cell(flexcpValue, .Rows - 1, 6)
                aPrecioTotal = aPrecioTotal + .Cell(flexcpValue, .Rows - 1, 4)
                
                RsAux.MoveNext
            Loop
        End If
        RsAux.Close
        .Sort = flexSortUseColSort
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 2) = aCantidadTotal
        .Cell(flexcpText, .Rows - 1, 4) = Format(aPrecioTotal, "##,##0.00")
        .Cell(flexcpText, .Rows - 1, 6) = Format(aVolumenTotal, "##,##0.00")
        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio
    End With
    
    '-----------------------------------------------------------------------------------------------------------------------
    'A las cantidades totales, hay que restarles las cantidades ya costeadas para los subniveles
    '1) Resto las cantidades costeadas para los embarques (q' no van a ZF) de la Carpeta
    Cons = "Select * From Embarque, ArticuloFolder, Articulo" _
            & " Where EmbCarpeta = " & idCarpeta _
            & " And AFoTipo = " & Folder.cFEmbarque _
            & " And AFoCodigo = EmbID " _
            & " And EmbCosteado = 1" _
            & " And EmbLocal <> " & paLocalZF _
            & " And AFoArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        aCantidadTotal = aCantidadTotal - RsAux!AFoCantidad
        If Not IsNull(RsAux!ArtVolumen) Then aVolumenTotal = aVolumenTotal - (RsAux!AFoCantidad * Format(RsAux!ArtVolumen, FormatoMonedaP)) Else aVolumenTotal = aVolumenTotal - RsAux!AFoCantidad
        aPrecioTotal = aPrecioTotal - (RsAux!AFoCantidad * RsAux!AFoPUnitario)
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    '2) Resto las cantidades costeadas para las SubCarpetas
    Cons = "Select * From Subcarpeta, ArticuloFolder, Articulo" _
            & " Where SubEmbarque IN (Select EmbID from Embarque Where EmbCarpeta = " & idCarpeta & ")" _
            & " And SubID = AFoCodigo " _
            & " And AFoTipo = " & Folder.cFSubCarpeta _
            & " And AFoArticulo = ArtID" _
            & " And SubCosteada = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        aCantidadTotal = aCantidadTotal - RsAux!AFoCantidad
        If Not IsNull(RsAux!ArtVolumen) Then aVolumenTotal = aVolumenTotal - (RsAux!AFoCantidad * Format(RsAux!ArtVolumen, FormatoMonedaP)) Else aVolumenTotal = aVolumenTotal - RsAux!AFoCantidad
        aPrecioTotal = aPrecioTotal - (RsAux!AFoCantidad * RsAux!AFoPUnitario)
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de la carpeta.", Err.Description
End Sub

Private Sub CargoArticulosEmbarque(idEmbarque As Long)
    On Error GoTo errCargar
    vsArticuloE.Rows = 1
    aPrecioTotalE = 0: aCantidadTotalE = 0: aVolumenTotalE = 0
    
    Cons = "Select * From ArticuloFolder, Articulo" _
            & " Where AFoTipo = " & Folder.cFEmbarque _
            & " And AFoCodigo = " & idEmbarque _
            & " And AFoArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    With vsArticuloE
    Do While Not RsAux.EOF
        .AddItem Format(RsAux!ArtCodigo, "000,000"), .Rows
        aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
        
        .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
        .Cell(flexcpText, .Rows - 1, 2) = RsAux!AFoCantidad
        
        .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!AFoPUnitario, "##,##0.00")
        .Cell(flexcpText, .Rows - 1, 4) = Format(.Cell(flexcpValue, .Rows - 1, 3) * .Cell(flexcpText, .Rows - 1, 2), "##,##0.00")   'Precio T
        
        If Not IsNull(RsAux!ArtVolumen) Then .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!ArtVolumen, "##,##0.00") Else: .Cell(flexcpText, .Rows - 1, 5) = "1.00"
        .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpValue, .Rows - 1, 5) * .Cell(flexcpText, .Rows - 1, 2), "##,##0.00")
        
        For I = 1 To vsArticulo.Rows - 1
            'If vsArticulo.Cell(flexcpData, I, 0) = RsAux!ArtID Then
            If vsArticulo.Cell(flexcpData, I, 0) = RsAux!ArtID And Format(RsAux!AFoPUnitario, "##,##0.00") = vsArticulo.Cell(flexcpText, I, 3) Then
                .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!AFoCantidad * 100 / vsArticulo.Cell(flexcpValue, I, 2), "##,##0.00")
                aVolumenTotalE = aVolumenTotalE + RsAux!AFoCantidad * vsArticulo.Cell(flexcpValue, I, 5)
                aPrecioTotalE = aPrecioTotalE + RsAux!AFoCantidad * vsArticulo.Cell(flexcpValue, I, 3)
                Exit For
            End If
        Next I
        aCantidadTotalE = aCantidadTotalE + RsAux!AFoCantidad
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    .Sort = flexSortUseColSort
    .AddItem ""
    .Cell(flexcpText, .Rows - 1, 2) = aCantidadTotalE
    .Cell(flexcpText, .Rows - 1, 4) = Format(aPrecioTotalE, "##,##0.00")
    .Cell(flexcpText, .Rows - 1, 6) = Format(aVolumenTotalE, "##,##0.00")
    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio
    End With
    
    '-----------------------------------------------------------------------------------------------------------------------
    'A las cantidades totales, hay que restarles las cantidades ya costeadas para los subniveles
    '1) Resto las cantidades costeadas para las SubCarpetas del Embarque
    Cons = "Select * From SubCarpeta, ArticuloFolder, Articulo" _
            & " Where SubEmbarque = " & idEmbarque _
            & " And SubID = AFoCodigo" _
            & " And AFoTipo = " & Folder.cFSubCarpeta _
            & " And AFoArticulo = ArtID" _
            & " And SubCosteada = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        aCantidadTotalE = aCantidadTotalE - RsAux!AFoCantidad
        If Not IsNull(RsAux!ArtVolumen) Then aVolumenTotalE = aVolumenTotalE - (RsAux!AFoCantidad * Format(RsAux!ArtVolumen, FormatoMonedaP)) Else aVolumenTotalE = aVolumenTotalE - RsAux!AFoCantidad
        aPrecioTotalE = aPrecioTotalE - (RsAux!AFoCantidad * RsAux!AFoPUnitario)
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del embarque.", Err.Description
End Sub

Private Sub CargoArticulosSubcarpeta(idSubcarpeta As Long)
    On Error GoTo errCargar
    vsArticuloS.Rows = 1
    aPrecioTotalS = 0: aCantidadTotalS = 0: aVolumenTotalS = 0
    
    Cons = "Select * From ArticuloFolder, Articulo" _
            & " Where AFoTipo = " & Folder.cFSubCarpeta _
            & " And AFoCodigo = " & idSubcarpeta _
            & " And AFoArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    With vsArticuloS
    Do While Not RsAux.EOF
        .AddItem Format(RsAux!ArtCodigo, "000,000"), .Rows
        aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
        
        .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
        .Cell(flexcpText, .Rows - 1, 2) = RsAux!AFoCantidad
        
        .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!AFoPUnitario, "##,##0.00")
        .Cell(flexcpText, .Rows - 1, 4) = Format(.Cell(flexcpValue, .Rows - 1, 3) * .Cell(flexcpText, .Rows - 1, 2), "##,##0.00")   'Precio T
        
        If Not IsNull(RsAux!ArtVolumen) Then .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!ArtVolumen, "##,##0.00") Else: .Cell(flexcpText, .Rows - 1, 5) = "1.00"
        .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpValue, .Rows - 1, 5) * .Cell(flexcpText, .Rows - 1, 2), "##,##0.00")
        
        For I = 1 To vsArticuloE.Rows - 1
            If vsArticuloE.Cell(flexcpData, I, 0) = RsAux!ArtID Then
                .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!AFoCantidad * 100 / vsArticuloE.Cell(flexcpValue, I, 2), "##,##0.00")
                aVolumenTotalS = aVolumenTotalS + RsAux!AFoCantidad * vsArticuloE.Cell(flexcpValue, I, 5)
                aPrecioTotalS = aPrecioTotalS + RsAux!AFoCantidad * vsArticuloE.Cell(flexcpValue, I, 3)
                Exit For
            End If
        Next I
        aCantidadTotalS = aCantidadTotalS + RsAux!AFoCantidad
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    .AddItem ""
    .Cell(flexcpText, .Rows - 1, 2) = aCantidadTotalS
    .Cell(flexcpText, .Rows - 1, 4) = Format(aPrecioTotalS, "##,##0.00")
    .Cell(flexcpText, .Rows - 1, 6) = Format(aVolumenTotalS, "##,##0.00")
    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio
    End With
    
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de la subcarpeta.", Err.Description
End Sub

Private Function BuscoGastoEnColeccion(ByVal colOrigen As Collection, ByVal idSubro As Long) As clsGastoTotal
    Dim oGasto As clsGastoTotal
    For Each oGasto In colOrigen
        If oGasto.idSubro = idSubro Then
            Set BuscoGastoEnColeccion = oGasto
            Exit Function
        End If
    Next
End Function

Private Sub CargoGastos()
Dim aVolumenP As Currency, aVolumenD As Currency
Dim aDivisaP As Currency, aDivisaD As Currency
Dim aLinealP As Currency, aLinealD As Currency
Dim aCF As Currency: aCF = 0
Dim pblnEsDifCambio As Boolean
Dim oGasto As clsGastoTotal

    On Error GoTo errGastos
    aVolumenP = 0: aVolumenD = 0: aDivisaP = 0: aDivisaD = 0: aLinealP = 0: aLinealD = 0
    vsGasto.Rows = 1
    
    Cons = "Select dbo.Dolar(1, ComFecha) as TipoCambio, * from GastoImportacion, Subrubro, Compra" _
           & " Where GImIDSubrubro = SRuID" _
           & " And GImIdCompra = ComCodigo" _
           & " And ((GImNivelFolder = " & Folder.cFCarpeta & " And GImFolder = " & aCarpeta & ")" _
            & " OR (GImNivelFolder = " & Folder.cFEmbarque & " And GImFolder = " & aEmbarque & ")"
    If aSub <> 0 Then Cons = Cons & " OR (GImNivelFolder = " & Folder.cFSubCarpeta & " And GImFolder = " & aSub & ")"
    Cons = Cons & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    With vsGasto
    Do While Not RsAux.EOF
        'Data Row = 0   -> ID_Compra
        'Data Row = 1   -> ID_NivelFolder
        'Data Row = 2   -> ID_SubRubro
        'Data Row = 3   -> ID_Folder
        'Data Row = 4   -> ID_Moneda
        
        .AddItem ""
        '.Cell(flexcpText, .Rows - 1, 0) = RetornoNombreFolder(RsAux!SRuNivel, True)
        '.Cell(flexcpText, .Rows - 1, 1) = RetornoNombreDistribucion(RsAux!SRuDistribucion, True)
        
        .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!ComFecha, "dd/mm/yy")
        If Not (IsNull(RsAux("SRuNivel"))) Then
            .Cell(flexcpText, .Rows - 1, 1) = Mid(RetornoNombreFolder(RsAux!SRuNivel, True), 1, 1) & "/" & Mid(RetornoNombreDistribucion(RsAux!SRuDistribucion, True), 1, 1)
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!SRuNombre)
        Else
            MsgBox "OPA"
        End If
        .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ComTC, "0.000")
        
        aValor = RsAux!GImIDCompra: .Cell(flexcpData, .Rows - 1, 0) = aValor        'ID Compra
        aValor = RsAux!GImNivelFolder: .Cell(flexcpData, .Rows - 1, 1) = aValor      'ID Nivel Folder
        aValor = RsAux!SRuId: .Cell(flexcpData, .Rows - 1, 2) = aValor                   'ID Subrubro
        aValor = RsAux!GImFolder: .Cell(flexcpData, .Rows - 1, 3) = aValor             'ID Folder
        aValor = RsAux!ComMoneda: .Cell(flexcpData, .Rows - 1, 4) = aValor          'ID Moneda
        
        If RsAux!ComMoneda = paMonedaPesos Then     'Gasto en pesos
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!GImImporte, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!GImCostear, FormatoMonedaP)
            
            'Formato para las diferencias de cambio "DC" & Format(tFecha.Text, "yyyymm")
            pblnEsDifCambio = Not IsNull(RsAux!ComDCDe)
            If Not pblnEsDifCambio And Not IsNull(RsAux!ComNumero) Then
                pblnEsDifCambio = (RsAux!ComNumero Like "DC*[0-9][0-9][0-9][0-9][0-9][0-9]")
            End If
            
            If Not pblnEsDifCambio Then
                Dim pdlbTC As Double
                pdlbTC = IIf(RsAux!ComTC = 1, RsAux!TipoCambio, RsAux!ComTC)
                .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!GImImporte / pdlbTC, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!GImCostear / pdlbTC, FormatoMonedaP)
            Else
                .Cell(flexcpText, .Rows - 1, 5) = "0.00": .Cell(flexcpText, .Rows - 1, 7) = "0.00"
                
            End If
        Else                                                                'Gasto en dolares
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!GImImporte, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!GImImporte * RsAux!ComTC, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!GImCostear, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!GImCostear * RsAux!ComTC, FormatoMonedaP)
        End If
        .Cell(flexcpForeColor, .Rows - 1, 6, , 7) = &H800000
        
        'Voy a sumar el C&F en pesos
        If RsAux!GImIDSubrubro = paSubrubroDivisa Or RsAux!GImIDSubrubro = paSubrubroTransporteT Or RsAux!GImIDSubrubro = paSubrubroTransporteM Then
            aCF = aCF + CCur(.Cell(flexcpText, .Rows - 1, 6))
        End If
        
        'Cargo porcentaje afectado-----------------------------------------------------------------------------------
        If aTipoFolder = Folder.cFEmbarque Then          'Folder seleccionado es EMBARQUE --> Gastos Carpeta y Embarque
            
            Select Case RsAux!SRuNivel
                Case Folder.cFEmbarque, Folder.cFSubCarpeta: .Cell(flexcpText, .Rows - 1, 8) = "100"
                Case Folder.cFCarpeta
                    Select Case RsAux!SRuDistribucion
                    
                        Case Distribucion.Divisa:
                            If aPrecioTotal = 0 Then
                                .Cell(flexcpText, .Rows - 1, 8) = 0
                            Else
                                .Cell(flexcpText, .Rows - 1, 8) = Format((aPrecioTotalE * 100 / aPrecioTotal), "##,##0.000")
                            End If
                        Case Distribucion.Lineal: .Cell(flexcpText, .Rows - 1, 8) = Format((aCantidadTotalE * 100 / aCantidadTotal), "##,##0.000")
                        Case Distribucion.Volumen: .Cell(flexcpText, .Rows - 1, 8) = Format((aVolumenTotalE * 100 / aVolumenTotal), "##,##0.000")
                    End Select
            End Select
            
        Else                        'Folder seleccionado es SUBCARPETA --> Gastos Carpeta, Embarque y Sub
        
            Select Case RsAux!SRuNivel
                Case Folder.cFSubCarpeta: .Cell(flexcpText, .Rows - 1, 8) = "100"
                
                Case Folder.cFEmbarque
                   Select Case RsAux!SRuDistribucion
                       Case Distribucion.Divisa: .Cell(flexcpText, .Rows - 1, 8) = Format((aPrecioTotalS * 100 / aPrecioTotalE), "##,##0.000")
                       Case Distribucion.Lineal: .Cell(flexcpText, .Rows - 1, 8) = Format((aCantidadTotalS * 100 / aCantidadTotalE), "##,##0.000")
                       Case Distribucion.Volumen: .Cell(flexcpText, .Rows - 1, 8) = Format((aVolumenTotalS * 100 / aVolumenTotalE), "##,##0.000")
                   End Select
                   
                Case Folder.cFCarpeta
                    Select Case RsAux!SRuDistribucion
                        Case Distribucion.Divisa:  .Cell(flexcpText, .Rows - 1, 8) = Format((aPrecioTotalS * 100 / aPrecioTotal), "##,##0.000")
                        Case Distribucion.Lineal: .Cell(flexcpText, .Rows - 1, 8) = Format((aCantidadTotalS * 100 / aCantidadTotal), "##,##0.000")
                        Case Distribucion.Volumen: .Cell(flexcpText, .Rows - 1, 8) = Format((aVolumenTotalS * 100 / aVolumenTotal), "##,##0.000")
                    End Select
            End Select
        End If
        If .Cell(flexcpText, .Rows - 1, 8) = "100.000" Then .Cell(flexcpText, .Rows - 1, 8) = "100"
        .Cell(flexcpForeColor, .Rows - 1, 8) = &H80&
        
        'Cargo cantidades por la DISTRIBUCION -----------------------------------------------------------------------------------
        .Cell(flexcpText, .Rows - 1, 10) = Format(.Cell(flexcpValue, .Rows - 1, 6) * (.Cell(flexcpValue, .Rows - 1, 8) / 100), FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 11) = Format(.Cell(flexcpValue, .Rows - 1, 7) * (.Cell(flexcpValue, .Rows - 1, 8) / 100), FormatoMonedaP)
        .Cell(flexcpBackColor, .Rows - 1, 10, , 11) = Colores.Obligatorio
        
        Select Case RsAux!SRuDistribucion
            Case Distribucion.Divisa
                aDivisaP = aDivisaP + .Cell(flexcpValue, .Rows - 1, 10) '.Cell(flexcpValue, .Rows - 1, 6) * (.Cell(flexcpValue, .Rows - 1, 8) / 100)
                aDivisaD = aDivisaD + .Cell(flexcpValue, .Rows - 1, 11)  '.Cell(flexcpValue, .Rows - 1, 7) * (.Cell(flexcpValue, .Rows - 1, 8) / 100)
            Case Distribucion.Lineal
                aLinealP = aLinealP + .Cell(flexcpValue, .Rows - 1, 10)  '.Cell(flexcpValue, .Rows - 1, 6) * (.Cell(flexcpValue, .Rows - 1, 8) / 100)
                aLinealD = aLinealD + .Cell(flexcpValue, .Rows - 1, 11)  '.Cell(flexcpValue, .Rows - 1, 7) * (.Cell(flexcpValue, .Rows - 1, 8) / 100)
            Case Distribucion.Volumen
                aVolumenP = aVolumenP + .Cell(flexcpValue, .Rows - 1, 10) '.Cell(flexcpValue, .Rows - 1, 6) * (.Cell(flexcpValue, .Rows - 1, 8) / 100)
                aVolumenD = aVolumenD + .Cell(flexcpValue, .Rows - 1, 11) '.Cell(flexcpValue, .Rows - 1, 7) * (.Cell(flexcpValue, .Rows - 1, 8) / 100)
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
        
    'Recorro para asignar los porcentajes de gastos por C&F
    If aCF <> 0 Then
        For I = 1 To .Rows - 1
            Select Case .Cell(flexcpData, I, 2)
                Case paSubrubroDivisa, paSubrubroTransporteT, paSubrubroTransporteM
                Case Else: .Cell(flexcpText, I, 9) = Format((.Cell(flexcpText, I, 6) * 100) / aCF, FormatoMonedaP)
            End Select
        Next
    End If
    
    If vsGasto.Rows > 1 Then
        MarcoGastosDobles
        
        'Ordeno por Tipo de Gasto
        .ColDataType(0) = flexDTDate
        .Select 1, 0, .Rows - 1, 0
        .Sort = flexSortGenericAscending
        '.Sort = flexSortUseColSort
        .Select 0, 0, 0, 0
        
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 6, , &H80&, vbWhite, , "Total"
        .Subtotal flexSTSum, -1, 7
        .Subtotal flexSTSum, -1, 10: .Subtotal flexSTSum, -1, 11
    End If
    End With

    'Cargo las grillas totalizadoras de la distribucion
    With vsTotales
        .Rows = 1
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = RetornoNombreDistribucion(Distribucion.Divisa)
        .Cell(flexcpText, .Rows - 1, 1) = Format(aDivisaP, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 2) = Format(aDivisaD, FormatoMonedaP)
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = RetornoNombreDistribucion(Distribucion.Lineal)
        .Cell(flexcpText, .Rows - 1, 1) = Format(aLinealP, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 2) = Format(aLinealD, FormatoMonedaP)
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = RetornoNombreDistribucion(Distribucion.Volumen)
        .Cell(flexcpText, .Rows - 1, 1) = Format(aVolumenP, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 2) = Format(aVolumenD, FormatoMonedaP)
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = "Total de Gastos"
        .Cell(flexcpText, .Rows - 1, 1) = Format(aDivisaP + aLinealP + aVolumenP, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 2) = Format(aDivisaD + aLinealD + aVolumenD, FormatoMonedaP)
        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = &H80&: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = vbWhite
    End With

    '14/02/2003 - Si el precio total del embarque es 0, hago distribucion lineal de los Gastos por Divisa   ------------------
    Select Case aTipoFolder
        Case Folder.cFEmbarque
                If aPrecioTotalE <> 0 Then
                    CargoDistribucionEmbarque aDivisaP, aDivisaD, aLinealP, aLinealD, aVolumenP, aVolumenD
                Else
                    CargoDistribucionEmbarque 0, 0, aLinealP + aDivisaP, aLinealD + aDivisaD, aVolumenP, aVolumenD
                End If
                
        Case Folder.cFSubCarpeta
                If aPrecioTotalS <> 0 Then
                    CargoDistribucionSub aDivisaP, aDivisaD, aLinealP, aLinealD, aVolumenP, aVolumenD
                Else
                    CargoDistribucionSub 0, 0, aLinealP + aDivisaP, aLinealD + aDivisaD, aVolumenP, aVolumenD
                End If
    End Select
    
    Exit Sub

errGastos:
    clsGeneral.OcurrioError "Error al cargar los gastos asignados a la carpeta.", Err.Description
End Sub

Private Sub MarcoGastosDobles()
    On Error GoTo errMarco
    Dim f1 As Integer
    
    Cons = "Select * from SubRubro Where SRuRubro = " & paRubroImportaciones & _
               " And SRuCantidad = 1" & " And SRuID <> " & paSubrubroDivisa
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        aValor = RsAux!SRuId
        f1 = 0
        With vsGasto
            For I = 1 To .Rows - 1
                If aValor = .Cell(flexcpData, I, 2) Then
                    If f1 = 0 Then
                        f1 = I
                    Else
                        .Cell(flexcpBackColor, f1, 0, , 2) = vbRed: .Cell(flexcpForeColor, f1, 0, , 2) = vbWhite
                        .Cell(flexcpBackColor, I, 0, , 2) = vbRed: .Cell(flexcpForeColor, I, 0, , 2) = vbWhite
                    End If
                End If
            Next
        End With
    
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Sub
errMarco:
    clsGeneral.OcurrioError "Error al marcar los gastos repetidos.", Err.Description
End Sub

Private Sub CargoDistribucionEmbarque(DivisaP As Currency, DivisaD As Currency, LinealP As Currency, LinealD As Currency, VolumenP As Currency, VolumenD As Currency)

Dim aCantidad As Long, aCalculoP As Currency, aCalculoD As Currency
Dim aPUnitario As Currency, aVUnitario As Currency

    On Error GoTo errDis
    With vsCosteo
        .MergeCells = flexMergeSpill
        .Rows = 1
        For I = 1 To vsArticuloE.Rows - 2
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = "(" & vsArticuloE.Cell(flexcpText, I, 0) & ") " & vsArticuloE.Cell(flexcpText, I, 1)
            aValor = vsArticuloE.Cell(flexcpData, I, 0): .Cell(flexcpData, .Rows - 1, 0) = aValor
                       
            aCantidad = vsArticuloE.Cell(flexcpText, I, 2)
            aPUnitario = vsArticuloE.Cell(flexcpText, I, 3)
            aVUnitario = vsArticuloE.Cell(flexcpText, I, 5)
        
            'Distribucion por Divisa
            aCalculoP = 0: aCalculoD = 0
            If DivisaP <> 0 Then
                aCalculoP = ((aPUnitario * 100) / aPrecioTotalE) * DivisaP / 100
                aCalculoD = ((aPUnitario * 100) / aPrecioTotalE) * DivisaD / 100
            End If
            .Cell(flexcpText, .Rows - 1, 3) = Format(aCalculoP, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 4) = Format(aCalculoD, FormatoMonedaP)
            
            'Lineal: Dividido por la cantidad de articulos --- % de la cantidad Total / Gastos Lineal
            aCalculoP = 0: aCalculoD = 0
            If LinealP <> 0 Then
                aCalculoP = ((aCantidad * 100) / aCantidadTotalE) * LinealP / 100
                aCalculoP = aCalculoP / aCantidad
                aCalculoD = ((aCantidad * 100) / aCantidadTotalE) * LinealD / 100
                aCalculoD = aCalculoD / aCantidad
            End If
            .Cell(flexcpText, .Rows - 1, 5) = Format(aCalculoP, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(aCalculoD, FormatoMonedaP)
        
            'Por Volumen --> % de la volumen Total / Gastos Volumen
            aCalculoP = 0: aCalculoD = 0
            If VolumenP <> 0 Then
                aCalculoP = ((aVUnitario * 100) / aVolumenTotalE) * VolumenP / 100
                aCalculoD = ((aVUnitario * 100) / aVolumenTotalE) * VolumenD / 100
            End If
            .Cell(flexcpText, .Rows - 1, 7) = Format(aCalculoP, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(aCalculoD, FormatoMonedaP)
        
            'Gastos Pesos y Dolares general
            aCalculoP = .Cell(flexcpValue, .Rows - 1, 3) + .Cell(flexcpValue, .Rows - 1, 5) + .Cell(flexcpValue, .Rows - 1, 7)
            .Cell(flexcpText, .Rows - 1, 1) = Format(aCalculoP, FormatoMonedaP)
            aCalculoD = .Cell(flexcpValue, .Rows - 1, 4) + .Cell(flexcpValue, .Rows - 1, 6) + .Cell(flexcpValue, .Rows - 1, 8)
            .Cell(flexcpText, .Rows - 1, 2) = Format(aCalculoD, FormatoMonedaP)
    Next
    If vsCosteo.Rows > vsCosteo.FixedRows Then
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, 2) = &H80&: .Cell(flexcpForeColor, 1, 1, .Rows - 1, 2) = vbWhite
    End If
    End With
    
    Exit Sub
errDis:
    clsGeneral.OcurrioError "Error al realizar la distribución de gastos para el embarque.", Err.Description
End Sub

Private Sub CargoDistribucionSub(DivisaP As Currency, DivisaD As Currency, LinealP As Currency, LinealD As Currency, VolumenP As Currency, VolumenD As Currency)

Dim aCantidad As Long, aCalculoP As Currency, aCalculoD As Currency
Dim aPUnitario As Currency, aVUnitario As Currency

    On Error GoTo errDis
    With vsCosteo
        .Rows = 1
        .MergeCells = flexMergeSpill
        For I = 1 To vsArticuloS.Rows - 2
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = "(" & vsArticuloS.Cell(flexcpText, I, 0) & ") " & vsArticuloS.Cell(flexcpText, I, 1)
            aValor = vsArticuloS.Cell(flexcpData, I, 0): .Cell(flexcpData, .Rows - 1, 0) = aValor
                       
            aCantidad = vsArticuloS.Cell(flexcpText, I, 2)
            aPUnitario = vsArticuloS.Cell(flexcpText, I, 3)
            aVUnitario = vsArticuloS.Cell(flexcpText, I, 5)
        
            'Distribucion por Divisa
            aCalculoP = 0: aCalculoD = 0
            If DivisaP <> 0 Then
                aCalculoP = ((aPUnitario * 100) / aPrecioTotalS) * DivisaP / 100
                aCalculoD = ((aPUnitario * 100) / aPrecioTotalS) * DivisaD / 100
            End If
            .Cell(flexcpText, .Rows - 1, 3) = Format(aCalculoP, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 4) = Format(aCalculoD, FormatoMonedaP)
            
            'Lineal: Dividido por la cantidad de articulos --- % de la cantidad Total / Gastos Lineal
            aCalculoP = 0: aCalculoD = 0
            If LinealP <> 0 Then
                aCalculoP = ((aCantidad * 100) / aCantidadTotalS) * LinealP / 100
                aCalculoP = aCalculoP / aCantidad
                aCalculoD = ((aCantidad * 100) / aCantidadTotalS) * LinealD / 100
                aCalculoD = aCalculoD / aCantidad
            End If
            .Cell(flexcpText, .Rows - 1, 5) = Format(aCalculoP, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(aCalculoD, FormatoMonedaP)
        
            'Por Volumen --> % de la volumen Total / Gastos Volumen
            aCalculoP = 0: aCalculoD = 0
            If VolumenP <> 0 Then
                aCalculoP = ((aVUnitario * 100) / aVolumenTotalS) * VolumenP / 100
                aCalculoD = ((aVUnitario * 100) / aVolumenTotalS) * VolumenD / 100
            End If
            .Cell(flexcpText, .Rows - 1, 7) = Format(aCalculoP, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(aCalculoD, FormatoMonedaP)
        
            'Gastos Pesos y Dolares general
            aCalculoP = .Cell(flexcpValue, .Rows - 1, 3) + .Cell(flexcpValue, .Rows - 1, 5) + .Cell(flexcpValue, .Rows - 1, 7)
            .Cell(flexcpText, .Rows - 1, 1) = Format(aCalculoP, FormatoMonedaP)
            aCalculoD = .Cell(flexcpValue, .Rows - 1, 4) + .Cell(flexcpValue, .Rows - 1, 6) + .Cell(flexcpValue, .Rows - 1, 8)
            .Cell(flexcpText, .Rows - 1, 2) = Format(aCalculoD, FormatoMonedaP)
    Next
    .Cell(flexcpBackColor, 1, 1, .Rows - 1, 2) = &H80&: .Cell(flexcpForeColor, 1, 1, .Rows - 1, 2) = vbWhite
    End With
    
    Exit Sub
errDis:
    clsGeneral.OcurrioError "Error al realizar la distribución de gastos para la subcarpeta.", Err.Description
End Sub

Private Sub AccionCostear()

Dim aCantidadS As Long, aCantidadE As Long, aCantidadC As Long
Dim bCarpeta As Boolean, bEmbarque As Boolean, bSub As Boolean
Dim aCantidadCosteada As Long
Dim aTexto As String

    If vsArticuloE.Rows = 1 Then Exit Sub
    
    If prmSRMerImpARecibir = 0 Or prmTipoCompAsiento = 0 Then
        MsgBox "No se cargaron los parámetros para el ingreso del comprobante en ZUREO." & vbCrLf & vbCrLf & "No se puede continuar.", vbCritical, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo ErrorCosteo
    Screen.MousePointer = 11
    
    'Valido se se calcularon las diferencias de cambio del mes anterior a la fecha de arribo del embarque.
    If Not IsDate(lArribo.Caption) Then
        If MsgBox("Esta carpeta no tiene fecha de arribo al local." & vbCrLf & _
                    "El sistema no puede verificar si se calcularon las diferencias de cambio." & vbCrLf & vbCrLf & _
                    "Desea continuar.", vbQuestion + vbYesNo, "Diferencias de Cambio") = vbNo Then Screen.MousePointer = 0: Exit Sub
    Else
        Dim miFechaDC As String, bHayDC As Boolean
        miFechaDC = Format(DateAdd("m", -1, CDate(lArribo.Caption)), "yyyymm")
        Cons = "Select Top 1 * from Compra " & _
              " Where ComNumero = 'DC" & miFechaDC & "' OR ComNumero = 'DC " & miFechaDC & "'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then bHayDC = False Else bHayDC = True
        RsAux.Close
        
        If Not bHayDC Then
            If MsgBox(Trim(miConexion.UsuarioLogueado(Nombre:=True)) & _
                        ": Faltan calcular las diferencias de cambio para el/los meses anteriores al arribo de la mercadería." & vbCrLf & vbCrLf & _
                        "Ud. debe registrarlas antes de costear la carpeta." & vbCrLf & _
                        "Desea continuar.", vbExclamation + vbYesNo + vbDefaultButton2, "Faltan Diferencias de Cambio") = vbNo Then Screen.MousePointer = 0: Exit Sub
        End If
    End If
    '----------------------------------------------------------------------------------------------------------------------------------
    
    bCarpeta = False: bEmbarque = False: bSub = False
    aTexto = ""
    'Debo Validar los gastos obligatorios-----------------------------------------------------------------
    Cons = "Select * from SubRubro Where SRuRubro = " & paRubroImportaciones & _
               " And SRuCosteo = 1 Order by SRuNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    'Data Row = 2   -> ID_SubRubro
    Dim bHayG As Boolean
    Do While Not RsAux.EOF
        bHayG = False
        With vsGasto
            For I = 1 To .Rows - 2      'por el total
                If .Cell(flexcpData, I, 2) = RsAux!SRuId Then bHayG = True: Exit For
            Next
        End With
        
        If Not bHayG Then aTexto = aTexto & Trim(RsAux!SRuNombre) & vbCrLf
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    frmMsgGastos.prmIdCarpeta = aCarpeta
    frmMsgGastos.prmTexto = Trim(aTexto)
    frmMsgGastos.Show vbModal, Me
    Me.Refresh
    If Not frmMsgGastos.prmOK Then Screen.MousePointer = 0: Exit Sub
    '----------------------------------------------------------------------------------------------------------
    
    'Saco cantidad articulos del Embarque
    Cons = "Select Sum(AFoCantidad) from ArticuloFolder" _
           & " Where AFoTipo = " & Folder.cFEmbarque _
           & " And AFoCodigo = " & aEmbarque
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aCantidadE = RsAux(0)
    RsAux.Close
    'Saco cantidad de la Carpeta
    Cons = "Select Sum(AFoCantidad) from ArticuloFolder" _
            & " Where AFoTipo = " & Folder.cFEmbarque _
            & " And AFoCodigo IN (Select EmbId from Embarque Where EmbCarpeta = " & aCarpeta & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aCantidadC = RsAux(0)
    RsAux.Close
    
    If aTipoFolder = Folder.cFEmbarque Then
        If aCantidadE = aCantidadC Then
            aTexto = "El sistema ha detectado que para la carpeta seleccionada hay un solo embarque." & Chr(vbKeyReturn) _
                      & "Se procederá a costear el embarque y la carpeta."
            bCarpeta = True: bEmbarque = True
            
        Else
            'Saco cantidad de articulos costeado para los otros embarques
            Cons = " Select Sum(AFoCantidad) from ArticuloFolder, Embarque" _
                    & " Where AFoTipo = " & Folder.cFEmbarque _
                    & " And AFoCodigo = EmbID" _
                    & " And EmbCarpeta = " & aCarpeta _
                    & " And AFoCodigo <> " & aEmbarque _
                    & " And EmbCosteado = 1"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If IsNull(RsAux(0)) Then aCantidadCosteada = 0 Else aCantidadCosteada = RsAux(0)
            RsAux.Close
            
            If aCantidadC = (aCantidadE + aCantidadCosteada) Then
                aTexto = "El sistema ha detectado que para la carpeta seleccionada todos los embarques han sido costeados." & Chr(vbKeyReturn) _
                           & "Se procederá a costear el embarque y la carpeta."
                bCarpeta = True: bEmbarque = True
            Else
                aTexto = "El sistema ha detectado que para la carpeta seleccionada hay embarques que no han sido costeados." & Chr(vbKeyReturn) _
                          & "Se procederá a costear el embarque."
                bEmbarque = True
            End If
        End If
    
    Else        'ES Subcarpeta
        'Si la cantidad de articulos de la Sub es igual a la del Embarque costeo el Embarque
        'Saco cantidad articulos de la Sub
        Cons = "Select Sum(AFoCantidad) from ArticuloFolder" _
               & " Where AFoTipo = " & Folder.cFSubCarpeta _
               & " And AFoCodigo = " & aSub
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        aCantidadS = RsAux(0)
        RsAux.Close
                
        If aCantidadS = aCantidadE Then
            If aCantidadC = aCantidadE Then
                'Cantidades iguales en Emb, Carp y Sub
                aTexto = "El sistema ha detectado que para la carpeta seleccionada hay un único embarque y una única subcarpeta." & Chr(vbKeyReturn) _
                           & "Se procederá a costear la subcarpeta, el embarque y la carpeta."
                bCarpeta = True: bEmbarque = True: bSub = True
            Else
                'Saco cantidad de articulos costeado para los otros embarques
                Cons = " Select Sum(AFoCantidad) from ArticuloFolder, Embarque" _
                        & " Where AFoTipo = " & Folder.cFEmbarque _
                        & " And AFoCodigo = EmbID" _
                        & " And EmbCarpeta = " & aCarpeta _
                        & " And AFoCodigo <> " & aEmbarque _
                        & " And EmbCosteado = 1"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If IsNull(RsAux(0)) Then aCantidadCosteada = 0 Else aCantidadCosteada = RsAux(0)
                RsAux.Close
                
                If aCantidadC = (aCantidadS + aCantidadCosteada) Then
                    aTexto = "El sistema ha detectado que para la carpeta seleccionada todos los demás embarques han sido costeados." & Chr(vbKeyReturn) _
                              & "Se procederá a costear la subcarpeta, el embarque y la carpeta."
                    bCarpeta = True: bEmbarque = True: bSub = True
                Else
                    aTexto = "El sistema ha detectado que para la carpeta seleccionada hay embarques que no han sido costeados." & Chr(vbKeyReturn) _
                              & "Se procederá a costear la subcarpeta y el embarque."
                    bEmbarque = True: bSub = True
                End If
            End If
            
        Else            'La cantidad de la Sub <> Cantidad Embarque
            'Saco cantidad costeada para las otras subcarpetas del embarque
            Cons = " Select Sum(AFoCantidad) from ArticuloFolder, Subcarpeta" _
                        & " Where AFoTipo = " & Folder.cFSubCarpeta _
                        & " And AFoCodigo = SubID" _
                        & " And SubEmbarque = " & aEmbarque _
                        & " And AFoCodigo <> " & aSub _
                        & " And SubCosteada = 1"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If IsNull(RsAux(0)) Then aCantidadCosteada = 0 Else aCantidadCosteada = RsAux(0)
            RsAux.Close
            
            If aCantidadE = (aCantidadS + aCantidadCosteada) Then
                'Toodo el embarque se costeó
                'Saco lo costeado para los demas embarques
                Cons = " Select Sum(AFoCantidad) from ArticuloFolder, Embarque" _
                        & " Where AFoTipo = " & Folder.cFEmbarque _
                        & " And AFoCodigo = EmbID" _
                        & " And EmbCarpeta = " & aCarpeta _
                        & " And AFoCodigo <> " & aEmbarque _
                        & " And EmbCosteado = 1"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If IsNull(RsAux(0)) Then aCantidadCosteada = 0 Else aCantidadCosteada = RsAux(0)
                RsAux.Close
                If aCantidadC = (aCantidadE + aCantidadCosteada) Then
                    'Se costeo TODO costeo Todo
                    aTexto = "El sistema ha detectado que para la carpeta seleccionada todos los demás embarques han sido costeados." & Chr(vbKeyReturn) _
                               & "Se procederá a costear la subcarpeta, el embarque y la carpeta."
                    bCarpeta = True: bEmbarque = True: bSub = True
                Else
                    aTexto = "El sistema ha detectado que para la carpeta seleccionada hay embarques que no han sido costeados." & Chr(vbKeyReturn) _
                             & "Se procederá a costear la subcarpeta y el embarque."
                     bEmbarque = True: bSub = True
                End If
            Else
                'No se costeo todo el embarque
                    aTexto = "El sistema ha detectado que para la carpeta seleccionada hay subcarpetas pertenecientes al embarque  que no han sido costeadas." & Chr(vbKeyReturn) _
                            & "Se procederá a costear la subcarpeta."
                    bSub = True
            End If
        End If
    End If
    Screen.MousePointer = 0
    If Not fnc_ValidoAcceso() Then
        MsgBox "Falló la conexión a ZUREO, no podrá continuar.", vbCritical, "Sin acceso a Zureo"
        Exit Sub
    End If
    
    If MsgBox(aTexto, vbOKCancel + vbDefaultButton2 + vbQuestion, "Costear Carpeta") = vbCancel Then Exit Sub
    GraboCosteo bCarpeta, bEmbarque, bSub
    Exit Sub
    
ErrorCosteo:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Se ha producido un error al realizar la consulta para el costeo. ", Err.Description
End Sub

Private Sub GraboCosteo(Carpeta As Boolean, Embarque As Boolean, SubCarpeta As Boolean)
Dim aIDCosteo As Long
Dim J As Integer
Dim mStrErr As String

    Screen.MousePointer = 11
    On Error GoTo errorBT
    FechaDelServidor
    Dim colGastos As New Collection
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    If SubCarpeta Then
        Cons = "Update SubCarpeta Set SubCosteada = 1 Where SubID = " & aSub
        cBase.Execute Cons
    End If
    If Embarque Then
        Cons = "Update Embarque Set EmbCosteado = 1 Where EmbID = " & aEmbarque
        cBase.Execute Cons
    End If
    If Carpeta Then
        Cons = "Update Carpeta Set CarCosteada = 1 Where CarID = " & aCarpeta
        cBase.Execute Cons
    End If
    
    'Inserto en Tabla CosteoCarpeta------------------------------------------------------------------------------
    Cons = "Select * from CosteoCarpeta Where CCaID = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!CCaNivelFolder = aTipoFolder
    RsAux!CCaFolder = aFolder
    RsAux!CCaCarpeta = Trim(cFolder.Text)
    RsAux!CCaFCosteo = Format(gFechaServidor, sqlFormatoFH)
    If IsDate(lArribo.Caption) Then RsAux!CCaFArribo = Format(lArribo.Caption, sqlFormatoFH)
    RsAux!CCaFApertura = Format(lApertura.Caption, sqlFormatoFH)
    RsAux!CCaUsuario = paCodigoDeUsuario
    RsAux.Update
    RsAux.Close
    
    'Saco el ID de Costeo
    Cons = "Select Max(CCaID) from CosteoCarpeta"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aIDCosteo = RsAux(0)
    RsAux.Close
    
    'Actualizo el monto para cada gasto de la lista
    'Data Row = 0   -> ID_Compra
    'Data Row = 1   -> ID_NivelFolder
    'Data Row = 2   -> ID_SubRubro
    'Data Row = 3   -> ID_Folder
    'Data Row = 4   -> ID_Moneda
    
    Dim RsGas As rdoResultset
    Cons = "Select * from CosteoGasto Where CGaIDCosteo = " & aIDCosteo
    Set RsGas = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    For I = 1 To vsGasto.Rows - 2       'por el total
        mStrErr = "GastoImportacion: " & vsGasto.Cell(flexcpData, I, 0)
        
        'Actualizo el importe del gasto en la tabla Gasto importacion----------------------------------------------
        Cons = "Select * from GastoImportacion" _
                & " Where GImIDCompra = " & vsGasto.Cell(flexcpData, I, 0) _
                & " And GImIDSubrubro = " & vsGasto.Cell(flexcpData, I, 2) _
                & " And GImNivelFolder = " & vsGasto.Cell(flexcpData, I, 1) _
                & " And GImFolder = " & vsGasto.Cell(flexcpData, I, 3)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        If vsGasto.Cell(flexcpData, I, 4) = paMonedaPesos Then     'Pesos
            RsAux!GImCostear = RsAux!GImCostear - (vsGasto.Cell(flexcpValue, I, 6) * (vsGasto.Cell(flexcpValue, I, 8) / 100))
        Else                                                                               'Dolares
            RsAux!GImCostear = RsAux!GImCostear - (vsGasto.Cell(flexcpValue, I, 7) * (vsGasto.Cell(flexcpValue, I, 8) / 100))
        End If
        RsAux.Update: RsAux.Close
        
        'Inserto los datos en la tabla CosteoGasto
        RsGas.AddNew
        RsGas!CGaIDCosteo = aIDCosteo
        RsGas!CGaIDCompra = vsGasto.Cell(flexcpData, I, 0)
        RsGas!CGaIDGasto = vsGasto.Cell(flexcpData, I, 2)
        RsGas!CGaNivelFolder = vsGasto.Cell(flexcpData, I, 1)
        RsGas!CGaFolder = vsGasto.Cell(flexcpData, I, 3)
        RsGas!CGaImporteOP = vsGasto.Cell(flexcpValue, I, 4)
        RsGas!CGaImporteOD = vsGasto.Cell(flexcpValue, I, 5)
        RsGas!CGaImporteAP = vsGasto.Cell(flexcpValue, I, 6) * (vsGasto.Cell(flexcpValue, I, 8) / 100)
        RsGas!CGaImporteAD = vsGasto.Cell(flexcpValue, I, 7) * (vsGasto.Cell(flexcpValue, I, 8) / 100)
        RsGas.Update
        
        Dim oGasto As clsGastoTotal
        Set oGasto = BuscoGastoEnColeccion(colGastos, vsGasto.Cell(flexcpData, I, 2))
        If oGasto Is Nothing Then
            Set oGasto = New clsGastoTotal
            colGastos.Add oGasto
            oGasto.idSubro = vsGasto.Cell(flexcpData, I, 2)
        End If
        oGasto.TotalMonComprobante = Format(CCur(vsGasto.Cell(flexcpValue, I, 6) * (vsGasto.Cell(flexcpValue, I, 8) / 100)) + oGasto.TotalMonComprobante, "#,##0.00")
        If vsGasto.Cell(flexcpData, I, 4) = paMonedaPesos Then     'Pesos
            oGasto.TotalMonCuenta = Format(CCur(vsGasto.Cell(flexcpValue, I, 6) * (vsGasto.Cell(flexcpValue, I, 8) / 100)) + oGasto.TotalMonCuenta, "#,##0.00")
        Else
            oGasto.TotalMonCuenta = Format(CCur(vsGasto.Cell(flexcpValue, I, 7) * (vsGasto.Cell(flexcpValue, I, 8) / 100)) + oGasto.TotalMonCuenta, "#,##0.00")
        End If
    Next
    RsGas.Close
    
    'Guardo informacion para cada articulo Tabla CosteoArticulo
    Cons = "Select * from CosteoArticulo Where CArIDCosteo = " & aIDCosteo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    For I = 1 To vsCosteo.Rows - 1
        mStrErr = "CosteoArticulo: " & vsCosteo.Cell(flexcpData, I, 0)
        
        RsAux.AddNew
        RsAux!CArIDCosteo = aIDCosteo
        RsAux!CArIDArticulo = vsCosteo.Cell(flexcpData, I, 0)
        
        If aTipoFolder = Folder.cFEmbarque Then
            With vsArticuloE
                For J = 1 To .Rows - 2
                    If vsCosteo.Cell(flexcpData, I, 0) = .Cell(flexcpData, J, 0) Then
                        RsAux!CArCantidad = .Cell(flexcpValue, J, 2)
                        RsAux!CArImporteO = .Cell(flexcpValue, J, 3)
                        Exit For
                    End If
                Next
            End With
        Else
            With vsArticuloS
                For J = 1 To .Rows - 2
                    If vsCosteo.Cell(flexcpData, I, 0) = .Cell(flexcpData, J, 0) Then
                        RsAux!CArCantidad = .Cell(flexcpValue, J, 2)
                        RsAux!CArImporteO = .Cell(flexcpValue, J, 3)
                        Exit For
                    End If
                Next
            End With
        End If
        
        RsAux!CArCostoP = vsCosteo.Cell(flexcpValue, I, 1)
        RsAux!CArCostoD = vsCosteo.Cell(flexcpValue, I, 2)
        RsAux.Update
    Next
    RsAux.Close
    
    cBase.CommitTrans    'FIN TRANSACCION------------------------------------------
    
    If colGastos.Count > 0 Then
        Dim aCompZureo As Long
        Dim sFolder As String
        sFolder = Mid(cFolder.Text, InStr(1, cFolder.Text, ".") + 1) & "-" & Mid(cFolder.Text, 1, InStr(1, cFolder.Text, ".") - 1)
        Dim Fecha As Date
        If IsDate(lArribo.Caption) Then
            Fecha = CDate(lArribo.Caption)
        Else
            Fecha = gFechaServidor
        End If
        
        aCompZureo = modGastoZureo.IngresoGastoZureo(Fecha, sFolder, "Costeo de carpeta: " & Trim(cFolder.Text), colGastos)
        If aCompZureo > 0 Then
            MsgBox "Se agregó el comprobante " & aCompZureo & " en Zureo.", vbInformation, "Comprobante en Zureo"
        Else
            MsgBox "NO se insertó el comprobante en Zureo que resume las cuentas del costeo.", vbExclamation, "Comprobante en Zureo"
        End If
    End If
    
    Screen.MousePointer = 0
    If MsgBox("La carpeta ha sido costeada con éxito. Para visualizar la información de costeo ver Carpetas Costeadas." _
            & Chr(vbKeyReturn) & "Desea imprimir una copia de la información de costeo ?.", vbInformation + vbYesNo, "ATENCIÓN") = vbYes Then
            AccionImprimir
    End If
    cFolder.RemoveItem cFolder.ListIndex
    AccionLimpiar
    Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al grabar. " & mStrErr, Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionImprimir()

Dim aTituloTabla As String, aComentario As String

    If vsArticulo.Rows = 1 Or cFolder.ListIndex = -1 Then MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN": Exit Sub
    
    On Error GoTo errPrint
    Screen.MousePointer = 11
    aTituloTabla = "": aComentario = ""
    
    With vsListado
    
        If Not .PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 0: Exit Sub
        
        .Preview = True: .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
    
        EncabezadoListado vsListado, "Importaciones - Costeo de Carpetas", False
        
        .FileName = "Costeo de Carpetas"
        .FontSize = 8: .FontBold = False
        
        .Paragraph = "": .FontSize = 9
        .Text = "Apertura: " & lApertura.Caption & Chr(vbKeyTab) & "Arribo de Mercadería: " & lArribo.Caption & Chr(vbKeyTab) & Chr(vbKeyTab) & Chr(vbKeyTab) & Chr(vbKeyTab) & Chr(vbKeyTab)
        .Paragraph = "": .Paragraph = "": .FontSize = 8: .FontBold = False
        
        'Listas de Articulos------------------------------------------------------------------------
        .Paragraph = "Artículos de la Carpeta Madre"
        vsArticulo.ExtendLastCol = False: .RenderControl = vsArticulo.hwnd: vsArticulo.ExtendLastCol = True
        
        .Paragraph = "": .Paragraph = "Artículos del Embarque"
        vsArticuloE.ExtendLastCol = False: .RenderControl = vsArticuloE.hwnd: vsArticuloE.ExtendLastCol = True
        If vsArticuloS.Rows > 1 Then
            .Paragraph = "": .Paragraph = "Artículos de la Subcarpeta"
            vsArticuloS.ExtendLastCol = False: .RenderControl = vsArticuloS.hwnd: vsArticuloS.ExtendLastCol = True
        End If
        '--------------------------------------------------------------------------------------------
        
        .Paragraph = "": .Paragraph = "Lista de Gastos Asignados"
        vsGasto.FontSize = 7
        vsGasto.ExtendLastCol = False: .RenderControl = vsGasto.hwnd: vsGasto.ExtendLastCol = True
        vsGasto.FontSize = 8
        
        .FontSize = 12: .FontBold = True: .Paragraph = ""
        .TextAlign = taRightTop: .Text = Format(lArribo.Caption, "Mmmm/yyyy"): .TextAlign = taLeftTop
        .FontSize = 8: .FontBold = False
        
        .Paragraph = "": .Paragraph = "Asignación de Gastos según Distribución"
        vsTotales.ExtendLastCol = False: .RenderControl = vsTotales.hwnd: vsTotales.ExtendLastCol = True
        
        .Paragraph = "": .Paragraph = "Costeo de Artículos por Unidad"
        vsCosteo.ExtendLastCol = False: .RenderControl = vsCosteo.hwnd: vsCosteo.ExtendLastCol = True
        
        .EndDoc
        .PrintDoc
        
    End With
    Screen.MousePointer = 0
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión. ", Err.Description
End Sub

Private Sub CargoComboFolder()
     
    On Error GoTo errCargo
    cFolder.Clear
    'Saco los embarques que no tengan subcarpetas  - Q no vayan a ZF
    Cons = "Select Nivel = " & Folder.cFEmbarque & ", ID = EmbID, Carpeta = CarCodigo, Embarque = EmbCodigo" _
            & " From Embarque, Carpeta " _
            & " Where EmbLocal <> " & paLocalZF _
            & " And EmbCosteado = 0" _
            & " And EmbCarpeta = CarID ORDER BY CarCodigo, EmbCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        cFolder.AddItem RsAux!Carpeta & "." & RsAux!Embarque
        cFolder.ItemData(cFolder.NewIndex) = RsAux!Nivel & RsAux!ID
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Saco todas las subcarpetas
    Cons = "Select Nivel = " & Folder.cFSubCarpeta & ", ID = SubID, Carpeta = CarCodigo, Embarque = EmbCodigo, Sub = SubCodigo" _
            & " From SubCarpeta, Embarque, Carpeta " _
            & " Where SubCosteada = 0" _
            & " And SubEmbarque = EmbID" _
            & " And EmbCarpeta = CarID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        cFolder.AddItem RsAux!Carpeta & "." & RsAux!Embarque & "/" & RsAux!Sub
        cFolder.ItemData(cFolder.NewIndex) = RsAux!Nivel & RsAux!ID
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Sub
    
errCargo:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar las carpetas a costear.", Err.Description
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsArticulo
        .Cols = 1:
        .WordWrap = False
        .FormatString = "Código|Artículo|>Q|>$ Unitario|>$ Total|>Vol. Unitario|>Vol. Total|"
            
        .ColWidth(0) = 850: .ColWidth(1) = 2600: .ColWidth(2) = 700: .ColWidth(3) = 1200: .ColWidth(4) = 1200: .ColWidth(5) = 1200: .ColWidth(6) = 1200
        .ColDataType(2) = flexDTCurrency: .ColDataType(3) = flexDTCurrency: .ColDataType(4) = flexDTCurrency: .ColDataType(5) = flexDTCurrency: .ColDataType(6) = flexDTCurrency
        .ColSort(0) = flexSortGenericAscending
    End With
    
    With vsArticuloE
         .Cols = 1: .Rows = 1
        .WordWrap = False
        .FormatString = "Código|Artículo|>Q|>$ Unitario|>$ Total|>Vol. Unitario|>Vol. Total|% Carpeta|"
            
        .ColWidth(0) = 850: .ColWidth(1) = 2600: .ColWidth(2) = 700: .ColWidth(3) = 1200: .ColWidth(4) = 1200: .ColWidth(5) = 1200: .ColWidth(6) = 1200: .ColWidth(7) = 900
        .ColDataType(2) = flexDTCurrency: .ColDataType(3) = flexDTCurrency: .ColDataType(4) = flexDTCurrency: .ColDataType(5) = flexDTCurrency: .ColDataType(6) = flexDTCurrency
        .ColSort(0) = flexSortGenericAscending
    End With
    
    With vsArticuloS
        .Cols = 1: .Rows = 1
        .FormatString = "Código|Artículo|>Q|>$ Unitario|>$ Total|>Vol. Unitario|>Vol. Total|% Emb.|"
        .WordWrap = False
        
        .ColWidth(0) = 850: .ColWidth(1) = 2600: .ColWidth(2) = 700: .ColWidth(3) = 1200: .ColWidth(4) = 1200: .ColWidth(5) = 1200: .ColWidth(6) = 1200: .ColWidth(7) = 900
        .ColDataType(2) = flexDTCurrency: .ColDataType(3) = flexDTCurrency: .ColDataType(4) = flexDTCurrency: .ColDataType(5) = flexDTCurrency: .ColDataType(6) = flexDTCurrency
        .ColSort(0) = flexSortGenericAscending
    End With
    
    'Grillas de Gastos y Distribucion    ---------------------------------------------------------------------------------------------------------------
    With vsGasto
        .Cols = 1 ':.Rows = 1
        '.FormatString = "Nivel|Distrib.|Gasto|TC|>Pesos (G)|>Dólares (G)|>A Costear $|>A Costear U$S|% Gasto|% C&F|>Pesos|>Dólares|"
        .FormatString = "Fecha|Ni/Di|Gasto|TC|>Pesos (G)|>Dólares (G)|>A Costear $|>A Costear U$S|% Gasto|% C&F|>Pesos|>Dólares|"
        .WordWrap = False
        .ColWidth(0) = 750
        .ColWidth(1) = 360
        .ColWidth(2) = 950: .ColWidth(3) = 600: .ColWidth(4) = 1000: .ColWidth(5) = 1000: .ColWidth(6) = 1200: .ColWidth(7) = 1200 ': .ColWidth(8) = 1000
        .ColWidth(10) = 1000: .ColWidth(11) = 1000
        .ColDataType(1) = flexDTCurrency: .ColDataType(2) = flexDTCurrency: .ColDataType(0) = flexDTDate
    End With
    
    With vsTotales
        .Cols = 1 ':.Rows = 1
        .FormatString = "Tipo Distribución|>Pesos|>Dólares|"
        .WordWrap = False
        .ColWidth(0) = 2000: .ColWidth(1) = 1200: .ColWidth(2) = 1200
        
        .ColDataType(1) = flexDTCurrency: .ColDataType(2) = flexDTCurrency
    End With
    
    With vsCosteo
        .Cols = 1 ':.Rows = 1
        .FormatString = "Artículo|>Costo $|>Costo U$S|>Divisa $|>Divisa U$S|>Lineal $|>Lineal U$S|>Volumen $|>Volumen U$S|"
        .WordWrap = False
        .ColWidth(0) = 2500
        .ColWidth(1) = 900: .ColWidth(2) = 900: .ColWidth(3) = 900: .ColWidth(4) = 900: .ColWidth(5) = 900: .ColWidth(6) = 900: .ColWidth(7) = 900
        .ColDataType(1) = flexDTCurrency: .ColDataType(2) = flexDTCurrency
        .ColDataType(3) = flexDTCurrency: .ColDataType(4) = flexDTCurrency: .ColDataType(5) = flexDTCurrency: .ColDataType(6) = flexDTCurrency: .ColDataType(7) = flexDTCurrency: .ColDataType(8) = flexDTCurrency
    End With
    
End Sub


Private Sub MnuBuEmbarques_Click()
On Error GoTo errApp
    Dim RetVal
    Screen.MousePointer = 11: Me.Refresh
    RetVal = Shell(App.Path & "\Consulta embarques ", 1)
    Screen.MousePointer = 0
    Exit Sub
errApp:
    clsGeneral.OcurrioError "Ocurrió un error al ejecutar la aplicación. " & Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuBuGastos_Click()
On Error GoTo errApp
    Dim RetVal
    Screen.MousePointer = 11: Me.Refresh
    RetVal = Shell(App.Path & "\Consulta de Gastos ", 1)
    Screen.MousePointer = 0
    Exit Sub
errApp:
    clsGeneral.OcurrioError "Ocurrió un error al ejecutar la aplicación. " & Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuVerEmbarque_Click()
On Error GoTo errApp
    Dim RetVal
    Screen.MousePointer = 11: Me.Refresh
    RetVal = Shell(App.Path & "\Embarque " & aEmbarque, 1)
    Screen.MousePointer = 0
    Exit Sub
errApp:
    clsGeneral.OcurrioError "Ocurrió un error al ejecutar la aplicación. " & Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuVerGasto_Click()
On Error GoTo errApp
    Dim RetVal
    Screen.MousePointer = 11: Me.Refresh
    RetVal = Shell(App.Path & "\Ingreso de Gastos " & vsGasto.Cell(flexcpData, vsGasto.Row, 0), 1)
    Screen.MousePointer = 0
    Exit Sub
errApp:
    clsGeneral.OcurrioError "Ocurrió un error al ejecutar la aplicación. " & Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tbArticulo_Click()
    
    Select Case tbArticulo.SelectedItem.Index
        Case 1: vsArticulo.ZOrder 0
        Case 2: vsArticuloE.ZOrder 0
        Case 3: vsArticuloS.ZOrder 0
    End Select
    
End Sub

Private Sub tbCosteo_Click()
    Select Case tbCosteo.SelectedItem.Index
        Case 1: vsTotales.ZOrder 0
        Case 2: vsCosteo.ZOrder 0
    End Select
End Sub

Private Sub vsCosteo_DblClick()

    If vsCosteo.Rows = 1 Then Exit Sub
    If Val(vsCosteo.Cell(flexcpData, vsCosteo.Row, 0)) = 0 Then Exit Sub
    
    On Error GoTo errLista
    Screen.MousePointer = 11
    Cons = " Select CCaID, 'Carpeta' = CCaCarpeta, 'Costeada' = CCaFCosteo, 'Arribó' =  CCaFArribo, 'Artículo' =  ArtNombre, 'Costo $' = CArCostoP, 'Costo U$S' = CArCostoD" & _
                " From CosteoArticulo, CosteoCarpeta, Articulo " & _
                " Where CCaID = CArIdCosteo " & _
                " And CArIdArticulo = ArtID" & _
                " And ArtID = " & vsCosteo.Cell(flexcpData, vsCosteo.Row, 0) & _
                " Order by CCaFCosteo DESC"
                
    Dim objLista As New clsListadeAyuda
    objLista.ActivarAyuda cBase, Cons, 8600, 1, "Lista de Costeos"
    'objLista.ActivoListaAyuda Cons, False, txtConeccion, 8600
    Set objLista = Nothing
    Screen.MousePointer = 0
    Exit Sub

errLista:
    clsGeneral.OcurrioError "Ocurrió un error al buscar los costes anteriores.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub vsCosteo_GotFocus()
    Status.Panels(3).Text = "Doble Click- Ver últimos costeos para el artículo seleccionado."
End Sub

Private Sub vsCosteo_LostFocus()
    Status.Panels(3).Text = ""
End Sub

'Private Sub vsGasto_Click()
'    With vsGasto
'        If .MouseRow = 0 Then
'            .ColSel = .MouseCol
'            If .ColSort(.MouseCol) = flexSortGenericAscending Then .ColSort(.MouseCol) = flexSortGenericDescending Else .ColSort(.MouseCol) = flexSortGenericAscending
'            .Sort = flexSortUseColSort
'        End If
'    End With
'End Sub

Private Sub vsGasto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        If vsGasto.Rows > 1 And vsGasto.Cell(flexcpData, vsGasto.MouseRow, 0) <> 0 Then
            MnuVerGasto.Enabled = True: MnuVerEmbarque.Enabled = True
        Else
            MnuVerGasto.Enabled = False: MnuVerEmbarque.Enabled = False
        End If
        vsGasto.Select vsGasto.MouseRow, 0, vsGasto.MouseRow, vsGasto.Cols - 1
        PopupMenu MnuAcceso
    End If
    
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionLimpiar(Optional DejarFolder As Boolean = False)
    lArribo.Caption = "N/D"
    lApertura.Caption = "N/D"
    If Not DejarFolder Then cFolder.Text = ""
    vsArticulo.Rows = 1: vsArticuloE.Rows = 1: vsArticuloS.Rows = 1
    vsGasto.Rows = 1
    vsTotales.Rows = 1: vsCosteo.Rows = 1
End Sub

Private Sub vsListado_NewPage()
    With vsListado
        .FontSize = 12: .FontBold = True
        .TextAlign = taRightTop
        .Text = "Carpeta: " & cFolder.Text
        .TextAlign = taLeftTop
    End With
End Sub
