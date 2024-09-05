VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{191D08B9-4E92-4372-BF17-417911F14390}#1.5#0"; "orGridPreview.ocx"
Object = "{923DD7D8-A030-4239-BCD4-51FDB459E0FE}#4.0#0"; "orgComboCalculator.ocx"
Object = "{D9D9E0F6-C86B-4B3A-BFD9-06B9B5B7A222}#2.1#0"; "orUserDigit.ocx"
Begin VB.Form frmListado 
   Caption         =   "Liquidación de Camioneros"
   ClientHeight    =   6540
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13905
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
   ScaleHeight     =   6540
   ScaleWidth      =   13905
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboOpcion 
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtCodImpresion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      TabIndex        =   40
      Top             =   120
      Width           =   1575
   End
   Begin VB.Timer tmStart 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   4200
   End
   Begin orGridPreview.GridPreview gpPrint 
      Left            =   480
      Top             =   2280
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5115
      Index           =   1
      Left            =   6360
      ScaleHeight     =   5115
      ScaleWidth      =   10275
      TabIndex        =   26
      Top             =   720
      Width           =   10275
      Begin orgCalculatorFlat.orgCalculator caPesos 
         Height          =   285
         Left            =   960
         TabIndex        =   38
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
      End
      Begin VB.TextBox tConAmp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   32
         Top             =   840
         Width           =   2055
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   900
         TabIndex        =   11
         Top             =   2340
         Width           =   690
         _ExtentX        =   1217
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsLiquidacion 
         Height          =   1455
         Left            =   3240
         TabIndex        =   28
         Top             =   3600
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   2566
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
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483643
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   3
         FixedRows       =   0
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
         MergeCells      =   6
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
         Editable        =   -1  'True
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
      Begin VB.TextBox tA 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   900
         TabIndex        =   14
         Top             =   2700
         Width           =   1215
      End
      Begin VB.TextBox tSon 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1620
         TabIndex        =   12
         Top             =   2340
         Width           =   1095
      End
      Begin VB.ComboBox cTipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1980
         Width           =   2115
      End
      Begin AACombo99.AACombo cLiquidar 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   480
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsDetMonto 
         Height          =   1575
         Left            =   3240
         TabIndex        =   7
         Top             =   120
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   2778
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
         ForeColor       =   -2147483640
         BackColorFixed  =   10526720
         ForeColorFixed  =   16777215
         BackColorSel    =   12582912
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483643
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
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
         MergeCells      =   6
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsPagaCon 
         Height          =   1695
         Left            =   3240
         TabIndex        =   35
         Top             =   1800
         Width           =   5355
         _ExtentX        =   9446
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
         BackColorSel    =   -2147483633
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483643
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   3
         FixedRows       =   0
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
         MergeCells      =   6
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
      Begin VB.Label lblAyudaCheque 
         BackStyle       =   0  'Transparent
         Caption         =   "'Enter' alta, 'F1' buscar"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackColor       =   &H00008000&
         Caption         =   " Paga"
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
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label6 
         BackColor       =   &H00008000&
         Caption         =   " Resumen Final de Pago"
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
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ampliación:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00008000&
         Caption         =   " Conceptos de Liquidación (Boleta)"
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
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label lLiquidar 
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidar:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lPesos 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lA 
         BackStyle       =   0  'Transparent
         Caption         =   "&por"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   2700
         Width           =   615
      End
      Begin VB.Label lSon 
         BackStyle       =   0  'Transparent
         Caption         =   "&Son:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2340
         Width           =   675
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Con:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1980
         Width           =   735
      End
   End
   Begin MSComctlLib.TabStrip tsOpcion 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Mo&vimientos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Liquidación"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3075
      Index           =   0
      Left            =   360
      ScaleHeight     =   3075
      ScaleWidth      =   2955
      TabIndex        =   25
      Top             =   840
      Width           =   2955
      Begin VB.PictureBox picHorizontal 
         BackColor       =   &H8000000D&
         Height          =   45
         Left            =   0
         ScaleHeight     =   45
         ScaleWidth      =   2655
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
         Height          =   1875
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3307
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorFixed  =   14857624
         ForeColorFixed  =   0
         BackColorSel    =   12582912
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   4
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
   End
   Begin VB.PictureBox picTextos 
      Height          =   375
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   9435
      TabIndex        =   24
      Top             =   5760
      Width           =   9495
      Begin orUserDigit.UserDigit tUsuario 
         Height          =   285
         Left            =   7320
         TabIndex        =   37
         Top             =   0
         Width           =   1455
         _ExtentX        =   2805
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   16
         Top             =   0
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Co&mentario:"
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "&Usuario:"
         Height          =   255
         Left            =   6660
         TabIndex        =   17
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   375
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   6075
      TabIndex        =   23
      Top             =   6240
      Width           =   6135
      Begin VB.CommandButton bLastHtml 
         Height          =   310
         Left            =   2160
         Picture         =   "frmListado.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Ver última liquidación"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bQuery 
         Height          =   310
         Left            =   600
         Picture         =   "frmListado.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Consultar."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   960
         Picture         =   "frmListado.frx":084E
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   60
         Picture         =   "frmListado.frx":0950
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   1380
         Picture         =   "frmListado.frx":0A52
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bPreview 
         Height          =   310
         Left            =   1740
         Picture         =   "frmListado.frx":0B54
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Preview."
         Top             =   0
         Width           =   310
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   6285
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21431
            Key             =   "msg"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Zureo OK"
            TextSave        =   "Zureo OK"
            Key             =   "zureo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Image1 
      Left            =   8640
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1086
            Key             =   "entrega"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1318
            Key             =   "retiro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1632
            Key             =   "visita"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":194C
            Key             =   "envio"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1C66
            Key             =   "last"
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cCamion 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
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
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar códigos separados con coma (sólo carga envíos del o los códigos ingresados)."
      ForeColor       =   &H00666666&
      Height          =   375
      Left            =   9720
      TabIndex        =   43
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "&Opción:"
      Height          =   255
      Left            =   3120
      TabIndex        =   41
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Impresión"
      Height          =   255
      Left            =   6240
      TabIndex        =   39
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Camión:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu MnuAcceso 
      Caption         =   "Accesos"
      Visible         =   0   'False
      Begin VB.Menu MnuAcEnvio 
         Caption         =   "Ver Envío"
      End
      Begin VB.Menu MnuAcServicio 
         Caption         =   "Seguimiento de Servicio"
      End
      Begin VB.Menu MnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMarcarTodos 
         Caption         =   "Seleccionar Todos"
      End
      Begin VB.Menu MnuDesmarcar 
         Caption         =   "Deseleccionar Todos"
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TipoComprobanteZureo
    SalidaCaja = 30
    EntradaCaja = 31
    DepósitoZureo = 61
End Enum

Private colQueCobro As New Collection
Private Const ct_KeyDisponibilidadCamioneros As String = "Diferencias liquidación camioneros"

Dim oTransRP As clsDesicionConQuePaga
Dim colLiqCobro As Collection

Dim lIDDispPesos As Long
Dim iIDSRPesos As Long

'Dim colReclamos As Collection
'Dim colPagos As Collection

Public paTipoArticuloServicio As Integer
Const cte_CelesteClaro As Long = &HFFE0D0

Private Enum TipoSuceso
    ModificacionDePrecios = 3
End Enum

Private Enum DocPendiente
    Servicio = 2
End Enum

Private Enum TipoPagoEnvio
    PagaAhora = 1
    PagaDomicilio = 2
    FacturaCamión = 3
End Enum

'T = al día y V = $
'AltaCheques.exe T 0|V 650

Private lIDLastSave As Long
Private lLastID As Long

Dim sSignMonedaPesos As String
Private Type tPendiente
    Envio As Long
    Pendiente As Long
    Moneda As Long
    Importe As Currency
    Activo As Boolean
End Type
Private arrPendiente() As tPendiente

Private aTexto As String
Private bSizeAjuste As Boolean

Dim objUsers As clsCheckIn
Private iUserZureo As Long

Private Sub CargarChequesDeTag(ByVal idtag As Integer, ByVal idCheque As Long)
On Error GoTo errTC
Dim rsT As rdoResultset
Dim iUtilizo As Integer
Dim TCME As String
    
    Dim sQueryCont As String
    'Si recibo el ID es por lista de ayuda.
    If idCheque > 0 Then
        sQueryCont = " WHERE CDiCodigo = " & idCheque
    Else
        sQueryCont = " WHERE CDiTag = " & idtag
    End If
    
    iUtilizo = 254
    Cons = "SELECT CDiCodigo, CDiMoneda, CDiImporte, rtrim(CDiSerie) + ' ' + CONVERT(varchar(25), CDiNumero) Cheque, MonSigno " & _
            "FROM ChequeDiferido INNER JOIN Moneda ON CDiMoneda = MonCodigo " & sQueryCont   'WHERE CDiTag = " & idtag
                        
    Set rsT = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsT.EOF Then
        With vsPagaCon
            If .Rows > .FixedRows Then
                If .IsSubtotal(.FixedRows) Then .RemoveItem .FixedRows
            End If
        End With
    End If
    
    Do While Not rsT.EOF
        
        With vsPagaCon
            
            If rsT("CDiMoneda") <> paMonedaPesos Then
                MsgBox "No se pueden asignar cheques en moneda extranjera." & vbCrLf & vbCrLf & "DEBE INGRESAR COMPRA DE Moneda Extranjera", vbExclamation, "ATENCIÓN"
'                TCME = InputBox("Ingrese a que TASA DE CAMBIO se tomó el cheque " & rsT("Cheque") & " de " & Trim(rsT("MonSigno")) & " " & Format(rsT("CDiImporte"), "#,##0.00"), "TASA DE CAMBIO")
'                If Not IsNumeric(TCME) Then
'                    TCME = 1
'                End If
'                .Cell(flexcpText, .Rows - 1, 1) = Format(CCur(TCME), "#,##0.00")
'                .Cell(flexcpText, .Rows - 1, 2) = Format(rsT("CDiImporte") * CCur(TCME), "#,##0.00")
            Else
                .AddItem "  Cheque " & Trim(rsT("Cheque"))
                .Cell(flexcpText, .Rows - 1, 2) = Format(rsT("CDiImporte"), "#,##0.00")
                .Cell(flexcpData, .Rows - 1, 0) = -1
                .Cell(flexcpData, .Rows - 1, 2) = CStr(rsT("CDiCodigo"))
            End If
            
'            .Cell(flexcpData, .Rows - 1, 0) = -1
'            .Cell(flexcpData, .Rows - 1, 2) = CStr(rsT("CDiCodigo"))
        End With
        
        rsT.MoveNext
    Loop
    rsT.Close
    
    
    vsPagaCon.Subtotal flexSTSum, -1, 2, , &HA0A000, vbWhite, True, "Total Paga Con (" & sSignMonedaPesos & ")"
    
    'LES CAMBIO EL TAG A LOS CHEQUES.
    Cons = "UPDATE ChequeDiferido SET CDiTAG = 1 " & sQueryCont  'WHERE CDiTAG = " & idtag
    cBase.Execute Cons
    
    ArmoDetalleLiquidacion
    
    Exit Sub
errTC:
    clsGeneral.OcurrioError "Error al cargar los cheques para el tag", Err.Description, "Tag de cheques"
End Sub

Private Function ObtenerTagCheques() As Integer
On Error GoTo errTC
Dim rsT As rdoResultset
Dim iUtilizo As Integer
    iUtilizo = 254
    Cons = "SELECT DISTINCT(Isnull(CDiTAG, 0)) CDiTAG FROM ChequeDiferido ORDER BY CDiTag DESC"
    Set rsT = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsT.EOF
        If iUtilizo < rsT("CDiTag") Then
            'nada
        ElseIf iUtilizo = rsT("CDiTag") Then
            iUtilizo = iUtilizo - 1
        Else
            Exit Do
        End If
        rsT.MoveNext
    Loop
    rsT.Close
    ObtenerTagCheques = iUtilizo
    Exit Function
errTC:
    clsGeneral.OcurrioError "Error al buscar el tag de cheques", Err.Description, "Tag de cheques"
End Function

Private Function GetImportePagoEfectivo() As Currency
    GetImportePagoEfectivo = GetImportePago(0) + GetImportePago(-2)
End Function

Private Function GetImportePago(ByVal tipoOdisponibilidad As Long) As Currency
Dim fila As Integer
Dim suma As Currency
    With vsPagaCon
        For fila = .FixedRows + 1 To .Rows - 1
            If Val(.Cell(flexcpData, fila, 0)) = tipoOdisponibilidad Then
                suma = CCur(.Cell(flexcpText, fila, 2)) + suma
            End If
        Next
    End With
    GetImportePago = suma
End Function

Private Function GetCantidadRenglonesPago(ByVal tipoOdisponibilidad As Long) As Integer
Dim fila As Integer
Dim suma As Currency
    With vsPagaCon
        For fila = .FixedRows + 1 To .Rows - 1
            If Val(.Cell(flexcpData, fila, 0)) = tipoOdisponibilidad Then
                GetCantidadRenglonesPago = GetCantidadRenglonesPago + 1
            End If
        Next
    End With
End Function

Private Function GetImportePagoTotal() As Currency
Dim fila As Integer
Dim suma As Currency
    With vsPagaCon
        For fila = .FixedRows + 1 To .Rows - 1
            suma = CCur(.Cell(flexcpText, fila, 2)) + suma
        Next
    End With
    GetImportePagoTotal = suma
End Function

Private Function GetDetalle(ByVal Tipo As TipoDetalleLiquidacion, ByVal colAR As Collection) As clsDetalleLiquidacion
    
    Set GetDetalle = Nothing
    Dim oRec As clsDetalleLiquidacion
    For Each oRec In colAR
        If oRec.Tipo = Tipo Then Set GetDetalle = oRec: Exit Function
    Next
    
End Function

'Private Function fnc_ValidoLiquidacionesPendientes() As Boolean
'Dim cPend As Currency
'
'    'cPend = fnc_GetImporteReclamo(PendientesLiquidaciones)         'fnc_GetTotalPendienteSeleccionado(0)
'    cPend = GetImporteReclamarDiferencias
'    If cPend > 0 Then
'
'        'Tengo pendientes.
'        Dim cPaga As Currency
'        'Cuanto me paga.
'        cPaga = fnc_GetImporteAporte(Pagos) + fnc_GetImporteAporte(PendientesLiquidaciones)           'CCur(vsLiquidacion.Cell(flexcpText, iGTReclamo + 1, 2))
'
'        If cPaga = 0 Then
'            MsgBox "Para saldar pendientes sí o sí debe pagar parte o el total de los mismos.", vbCritical, "ATENCIÓN"
'            Exit Function
'        End If
'
'        If cPaga >= cPend Then
'
'            'Lo que paga supera o salda el pendiente.
'            fnc_ValidoLiquidacionesPendientes = True
'            Exit Function
'
'        Else
'
'            'tengo que recorrer para ver si tengo + de 1 liquidación pendiente
'            'que no salde la diferencia.
'            Dim cDif As Currency
'            cDif = cPaga
'            With vsConsulta
'
'                For i = 1 To .Rows - 1
'                    If .Cell(flexcpBackColor, i, 0) <> cte_CelesteClaro And .Cell(flexcpData, i, 0) = 4 And CCur(.Cell(flexcpData, i, 3)) > 0 Then
'                        If cDif > 0 Then
'                            cDif = cDif - CCur(.Cell(flexcpData, i, 3))
'                        Else
'                            MsgBox "ATENCIÓN!!!" & vbCrLf & "Lo que ud cobra no salda todos los pendientes, debe eliminar aquellos pendientes que exceden el pago (pago parcial sólo puede haber en 1 pendiente).", vbCritical, "POSIBLE ERROR"
'                            Exit Function
'                        End If
'                    End If
'                Next
'
'            End With
'
'            'Si sale es por que puedo seguir
'            fnc_ValidoLiquidacionesPendientes = True
'
'        End If
'
'    Else
'
'        If fnc_GetImporteAporte(PendientesLiquidaciones) > 0 And Not fnc_HayEnviosEnLiquidacion Then
'            MsgBox "No puede utilizar aportes sin liquidar envíos o pendientes.", vbExclamation, "ATENCIÓN"
'            fnc_ValidoLiquidacionesPendientes = False
'        Else
'            fnc_ValidoLiquidacionesPendientes = True
'        End If
'
'    End If
'End Function

Private Function fnc_GetTotalPendienteSeleccionado(ByVal tipoImporte As Byte) As Currency
    
    fnc_GetTotalPendienteSeleccionado = 0
    With vsConsulta
        For i = 1 To .Rows - 1
            If .Cell(flexcpBackColor, i, 0) <> cte_CelesteClaro And .Cell(flexcpData, i, 0) = 4 Then
                If (tipoImporte = 1 And CCur(.Cell(flexcpData, i, 3)) > 0) Or tipoImporte = 0 Then
                    fnc_GetTotalPendienteSeleccionado = fnc_GetTotalPendienteSeleccionado + CCur(.Cell(flexcpData, i, 3))
                ElseIf tipoImporte = 2 And CCur(.Cell(flexcpData, i, 3)) < 0 Then
                    fnc_GetTotalPendienteSeleccionado = fnc_GetTotalPendienteSeleccionado + CCur(.Cell(flexcpData, i, 3))
                End If
            End If
        Next
    End With
End Function

Private Function fnc_HayEnviosEnLiquidacion() As Boolean
    fnc_HayEnviosEnLiquidacion = False
    With vsConsulta
        For i = 1 To .Rows - 1
            If .Cell(flexcpBackColor, i, 0) <> cte_CelesteClaro And .Cell(flexcpData, i, 0) <> 4 Then
                fnc_HayEnviosEnLiquidacion = True
                Exit Function
            End If
        Next
    End With
End Function

Private Function fnc_GetStringDocumentosPendientes(ByVal iEnv As Long, ByVal Tipo As Byte) As String
    fnc_GetStringDocumentosPendientes = "Select * From DocumentoPendiente, Documento " & _
        "Where DPeTipo = " & Tipo & " And DPeIDTipo = " & iEnv & " And DPeDocumento = DocCodigo And DocAnulado = 0 And DPeIDLiquidacion Is Null"
End Function

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bImprimir_Click()
    AccionImprimir True
End Sub

Private Sub bLastHtml_Click()
    If prmPlantilla <> 0 Then EjecutarApp App.Path & "\appExploreMsg.exe ", prmPlantilla & ":C" & lLastID
End Sub

Private Sub bPreview_Click()
    AccionImprimir
End Sub

Private Sub bQuery_Click()
    AccionConsultar
End Sub

Private Sub caPesos_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If cLiquidar.ListIndex > -1 And cCamion.ListIndex > -1 Then ' And iGTReclamo > 0 Then
            
            With vsDetMonto
                .AddItem Trim(cLiquidar.Text)
                .Cell(flexcpText, .Rows - 1, 1) = Trim(tConAmp.Text)
                .Cell(flexcpText, .Rows - 1, 2) = 1
                .Cell(flexcpText, .Rows - 1, 3) = sSignMonedaPesos & " " & Format(caPesos.Text, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 4) = Format(caPesos.Text, "#,##0.00")
            End With
'            s_SetTotalBoleta
'            SetTotalLiquidacion

            ArmoDetalleLiquidacion
            
            cLiquidar.Text = ""
            caPesos.Text = 0
            tConAmp.Text = ""
            cLiquidar.SetFocus
        End If
    End If
End Sub

Private Sub cboOpcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cCamion_KeyPress 13
End Sub

Private Sub cLiquidar_GotFocus()
    With cLiquidar
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione un concepto de liquidación."
End Sub

Private Sub cLiquidar_KeyPress(KeyAscii As Integer)
Dim iCont As Integer
    If KeyAscii = vbKeyReturn Then
        If cLiquidar.ListIndex > -1 And cCamion.ListIndex > -1 Then
            'Válido si ya existe el item ingresado.
            With vsDetMonto
                For iCont = 1 To .Rows - 1
                    If Trim(.Cell(flexcpText, iCont, 0)) = Trim(cLiquidar.Text) Then
                        If MsgBox("El concepto a ingresar ya esta en la lista." & vbCr & "¿Desea ingresar una nueva liquidación con este concepto?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible Error") = vbNo Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    End If
                Next iCont
            End With
            '.........................................................
            tConAmp.SetFocus
        Else
            If tComentario.Enabled Then cTipo.SetFocus
        End If
    End If
End Sub

Private Sub cLiquidar_LostFocus()
    Ayuda ""
End Sub

Private Sub cMoneda_Click()
    EtiquetaLiquidacion
End Sub

Private Sub cMoneda_GotFocus()
    s_ChangeMoneda
    With cMoneda
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione una moneda."
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If cMoneda.ListIndex > -1 Then tSon.SetFocus Else tComentario.SetFocus
End Sub

Private Sub cMoneda_LostFocus()
    Ayuda ""
End Sub

Private Sub cTipo_Click()
    cMoneda.ListIndex = -1
    EtiquetaLiquidacion
End Sub

Private Sub cTipo_GotFocus()
    Ayuda "Seleccione con que le paga el camión."
End Sub

Private Sub cTipo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errCT

'ANULE TODO ESTO A PEDIDO DE JULIANA. 25/09/2013

'    If Shift = 0 And KeyCode = vbKeyF1 And cTipo.ListIndex > -1 Then
'        If cTipo.ItemData(cTipo.ListIndex) = -1 Then
'
'            'Cargo los últimos cheques ingresados (no filtro por el tag).
'            Cons = "SELECT Top 15 CDiCodigo, RTRIM(RTRIM(CDiSerie) + ' ' + CONVERT(varchar(10), CDiNumero)) 'Serie Nro.', MonSigno Moneda, CDiImporte Importe, UsuIdentificacion Usuario" & _
'                " FROM ChequeDiferido INNER JOIN Moneda ON MonCodigo = CDiMoneda INNER JOIN Usuario ON UsuCodigo = CDiUsuario ORDER BY CDiCodigo DESC"
'            Dim oAyuda As New clsListadeAyuda
'            If oAyuda.ActivarAyuda(cBase, Cons, 5000, 1, "Buscar cheques") > 0 Then
'                If MsgBox("¿Confirma que el cheque seleccionado paga una liquidación?", vbQuestion + vbYesNo, "Cheque de camioneros") = vbYes Then
'                    CargarChequesDeTag 0, oAyuda.RetornoDatoSeleccionado(0)
'                End If
'            End If
'            Set oAyuda = Nothing
'
'        End If
'    End If
    Exit Sub
errCT:
    clsGeneral.OcurrioError "Error al buscar cheques.", Err.Description, "Ayuda de cheques"
End Sub

Private Sub cTipo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And cTipo.ListIndex > -1 Then
        
        If cTipo.ItemData(cTipo.ListIndex) >= 0 Then
            s_ChangeMoneda
            tSon.SetFocus
        ElseIf cTipo.ItemData(cTipo.ListIndex) = -3 Then
            'Abro formulario con que cobra.
            tSon.SetFocus
        ElseIf cTipo.ItemData(cTipo.ListIndex) = -4 Then
            'Abro formulario con que cobra.
            cMoneda.ListIndex = 0
            BuscoCodigoEnCombo cMoneda, CLng(paMonedaPesos)
            tSon.SetFocus
        ElseIf cTipo.ListIndex > -1 And cCamion.ListIndex > -1 Then
            If cTipo.ItemData(cTipo.ListIndex) = -1 Then
                
                Dim idtag As Integer
                idtag = ObtenerTagCheques
                If idtag <= 0 Then Exit Sub
                
                EjecutarApp App.Path & "\altacheques.exe", "L " & idtag, True
                
                'Busco si me ingreso cheques por el tag si es así entonces los agrego.
                CargarChequesDeTag idtag, 0
            Else
                cMoneda.SetFocus
            End If
        End If
        
    End If
    
End Sub

Private Sub cTipo_LostFocus()
    Ayuda ""
End Sub

Private Sub cTipo_Validate(Cancel As Boolean)
    s_ChangeMoneda
End Sub

Private Sub Label8_Click()
    tUsuario.SetFocus
End Sub

Private Sub lLiquidar_Click()
    Foco cLiquidar
End Sub

Private Sub lSon_Click()
    cMoneda.SetFocus
End Sub

Private Sub MnuAcEnvio_Click()
    If vsConsulta.Cell(flexcpText, vsConsulta.Row, 0) = "DifL" Then Exit Sub
    Dim objEnvio As New clsEnvio
    objEnvio.InvocoEnvio vsConsulta.Cell(flexcpText, vsConsulta.Row, 1), gPathListados
    Set objEnvio = Nothing
    ReturnEnvio vsConsulta.Row, vsConsulta.Cell(flexcpText, vsConsulta.Row, 1)
End Sub

Private Sub MnuAcServicio_Click()
    EjecutarApp App.Path & "\Seguimiento de Servicios.exe", vsConsulta.Cell(flexcpText, vsConsulta.Row, 1), True
    AccionConsultar
End Sub

Private Sub MnuDesmarcar_Click()
    MarcaroDesmarcar False
End Sub

Private Sub MnuMarcarTodos_Click()
    MarcaroDesmarcar True
End Sub

Private Sub Status_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "zureo" Then
        fnc_DoTestLogin
    End If
End Sub

Private Function fnc_EsZureoCGSA() As Boolean
On Error GoTo errZ
    Set RsAux = rdoCZureo.OpenResultset("Select Top 1 ParValor From ZUREOCGSA.dbo.genParametros", rdOpenDynamic, rdConcurValues)
    fnc_EsZureoCGSA = Not RsAux.EOF
    RsAux.Close
    Exit Function
errZ:
fnc_EsZureoCGSA = False
End Function

Private Function fnc_DoTestLogin(Optional ByVal bSave As Boolean = False) As Boolean
    Screen.MousePointer = 11
    Dim sRet As String
    If objUsers Is Nothing Then Set objUsers = New clsCheckIn
        
    If Not fnc_ValidoAcceso(sRet) Then
        fnc_DoTestLogin = False
        Status.Panels("zureo").Text = "ZUREO OFF"
        If Not bSave Then
            MsgBox "El programa no logró hacer el login en zureo." & vbCrLf & vbCrLf & "Retorno: " & sRet, vbExclamation, "Acceso a Zureo"
        End If
    Else
        fnc_DoTestLogin = True
        Status.Panels("zureo").Text = "ZUREO ON"
    End If
    Screen.MousePointer = 0
End Function

Private Sub tA_GotFocus()
    With tA
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Function BuscoNota() As Long
On Error GoTo errDoc

    tA.Tag = ""
    tSon.Tag = ""
    If Trim(tSon.Text) = "" Then Exit Function
    
    Dim mDSerie As String, mDNumero As Long
    Dim adQ As Integer, adCodigo As Long, adTexto As String
    
    adTexto = Trim(tSon.Text)
    If InStr(adTexto, "-") <> 0 Then
        mDSerie = Mid(adTexto, 1, InStr(adTexto, "-") - 1)
        mDNumero = Val(Mid(adTexto, InStr(adTexto, "-") + 1))
    Else
        adTexto = Replace(adTexto, " ", "")
        If IsNumeric(Mid(adTexto, 2, 1)) Then
            mDSerie = Mid(adTexto, 1, 1)
            mDNumero = Val(Mid(adTexto, 2))
        Else
            mDSerie = Mid(adTexto, 1, 2)
            mDNumero = Val(Mid(adTexto, 3))
        End If
    End If
    tA.Text = UCase(mDSerie) & "-" & mDNumero
    
    Screen.MousePointer = 11
    adQ = 0: adTexto = ""
    
    Cons = "Select DocCodigo, DocFecha as Fecha" & _
               ", CASE DocTipo WHEN 1 Then 'Contado' WHEN 2 Then 'Crédito' WHEN 3 Then 'Nota Dev' WHEN 4 then 'Nota Créd' WHEN 5 then 'Recibo' When 6 Then 'Remito' When 10 Then 'Nota Esp' ELSE '' END Documento " & _
               ", rTrim(DocSerie) + '-' + rtrim(Convert(Varchar(6), DocNumero)) as Número, DocTotal Importe" & _
               " From Documento " & _
               " Where DocSerie = '" & mDSerie & "'" & _
               " And DocNumero = " & mDNumero & _
               " And DocTipo IN (" & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")"
        
    Dim miLDocs As New clsListadeAyuda
    adCodigo = miLDocs.ActivarAyuda(cBase, Cons, 4400, 1)
    Me.Refresh
    If adCodigo > 0 Then
        adCodigo = miLDocs.RetornoDatoSeleccionado(0)
        tSon.Tag = miLDocs.RetornoDatoSeleccionado(4)
    End If
    Set miLDocs = Nothing
    
    If adCodigo > 0 Then
        tA.Tag = adCodigo
        s_AddRenglonPagaCon
    End If
    
    Screen.MousePointer = 0
    Exit Function
errDoc:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub tA_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If cCamion.ListIndex = -1 Then 'Or iGTReclamo = 0 Then
            MsgBox "Debe seleccionar un camión y dar Enter.", vbInformation, "Atención"
            cCamion.SetFocus
            Exit Sub
        End If
        'Válido
        If cTipo.ListIndex > -1 And cMoneda.ListIndex > -1 And IsNumeric(tSon.Text) Then
            If IsNumeric(tA.Text) Then
            'Si puso m/e y la moneda es $ le aviso.
                If cMoneda.ItemData(cMoneda.ListIndex) = paMonedaPesos And cTipo.ItemData(cTipo.ListIndex) = -2 Then
                    MsgBox "Ingresó moneda extranjera para la moneda pesos.", vbExclamation, "ATENCIÓN"
                    cTipo.SetFocus
                    Exit Sub
                ElseIf cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos And cTipo.ItemData(cTipo.ListIndex) = 0 Then
                    MsgBox "Ingresó moneda extranjera para los billetes.", vbExclamation, "ATENCIÓN"
                    cTipo.SetFocus
                    Exit Sub
                ElseIf cTipo.ItemData(cTipo.ListIndex) > 0 And cTipo.ItemData(cTipo.ListIndex) <> paMCCtaComisionRP Then
                    If cMoneda.ItemData(cMoneda.ListIndex) <> f_SelectMonedaDisponibilidad Then
                        MsgBox "Posible error, la disponibilidad seleccionada tiene asignada una moneda distinta a la del combo.", vbExclamation, "Atención"
                        cTipo.SetFocus
                        Exit Sub
                    End If
                End If
                s_AddRenglonPagaCon
            End If
        End If
    End If
    
End Sub

Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tUsuario.Enabled Then
            tUsuario.SetFocus
        Else
            AccionGrabar
        End If
    End If
End Sub

Private Sub cCamion_Change()
    LimpioDatos
End Sub

Private Sub cCamion_Click()
    LimpioDatos
End Sub

Private Sub cCamion_GotFocus()
    With cCamion: .SelStart = 0: .SelLength = Len(.Text): End With
    Ayuda "Seleccione un camión. ([Enter] Consulta)"
End Sub

Private Sub cCamion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cCamion.ListIndex <> -1 Then
        AccionConsultar
        If tsOpcion.SelectedItem.Index = 1 Then
            If vsConsulta.Rows > 1 Then vsConsulta.SetFocus
        End If
    End If
End Sub

Private Sub cCamion_LostFocus()
    Ayuda ""
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me
    picBotones.BorderStyle = vbBSNone: picTextos.BorderStyle = vbBSNone
    InicializoGrillas
    
    paTipoArticuloServicio = 151
    
    lblAyudaCheque.Visible = False
    lblAyudaCheque.Caption = "Presione 'Enter' para dar el alta"
    
    cboOpcion.Clear
    cboOpcion.AddItem "Envíos y servicios"
    cboOpcion.AddItem "Sólo traslados"
    cboOpcion.AddItem "Todos"
    cboOpcion.ListIndex = 0
    
    Dim fHeader As New StdFont, ffooter As New StdFont
    bLastHtml.Picture = Image1.ListImages("last").ExtractIcon
    
    With fHeader
        .Bold = True
        .Name = "Arial"
        .Size = 11
    End With
    With ffooter
        .Bold = True
        .Name = "Tahoma"
        .Size = 10
    End With
    
    With gpPrint
        .Caption = "Liquidación de Camioneros"
        .FileName = "LiquidacionCamionero"
        .Font = Font
        Set .HeaderFont = fHeader
        .Orientation = opPortrait
        .PaperSize = 1
        .PageBorder = opTopBottom
        .MarginLeft = 500
        .MarginTop = 700
        .MarginRight = 400
    End With
    
    With tUsuario
        .Connect cBase
        .SetApp App.Title
        .UserID = 0
        .UserLog = paCodigoDeUsuario
        .Terminal = paCodigoDeTerminal
        If .GetConfigUser = 2 Then .UserID = paCodigoDeUsuario
    End With
    
    picTab(0).ZOrder 0
    
    s_CargoComboFormaPago
    
    FechaDelServidor
    Cons = "Select CamCodigo, CamNombre From Camion Order by CamNombre"
    CargoCombo Cons, cCamion
    
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda
    
    sSignMonedaPesos = BuscoSignoMoneda(CLng(paMonedaPesos))
        
    Cons = "Select CLiCodigo, CLiDescripcion From ConceptoLiquidacion Where CLiTipoEnte = 1 Order by CLiDescripcion"
    CargoCombo Cons, cLiquidar
    
    lIDDispPesos = modMaeDisponibilidad.dis_DisponibilidadPara(paCodigoDeSucursal, CLng(paMonedaPesos))
    loc_GetDatosDisponibilidad lIDDispPesos, iIDSRPesos, 0
    
    tmStart.Enabled = True
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            
            Case vbKeyE: AccionConsultar
            Case vbKeyI: AccionImprimir True
            Case vbKeyP: AccionImprimir False
            
            Case vbKeyG: If bGrabar.Enabled Then AccionGrabar
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11
    picBotones.Top = ScaleHeight - (picBotones.Height + Status.Height)
    picTextos.Top = picBotones.Top - picTextos.Height
    picTextos.Width = ScaleWidth
    
    With tsOpcion
        .Left = 0
        .Top = cCamion.Top + cCamion.Height + 30
        .Width = ScaleWidth
        .Height = picTextos.Top - (.Top + 30)
    End With
    
    With picTab(0)
        .Left = tsOpcion.ClientLeft
        .Height = tsOpcion.ClientHeight
        .Top = tsOpcion.ClientTop
        .Width = tsOpcion.ClientWidth
    End With
    
    With picTab(1)
        .Left = tsOpcion.ClientLeft
        .Height = tsOpcion.ClientHeight
        .Top = tsOpcion.ClientTop
        .Width = tsOpcion.ClientWidth
    End With
    
    With vsConsulta
        .Top = picTab(0).ScaleTop
        .Left = picTab(0).ScaleLeft
        .Width = picTab(0).ScaleWidth
        .Height = picTab(0).ScaleHeight
    End With
        
    vsLiquidacion.Move vsDetMonto.Left, vsLiquidacion.Top, 7600, picTab(1).ScaleHeight - vsLiquidacion.Top
    With vsPagaCon
        .Left = vsLiquidacion.Left
        .Width = vsLiquidacion.Width
    End With
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    Set objUsers = Nothing
    Erase arrPendiente
    
    GuardoSeteoForm Me
    CierroConexion
    rdoCZureo.Close
    Set rdoCZureo = Nothing
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Sub CargoTrasladosDeMercaderia()
On Error GoTo errCTM
    Cons = "SELECT TraCodigo, TraFechaEntregado,  RTrim(Origen.SucAbreviacion) + ' a ' + RTrim(Destino.SucAbreviacion) DeA, TraFModificacion " & _
            "FROM Traspaso INNER JOIN Sucursal Origen ON TraLocalOrigen = Origen.SucCodigo " & _
            "INNER JOIN Sucursal Destino ON TraLocalDestino = Destino.SucCodigo " & _
            "WHERE TraLocalIntermedio = " & cCamion.ItemData(cCamion.ListIndex) & _
            "AND TraLocalOrigen IN (1, 5, 6, 9, 14, 41) AND TraLocalDestino IN (1, 5, 6, 9, 41) " & _
            "AND TraAnulado IS NUll AND TraFechaEntregado IS NOT NULL AND TraComentario NOT LIKE '%L:%'"
    Dim rsTM As rdoResultset
    Set rsTM = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsTM.EOF
        With vsConsulta
            .AddItem "Tra"
            .Cell(flexcpData, .Rows - 1, 0) = 3 'Me digo que es traslado
            .Cell(flexcpData, .Rows - 1, 2) = 1 'Moneda pesos
            .Cell(flexcpData, .Rows - 1, 4) = prmCostoParada
            .Cell(flexcpData, .Rows - 1, 5) = 0
            .Cell(flexcpData, .Rows - 1, 6) = 0
            
            .Cell(flexcpText, .Rows - 1, 9) = Trim(rsTM!TraFModificacion)
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsTM!TraCodigo)
            .Cell(flexcpText, .Rows - 1, 2) = rsTM("DeA")
            .Cell(flexcpText, .Rows - 1, 4) = Format(prmCostoParada, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 5) = Format(0, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 6) = Format(0, "#,##0.00")
            AgregoQueLiquido Traslados, rsTM("TraCodigo"), 0, prmCostoParada, 0, 0, 0
        End With
        
        rsTM.MoveNext
    Loop
    rsTM.Close
    Exit Sub
errCTM:
    clsGeneral.OcurrioError "Error al cargar los traslados de mercadería.", Err.Description, "Liquidación"
End Sub

Private Sub AgregoQueLiquido(ByVal tipoLL As TipoQueCobro, ByVal ID As Long, ByVal Reclamar As Currency, ByVal Liquidar As Currency, ByVal Diferencia As Currency, ByVal PendienteCaja As Currency, ByVal documentopendiente As Currency)
Dim oQueCobro  As clsQueLiquido
    
    Set oQueCobro = New clsQueLiquido
    With oQueCobro
        .ID = ID
        .Liquidar = Liquidar
        .TipoID = tipoLL
        .PendienteCaja = PendienteCaja
        .Reclamar = Reclamar
        .Diferencia = Diferencia
        .ImporteDocumentoPendiente = documentopendiente
        .SeCobra = True                 'Cuando inserto van todos prendidos.
    End With
    colQueCobro.Add oQueCobro
    
End Sub

Private Function BuscarQueCobro(ByVal tipoLL As TipoQueCobro, ByVal ID As Long) As clsQueLiquido
Dim oQueCobro  As clsQueLiquido
    For Each oQueCobro In colQueCobro
        If (oQueCobro.ID = ID And oQueCobro.TipoID = tipoLL) Then
            Set BuscarQueCobro = oQueCobro
            Exit Function
        End If
    Next
End Function

Private Function BuscarQueCobroServicio(ByVal tipoLL As TipoQueCobro, ByVal ID As Long, ByVal Reclamo As Currency) As clsQueLiquido
Dim oQueCobro  As clsQueLiquido
    For Each oQueCobro In colQueCobro
        If (oQueCobro.ID = ID And oQueCobro.TipoID = tipoLL And oQueCobro.Reclamar = Reclamo) Then
            Set BuscarQueCobroServicio = oQueCobro
            Exit Function
        End If
    Next
End Function

Private Sub SeCobraSINO(ByVal tipoLL As TipoQueCobro, ByVal ID As Long, ByVal SeCobra As Boolean, ByVal Liquida As Currency, ByVal Reclamo As Currency)
Dim oQueCobro  As clsQueLiquido
    For Each oQueCobro In colQueCobro
        If (oQueCobro.ID = ID And oQueCobro.TipoID = tipoLL And oQueCobro.Liquidar = Liquida And ((oQueCobro.Reclamar = Reclamo And tipoLL <> LiquidacionesPendientes) Or (oQueCobro.Diferencia = Reclamo And tipoLL = LiquidacionesPendientes))) Then
            oQueCobro.SeCobra = SeCobra: Exit Sub
        End If
    Next
End Sub

Private Function GetImporteLiquidar() As Currency
Dim oQueCobro  As clsQueLiquido
Dim Liquidar As Currency
    For Each oQueCobro In colQueCobro
        If (oQueCobro.SeCobra And oQueCobro.Liquidar <> 0) Then
            Liquidar = Liquidar + oQueCobro.Liquidar
        End If
    Next
    GetImporteLiquidar = Liquidar
End Function

Private Function GetImporteReclamar() As Currency
Dim oQueCobro  As clsQueLiquido
Dim Liquidar As Currency
    For Each oQueCobro In colQueCobro
        If (oQueCobro.SeCobra And oQueCobro.Reclamar > 0) Then
            Liquidar = Liquidar + oQueCobro.Reclamar
        End If
    Next
    GetImporteReclamar = Liquidar
End Function

Private Function GetImporteReclamarDiferencias() As Currency
Dim oQueCobro  As clsQueLiquido
Dim retorno As Currency
    For Each oQueCobro In colQueCobro
        If (oQueCobro.SeCobra And oQueCobro.Diferencia > 0 And oQueCobro.TipoID = LiquidacionesPendientes) Then
            retorno = retorno + oQueCobro.Diferencia
        End If
    Next
    GetImporteReclamarDiferencias = retorno
End Function

Private Function GetImporteAporteDiferencias() As Currency
Dim oQueCobro  As clsQueLiquido
Dim Liquidar As Currency
    For Each oQueCobro In colQueCobro
        If (oQueCobro.SeCobra And oQueCobro.Diferencia < 0 And oQueCobro.TipoID = LiquidacionesPendientes) Then
            Liquidar = Liquidar + oQueCobro.Diferencia
        End If
    Next
    GetImporteAporteDiferencias = Abs(Liquidar)
End Function

Private Function GetImportePendienteCaja(ByVal Reclamar As Boolean) As Currency
Dim iCont As Integer
Dim retorno As Currency
On Error Resume Next
    For iCont = 1 To UBound(arrPendiente)
        If arrPendiente(iCont).Activo Then
            If Reclamar And arrPendiente(iCont).Importe > 0 Then
                retorno = retorno + arrPendiente(iCont).Importe
            ElseIf Not Reclamar And arrPendiente(iCont).Importe < 0 Then
                retorno = retorno + Abs(arrPendiente(iCont).Importe)
            End If
        End If
    Next
    'Retorno SIEMPRE EL POSITIVO
    GetImportePendienteCaja = Abs(retorno)
End Function

Private Function GetRenglonesBoleta() As Collection
Dim oQueCobro  As clsQueLiquido
Dim Liquidar As Currency
Dim oLiq As clsImporteCantidad, oLiqFind As clsImporteCantidad
Dim retorno As Collection

    Set retorno = New Collection
    
    For Each oQueCobro In colQueCobro
        
        If oQueCobro.SeCobra And oQueCobro.Liquidar <> 0 Then
            Set oLiq = New clsImporteCantidad
            For Each oLiqFind In retorno
                If oLiqFind.Importe = oQueCobro.Liquidar Then
                    Set oLiq = oLiqFind
                    Exit For
                End If
            Next
            If oLiq.Cantidad = 0 And oLiq.Importe = 0 Then
                oLiq.Cantidad = 1
                oLiq.Importe = oQueCobro.Liquidar
                retorno.Add oLiq
            Else
                oLiq.Cantidad = oLiq.Cantidad + 1
            End If
        End If
        
    Next
    
    Set GetRenglonesBoleta = retorno
    
End Function

Private Sub CargoLiquidacionesPendientes()
Dim rsLP As rdoResultset
Dim sQy As String
Dim oLiCobro As clsLiquidacionCobro
Dim cImp As Currency
    
    sQy = "SELECT LiqID, LCoMoneda, LiqTotal, IsNull(SUM(LCoCobrado), 0) Cobrado FROM Liquidacion " & _
        "LEFT OUTER JOIN LiquidacionCobro ON LCoLiquidacion = LiqID " & _
        "WHERE LiqTipo = 1 AND LiqEnte = " & cCamion.ItemData(cCamion.ListIndex) & _
        " GROUP BY LiqID, LCoMoneda, LiqTotal " & _
        "HAVING (IsNull(Sum(LCoCobrado), 0) <> LiqTotal)"
    
    Set rsLP = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    
    Dim oQueCobro As clsQueLiquido
    
    Do While Not rsLP.EOF
        'Agrego a la grilla.
        With vsConsulta
            .AddItem "DifL"
            
            .Cell(flexcpData, .Rows - 1, 0) = TipoQueCobro.LiquidacionesPendientes
            
            .Cell(flexcpText, .Rows - 1, 1) = rsLP("LiqID") 'oLiCobro.Liquidacion
            .Cell(flexcpText, .Rows - 1, 2) = ct_KeyDisponibilidadCamioneros
            
            cImp = Format(rsLP("LiqTotal") - rsLP("Cobrado"), "#,##0.00")
            
            .Cell(flexcpText, .Rows - 1, 4) = Format(0, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 5) = Format(cImp, "#,##0.00")
            
            .Cell(flexcpText, .Rows - 1, 6) = Format(0, "#,##0.00")
            
            .Cell(flexcpData, .Rows - 1, 2) = paMonedaPesos
            .Cell(flexcpData, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 5)
            .Cell(flexcpData, .Rows - 1, 9) = .Cell(flexcpText, .Rows - 1, 5)
            
            AgregoQueLiquido LiquidacionesPendientes, rsLP("LiqID"), 0, 0, cImp, 0, 0
           
        End With
        
        rsLP.MoveNext
    Loop
    rsLP.Close
    
    Exit Sub
    
    Set oLiCobro = New clsLiquidacionCobro
    oLiCobro.Liquidacion = 21650
    oLiCobro.Moneda = 1
    oLiCobro.TotalLiquidacion = 0
    oLiCobro.TotalCobrado = 500
    colLiqCobro.Add oLiCobro

    'Agrego a la grilla.
    With vsConsulta
        .AddItem "DifL"
        .Cell(flexcpData, .Rows - 1, 0) = 4 'Me digo que es por liquidación pendiente.
        .Cell(flexcpText, .Rows - 1, 1) = oLiCobro.Liquidacion
        .Cell(flexcpText, .Rows - 1, 2) = ct_KeyDisponibilidadCamioneros

        .Cell(flexcpText, .Rows - 1, 4) = Format(0, "#,##0.00")
        .Cell(flexcpText, .Rows - 1, 5) = Format(oLiCobro.TotalLiquidacion - oLiCobro.TotalCobrado, "#,##0.00")

        .Cell(flexcpText, .Rows - 1, 6) = Format(0, "#,##0.00")

        .Cell(flexcpData, .Rows - 1, 2) = paMonedaPesos
        .Cell(flexcpData, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 5)
        .Cell(flexcpData, .Rows - 1, 9) = .Cell(flexcpText, .Rows - 1, 5)
        
        AgregoQueLiquido LiquidacionesPendientes, 21650, 0, 0, .Cell(flexcpText, .Rows - 1, 5), 0, 0

    End With

    Exit Sub

    Set oLiCobro = New clsLiquidacionCobro
    oLiCobro.Liquidacion = 2150
    oLiCobro.Moneda = 1
    oLiCobro.TotalLiquidacion = 10000
    oLiCobro.TotalCobrado = 15000
    colLiqCobro.Add oLiCobro

    'Agrego a la grilla.
    With vsConsulta
        .AddItem "DifL"
        .Cell(flexcpData, .Rows - 1, 0) = 4 'Me digo que es por liquidación pendiente.
        .Cell(flexcpText, .Rows - 1, 1) = oLiCobro.Liquidacion
        .Cell(flexcpText, .Rows - 1, 2) = "Liquidación pendiente"

        .Cell(flexcpText, .Rows - 1, 4) = Format(0, "#,##0.00")
        .Cell(flexcpText, .Rows - 1, 5) = Format(oLiCobro.TotalLiquidacion - oLiCobro.TotalCobrado, "#,##0.00")

        .Cell(flexcpText, .Rows - 1, 6) = Format(0, "#,##0.00")

        .Cell(flexcpData, .Rows - 1, 2) = paMonedaPesos
        .Cell(flexcpData, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 5)
        .Cell(flexcpData, .Rows - 1, 9) = .Cell(flexcpText, .Rows - 1, 5)
'        InsertoEnTablaPago PendientesLiquidaciones, CLng(paMonedaPesos), 0, 0, 1, 0, .Cell(flexcpText, .Rows - 1, 5)
    End With
    
End Sub

Private Sub AccionConsultar()
Dim rs As rdoResultset
Dim sErr As String
    
    On Error GoTo errConsultar
    Erase arrPendiente
    ReDim Preserve arrPendiente(0)
    
    lIDLastSave = 0
    If cCamion.ListIndex = -1 And txtCodImpresion.Text <> "" Then
        MsgBox "Seleccione un camión.", vbExclamation, "ATENCIÓN": Foco cCamion: Exit Sub
    End If
'    Set colReclamos = New Collection
'    Set colPagos = New Collection
    Set colQueCobro = New Collection
    
'    iGTReclamo = -1
    Screen.MousePointer = 11
    InicializoGrillas
    vsConsulta.Refresh
    vsConsulta.Redraw = False
    Status.SimpleText = "Consultando ..."
    Set colLiqCobro = New Collection
    
    ReDim arrPendiente(0)
    If cboOpcion.ListIndex = 0 Or cboOpcion.ListIndex = 2 Then
        CargoLiquidacionesPendientes
        CargoEnviosReparto
    End If
    
    If cCamion.ListIndex > -1 And txtCodImpresion.Text = "" And (cboOpcion.ListIndex = 0 Or cboOpcion.ListIndex = 2) Then
        sErr = CargoServicios
        If sErr <> "" Then GoTo errConsultar
    End If
    
    If cCamion.ListIndex > -1 And txtCodImpresion.Text = "" And (cboOpcion.ListIndex = 1 Or cboOpcion.ListIndex = 2) Then
        CargoTrasladosDeMercaderia
    End If
    
    vsConsulta.Redraw = True
    bGrabar.Enabled = True
    ArmoDetalleLiquidacion
    Status.SimpleText = ""
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    Status.SimpleText = ""
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", IIf(sErr <> "", "Error: " & sErr, Err.Description)
    vsConsulta.Redraw = True
End Sub

Private Sub Label4_Click()
    Foco cCamion
End Sub

Private Sub tConAmp_GotFocus()
    With tConAmp
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tConAmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then caPesos.SetFocus
End Sub

Private Sub tmStart_Timer()
    tmStart.Enabled = False
    fnc_DoTestLogin
End Sub

Private Sub tSon_GotFocus()
    With tSon
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Indique la cantidad "
End Sub

Private Sub tSon_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If cTipo.ItemData(cTipo.ListIndex) = -4 Then
                If BuscoNota() Then
                    tA.Text = ""
                    cTipo.SetFocus
                End If
                tA.Tag = ""
                tSon.Tag = ""
        ElseIf IsNumeric(tSon.Text) Then
            If (cTipo.ItemData(cTipo.ListIndex) = -3) Then
                BuscoTransaccionRP tSon.Text
            Else
                If cMoneda.ListIndex > -1 Then
                    If cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos Then
                        tA.Text = TasadeCambio(cMoneda.ItemData(cMoneda.ListIndex), paMonedaPesos, gFechaServidor, , paTC)
                    End If
                    tA.SetFocus
                Else
                    If cMoneda.Visible Then cMoneda.SetFocus
                End If
            End If
        ElseIf Trim(tSon.Text) = "" Then
            tComentario.SetFocus
        Else
            MsgBox "Formato incorrecto.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub

Private Sub tSon_LostFocus()
    If IsNumeric(tSon.Text) Then
        If (cTipo.ItemData(cTipo.ListIndex) <> -3) Then
            tSon.Text = Format(tSon.Text, "#,##0.00")
        End If
    Else
        tSon.Text = ""
    End If
End Sub

Private Sub tsOpcion_Click()
    picTab(tsOpcion.SelectedItem.Index - 1).ZOrder 0
End Sub

Private Sub tsOpcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tsOpcion.SelectedItem.Index = 2 Then cLiquidar.SetFocus
    End If
End Sub

Private Sub tUsuario_AfterDigit()
    If vsConsulta.Rows > 1 Or GetImportePagoTotal <> 0 Then AccionGrabar
End Sub

Private Sub txtCodImpresion_Change()
    cCamion.Text = ""
    LimpioDatos
End Sub

Private Sub txtCodImpresion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        AccionConsultar
    End If
End Sub

Private Sub vsConsulta_DblClick()
Dim iCont As Integer, iAlta As Integer
    
    If vsConsulta.Row >= 1 Then
        
        If vsConsulta.Cell(flexcpBackColor, vsConsulta.Row, 0) = cte_CelesteClaro Then
            iAlta = 1
            vsConsulta.Cell(flexcpBackColor, vsConsulta.Row, 0, , vsConsulta.Cols - 1) = vbWhite
        Else
            iAlta = -1
            vsConsulta.Cell(flexcpBackColor, vsConsulta.Row, 0, , vsConsulta.Cols - 1) = cte_CelesteClaro
        End If
        SeCobraSINO vsConsulta.Cell(flexcpData, vsConsulta.Row, 0), CLng(vsConsulta.Cell(flexcpText, vsConsulta.Row, 1)), (iAlta = 1), CCur(vsConsulta.Cell(flexcpData, vsConsulta.Row, 4)), CCur(vsConsulta.Cell(flexcpData, vsConsulta.Row, 3))
        
        If vsConsulta.Cell(flexcpData, vsConsulta.Row, 8) <> 0 Then
            For iCont = 1 To UBound(arrPendiente)
                If arrPendiente(iCont).Envio = CLng(vsConsulta.Cell(flexcpText, vsConsulta.Row, 1)) Then
                    arrPendiente(iCont).Activo = (iAlta = 1)
'                    InsertoPendienteDePago arrPendiente(iCont).Moneda, arrPendiente(iCont).Importe, iAlta
                End If
            Next
        End If
        ArmoDetalleLiquidacion
    End If
    
End Sub

Private Sub vsConsulta_GotFocus()
    Ayuda "Lista a liquidar. ([Espaciador], [DblClick] ó [Supr] elimina de la lista)"
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
Dim iCont As Integer
Dim iAlta As Integer
    
    If (KeyCode = vbKeySpace Or KeyCode = vbKeyDelete) And vsConsulta.Row > 0 Then
        
        If vsConsulta.Cell(flexcpBackColor, vsConsulta.Row, 0) = cte_CelesteClaro Then
            iAlta = 1
            vsConsulta.Cell(flexcpBackColor, vsConsulta.Row, 0, , vsConsulta.Cols - 1) = vbWhite
        Else
            iAlta = -1
            vsConsulta.Cell(flexcpBackColor, vsConsulta.Row, 0, , vsConsulta.Cols - 1) = cte_CelesteClaro
        End If
        
        SeCobraSINO vsConsulta.Cell(flexcpData, vsConsulta.Row, 0), CLng(vsConsulta.Cell(flexcpText, vsConsulta.Row, 1)), (iAlta = 1), CCur(vsConsulta.Cell(flexcpData, vsConsulta.Row, 4)), CCur(vsConsulta.Cell(flexcpData, vsConsulta.Row, 3))
        
        If vsConsulta.Cell(flexcpData, vsConsulta.Row, 8) <> 0 Then
            For iCont = 1 To UBound(arrPendiente)
                If arrPendiente(iCont).Envio = CLng(vsConsulta.Cell(flexcpText, vsConsulta.Row, 1)) Then
                    arrPendiente(iCont).Activo = (iAlta = 1)
                End If
            Next
        End If
        ArmoDetalleLiquidacion
    End If
End Sub

Private Sub vsConsulta_LostFocus()
    Ayuda ""
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 And vsConsulta.Row >= 1 Then
        MnuAcEnvio.Enabled = (TipoQueCobro.EnviosMerc = vsConsulta.Cell(flexcpData, vsConsulta.Row, 0))
        MnuAcServicio.Enabled = (TipoQueCobro.ServicioEntRet = vsConsulta.Cell(flexcpData, vsConsulta.Row, 0))
'        If vsConsulta.Cell(flexcpData, vsConsulta.Row, 0) = 2 Then
'            'Servicio
'            MnuAcEnvio.Enabled = False
'            MnuAcServicio.Enabled = True
'        Else
'            MnuAcEnvio.Enabled = True
'            MnuAcServicio.Enabled = False
'        End If
        PopupMenu MnuAcceso
    End If
    
End Sub

Private Sub vsDetMonto_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDelete And vsDetMonto.Rows >= 1 Then
        If Val(vsDetMonto.Cell(flexcpData, vsDetMonto.Row, 0)) = 0 And Not vsDetMonto.IsSubtotal(vsDetMonto.Row) Then
            
            'Elimino una línea de la boleta.
'            vsLiquidacion.Cell(flexcpText, iTConcepto - 2, 2) = Format(CCur(vsLiquidacion.Cell(flexcpValue, iTConcepto - 2, 2)) - CCur(vsDetMonto.Cell(flexcpValue, vsDetMonto.Row, 4)), "#,##0.00")

            vsDetMonto.RemoveItem vsDetMonto.Row
            
'            s_SetTotalBoleta
'            SetTotalLiquidacion

            ArmoDetalleLiquidacion
        End If
    End If
    
End Sub

Private Sub vsLiquidacion_GotFocus()
    Ayuda "Lista de liquidación."
End Sub

Private Sub vsLiquidacion_LostFocus()
    Ayuda ""
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False, Optional lLiqC As Long = 0, Optional sIDLiq As String = "")
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    vsConsulta.ExtendLastCol = False
    vsPagaCon.ExtendLastCol = False
    With gpPrint
        .Header = "Liquidación del Camión: " & Trim(cCamion.Text)
        If lLiqC > 0 Then
            .LineBeforeGrid "Código de liquidación de camionero:" & lLiqC, , , True
            .LineBeforeGrid ""
        End If
        If sIDLiq <> "" Then
            .LineBeforeGrid "Código(s) de Movimientos de disponibilidad: " & sIDLiq, , , True
            .LineBeforeGrid ""
        End If
        .AddGrid vsConsulta.hwnd
        
        .AddGrid vsDetMonto.hwnd
        .AddGrid vsPagaCon.hwnd
        .AddGrid vsLiquidacion.hwnd
    End With
    If Trim(tComentario.Text) <> "" Then
        gpPrint.LineAfterGrid ""
        gpPrint.LineAfterGrid "Comentario: " & Trim(tComentario.Text)
    End If
    gpPrint.LineAfterGrid ""
    gpPrint.LineAfterGrid "Usuario: " & tUsuario.UserName
    If Imprimir Then
        gpPrint.GoPrint
    Else
        gpPrint.ShowPreview
    End If
    vsConsulta.ExtendLastCol = True
    vsPagaCon.ExtendLastCol = True
'    s_HideUnSelect False
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    vsPagaCon.ExtendLastCol = True
    vsConsulta.ExtendLastCol = True
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
'    s_HideUnSelect False
End Sub

Private Sub Ayuda(strTexto As String)
    Status.Panels("msg").Text = strTexto
End Sub

Private Function CamionPagoPorHora(ByVal idCamion As Long) As Boolean
Dim rsC As rdoResultset
    
    Cons = "Select * From Camion Where CamCodigo = " & idCamion
    Set rsC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If rsC!CamPorHora = 0 Then CamionPagoPorHora = False Else CamionPagoPorHora = True
    rsC.Close
    
End Function
Private Sub CargoEnviosReparto()
Dim sPorHora As Boolean, cLiquidar As Currency, cCobrar As Currency, cFacCamion As Currency
Dim aModificacion As String
Dim RsDifEnvio As rdoResultset
Dim lIDEnv As Long, cImpDocPend As Currency

    If cCamion.ListIndex > -1 Then
        sPorHora = CamionPagoPorHora(cCamion.ItemData(cCamion.ListIndex))
    End If
    'Estado Envio = 4 = Entregado
    Cons = "Select * From Envio " _
            & " Left Outer Join PendientesCaja On EnvCodigo = PCaEnvio And PCaFLiquidacion Is Null " _
            & " Left Outer Join Disponibilidad On PCaDisponibilidad = DisID " _
        & " , Direccion, TipoFlete, Moneda" _
        & " Where EnvCamion = " & cCamion.ItemData(cCamion.ListIndex)
    
    If Trim(txtCodImpresion.Text) <> "" Then
        Cons = Cons & "And EnvEstado IN (3,4) And EnvCodImpresion IN (" & txtCodImpresion.Text & ")"
    Else
        Cons = Cons & " And EnvEstado = 4"
    End If
    
    Cons = Cons & " And EnvLiquidacion Is Null  And EnvDireccion = DirCodigo " _
        & " And EnvTipoFlete = TFlCodigo And EnvMoneda = MonCodigo" _
        & " ORDER BY EnvCodigo"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Al agregar el pendiente de caja, un envío puede tener + de uno
    'entonces me puede devolver + de una tupla, por ello agregue control de lidEnv
    lIDEnv = 0
    
    If Not RsAux.EOF And cCamion.ListIndex = -1 Then
        
        BuscoCodigoEnCombo cCamion, RsAux("EnvCamion")
        sPorHora = CamionPagoPorHora(RsAux("EnvCamion"))
        ReDim arrPendiente(0)
        
    End If
    
    Do While Not RsAux.EOF
    
        cLiquidar = 0: cCobrar = 0: cFacCamion = 0
        With vsConsulta
            If lIDEnv <> RsAux!EnvCodigo Then
                cImpDocPend = 0
                
                lIDEnv = RsAux!EnvCodigo
                .AddItem "Env"
                '19/09/2011
                'Agrego el estado ya que si me ingresa por código de impresión tengo que dar por cumplido el envío
                .Cell(flexcpData, .Rows - 1, 6) = CStr(RsAux("EnvEstado"))
                
                .Cell(flexcpData, .Rows - 1, 0) = 1 'Me digo que es por envío
                .Cell(flexcpText, .Rows - 1, 9) = Trim(RsAux!EnvFModificacion)
                
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!EnvCodigo)
                .Cell(flexcpText, .Rows - 1, 2) = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!EnvDireccion)
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!TFlNombreCorto)
                
                'Si tiene para liquidar y no es pago por hora
                If Not IsNull(RsAux!EnvLiquidar) And Not sPorHora Then cLiquidar = RsAux!EnvLiquidar
                
                'Busco documentos pendientes nuevos
                'Si no tengo para el envío procedo por forma anterior.
                Cons = fnc_GetStringDocumentosPendientes(RsAux("EnvCodigo"), 1)
                Set RsDifEnvio = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsDifEnvio.EOF Then
                
                    'Método nuevo
                    Do While Not RsDifEnvio.EOF
                        .Cell(flexcpText, .Rows - 1, 8) = IIf(.Cell(flexcpText, .Rows - 1, 8) <> "", .Cell(flexcpText, .Rows - 1, 8) & ", ", "") & RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero
                        cCobrar = cCobrar + RsDifEnvio("DPeImporte")
                        RsDifEnvio.MoveNext
                    Loop
                    RsDifEnvio.Close
                    
                    cImpDocPend = cCobrar
                    
'                Else
'
'                    'Veo que le reclamo y que le pago.
'                    'Si el envío fue pago en caja, se le pudo agregar diferencia de envío a cobrar en el domicilio.
'                    If RsAux!EnvFormaPago = TipoPagoEnvio.PagaAhora Then
'
'                        Cons = "Select DiferenciaEnvio.*, DocSerie, DocNumero " _
'                                & " From DiferenciaEnvio, Documento Where DEvEnvio = " & RsAux!EnvCodigo _
'                                & " And DEvFormaPago = " & TipoPagoEnvio.PagaDomicilio _
'                                & " And DEvDocumento = DocCodigo And DocAnulado = 0"
'
'                        Set RsDifEnvio = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'                        If Not RsDifEnvio.EOF Then
'                            If Not IsNull(RsDifEnvio!DevValorFlete) Then cCobrar = cCobrar + RsDifEnvio!DevValorFlete
'                            If Not IsNull(RsDifEnvio!DevValorPiso) Then cCobrar = cCobrar + RsDifEnvio!DevValorPiso
'                            .Cell(flexcpText, .Rows - 1, 8) = RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero
'                        End If
'                        RsDifEnvio.Close
'
'                    Else
'
'                        If RsAux!EnvFormaPago = TipoPagoEnvio.PagaDomicilio Then
'                            If Not IsNull(RsAux!EnvDocumentoFactura) Then
'                                Cons = "Select DocSerie, DocNumero, DocAnulado From Documento Where DocCodigo = " & RsAux!EnvDocumentoFactura
'                                Set RsDifEnvio = cBase.OpenResultset(Cons, rdOpenForwardOnly)
'                                If Not RsDifEnvio.EOF Then
'                                    If RsDifEnvio("DocAnulado") Then
'                                        MsgBox "ATENCIÓN!!! El envío Nº " & RsAux("EnvCodigo") & " tiene el documento de cobro anulado, verifique que se haya cobrado el envío al cliente.", vbExclamation, "POSIBLE ERROR"
'                                    Else
'                                        .Cell(flexcpText, .Rows - 1, 8) = RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero
'                                    End If
'                                End If
'                                RsDifEnvio.Close
'                            End If
'                            If Not IsNull(RsAux!EnvValorFlete) Then cCobrar = RsAux!EnvValorFlete
'                        End If
'                    End If
'
'                    If Not IsNull(RsAux!EnvReclamoCobro) Then
'                        cCobrar = cCobrar + RsAux!EnvReclamoCobro
'
'                        'Si es vta. telefonica busco el código.
'                        If RsAux!EnvFormaPago <> TipoPagoEnvio.PagaAhora Then
'                            'Si el join devuelve algo es porque es una vta. telefónica.
'                            Cons = "Select DocSerie, DocNumero From Documento, VentaTelefonica Where DocCodigo = " & RsAux!EnvDocumento _
'                                 & " And DocCodigo = VTeDocumento And DocAnulado = 0"
'                            Set RsDifEnvio = cBase.OpenResultset(Cons, rdOpenForwardOnly)
'                            If Not RsDifEnvio.EOF Then
'                                If InStr(1, .Cell(flexcpText, .Rows - 1, 8), RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero) = 0 Then
'                                    If Trim(.Cell(flexcpText, .Rows - 1, 8)) = vbNullString Then
'                                        .Cell(flexcpText, .Rows - 1, 8) = RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero
'                                    Else
'                                        .Cell(flexcpText, .Rows - 1, 8) = Trim(.Cell(flexcpText, .Rows - 1, 8)) & ", " & RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero
'                                    End If
'                                End If
'                            End If
'                            RsDifEnvio.Close
'                        End If
'                    End If

                End If
                
                'Cobro camión.
                If Not IsNull(RsAux!EnvValorFlete) And RsAux!EnvFormaPago = TipoPagoEnvio.FacturaCamión Then cFacCamion = RsAux!EnvValorFlete
                            
                'Pongo en el data 9 lo pendiente a cobrar
                .Cell(flexcpData, .Rows - 1, 9) = cImpDocPend
                            
                aModificacion = Trim(BuscoSignoMoneda(RsAux!EnvMoneda))
                If cLiquidar > 0 Then
                    .Cell(flexcpText, .Rows - 1, 4) = Trim(aModificacion) & " " & Format(cLiquidar, FormatoMonedaP)
                Else
                    .Cell(flexcpText, .Rows - 1, 4) = Trim(aModificacion) & " " & "0.00"
                End If
                .Cell(flexcpData, .Rows - 1, 4) = cLiquidar
                
                If cCobrar > 0 Then
                    .Cell(flexcpText, .Rows - 1, 5) = Trim(aModificacion) & " " & Format(cCobrar, FormatoMonedaP)
                Else
                    .Cell(flexcpText, .Rows - 1, 5) = Trim(aModificacion) & " " & "0.00"
                End If
                .Cell(flexcpText, .Rows - 1, 6) = Trim(aModificacion) & " " & Format(cFacCamion, FormatoMonedaP)
                
                Dim iMoneda As Long
                iMoneda = RsAux!EnvMoneda: .Cell(flexcpData, .Rows - 1, 2) = iMoneda
                .Cell(flexcpData, .Rows - 1, 3) = cCobrar
                
                AgregoQueLiquido TipoQueCobro.EnviosMerc, RsAux!EnvCodigo, cCobrar, cLiquidar, 0, 0, cImpDocPend

            End If
            
            If Not IsNull(RsAux!PCaEnvio) Then
                
                'If arrPendiente Is Nothing Then ReDim arrPendiente(0)
                
                ReDim Preserve arrPendiente(UBound(arrPendiente) + 1)
                With arrPendiente(UBound(arrPendiente))
                    .Envio = RsAux!EnvCodigo
                    .Pendiente = RsAux!PCaID
                    .Moneda = RsAux!DisMoneda
                    .Importe = RsAux!PCaImporte
                    .Activo = True
                End With
                cCobrar = RsAux!DisMoneda
                If .Cell(flexcpData, .Rows - 1, 8) = "" Then
                    .Cell(flexcpData, .Rows - 1, 8) = cCobrar
                    .Cell(flexcpData, .Rows - 1, 7) = arrPendiente(UBound(arrPendiente)).Importe
                    If RsAux!DisMoneda <> RsAux!EnvMoneda Then
                        aModificacion = Trim(BuscoSignoMoneda(RsAux!DisMoneda))
                    End If
                    .Cell(flexcpText, .Rows - 1, 7) = aModificacion & " " & Format(arrPendiente(UBound(arrPendiente)).Importe, "#,##0.00")
                Else
                    If .Cell(flexcpData, .Rows - 1, 8) = cCobrar Then
                        .Cell(flexcpData, .Rows - 1, 7) = CCur(.Cell(flexcpData, .Rows - 1, 7)) + RsAux!PCaImporte
                        .Cell(flexcpText, .Rows - 1, 7) = aModificacion & " " & Format(CCur(.Cell(flexcpData, .Rows - 1, 7)), "#,##0.00")
                    Else
                        'Tengo que hacer tasa de cambio para esta moneda
                        .Cell(flexcpText, .Rows - 1, 7) = "Varias Monedas"
                    End If
                End If
                
            Else
                .Cell(flexcpData, .Rows - 1, 7) = "0.00"
            End If
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close

End Sub
Private Function CargoServicios() As String
On Error GoTo errCSE
Dim aModificacion As String
Dim aValor As Currency
Dim sInserto As Boolean
Dim RsPen As rdoResultset
Dim sErr As String

    'Primero Cargo los envíos
    CargoServicios = ""
    
    'Cargo servicios cumplidos
    Cons = "Select * From Servicio, ServicioVisita " _
            & "Left Outer Join TipoFlete On VisTipoFlete = TFlCodigo" _
        & " , Producto, Cliente " _
        & " Where VisCamion = " & cCamion.ItemData(cCamion.ListIndex) _
        & " And VisLiquidada Is Null And  SerCodigo = VisServicio And SerProducto = ProCodigo And ProCliente = CliCodigo " '_
        '& " Order by VisServicio "
    
    sErr = "Hago consulta"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Dim iRow As Integer
    iRow = vsConsulta.Rows - 1
    
    Do While Not RsAux.EOF
        sInserto = False
        
        With vsConsulta
            
            'Aca tengo que agregar los retiros que no esten en retiros o anulados, visitas que no esten en visitas  o anuladas y entregas que esten cumplidas o anuladas.
            Select Case RsAux!VisTipo
                Case TipoServicio.Retiro
                    If RsAux!SerEstadoServicio <> EstadoS.Retiro Or RsAux!VisSinEfecto = 1 Then sInserto = True
                Case TipoServicio.Entrega
                    If RsAux!SerEstadoServicio <> EstadoS.Entrega Or RsAux!VisSinEfecto = 1 Then sInserto = True
                Case TipoServicio.Visita
                    If RsAux!SerEstadoServicio <> EstadoS.Visita Or RsAux!VisSinEfecto = 1 Then sInserto = True
            End Select
            
            If sInserto And RsAux("SerEstadoServicio") <> EstadoS.Anulado Then
                Select Case RsAux!VisTipo
                    Case TipoServicio.Entrega: .AddItem "Ent"
                    Case TipoServicio.Retiro: .AddItem "Ret"
                    Case TipoServicio.Visita: .AddItem "Vis"
                End Select
                
                sErr = "Servicio: " & RsAux("SerCodigo") & " paso cargo Data SerModificacion y VisFModificacion"
                .Cell(flexcpText, .Rows - 1, 7) = "0.00"
                .Cell(flexcpData, .Rows - 1, 0) = 2 'Me digo que es servicio
                .Cell(flexcpText, .Rows - 1, 9) = RsAux!SerModificacion
                .Cell(flexcpText, .Rows - 1, 10) = RsAux!VisFModificacion
                aValor = RsAux!VisCodigo
                .Cell(flexcpData, .Rows - 1, 10) = aValor
                
                sErr = "Servicio: " & RsAux("SerCodigo") & " paso: cargo VisMoneda y VisLiquidarAlCamion"
                If Not IsNull(RsAux!VisMoneda) Then
                    aValor = RsAux!VisMoneda: .Cell(flexcpData, .Rows - 1, 2) = aValor
                    aValor = RsAux!VisLiquidarAlCamion: .Cell(flexcpData, .Rows - 1, 4) = aValor
                    .Cell(flexcpText, .Rows - 1, 4) = BuscoSignoMoneda(RsAux!VisMoneda) & " " & Format(RsAux!VisLiquidarAlCamion, FormatoMonedaP)
                Else
                    aValor = 0: .Cell(flexcpData, .Rows - 1, 4) = 0
                End If
                
                .Cell(flexcpText, .Rows - 1, 1) = RsAux!SerCodigo
                sErr = "Servicio: " & RsAux("SerCodigo") & " paso: ArmoDirección ProDireccion o CliDireccion"
                If Not IsNull(RsAux!ProDireccion) Then
                    .Cell(flexcpText, .Rows - 1, 2) = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!ProDireccion)
                Else
                    .Cell(flexcpText, .Rows - 1, 2) = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion)
                End If
                
                sErr = "Servicio: " & RsAux("SerCodigo") & " paso: TipoFlete TFLNombreCorto para VisTipoFlete"
                If Not IsNull(RsAux!VisTipoFlete) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!TFlNombreCorto)
                
                
                If RsAux!VisSinEfecto Then .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = &HD0D0FF
                                
                Dim curDocPend  As Currency
                curDocPend = 0
                'Si es una entrega verifico si tiene en la tabla servicio el SerDocumento.
                If RsAux!VisTipo = TipoServicio.Entrega Then
                    If Not RsAux!VisSinEfecto Then
                        '20111102
                        'SI ES UN ENTREGA SI O SI LEVANTO LOS PENDIENTES x si me quedó alguno colgado.
                        'tengo documento asociado veo si esta pendiente.
                        Cons = fnc_GetStringDocumentosPendientes(RsAux("SerCodigo"), 2)
                        Set RsPen = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        If Not RsPen.EOF Then
                            'Método nuevo
                            .Cell(flexcpText, .Rows - 1, 5) = BuscoSignoMoneda(RsPen!DPeMoneda)
                            Do While Not RsPen.EOF
                                .Cell(flexcpText, .Rows - 1, 8) = IIf(.Cell(flexcpText, .Rows - 1, 8) <> "", .Cell(flexcpText, .Rows - 1, 8) & ", ", "") & RsPen!DocSerie & " " & RsPen!DocNumero
                                curDocPend = curDocPend + RsPen("DPeImporte")
                                RsPen.MoveNext
                            Loop
                            .Cell(flexcpText, .Rows - 1, 5) = .Cell(flexcpText, .Rows - 1, 5) & " " & Format(curDocPend, FormatoMonedaP)
                            .Cell(flexcpData, .Rows - 1, 3) = curDocPend
                        Else
                            .Cell(flexcpText, .Rows - 1, 5) = "0"
                            .Cell(flexcpData, .Rows - 1, 3) = 0
                            .Cell(flexcpData, .Rows - 1, 3) = 0
                        End If
                        RsPen.Close
                        
'                        Cons = "Select * From DocumentoPendiente, Documento " _
'                            & " Where DPeDocumento = " & RsAux!SerDocumento & " And DPeDocumento = DocCodigo AND DocAnulado = 0"
'                        Set RsPen = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
'                        If Not RsPen.EOF Then
'                            sErr = "Servicio: " & RsAux("SerCodigo") & " paso: DocumentoPendiente donde SerDocumento = DPeDocumento leo DPeMoneda, DPeImporte, DocSerie y DocNumero"
'                            .Cell(flexcpText, .Rows - 1, 5) = BuscoSignoMoneda(RsPen!DPeMoneda) & " " & Format(RsPen!DPeImporte, FormatoMonedaP)
'                            .Cell(flexcpText, .Rows - 1, 8) = RsPen!DocSerie & " " & RsPen!DocNumero
'
'                            curDocPend = RsPen!DPeImporte
'                            .Cell(flexcpData, .Rows - 1, 3) = curDocPend
'
'                        Else
'                            .Cell(flexcpText, .Rows - 1, 5) = "0"
'                            .Cell(flexcpData, .Rows - 1, 3) = 0 '
'                            .Cell(flexcpData, .Rows - 1, 3) = 0
'                        End If
'                        RsPen.Close
                    Else
                        .Cell(flexcpText, .Rows - 1, 5) = "0.00"
                    End If
                    
                Else
                
                    'ACA SI O SI TENGO QUE TENER EL VISDOCUMENTO.
                    If Not IsNull(RsAux!VisDocumento) Then
                        'tengo documento asociado veo si esta pendiente.
                        Cons = "Select * From DocumentoPendiente, Documento " _
                            & " Where DPeDocumento = " & RsAux!VisDocumento & " And DPeDocumento = DocCodigo AND DocAnulado = 0 AND DPeIDLiquidacion IS NULL"
                        Set RsPen = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                        If Not RsPen.EOF Then
                            sErr = "Servicio: " & RsAux("SerCodigo") & " no entrega paso: DocumentoPendiente donde SerDocumento = DPeDocumento leo DPeMoneda, DPeImporte, DocSerie y DocNumero"
                            .Cell(flexcpText, .Rows - 1, 5) = BuscoSignoMoneda(RsPen!DPeMoneda) & " " & Format(RsPen!DPeImporte, FormatoMonedaP)
                            .Cell(flexcpText, .Rows - 1, 8) = RsPen!DocSerie & " " & RsPen!DocNumero
                            curDocPend = RsPen!DPeImporte: .Cell(flexcpData, .Rows - 1, 3) = curDocPend
                        Else
                            .Cell(flexcpText, .Rows - 1, 5) = "0"
                            .Cell(flexcpData, .Rows - 1, 3) = 0 '
                        End If
                        RsPen.Close
                    Else
                        .Cell(flexcpText, .Rows - 1, 5) = sSignMonedaPesos
                        curDocPend = 0: .Cell(flexcpData, .Rows - 1, 3) = curDocPend
                    End If
                End If
                
                sErr = "Servicio: " & RsAux("SerCodigo") & "Cargo VisFormaPago, VisCosto, VisMoneda "
                If Not IsNull(RsAux!VisFormaPago) Then
                    If RsAux!VisFormaPago = 2 Then
                        .Cell(flexcpText, .Rows - 1, 6) = BuscoSignoMoneda(RsAux!VisMoneda) & " " & Format(RsAux!VisCosto, FormatoMonedaP)
                    Else
                        .Cell(flexcpText, .Rows - 1, 6) = BuscoSignoMoneda(RsAux!VisMoneda) & " 0.00"
                    End If
                End If
                sErr = ""
                
                AgregoQueLiquido ServicioEntRet, RsAux!SerCodigo, curDocPend, CCur(.Cell(flexcpData, .Rows - 1, 4)), 0, 0, curDocPend
                
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If iRow <> vsConsulta.Rows - 1 Then
        With vsConsulta
            .Select iRow + 1, 1, .Rows - 1, 1
            .Sort = flexSortGenericAscending
            .Select 0, 1
        End With
    End If
    
    Status.SimpleText = ""
    Exit Function
errCSE:
    CargoServicios = "(Servicios)" & sErr & " Desc: " & Err.Description
End Function

Private Sub InsertoEnTablaDetalle(ByVal lMoneda As Long, ByVal cLiquidar As Currency, iAlta As Integer)
Dim iCont As Integer
Dim bInsert As Boolean
Dim cAux As Currency

    bInsert = False
    If cLiquidar = 0 Then Exit Sub
    'Busco el monto en la grilla.
    For iCont = 1 To vsDetMonto.Rows - 1
        cAux = vsDetMonto.Cell(flexcpData, iCont, 1)
        
        If cAux = cLiquidar _
            And Val(vsDetMonto.Cell(flexcpData, iCont, 0)) = lMoneda Then
            
            If iAlta <> -1 Then
                vsDetMonto.Cell(flexcpText, iCont, 2) = Val(vsDetMonto.Cell(flexcpValue, iCont, 2)) + 1
                vsDetMonto.Cell(flexcpText, iCont, 4) = Format(CCur(vsDetMonto.Cell(flexcpValue, iCont, 4)) + cLiquidar, "#,##0.00")
            Else
                If Val(vsDetMonto.Cell(flexcpValue, iCont, 2)) = 1 Then
                    vsDetMonto.RemoveItem iCont
                Else
                    vsDetMonto.Cell(flexcpText, iCont, 2) = Val(vsDetMonto.Cell(flexcpValue, iCont, 2)) - 1
                    vsDetMonto.Cell(flexcpText, iCont, 4) = Format(CCur(vsDetMonto.Cell(flexcpValue, iCont, 4)) - cLiquidar, "#,##0.00")
                End If
            End If
            bInsert = True
            Exit For
        End If
        
    Next iCont
    
    If Not bInsert Then
        With vsDetMonto
            .AddItem "Reparto"
            .Cell(flexcpData, .Rows - 1, 0) = lMoneda
            .Cell(flexcpData, .Rows - 1, 1) = cLiquidar
            .Cell(flexcpText, .Rows - 1, 2) = 1
            .Cell(flexcpText, .Rows - 1, 3) = BuscoSignoMoneda(lMoneda) & "  " & Format(cLiquidar, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 4) = Format(cLiquidar, "#,##0.00")
            
            .Select 1, 1, .Rows - 1, 1
            .Sort = flexSortNumericAscending
        End With
    End If
    
    ArmoDetalleLiquidacion
'    s_SetTotalBoleta
    
End Sub

Sub PagoLiquidacion(ByVal IDLiquidacion As Long, ByVal NroPago As Byte, ByVal Moneda As Integer, ByVal Importe As Currency, ByVal Usuario As Long)
    
    If Importe = 0 Then Exit Sub
    If NroPago = 0 Then
        'Busco el mayor.
        Dim sQy As String
        sQy = "SELECT IsNull(Max(LCoNro), 0) From LiquidacionCobro WHERE LCoLiquidacion = " & IDLiquidacion
        Dim rsNP As rdoResultset
        Set rsNP = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
        If Not rsNP.EOF Then NroPago = rsNP(0) + 1 Else NroPago = 1
        rsNP.Close
    End If
    
    cBase.Execute "INSERT INTO LiquidacionCobro ([LCoLiquidacion] ,[LCoNro],[LCoFecha],[LCoMoneda],[LCoCobrado],[LCoUsrLiquid]) " & _
            " Values (" & IDLiquidacion & "," & NroPago & ",'" & Format(gFechaServidor, "yyyyMMdd HH:nn:ss") & "', " & Moneda & ", " & Importe & ", " & Usuario & ")"
End Sub

Private Sub AccionGrabar()
Dim Msg As String, J As Integer, aJ As Integer, Hora As String
Dim lUIDSuc As Long, lUIDAut As Long
Dim sDefSuc As String, sDescSuc As String
Dim sIDNro As String
Dim sErrAtrapado As String
Dim QueMontoCobro As Currency

    If Not tComentario.Enabled Then Exit Sub
    
    If GetImportePagoTotal = 0 And vsConsulta.Rows = 1 Then
        MsgBox "Ud. no hizo ningún movimiento para poder almacenar.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    If vsLiquidacion.Rows = 1 Then MsgBox "No hay datos en la consulta.", vbExclamation, "ATENCIÓN": Exit Sub
    If Val(tUsuario.UserID) = 0 Then MsgBox "Ingrese su dígito de usuario.", vbExclamation, "ATENCIÓN": tUsuario.SetFocus: Exit Sub
    
'    ValidoDisponibilidadMoneda
    'Valido si hay liquidaciones pendientes.
'    If Not fnc_ValidoLiquidacionesPendientes() Then Exit Sub
    
    Dim cAZureo As Currency
    
    If Not fnc_EsZureoCGSA Then
        MsgBox "Su PC no tiene bien configurado ZUREO no se puede continuar.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If Not fnc_DoTestLogin(False) Then
        MsgBox "Hay que hacer movimientos en Zureo y no retornó conexión, no puede continuar.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If

    If MsgBox("¿Confirma grabar la liquidación?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    With vsPagaCon
        For i = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpData, i, 0)) > 0 Then
                'PASO X ZUREO las transferencias.
                If Not fnc_DoTestLogin(True) Then
                    MsgBox "Existen transferencias y no hay conexión a zureo, reintente.", vbExclamation, "Sin acceso a zureo"
                    Exit Sub
                End If
                '---------------------------------------------
                Exit For
            End If
        Next
    End With
    
    
    Dim iGRP As Integer
    With vsPagaCon
        For i = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpData, i, 0)) = -3 Then
                iGRP = iGRP + 1
            End If
        Next
    End With
    
    If iGRP > 0 Then
    'Veo si tengo aportes de comisión.  paMCCtaComisionRP
        If GetImportePago(paMCCtaComisionRP) = 0 Then
            If MsgBox("POSIBLE ERROR, se ingresó al menos un giro de redpagos y no se asignó ninguna COMISIÓN." & vbCrLf & vbCrLf & "¿Desea continuar de todas formas?", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then
                Exit Sub
            End If
        ElseIf GetCantidadRenglonesPago(paMCCtaComisionRP) <> iGRP Then
        
            If MsgBox("POSIBLE ERROR, la cantidad de GIROS y COMISIONES RedPagos no son iguales." & vbCrLf & vbCrLf & "¿Desea continuar de todas formas?", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then
                Exit Sub
            End If
        
        End If
    End If
    tComentario.Enabled = False
    Screen.MousePointer = 11
    FechaDelServidor
    
    Dim idZureoDiferencia As Long, idZureoDepositoEfectivo As Long, idZureoDepositoComision As Long
    bGrabar.Enabled = False
    cBase.BeginTrans
    On Error GoTo ErrRB
    
    Dim lnCodEntrada As Long, lMPesos As Long
    lMPesos = 0
    lnCodEntrada = 0
    'Primero inserto la liquidación y le asigno a cada envío el nro. de liquidación.
    'Saco el mayor código de Salida de Caja.
    'Puedo tener distintas monedas entonces inserto para cada moneda y
    'en la misma lista le inserto el código de salida de caja
    Dim cImpMC As Currency
    Dim bInsert As Boolean
    
    'Inserto en tabla LiquidacionCamiones la misma
    Cons = "Select * from Liquidacion Where LiqID = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!LiqFecha = Format(Now, "yyyy/mm/dd hh:nn")
    RsAux!LiqTipo = 1
    RsAux!LiqEnte = cCamion.ItemData(cCamion.ListIndex)
    If lMPesos > 0 Then RsAux!LiqMovDisponibilidad = lMPesos
    If GetImporteReclamar + GetImportePendienteCaja(True) - GetImportePendienteCaja(False) > 0 Then
        'El total SIEMPRE es de los envíos
        RsAux("LiqTotal") = GetImporteReclamar + GetImportePendienteCaja(True) - GetImportePendienteCaja(False)
    Else
        RsAux("LiqTotal") = 0
    End If
    RsAux.Update
    RsAux.Close

    Cons = "Select Max(LiqID) from Liquidacion Where LiqEnte = " & cCamion.ItemData(cCamion.ListIndex) & " And LiqTipo = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    If Not IsNull(RsAux(0)) Then lnCodEntrada = RsAux(0) Else lnCodEntrada = 0
    RsAux.Close
    '.....................................................................
   
    Dim cSaldoReclamo As Currency
    Dim cSaldoAporte As Currency
    
    If GetImporteReclamarDiferencias <> 0 Or GetImporteAporteDiferencias <> 0 Then
    
        'Tengo que validar cuanto se salda.
        'Condiciones:
        '   el reclamo es menor o igual que PAGA + Pendientes --> aca saldo todos los reclamos.
        '   el reclamo es mayor que PAGA + Pendientes --> saldo hasta que cubra los valores.
        If GetImporteReclamarDiferencias <= GetImportePagoTotal + GetImporteAporteDiferencias Then
            cSaldoReclamo = GetImporteReclamarDiferencias
        Else
            cSaldoReclamo = GetImportePagoTotal + GetImporteAporteDiferencias
        End If
        
        '   el aporte es mayor que los (envíos + pendientes) - Paga --> saldo importe que cubre  (paga - (envíos + pendientes)
        '   el aporte es menor que los (envíos + pendientes) - Paga --> saldo todo.
        If GetImporteAporteDiferencias > 0 And GetImporteAporteDiferencias > (GetImporteReclamar + GetImportePendienteCaja(True) - GetImportePendienteCaja(False)) - GetImportePagoTotal Then
            If (GetImporteReclamar + GetImporteReclamarDiferencias + GetImportePendienteCaja(True) - GetImportePendienteCaja(False)) - GetImportePagoTotal < 0 Then
                'Deja más plata entonces no utilizo este dinero.
            Else
                cSaldoAporte = (GetImporteReclamar + GetImporteReclamarDiferencias + GetImportePendienteCaja(True) - GetImportePendienteCaja(False)) - GetImportePagoTotal
            End If
        Else
            cSaldoAporte = GetImporteAporteDiferencias
            'Tomo el aporte como pago de esta liquidación por lo tanto esto se lo tengo que restar al ingreso de efectivo.
        End If
        
    End If
    
    'Puedo tener un nuevo pendiente, las condiciones son:
    '1) que falte plata (sí o sí tengo que tener a reclamar)
    '2) que sobre plata.
    If fnc_HayEnviosEnLiquidacion Then
        
        Dim cInsert As Currency
        cInsert = 0
        If GetImporteAporteDiferencias = 0 And GetImporteReclamarDiferencias = 0 Then
        
            'No tengo pendientes procedo derecho (tomo el saldo)  .
'            cInsert = GetImportePagoTotal - (GetImporteReclamar + GetImportePendienteCaja(True) - GetImportePendienteCaja(False))
'            If cInsert >= 0 Then
'                cInsert = GetImportePagoTotal
'            Else
'                cInsert = (GetImporteReclamar + GetImportePendienteCaja(True) - GetImportePendienteCaja(False))
'            End If
            cInsert = GetImportePagoTotal
            
        Else
            
            'Tengo que validar si lo que paga cubre el total del pendiente.
            If GetImportePagoTotal + GetImporteAporteDiferencias <= GetImporteReclamarDiferencias Then
                'SOLO LO QUE PAGA (queda pendiente).
                cInsert = 0
            Else
                'Aquí tengo que volcar el reclamo.
                'SI TENGO APORTE POR DIFERENCIAS TENGO QUE VER CUANTO UTILIZO DE EL.
                If cSaldoAporte > 0 And cSaldoAporte < GetImporteAporteDiferencias Then
                    'Utilizo parte del saldo anterior para pagar esta liquidación
                    cInsert = GetImportePagoTotal + cSaldoAporte - GetImporteReclamarDiferencias
                Else
                    If (GetImporteReclamar + GetImportePendienteCaja(True) - GetImportePendienteCaja(False)) <= GetImportePagoTotal - GetImporteReclamarDiferencias Then
                        cInsert = GetImportePagoTotal - GetImporteReclamarDiferencias
                    Else
                        cInsert = GetImportePagoTotal + GetImporteAporteDiferencias - GetImporteReclamarDiferencias
                    End If
                End If
                
            End If
            
        End If
        
'        If cInsert <> 0 Then
            'Aca tengo que incluir el pendiente.
            'ES EL primero así que ya paso el ID
            PagoLiquidacion lnCodEntrada, 1, paMonedaPesos, cInsert, tUsuario.UserID
'        End If
        
    Else
        
        'Tengo que grabar lo que pague de más.
        If GetImporteReclamarDiferencias = 0 Or GetImportePagoTotal > GetImporteReclamarDiferencias Then '  GetImporteAporteDiferencias > GetImporteReclamarDiferencias Then
            'SI NO TENGO NADA A RECLAMAR O LO QUE RECLAMO ES MENOR QUE LO QUE PAGO CON EL DINERO DEL EFECTIVO --> es un nuevo aporte.
            PagoLiquidacion lnCodEntrada, 1, paMonedaPesos, GetImportePagoTotal - GetImporteReclamarDiferencias, tUsuario.UserID
        End If
    End If
    
'Asigno Pagos de redpagos
    Dim oAsignados As New clsDesicionConQuePaga
    Set oAsignados.ConQuePaga = New Collection
    Dim oCQC As clsConQuePaga
    With vsPagaCon
        For i = .FixedRows To .Rows - 1
            
            If Val(.Cell(flexcpData, i, 0)) = -3 Then
                
                'Agrego otra transacción a la colección.
                Set oCQC = New clsConQuePaga
                oCQC.IDDocumentoPaga = Val(.Cell(flexcpData, i, 1))
                oCQC.Importe = CCur(.Cell(flexcpValue, i, 2))
                oCQC.TipoConQuePaga = GiroRedPagos
                
                oAsignados.ConQuePaga.Add oCQC
                
            ElseIf Val(.Cell(flexcpData, i, 0)) = -1 Then
                
                Set oCQC = New clsConQuePaga
                oCQC.IDDocumentoPaga = .Cell(flexcpData, i, 2)
                oCQC.Importe = .Cell(flexcpValue, i, 2)
                oCQC.TipoConQuePaga = ChequeADepositar
                oAsignados.ConQuePaga.Add oCQC
            
            ElseIf Val(.Cell(flexcpData, i, 0)) = -4 Then
            
                Set oCQC = New clsConQuePaga
                
                oCQC.IDDocumentoPaga = .Cell(flexcpData, i, 1)
                oCQC.Importe = .Cell(flexcpValue, i, 2)
                oCQC.TipoConQuePaga = NotaDevolucion
                
                oAsignados.ConQuePaga.Add oCQC
            
            End If
        Next
    End With
        
    Dim rsD As rdoResultset
    With vsConsulta
        
        For i = 1 To .Rows - 1
            
            If vsConsulta.Cell(flexcpBackColor, i, 0) <> cte_CelesteClaro Then
                    
                If .Cell(flexcpData, i, 0) = 2 Then
                    'TABLA SERVICIO
                    Msg = "Servicio " & Val(.Cell(flexcpText, i, 1))
                    
                    Dim bDocPendServ As Byte, DocRetiro As Long
                    bDocPendServ = 0
                    
                    Cons = "Select * From Servicio Where SerCodigo = " & Val(.Cell(flexcpText, i, 1))
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsAux.EOF Then
                        Msg = "Otra terminal elimino el servicio = " & Val(.Cell(flexcpText, i, 1))
                        RsAux.Close: RsAux.Edit 'Provoco error.
                    Else
                        If RsAux!SerModificacion = CDate(.Cell(flexcpText, i, 9)) Then
                            
                            If Not IsNull(RsAux("SerDocumento")) Then bDocPendServ = 2
                            
                            RsAux.Edit
                            RsAux!SerModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:nn:00")
                            Hora = Format(RsAux!SerModificacion, "dd/mm/yyyy hh:mm:00")
                            RsAux.Update
                            RsAux.Close
                            aJ = i + 1
                            For J = aJ To .Rows - 1
                                If Val(.Cell(flexcpText, J, 1)) = Val(.Cell(flexcpText, i, 1)) Then
                                    .Cell(flexcpText, J, 9) = Hora
                                End If
                            Next J
                        Else
                            Msg = "Otra terminal modifico el servicio = " & Val(.Cell(flexcpText, i, 1))
                            RsAux.Close: RsAux.Edit 'Provoco error.
                        End If
                    End If
                    
                    Cons = "Select * From ServicioVisita Where VisCodigo = " & Val(.Cell(flexcpData, i, 10))
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsAux.EOF Then
                        Msg = "Otra terminal elimino los datos  de la tabla de ServicioVisita para el servicio " & Val(.Cell(flexcpText, i, 1))
                        RsAux.Close: RsAux.Edit 'Provoco error.
                    Else
                        If Not IsNull(RsAux("VisDocumento")) And RsAux!VisTipo = TipoServicio.Retiro And bDocPendServ = 0 Then
                            bDocPendServ = 1
                        Else
                            'Si tiene documento asociado o es entrega ahí busco los documentos pendientes.
                            If Not IsNull(RsAux("VisDocumento")) And RsAux!VisTipo <> TipoServicio.Retiro Then
                                bDocPendServ = 2
                                DocRetiro = RsAux("VisDocumento")
                            End If
                            
                        End If
                    
                        RsAux.Edit
                        RsAux!VisFModificacion = Format(gFechaServidor, sqlFormatoFH)
                        RsAux!VisLiquidada = lnCodEntrada
                        RsAux.Update
                        RsAux.Close
                        
                    End If
                    
                    If bDocPendServ > 0 Then
                        
                        If bDocPendServ = 2 Then
                            Cons = fnc_GetStringDocumentosPendientes(Val(.Cell(flexcpText, i, 1)), 2)
                        Else
                            Cons = "SELECT * FROM DocumentoPendiente WHERE DPeDocumento = " & DocRetiro
                        End If
                        Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        QueMontoCobro = 0
                        Do While Not rsD.EOF
                            QueMontoCobro = QueMontoCobro + rsD("DPeImporte")
                            rsD.Edit
                            rsD("DPeIDLiquidacion") = lnCodEntrada
                            rsD("DPeFLiquidacion") = Format(Now, "yyyy/mm/dd hh:nn:ss")
                            rsD.Update
                            rsD.MoveNext
                        Loop
                        rsD.Close
                        If QueMontoCobro <> 0 Then
                            If (BuscarQueCobroServicio(ServicioEntRet, Val(.Cell(flexcpText, i, 1)), CCur(.Cell(flexcpData, i, 3))).ImporteDocumentoPendiente <> QueMontoCobro) Then
                                rsD.Edit
                                sErrAtrapado = "Para el servicio " & Val(.Cell(flexcpText, i, 1)) & " se asigno en pendiente " & BuscarQueCobro(ServicioEntRet, Val(.Cell(flexcpData, i, 10))).ImporteDocumentoPendiente & " y se está cobrando " & QueMontoCobro
                            End If
                        End If
                        
                    End If
                    
                ElseIf .Cell(flexcpData, i, 0) = 3 Then
                    
                    'Trabla Traslados
                    Cons = "SELECT * " & _
                        "FROM Traspaso WHERE TraCodigo = " & Val(.Cell(flexcpText, i, 1))
                        
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsAux.EOF Then
                        Msg = "No encontré el traslado " & .Cell(flexcpText, i, 1) & ", verifique si no fue eliminado."
                        RsAux.Close
                        RsAux.Edit
                    Else
                        If RsAux("TraFModificacion") <> CDate(.Cell(flexcpText, i, 9)) Then
                            Msg = "El traslado fué editado por otro usuario."
                            RsAux.Close
                            RsAux.Edit
                        ElseIf Not IsNull(RsAux("TraAnulado")) Then
                            Msg = "El traslado fué anulado por otro usuario."
                            RsAux.Close
                            RsAux.Edit
                        Else
                            RsAux.Edit
                            If Not IsNull(RsAux("TraComentario")) Then
                                RsAux("TraComentario") = Trim(RsAux("TraComentario")) & "[L:" & lnCodEntrada & "]"
                            Else
                                RsAux("TraComentario") = "[L:" & lnCodEntrada & "]"
                            End If
                            RsAux("TraFModificacion") = Format(Now, "yyyy/mm/dd hh:nn:ss")
                            RsAux.Update
                        End If
                    End If
                    RsAux.Close
                    
                ElseIf .Cell(flexcpData, i, 0) = 4 Then
                    
                    'Liquidación pendiente.
                    'Voy tomando lo que tengo del pago y lo voy asignando a los pendientes.
                    Dim sQy As String
                    sQy = "SELECT LiqID, LiqTotal, IsNull(SUM(LCoCobrado), 0) Cobrado FROM Liquidacion " & _
                        "LEFT OUTER JOIN LiquidacionCobro ON LCoLiquidacion = LiqID " & _
                        "WHERE LiqID = " & vsConsulta.Cell(flexcpText, i, 1) & _
                        " GROUP BY LiqID, LiqTotal"
                    Dim rsPend As rdoResultset
                    Set rsPend = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
                    If rsPend.EOF Then
                        sErrAtrapado = "ATENCIÓN!!!" & vbCrLf & "El pendiente " & vsConsulta.Cell(flexcpText, i, 1) & " no existe, verifique."
                        rsPend.Close
                        rsPend.Edit
                    Else
                        If CCur(vsConsulta.Cell(flexcpText, i, 5)) <> CCur(Format(rsPend("LiqTotal") - rsPend("Cobrado"), "#,##0.00")) Then
                            sErrAtrapado = "ATENCIÓN!!!" & vbCrLf & "El pendiente " & vsConsulta.Cell(flexcpText, i, 1) & " fue modificado, verifique."
                            rsPend.Close
                            rsPend.Edit
                        Else
                            If rsPend("LiqTotal") - rsPend("Cobrado") > 0 Then
                                If CCur(Format(rsPend("LiqTotal") - rsPend("Cobrado"), "#,##0.00")) >= cSaldoReclamo Then
                                    PagoLiquidacion vsConsulta.Cell(flexcpText, i, 1), 0, CInt(paMonedaPesos), cSaldoReclamo, tUsuario.UserID
                                    cSaldoReclamo = 0
                                Else
                                    PagoLiquidacion vsConsulta.Cell(flexcpText, i, 1), 0, CInt(paMonedaPesos), rsPend("LiqTotal") - rsPend("Cobrado"), tUsuario.UserID
                                    cSaldoReclamo = cSaldoReclamo - (rsPend("LiqTotal") - rsPend("Cobrado"))
                                End If
                            Else
                                If ((rsPend("LiqTotal") - rsPend("Cobrado")) * -1) >= cSaldoAporte Then
                                    PagoLiquidacion vsConsulta.Cell(flexcpText, i, 1), 0, CInt(paMonedaPesos), cSaldoAporte * -1, tUsuario.UserID
                                    cSaldoAporte = 0
                                Else
                                    PagoLiquidacion vsConsulta.Cell(flexcpText, i, 1), 0, CInt(paMonedaPesos), rsPend("LiqTotal") - rsPend("Cobrado"), tUsuario.UserID
                                    cSaldoAporte = cSaldoAporte - ((rsPend("LiqTotal") - rsPend("Cobrado")) * -1)
                                End If

                            End If
                        
                            
                        End If
                    End If
                    
                Else
                    'TABLA Envío
                    '19/9/2011
                    'si el estado es 3 entonces lo tengo que dar por cumplido.
                    'para evitar lios con la fecha de edición lo hago posterior a este update.
                    
                    Dim idCodImpresion As Long
                    
                    Msg = "Envío " & Val(.Cell(flexcpText, i, 1))
                    
                    'Ya que lo puedo dar x entregado tengo que controlar si es un vacon, ya que al entregar le modifica la fecha a todos.
                    'Para eso sólo valido el ID de impresión y si el envío pertenece a un vacon.
                    
                    Cons = "Select * From Envio Where EnvCodigo = " & Val(.Cell(flexcpText, i, 1))
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsAux.EOF Then
                        Msg = "No se encontro el envío " & .Cell(flexcpText, i, 1) & ", verifique si no fue eliminado."
                        RsAux.Close
                        RsAux.Edit
                    Else
                        idCodImpresion = RsAux("EnvCodImpresion")
                        If RsAux!EnvFModificacion = CDate(.Cell(flexcpText, i, 9)) Then
                            If IsNull(RsAux!EnvLiquidacion) Then
                                RsAux.Edit
                                RsAux!EnvLiquidacion = lnCodEntrada
                                RsAux.Update
                                RsAux.Close
                            Else
                                Msg = "Envíos ya liquidados."
                                RsAux.Close
                                RsAux.Edit
                            End If
                        Else
                            Dim bOK As Boolean
                            'Si es un código de impresión busco si tengo otro envío del vacon con este id de código de liquidación, si es así lo dejo ok.
                            If Val(.Cell(flexcpData, i, 6)) = 3 Then
                                Cons = "SELECT EnvFModificacion, EnvCodigo, EnvEstado " _
                                        & "FROM EnvioVaCon INNER JOIN Envio ON EVCEnvio = EnvCodigo AND EnvCodigo <> " & Val(.Cell(flexcpText, i, 1)) _
                                                            & " AND EnvLiquidacion = " & lnCodEntrada & " AND EnvCodImpresion = " & idCodImpresion _
                                        & " WHERE EVCID = (SELECT EVCID FROM EnvioVaCon WHERE EVCEnvio = " & Val(.Cell(flexcpText, i, 1)) & ")"
                                Dim rsVC As rdoResultset
                                Set rsVC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                                If Not rsVC.EOF Then
                                    bOK = (RsAux("EnvFModificacion") = rsVC("EnvFModificacion"))
                                End If
                                rsVC.Close
                                
                                If bOK Then
                                    RsAux.Edit
                                    RsAux!EnvLiquidacion = lnCodEntrada
                                    RsAux.Update
                                    RsAux.Close
                                End If

                            End If
                            
                            If Not bOK Then
                                Msg = " El envío " & .Cell(flexcpText, i, 1) & " fue modificado por otra terminal, verifique."
                                RsAux.Close
                                RsAux.Edit
                            End If
                        End If
                    End If
                    
'                    [dbo].[prg_RecepcionEnvio_DoyPorEntregadoEnvio] @iCodImpresion int, @iCodEnvio int, @iTipoArticuloServicio int, @iUser int,
'                    @iLocal int, @iTerminal int, @porLiqCamion tinyint = 0
                    If Val(.Cell(flexcpData, i, 6)) = 3 Then
                        Dim rsSP As rdoResultset
                        Debug.Print Val(.Cell(flexcpText, i, 1))
                        Cons = "EXEC prg_RecepcionEnvio_DoyPorEntregadoEnvio " & idCodImpresion & ", " & Val(.Cell(flexcpText, i, 1)) & ", " & paTipoArticuloServicio & ", " & tUsuario.UserID & ", " & paCodigoDeSucursal & ", " & paCodigoDeTerminal & ", 1"
                        Set rsSP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        If rsSP(0) = -1 Then
                            Msg = "Error en envío " & .Cell(flexcpText, i, 1) & " " & rsSP(1)
                            rsSP.Close
                            rsSP.Edit
                        End If
                        rsSP.Close
                        '[dbo].[prg_RecepcionEnvio_DoyPorEntregadoEnvio] @iCodImpresion int, @iCodEnvio int, @iTipoArticuloServicio int, @iUser int,
                        '@iLocal int, @iTerminal int
                    End If
                    
                    If .Cell(flexcpData, i, 8) <> "" Then
                        'Updateo la información de pendiente de caja.
                        Cons = "Update PendientesCaja Set PCaFLiquidacion = '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
                                & ", PCaDetalleLiquidacion = ' Liquidación de camión " & cCamion.Text _
                                & " ', PCaUsrLiquidacion = " & tUsuario.UserID & "  Where PCaEnvio = " & Val(.Cell(flexcpText, i, 1)) & " And PCaFLiquidacion Is Null"
                        cBase.Execute (Cons)
                    End If
                    
                    QueMontoCobro = 0
                    Cons = fnc_GetStringDocumentosPendientes(Val(.Cell(flexcpText, i, 1)), 1)
                    Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    Do While Not rsD.EOF
                        QueMontoCobro = QueMontoCobro + rsD("DPeImporte")
                        rsD.Edit
                        rsD("DPeIDLiquidacion") = lnCodEntrada
                        rsD("DPeFLiquidacion") = Format(Now, "yyyy/mm/dd hh:nn:ss")
                        rsD.Update
                        rsD.MoveNext
                    Loop
                    rsD.Close

                    If BuscarQueCobro(EnviosMerc, Val(.Cell(flexcpText, i, 1))).ImporteDocumentoPendiente <> QueMontoCobro Then
                        rsD.Edit
                        sErrAtrapado = "Para el servicio " & Val(.Cell(flexcpText, i, 1)) & " se asigno en pendiente " & BuscarQueCobro(ServicioEntRet, Val(.Cell(flexcpData, i, 10))).ImporteDocumentoPendiente & " y se está cobrando " & QueMontoCobro
                    End If
                    
                End If
                
            End If
        Next i
    End With
        
    
    If lUIDSuc > 0 Then
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, TipoSuceso.ModificacionDePrecios, paCodigoDeTerminal, lUIDSuc, 0, 0, sDescSuc, Trim(sDefSuc), 0, 0, lUIDAut
    End If
    
    Dim sIDSComprobantes As String
    'DIFERENCIA DE LIQUIDACION EN ZUREO
    'fórmula a zureo -->> ((Reclamo + AReclamar) - Paga) - (Reclamo)
    'Saco el reclamo y hago sólo lo que paga contra lo que debería pagar.
    cAZureo = (GetImporteReclamar + GetImportePendienteCaja(True) - GetImportePendienteCaja(False)) - GetImportePagoTotal
    If cAZureo <> 0 Then
            '(IIf(tipoCompZureo = EntradaCaja, 1, 0))
        idZureoDiferencia = Zureo_InsertoEntradaSalida(IIf(cAZureo < 0, EntradaCaja, SalidaCaja), CStr(lnCodEntrada), "Liquidación de camionero id: " & lnCodEntrada & " camionero: " & cCamion.Text _
                                                    , IIf(cAZureo < 0, iIDSRPesos, paMCDifLiqCamion), Abs(cAZureo), IIf(cAZureo < 0, paMCDifLiqCamion, iIDSRPesos), 0)
        If idZureoDiferencia = 0 Then
            MsgBox "ATENCIÓN!!!" & vbCrLf & "No se obtuvo el id de comprobante que almacena la diferencia de la liquidación, se debe ingresar una " & IIf(cAZureo < 0, "ENTRADA DE CAJA", "SALIDA DE CAJA") & " por $ " & Abs(cAZureo), vbExclamation, "ATENCIÓN"
        Else
            sIDSComprobantes = "Diferencia Liquidación " & idZureoDiferencia & "<BR>"
            Set oCQC = New clsConQuePaga
            oCQC.IDDocumentoPaga = idZureoDiferencia
            oCQC.Importe = Abs(cAZureo)     'EL IMPORTE NO SE GRABA.
            oCQC.TipoConQuePaga = DiferenciaLiquidacion
            
            oAsignados.ConQuePaga.Add oCQC
        End If
    End If
    
    cAZureo = GetImportePagoEfectivo
    If cAZureo <> 0 Then
            '(IIf(tipoCompZureo = EntradaCaja, 1, 0))
        idZureoDepositoEfectivo = Zureo_InsertoEntradaSalida(DepósitoZureo, CStr(lnCodEntrada), "Liquidación de camionero id: " & lnCodEntrada & " camionero: " & cCamion.Text _
                                                    , paMCCtaDepositoEfectivo, cAZureo, iIDSRPesos, 31)
        If idZureoDepositoEfectivo = 0 Then
            MsgBox "ATENCIÓN!!!" & vbCrLf & "No se obtuvo el id de comprobante que almacena el depósito efectivo de la liquidación, se debe ingresar un depósito " & " por $ " & Abs(cAZureo), vbExclamation, "ATENCIÓN"
        Else
            sIDSComprobantes = sIDSComprobantes & "Depósito efectivo: " & idZureoDepositoEfectivo & "<BR>"
            Set oCQC = New clsConQuePaga
            oCQC.IDDocumentoPaga = idZureoDepositoEfectivo
            oCQC.Importe = cAZureo
            oCQC.TipoConQuePaga = DepositoEfectivo
            
            oAsignados.ConQuePaga.Add oCQC
        End If
    End If
    
    'Veo si tengo aportes de comisión.  paMCCtaComisionRP
    cAZureo = GetImportePago(paMCCtaComisionRP)
    If cAZureo <> 0 Then
            '(IIf(tipoCompZureo = EntradaCaja, 1, 0))
        idZureoDepositoComision = Zureo_InsertoEntradaSalida(SalidaCaja, CStr(lnCodEntrada), "Liquidación de camionero id: " & lnCodEntrada & " camionero: " & cCamion.Text _
                                                    , paMCCtaComisionRP, cAZureo, iIDSRPesos, 0)
        If idZureoDepositoComision = 0 Then
            MsgBox "ATENCIÓN!!!" & vbCrLf & "No se obtuvo el id de comprobante que pasa el gasto de comisión de la liquidación, se debe ingresar una salida de caja " & " por $ " & Abs(cAZureo), vbExclamation, "ATENCIÓN"
        Else
            
            sIDSComprobantes = sIDSComprobantes & "Comisión Redpagos: " & idZureoDepositoComision & "<BR>"
            
            Set oCQC = New clsConQuePaga
            oCQC.IDDocumentoPaga = idZureoDepositoComision
            oCQC.Importe = cAZureo
            oCQC.TipoConQuePaga = GastoComision
            
            oAsignados.ConQuePaga.Add oCQC
        End If
    End If
    
    
    If oAsignados.ConQuePaga.Count > 0 Then
    
        Dim oTrans As New clsCobrarConQuePaga
        Set oTrans.Conexion = cBase
        oAsignados.DocumentoQueSalda = lnCodEntrada
        oAsignados.TipoDocQueSalda = LiquidacionDeCamioneros
        
        oAsignados.Sucursal = paCodigoDeSucursal
        oAsignados.Terminal = paCodigoDeTerminal
        
        oTrans.GrabarAsignaciones oAsignados
        'lnCodEntrada
    End If
    cBase.CommitTrans
    
    
    lLastID = lnCodEntrada
    act_SaveHTML sIDNro, lnCodEntrada, sIDSComprobantes
    AccionImprimir False, lnCodEntrada, sIDNro
    Foco cCamion
    tComentario.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
    
errBT:
    bGrabar.Enabled = True
    tComentario.Enabled = True
    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub

ErrVA:
    bGrabar.Enabled = True
    tComentario.Enabled = True
    cBase.RollbackTrans
    If sErrAtrapado <> "" Then
        MsgBox sErrAtrapado, vbExclamation, "ATENCIÓN"
    Else
        clsGeneral.OcurrioError "Ocurrió un error al intentar almacenar la información." & Chr(13) & Msg, Err.Description
    End If
'    Set objSuceso = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
ErrRB:
    Resume ErrVA
End Sub

Private Function BuscoSignoMoneda(idMoneda As Long)
Dim rsMon As rdoResultset
On Error GoTo errBSM
    BuscoSignoMoneda = ""
    Cons = "Select * From Moneda Where MonCodigo = " & idMoneda
    Set rsMon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not rsMon.EOF Then BuscoSignoMoneda = Trim(rsMon!MonSigno)
    rsMon.Close
errBSM:
End Function

Private Sub LimpioDatos()
    InicializoGrillas
    tComentario.Text = ""
    Set oTransRP = Nothing
'    Set colReclamos = Nothing
'    Set colPagos = Nothing
'
'    Set colReclamos = New Collection
'    Set colPagos = New Collection
    
    Set colLiqCobro = Nothing
    Set colLiqCobro = New Collection
    
    If tUsuario.GetConfigUser <> 2 Then tUsuario.UserID = 0
    Erase arrPendiente
End Sub

Private Sub vsLiquidacion_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)

    If IsNumeric(vsLiquidacion.EditText) Then
        vsLiquidacion.Cell(flexcpText, Row, 2) = Format(vsLiquidacion.EditText * vsLiquidacion.Cell(flexcpData, Row, 2), "#,##0.00")
        vsLiquidacion.EditText = Format(vsLiquidacion.EditText, "#,##0.00")
    End If
End Sub

Private Sub vsPagaCon_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If vbKeyDelete = KeyCode Then
        With vsPagaCon
            If .Row >= .FixedRows + 1 Then
                .RemoveItem .Row
                .RemoveItem .FixedRows            'Total
                .Subtotal flexSTSum, -1, 2, , &HA0A000, vbWhite, True, "Total Paga Con (" & sSignMonedaPesos & ")"
                ArmoDetalleLiquidacion
            End If
        End With
    End If
End Sub

Private Sub ArmoDetalleLiquidacion()
Dim iCont As Integer
Dim cPedir As Currency

    With vsLiquidacion
        .Rows = 0
        .AddItem "A Reclamar"
        .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
        
        .AddItem "  Por cobranza"
        .Cell(flexcpText, .Rows - 1, 2) = Format(GetImporteReclamar, "#,##0.00")
        cPedir = CCur(.Cell(flexcpText, .Rows - 1, 2))
        
        .AddItem "  " & ct_KeyDisponibilidadCamioneros
        .Cell(flexcpText, .Rows - 1, 2) = Format(GetImporteReclamarDiferencias, "#,##0.00")
        cPedir = cPedir + CCur(.Cell(flexcpText, .Rows - 1, 2))
        
        .AddItem "  Pendientes caja"
        .Cell(flexcpText, .Rows - 1, 2) = Format(GetImportePendienteCaja(True) - GetImportePendienteCaja(False), "#,##0.00")
        cPedir = cPedir + CCur(.Cell(flexcpText, .Rows - 1, 2))
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 2) = Space(200)
        .Cell(flexcpFontUnderline, .Rows - 1, 0, , .Cols - 1) = True
        
        .AddItem "Total a reclamar"
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
        .Cell(flexcpText, .Rows - 1, 2) = Format(cPedir, "#,##0.00")
        .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = &HE16949
        If cPedir < 0 Then .Cell(flexcpForeColor, .Rows - 1, 2) = &H3C14DC
        
        Dim cPaga As Currency
        .AddItem "Pagos"
        .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True

        Dim cEfectivo As Currency, cCheque As Currency, cPagoRP As Currency
        cEfectivo = GetImportePagoEfectivo
        .AddItem "  Efectivo a depositar"
        .Cell(flexcpText, .Rows - 1, 2) = Format(cEfectivo, "#,##0.00")
        cPaga = CCur(.Cell(flexcpText, .Rows - 1, 2))
        
        cCheque = GetImportePago(-1)
        .AddItem "  Cheques a depositar"
        .Cell(flexcpText, .Rows - 1, 2) = Format(cCheque, "#,##0.00")
        cPaga = cPaga + CCur(.Cell(flexcpText, .Rows - 1, 2))
        
        cPagoRP = GetImportePago(-3)
        .AddItem "  Giros Redpagos"
        .Cell(flexcpText, .Rows - 1, 2) = Format(cPagoRP, "#,##0.00")
        cPaga = cPaga + CCur(.Cell(flexcpText, .Rows - 1, 2))
        
        .AddItem "  Otros conceptos"
        .Cell(flexcpText, .Rows - 1, 2) = Format(GetImportePagoTotal - (cPagoRP + cCheque + cEfectivo), "#,##0.00")
        cPaga = cPaga + CCur(.Cell(flexcpText, .Rows - 1, 2))
        
        .AddItem "  " & ct_KeyDisponibilidadCamioneros
        .Cell(flexcpText, .Rows - 1, 2) = Format(GetImporteAporteDiferencias, "#,##0.00")
        cPaga = cPaga + CCur(.Cell(flexcpText, .Rows - 1, 2))
        
        'LE SUMO LOS PENDIENTES DE CAJA NEGATIVOS.
'        cPaga = cPaga + GetImportePendienteCaja(False)
                
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 2) = Space(200)
        .Cell(flexcpFontStrikethru, .Rows - 1, 0, , .Cols - 1) = True
        .Cell(flexcpFontUnderline, .Rows - 1, 0, , .Cols - 1) = True
        
        .AddItem IIf(cPedir > cPaga, "Pendiente a reclamar", "Sobrante")
        .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
        .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = IIf(cPaga - cPedir < 0, vbRed, &HE16949)
        .Cell(flexcpText, .Rows - 1, 2) = Format(cPaga - cPedir, "#,##0.00")
                
                
        'ACA Tengo el total con que CGSA tiene a Reclamar o que le sobra.
        .AddItem ""
    End With
    
    s_SetTotalBoleta
    
End Sub

Private Sub EtiquetaLiquidacion()
        
    tSon.Tag = ""
    
    Dim bVisible As Boolean
    bVisible = (cTipo.ItemData(cTipo.ListIndex) > -3 And cTipo.ItemData(cTipo.ListIndex) <> -1)
    
    cMoneda.Visible = bVisible
    tA.Visible = bVisible
    lA.Visible = bVisible
    
    bVisible = (cTipo.ItemData(cTipo.ListIndex) <> -1)
    
    tSon.Visible = bVisible
    lSon.Visible = bVisible
    
    If (cTipo.ItemData(cTipo.ListIndex) <= -3) Then
        tSon.Left = cMoneda.Left
    Else
        tSon.Left = cMoneda.Left + cMoneda.Width + 30
    End If
    
    lblAyudaCheque.Visible = (cTipo.ItemData(cTipo.ListIndex) = -1)
    
    Select Case cTipo.ItemData(cTipo.ListIndex)
         'Son depósitos
         Case Is > 0
            lSon.Caption = "&Son": lA.Caption = "&Q"
            
        'Cheque
        Case -1
            lSon.Caption = "&Son": lA.Caption = "&Q"
            If cMoneda.ListIndex > -1 Then
                If cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos Then lA.Caption = "&T.C."
            End If
            
        'ME
        Case -2
            lSon.Caption = "&Son": lA.Caption = "&T.C."
            
        Case -3
            lSon.Caption = "ID:"
        
        Case -4
            lSon.Caption = "Nº Nota"
        
        Case 0
            lSon.Caption = "&De": lA.Caption = "&Q"
        
    End Select
    

End Sub

Private Sub s_AddRenglonPagaCon()
Dim iRow As Integer
    
'    If iGTReclamo = 0 Then ArmoDetalleLiquidacion
    With vsPagaCon
    
    'Quito el total.
        If .Rows > .FixedRows Then
            If .IsSubtotal(.FixedRows) Then .RemoveItem .FixedRows
        End If
    
        Select Case cTipo.ItemData(cTipo.ListIndex)
            Case 0
                .AddItem "  Con " & Trim(tA.Text) & " billetes de " & Trim(cMoneda.Text) & " " & tSon.Text
                .Cell(flexcpText, .Rows - 1, 2) = Format(CCur(tSon.Text) * CCur(tA.Text), "#,##0.00")
        
'            Case -1
'                If cMoneda.ItemData(cMoneda.ListIndex) = paMonedaPesos Then
'                    .AddItem "  Con " & Trim(tA.Text) & " cheques "
'                    .Cell(flexcpText, .Rows - 1, 2) = Format(tSon.Text, "#,##0.00")
'                Else
'                    .AddItem "  Con cheque(s) por " & Trim(cMoneda.Text) & " " & Format(tSon.Text, "#,##0.00")
'                    .Cell(flexcpText, .Rows - 1, 1) = Format(tA.Text, "#,##0.00")
'                    .Cell(flexcpText, .Rows - 1, 2) = Format(CCur(tSon.Text) * CCur(tA.Text), "#,##0.00")
'                    EjecutarApp App.Path & "\comprame.exe", "MV " & cMoneda.ItemData(cMoneda.ListIndex) & "|IV " & CCur(tSon.Text) & "|TC " & CCur(tA.Text)
'                End If
                    
                    
            Case -2
                .AddItem "  Con M/E por " & Trim(cMoneda.Text) & " " & Format(tSon.Text, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 1) = Format(tA.Text, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 2) = Format(CCur(tSon.Text) * CCur(tA.Text), "#,##0.00")
    
                EjecutarApp App.Path & "\comprame.exe", "MV " & cMoneda.ItemData(cMoneda.ListIndex) & "|IV " & CCur(tSon.Text) & "|TC " & CCur(tA.Text)
                        
            Case -3
                If tSon.Tag <> "" Then
                    .AddItem tSon.Tag
                    .Cell(flexcpText, .Rows - 1, 1) = Format(tSon.Text, "#,##0")
                    .Cell(flexcpText, .Rows - 1, 2) = Format(CCur(lSon.Tag), "#,##0.00")
                    
                    .Cell(flexcpData, .Rows - 1, 0) = cTipo.ItemData(cTipo.ListIndex)   'ID Disponibilidad
                    .Cell(flexcpData, .Rows - 1, 1) = tSon.Text
                Else
                    Exit Sub
                End If
                
            Case -4
                .AddItem "Nota devolución"
                .Cell(flexcpText, .Rows - 1, 1) = tA.Text
                .Cell(flexcpText, .Rows - 1, 2) = Format(tSon.Tag, "#,##0.00")
                
                .Cell(flexcpData, .Rows - 1, 0) = cTipo.ItemData(cTipo.ListIndex)   'ID Disponibilidad
                .Cell(flexcpData, .Rows - 1, 1) = tA.Tag    'id del documento.
            
            Case Is > 0
                If cMoneda.ItemData(cMoneda.ListIndex) = paMonedaPesos Then
                    .AddItem "  Con " & Trim(tA.Text) & " " & cTipo.Text
                    .Cell(flexcpText, .Rows - 1, 2) = Format(tSon.Text, "#,##0.00")
                Else
                    .AddItem "  Con depósito(s) por " & Trim(cMoneda.Text) & " " & Format(tSon.Text, "#,##0.00")
                    .Cell(flexcpText, .Rows - 1, 1) = Format(tA.Text, "#,##0.00")
                    .Cell(flexcpText, .Rows - 1, 2) = Format(CCur(tSon.Text) * CCur(tA.Text), "#,##0.00")
                    EjecutarApp App.Path & "\comprame.exe", "MV " & cMoneda.ItemData(cMoneda.ListIndex) & "|IV " & CCur(tSon.Text) & "|TC " & CCur(tA.Text)
                End If
                'ME DIGO QUE ES DEPÓSITO
'                .Cell(flexcpData, .Rows - 1, 0) = cTipo.ItemData(cTipo.ListIndex)   'ID Disponibilidad
                .Cell(flexcpData, .Rows - 1, 1) = cMoneda.ItemData(cMoneda.ListIndex)
                .Cell(flexcpData, .Rows - 1, 2) = CCur(tSon.Text)
                
        End Select
        'ID DE DISPONIBILIDAD O SELECCIÓN
        .Cell(flexcpData, .Rows - 1, 0) = cTipo.ItemData(cTipo.ListIndex)   'ID Disponibilidad
        
        .Subtotal flexSTSum, -1, 2, , &HA0A000, vbWhite, True, "Total Paga Con (" & sSignMonedaPesos & ")"
    End With

'    SetTotalLiquidacion
    
    ArmoDetalleLiquidacion
    
    tA.Text = ""
    tSon.Text = ""
    cMoneda.Text = ""
    cTipo.SetFocus

End Sub

Private Sub ReturnEnvio(ByVal iRow As Integer, ByVal lIDEnvio As Long)
'Cdo accede al formulario de envío, valido si cambio algún dato.
Dim sPorHora As Boolean, cLiquidar As Currency, cCobrar As Currency, cFacCamion As Currency
Dim aModificacion As String
Dim RsDifEnvio As rdoResultset
Dim iCont As Integer
Dim cImpDocPend As Currency

    
    If cCamion.ListIndex > -1 Then sPorHora = CamionPagoPorHora(cCamion.ItemData(cCamion.ListIndex))
    
    Cons = "Select * " _
        & " From Envio " _
            & " Left Outer Join PendientesCaja On EnvCodigo = PCaEnvio And PCaFLiquidacion Is Null " _
            & " Left Outer Join Disponibilidad On PCaDisponibilidad = DisID " _
        & " , TipoFlete, Moneda" _
        & " Where EnvCodigo = " & lIDEnvio & " And EnvLiquidacion Is Null " _
        & " And EnvTipoFlete = TFlCodigo And EnvMoneda = MonCodigo"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Al agregar el pendiente de caja, un envío puede tener + de uno
    'entonces me puede devolver + de una tupla, por ello agregue control de lidEnv
    lIDEnvio = 0
    If RsAux.EOF Then
        'Anulo esta fila
        RsAux.Close
        If vsConsulta.Cell(flexcpBackColor, iRow, 0) <> cte_CelesteClaro Then
'            InsertoEnTablaPago EnviosAReclamar, vsConsulta.Cell(flexcpData, iRow, 2), vsConsulta.Cell(flexcpData, iRow, 3), vsConsulta.Cell(flexcpData, iRow, 4), -1, vsConsulta.Cell(flexcpData, iRow, 9), 0
            ArmoDetalleLiquidacion
        End If
        vsConsulta.RemoveItem iRow
        Exit Sub
    End If
    
    If IsNull(RsAux!EnvCamion) Then
        'Anulo esta fila
        RsAux.Close
        If vsConsulta.Cell(flexcpBackColor, iRow, 0) <> cte_CelesteClaro Then
'            InsertoEnTablaPago EnviosAReclamar, vsConsulta.Cell(flexcpData, iRow, 2), vsConsulta.Cell(flexcpData, iRow, 3), vsConsulta.Cell(flexcpData, iRow, 4), -1, vsConsulta.Cell(flexcpData, iRow, 9), 0
            ArmoDetalleLiquidacion
        End If
        vsConsulta.RemoveItem iRow
        Exit Sub
    Else
        If cCamion.ItemData(cCamion.ListIndex) <> RsAux!EnvCamion Then
            'Anulo esta fila.
            RsAux.Close
            If vsConsulta.Cell(flexcpBackColor, iRow, 0) <> cte_CelesteClaro Then
'                InsertoEnTablaPago EnviosAReclamar, vsConsulta.Cell(flexcpData, iRow, 2), vsConsulta.Cell(flexcpData, iRow, 3), vsConsulta.Cell(flexcpData, iRow, 4), -1, vsConsulta.Cell(flexcpData, iRow, 9), 0
                ArmoDetalleLiquidacion
            End If
            vsConsulta.RemoveItem iRow
            Exit Sub
        End If
    End If
    
    'Válido el estado del envío.
    If (RsAux!EnvEstado <> 4) And txtCodImpresion.Text = "" Then              'EstadoEnvio.Entregado
        RsAux.Close
        If vsConsulta.Cell(flexcpBackColor, iRow, 0) <> cte_CelesteClaro Then
'            InsertoEnTablaPago EnviosAReclamar, vsConsulta.Cell(flexcpData, iRow, 2), vsConsulta.Cell(flexcpData, iRow, 3), vsConsulta.Cell(flexcpData, iRow, 4), -1, vsConsulta.Cell(flexcpData, iRow, 9), 0
            ArmoDetalleLiquidacion
        End If
        vsConsulta.RemoveItem iRow
        Exit Sub
    End If
    
    If vsConsulta.Cell(flexcpBackColor, iRow, 0) = cte_CelesteClaro Then
        RsAux.Close
        Exit Sub
    End If
    
    
        'ELIMINO LOS PENDIENTES
    If vsConsulta.Cell(flexcpData, iRow, 8) <> 0 Then
        For iCont = 1 To UBound(arrPendiente)
            If arrPendiente(iCont).Envio = CLng(vsConsulta.Cell(flexcpText, iRow, 1)) Then
                If arrPendiente(iCont).Activo Then
'                    InsertoPendienteDePago arrPendiente(iCont).Moneda, arrPendiente(iCont).Importe, -1
                End If
                With arrPendiente(iCont)
                    .Envio = 0
                    .Activo = False
                End With
            End If
        Next
    End If
    
    Do While Not RsAux.EOF
        cLiquidar = 0: cCobrar = 0: cFacCamion = 0
        
        With vsConsulta
            If lIDEnvio <> RsAux!EnvCodigo Then
                lIDEnvio = RsAux!EnvCodigo
                .Cell(flexcpText, iRow, 9) = Trim(RsAux!EnvFModificacion)
                .Cell(flexcpText, iRow, 3) = Trim(RsAux!TFlNombreCorto)
                .Cell(flexcpText, iRow, 8) = ""     'Limpio
                .Cell(flexcpData, iRow, 8) = ""
                .Cell(flexcpData, iRow, 7) = ""
                
                'Si tiene para liquidar y no es pago por hora
                If Not IsNull(RsAux!EnvLiquidar) And Not sPorHora Then cLiquidar = RsAux!EnvLiquidar
                
                'Voy a la tabla documentopendiente para validar los documentos pendientes del envío impresos en depósito.
                Cons = fnc_GetStringDocumentosPendientes(RsAux("EnvCodigo"), 1)
                Set RsDifEnvio = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsDifEnvio.EOF Then
                    
                    'Método nuevo
                    Do While Not RsDifEnvio.EOF
                        .Cell(flexcpText, iRow, 8) = IIf(.Cell(flexcpText, iRow, 8) <> "", .Cell(flexcpText, iRow, 8) & ", ", "") & RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero
                        cCobrar = cCobrar + RsDifEnvio("DPeImporte")
                        RsDifEnvio.MoveNext
                    Loop
                    RsDifEnvio.Close
                    cImpDocPend = cCobrar
                    
                Else
                
                    'Veo que le reclamo y que le pago.
                    'Si el envío fue pago en caja, se le pudo agregar diferencia de envío a cobrar en el domicilio.
                    If RsAux!EnvFormaPago = TipoPagoEnvio.PagaAhora Then
                    
                        Cons = "Select DiferenciaEnvio.*, DocSerie, DocNumero " _
                                & " From DiferenciaEnvio, Documento Where DEvEnvio = " & RsAux!EnvCodigo _
                                & " And DEvFormaPago = " & TipoPagoEnvio.PagaDomicilio _
                                & " And DEvDocumento = DocCodigo"
                                
                        Set RsDifEnvio = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        If Not RsDifEnvio.EOF Then
                            If Not IsNull(RsDifEnvio!DevValorFlete) Then cCobrar = cCobrar + RsDifEnvio!DevValorFlete
                            If Not IsNull(RsDifEnvio!DevValorPiso) Then cCobrar = cCobrar + RsDifEnvio!DevValorPiso
                            .Cell(flexcpText, iRow, 8) = RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero
                        End If
                        RsDifEnvio.Close
                        
                    Else
                        If RsAux!EnvFormaPago = TipoPagoEnvio.PagaDomicilio Then
                            If Not IsNull(RsAux!EnvDocumentoFactura) Then
                                Cons = "Select DocSerie, DocNumero From Documento Where DocCodigo = " & RsAux!EnvDocumentoFactura
                                Set RsDifEnvio = cBase.OpenResultset(Cons, rdOpenForwardOnly)
                                If Not RsDifEnvio.EOF Then .Cell(flexcpText, iRow, 8) = RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero
                                RsDifEnvio.Close
                            End If
                            If Not IsNull(RsAux!EnvValorFlete) Then cCobrar = RsAux!EnvValorFlete
                        Else
                            'Cobro camión.
                            If Not IsNull(RsAux!EnvValorFlete) And RsAux!EnvFormaPago = TipoPagoEnvio.FacturaCamión Then cFacCamion = RsAux!EnvValorFlete
                        End If
                    End If
                    
                    If Not IsNull(RsAux!EnvReclamoCobro) Then
                        cCobrar = cCobrar + RsAux!EnvReclamoCobro
                        
                        'Si es vta. telefonica busco el código.
                        If RsAux!EnvFormaPago <> TipoPagoEnvio.PagaAhora Then
                            'Si el join devuelve algo es porque es una vta. telefónica.
                            Cons = "Select DocSerie, DocNumero From Documento, VentaTelefonica Where DocCodigo = " & RsAux!EnvDocumento _
                                 & " And DocCodigo = VTeDocumento"
                            Set RsDifEnvio = cBase.OpenResultset(Cons, rdOpenForwardOnly)
                            If Not RsDifEnvio.EOF Then
                                If InStr(1, .Cell(flexcpText, iRow, 8), RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero) = 0 Then
                                    If Trim(.Cell(flexcpText, iRow, 8)) = vbNullString Then
                                        .Cell(flexcpText, iRow, 8) = RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero
                                    Else
                                        .Cell(flexcpText, iRow, 8) = Trim(.Cell(flexcpText, iRow, 8)) & ", " & RsDifEnvio!DocSerie & " " & RsDifEnvio!DocNumero
                                    End If
                                End If
                            End If
                            RsDifEnvio.Close
                        End If
                    End If
                End If
                
                'La bajo y la agrego abajo.
'                InsertoEnTablaPago EnviosAReclamar, .Cell(flexcpData, iRow, 2), vsConsulta.Cell(flexcpData, iRow, 3), vsConsulta.Cell(flexcpData, iRow, 4), -1, .Cell(flexcpData, iRow, 9), 0
                
                'Asigno el valor de importe en doc pendientes.
                .Cell(flexcpData, iRow, 9) = cImpDocPend
                
                aModificacion = Trim(BuscoSignoMoneda(RsAux!EnvMoneda))
                If cLiquidar > 0 Then
                    .Cell(flexcpText, iRow, 4) = Trim(aModificacion) & " " & Format(cLiquidar, FormatoMonedaP)
                Else
                    .Cell(flexcpText, iRow, 4) = Trim(aModificacion) & " " & "0.00"
                End If
                .Cell(flexcpData, iRow, 4) = cLiquidar
                
                If cCobrar > 0 Then
                    .Cell(flexcpText, iRow, 5) = Trim(aModificacion) & " " & Format(cCobrar, FormatoMonedaP)
                Else
                    .Cell(flexcpText, iRow, 5) = Trim(aModificacion) & " " & "0.00"
                End If
                .Cell(flexcpText, iRow, 6) = Trim(aModificacion) & " " & Format(cFacCamion, FormatoMonedaP)
                
                
                cLiquidar = RsAux!EnvMoneda: .Cell(flexcpData, iRow, 2) = cLiquidar
                .Cell(flexcpData, iRow, 3) = cCobrar
                
'                InsertoEnTablaPago EnviosAReclamar, RsAux!EnvMoneda, cCobrar, .Cell(flexcpData, iRow, 4), 1, cImpDocPend, 0
                
            End If
            
            If Not IsNull(RsAux!PCaEnvio) Then
                ReDim Preserve arrPendiente(UBound(arrPendiente) + 1)
                With arrPendiente(UBound(arrPendiente))
                    .Envio = RsAux!EnvCodigo
                    .Pendiente = RsAux!PCaID
                    .Moneda = RsAux!DisMoneda
                    .Importe = RsAux!PCaImporte
                    .Activo = True
                End With
                cCobrar = RsAux!DisMoneda
                If .Cell(flexcpData, iRow, 8) = "" Then
                    .Cell(flexcpData, iRow, 8) = cCobrar
                    .Cell(flexcpData, iRow, 7) = arrPendiente(UBound(arrPendiente)).Importe
                    If RsAux!DisMoneda <> RsAux!EnvMoneda Then
                        aModificacion = Trim(BuscoSignoMoneda(RsAux!DisMoneda))
                    End If
                    .Cell(flexcpText, iRow, 7) = aModificacion & " " & Format(RsAux!PCaImporte, "#,##0.00")
                Else
                    If .Cell(flexcpData, iRow, 8) = cCobrar Then
                        .Cell(flexcpData, iRow, 7) = CCur(.Cell(flexcpData, iRow, 7)) + RsAux!PCaImporte
                        .Cell(flexcpText, iRow, 7) = aModificacion & " " & Format(CCur(.Cell(flexcpData, iRow, 7)), "#,##0.00")
                    Else
                        'Tengo que hacer taza de cambio para esta moneda
                        .Cell(flexcpText, iRow, 7) = "Varias Monedas"
                    End If
                End If
'                InsertoPendienteDePago RsAux!DisMoneda, RsAux!PCaImporte, 1
            Else
                .Cell(flexcpData, iRow, 7) = "0.00"
            End If
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    ArmoDetalleLiquidacion
    
End Sub

Private Sub act_SaveHTML(ByVal sCodComun As String, lIDLiq As Long, ByVal sIDSComprobantes As String)
On Error Resume Next
    
    act_SaveFileHTML lIDLiq, sCodComun, sIDSComprobantes
    
End Sub

Private Sub act_SaveFileHTML(ByVal lIDLiq As Long, ByVal sCodMD As String, ByVal sIDSComprobantes As String)
Dim sFile As String
Dim iFile As Integer

    On Error GoTo errArmo
    
    sFile = "<HTML>" & vbCrLf & _
                "<HEAD>" & vbCrLf & _
                    "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html;charset=windows-1252"">" & vbCrLf & _
                    "<META NAME=""Generator"" >" & vbCrLf & _
                    "<TITLE>Liquidación de Camionero</TITLE>" & vbCrLf & _
                "</HEAD>" & vbCrLf & _
                    "<BODY>" & vbCrLf & vbCrLf & _
                    "<BR><b> Liquidación de Camión: " & Trim(cCamion.Text) & "<BR><br>" & _
                    "Código de liquidación: " & lIDLiq & "<br><br>" & vbCrLf & _
                    "Movimientos de disponibilidad: " & sCodMD & "</b><br><br>" & vbCrLf
                    
    vsConsulta.ExtendLastCol = False
    vsConsulta.ColHidden(vsConsulta.Cols - 1) = True
    sFile = sFile & GetFlexGridToHTML(vsConsulta) & "<BR>" & vbCrLf & "<BR><b>Conceptos Liquidación:<b><BR>" & vbCrLf
    vsConsulta.ColHidden(vsConsulta.Cols - 1) = False
    vsConsulta.ExtendLastCol = True
    'sFile = sFile & GetFlexGridToHTML(vsPago) & "<BR>" & vbCrLf & "<BR><b>Detalle de la boleta<b><BR>" & vbCrLf
    sFile = sFile & "<BR><b>Detalle de la boleta<b><BR>" & vbCrLf
    sFile = sFile & GetFlexGridToHTML(vsDetMonto) & "<BR>" & vbCrLf & "<BR><b>Paga Con:</b><BR>" & vbCrLf
    sFile = sFile & GetFlexGridToHTML(vsPagaCon) & "<BR>" & vbCrLf & "<BR><b>Resumen Final de Liquidación:</b><BR>" & vbCrLf
    sFile = sFile & GetFlexGridToHTML(vsLiquidacion) & "<BR>" & vbCrLf & "<BR><BR>" & vbCrLf
    sFile = sFile & "<BR><BR> Comentario: " & Trim(tComentario.Text) & _
                IIf(sIDSComprobantes <> "", "<BR><BR>" & sIDSComprobantes, "") & _
                "<BR><BR>Usuario: " & tUsuario.UserName & "<BR> Terminal: " & miConexion.NombreTerminal
    
    sFile = sFile & "</BODY>" & vbCrLf & "</HTML>" & vbCrLf

    'ALMACENO EL ARCHIVO
    On Error GoTo errSaveLocal
    iFile = FreeFile
    Open prmPahtHTML & "LiqCamion" & Format(lIDLiq, "000000") & ".htm" For Output As iFile
    Print #iFile, sFile
    Close iFile
    
    Exit Sub
    
errArmo:
    MsgBox "Ocurrió el siguiente error al intentar crear el html para almacenar en el archivo.", vbCritical, "ATENCIÓN"
    
errSaveLocal:
    MsgBox "Atención ocurrió el siguiente error " & Err.Description & _
        "  al grabar el archivo html " & vbCrLf & "El mismo será almacenado en su terminal." & vbCrLf & _
        " COMUNIQUE ESTE PROBLEMA AL ADMINISTRADOR ", vbCritical, "ERROR"
    Open App.Path & "LiqCamion" & Format(lIDLiq, "000000") & ".htm" For Output As iFile
    Print #iFile, sFile
    Close iFile
    Exit Sub

End Sub

Private Sub s_SetTotalBoleta()
Dim iCont As Integer
    With vsDetMonto
        
        iCont = .FixedRows
        'Primero borro el subtotal.
        Do While iCont <= .Rows - 1
            If .IsSubtotal(iCont) Or Val(vsDetMonto.Cell(flexcpData, iCont, 0)) = 1 Then
                .RemoveItem iCont
            Else
                iCont = iCont + 1
            End If
        Loop
        
        'Armo en base a lo que tengo en la liquidación.
        Dim colBoleta As Collection
        Set colBoleta = GetRenglonesBoleta()
        If colBoleta.Count > 0 Then
            Dim oReg As clsImporteCantidad
            For Each oReg In colBoleta
                .AddItem "Reparto", .FixedRows
                .Cell(flexcpData, .FixedRows, 0) = 1 'Marco que es detalle a liquidar.

                .Cell(flexcpText, .FixedRows, 2) = oReg.Cantidad
                .Cell(flexcpText, .FixedRows, 3) = sSignMonedaPesos & " " & Format(oReg.Importe, "#,##0.00")
                .Cell(flexcpText, .FixedRows, 4) = Format(oReg.Cantidad * oReg.Importe, "#,##0.00")
            Next
        End If
        
        If .Rows > .FixedRows Then
            .Subtotal flexSTSum, -1, 4, , , &H40C0&, True, "Total Boleta"
            'Le pongo el IVA
            .Cell(flexcpText, 1, 1) = "I.V.A.: " & Format(CCur(.Cell(flexcpValue, 1, 4)) - (CCur(.Cell(flexcpValue, 1, 4)) / paxIVA), "#,##0.00")
        End If
        
    End With
End Sub

Private Sub MarcaroDesmarcar(ByVal bSelect As Boolean)
Dim iQ As Integer, iAlta As Integer, iCont As Integer

    With vsConsulta
        For iQ = .FixedRows To .Rows - .FixedRows
        
            If bSelect Then
                .Cell(flexcpBackColor, iQ, 0, , .Cols - 1) = cte_CelesteClaro
            Else
                .Cell(flexcpBackColor, iQ, 0, , .Cols - 1) = vbWhite
            End If
            
            'SeCobraSINO vsConsulta.Cell(flexcpData, iQ, 0), CLng(vsConsulta.Cell(flexcpText, iQ, 1)), Not bSelect
            SeCobraSINO vsConsulta.Cell(flexcpData, iQ, 0), CLng(vsConsulta.Cell(flexcpText, iQ, 1)), Not bSelect, CCur(vsConsulta.Cell(flexcpData, iQ, 4)), CCur(vsConsulta.Cell(flexcpData, iQ, 3))
            
            If .Cell(flexcpData, iQ, 8) <> 0 Then
                For iCont = 1 To UBound(arrPendiente)
                    If arrPendiente(iCont).Envio = CLng(vsConsulta.Cell(flexcpText, iQ, 1)) Then
                        arrPendiente(iCont).Activo = Not bSelect
                    End If
                Next
            End If
        Next
    End With
    
    ArmoDetalleLiquidacion
    
End Sub
Private Sub s_HideUnSelect(ByVal bHide As Boolean)
On Error Resume Next
Dim iQ As Integer
    With vsConsulta
        For iQ = .FixedRows To .Rows - .FixedRows
            If Not bHide Then
                .RowHidden(iQ) = False
            Else
                If .Cell(flexcpBackColor, iQ, 0) = cte_CelesteClaro Then .RowHidden(iQ) = True
            End If
        Next
    End With
End Sub

Public Function fnc_ValidoAcceso(Optional ByRef sRet As String = "") As Boolean
On Error GoTo errVAcc
Dim prmAccesosUserLog As String

    fnc_ValidoAcceso = False
    sRet = "Instancio ValidateAccess"
    
    prmAccesosUserLog = objUsers.ValidateAccess(rdoCZureo, "org01", "Comprobantes", 1)
    
    If prmAccesosUserLog <> "" Then
        
        fnc_ValidoAcceso = True
    
    Else

    'Veo si es multiusuario
        Set RsAux = rdoCZureo.OpenResultset("Select ParValor From genParametros Where ParNombre = 'sis_MultiUsuario'", rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If RsAux(0) = 0 Then
                iUserZureo = 0
                RsAux.Close
                fnc_ValidoAcceso = True
                Exit Function
            End If
        End If
        RsAux.Close
        
        Dim mRet As Integer
        'Como valor Q usuarios (-1 error, 0 no hay, 1 hay uno ,2 hay mas de 1)
        iUserZureo = 0
        mRet = objUsers.GetUserData(rdoCZureo, "org01", UserID:=iUserZureo)

        If (mRet <> -1) And (iUserZureo = -1) Then
            '0- No hay acceso;  1- Hay acceso

            mRet = objUsers.doLogIn(rdoCZureo, "org01", CStr(prmUsuario), prmPWDZureo)

            If mRet = 1 Then
                prmAccesosUserLog = objUsers.ValidateAccess(rdoCZureo, "org01", "Comprobantes", 1)
                If prmAccesosUserLog <> "" Then
                
                    objUsers.GetUserData rdoCZureo, "org01", UserID:=iUserZureo
                    'iUserZureo = prmUsuario:
                    fnc_ValidoAcceso = True
                    
                End If
            End If
        End If
    End If
    
    Exit Function
errVAcc:
    sRet = "Error: " & Err.Description & " Paso: " & sRet
End Function

Private Sub loc_GetDatosDisponibilidad(ByVal iDisp As Long, ByRef iSubRubro As Long, ByRef iMoneda As Integer)
Dim sQy As String
Dim rsQ As rdoResultset
    iSubRubro = 0
    iMoneda = 0
    sQy = "Select DisIDSubRubro, DisMoneda From Disponibilidad Where DisID = " & iDisp
    Set rsQ = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not rsQ.EOF Then
        If Not IsNull(rsQ("DisIDSubRubro")) Then iSubRubro = rsQ("DisIDSubRubro")
        If Not IsNull(rsQ("DisMoneda")) Then iMoneda = rsQ("DisMoneda")
    End If
    rsQ.Close
End Sub

Private Sub loc_TransferenciasZureo(ByVal lIDLiquidacion As Long)
On Error GoTo errTZ
'Realizo transferencia para los depósitos....................................................................
    'Recorro la lista de liquidación y me fijo si hay depósitos.
    'Valido usuario zureo
    
    If fnc_ValidoAcceso Then
    
'        Dim arrDispME() As tDatosBco
'        ReDim arrDispME(0)
        
        Dim iQ As Integer, bAdd As Boolean
        Dim iIDCtaBco As Long, iMonBco As Integer
        
        'Para todos los datos ingresados sumo la misma disponibilidad
        
        Dim i As Integer
        With vsPagaCon
            For i = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, i, 0)) > 0 Then
                    
                    loc_GetDatosDisponibilidad Val(.Cell(flexcpData, i, 0)), iIDCtaBco, iMonBco
                    s_InsertoTransfZureo iIDSRPesos, CCur(.Cell(flexcpValue, i, 2)), iIDCtaBco, CCur(.Cell(flexcpData, i, 2)), lIDLiquidacion, iMonBco
                    
'                    bAdd = True
                    'Recorro el array y verifico si ya inserte esta disp.
 '                   For iQ = 0 To UBound(arrDispME)
  '                      If arrDispME(iQ).iDisp = Val(.Cell(flexcpData, i, 0)) Then
   '                         bAdd = False
    '
     '                       arrDispME(iQ).iImpMC = CCur(.Cell(flexcpValue, i, 2)) + arrDispME(iQ).iImpMC
      '                      arrDispME(iQ).iImpME = CCur(.Cell(flexcpData, i, 2)) + arrDispME(iQ).iImpME
       '
        '                    Exit For
         '               End If
          '          Next
           '         If bAdd Then
            '            ReDim Preserve arrDispME(UBound(arrDispME) + 1)
             '           arrDispME(UBound(arrDispME)).iDisp = .Cell(flexcpData, i, 0)
              '          arrDispME(UBound(arrDispME)).iImpMC = CCur(.Cell(flexcpValue, i, 2))
               '         arrDispME(UBound(arrDispME)).iImpME = CCur(.Cell(flexcpData, i, 2))
                '    End If
                
                End If
            Next
        End With
        
'        If UBound(arrDispME) > 0 Then
            
'            If lIDDispPesos = 0 Then lIDDispPesos = modMaeDisponibilidad.dis_DisponibilidadPara(paCodigoDeSucursal, CLng(paMonedaPesos))
 '           loc_GetDatosDisponibilidad lIDDispPesos, iIDSRPesos, 0
            
'            For iQ = 1 To UBound(arrDispME)
 '               loc_GetDatosDisponibilidad arrDispME(iQ).iDisp, iIDCtaBco, iMonBco
  '              s_InsertoTransfZureo iIDSRPesos, arrDispME(iQ).iImpMC, iIDCtaBco, arrDispME(iQ).iImpME, lIDLiquidacion, iMonBco
   '         Next
            
 '       End If
        
    Else
        With vsPagaCon
            For i = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, i, 0)) > 0 Then
                    MsgBox "Sin permisos para hacer la transferencia.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
            Next
        End With
    End If

Exit Sub
errTZ:
    clsGeneral.OcurrioError "Error al intentar realizar las transferencias a Zureo.", Err.Description, "Transferencias"
End Sub

Private Function Zureo_InsertoEntradaSalida(ByVal tipoCompZureo As TipoComprobanteZureo _
                                            , ByVal NumeroDoc As String _
                                            , ByVal Memo As String _
                                            , ByVal cuentaDebe As Long _
                                            , ByVal cImporte As Currency _
                                            , ByVal cuentaHaber As Long _
                                            , ByVal idProveedor As Long) As Long
Dim m_ReturnID As Long

    Dim objComp As New clsComprobantes
    Dim OBJ_COM As clsDComprobante, OBJ_CTA As clsDCuenta
    Dim colCuentas As New Collection
    
    If iUserZureo <= 0 Then iUserZureo = 23
        
    Set OBJ_COM = New clsDComprobante
    With OBJ_COM
        .doAccion = 1
        .Ente = idProveedor
        .Empresa = 1
        .Numero = NumeroDoc        'guardamos el id de liquidación (sirve para buscar)
        .Fecha = Date
        .Tipo = tipoCompZureo
        .Moneda = paMonedaPesos
        .ImporteTotal = cImporte
        .Memo = Memo
        .UsuarioAlta = iUserZureo
        .UsuarioAutoriza = iUserZureo  'IIf(iUserZureo = 0, 23, iUserZureo)
    End With
    Set OBJ_CTA = New clsDCuenta
    With OBJ_CTA
        .VaAlDebe = 1
        .Cuenta = cuentaDebe
        .ImporteComp = cImporte
        .ImporteCta = cImporte
        .MonedaCta = paMonedaPesos
        If tipoCompZureo = DepósitoZureo Then
            .Referencia.ID = 0
            .Referencia.Importe = cImporte
        End If
    End With
    colCuentas.Add OBJ_CTA
    Set OBJ_CTA = Nothing
    
    Set OBJ_CTA = New clsDCuenta
    With OBJ_CTA
        .VaAlDebe = 0 '(IIf(tipoCompZureo = EntradaCaja, 1, 0))
        .Cuenta = cuentaHaber 'iIDSRPesos
        .ImporteComp = cImporte
        .ImporteCta = cImporte
        .MonedaCta = paMonedaPesos
    End With
    colCuentas.Add OBJ_CTA
    Set OBJ_CTA = Nothing
    
    Set OBJ_COM.Cuentas = colCuentas
    If objComp.fnc_PasarComprobante(rdoCZureo, OBJ_COM) Then m_ReturnID = objComp.prm_Comprobante
    Set OBJ_COM = Nothing
    
    Zureo_InsertoEntradaSalida = m_ReturnID
    
End Function

Private Sub s_InsertoTransfZureo(ByVal iIDSRPesos As Long, ByVal cImpPesos As Currency, ByVal lIDSRBco As Long, ByVal cImpDest As Currency, ByVal lIDLiquidacion As Long, ByVal iMonBco As Integer)
Dim m_ReturnID As Long

    Dim objComp As New clsComprobantes
    Dim OBJ_COM As clsDComprobante, OBJ_CTA As clsDCuenta
    Dim colCuentas As New Collection
        
    Set OBJ_COM = New clsDComprobante
    With OBJ_COM
        .doAccion = 1
        .Ente = 0 'idProveedor
        .Empresa = 1
        .Numero = "" 'Trim(sserie & " " & Numero)
        .Fecha = Date                            'CDate(Format(FPago, "dd/mm/yyyy"))
        .Tipo = paMCDeposito
        .Moneda = paMonedaPesos
        .ImporteTotal = cImpPesos
        .Memo = "Liquidación de camionero id: " & lIDLiquidacion
        .UsuarioAlta = iUserZureo
        .UsuarioAutoriza = iUserZureo
    End With
    
    Set OBJ_CTA = New clsDCuenta
    With OBJ_CTA
        .VaAlDebe = 0
        .Cuenta = iIDSRPesos
        .ImporteComp = cImpPesos
        .ImporteCta = cImpPesos
        .MonedaCta = paMonedaPesos          '  xCuentaS_M 'prmMonedaContabilidad
    End With
    colCuentas.Add OBJ_CTA
    Set OBJ_CTA = Nothing
    
    Set OBJ_CTA = New clsDCuenta
    With OBJ_CTA
        .VaAlDebe = 1
        .Cuenta = lIDSRBco
        .ImporteComp = cImpPesos
        .ImporteCta = cImpDest
        .MonedaCta = iMonBco               ' xCuentaE_M 'prmMonedaContabilidad
    End With
    
    Dim oRef As New clsDReferencia
    With oRef
        .ID = 0
        .Cuenta = OBJ_CTA.Cuenta
        .Vencimiento = Date
        .Importe = cImpDest
        .Fecha = Date
        .QPagos = 1
        .Numero = ""
    End With
    Set OBJ_CTA.Referencia = oRef
    Set oRef = Nothing
    
    colCuentas.Add OBJ_CTA
    Set OBJ_CTA = Nothing
    
    Set OBJ_COM.Cuentas = colCuentas
    If objComp.fnc_PasarComprobante(rdoCZureo, OBJ_COM) Then m_ReturnID = objComp.prm_Comprobante
    
    Set OBJ_COM = Nothing
End Sub

Private Sub s_CargoComboFormaPago()

'        Cons = "Select DisID, DisNombre From Disponibilidad Where DisID IN (" & paDispBuzon & ")"
'        CargoCombo Cons, cTipo
        With cTipo
            .AddItem "Billete": .ItemData(.NewIndex) = 0
            .AddItem "Cheque": .ItemData(.NewIndex) = -1
            .AddItem "Comisión Redpagos": .ItemData(.NewIndex) = paMCCtaComisionRP
            .AddItem "Notas de devolución": .ItemData(.NewIndex) = -4
            .AddItem "M/E": .ItemData(.NewIndex) = -2
            .AddItem "Transacciones Redpagos": .ItemData(.NewIndex) = -3
        End With
End Sub
Private Function f_SelectMonedaDisponibilidad() As Long
On Error GoTo errSMD
Dim rsM As rdoResultset
    f_SelectMonedaDisponibilidad = 0
    Cons = "Select DisMoneda From Disponibilidad Where DisID = " & cTipo.ItemData(cTipo.ListIndex)
    Set rsM = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsM.EOF Then f_SelectMonedaDisponibilidad = rsM!DisMoneda
    rsM.Close
    Exit Function
errSMD:
    clsGeneral.OcurrioError "Error al cargar la moneda para la disponibilidad.", Err.Description
End Function

Private Sub s_ChangeMoneda()
    
    cMoneda.Enabled = True
    If cTipo.ListIndex = -1 Then Exit Sub
    Select Case cTipo.ItemData(cTipo.ListIndex)
        Case paMCCtaComisionRP
            cMoneda.ListIndex = 0
            BuscoCodigoEnCombo cMoneda, CLng(paMonedaPesos)
            If cMoneda.ListIndex = -1 Then Exit Sub
            If cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos Then lA.Caption = "&T.C."
            cMoneda.Enabled = False
            
        Case Is > 0
            cMoneda.ListIndex = 0
            BuscoCodigoEnCombo cMoneda, f_SelectMonedaDisponibilidad
            If cMoneda.ListIndex = -1 Then Exit Sub
            If cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos Then lA.Caption = "&T.C."
            cMoneda.Enabled = False
            
        Case 0
            BuscoCodigoEnCombo cMoneda, CLng(paMonedaPesos)
            lA.Caption = "&Q"
            
        Case -1
            lSon.Caption = "&Son": lA.Caption = "&Q"
            If cMoneda.ListIndex > -1 Then
                If cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos Then lA.Caption = "&T.C."
            End If
        Case -2
            lSon.Caption = "&Son": lA.Caption = "&T.C."
            If cMoneda.ListIndex > -1 Then
                If cMoneda.ItemData(cMoneda.ListIndex) = paMonedaPesos Then cMoneda.ListIndex = -1
            End If
    End Select
    
End Sub

Private Sub InicializoGrillas()
    On Error Resume Next
    With vsConsulta
        .Redraw = False
        .ExtendLastCol = True
        .WordWrap = False
        .Cols = 1: .Rows = 1
        .FormatString = "Tipo|>Código|<Dirección|<Flete|>Liquidar|>Reclamar|>Fac.Camión|>Pendiente|<Facturas|SerModificacion|VisModificacion|"
        .ColWidth(1) = 700: .ColWidth(3) = 700
        .ColWidth(2) = 2500: .ColWidth(4) = 1050: .ColWidth(5) = 1350: .ColWidth(6) = 1060: .ColWidth(7) = 1300: .ColWidth(8) = 2200
        .ColWidth(.Cols - 1) = 20
        .ColHidden(9) = True
        .ColHidden(10) = True
        .Redraw = True
    End With
    
    With vsDetMonto
        .Redraw = False
        .Cols = 4: .Rows = 1
        .FixedRows = 1
        .FormatString = "<Detalle|<Ampliación|>Q|>Monto|>Total"
        .ColWidth(2) = 300
        .ColWidth(0) = 1350
        .ColWidth(1) = 2900
        .ColWidth(3) = 1350: .ColWidth(4) = 1350
        .Redraw = True
        .Editable = False
    End With
    With vsLiquidacion
        .ExtendLastCol = False
        .Cols = 3: .Rows = 0
        .ColWidth(0) = 4000
        .ColWidth(1) = 1050
        .ColWidth(2) = 2000
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
    End With
    With vsPagaCon
        .ExtendLastCol = True
        .Cols = 3: .Rows = 0
        .ColWidth(0) = 4000
        .ColWidth(1) = 1050
        .ColWidth(2) = 2000
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
    End With
    
End Sub

Private Sub BuscoTransaccionRP(ByVal idRP As Long)
On Error GoTo errBT
    
    tSon.Tag = ""
    lSon.Tag = ""
    
    Cons = "SELECT TraID ID, IsNull(TraCIRUC, '') CIRUC" & _
        ", Sum(TItImporte)  Importe" & _
        " FROM comTransacciones INNER JOIN comTransaccionItems ON TraID = TItTransaccion" & _
        " WHERE TraEstado = 1 AND (TraID = " & idRP & " OR TraIdZureo = " & idRP & ")" & _
        " GROUP BY TraID, TraCIRUC"

    Dim rsT As rdoResultset
    Set rsT = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsT.EOF Then
    
        'Valido duplicados.
        Dim iR As Integer
        For iR = 1 To vsPagaCon.Rows - 1
            If Val(vsPagaCon.Cell(flexcpData, iR, 0)) = -3 And Val(vsPagaCon.Cell(flexcpData, iR, 1)) = rsT("ID") Then
                MsgBox "Transacción ya insertada, verifique.", vbExclamation, "ATENCIÓN"
                rsT.Close
                Exit Sub
            End If
        Next
    
    
        tSon.Tag = "Con Redpagos " & rsT("ID") & ", CI: " & Trim(rsT("CIRUC"))
        lSon.Tag = rsT("Importe")
        tSon.Text = rsT("ID")
    Else
        MsgBox "No hay datos para el id ingresado.", vbExclamation, "ATENCIÓN"
    End If
    rsT.Close
    
    If lSon.Tag <> "" Then s_AddRenglonPagaCon
    
    Exit Sub
errBT:
    clsGeneral.OcurrioError "Error al buscar la transacción.", Err.Description, "Buscar transacción"
End Sub
