VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCierre 
   BackColor       =   &H8000000B&
   Caption         =   "Cierre de Caja"
   ClientHeight    =   7545
   ClientLeft      =   3075
   ClientTop       =   2100
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCierre.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   9840
   Begin VB.PictureBox picArqueo 
      Height          =   4755
      Left            =   3120
      ScaleHeight     =   4695
      ScaleWidth      =   1935
      TabIndex        =   32
      Top             =   1980
      Width           =   1995
      Begin VB.Label Label2 
         Caption         =   "(Vtas telef, envíos, camioneros)"
         Height          =   315
         Left            =   3540
         TabIndex        =   51
         Top             =   1440
         Width           =   2355
      End
      Begin VB.Label lATPendienteCam 
         Caption         =   "Pendiente Camioneros"
         Height          =   315
         Left            =   180
         TabIndex        =   50
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lAPendienteCam 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "100,000.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1860
         TabIndex        =   49
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lATTotalSaldo 
         Caption         =   "Saldo del Día:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   46
         Top             =   2580
         Width           =   1515
      End
      Begin VB.Label lATotalSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "100,000.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   45
         Top             =   2580
         Width           =   1575
      End
      Begin VB.Label lATSaldoCaja 
         Caption         =   "Saldo del Día"
         Height          =   315
         Left            =   180
         TabIndex        =   44
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Label lASaldoCaja 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "100,000.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   43
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lATotalEfectivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "100,000.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1860
         TabIndex        =   42
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lATTotalEfectivo 
         Caption         =   "Total Efectivo"
         Height          =   315
         Left            =   180
         TabIndex        =   41
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label lAPendienteCaja 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "100,000.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1860
         TabIndex        =   40
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label lATPendienteCaja 
         Caption         =   "Pendiente de Caja"
         Height          =   315
         Left            =   180
         TabIndex        =   39
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label lAChequesReb 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "100,000.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1860
         TabIndex        =   38
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lATChequesReb 
         Caption         =   "Cheques Rebotados"
         Height          =   315
         Left            =   180
         TabIndex        =   37
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label lAChequesDia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "100,000.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1860
         TabIndex        =   36
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label lATChequesDia 
         Caption         =   "Cheques del Día"
         Height          =   315
         Left            =   180
         TabIndex        =   35
         Top             =   540
         Width           =   1515
      End
      Begin VB.Label lABilletes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "100,000.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1860
         TabIndex        =   34
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lATBilletes 
         Caption         =   "Efectivo (Billetes)"
         Height          =   315
         Left            =   180
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox picErrores 
      Height          =   3735
      Left            =   3600
      ScaleHeight     =   3675
      ScaleWidth      =   4635
      TabIndex        =   47
      Top             =   1200
      Width           =   4695
      Begin VSFlex6DAOCtl.vsFlexGrid vsErrores 
         Height          =   1575
         Left            =   60
         TabIndex        =   48
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
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
         SelectionMode   =   0
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
   End
   Begin MSComctlLib.TabStrip tabCaja 
      Height          =   1395
      Left            =   180
      TabIndex        =   31
      Top             =   1020
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2461
      TabWidthStyle   =   1
      TabMinWidth     =   2646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&1) Cierre"
            Key             =   "cierre"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&2) Arqueo"
            Key             =   "arqueo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&3) Errores"
            Key             =   "errores"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCierre 
      Height          =   3375
      Left            =   300
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   26
      Top             =   3000
      Width           =   2655
      Begin VSFlex6DAOCtl.vsFlexGrid vsIngresos 
         Height          =   1755
         Left            =   60
         TabIndex        =   27
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3096
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
         Height          =   1575
         Left            =   60
         TabIndex        =   28
         Top             =   2475
         Width           =   5895
         _ExtentX        =   10398
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
      Begin VB.Label lbIngresos 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DETALLE DE MOVIMIENTOS (POR FACTURACIÓN)"
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
         Left            =   60
         TabIndex        =   30
         Top             =   60
         Width           =   5175
      End
      Begin VB.Label lbMovimientos 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DETALLE DE MOVIMIENTOS"
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
         Left            =   60
         TabIndex        =   29
         Top             =   2160
         Width           =   5175
      End
   End
   Begin VB.Frame fResumen 
      Caption         =   "Resumen Final"
      ForeColor       =   &H00000080&
      Height          =   1020
      Left            =   120
      TabIndex        =   7
      Top             =   5190
      Width           =   9495
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Anterior:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lSubtotal 
         Alignment       =   1  'Right Justify
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
         Left            =   4320
         TabIndex        =   24
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lSaldoN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7320
         TabIndex        =   22
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo Saldo:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lSaldoA 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         TabIndex        =   20
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lTIngreso 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
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
         Left            =   7320
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Movimientos:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lTMovimiento 
         Alignment       =   1  'Right Justify
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
         Left            =   4320
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Facturación:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   675
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   9495
      Begin VB.PictureBox picBotones 
         Height          =   450
         Left            =   6720
         ScaleHeight     =   390
         ScaleWidth      =   2595
         TabIndex        =   14
         Top             =   180
         Width           =   2655
         Begin VB.CommandButton bCierre 
            Height          =   310
            Left            =   480
            Picture         =   "frmCierre.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Dar cierre."
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   310
         End
         Begin VB.CommandButton bImprimir 
            Height          =   310
            Left            =   1080
            Picture         =   "frmCierre.frx":068C
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Imprimir."
            Top             =   60
            Width           =   310
         End
         Begin VB.CommandButton bNoFiltros 
            Height          =   310
            Left            =   1440
            Picture         =   "frmCierre.frx":078E
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Quitar filtros."
            Top             =   60
            Width           =   310
         End
         Begin VB.CommandButton bCancelar 
            Height          =   310
            Left            =   2160
            Picture         =   "frmCierre.frx":0B54
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Salir."
            Top             =   60
            Width           =   310
         End
         Begin VB.CommandButton bConsultar 
            Height          =   310
            Left            =   120
            Picture         =   "frmCierre.frx":0C56
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Ejecutar."
            Top             =   60
            Width           =   310
         End
      End
      Begin AACombo99.AACombo cDisponibilidad 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   240
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
      Begin MSComCtl2.DTPicker dFecha 
         Height          =   315
         Left            =   5340
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   46268417
         CurrentDate     =   37543
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Cierre al:"
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   2
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Disponibilidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   315
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   7290
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "sucursal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "bd"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6562
         EndProperty
      EndProperty
   End
   Begin vsViewLib.vsPrinter vsPrinter 
      Height          =   5775
      Left            =   7560
      TabIndex        =   5
      Top             =   1140
      Visible         =   0   'False
      Width           =   1995
      _Version        =   196608
      _ExtentX        =   3519
      _ExtentY        =   10186
      _StockProps     =   229
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Zoom            =   70
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   120
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierre.frx":0F58
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierre.frx":1272
            Key             =   "help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierre.frx":1384
            Key             =   "consultar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierre.frx":14DE
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierre.frx":1638
            Key             =   "print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierre.frx":1792
            Key             =   "limpiar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierre.frx":18A4
            Key             =   "error50"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierre.frx":1CF6
            Key             =   "error100"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierre.frx":2148
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierre.frx":259A
            Key             =   "error0"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuIrA 
      Caption         =   "&Ir a ..."
      Begin VB.Menu MnuBilletes 
         Caption         =   "Conteo de &Billetes"
      End
      Begin VB.Menu MnuControlCheques 
         Caption         =   "Listado de Control de &Cheques"
      End
      Begin VB.Menu MnuPendientes 
         Caption         =   "&Pendientes de Caja"
      End
      Begin VB.Menu MnuIRL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuABMCheques 
         Caption         =   "Seguimiento de Cheques"
      End
      Begin VB.Menu MnuAltaCheque 
         Caption         =   "Ingreso de Cheques"
      End
      Begin VB.Menu MnuIRL2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPPenCaja 
         Caption         =   "Plantilla de Pendientes de Caja"
      End
      Begin VB.Menu mnuPPenCajaCam 
         Caption         =   "Plantilla de Pendientes de Camioneros"
      End
      Begin VB.Menu MnuPSeñasRecibidas 
         Caption         =   "Plantilla de señas recibidas"
      End
      Begin VB.Menu MnuPMovPendienteCamion 
         Caption         =   "Plantilla movimientos en pendientes camioneros"
      End
   End
   Begin VB.Menu MnuExit 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalir 
         Caption         =   "Cerrar Aplicación"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuMousse 
      Caption         =   "MnuMousse"
      Visible         =   0   'False
      Begin VB.Menu MnuMovimientos 
         Caption         =   "Ver Movimientos"
      End
      Begin VB.Menu MnuMoL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "Cancelar"
      End
   End
   Begin VB.Menu MnuErrores 
      Caption         =   "MnuErrores"
      Visible         =   0   'False
      Begin VB.Menu MnuEVOpe 
         Caption         =   "&Visualización de Operaciones"
      End
      Begin VB.Menu MnuEL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTDocs 
         Caption         =   "Ver Documentos Asociados"
      End
      Begin VB.Menu MnuEDocX 
         Caption         =   "MnuEDocX"
         Index           =   0
      End
      Begin VB.Menu MnuEL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEAddAnuladas 
         Caption         =   "Agregar &Anulaciones"
      End
      Begin VB.Menu MnuEAddNotas 
         Caption         =   "Agregar &Devoluciones"
      End
   End
End
Attribute VB_Name = "frmCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aTexto As String

Dim prmSucursalesDisp As String
Dim mlngIDCuenta As Long

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bCierre_Click()
On Error GoTo errorBT
    
Dim bCorregirSaldo As Boolean

    'Valido el ingreso de los campos para realizar la consulta.------------------------------------------
    If cDisponibilidad.ListIndex = -1 Then
        MsgBox "Debe seleccionar la disponibilidad para realizar el cierre.", vbExclamation, "ATENCIÓN"
        Foco cDisponibilidad: Exit Sub
    End If
    If Not IsDate(dFecha.Value) Then
        MsgBox "La fecha ingresada para realizar el cierre no es correcta.", vbExclamation, "ATENCIÓN"
        dFecha.SetFocus: Exit Sub
    End If
    '------------------------------------------------------------------------------------------------------------
    
    'Saco los Datos de la disponibilidad -------------------------------------------------------------------
    Dim mSucursales As String, mDisponibilidad As Long, mMonedaD As Integer
    
    mDisponibilidad = cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    mSucursales = dis_SucursalesConDisponibilidad(mDisponibilidad)
    
    Cons = "Select * from Disponibilidad Where DisID = " & mDisponibilidad
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    mMonedaD = RsAux!DisMoneda
    mlngIDCuenta = RsAux!DisIDSubrubro
    RsAux.Close
    
    C_KEY_MEMO = "#CC" & mlngIDCuenta & "# "
    '------------------------------------------------------------------------------------------------------------
    
    bCorregirSaldo = True
    
    Dim pblnMovsZureo As Boolean: pblnMovsZureo = False
    'Valido si hay un cierre de caja para la disponibilidad -------------------------------
    pblnMovsZureo = HayMovimientosZureo(C_KEY_MEMO)
    'cons = "Select * from ZureoCGSA.dbo.cceComprobantes " & _
          " Where ComFecha = '" & Format(dFecha.Value, "mm/dd/yyyy") & "'" & _
          " And ComMemo Like '" & C_KEY_MEMO & "%'"
          
    'Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    'If Not rsAux.EOF Then
    '    pblnMovsZureo = True
    'cons = "Delete ZureoCGSA.dbo.cceComprobanteCuenta " & _
          " Where CCuIdComprobante IN (Select ComID from ZureoCGSA.dbo.cceComprobantes " & _
                                        " Where ComFecha = '" & Format(dFecha.Value, "mm/dd/yyyy") & "'" & _
                                        " And ComMemo Like '" & C_KEY_MEMO & "%' )"

    'cons = "Delete ZureoCGSA.dbo.cceComprobantes " & _
        " Where ComFecha = '" & Format(dFecha.Value, "mm/dd/yyyy") & "'" & _
        " And ComMemo Like '" & C_KEY_MEMO & "%'"
    'End If
    'rsAux.Close
    '--------------------------------------------------------------------------------------
    
    Cons = "Select  * from MovimientoDisponibilidad, MovimientoDisponibilidadRenglon" _
           & " Where MDiID = MDRIdMovimiento " _
           & " And MDiFecha = '" & Format(dFecha.Value, "mm/dd/yyyy") & "'" _
           & " And MDiHora = '23:59:59'" _
           & " And MDiTipo = " & paMCIngresosOperativos _
           & " And MDRIdDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If MsgBox("La disponibilidad ya fue cerrada. " & Chr(vbKeyReturn) _
                    & "Si realiza un nuevo cierre se reemplazará la información existente por los nuevos datos (movimientos de disponibilidades y acumulado de ventas por sucursal)." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                    & "Desea continuar con la operación.", vbQuestion + vbYesNo + vbDefaultButton2, "GRABAR") = vbNo Then
                
                RsAux.Close: Exit Sub
        End If
        'Borro los renglones-------------------------------------------------------------------------------------------
        Do While Not RsAux.EOF
            Cons = "Delete MovimientoDisponibilidadRenglon Where MDRIdMovimiento = " & RsAux!MDiId
            cBase.Execute Cons
            
            Cons = "Delete MovimientoDisponibilidad Where MDiId = " & RsAux!MDiId
            cBase.Execute Cons
            
            RsAux.MoveNext
        Loop
        RsAux.Close
        bCorregirSaldo = False
        '------------------------------------------------------------------------------------------------------------------
    Else
        RsAux.Close
        If MsgBox("Confirma procesar el cierre de caja." & Chr(vbKeyReturn) & "Disponibilidad: " & cDisponibilidad.Text, vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    End If
    
    If Not modZureo.CargoDatosEmpresa Then
        If MsgBox("Error al cargar los datos de la empresa. ¿Continúa?" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical + vbYesNo, "Error") = vbNo Then
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = 11
    Dim RsSuc As rdoResultset
        
    cBase.BeginTrans            'Comienzo Transaccion-----------------------------------------------------------------
    On Error GoTo errorET

    AccionCerrar mDisponibilidad, dFecha.Value, bCorregirSaldo, mMonedaD
    
    Cons = "Select * from Sucursal Where SucCodigo IN (" & mSucursales & ")"
    Set RsSuc = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsSuc.EOF
        GeneroAcumuladoVenta mMonedaD, paCodigoDeUsuario, dFecha.Value, RsSuc!SucCodigo
        
        RsSuc.MoveNext
    Loop
    RsSuc.Close
    
    cBase.CommitTrans         'Fin Transaccion-------------------------------------------------------------------------
    Screen.MousePointer = 0
    
    MsgBox "El cierre de caja para la disponibilidad seleccionada y el acumulado de ventas se han realizado con éxito.", vbInformation, "Cierre de Caja"
    
    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Screen.MousePointer = 0
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación. " & Err.Description
End Sub

Private Function CerrarCaja_Carlos(Optional ConMensaje As Boolean = True) As Boolean
    CerrarCaja_Carlos = False
    
    'Valido el ingreso de los campos para realizar la consulta.------------------------------------------
    If cDisponibilidad.ListIndex = -1 Then
        MsgBox "Debe seleccionar la disponibilidad para realizar el cierre.", vbExclamation, "ATENCIÓN"
        Foco cDisponibilidad: Exit Function
    End If
    If Not IsDate(dFecha.Value) Then
        MsgBox "La fecha ingresada para realizar el cierre no es correcta.", vbExclamation, "ATENCIÓN"
        dFecha.SetFocus: Exit Function
    End If
    '------------------------------------------------------------------------------------------------------------
    
    If ConMensaje Then
        If MsgBox("Confirma procesar el cierre de caja." & Chr(vbKeyReturn) & _
                  "Disponibilidad: " & cDisponibilidad.Text, vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Function
    End If
    
    If Not modZureo.CargoDatosEmpresa Then
        If MsgBox("Error al cargar los datos de la empresa. ¿Continúa?" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical + vbYesNo, "Error") = vbNo Then
            Exit Function
        End If
    End If
    
    Screen.MousePointer = 11
    On Error GoTo errorBT

    'Saco los Datos de la disponibilidad -------------------------------------------------------------------
    Dim mDisponibilidad As Long, mMonedaD As Integer
    mDisponibilidad = cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    
    Cons = "Select * from Disponibilidad Where DisID = " & mDisponibilidad
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    mMonedaD = RsAux!DisMoneda
    mlngIDCuenta = RsAux!DisIDSubrubro
    RsAux.Close
    
    C_KEY_MEMO = "#CC" & mlngIDCuenta & "# "
    '------------------------------------------------------------------------------------------------------------
    
    cBase.BeginTrans            'Comienzo Transaccion-----------------------------------------------------------------
    On Error GoTo errorET

    AccionCerrar mDisponibilidad, dFecha.Value, False, mMonedaD
      
    cBase.CommitTrans         'Fin Transaccion-------------------------------------------------------------------------
    Screen.MousePointer = 0
    
    If ConMensaje Then MsgBox "El cierre de caja se ha realizado con éxito.", vbInformation, "Cierre de Caja"
    
    CerrarCaja_Carlos = True
    Exit Function
    
errorBT:
    clsGeneral.OcurrioError modZureo.prmErrorText & vbCrLf & "No se ha podido inicializar la transacción." & vbCrLf & Err.Description
    Screen.MousePointer = 0
    Exit Function
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError modZureo.prmErrorText & vbCrLf & "No se ha podido realizar la transacción. " & vbCrLf & Err.Description
End Function

Private Function HayMovimientosZureo(KeyMemo As String) As Boolean
Dim rsVal As rdoResultset

    Cons = "Select Top 1 * from ZureoCGSA.dbo.cceComprobantes " & _
          " Where ComFecha = '" & Format(dFecha.Value, "mm/dd/yyyy") & "'" & _
          " And ComMemo Like '" & KeyMemo & "%'" & _
          " And (ComEstado <> 9 OR ComEstado IS NULL)"
          
    Set rsVal = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    HayMovimientosZureo = Not rsVal.EOF
    rsVal.Close

End Function

Private Sub AccionCerrar(aDisponibilidad As Long, aFecha As Date, aGeneroSaldo As Boolean, aMoneda As Integer)

Dim bSalida As Boolean
    
    'Paso los movimientos del comercio (vtas, notas, cobranzas, etc)    -----------------------------------------
    modZureo.CGSA_VentasContado aFecha, prmSucursalesDisp, mlngIDCuenta
    modZureo.CGSA_VentasCredito aFecha, prmSucursalesDisp
    modZureo.CGSA_VentasCreditoNotas aFecha, prmSucursalesDisp
    modZureo.CGSA_VentasContadoNotas aFecha, prmSucursalesDisp, mlngIDCuenta
    modZureo.CGSA_VentasContadoNotasE aFecha, prmSucursalesDisp, mlngIDCuenta
    
    modZureo.CGSA_Cobranza aFecha, prmSucursalesDisp, mlngIDCuenta
    modZureo.CGSA_CobranzaMoras aFecha, prmSucursalesDisp, mlngIDCuenta
    modZureo.CGSA_SeñasRecibidas aFecha, prmSucursalesDisp, mlngIDCuenta
    '------------------------------------------------------------------------------------------------------------
    
    Dim pintTipoDoc As Integer, plngCtaAsociada As Long, pcurImporteME As Currency, pdblTC As Double
    Dim pblnYaPasado As Boolean, pstrMemo As String
    
    With vsLista        '<Concepto|>Cantidad|>Importe|IDCuenta|IDComprobante|>Importe ME|Transferencia
        For i = 1 To .Rows - 1
            If .Cell(flexcpValue, i, 2) <> 0 Then
                If .Cell(flexcpValue, i, 2) > 0 Then bSalida = False Else bSalida = True

                plngCtaAsociada = Val(.Cell(flexcpText, i, 3))
                pintTipoDoc = Val(.Cell(flexcpText, i, 4))
                pdblTC = 1
                
                If .Cell(flexcpText, i, 5) <> "" And .Cell(flexcpValue, i, 5) <> .Cell(flexcpValue, i, 2) Then
                    pcurImporteME = Abs(.Cell(flexcpValue, i, 5))
                    If prmMonedaDisp <> prmMonedaContabilidad Then
                        pdblTC = .Cell(flexcpValue, i, 5) / .Cell(flexcpValue, i, 2)
                    End If
                Else
                    pcurImporteME = 0
                End If
                
                If pintTipoDoc <> 0 And plngCtaAsociada <> 0 Then
                
                    pblnYaPasado = False
                    pstrMemo = .Cell(flexcpText, i, 0) & " QMovs:" & .Cell(flexcpText, i, 1)
                    If .Cell(flexcpValue, i, 6) = 1 Then 'Si es Transferencia Valido que no este pasado
                    
                        Dim pdblImporteMContab As Double
                        If prmMonedaDisp = prmMonedaContabilidad Then
                            pdblImporteMContab = Abs(.Cell(flexcpValue, i, 2))
                        Else
                            pdblImporteMContab = Abs(.Cell(flexcpValue, i, 5))
                        End If
                        
                        'cons = "Select Top 1 * from ZureoCGSA.dbo.cceComprobantes " & _
                              " Where ComFecha = '" & Format(dFecha.Value, "mm/dd/yyyy") & "'" & _
                              " And (ComEstado <> 9 OR ComEstado IS NULL) " & _
                              " And ComTotal = " & Abs(.Cell(flexcpValue, I, 2)) & _
                              " And ComMoneda = " & prmMonedaContabilidad & _
                              " And ComTipo = " & pintTipoDoc & _
                              " And ( ComMemo Like '#CC" & mlngIDCuenta & "# " & pstrMemo & "%' OR ComMemo Like '#CC" & plngCtaAsociada & "# " & pstrMemo & "%' )"
                              
                        Cons = "Select Top 1 * from ZureoCGSA.dbo.cceComprobantes " & _
                              " Where ComFecha = '" & Format(dFecha.Value, "mm/dd/yyyy") & "'" & _
                              " And (ComEstado <> 9 OR ComEstado IS NULL) " & _
                              " And ( ( (ComTotal * ComTC) - " & pdblImporteMContab & ")  < 0.001 ) " & _
                              " And ComTipo = " & pintTipoDoc & _
                              " And ( ComMemo Like '#CC" & mlngIDCuenta & "# " & pstrMemo & "%' OR ComMemo Like '#CC" & plngCtaAsociada & "# " & pstrMemo & "%' )"
                                                            
                        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        pblnYaPasado = Not RsAux.EOF
                        RsAux.Close
                    End If

                    If Not pblnYaPasado Then
                        modZureo.fnc_AltaComprobante cBase, aFecha, pintTipoDoc, C_KEY_MEMO & pstrMemo, _
                                     bSalida, plngCtaAsociada, Abs(.Cell(flexcpValue, i, 2)), 0, 0, _
                                                mlngIDCuenta, Abs(.Cell(flexcpValue, i, 2)), dICta1ME:=pcurImporteME, dTC:=pdblTC
                    End If
                End If
            End If
        Next
    End With
    
    'With vsIngresos
    '    For I = 1 To .Rows - 1
    '        If .Cell(flexcpValue, I, 2) <> 0 Then
    '            If .Cell(flexcpValue, I, 2) > 0 Then bSalida = False Else bSalida = True
    '
    '            MovimientoDeCaja .Cell(flexcpData, I, 0), aFecha & " 23:59:59", aDisponibilidad, aMoneda, Abs(.Cell(flexcpValue, I, 2)), .Cell(flexcpText, I, 0), bSalida
    '        End If
    '    Next
    'End With
    
    'Saldo inicial para el dia siguiente a las 00:00:00-----------------------------------------------------------
    If aGeneroSaldo Then
        Cons = "Select * FROM SaldoDisponibilidad " _
                & " Where SDiFecha = '" & Format(aFecha + 1, "mm/dd/yyyy") & "'" _
                & " And SDiHora = '00:00:00'" _
                & " And SDiDisponibilidad = " & aDisponibilidad
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then RsAux.AddNew Else RsAux.Edit
        RsAux!SDiDisponibilidad = aDisponibilidad
        RsAux!SDiFecha = Format(aFecha + 1, "mm/dd/yyyy")
        RsAux!SDiHora = "00:00:00"
        RsAux!SDiSaldo = CCur(lSaldoN.Caption)
        RsAux!SDiComentario = "Saldo Inicial (por cierre de caja)"
        RsAux!SDiUsuario = paCodigoDeUsuario
        RsAux.Update: RsAux.Close
    End If
    
End Sub

Private Sub GeneroAcumuladoVenta(Moneda As Integer, Usuario As Long, Fecha As String, Sucursal As Long)

Dim aContado As Currency, aCredito As Currency, aNContado As Currency, aNCredito As Currency
Dim aContadoIva As Currency, aCreditoIva As Currency, aNContadoIva As Currency, aNCreditoIva As Currency
Dim aContadoCofis As Currency, aCreditoCofis As Currency, aNContadoCofis As Currency, aNCreditoCofis As Currency
Dim aContadoC As Long, aCreditoC As Long, aNContadoC As Long, aNCreditoC As Long

    Screen.MousePointer = 11
    aContado = 0: aCredito = 0: aNContado = 0: aNCredito = 0
    aContadoIva = 0: aCreditoIva = 0: aNContadoIva = 0: aNCreditoIva = 0
    aContadoCofis = 0: aCreditoCofis = 0: aNContadoCofis = 0: aNCreditoCofis = 0
    aContadoC = 0: aCreditoC = 0: aNContadoC = 0: aNCreditoC = 0
    
    
    'Saco datos de Facturas Creditos ---------------------------------------------------------------------------------------------
    Cons = "Select Sum(DocTotal) Suma, Count(*)  Cantidad, Sum(DocIva) Iva, Sum(DocCofis) Cofis from Documento " _
            & " Where DocFecha Between '" & Format(Fecha, "mm/dd/yyyy") & " 00:00' And '" & Format(Fecha, "mm/dd/yyyy") & " 23:59'" _
            & " And DocSucursal = " & Sucursal _
            & " And DocTipo = " & TipoDocumento.Credito _
            & " And DocMoneda = " & Moneda _
            & " And DocAnulado = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aCreditoC = RsAux!Cantidad
        If Not IsNull(RsAux!IVA) Then aCreditoIva = Format(RsAux!IVA, FormatoMonedaP)
        If Not IsNull(RsAux!Suma) Then aCredito = Format(RsAux!Suma, FormatoMonedaP)
        If Not IsNull(RsAux!COFIS) Then aCreditoCofis = Format(RsAux!COFIS, FormatoMonedaP)
        aCredito = aCredito - aCreditoIva - aCreditoCofis
    End If
    RsAux.Close
    
    'Notas de Credito
    Cons = "Select Sum(DocTotal) Suma, Count(*)  Cantidad, Sum(DocIva) Iva, Sum(DocCofis) Cofis from Documento " _
            & " Where DocFecha Between '" & Format(Fecha, "mm/dd/yyyy") & " 00:00' And '" & Format(Fecha, "mm/dd/yyyy") & " 23:59'" _
            & " And DocSucursal = " & Sucursal _
            & " And DocTipo = " & TipoDocumento.NotaCredito _
            & " And DocMoneda = " & Moneda _
            & " And DocAnulado = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aNCreditoC = RsAux!Cantidad
        If Not IsNull(RsAux!IVA) Then aNCreditoIva = Format(RsAux!IVA, FormatoMonedaP)
        If Not IsNull(RsAux!Suma) Then aNCredito = Format(RsAux!Suma, FormatoMonedaP)
        If Not IsNull(RsAux!COFIS) Then aNCreditoCofis = Format(RsAux!COFIS, FormatoMonedaP)
        aNCredito = aNCredito - aNCreditoIva - aNCreditoCofis
    End If
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------------------------
    
    'Saco datos de Facturas Contados
    Cons = "Select Sum(DocTotal) Suma, Count(*)  Cantidad, Sum(DocIva) Iva, Sum(DocCofis) Cofis from Documento " _
            & " Where DocFecha Between '" & Format(Fecha, "mm/dd/yyyy") & " 00:00' And '" & Format(Fecha, "mm/dd/yyyy") & " 23:59'" _
            & " And DocSucursal = " & Sucursal _
            & " And DocTipo = " & TipoDocumento.Contado _
            & " And DocMoneda = " & Moneda _
            & " And DocAnulado = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aContadoC = RsAux!Cantidad
        If Not IsNull(RsAux!IVA) Then aContadoIva = Format(RsAux!IVA, FormatoMonedaP)
        If Not IsNull(RsAux!Suma) Then aContado = Format(RsAux!Suma, FormatoMonedaP)
        If Not IsNull(RsAux!COFIS) Then aContadoCofis = Format(RsAux!COFIS, FormatoMonedaP)
        aContado = aContado - aContadoIva - aContadoCofis
    End If
    RsAux.Close
    
    'Saco datos de Notas Contados
    Cons = "Select Sum(DocTotal) Suma, Count(*)  Cantidad, Sum(DocIva) Iva, Sum(DocCofis) Cofis from Documento " _
            & " Where DocFecha Between '" & Format(Fecha, "mm/dd/yyyy") & " 00:00' And '" & Format(Fecha, "mm/dd/yyyy") & " 23:59'" _
            & " And DocSucursal = " & Sucursal _
            & " And DocTipo = " & TipoDocumento.NotaDevolucion _
            & " And DocMoneda = " & Moneda _
            & " And DocAnulado = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aNContadoC = RsAux!Cantidad
        If Not IsNull(RsAux!IVA) Then aNContadoIva = Format(RsAux!IVA, FormatoMonedaP)
        If Not IsNull(RsAux!Suma) Then aNContado = Format(RsAux!Suma, FormatoMonedaP)
        If Not IsNull(RsAux!COFIS) Then aNContadoCofis = Format(RsAux!COFIS, FormatoMonedaP)
    End If
    RsAux.Close
    
    'Saco datos de Notas Especiales
    Cons = "Select Sum(DocTotal) Suma, Count(*)  Cantidad, Sum(DocIva) Iva, Sum(DocCofis) Cofis from Documento " _
            & " Where DocFecha Between '" & Format(Fecha, "mm/dd/yyyy") & " 00:00' And '" & Format(Fecha, "mm/dd/yyyy") & " 23:59'" _
            & " And DocSucursal = " & Sucursal _
            & " And DocTipo = " & TipoDocumento.NotaEspecial _
            & " And DocMoneda = " & Moneda _
            & " And DocAnulado = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aNContadoC = aNContadoC + RsAux!Cantidad
        If Not IsNull(RsAux!IVA) Then aNContadoIva = aNContadoIva + Format(RsAux!IVA, FormatoMonedaP)
        If Not IsNull(RsAux!Suma) Then aNContado = aNContado + Format(RsAux!Suma, FormatoMonedaP)
        If Not IsNull(RsAux!COFIS) Then aNContadoCofis = aNContadoCofis + Format(RsAux!COFIS, FormatoMonedaP)
    End If
    RsAux.Close
    aNContado = aNContado - aNContadoIva - aNContadoCofis
    '-----------------------------------------------------------------------------------------------------------------------
               
    GraboBDAcumuladoVenta TipoDocumento.Contado, aContado, aContadoIva, aContadoC, aNContado, aNContadoIva, aNContadoC, Moneda, Usuario, Fecha, Sucursal, aContadoCofis, aNContadoCofis
    GraboBDAcumuladoVenta TipoDocumento.Credito, aCredito, aCreditoIva, aCreditoC, aNCredito, aNCreditoIva, aNCreditoC, Moneda, Usuario, Fecha, Sucursal, aCreditoCofis, aNCreditoCofis
    
End Sub

Private Sub GraboBDAcumuladoVenta(Tipo As Integer, Importe As Currency, IVA As Currency, Cantidad As Long, ImporteN As Currency, IvaN As Currency, CantidadN As Long, _
                                                       Moneda As Integer, Usuario As Long, Fecha As String, Sucursal As Long, COFIS As Currency, CofisN As Currency)


    Cons = "Select * from AcumuladoVenta" _
            & " Where AVeFecha = '" & Format(Fecha, "mm/dd/yyyy") & "'" _
            & " And AVeTipo = " & Tipo _
            & " And AVeSucursal = " & Sucursal _
            & " And AVeMoneda = " & Moneda
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.AddNew
        RsAux!AVeFecha = Format(Fecha, "mm/dd/yyyy")
        RsAux!AVeTipo = Tipo
        RsAux!AVeSucursal = Sucursal
        RsAux!AVeMoneda = Moneda
        RsAux!AveNeto = Importe
        RsAux!AveIva = IVA
        RsAux!AVeCantidad = Cantidad
        RsAux!AveNNeto = ImporteN
        RsAux!AveNIva = IvaN
        RsAux!AVeNCantidad = CantidadN
        RsAux!AveUsuario = Usuario
        RsAux!AVeHora = Format(gFechaServidor, sqlFormatoFH)
        
        RsAux!AveCofis = COFIS
        RsAux!AveNCofis = CofisN
        RsAux.Update
    Else
        RsAux.Edit
        RsAux!AveNeto = Importe
        RsAux!AveIva = IVA
        RsAux!AVeCantidad = Cantidad
        RsAux!AveNNeto = ImporteN
        RsAux!AveNIva = IvaN
        RsAux!AVeNCantidad = CantidadN
        RsAux!AveUsuario = Usuario
        RsAux!AVeHora = Format(gFechaServidor, sqlFormatoFH)
        
        RsAux!AveCofis = COFIS
        RsAux!AveNCofis = CofisN
        RsAux.Update
    End If
    
    RsAux.Close
    
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub AccionConsultar()

Dim aValor As Currency

    'Valido el ingreso de los campos para realizar la consulta.------------------------------------------
    If cDisponibilidad.ListIndex = -1 Then
        MsgBox "Debe seleccionar una disponibilidad para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco cDisponibilidad: Exit Sub
    End If
    If Not IsDate(dFecha.Value) Then
        MsgBox "La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN"
        dFecha.SetFocus: Exit Sub
    End If
    '------------------------------------------------------------------------------------------------------------

    Screen.MousePointer = 11
       
    'Saco la moneda de la disponibilidad-------------------------------------------------------------------
    Dim aDisponibilidad As Long, aMonedaD As Long
    aDisponibilidad = CLng(cDisponibilidad.ItemData(cDisponibilidad.ListIndex))
    Cons = "Select * from Disponibilidad Where DisId = " & aDisponibilidad
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aMonedaD = RsAux!DisMoneda
    mlngIDCuenta = RsAux!DisIDSubrubro
    RsAux.Close
    
    C_KEY_MEMO = "#CC" & mlngIDCuenta & "# "
    '------------------------------------------------------------------------------------------------------------
    
    prmMonedaDisp = aMonedaD
    prmSucursalesDisp = dis_SucursalesConDisponibilidad(aDisponibilidad)
    
    'Cargo la Lista de ingresos------------------------------------------------------
    CargoIngresosContado 1, prmSucursalesDisp, aMonedaD
    CargoIngresosCuotas 2, 3, 6, prmSucursalesDisp, aMonedaD
    CargoEgresosContado 4, prmSucursalesDisp, aMonedaD
    CargoEgresosNEspecial 5, prmSucursalesDisp, aMonedaD
    
    aValor = 0
    For i = 1 To vsIngresos.Rows - 1
        aValor = aValor + vsIngresos.Cell(flexcpValue, i, 2)
    Next
    lTIngreso.Caption = Format(aValor, FormatoMonedaP)
    '-------------------------------------------------------------------------------------
    
    'Cargo la lista de Movimientos----------------------------------------------------
    CargoMovimientosDeCaja aDisponibilidad, aMonedaD
    aValor = 0
    For i = 1 To vsLista.Rows - 1
        aValor = aValor + vsLista.Cell(flexcpValue, i, 2)
    Next
    lTMovimiento.Caption = Format(aValor, FormatoMonedaP)
    '-------------------------------------------------------------------------------------
        
    lTotal.Caption = Format(CCur(lTIngreso.Caption) + CCur(lTMovimiento.Caption), FormatoMonedaP)
    lSubtotal.Caption = lTotal.Caption
    
    
    'Busco el saldo anterior a la fecha ingresada---------------------------------------------------------
    Cons = "Select * from SaldoDisponibilidad " _
           & " Where SDiFecha = (Select Max(SDiFecha) from SaldoDisponibilidad " _
                            & " Where SDiDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
                            & " And SDiFecha <= '" & Format(dFecha.Value, sqlFormatoF) & "')" _
          & " And SDiDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        lSaldoA.Caption = Format(RsAux!SDiSaldo, FormatoMonedaP)
    Else
        lSaldoA.Caption = "0.00"
    End If
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------------
    lSaldoN.Caption = Format(CCur(lSaldoA.Caption) + CCur(lSubtotal.Caption), FormatoMonedaP)
    
    CargoDatosArqueo aMonedaD
    CargoDatosErrores prmSucursalesDisp, aMonedaD
    
    bCierre.Enabled = True
    Screen.MousePointer = 0
    
End Sub

Private Sub bImprimir_Click()
    AccionImprimir
End Sub

Private Sub bNoFiltros_Click()
    cDisponibilidad.Text = ""
    Exit Sub
    
Dim pdatF1 As Date, pdatF2 As Date, pstrTXT As String

    pstrTXT = InputBox("Procesar los cierres de caja Desde:", "Cierres de Caja")
    If pstrTXT = "" Then Exit Sub
    pdatF1 = CDate(pstrTXT)

    pstrTXT = InputBox("Procesar los cierres de caja Hasta:", "Cierres de Caja")
    If pstrTXT = "" Then Exit Sub
    pdatF2 = CDate(pstrTXT)

    If pdatF1 > pdatF2 Then Exit Sub
    
    If MsgBox("¿Confirma procesar los cierres de caja desde " & pdatF1 & " al " & pdatF2 & " ?", vbQuestion + vbYesNo, "Procesar") = vbNo Then Exit Sub
    
    Dim pblnSalir As Boolean
    pblnSalir = False
    Do While pdatF1 <= pdatF2 Or pblnSalir
        
        dFecha.Value = pdatF1
        Call bConsultar_Click
        Me.Refresh
        
        pblnSalir = Not CerrarCaja_Carlos(ConMensaje:=False)
        pdatF1 = DateAdd("d", 1, pdatF1)
    Loop

    MsgBox "Cierres de Caja procesados.", vbInformation, "Cierre de Caja"
    
End Sub

Private Sub cDisponibilidad_Change()
    bCierre.Enabled = False
End Sub

Private Sub cDisponibilidad_Click()
    bCierre.Enabled = False
End Sub

Private Sub cDisponibilidad_GotFocus()
    cDisponibilidad.SelStart = 0: cDisponibilidad.SelLength = Len(cDisponibilidad.Text)
End Sub

Private Sub cDisponibilidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then dFecha.SetFocus
End Sub

Private Sub dFecha_Change()
    bCierre.Enabled = False
End Sub

Private Sub dFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco bConsultar
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Screen.MousePointer = 11
    ObtengoSeteoForm Me, 105, 105, 9960, 7950
    
    Dim aSucursal As String
    aSucursal = CargoParametrosSucursal
    
    Status.Panels("sucursal") = "Sucursal: " & aSucursal
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        
    picBotones.BorderStyle = 0
    FechaDelServidor
    
    'Cargo las Disponibilidades-------------------------------------------------------------------
    Dim mDisME As String
    Cons = "Select * from Sucursal Where SucDisponibilidadME Is not Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If Trim(RsAux!SucDisponibilidadME) <> "" Then
            If mDisME <> "" Then
                If Right(mDisME, 1) <> "," Then mDisME = mDisME & ","
            End If
            mDisME = mDisME & Trim(RsAux!SucDisponibilidadME)
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Cons = "Select DisID, DisNombre from Disponibilidad " _
            & " Where DisID In (Select SucDisponibilidad from Sucursal) "
    If Trim(mDisME) <> "" Then Cons = Cons & " OR DisID In (" & mDisME & ") "
    Cons = Cons & " Order by DisNombre"
    
    CargoCombo Cons, cDisponibilidad, ""
    BuscoCodigoEnCombo cDisponibilidad, paDisponibilidad
    '-----------------------------------------------------------------------------------------
    
    dFecha.Value = Format(gFechaServidor, "dd/mm/yyyy")
    
    InicializoGrillas
    PreparoFormulario
    bCierre.Enabled = False
    
    Set tabCaja.SelectedItem = tabCaja.Tabs("cierre")
    
End Sub

Private Sub PreparoFormulario()

    With img1
        bConsultar.Picture = .ListImages("consultar").ExtractIcon
        bCierre.Picture = .ListImages("ok").ExtractIcon
        bImprimir.Picture = .ListImages("print").ExtractIcon
        bNoFiltros.Picture = .ListImages("limpiar").ExtractIcon
        bCancelar.Picture = .ListImages("salir").ExtractIcon
    End With
    
    'INGRESOS
    With vsIngresos
        .AddItem "Ventas Contado": .Cell(flexcpData, .Rows - 1, 0) = paMCIngresosOperativos
        .AddItem "Cobranza de Cuotas": .Cell(flexcpData, .Rows - 1, 0) = paMCIngresosOperativos
        .AddItem "Cobranza de Morosidades": .Cell(flexcpData, .Rows - 1, 0) = paMCIngresosOperativos
        .AddItem "Notas de Devolución": .Cell(flexcpData, .Rows - 1, 0) = paMCIngresosOperativos
        .AddItem "Notas Contado Especial": .Cell(flexcpData, .Rows - 1, 0) = paMCIngresosOperativos
        .AddItem "Señas Recibidas": .Cell(flexcpData, .Rows - 1, 0) = paMCIngresosOperativos
    End With
    
    'SUBTOTALES Y TOTALES
    lTIngreso.Caption = "N/D": lTMovimiento.Caption = "N/D": lTotal.Caption = "N/D"
    lSaldoA.Caption = "N/D": lSubtotal.Caption = "N/D": lSaldoN.Caption = "N/D"

    LimpioDatosArqueo
    LimpioDatosErrores
End Sub

Private Sub CargoIngresosContado(Row As Integer, idsSucursales As String, Moneda As Long)

    On Error GoTo errConsulta
    Cons = "Select Sum(DocTotal) Suma, Count(*)  Cantidad from Documento " _
           & " Where DocFecha Between '" & Format(dFecha.Value, "mm/dd/yyyy") & " 00:00' And '" & Format(dFecha.Value, "mm/dd/yyyy") & " 23:59'"
    
    Cons = Cons & " And DocSucursal IN (" & idsSucursales & ")"
    
    Cons = Cons & " And DocTipo = " & TipoDocumento.Contado _
                       & " And DocMoneda = " & Moneda _
                       & " And DocAnulado = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    With vsIngresos
        If Not RsAux.EOF Then
            If Not IsNull(RsAux!Cantidad) Then .Cell(flexcpText, Row, 1) = RsAux!Cantidad Else .Cell(flexcpText, Row, 1) = 0
            If Not IsNull(RsAux!Suma) Then .Cell(flexcpText, Row, 2) = Format(RsAux!Suma, FormatoMonedaP) Else .Cell(flexcpText, Row, 2) = "0.00"
        Else
            .Cell(flexcpText, Row, 1) = "0"
            .Cell(flexcpText, Row, 2) = "0.00"
        End If
    End With
    
    RsAux.Close
    Exit Sub
    
errConsulta:
    clsGeneral.OcurrioError "Error al procesar las ventas contado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoIngresosCuotas(RowCta As Integer, RowMora As Integer, RowSenia As Integer, idsSucursales As String, Moneda As Long)

Dim aCantidadC As Long, aSumaC As Currency

    On Error GoTo errConsulta
    With vsIngresos
    
    'Consulta para sacar los montos de los recibos (Hay que restarle las moras)------------------------------------------
    Cons = "Select Sum(DocTotal) Suma, Count(*)  Cantidad from Documento " _
            & " Where DocFecha Between '" & Format(dFecha.Value, "mm/dd/yyyy") & " 00:00:00' And '" & Format(dFecha.Value, "mm/dd/yyyy") & " 23:59:59'"
            
    Cons = Cons & " And DocSucursal IN (" & idsSucursales & ")"
    
    Cons = Cons & " And DocTipo = " & TipoDocumento.ReciboDePago _
                        & " And DocMoneda = " & Moneda _
                        & " And DocAnulado = 0"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aCantidadC = 0: aSumaC = 0
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aCantidadC = RsAux!Cantidad
        If Not IsNull(RsAux!Suma) Then aSumaC = RsAux!Suma
    End If
    RsAux.Close
    
    'Consulta para sacar las moras pagas ------------------------------------------------------------------------------------
    Cons = "Select Sum(DPaMora) Suma, Count(*)  Cantidad from Documento, DocumentoPago " _
            & " Where DocFecha Between '" & Format(dFecha.Value, "mm/dd/yyyy") & " 00:00' And '" & Format(dFecha.Value, "mm/dd/yyyy") & " 23:59'"
    
    Cons = Cons & " And DocSucursal IN (" & idsSucursales & ")"
    
    Cons = Cons & " And DocTipo = " & TipoDocumento.ReciboDePago _
                        & " And DocMoneda = " & Moneda _
                        & " And DocAnulado = 0" _
                        & " And DocCodigo = DPaDocQSalda and DPaMora <> 0"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then .Cell(flexcpText, RowMora, 1) = RsAux!Cantidad Else .Cell(flexcpText, RowMora, 1) = "0"
        If Not IsNull(RsAux!Suma) Then .Cell(flexcpText, RowMora, 2) = Format(RsAux!Suma, FormatoMonedaP) Else .Cell(flexcpText, RowMora, 2) = "0.00"
    Else
        .Cell(flexcpText, RowMora, 1) = "0"
        .Cell(flexcpText, RowMora, 2) = "0.00"
    End If
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------
        
    'Consulta para sacar los pagos a Perdidas y Sacarlos de la cobranza y ponerlos en Mora-------------------------------
    'Como la Mora ya la Sume --> Tengo que restarle la amortizacion a la Cbza Ctas
    Cons = "Select Sum(DPAAmortizacion) Suma, Count(*)  Cantidad from Documento, DocumentoPago, Credito " _
            & " Where DocFecha Between '" & Format(dFecha.Value, "mm/dd/yyyy") & " 00:00' And '" & Format(dFecha.Value, "mm/dd/yyyy") & " 23:59'"
    
    Cons = Cons & " And DocSucursal IN (" & idsSucursales & ")"
    
    Cons = Cons & " And DocTipo = " & TipoDocumento.ReciboDePago _
                        & " And DocMoneda = " & Moneda _
                        & " And DocAnulado = 0" _
                        & " And DocCodigo = DPaDocQSalda And DPaDocASaldar = CreFactura" _
                        & " And CreTipo = " & TipoCredito.Incobrable
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then
            .Cell(flexcpText, RowMora, 1) = .Cell(flexcpValue, RowMora, 1) + RsAux!Cantidad
            aCantidadC = aCantidadC - RsAux!Cantidad
        End If
        If Not IsNull(RsAux!Suma) Then .Cell(flexcpText, RowMora, 2) = Format(.Cell(flexcpValue, RowMora, 2) + RsAux!Suma, FormatoMonedaP)
            
    End If
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------
    
    .Cell(flexcpText, RowCta, 1) = aCantidadC
    .Cell(flexcpText, RowCta, 2) = Format(aSumaC - .Cell(flexcpValue, RowMora, 2), FormatoMonedaP)
    
    
    'Consulta para sacar los recibos que no estan asignados a facturas (senias recibidas)-------------------------------------
    Cons = "Select Sum(DocTotal) Suma, Count(*)  Cantidad from Documento Left Outer join DocumentoPago on DocCodigo = DPaDocQSalda " _
            & " Where DocFecha Between '" & Format(dFecha.Value, "mm/dd/yyyy") & " 00:00:00' And '" & Format(dFecha.Value, "mm/dd/yyyy") & " 23:59:59'"
            
    Cons = Cons & " And DocSucursal IN (" & idsSucursales & ")"
    'Cons = Cons & " And DocSucursal IN (Select SucCodigo from Sucursal Where SucDisponibilidad = " & Disponibilidad & " OR SucDisponibilidadME = " & Disponibilidad & ")"
    
    Cons = Cons & " And DocTipo = " & TipoDocumento.ReciboDePago _
                        & " And DocMoneda = " & Moneda _
                        & " And DocAnulado = 0" _
                        & " And DPaDocQSalda Is null And DPaDocASaldar Is null"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aCantidadC = 0: aSumaC = 0
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aCantidadC = RsAux!Cantidad
        If Not IsNull(RsAux!Suma) Then aSumaC = RsAux!Suma
        
        .Cell(flexcpText, RowCta, 1) = .Cell(flexcpValue, RowCta, 1) - aCantidadC
        .Cell(flexcpText, RowCta, 2) = Format(.Cell(flexcpValue, RowCta, 2) - Format(aSumaC, FormatoMonedaP), FormatoMonedaP)
    End If
    RsAux.Close
    
    .Cell(flexcpText, RowSenia, 1) = aCantidadC
    .Cell(flexcpText, RowSenia, 2) = Format(aSumaC, FormatoMonedaP)
    '-------------------------------------------------------------------------------------------------------------------------------------------------
    End With
    Exit Sub
    
errConsulta:
    clsGeneral.OcurrioError "Error al procesar los recibos de pago.", Err.Description
End Sub

Private Sub CargoEgresosContado(Row As Integer, idsSucursales As String, Moneda As Long)

    On Error GoTo errConsulta
    Cons = "Select Sum(DocTotal) Suma, Count(*)  Cantidad from Documento " _
            & " Where DocFecha Between '" & Format(dFecha.Value, "mm/dd/yyyy") & " 00:00' And '" & Format(dFecha.Value, "mm/dd/yyyy") & " 23:59'"
            
    Cons = Cons & " And DocSucursal IN (" & idsSucursales & ")"
        
    Cons = Cons & " And DocTipo = " & TipoDocumento.NotaDevolucion _
                        & " And DocMoneda = " & Moneda _
                        & " And DocAnulado = 0"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    With vsIngresos
    
    If Not RsAux.EOF Then
        .Cell(flexcpForeColor, Row, 1, , 2) = Colores.Rojo
        If Not IsNull(RsAux!Cantidad) Then .Cell(flexcpText, Row, 1) = RsAux!Cantidad Else: .Cell(flexcpText, Row, 1) = "0"
        If Not IsNull(RsAux!Suma) Then .Cell(flexcpText, Row, 2) = Format(RsAux!Suma * -1, FormatoMonedaP) Else: .Cell(flexcpText, Row, 2) = "0.00"
    Else
        .Cell(flexcpText, Row, 1) = "0"
        .Cell(flexcpText, Row, 2) = "0.00"
    End If
    
    RsAux.Close
    End With
    Exit Sub
    
errConsulta:
    clsGeneral.OcurrioError "Error al procesar las notas de devolución."
End Sub

Private Sub CargoEgresosNEspecial(Row As Integer, idsSucursales As String, Moneda As Long)

    On Error GoTo errConsulta
    Cons = "Select Sum(DocTotal) Suma, Count(*)  Cantidad from Documento " _
            & " Where DocFecha Between '" & Format(dFecha.Value, "mm/dd/yyyy") & " 00:00' And '" & Format(dFecha.Value, "mm/dd/yyyy") & " 23:59'"
            
    Cons = Cons & " And DocSucursal IN (" & idsSucursales & ")"
    
    Cons = Cons & " And DocTipo = " & TipoDocumento.NotaEspecial _
                        & " And DocMoneda = " & Moneda _
                        & " And DocAnulado = 0"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    With vsIngresos
    If Not RsAux.EOF Then
        .Cell(flexcpForeColor, Row, 1, , 2) = Colores.Rojo
        If Not IsNull(RsAux!Cantidad) Then .Cell(flexcpText, Row, 1) = RsAux!Cantidad Else .Cell(flexcpText, Row, 1) = "0"
        If Not IsNull(RsAux!Suma) Then .Cell(flexcpText, Row, 2) = Format(RsAux!Suma * -1, FormatoMonedaP) Else .Cell(flexcpText, Row, 2) = "0.00"
    Else
        .Cell(flexcpText, Row, 1) = "0"
        .Cell(flexcpText, Row, 2) = "0.00"
    End If
    RsAux.Close
    
    End With
    Exit Sub
    
errConsulta:
    clsGeneral.OcurrioError "Error al procesar las notas especiales."
End Sub

Private Function CargoMovimientosZureo() As Currency
On Error GoTo errZureo

    CargoMovimientosZureo = 0
    '1) Cargo los movimientos de la disponibiliad según Zureo   -------------------------------------------
    Dim objGeneric As New clsDBFncs
    Dim rdoCZureo As rdoConnection
    If objGeneric.get_Connection(rdoCZureo, "ORG01", 10) Then
    
        Dim pcurTotalZureo As Currency, pbytSigno As Byte, pbytAlDebe As Byte
        Dim objAcc As New clsAccSaldos
        Set objAcc.Connect = rdoCZureo
        
        objAcc.CalcularMovimientos 1, 10, dFecha.Value, dFecha.Value, 1, CStr(mlngIDCuenta), 0, 0
        pcurTotalZureo = objAcc.GetCuenta(1).SaldoMCuenta
        pbytSigno = objAcc.GetCuenta(1).Signo
        
        pbytAlDebe = objAcc.SaldoAValorDH(pbytSigno, CDbl(pcurTotalZureo)).AlDebe
        pcurTotalZureo = objAcc.SaldoAValorDH(pbytSigno, CDbl(pcurTotalZureo)).Valor
        
        If pbytAlDebe = 0 Then pcurTotalZureo = pcurTotalZureo * -1
        
        Set objAcc = Nothing
        
        With vsLista
            .AddItem "Totales Zureo"
            .Cell(flexcpData, .Rows - 1, 0) = 0
            
            .Cell(flexcpText, .Rows - 1, 1) = 1
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(pcurTotalZureo, FormatoMonedaP)
            If (pbytAlDebe = 0) Then .Cell(flexcpForeColor, .Rows - 1, 1, , 2) = Colores.Rojo
        End With
        
        CargoMovimientosZureo = pcurTotalZureo
    End If
    rdoCZureo.Close
    Set objGeneric = Nothing

errZureo:
End Function

Private Sub CargoMovimientosDeCaja(Disponibilidad As Long, Moneda As Long)

    On Error GoTo errConsulta
    vsLista.Rows = 1
        
'    cons = "Select TMDCodigo, TMDNombre, TMDSubRubro, TMDComprobante, Count(*) Cantidad, Sum(MDRDebe) Debe, Sum(MDRHaber) Haber  " _
            & " From MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, TipoMovDisponibilidad " _
            & " Where MDiFecha = '" & Format(dFecha.Value, "mm/dd/yyyy") & "'" _
            & " And MDiId = MDRIdMovimiento" & " And MDiTipo <> " & paMCIngresosOperativos
            
'    cons = cons & " And MDRIdDisponibilidad = " & Disponibilidad
            
'    cons = cons & " And MDiTipo = TMDCodigo" _
                & " Group by TMDCodigo, TMDNombre, TMDSubRubro, TMDComprobante"
            
            
    Cons = "Select TMDCodigo, TMDNombre, IsNull(DisIdSubrubro, TMDSubRubro) TMDSubRubro, TMDComprobante, TMDNoPasarAZureo, TMDTransferencia, " & _
                " Count(*) Cantidad, Sum(MDR1.MDRDebe) Debe, Sum(MDR1.MDRHaber) Haber, " & _
                " Sum(MDR2.MDRDebe) Haber2, Sum(MDR2.MDRHaber) Debe2 " & _
          " From CGSA.dbo.MovimientoDisponibilidad" & _
                " Left Outer JOIN CGSA.dbo.MovimientoDisponibilidadRenglon as MDR2 ON MDiId = MDR2.MDRIdMovimiento AND MDR2.MDRIdDisponibilidad <> " & Disponibilidad & _
                " Left Outer Join CGSA.dbo.Disponibilidad ON MDR2.MDRIdDisponibilidad = DisID " & _
          " , CGSA.dbo.MovimientoDisponibilidadRenglon MDR1, CGSA.dbo.TipoMovDisponibilidad " & _
          " Where MDiFecha = '" & Format(dFecha.Value, "mm/dd/yyyy") & "'" & _
          " And MDiId = MDR1.MDRIdMovimiento And MDiTipo <> " & paMCIngresosOperativos & _
          " And MDR1.MDRIdDisponibilidad = " & Disponibilidad & " And MDiTipo = TMDCodigo" & _
          " Group by TMDCodigo, TMDNombre, TMDSubRubro, TMDComprobante, TMDNoPasarAZureo, DisIDSubRubro, TMDTransferencia" & _
          " Order by TMDNombre"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Dim aValor As Long, aDebe As Currency, aHaber As Currency
    Dim bMSG As Boolean: bMSG = False
    
    Do While Not RsAux.EOF
        With vsLista    '<Concepto|>Cantidad|>Importe|IDCuenta|IDComprobante|>Importe ME|Transferencia
            .AddItem Trim(RsAux!TMDNombre)
            aValor = RsAux!TMDCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!Cantidad)
            
            If Not IsNull(RsAux!Haber) Then aHaber = RsAux!Haber Else aHaber = 0
            If Not IsNull(RsAux!Debe) Then aDebe = RsAux!Debe Else aDebe = 0
            .Cell(flexcpText, .Rows - 1, 2) = Format(aDebe - aHaber, FormatoMonedaP)
            If (aDebe - aHaber) < 0 Then .Cell(flexcpForeColor, .Rows - 1, 1, , 2) = Colores.Rojo
            
            If Not IsNull(RsAux!TMDSubRubro) Then .Cell(flexcpText, .Rows - 1, 3) = RsAux!TMDSubRubro Else bMSG = True
            If Not IsNull(RsAux!TMDComprobante) Then .Cell(flexcpText, .Rows - 1, 4) = RsAux!TMDComprobante Else bMSG = True
            
            If Not IsNull(RsAux!TMDNoPasarAZureo) Then
                If RsAux!TMDNoPasarAZureo = 1 Then
                    .Cell(flexcpText, .Rows - 1, 3) = 0: .Cell(flexcpText, .Rows - 1, 4) = 0
                End If
            End If
            
            If Not IsNull(RsAux!Haber2) Then aHaber = RsAux!Haber2 Else aHaber = 0
            If Not IsNull(RsAux!Debe2) Then aDebe = RsAux!Debe2 Else aDebe = 0
            If (aDebe - aHaber) <> 0 And Abs(aDebe - aHaber) <> Abs(.Cell(flexcpValue, .Rows - 1, 2)) Then
                .Cell(flexcpText, .Rows - 1, 5) = Format(aDebe - aHaber, FormatoMonedaP)
            End If
            If (aDebe - aHaber) < 0 Then .Cell(flexcpForeColor, .Rows - 1, 5) = Colores.Rojo
            
            If Not IsNull(RsAux!TMDTransferencia) Then
                If RsAux!TMDTransferencia = 1 Then .Cell(flexcpText, .Rows - 1, 6) = 1
            End If
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close

    'If bMSG Then
    '    MsgBox "Hay Tipos de Movimientos que no tienen asociado una cuenta o comprobante en Zureo." & vbCrLf & _
                "Se recomienda NO CERRAR LA CAJA !", vbExclamation, "Posible ERROR "
    'End If

    
    Dim pcurTotalZureo As Currency, pblnMovsZureo As Boolean
    pcurTotalZureo = 0: pblnMovsZureo = False
    If dFecha.Value > CDate("20/10/2007") Then
        pcurTotalZureo = CargoMovimientosZureo
        
        'Resto los movimientos que fueron pasados a Zureo de CGSA (x cierres) y que afectan la SQL de Movimientos ------------
        Cons = "Select Sum(CCuImporteCuenta) from ZureoCGSA.dbo.cceComprobantes, ZureoCGSA.dbo.cceComprobanteCuenta " & _
            " Where ComID = CCuIDComprobante " & _
            " And ComFecha = '" & Format(dFecha.Value, "mm/dd/yyyy") & "'" & _
            " And ComMemo Like '#CC[1-9]%#%' And (ComEstado <> 9 OR ComEstado IS NULL)" & _
            " And CCuIDCuenta = " & mlngIDCuenta & " And CCuDebe IS NOT NULL" & _
                " UNION ALL " & _
            "Select Sum(CCuImporteCuenta) * -1 from ZureoCGSA.dbo.cceComprobantes, ZureoCGSA.dbo.cceComprobanteCuenta " & _
            " Where ComID = CCuIDComprobante " & _
            " And ComFecha = '" & Format(dFecha.Value, "mm/dd/yyyy") & "'" & _
            " And ComMemo Like '#CC[1-9]%#%' And (ComEstado <> 9 OR ComEstado IS NULL)" & _
            " And CCuIDCuenta = " & mlngIDCuenta & " And CCuHaber IS NOT NULL"
          
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            Do While Not RsAux.EOF
                If Not IsNull(RsAux(0)) Then
                    pcurTotalZureo = pcurTotalZureo - RsAux(0)
                End If
                RsAux.MoveNext
            Loop
        
            vsLista.Cell(flexcpText, vsLista.Rows - 1, 2) = Format(pcurTotalZureo, FormatoMonedaP)
        End If
        RsAux.Close
    End If


'    If pblnMovsZureo And pcurTotalZureo <> 0 Then   'Si hay movs cierre se los resto al Totales Zureo
'        Dim pcurSuma As Currency
'        pcurSuma = 0
'        For I = vsLista.FixedRows To vsLista.Rows - 2
'            'Tengo que sumar (para restar) los que yo paso !!
'
'            If vsLista.Cell(flexcpValue, I, 2) <> 0 And vsLista.Cell(flexcpValue, I, 3) <> 0 And vsLista.Cell(flexcpText, I, 4) <> 0 Then
'                pcurSuma = pcurSuma + vsLista.Cell(flexcpValue, I, 2)
'            End If
'        Next
'        pcurTotalZureo = pcurTotalZureo - pcurSuma - CCur(lTIngreso.Caption)
'        vsLista.Cell(flexcpText, vsLista.Rows - 1, 2) = Format(pcurTotalZureo, FormatoMonedaP)
'    End If
    
    Exit Sub
    
errConsulta:
    clsGeneral.OcurrioError "Error al cargar los movimientos de caja.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fFiltros.Width = Me.Width - 320
    
    With fResumen
        .Top = Me.ScaleHeight - .Height - Status.Height - 100
        .Width = fFiltros.Width
    End With
    
    With tabCaja
        .Top = fFiltros.Top + fFiltros.Height + 60
        .Left = fFiltros.Left
        .Width = fFiltros.Width
        .Height = fResumen.Top - .Top - 60
    End With
    
    With picCierre
        .Left = tabCaja.ClientLeft
        .Top = tabCaja.ClientTop
        .Width = tabCaja.ClientWidth
        .Height = tabCaja.ClientHeight
        .BorderStyle = 0
    End With
    With picArqueo
        .Left = tabCaja.ClientLeft
        .Top = tabCaja.ClientTop
        .Width = tabCaja.ClientWidth
        .Height = tabCaja.ClientHeight
        .BorderStyle = 0
    End With
    With picErrores
        .Left = tabCaja.ClientLeft
        .Top = tabCaja.ClientTop
        .Width = tabCaja.ClientWidth
        .Height = tabCaja.ClientHeight
        .BorderStyle = 0
    End With
    
    lbIngresos.Left = picCierre.ScaleLeft
    lbIngresos.Width = picCierre.ScaleWidth
    vsIngresos.Left = picCierre.ScaleLeft
    vsIngresos.Width = picCierre.ScaleWidth
    
    lbMovimientos.Width = picCierre.ScaleWidth
    lbMovimientos.Left = picCierre.ScaleLeft
    vsLista.Width = picCierre.ScaleWidth
    vsLista.Left = picCierre.ScaleLeft
    vsLista.Height = picCierre.ScaleHeight - vsLista.Top - 30
    
    With vsErrores
        .Width = picErrores.ScaleWidth
        .Left = picErrores.ScaleLeft
        .Height = picErrores.ScaleHeight
        .Top = picErrores.ScaleTop
    End With
    
    Me.Refresh
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    
End Sub

Private Sub Label1_Click(Index As Integer)
    dFecha.SetFocus
End Sub

Private Sub Label4_Click()
    Foco cDisponibilidad
End Sub

Private Sub AccionImprimir()
    
Dim aFormato As String, aEncabezado As String
Dim bSoloErrores As Boolean

    MousePointer = 11
    bSoloErrores = tabCaja.Tabs("errores").Selected
    
    On Error GoTo errPrint
    With vsPrinter
    
    If Not .PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 0: Exit Sub
    
    .Preview = True
    .StartDoc
            
    If .Error Then
        MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
        Screen.MousePointer = vbDefault: Exit Sub
    End If

    EncabezadoListado vsPrinter, "Caja - Consulta de Cierre de Caja", False
    
    .FileName = "Cierre de Caja"
    .FontSize = 8: .FontBold = False
    
    .FontSize = 10: .Text = "Disponibilidad: ": vsPrinter.Text = Trim(cDisponibilidad.Text)
    vsPrinter = ""
    
    .Text = "Fecha de Cierre: ": vsPrinter.Text = Format(dFecha.Value, "Long Date")
    
    If Not bSoloErrores Then
        vsPrinter = "": vsPrinter = ""
        .FontSize = 8: .FontBold = False
        
        .Paragraph = Trim(lbIngresos.Caption)
        vsIngresos.ExtendLastCol = False
        .RenderControl = vsIngresos.hwnd
        vsIngresos.ExtendLastCol = True
        
        If vsLista.Rows > 1 Then
            .Paragraph = ""
            .Paragraph = Trim(lbMovimientos.Caption)
            vsLista.ExtendLastCol = False
            .RenderControl = vsLista.hwnd
            vsLista.ExtendLastCol = True
        End If
        
        .Paragraph = ""
        .FontItalic = True: .Paragraph = "RESUMEN FINAL": .FontItalic = False
        .FontSize = 10
        
        .TableBorder = tbNone
        aTexto = "|Facturación:|" & Trim(lTIngreso.Caption) & "||Saldo Anterior:|" & Trim(lSaldoA.Caption)
        .AddTable "200|<2000|>1300|400|<2000|>1300", "", aTexto, 0, , True
        aTexto = "|Movimientos:|" & Trim(lTMovimiento.Caption) & "||Subtotal:|" & Trim(lSubtotal.Caption)
        .AddTable "200|<2000|>1300|400|<2000|>1300", "", aTexto, 0, , True
        aTexto = "|Subtotal:|" & Trim(lTotal.Caption) & "||Nuevo Saldo:|" & Trim(lSaldoN.Caption)
        .AddTable "200|<2000|>1300|400|<2000|>1300", "", aTexto, 0, , True
        
        'Datos del Arqueo de Caja   -----------------------------------------------------------------------------------------
        .Paragraph = "": .Paragraph = ""
        .FontItalic = True: .Paragraph = "ARQUEO DE CAJA": .FontItalic = False
        .FontSize = 10
        
        Dim mFormat As String
        mFormat = "200|<2000|>1300|>1300"
        
        aTexto = "|" & Trim(lATBilletes.Caption) & "|" & Trim(lABilletes.Caption) & "|"
        .AddTable mFormat, "", aTexto, 0, , True
        aTexto = "|" & Trim(lATChequesDia.Caption) & "|" & Trim(lAChequesDia.Caption) & "|"
        .AddTable mFormat, "", aTexto, 0, , True
        aTexto = "|" & Trim(lATChequesReb.Caption) & "|" & Trim(lAChequesReb.Caption) & "|"
        .AddTable mFormat, "", aTexto, 0, , True
        aTexto = "|" & Trim(lATPendienteCaja.Caption) & "|" & Trim(lAPendienteCaja.Caption) & "|"
        .AddTable mFormat, "", aTexto, 0, , True
        aTexto = "|" & Trim(lATPendienteCam.Caption) & "|" & Trim(lAPendienteCam.Caption) & "|"
        .AddTable mFormat, "", aTexto, 0, , True
        aTexto = "|" & Trim(lATTotalEfectivo.Caption) & "|" & Trim(lATotalEfectivo.Caption) & "|"
        .AddTable mFormat, "", aTexto, 0, , True
        
        aTexto = "|" & Trim(lATSaldoCaja.Caption) & "||" & Trim(lASaldoCaja.Caption)
        .AddTable mFormat, "", aTexto, 0, , True
        
        aTexto = "|" & Trim(lATTotalSaldo.Caption) & "|"
        If lATotalSaldo.Left = lASaldoCaja.Left Then
            aTexto = aTexto & "|" & Trim(lATotalSaldo.Caption)
        Else
            aTexto = aTexto & Trim(lATotalSaldo.Caption) & "|"
        End If
        .AddTable mFormat, "", aTexto, 0, , True
    End If
    
    If vsErrores.Rows > 1 Then
        .Paragraph = "": .Paragraph = ""
        .Paragraph = "Listado de Errores"
        .RenderControl = vsErrores.hwnd
    End If
    
    .EndDoc            'Cierro el Documento--------------------------------------------------------------!!!!!!!!!!!!!!
    .PrintDoc
        
    '.ZOrder 0: .Visible = True
    '.Left = 0: .Width = Me.ScaleWidth
    
    End With

    MousePointer = 0
    Exit Sub
    
errPrint:
    clsGeneral.OcurrioError "Error al realizar la impresión. " & Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub MnuABMCheques_Click()
    EjecutarApp prmPathApp & "SeguimientoCheques.exe"
End Sub

Private Sub MnuAltaCheque_Click()
    EjecutarApp prmPathApp & "AltaCheques.exe"
End Sub

Private Sub MnuBilletes_Click()
    On Error Resume Next
    
    Dim mTexto As String
    If cDisponibilidad.ListIndex <> -1 Then
        mTexto = "D " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) & "|"
    Else
        mTexto = "D 0|"
    End If
    
    mTexto = mTexto & "F " & Format(dFecha.Value, "dd/mm/yyyy")
    
    EjecutarApp prmPathApp & "ConteoBilletes.exe", mTexto
    
End Sub

Private Sub MnuControlCheques_Click()
    EjecutarApp prmPathApp & "Control_Cheques.exe"
End Sub

Private Sub MnuEAddAnuladas_Click()
    If errAddAnuladas(prmMonedaDisp) Then MnuEAddAnuladas.Enabled = False
    vsErrores.SetFocus
End Sub

Private Sub MnuEAddNotas_Click()
    If errAddNotas(prmMonedaDisp) Then MnuEAddNotas.Enabled = False
End Sub

Private Sub MnuEDocX_Click(Index As Integer)
    If Trim(MnuEDocX(Index).Tag) = "" Then Exit Sub
    EjecutarApp Trim(MnuEDocX(Index).Tag)
End Sub

Private Sub MnuEVOpe_Click()
    EjecutarApp prmPathApp & "Visualizacion de Operaciones.exe", CStr(MnuEVOpe.Tag)
End Sub

Private Sub MnuMovimientos_Click()
    
    On Error Resume Next
    Dim aStr As String
    
    If cDisponibilidad.ListIndex <> -1 And vsLista.Cell(flexcpData, vsLista.Row, 0) <> 0 And IsDate(dFecha.Value) Then
        aStr = cDisponibilidad.ItemData(cDisponibilidad.ListIndex) & ":"
        aStr = aStr & vsLista.Cell(flexcpData, vsLista.Row, 0) & ":"
        aStr = aStr & Format(dFecha.Value, "dd/mm/yyyy")
        EjecutarApp prmPathApp & "Movimientos de Caja", aStr
    Else
        EjecutarApp prmPathApp & "Movimientos de Caja"
    End If
    
End Sub

Private Sub MnuPendientes_Click()
    EjecutarApp prmPathApp & "PendientesCaja.exe"
End Sub

Private Sub MnuPMovPendienteCamion_Click()
    EjecutarApp prmPathApp & "\appExploreMsg.exe ", "561:" & dFecha.Value
End Sub

Private Sub mnuPPenCaja_Click()
    EjecutarApp prmPathApp & "\appExploreMsg.exe ", prmPlaPendienteCaja & ":0"
End Sub

Private Sub mnuPPenCajaCam_Click()
    EjecutarApp prmPathApp & "\appExploreMsg.exe ", prmPlaPendienteCajaCam & ":" & dFecha.Value
End Sub

Private Sub MnuPSeñasRecibidas_Click()
    EjecutarApp prmPathApp & "\appExploreMsg.exe ", 562 & ":" & dFecha.Value
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub tabCaja_Click()
    
    Select Case LCase(Trim(tabCaja.SelectedItem.Key))
        Case "cierre": picCierre.ZOrder 0
        Case "arqueo": picArqueo.ZOrder 0
        Case "errores": picErrores.ZOrder 0: vsLista.SetFocus
    End Select
    
End Sub

Private Sub vsErrores_DblClick()
    Call vsErrores_KeyDown(vbKeySpace, 0)
End Sub

Private Sub vsErrores_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsErrores.Rows = 1 Then Exit Sub
    
On Error GoTo errSpace
    If Trim(vsErrores.Cell(flexcpText, vsErrores.Row, 0)) = "" Then Exit Sub
    
    Select Case KeyCode
        Case vbKeySpace      'Corrigo el Total
                Dim mValor As Currency
                With vsErrores
                    If .Cell(flexcpBackColor, .Row, 1) = Colores.Inactivo Then     'lo voy a sacar del subtotal
                        mValor = .Cell(flexcpValue, .Row, 2) * -1
                        .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = .BackColor
                    Else
                        mValor = .Cell(flexcpValue, .Row, 2)
                        .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = Colores.Inactivo
                    End If
                    
                    mValor = .Cell(flexcpValue, .Rows - 3, 2) + mValor
                    .Cell(flexcpText, .Rows - 3, 2) = Format(mValor, "#,##0.00")
                
                    .Cell(flexcpText, .Rows - 1, 1) = "Total"
                    .Cell(flexcpText, .Rows - 1, 2) = Format(mValor + CCur(lATotalSaldo.Caption), "#,##0.00")
                End With
                
        Case 93
            Dim mY As Single, mX As Single
            mY = tabCaja.Top + tabCaja.Tabs(1).Height
            mY = mY + (vsErrores.Row * vsErrores.RowHeight(1)) + vsErrores.RowHeight(0)
            mX = tabCaja.Left + vsErrores.ColWidth(0) + vsErrores.ColWidth(1) + vsErrores.ColWidth(2)
                           
            errMenuPopUp mY, mX
    End Select
    
errSpace:
End Sub

Private Sub vsErrores_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'If vsErrores.Rows = 1 Then Exit Sub
    On Error Resume Next
    
    If Button = vbRightButton Then
        vsErrores.SetFocus
        vsErrores.Select vsErrores.MouseRow, vsErrores.MouseCol
        
         If Trim(vsErrores.Cell(flexcpText, vsErrores.Row, 0)) = "" Then Exit Sub
        errMenuPopUp tabCaja.Top + tabCaja.Tabs(1).Height + Y, X
    End If
    
End Sub

Private Sub errMenuPopUp(mY As Single, mX As Single)
 
    'Limpio Menu    ------------------------------
    Dim mIdX As Integer
    MnuEDocX(0).Caption = "....": MnuEDocX(0).Tag = ""
    For i = 1 To MnuEDocX.UBound
        Unload MnuEDocX(i)
    Next
        
    With vsErrores
        MnuEVOpe.Tag = .Cell(flexcpData, .Row, 1)
        
        If Trim(.Cell(flexcpText, .Row, 5)) <> "" Then
            Dim arrData() As String
            arrData = Split(.Cell(flexcpText, .Row, 5), ";")
        
            If arrData(0) = "G" Then     'ID del Gasto
                MnuEDocX(0).Caption = "Gasto ID: " & arrData(1)
                MnuEDocX(0).Tag = prmPathApp & "Ingreso de Facturas.exe " & arrData(1)
            
            Else    'Documentos asociados (comercio)
                
                For mIdX = LBound(arrData) To UBound(arrData)
                    If mIdX <> 0 Then Load MnuEDocX(mIdX)
                    MnuEDocX(mIdX).Tag = prmPathApp & "Detalle de Factura.exe " & Trim(arrData(mIdX))
                Next
                
                arrData = Split(.Cell(flexcpText, .Row, 3), ";")
                For mIdX = LBound(arrData) To UBound(arrData)
                    MnuEDocX(mIdX).Caption = Trim(arrData(mIdX))
                Next
            End If
        End If
    End With
    
    PopupMenu MnuErrores, , mX, mY
    
End Sub

Private Sub vsLista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton And vsLista.Rows > 1 Then
        If vsLista.Cell(flexcpData, vsLista.MouseRow, 0) <> 0 Then PopupMenu MnuMousse
    End If
    
End Sub

Private Sub vsPrinter_EndDoc()
    EnumeroPiedePagina vsPrinter
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsIngresos
        .Cols = 1: .Rows = 1: .ExtendLastCol = True
        .FormatString = "<Concepto|>Cantidad|>Importe|"
        .WordWrap = False
        .ColWidth(0) = 3600: .ColWidth(1) = 800: .ColWidth(2) = 1800
    End With
    
    With vsLista
        .Cols = 1: .Rows = 1: .ExtendLastCol = True
        .FormatString = "<Concepto|>Cantidad|>Importe|IDCuenta|IDComprobante|>Importe ME|Transferencia|"
        .WordWrap = False
        .ColWidth(0) = 3600: .ColWidth(1) = 800: .ColWidth(2) = 1800: .ColWidth(5) = 1500
        .ColHidden(3) = True: .ColHidden(4) = True: .ColHidden(6) = True
    End With
    
    With vsErrores
        .Cols = 1: .Rows = 1: .ExtendLastCol = True
        .FormatString = ">Hora|<Posible Error|>Importe|<Descripción|Orden|Docs"
        .WordWrap = False
        .ColWidth(0) = 800: .ColWidth(1) = 2000: .ColWidth(2) = 1100
        .ColHidden(4) = True: .ColHidden(5) = True
        .SubtotalPosition = flexSTBelow
    End With
    
    tabCaja.ImageList = img1
    
End Sub

Private Sub CargoDatosArqueo(mMonedaD As Long)

On Error GoTo errArqueo

    Dim mDisponibilidad As Long, mFecha As String

    mDisponibilidad = cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    mFecha = Format(dFecha.Value, "dd/mm/yyyy")
    
    '1) Cargo el conteo de billetes ----------------------------------------------------------
    Dim mATotal As Currency
    mATotal = 0
    Cons = "Select * from ConteoBillete " & _
            " Where CBiDisponibilidad = " & mDisponibilidad & _
            " And CBiFecha = " & Format(mFecha, "'mm/dd/yyyy'")
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        mATotal = mATotal + (RsAux!CBiQ * RsAux!CBiBilleteDe)
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    lABilletes.Caption = Format(mATotal, "#,##0.00")
    
    If mDisponibilidad = prmDispCierreCheques Then
        '2) Cargo Cheques del Día  ----------------------------------------------------------
            'Van No Rebotados o Rebotados con Fecha Mayor a Hoy a última hora
            'Van No Eliminados o Eliminados con Fecha Mayor a Hoy
            'Van Cheques Al día Librados con fecha Menor o igual a hoy
            'Van Cheques No Cobrados o Cobrados Con Fecha Mayor a Hoy a última hora
            
        mATotal = 0
        Cons = "Select * from ChequeDiferido" & _
                    " Where CDiMoneda = " & mMonedaD & _
                    " And CDiVencimiento Is Null " & _
                    " And CDiLibrado <= " & Format(mFecha, "'mm/dd/yyyy'") & _
                    " And ( CDiRebotado is Null " & _
                             " Or CDiRebotado > " & Format(mFecha, "'mm/dd/yyyy 23:59'") & " )" & _
                    " And ( CDiCobrado Is Null " & _
                            " OR CDiCobrado > " & Format(mFecha, "'mm/dd/yyyy 23:59'") & " )" & _
                    " And ( CDiEliminado Is Null Or CDiEliminado > " & Format(mFecha, "'mm/dd/yyyy 23:59'") & " )"
                    
'        cons = "Select * from ChequeDiferido" & _
                    " Where CDiMoneda = " & mMonedaD & _
                    " And ( CDiVencimiento Is Null " & _
                             " And CDiLibrado <= " & Format(mFecha, "'mm/dd/yyyy'") & " )" & _
                    " And CDiRebotado is Null" & _
                    " And ( CDiCobrado Is Null " & _
                            " OR CDiCobrado > " & Format(mFecha, "'mm/dd/yyyy 23:59'") & " )"

        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            mATotal = mATotal + (RsAux!CDiImporte)
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        lAChequesDia.Caption = Format(mATotal, "#,##0.00")
        
        '3) Cargo Cheques Rebotados ----------------------------------------------------------
        mATotal = 0
        Cons = "Select * from ChequeDiferido" & _
                    " Where CDiMoneda = " & mMonedaD & _
                    " And ( ( CDiVencimiento Is Null " & _
                             " And CDiLibrado <= " & Format(mFecha, "'mm/dd/yyyy'") & " ) OR " & _
                             " ( CDiVencimiento <= " & Format(mFecha, "'mm/dd/yyyy'") & " ) )" & _
                    " And ( CDiRebotado Is Not Null " & _
                             " And CDiRebotado < " & Format(mFecha, "'mm/dd/yyyy 23:59'") & " )" & _
                    " And ( CDiEliminado Is Null Or CDiEliminado > " & Format(mFecha, "'mm/dd/yyyy 23:59'") & " )"
    
        'cons = "Select * from ChequeDiferido" & _
                    " Where CDiMoneda = " & mMonedaD & _
                    " And ( ( CDiVencimiento Is Null " & _
                             " And CDiLibrado <= " & Format(mFecha, "'mm/dd/yyyy'") & " ) OR " & _
                             " ( CDiVencimiento <= " & Format(mFecha, "'mm/dd/yyyy'") & " ) )" & _
                    " And CDiRebotado is Not Null"
                    
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            mATotal = mATotal + (RsAux!CDiImporte)
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        lAChequesReb.Caption = Format(mATotal, "#,##0.00")
    
    Else
        lAChequesDia.Caption = "0.00"
        lAChequesReb.Caption = "0.00"
    End If
    
    '4) Cargo Pendientes de Caja ----------------------------------------------------------
    mATotal = 0
    Cons = "Select * from PendientesCaja" & _
                " Where PCaFPendiente <= " & Format(mFecha, "'mm/dd/yyyy'") & _
                " And PCaDisponibilidad = " & mDisponibilidad & _
                " And (PCaFLiquidacion is Null OR PCaFLiquidacion > " & Format(mFecha, "'mm/dd/yyyy 23:59'") & ")"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        mATotal = mATotal + (RsAux!PCaImporte)
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    lAPendienteCaja.Caption = Format(mATotal, "#,##0.00")
    
    '5) Cargo Pendientes de Caja por Camioneros----------------------------------------------------------
    mATotal = 0
    Cons = "Select Sum(DocTotal) as Total" & _
                " From DocumentoPendiente Left Outer Join Liquidacion On DPeIDLiquidacion = LiqID," & _
                " Documento" & _
                " Where DPeDocumento = DocCodigo" & _
                " And DocAnulado = 0" & _
                " And DocFecha <= " & Format(mFecha, "'mm/dd/yyyy 23:59:59'") & _
                " And DPeDisponibilidad = " & mDisponibilidad & _
                " And (DPeIdLiquidacion is Null OR DPeFLiquidacion > " & Format(mFecha, "'mm/dd/yyyy 23:59:59'") & ")"

    '3/10/2011 modifique query en lugar de OR LiqFecha > cambié x DPeFLiquidacion

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Total) Then mATotal = (RsAux!Total)
    End If
    RsAux.Close
    
    lAPendienteCam.Caption = Format(mATotal, "#,##0.00")
    
    'Total de Efectivo  ----
    mATotal = CCur(lABilletes.Caption) + CCur(lAChequesDia.Caption) + CCur(lAChequesReb.Caption) + CCur(lAPendienteCaja.Caption) + CCur(lAPendienteCam.Caption)
    lATotalEfectivo.Caption = Format(mATotal, "#,##0.00")
    
    mATotal = 0
    If IsNumeric(lSaldoN.Caption) Then mATotal = CCur(lSaldoN.Caption)
    lASaldoCaja.Caption = Format(mATotal, "#,##0.00")
    
    mATotal = CCur(lATotalEfectivo.Caption) - CCur(lASaldoCaja.Caption)
    If mATotal < 0 Then
        lATTotalSaldo.Caption = "FALTA"
        lATotalSaldo.Left = lASaldoCaja.Left
        lATotalSaldo.BackColor = lASaldoCaja.BackColor
    Else
        lATTotalSaldo.Caption = "SOBRA"
        lATotalSaldo.Left = lATotalEfectivo.Left
        lATotalSaldo.BackColor = lATotalEfectivo.BackColor
    End If
    
    lATotalSaldo.Caption = Format(mATotal, "#,##0.00")
    lATTotalSaldo.Visible = True: lATotalSaldo.Visible = True
    
    Exit Sub
    
errArqueo:
    clsGeneral.OcurrioError "Error al cagar los datos del Arqueo.", Err.Description
End Sub

Private Sub CargoDatosErrores(mIdsSucursales As String, mIDMoneda As Long)
On Error GoTo errCErrores

    LimpioDatosErrores
    
Dim mIcono As Integer
Dim mFecha As String
Dim mRs As rdoResultset
Dim mTXT1 As String, mTXT2 As String, mTXT3 As String
Dim mValorD As Currency
Dim mData As Long

    mFecha = Format(dFecha.Value, "dd/mm/yyyy")
    mIcono = 0
    '">Hora|<Tipo de Error|>Importe|<Descripción"
    
    '1) Factura Duplicada: 2 doc iguales pal mismo   (+)    -----------------------------------------
    Cons = " Select DocCliente, DocTipo, DocTotal, Count(*) as Q " & _
            " From Documento" & _
            " Where DocFecha Between " & Format(mFecha, "'mm/dd/yyyy'") & _
                                        " AND " & Format(mFecha, "'mm/dd/yyyy 23:59:59'") & _
            " And DocAnulado = 0 " & _
            " And DocMoneda = " & mIDMoneda & _
            " And DocTotal > 0 " & _
            " And DocSucursal IN (" & mIdsSucursales & ")" & _
            " Group by DocCliente, DocTipo, DocTotal " & _
            " Having Count(*) > 1"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        mTXT2 = "": mTXT3 = ""
        Cons = " Select * from Documento" & _
                " Where DocCliente = " & RsAux!DocCliente & _
                " And DocTipo = " & RsAux!DocTipo & _
                " And DocFecha Between " & Format(mFecha, "'mm/dd/yyyy'") & _
                                        " AND " & Format(mFecha, "'mm/dd/yyyy 23:59:59'") & _
                " And DocAnulado = 0 " & _
                " And DocTotal > 0 " & _
                " And DocMoneda = " & mIDMoneda & _
                " And DocSucursal IN (" & mIdsSucursales & ")" & _
                " And DocTotal = " & RsAux!DocTotal
                
        Set mRs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not mRs.EOF
            mTXT1 = Format(mRs!DocFecha, "hh:mm")
            mTXT2 = mTXT2 & Trim(mRs!DocSerie) & "-" & mRs!DocNumero
            mTXT3 = mTXT3 & Trim(mRs!DocCodigo)
            mRs.MoveNext
            If Not mRs.EOF Then mTXT2 = mTXT2 & "; ": mTXT3 = mTXT3 & "; "
        Loop
        mRs.Close
        
        mValorD = RsAux!DocTotal * (RsAux!Q - 1)
        If RsAux!DocTipo = TipoDocumento.NotaCredito Or RsAux!DocTipo = TipoDocumento.NotaDevolucion Or _
                RsAux!DocTipo = TipoDocumento.NotaEspecial Then mValorD = mValorD * -1
        
        With vsErrores
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = mTXT1
            .Cell(flexcpText, .Rows - 1, 1) = RetornoNombreDocumento(RsAux!DocTipo) & " Duplicado/a"
            mData = RsAux!DocCliente: .Cell(flexcpData, .Rows - 1, 1) = mData
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(mValorD, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 3) = mTXT2
            .Cell(flexcpText, .Rows - 1, 5) = mTXT3
            
            If RsAux!DocTipo <> TipoDocumento.ReciboDePago Then
                .Cell(flexcpText, .Rows - 1, 4) = 100
                .Cell(flexcpPicture, .Rows - 1, 0) = img1.ListImages("error100").ExtractIcon
                If mIcono < 100 Then mIcono = 100
            Else
                .Cell(flexcpText, .Rows - 1, 4) = 50
                .Cell(flexcpPicture, .Rows - 1, 0) = img1.ListImages("error50").ExtractIcon
                If mIcono < 50 Then mIcono = 50
            End If
            
        End With
        RsAux.MoveNext
        
        Me.Refresh
    Loop
    RsAux.Close
    '-------------------------------------------------------------------------------------------------------
    
    'Anula Reimpresión: anular 1 que fue reimpreso ( xsucesos anulacion y reimpresion) (-)
    Cons = " Select Suceso.*, DocTipo, DocTotal, DocSerie, DocNumero, DocCliente From Suceso, Documento " & _
            " Where SucFecha Between " & Format(mFecha, "'mm/dd/yyyy'") & _
                                        " And " & Format(mFecha, "'mm/dd/yyyy 23:59:59'") & _
            " And SucTipo = " & TipoSuceso.AnulacionDeDocumentos & _
            " And SucDocumento = DocCodigo " & _
            " And SucDocumento In (  Select SucDocumento From Suceso " & _
                                                " Where SucFecha Between " & Format(mFecha, "'mm/dd/yyyy'") & _
                                                                            " AND " & Format(mFecha, "'mm/dd/yyyy 23:59:59'") & _
                                                " And SucTipo = " & TipoSuceso.Reimpresiones & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        
        mValorD = RsAux!DocTotal * -1
        If RsAux!DocTipo = TipoDocumento.NotaCredito Or RsAux!DocTipo = TipoDocumento.NotaDevolucion Or _
                RsAux!DocTipo = TipoDocumento.NotaEspecial Then mValorD = mValorD * -1
                
        With vsErrores
            .AddItem ""
            .Cell(flexcpPicture, .Rows - 1, 0) = img1.ListImages("error100").ExtractIcon
            .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!SucFecha, "hh:mm")
            .Cell(flexcpText, .Rows - 1, 1) = "Anula Reimpresión"
            mData = RsAux!DocCliente: .Cell(flexcpData, .Rows - 1, 1) = mData
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(mValorD, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero
            .Cell(flexcpText, .Rows - 1, 5) = RsAux!SucDocumento
            .Cell(flexcpText, .Rows - 1, 4) = 100
            If mIcono < 100 Then mIcono = 100
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Me.Refresh
    '-------------------------------------------------------------------------------------------------------
    
'    '3) Compras sin pagos ingresados (-)
'    cons = "Select Compra.* from Compra " & _
                            " Left Outer Join MovimientoDisponibilidad On ComCodigo = MDiIdCompra" & _
                " Where ComFecha = " & Format(mFecha, "'mm/dd/yyyy'") & _
                " And ComMoneda = " & mIDMoneda & _
                " And MDiID is Null" & _
                " And ComTipoDocumento Not In (" & TipoDocumento.CompraNotaCredito & ", " & _
                                                                        TipoDocumento.CompraCredito & ", " & _
                                                                        TipoDocumento.CompraNotaDevolucion & ")"
'    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
'    Do While Not rsAux.EOF
        
'        mValorD = rsAux!ComImporte
'        If Not IsNull(rsAux!ComIVA) Then mValorD = mValorD + rsAux!ComIVA
'        If Not IsNull(rsAux!ComCofis) Then mValorD = mValorD + rsAux!ComCofis
                
'        With vsErrores
'            .AddItem ""
'            .Cell(flexcpPicture, .Rows - 1, 0) = img1.ListImages("error50").ExtractIcon
'            .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!ComFModificacion, "hh:mm")
'            .Cell(flexcpText, .Rows - 1, 1) = "Gastos Sin Pagos"
'            mData = rsAux!ComProveedor: .Cell(flexcpData, .Rows - 1, 1) = mData
            
'            .Cell(flexcpText, .Rows - 1, 2) = Format(mValorD, "#,##0.00")
'            .Cell(flexcpText, .Rows - 1, 3) = "Gasto ID: " & Trim(rsAux!ComCodigo)
'            .Cell(flexcpText, .Rows - 1, 5) = "G;" & Trim(rsAux!ComCodigo)
            
'            .Cell(flexcpText, .Rows - 1, 4) = 50
'            If mIcono < 50 Then mIcono = 50
            
'        End With
'        rsAux.MoveNext
'    Loop
'    rsAux.Close
    
    Me.Refresh
    
    If vsErrores.Rows = 1 Then
        MnuEVOpe.Enabled = False
        MnuEDocX.Item(0).Enabled = False
        MnuTDocs.Enabled = False
        Exit Sub
    Else
        MnuEVOpe.Enabled = True
        MnuEDocX.Item(0).Enabled = True
        MnuTDocs.Enabled = True
    End If
    
    
    tabCaja.Tabs("errores").Image = img1.ListImages("error" & CStr(mIcono)).Index
    
    errTotalizo
    
    Exit Sub
    
errCErrores:
    clsGeneral.OcurrioError "Error al cargar la lista de Errores.", Err.Description
End Sub

Private Sub errTotalizo(Optional bRemove As Boolean = False)

On Error Resume Next
Dim IRow As Integer

    If bRemove Then
        With vsErrores
            IRow = 1
            For i = 1 To .Rows - 1
                If Trim(.Cell(flexcpText, IRow, 5)) = "T" Then .RemoveItem IRow Else IRow = IRow + 1
            Next i
        End With
        Exit Sub
    End If
    
    Dim mValorD As Currency
    With vsErrores
        If .Rows > 1 Then
            .Select 1, 4
            .Sort = flexSortGenericDescending
            
            .AddItem "": .Cell(flexcpText, .Rows - 1, 5) = "T"
            
            mValorD = 0
            For IRow = 1 To .Rows - 1
                If .Cell(flexcpBackColor, IRow, 1) = Colores.Inactivo Then mValorD = mValorD + .Cell(flexcpValue, IRow, 2)
            Next
            .AddItem "": .Cell(flexcpText, .Rows - 1, 5) = "T"
            .Cell(flexcpText, .Rows - 1, 1) = "Sub Total"
            .Cell(flexcpText, .Rows - 1, 2) = Format(mValorD, "#,##0.00")
            
            .AddItem "": .Cell(flexcpText, .Rows - 1, 5) = "T"
            .Cell(flexcpText, .Rows - 1, 1) = lATTotalSaldo.Caption
            .Cell(flexcpText, .Rows - 1, 2) = Format(lATotalSaldo.Caption, "#,##0.00")
            
            .AddItem "": .Cell(flexcpText, .Rows - 1, 5) = "T"
            .Cell(flexcpText, .Rows - 1, 1) = "Total"
            .Cell(flexcpText, .Rows - 1, 2) = Format(mValorD + CCur(lATotalSaldo.Caption), "#,##0.00")
        End If
    End With
    
End Sub

Private Sub LimpioDatosArqueo()

    lABilletes.Caption = ""
    lAChequesDia.Caption = ""
    lAChequesReb.Caption = ""
    lAPendienteCaja.Caption = ""
    lAPendienteCam.Caption = ""
    lATotalEfectivo.Caption = ""
    
    lASaldoCaja.Caption = ""
    
    lATTotalSaldo.Visible = False
    lATotalSaldo.Visible = False
    
End Sub

Private Sub LimpioDatosErrores()
    tabCaja.Tabs("errores").Image = 0
    vsErrores.Rows = 1
    
    MnuEAddAnuladas.Enabled = True
    MnuEAddNotas.Enabled = True
    
End Sub

Private Function errAddAnuladas(mIDMoneda As Long) As Boolean

On Error GoTo errAddA

Dim mValorD As Currency
Dim mFecha As String
Dim mData As Long

    errAddAnuladas = False
    Screen.MousePointer = 11
    mFecha = Format(dFecha.Value, "dd/mm/yyyy")

    Cons = " Select * From Documento" & _
                " Where DocFecha Between " & Format(mFecha, "'mm/dd/yyyy'") & _
                                            " AND " & Format(mFecha, "'mm/dd/yyyy 23:59:59'") & _
                " And DocAnulado = 1 " & _
                " And DocMoneda = " & mIDMoneda & _
                " And DocSucursal IN (" & prmSucursalesDisp & ")"

    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        errTotalizo True
        Do While Not RsAux.EOF
        
            mValorD = RsAux!DocTotal
            If RsAux!DocTipo = TipoDocumento.NotaCredito Or RsAux!DocTipo = TipoDocumento.NotaDevolucion Or _
                    RsAux!DocTipo = TipoDocumento.NotaEspecial Then mValorD = mValorD * -1
            
            With vsErrores
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!DocFecha, "hh:mm")
                .Cell(flexcpText, .Rows - 1, 1) = RetornoNombreDocumento(RsAux!DocTipo) & " Anulado/a"
                mData = RsAux!DocCliente: .Cell(flexcpData, .Rows - 1, 1) = mData
                
                .Cell(flexcpText, .Rows - 1, 2) = Format(mValorD, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero
                .Cell(flexcpText, .Rows - 1, 5) = RsAux!DocCodigo
                
                .Cell(flexcpText, .Rows - 1, 4) = 5
                .Cell(flexcpPicture, .Rows - 1, 0) = img1.ListImages("error0").ExtractIcon
                
            End With
            RsAux.MoveNext
            
        Loop
        errTotalizo
    End If
    RsAux.Close
    '-------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    errAddAnuladas = True
    Exit Function
    
errAddA:
    clsGeneral.OcurrioError "Error al agregar los documentos anulados.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function errAddNotas(mIDMoneda As Long) As Boolean

On Error GoTo errAddN

Dim mValorD As Currency
Dim mFecha As String
Dim mData As Long

    errAddNotas = False
    Screen.MousePointer = 11
    mFecha = Format(dFecha.Value, "dd/mm/yyyy")

    Cons = " Select * From Documento" & _
                " Where DocFecha Between " & Format(mFecha, "'mm/dd/yyyy'") & _
                                            " AND " & Format(mFecha, "'mm/dd/yyyy 23:59:59'") & _
                " And DocAnulado = 0 " & _
                " And DocMoneda = " & mIDMoneda & _
                " And DocSucursal IN (" & prmSucursalesDisp & ")" & _
                " And DocTipo IN (" & TipoDocumento.NotaCredito & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        errTotalizo True
        Do While Not RsAux.EOF
        
            mValorD = RsAux!DocTotal * -1
                        
            With vsErrores
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!DocFecha, "hh:mm")
                .Cell(flexcpText, .Rows - 1, 1) = RetornoNombreDocumento(RsAux!DocTipo)
                mData = RsAux!DocCliente: .Cell(flexcpData, .Rows - 1, 1) = mData
                
                .Cell(flexcpText, .Rows - 1, 2) = Format(mValorD, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero
                .Cell(flexcpText, .Rows - 1, 5) = RsAux!DocCodigo
                
                .Cell(flexcpText, .Rows - 1, 4) = 5
                .Cell(flexcpPicture, .Rows - 1, 0) = img1.ListImages("error0").ExtractIcon
                
            End With
            RsAux.MoveNext
            
        Loop
        errTotalizo
    End If
    RsAux.Close
    '-------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    errAddNotas = True
    Exit Function
    
errAddN:
    clsGeneral.OcurrioError "Error al agregar las devoluciones.", Err.Description
    Screen.MousePointer = 0
End Function
