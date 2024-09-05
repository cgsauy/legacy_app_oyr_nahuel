VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{923DD7D8-A030-4239-BCD4-51FDB459E0FE}#4.0#0"; "orgComboCalculator.ocx"
Begin VB.Form frmPresupuestacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validación de Presupuestos"
   ClientHeight    =   6900
   ClientLeft      =   2145
   ClientTop       =   615
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmValidoPto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHistoria 
      Height          =   735
      Left            =   3780
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   40
      Top             =   3200
      Width           =   1095
      Begin VSFlex6DAOCtl.vsFlexGrid vsHistoria 
         Height          =   795
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
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
   Begin VB.PictureBox PicPresupuesto 
      Height          =   2655
      Left            =   840
      ScaleHeight     =   2595
      ScaleWidth      =   8355
      TabIndex        =   41
      Top             =   4560
      Width           =   8415
      Begin orgCalculatorFlat.orgCalculator caCostoFinal 
         Height          =   315
         Left            =   5640
         TabIndex        =   63
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorCalculator=   12564658
         Text            =   "0.00"
      End
      Begin VB.TextBox tCosto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5640
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox tPresupuesto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   60
         MaxLength       =   30
         TabIndex        =   5
         Top             =   300
         Width           =   3135
      End
      Begin VB.TextBox tCantidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3180
         MaxLength       =   15
         TabIndex        =   6
         Top             =   300
         Width           =   435
      End
      Begin VB.TextBox tComentarioP 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   960
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1920
         Width           =   7335
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   4980
         TabIndex        =   9
         Top             =   1200
         Width           =   615
         _ExtentX        =   1085
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsMotivo 
         Height          =   1215
         Left            =   60
         TabIndex        =   7
         Top             =   660
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   2143
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
      Begin AACombo99.AACombo cMonedaFinal 
         Height          =   315
         Left            =   4980
         TabIndex        =   12
         Top             =   1560
         Width           =   615
         _ExtentX        =   1085
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
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Reclamo:"
         Height          =   195
         Left            =   6360
         TabIndex        =   65
         Top             =   900
         Width           =   795
      End
      Begin VB.Label lReclamo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-May 2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7140
         TabIndex        =   64
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label labCostoGrilla 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "125.00"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6780
         TabIndex        =   56
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label labFReparado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1/1/2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7140
         TabIndex        =   49
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label11 
         Caption         =   "Reparado:"
         Height          =   195
         Left            =   6360
         TabIndex        =   48
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Técnico:"
         Height          =   195
         Left            =   3780
         TabIndex        =   47
         Top             =   600
         Width           =   855
      End
      Begin VB.Label labTecnico 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coquito"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4440
         TabIndex        =   46
         Top             =   600
         Width           =   1755
      End
      Begin VB.Label labFPresupuesto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-May 2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4980
         TabIndex        =   45
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label labIDLocalReparacion 
         Caption         =   "Presupuestado:"
         Height          =   195
         Left            =   3780
         TabIndex        =   44
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "C&osto Final:"
         Height          =   255
         Left            =   3780
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label LabFAceptado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-May 2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7140
         TabIndex        =   43
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label labTitAceptado 
         Caption         =   "Aceptado:"
         Height          =   195
         Left            =   6360
         TabIndex        =   42
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "&Mano Obra:"
         Height          =   255
         Left            =   3780
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label labMotivo 
         Caption         =   "&Artículo: [F12 a Presupuesto]"
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   3195
      End
      Begin VB.Label Label17 
         Caption         =   "Co&mentario:"
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   1920
         Width           =   855
      End
   End
   Begin VB.PictureBox PicProductos 
      Height          =   1455
      Left            =   2160
      ScaleHeight     =   1395
      ScaleWidth      =   2835
      TabIndex        =   50
      Top             =   3540
      Width           =   2895
      Begin VSFlex6DAOCtl.vsFlexGrid vsProducto 
         Height          =   1035
         Left            =   60
         TabIndex        =   51
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
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
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsHistoria2 
         Height          =   795
         Left            =   60
         TabIndex        =   52
         Top             =   1140
         Width           =   1695
         _ExtentX        =   2990
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   6645
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin VB.Frame fFicha 
      Caption         =   "Ficha"
      ForeColor       =   &H00000080&
      Height          =   3075
      Left            =   60
      TabIndex        =   21
      Top             =   420
      Width           =   8475
      Begin VB.TextBox tMemoInt 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   525
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   66
         Text            =   "frmValidoPto.frx":0312
         Top             =   2400
         Width           =   7395
      End
      Begin VB.TextBox tComentarioS 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         Height          =   465
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   1920
         Width           =   8235
      End
      Begin VB.TextBox tServicio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label9 
         Caption         =   "Interno:"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Estado:"
         Height          =   195
         Left            =   5940
         TabIndex        =   55
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label labFactura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   39
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label LabEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6660
         TabIndex        =   36
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label LabGarantia 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12 Meses"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6660
         TabIndex        =   35
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Garantía:"
         Height          =   195
         Left            =   5940
         TabIndex        =   34
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label LabFCompra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-May 2000"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "F/Compra:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Factura:"
         Height          =   255
         Left            =   2340
         TabIndex        =   15
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label labIngreso 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-May 2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5040
         TabIndex        =   31
         Top             =   240
         Width           =   3295
      End
      Begin VB.Label Label5 
         Caption         =   "Ingreso:"
         Height          =   195
         Left            =   4320
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.Label labTelefono 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   960
         TabIndex        =   29
         Top             =   900
         Width           =   7395
      End
      Begin VB.Label Label4 
         Caption         =   "Teléfonos:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   900
         Width           =   735
      End
      Begin VB.Label labEstadoServicio 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TALLER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   1800
         TabIndex        =   27
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label3 
         Caption         =   "&Servicio:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label LabCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WALTER ADRIAN OCCHIUZZI MARTINEZ"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   600
         Width           =   7395
      End
      Begin VB.Label LabProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(190111) REFRIGERADOR PANAVOX ALTO DE 10 PULGAS"
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
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   1260
         Width           =   4875
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   375
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   8355
      TabIndex        =   20
      Top             =   70
      Width           =   8415
      Begin VB.CommandButton bHelp 
         Height          =   310
         Left            =   7980
         Picture         =   "frmValidoPto.frx":0328
         Style           =   1  'Graphical
         TabIndex        =   62
         TabStop         =   0   'False
         ToolTipText     =   "Ayuda [Ctrl+H]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bNextServicio 
         Height          =   310
         Left            =   1920
         Picture         =   "frmValidoPto.frx":0632
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "Siguiente Servicio"
         Top             =   0
         Width           =   310
      End
      Begin VB.CheckBox chRepuesto 
         Height          =   310
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Contenga algún Repuestos"
         Top             =   0
         Width           =   310
      End
      Begin VB.CheckBox chComentario 
         Height          =   310
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Sin Comentario"
         Top             =   0
         Width           =   310
      End
      Begin VB.CheckBox chArticulo 
         Height          =   310
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Tipo de Artículo"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bAyuda 
         Height          =   310
         Left            =   5940
         Picture         =   "frmValidoPto.frx":0734
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Ayuda de Presupuesto [Ctrl+Y]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bVisualizacion 
         Height          =   310
         Left            =   3420
         Picture         =   "frmValidoPto.frx":0FFE
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "Visualización de Operaciones"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bProducto 
         Height          =   310
         Left            =   2820
         Picture         =   "frmValidoPto.frx":18C8
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Ficha de Producto. [Ctrl+P]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bHistoria 
         Height          =   310
         Left            =   2460
         Picture         =   "frmValidoPto.frx":2192
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Historia. [Ctrl+H]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   1500
         Picture         =   "frmValidoPto.frx":2A5C
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Cancelar. [Ctrl+C]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bModificar 
         Height          =   310
         Left            =   780
         Picture         =   "frmValidoPto.frx":2B5E
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Modificar. [Ctrl+M]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bSalir 
         Height          =   310
         Left            =   60
         Picture         =   "frmValidoPto.frx":2CA8
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Salir. [Ctrl+X]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   1140
         Picture         =   "frmValidoPto.frx":2DAA
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Grabar [Ctrl+G]."
         Top             =   0
         Width           =   310
      End
   End
   Begin ComctlLib.TabStrip TabValidacion 
      Height          =   2955
      Left            =   60
      TabIndex        =   2
      Top             =   3600
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   5212
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Historia"
            Key             =   "historia"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "P&resupuesto"
            Key             =   "presupuesto"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pro&ductos"
            Key             =   "productos"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8640
      Y1              =   0
      Y2              =   0
   End
   Begin ComctlLib.ImageList Image1 
      Left            =   7560
      Top             =   5580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":2EAC
            Key             =   "historia"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":31C6
            Key             =   "Valido"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":34E0
            Key             =   "servicio"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":37FA
            Key             =   "producto"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":3B14
            Key             =   "visualizacion"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":3E2E
            Key             =   "helptipo"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":4148
            Key             =   "helpcomentario"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":4462
            Key             =   "helprepuesto"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":477C
            Key             =   "repuesto"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":4A96
            Key             =   "sincomentario"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":4DB0
            Key             =   "ayuda"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":50CA
            Key             =   "sinrepuesto"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":53E4
            Key             =   "siguienteservicio"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmValidoPto.frx":56FE
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuOpModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuOpL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuOpCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuOpL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpNextServicio 
         Caption         =   "&Siguiente Servicio"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu MnuAccesos 
      Caption         =   "&Accesos"
      Begin VB.Menu MnuAccSeguimiento 
         Caption         =   "Seguimiento de Servicios"
      End
      Begin VB.Menu MnuHistoria 
         Caption         =   "Historia de Servicios"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuProducto 
         Caption         =   "Mantenimiento de Productos"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuCpasArt 
         Caption         =   "Compras de ese articulo"
      End
      Begin VB.Menu MnuAcL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVisualizacion 
         Caption         =   "&Visualización de Operaciones"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu MnuAyuTipoArticulo 
         Caption         =   "Tipo de Artículo"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuAyuArticulo 
         Caption         =   "Articulo"
      End
      Begin VB.Menu MnuAyuTipoComentario 
         Caption         =   "Con Comentario"
      End
      Begin VB.Menu MnuAyuNoRepuesto 
         Caption         =   "No Conciderar Repuestos"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuAyuAlgunRepuesto 
         Caption         =   "Algún repuesto"
      End
      Begin VB.Menu MnuAyuTipoRepuesto 
         Caption         =   "Todos los Repuestos"
      End
      Begin VB.Menu MnuAyuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAyuEjecutar 
         Caption         =   "Consultar Ayuda"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSaSalir 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuAyudaHelp 
      Caption         =   "?"
      Begin VB.Menu MnuHelp 
         Caption         =   "Ayuda ..."
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmPresupuestacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------
'28/8/2001  Cambios: agregue botón ayuda. Seteo botón comentario y aviso de que afectará el stock.
'4-11-2002  Modifique dll x orCGSA, además cambie las consultas de listas (" " x "%")
'12-3-2008  presento la historia cdo dan sgte.
'           presento cuotas vdas mayores a 30 días.
'           se agregó local de ingreso.
'           si el costo = 0 --> lo doy como aceptado
'           si aumenta el costo --> dejo sin aceptar el pto.
'------------------------------------------------------------------------------------------
Private bAvisarStock As Boolean
Private sModEnt As Boolean
Private aTexto As String
Private gSeleccionado As Long

Public Property Get prmCodigo() As Long
    prmCodigo = gSeleccionado
End Property
Public Property Let prmCodigo(Codigo As Long)
    gSeleccionado = Codigo
End Property

Private Sub bAyuda_Click()
    If Val(tServicio.Tag) > 0 Then ConsultaAyuda
End Sub

Private Sub bCancelar_Click()
    CargoServicio
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bHelp_Click()
    AccionAyuda
End Sub

Private Sub bHistoria_Click()
    If Val(LabProducto.Tag) > 0 Then EjecutarApp pathApp & "\Historia Servicio.exe", CStr(LabProducto.Tag)
End Sub

Private Sub bModificar_Click()
    AccionModificar
End Sub

Private Sub bNextServicio_Click()
    AccionSiguienteServicio
End Sub

Private Sub bProducto_Click()
    If Val(LabCliente.Tag) > 0 Then
        EjecutarApp pathApp & "\Productos.exe", LabCliente.Tag & ";p" & Val(LabProducto.Tag)
    End If
End Sub

Private Sub bSalir_Click()
    Unload Me
End Sub

Private Sub bVisualizacion_Click()
    If Val(LabCliente.Tag) <> 0 Then EjecutarApp App.Path & "\Visualizacion de operaciones", CStr(LabCliente.Tag)
End Sub

Private Sub caCostoFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub cMoneda_GotFocus()
    With cMoneda
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tCosto
End Sub

Private Sub cMoneda_LostFocus()
    cMoneda.SelStart = 0
End Sub

Private Sub cMonedaFinal_GotFocus()
    With cMonedaFinal
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cMonedaFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco caCostoFinal
End Sub
Private Sub cMonedaFinal_LostFocus()
    cMonedaFinal.SelStart = 0
End Sub

Private Sub chArticulo_Click()
    If chArticulo.Value = 0 Then
        chArticulo.ToolTipText = "Tipo de Artículo"
        MnuAyuArticulo.Checked = False
        MnuAyuTipoArticulo.Checked = True
    Else
        chArticulo.ToolTipText = "Artículo"
        MnuAyuArticulo.Checked = True
        MnuAyuTipoArticulo.Checked = False
    End If
End Sub

Private Sub chComentario_Click()
    If chComentario.Value = 0 Then
        chComentario.ToolTipText = "Sin Comentario"
        MnuAyuTipoComentario.Checked = False
    Else
        chComentario.ToolTipText = "Con Comentario"
        MnuAyuTipoComentario.Checked = True
    End If
End Sub

Private Sub chRepuesto_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        chRepuesto.Value = 1    'Levanto el botón.
        chRepuesto.Value = vbGrayed
        chRepuesto.ToolTipText = "No conciderar repuestos."
        chRepuesto.Picture = Image1.ListImages("sinrepuesto").ExtractIcon
        MnuAyuNoRepuesto.Checked = True
        MnuAyuAlgunRepuesto.Checked = False
        MnuAyuTipoRepuesto.Checked = False
    Else
        If chRepuesto.Value = 1 Then
            chRepuesto.ToolTipText = "Contenga todos los Repuestos"
            MnuAyuNoRepuesto.Checked = False
            MnuAyuAlgunRepuesto.Checked = False
            MnuAyuTipoRepuesto.Checked = True
        ElseIf chRepuesto.Value = 0 Then
            chRepuesto.ToolTipText = "Contenga algún Repuestos"
            MnuAyuNoRepuesto.Checked = False
            MnuAyuAlgunRepuesto.Checked = True
            MnuAyuTipoRepuesto.Checked = False
            chRepuesto.Picture = Image1.ListImages("repuesto").ExtractIcon
        End If
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
    If sModEnt Then
        sModEnt = False
        If caCostoFinal.Enabled Then Foco caCostoFinal Else Foco tPresupuesto
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    If App.PrevInstance Then
        MsgBox "Hay una instancia abierta de Validación de Presupuestos.", vbExclamation, "ATENCIÓN"
        End
        Exit Sub
    End If
    FechaDelServidor
    bAvisarStock = False
    ObtengoSeteoForm Me
    If Me.Height <> 7560 Then Me.Height = 7560
    If Me.Width <> 8730 Then Me.Width = 8730
    picBotones.BorderStyle = vbBSNone
    PicPresupuesto.BorderStyle = 0: picHistoria.BorderStyle = 0: PicProductos.BorderStyle = 0
    
    PicPresupuesto.Height = TabValidacion.ClientHeight
    picHistoria.Height = PicPresupuesto.Height
    
    picHistoria.Width = TabValidacion.Width - 100
    PicPresupuesto.Width = picHistoria.Width
    
    picHistoria.Top = TabValidacion.ClientTop: picHistoria.Left = TabValidacion.ClientLeft
    PicPresupuesto.Top = picHistoria.Top: PicPresupuesto.Left = picHistoria.Left
    vsHistoria.Width = picHistoria.Width - (vsHistoria.Left * 2)
    vsHistoria.Height = picHistoria.Height - (vsHistoria.Top * 2)
    picHistoria.ZOrder 0
    
    PicProductos.Top = PicPresupuesto.Top: PicProductos.Left = PicPresupuesto.Left
    PicProductos.Height = PicPresupuesto.Height: PicProductos.Width = PicPresupuesto.Width
    vsHistoria2.Left = 30
    vsHistoria2.Width = PicProductos.Width - (vsHistoria.Left * 2)
    vsHistoria2.Height = picHistoria.Height - (vsProducto.Top + vsProducto.Height + vsHistoria.Top - (vsProducto.Top + vsProducto.Height))
    vsProducto.Width = vsHistoria2.Width
    
    bNextServicio.Picture = Image1.ListImages("siguienteservicio").ExtractIcon
    bHistoria.Picture = Image1.ListImages("historia").ExtractIcon
    bProducto.Picture = Image1.ListImages("producto").ExtractIcon
    bVisualizacion.Picture = Image1.ListImages("visualizacion").ExtractIcon
    
    bAyuda.Picture = Image1.ListImages("ayuda").ExtractIcon
    
    chArticulo.Picture = Image1.ListImages("helptipo").ExtractIcon
    chArticulo.DownPicture = Image1.ListImages("producto").ExtractIcon
    chComentario.DownPicture = Image1.ListImages("helpcomentario").ExtractIcon
    chComentario.Picture = Image1.ListImages("sincomentario").ExtractIcon
    chRepuesto.Picture = Image1.ListImages("sinrepuesto").ExtractIcon
    chRepuesto.DownPicture = Image1.ListImages("helprepuesto").ExtractIcon
    bHelp.Picture = Image1.ListImages("help").ExtractIcon
    
    chRepuesto.Value = vbGrayed
    chRepuesto.ToolTipText = "No conciderar repuestos."
    
    With TabValidacion
        Set .ImageList = Image1
        .Tabs(1).Image = Image1.ListImages("historia").Index
        .Tabs(2).Image = Image1.ListImages("Valido").Index
        .Tabs(3).Image = Image1.ListImages("producto").Index
    End With
    
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda
    Cons = "Select MonCodigo, MonSigno From Moneda Where MonCodigo = " & paMonedaPesos & " Order by MonSigno"
    CargoCombo Cons, cMonedaFinal
    
    labMotivo.Caption = "Re&puesto: [F12 a Comb.Repuesto]": labMotivo.Tag = 0
    
    chComentario.Value = Val(ObtengoSeteoControl("chComentarioValue", 0))
    
    If gSeleccionado <> 0 Then
        tServicio.Text = gSeleccionado
        tServicio.Tag = gSeleccionado
        'CargoServicio
        AccionModificar
        sModEnt = True
    Else
        LimpioCamposFicha
        LimpioCamposPresupuesto
        OcultoCamposPresupuesto
        MeBotones False, False, False
    End If
    FechaDelServidor
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    GuardoSeteoControl "chComentarioValue", chComentario.Value
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Sub

Private Sub InicializoGrilla()
    On Error Resume Next
    With vsMotivo
        .Editable = False
        .ExtendLastCol = True
        .Redraw = False
        .WordWrap = False
        .Rows = 1: .Cols = 1
        .FormatString = "Q|Repuestos Utilizados|>Costo"
        .ColWidth(0) = 250: .ColWidth(1) = 2400
        .Redraw = True
    End With
    With vsHistoria
        .Tag = ""
        .Rows = 1
        .WordWrap = True
        .FormatString = "Fecha|Motivos|Estado|>Importe|"
        .ColWidth(0) = 750: .ColWidth(1) = 5200: .ColWidth(3) = 1300: .ColWidth(5) = 10
        .ColAlignment(0) = flexAlignLeftTop: .ColAlignment(1) = flexAlignLeftTop: .ColAlignment(2) = flexAlignLeftTop: .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(4) = flexAlignRightTop: .ColAlignment(5) = flexAlignLeftTop
    End With
    With vsHistoria2
        .Rows = 1
        .WordWrap = True
        .FormatString = "Fecha|Motivos|Estado|>Importe|"
        .ColWidth(0) = 750: .ColWidth(1) = 5200: .ColWidth(3) = 1300: .ColWidth(5) = 10
        .ColAlignment(0) = flexAlignLeftTop: .ColAlignment(1) = flexAlignLeftTop: .ColAlignment(2) = flexAlignLeftTop: .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(4) = flexAlignRightTop: .ColAlignment(5) = flexAlignLeftTop
    End With
    With vsProducto
        .Rows = 1
        .Cols = 1
        .ExtendLastCol = True
        .FormatString = "Artículo|Estado|>F.Compra|Garantía|N° Serie|Factura"
        .ColWidth(0) = 4200: .ColWidth(2) = 1000: .ColWidth(4) = 800
        .ColHidden(5) = True
    End With

End Sub

Private Function BuscoSignoMoneda(IdMoneda As Long)
Dim RsMon As rdoResultset
    BuscoSignoMoneda = ""
    Cons = "Select * From Moneda Where MonCodigo = " & IdMoneda
    Set RsMon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsMon.EOF Then BuscoSignoMoneda = Trim(RsMon!MonSigno)
    RsMon.Close
End Function

Private Sub Label15_Click()
    Foco cMoneda
End Sub
Private Sub Label17_Click()
    Foco tComentarioP
End Sub
Private Sub Label3_Click()
    Foco tServicio
End Sub

Private Sub Label8_Click()
    Foco cMonedaFinal
End Sub

Private Sub labMotivo_Click()
    Foco tPresupuesto
End Sub

Private Sub MnuAccSeguimiento_Click()
    If Val(tServicio.Tag) = 0 Then Exit Sub
    EjecutarApp App.Path & "\Seguimiento de Servicios.exe", Val(tServicio.Tag)
End Sub

Private Sub MnuAyuAlgunRepuesto_Click()
    chRepuesto.Value = 0
    chRepuesto.ToolTipText = "Contenga algún Repuestos"
    chRepuesto.Picture = Image1.ListImages("repuesto").ExtractIcon
    MnuAyuNoRepuesto.Checked = False
    MnuAyuAlgunRepuesto.Checked = True
    MnuAyuTipoRepuesto.Checked = False
End Sub

Private Sub MnuAyuArticulo_Click()
    If chArticulo.Value = 0 Then
        chArticulo.Value = 1
    Else
        chArticulo.Value = 0
    End If
End Sub

Private Sub MnuAyuEjecutar_Click()
    If Val(tServicio.Tag) > 0 Then ConsultaAyuda
End Sub

Private Sub MnuAyuNoRepuesto_Click()
    MnuAyuNoRepuesto.Checked = True
    MnuAyuAlgunRepuesto.Checked = False
    MnuAyuTipoRepuesto.Checked = False
    
    chRepuesto.Value = 0    'Levanto el botón
    chRepuesto.Value = vbGrayed
    chRepuesto.ToolTipText = "No conciderar repuestos."
    chRepuesto.Picture = Image1.ListImages("sinrepuesto").ExtractIcon
End Sub

Private Sub MnuAyuTipoArticulo_Click()
    If chArticulo.Value = 0 Then
        chArticulo.Value = 1
    Else
        chArticulo.Value = 0
    End If
End Sub

Private Sub MnuAyuTipoComentario_Click()
    If chComentario.Value = 0 Then
        chComentario.Value = 1
    Else
        chComentario.Value = 0
    End If
End Sub

Private Sub MnuAyuTipoRepuesto_Click()

    MnuAyuNoRepuesto.Checked = False
    MnuAyuAlgunRepuesto.Checked = False
    MnuAyuTipoRepuesto.Checked = True
    chRepuesto.Value = 1
    chRepuesto.ToolTipText = "Contenga todos los repuestos Repuestos"
    chRepuesto.Picture = Image1.ListImages("repuesto").ExtractIcon
    
End Sub

Private Sub MnuCpasArt_Click()
'Plantilla.
If Val(MnuCpasArt.Tag) > 0 Then EjecutarApp App.Path & "\appExploreMsg.exe ", prmPlaCpas & ":" & LabCliente.Tag & ";;" & Val(MnuCpasArt.Tag)
End Sub

Private Sub MnuHelp_Click()
    AccionAyuda
End Sub

Private Sub MnuHistoria_Click()
    Call bHistoria_Click
End Sub
Private Sub MnuOpCancelar_Click()
    CargoServicio
End Sub

Private Sub MnuOpGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuOpModificar_Click()
    AccionModificar
End Sub

Private Sub MnuOpNextServicio_Click()
    AccionSiguienteServicio
End Sub

Private Sub MnuProducto_Click()
    Call bProducto_Click
End Sub

Private Sub MnuSaSalir_Click()
    Unload Me
End Sub

Private Sub MnuVisualizacion_Click()
    Call bVisualizacion_Click
End Sub

Private Sub TabValidacion_Click()
    If TabValidacion.SelectedItem.Index = 1 Then
        picHistoria.ZOrder 0
    Else
        If TabValidacion.SelectedItem.Index = 2 Then PicPresupuesto.ZOrder 0 Else CargoOtrosProducto: PicProductos.ZOrder 0
    End If
End Sub

Private Sub tCantidad_GotFocus()
    With tCantidad
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tCantidad.Text = "" And tPresupuesto.Text = "" Then Foco tCosto: Exit Sub
        If tCantidad.Text = "" Then Exit Sub
        If Not IsNumeric(tCantidad.Text) Then MsgBox "El formato debe ser numérico.", vbInformation, "ATENCIÓN": Exit Sub
        If Val(tCantidad.Text) <= 0 Then MsgBox "Debe ingresar un valor positivo.", vbInformation, "ATENCIÓN": Exit Sub
        AgregoMotivo CLng(tPresupuesto.Tag)
        tPresupuesto.Text = "": tCantidad.Text = ""
    End If
End Sub
Private Sub tComentarioP_GotFocus()
    With tComentarioP
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tComentarioP_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tCosto_GotFocus()
    With tCosto
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tCosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If cMonedaFinal.Enabled Then Foco cMonedaFinal Else Foco tComentarioP
End Sub
Private Sub tCosto_LostFocus()
    If IsNumeric(tCosto.Text) Then tCosto.Text = Format(tCosto.Text, FormatoMonedaP) Else tCosto.Text = ""
End Sub

Private Sub tPresupuesto_Change()
    tPresupuesto.Tag = ""
End Sub

Private Sub tPresupuesto_GotFocus()
    With tPresupuesto
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tPresupuesto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF12
            If labMotivo.Tag = 0 Then
                labMotivo.Caption = "Comb. Re&puesto: [F12 a Repuesto]": labMotivo.Tag = 1
                tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
            Else
                labMotivo.Caption = "Re&puesto: [F12 a Comb.Repuesto]": labMotivo.Tag = 0
                tCantidad.Enabled = True: tCantidad.BackColor = Blanco
            End If
        Case vbKeyReturn
            If tPresupuesto.Text <> "" Then
                If IsNumeric(tPresupuesto.Text) Then
                    If labMotivo.Tag = 1 Then   'Presupuesto
                        BuscoPresupuestoXCodigo tPresupuesto.Text
                    Else
                        BuscoArticuloXCodigo tPresupuesto.Text
                    End If
                Else
                    If labMotivo.Tag = 1 Then   'Presupuesto
                        BuscoPresupuestoXNombre
                    Else
                        BuscoArticuloXNombre
                    End If
                End If
            Else
                tCantidad.Text = ""
                Foco cMoneda
            End If
    End Select
End Sub

Private Sub tServicio_Change()
    LimpioCamposFicha
    LimpioCamposPresupuesto
    OcultoCamposPresupuesto
    tServicio.Enabled = True
    MeBotones False, False, False
End Sub

Private Sub tServicio_GotFocus()
    With tServicio
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tServicio_KeyPress(KeyAscii As Integer)
On Error GoTo ErrCS

    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tServicio.Text) Then
            Screen.MousePointer = 11
            CargoServicio
            If TabValidacion.SelectedItem.Index = 3 Then CargoOtrosProducto
            Screen.MousePointer = 0
        Else
            If Trim(tServicio.Text) <> "" Then MsgBox "Formato incorrecto.", vbExclamation, "ATENCIÓN"
        End If
    End If
    Exit Sub
ErrCS:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los datos del servicio.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub CargoServicio()
Dim LocReparacion As Integer
    
    LimpioCamposFicha
    LimpioCamposPresupuesto
    OcultoCamposPresupuesto
    tServicio.Enabled = True
    MnuCpasArt.Tag = ""
    
    MeBotones False, False, False
    Cons = "Select Servicio.*, ArtCodigo, ArtNombre, ArtID, Producto.*, SucAbreviacion From (Servicio LEFT OUTER JOIN Sucursal ON SerLocalIngreso = SucCodigo), Producto, Articulo " _
        & " Where SerCodigo = " & Val(tServicio.Text) _
        & " And SerProducto = ProCodigo And ProArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No existe un servicio con ese código.", vbInformation, "ATENCIÓN"
    Else
        If IsNull(RsAux!SerLocalReparacion) Then
            RsAux.Close: Screen.MousePointer = 0
            MsgBox "Hay inconsistencias en los datos, no esta ingresado el Local de Reparación.", vbInformation, "ATENCIÓN"
            Exit Sub
        End If
        MnuCpasArt.Tag = RsAux("ArtCodigo")
        LocReparacion = RsAux!SerLocalReparacion
        CargoDatosServicio
        RsAux.Close
        CargoHistoria LabProducto.Tag, vsHistoria
        
        'VEO si le dejo entrar
        If Val(labEstadoServicio.Tag) = EstadoS.Taller Then
            CargoDatosTaller tServicio.Tag, LocReparacion
        Else
            'Veo si tiene datos en la ficha de taller
            Cons = "Select * From Taller Where TalServicio = " & tServicio.Tag
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
            If Not RsAux.EOF Then
                PresentoDatosTaller
            Else
                cMonedaFinal.Text = "": caCostoFinal.Clean: caCostoFinal.Tag = "0"
                OcultoCamposPresupuesto
            End If
            RsAux.Close
        End If
    End If
    If vsHistoria.Rows = 1 Then TabValidacion.Tabs(2).Selected = True
    
End Sub
Private Sub CargoDatosServicio()
    
    tServicio.Tag = RsAux!SerCodigo
    
    If Not IsNull(RsAux!SerLocalReparacion) Then
        labIDLocalReparacion.Tag = RsAux!SerLocalReparacion
    Else
        labIDLocalReparacion.Tag = "0"
    End If
    If Not IsNull(RsAux!SerReclamoDe) Then lReclamo.Caption = RsAux!SerReclamoDe Else lReclamo.Caption = ""
    labEstadoServicio.Caption = UCase(EstadoServicio(RsAux!SerEstadoServicio))
    labEstadoServicio.Tag = RsAux!SerEstadoServicio
    
    labIngreso.Caption = " " & Format(RsAux!SerFecha, FormatoFP) & " en " & Trim(RsAux("SucAbreviacion"))
    labIngreso.Tag = RsAux!SerModificacion
    
    If Not IsNull(RsAux!SerCostoFinal) And Not IsNull(RsAux!SerMoneda) Then caCostoFinal.Text = RsAux!SerCostoFinal: caCostoFinal.Tag = "1": BuscoCodigoEnCombo cMonedaFinal, RsAux!SerMoneda
    
    CargoDatosCliente RsAux!ProCliente
    
    LabProducto.Caption = " (" & Format(RsAux!ProCodigo, "#,000") & ") " & Trim(RsAux!ArtNombre)
    LabProducto.Tag = RsAux!ProCodigo
    LabFCompra.Tag = RsAux!ProFModificacion
    If Not IsNull(RsAux!ProFacturaS) Then labFactura.Caption = Trim(RsAux!ProFacturaS)
    If Not IsNull(RsAux!ProFacturaN) Then labFactura.Caption = Trim(labFactura.Caption) & " " & Trim(RsAux!ProFacturaN)
    If Not IsNull(RsAux!ProCompra) Then LabFCompra.Caption = Format(RsAux!ProCompra, "dd/mm/yyyy")
    
    LabEstado.Caption = " " & EstadoProducto(RsAux!SerEstadoProducto, False)
    LabEstado.Tag = RsAux!SerEstadoProducto
    LabGarantia.Caption = " " & RetornoGarantia(RsAux!ArtID)
    
    If Not IsNull(RsAux!SerComentario) Then tComentarioS.Text = Trim(RsAux!SerComentario)
    If Not IsNull(RsAux("SerComInterno")) Then tMemoInt.Text = Trim(RsAux("SerComInterno"))
    Dim RsSR As rdoResultset
    Cons = "Select * From ServicioRenglon, MotivoServicio " _
        & " Where SReServicio = " & RsAux!SerCodigo _
        & " And SReTipoRenglon = " & TipoRenglonS.Llamado _
        & " And SReMotivo = MSeID"
    Set RsSR = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsSR.EOF Then
        If Trim(tComentarioS.Text) <> "" Then tComentarioS.Text = Trim(tComentarioS.Text) & Chr(13) & Chr(10)
        tComentarioS.Text = Trim(tComentarioS.Text) & Trim(RsSR!MSeNombre): RsSR.MoveNext
    End If
    Do While Not RsSR.EOF
        tComentarioS.Text = Trim(tComentarioS.Text) & ", " & Trim(RsSR!MSeNombre)
        RsSR.MoveNext
    Loop
    RsSR.Close
    
End Sub

Private Sub LimpioCamposFicha()
    tServicio.Tag = ""
    labEstadoServicio.Caption = "": labEstadoServicio.Tag = ""
    labIngreso.Caption = "": labIngreso.Tag = ""
    LabCliente.Caption = ""
    labTelefono.Caption = ""
    LabProducto.Caption = "": LabProducto.Tag = ""
    labFactura.Caption = ""
    LabFCompra.Caption = "": LabFCompra.Tag = ""
    LabEstado.Caption = "": LabEstado.Tag = ""
    LabGarantia.Caption = ""
    tComentarioS.Text = ""
    tMemoInt.Text = ""
    MnuCpasArt.Tag = ""
End Sub
Private Sub LimpioCamposPresupuesto()
    
    InicializoGrilla
    LabFAceptado.BackColor = Inactivo: LabFAceptado.ForeColor = vbBlack
    LabFAceptado.Caption = "": LabFAceptado.Tag = ""
    labTitAceptado.Caption = "Aceptado:"
    labFPresupuesto.Caption = ""
    labFReparado.Caption = ""
    labTecnico.Caption = ""
    lReclamo.Caption = ""
    tPresupuesto.Text = ""
    tCantidad.Text = ""
    cMoneda.Text = ""
    tCosto.Text = ""
    labCostoGrilla.Caption = ""
    cMonedaFinal.Text = ""
    caCostoFinal.Clean: caCostoFinal.Tag = "0"
    tComentarioP.Text = ""
    
End Sub
Private Sub CargoDatosCliente(idCliente As Long)
Dim RsCli As rdoResultset

    Cons = "Select * from Cliente " _
                & " Left Outer Join CPersona ON CliCodigo = CPeCliente " _
                & " Left Outer Join CEmpresa ON CliCodigo = CEmCliente " _
           & " Where CliCodigo = " & idCliente
           
    Set RsCli = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If Not RsCli.EOF Then       'CI o RUC
        With LabCliente
            Select Case RsCli!CliTipo
                Case TipoCliente.Cliente
                    If Not IsNull(RsCli!CliCiRuc) Then .Caption = " (" & clsGeneral.RetornoFormatoCedula(RsCli!CliCiRuc) & ")"
                    .Caption = .Caption & " " & Trim(Trim(Format(RsCli!CPeApellido1, "#")) & " " & Trim(Format(RsCli!CPeApellido2, "#"))) & ", " & Trim(Trim(Format(RsCli!CPeNombre1, "#")) & " " & Trim(Format(RsCli!CPeNombre2, "#")))
                Case TipoCliente.Empresa
                    If Not IsNull(RsCli!CliCiRuc) Then .Caption = " (" & Trim(RsCli!CliCiRuc) & ")"
                    If Not IsNull(RsCli!CEmNombre) Then .Caption = .Caption & " " & Trim(RsCli!CEmNombre)
                    If Not IsNull(RsCli!CEmFantasia) Then .Caption = .Caption & " (" & Trim(RsCli!CEmFantasia) & ")"
            End Select
            .Tag = RsCli!CliCodigo
        End With
    End If
    RsCli.Close
    
    labTelefono.Caption = " " & TelefonoATexto(idCliente)     'Telefonos
    
End Sub
Private Sub OcultoCamposPresupuesto()
    vsMotivo.BackColor = Inactivo
    tPresupuesto.Enabled = False: tPresupuesto.BackColor = Inactivo
    tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tCosto.Enabled = False: tCosto.BackColor = Inactivo
    cMonedaFinal.Enabled = False: cMonedaFinal.BackColor = Inactivo
    caCostoFinal.Enabled = False: caCostoFinal.BackColorDisplay = Inactivo
    tComentarioP.Enabled = False: tComentarioP.BackColor = Inactivo
End Sub
Private Sub MuestroCamposPresupuesto()
    vsMotivo.BackColor = Blanco
    tPresupuesto.Enabled = True: tPresupuesto.BackColor = Blanco
    tCantidad.Enabled = True: tCantidad.BackColor = Blanco
    cMoneda.Enabled = True: cMoneda.BackColor = Blanco
    tCosto.Enabled = True: tCosto.BackColor = Blanco
    cMonedaFinal.Enabled = True: cMonedaFinal.BackColor = Blanco
    caCostoFinal.Enabled = True: caCostoFinal.BackColorDisplay = Blanco
    tComentarioP.Enabled = True: tComentarioP.BackColor = Blanco
    If labMotivo.Tag = "1" Then tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
End Sub
Private Sub CargoDatosTaller(IdServicio As Long, IdLocalRepara As Integer)
    
    Cons = "Select * From Taller Where TalServicio = " & IdServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "Este servicio no tiene ingreso en taller, verifique.", vbExclamation, "ATENCIÓN"
    Else
        If Not IsNull(RsAux!TalFIngresoRecepcion) Then
            PresentoDatosTaller
        Else
            RsAux.Close
            MsgBox "Este servicio está en traslado al local de reparación.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub
Private Sub PresentoDatosTaller()
    LabFAceptado.Tag = RsAux!TalModificacion
    'Veo las distintas posibilidades.-------------------
    If Not IsNull(RsAux!TalFPresupuesto) Then labFPresupuesto.Caption = Format(RsAux!TalFPresupuesto, FormatoFP)
    If Not IsNull(RsAux!TalFAceptacion) Then
        LabFAceptado.Caption = Format(RsAux!TalFAceptacion, FormatoFP)
        If Not RsAux!TalAceptado Then LabFAceptado.BackColor = Rojo: LabFAceptado.ForeColor = vbWhite: labTitAceptado.Caption = "No Aceptado:"
    End If
    If Not IsNull(RsAux!TalMonedaCosto) Then BuscoCodigoEnCombo cMoneda, RsAux!TalMonedaCosto
    If Not IsNull(RsAux!TalCostoTecnico) Then tCosto.Text = Format(RsAux!TalCostoTecnico, FormatoMonedaP)
    If Not IsNull(RsAux!TalComentario) Then tComentarioP.Text = modStart.f_QuitarClavesDelComentario(Trim(RsAux!TalComentario))
    If Not IsNull(RsAux!TalTecnico) Then labTecnico.Caption = BuscoUsuario(RsAux!TalTecnico, True)
    If Not IsNull(RsAux!TalFReparado) Then labFReparado.Caption = Format(RsAux!TalFReparado, FormatoFP)
    If Not IsNull(RsAux!TalFPresupuesto) Then MeBotones True, False, False
    CargoRenglones
End Sub
Private Sub MeBotones(Modif As Boolean, Grabo As Boolean, Cance As Boolean)
    bModificar.Enabled = Modif: MnuOpModificar.Enabled = Modif
    bGrabar.Enabled = Grabo: MnuOpGrabar.Enabled = Grabo
    bCancelar.Enabled = Cance: MnuOpCancelar.Enabled = Cance
    bNextServicio.Enabled = Modif: MnuOpNextServicio.Enabled = Modif
End Sub
Private Sub BuscoPresupuestoXCodigo(IdPresupuesto As Long)
On Error GoTo ErrBP
    Screen.MousePointer = 11
    Cons = "Select * From Presupuesto " _
        & " Where PreCodigo = " & IdPresupuesto & " And PreEsPresupuesto = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un presupuesto con ese código.", vbInformation, "ATENCIÓN"
    Else
        'Veo si lo ingresó
        tPresupuesto.Tag = RsAux!PreID
        RsAux.Close
        AgregoMotivo tPresupuesto.Tag, True
        tPresupuesto.Text = ""
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el presupuesto.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoPresupuestoXNombre()
On Error GoTo ErrBP
Dim aValor As Long
    Screen.MousePointer = 11
    Cons = "Select ID = PreID, Código = PreCodigo, Nombre = PreNombre From Presupuesto " _
        & " Where PreNombre Like '" & Replace(tPresupuesto.Text, " ", "%") & "%' And PreEsPresupuesto = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un presupuesto con ese nombre.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            aValor = RsAux!ID
            RsAux.Close
            AgregoMotivo aValor, True
            tPresupuesto.Text = ""
        Else
            RsAux.Close
            Dim objLista As New clsListadeAyuda
            If objLista.ActivarAyuda(cBase, Cons, 4500, 1, "Ayuda") > 0 Then
                aValor = objLista.RetornoDatoSeleccionado(0)
            End If
            Set objLista = Nothing
            If aValor <> 0 Then AgregoMotivo aValor, True: tPresupuesto.Text = ""
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el presupuesto por nombre.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoArticuloXCodigo(IdPresupuesto As Long)
On Error GoTo ErrBP
    Screen.MousePointer = 11
    Cons = "Select * From Articulo " _
        & " Where ArtCodigo = " & IdPresupuesto
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un artículo con ese código.", vbInformation, "ATENCIÓN"
    Else
        tPresupuesto.Text = Trim(RsAux!ArtNombre)
        tPresupuesto.Tag = RsAux!ArtID
        RsAux.Close
        tCantidad.Text = "1"
        Foco tCantidad
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el presupuesto.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoArticuloXNombre()
On Error GoTo ErrBP
    Screen.MousePointer = 11
    Cons = "Select ArtID, Código = ArtCodigo, Nombre = ArtNombre From Articulo " _
        & " Where ArtNombre Like '" & Replace(tPresupuesto.Text, " ", "%") & "%'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un artículo con ese nombre.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            tPresupuesto.Text = Trim(RsAux!Nombre)
            tPresupuesto.Tag = RsAux!ArtID
            RsAux.Close
            tCantidad.Text = "1"
            Foco tCantidad
        Else
            RsAux.Close
            Dim objLista As New clsListadeAyuda
            Dim aValor As Long
            If objLista.ActivarAyuda(cBase, Cons, 4500, 1, "Ayuda") > 0 Then
                aValor = objLista.RetornoDatoSeleccionado(0)
            End If
            Set objLista = Nothing
            If aValor <> 0 Then
                Cons = "Select * From Articulo Where ArtID = " & aValor
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                tPresupuesto.Text = Trim(RsAux!ArtNombre)
                tPresupuesto.Tag = RsAux!ArtID
                RsAux.Close
                tCantidad.Text = "1"
                Foco tCantidad
            End If
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo por nombre.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function RetornoPrecioVigenteEnPesos(IDArticulo As Long) As String
Dim RsPD As rdoResultset
Dim TC As Currency
    RetornoPrecioVigenteEnPesos = "0.00"
    Cons = "Select * From PrecioVigente Where PViArticulo  = " & IDArticulo & _
            " And PViTipoCuota = " & paTipoCuotaContado & " And PViMoneda = " & paMonedaPesos
    Set RsPD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsPD.EOF Then
        RetornoPrecioVigenteEnPesos = Format(RsPD!PViPrecio, FormatoMonedaP)
    End If
    RsPD.Close
End Function

Private Function RetornoPrecioDolarEnPesos(IDArticulo As Long) As String
Dim RsPD As rdoResultset
Dim TC As Currency
    RetornoPrecioDolarEnPesos = "0.00"
    Cons = "Select * From PrecioVigente Where PViArticulo  = " & IDArticulo & _
            " And PViTipoCuota = " & paTipoCuotaContado & " And PViMoneda = " & paMonedaDolar
    Set RsPD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsPD.EOF Then
        'Encontre un precio en dolar ahora lo convierto a pesos.
        TC = TasadeCambio(CInt(paMonedaDolar), CInt(paMonedaPesos), gFechaServidor)
        RetornoPrecioDolarEnPesos = Format(RsPD!PViPrecio * TC, FormatoMonedaP)
    End If
    RsPD.Close
End Function
Private Sub AgregoMotivo(idMotivo As Long, Optional Presupuesto As Boolean = False)
    On Error GoTo errAgregar
    
    Screen.MousePointer = 11
    Dim aValor As Long, I As Integer
        
    If Presupuesto Then
        Cons = "Select * from PresupuestoArticulo, Articulo " & _
                                 "Left Outer Join PrecioVigente On ArtID = PViArticulo " & _
                                                                          " And PViTipoCuota = " & paTipoCuotaContado & _
                                                                          " And PViMoneda = " & paMonedaPesos & _
                    " Where PArPresupuesto = " & idMotivo & _
                    " And PArArticulo = ArtID"
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        Do While Not RsAux.EOF
            With vsMotivo
                If Not ArticuloIngresado(RsAux!ArtID) Then
                    .AddItem RsAux!PArCantidad
                    aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                    .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                    If Not IsNull(RsAux!PViPrecio) Then
                        .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PViPrecio, FormatoMonedaP)
                    Else
                        .Cell(flexcpText, .Rows - 1, 2) = RetornoPrecioDolarEnPesos(RsAux!ArtID)
                    End If
                End If
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        'Agrego el Articulo del Preosupuesto (Bonificacion)
        Cons = "Select * from Presupuesto, Articulo" & _
                   " Where PreID = " & idMotivo & _
                   " And PreArticulo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            With vsMotivo
                If Not ArticuloIngresado(RsAux!ArtID) Then
                    .AddItem "1"
                    aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                    .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                    .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PreImporte, FormatoMonedaP)
                End If
            End With
        End If
        RsAux.Close
        tPresupuesto.Text = "": tCantidad.Text = "": Foco tPresupuesto
    Else
        Cons = "Select * from Articulo " & _
                                 "Left Outer Join PrecioVigente On ArtID = PViArticulo " & _
                                                                          " And PViTipoCuota = " & paTipoCuotaContado & _
                                                                          " And PViMoneda = " & paMonedaPesos & _
                    " Where ArtId = " & idMotivo
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            With vsMotivo
                If Not ArticuloIngresado(RsAux!ArtID) Then
                    .AddItem tCantidad.Text
                    aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                    .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                    If Not IsNull(RsAux!PViPrecio) Then
                        .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PViPrecio, FormatoMonedaP)
                    Else
                        .Cell(flexcpText, .Rows - 1, 2) = RetornoPrecioDolarEnPesos(RsAux!ArtID)
                    End If
                End If
            End With
        End If
        RsAux.Close
        tPresupuesto.Text = "": tCantidad.Text = "": Foco tPresupuesto
    End If
    PongoTotalEnEtiqueta
    Screen.MousePointer = 0
    Exit Sub
    
errAgregar:
    clsGeneral.OcurrioError "Ocurrió un error al agregar el item a la lista.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function ArticuloIngresado(IDArticulo As Long) As Boolean

    On Error GoTo errFunction
    ArticuloIngresado = True
    With vsMotivo
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 0) = IDArticulo Then
                MsgBox "El artículo " & .Cell(flexcpText, I, 1) & " ya está ingresado en la lista." & Chr(vbKeyReturn) & "Para modifcar la cantidad elimínelo de la lista y vuelva a ingresarlo.", vbInformation, "Item Ingresado"
                Screen.MousePointer = 0: Exit Function
            End If
        Next
    End With
    '-----------------------------------------------------------------------------------------------------
    ArticuloIngresado = False
    Exit Function

errFunction:
End Function

Private Sub AccionSiguienteServicio()
On Error GoTo errASS
    Screen.MousePointer = 11
    Cons = "Select * from Servicio, Taller" & _
                " Where SerCodigo = TalServicio And TalFPresupuesto is Not Null " & _
                " And SerCostoFinal is null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        tServicio.Text = RsAux!SerCodigo
        RsAux.Close
        CargoServicio
        If vsHistoria.Rows > 1 Then TabValidacion.Tabs(1).Selected = True
        AccionModificar
    Else
        RsAux.Close
        MsgBox "No hay servicios pendientes para presupuestar.", vbInformation, "ATENCIÓN"
    End If
    Screen.MousePointer = 0
    Exit Sub
errASS:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el siguiente servicio.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub AccionGrabar()
    If Not ValidoDatos Then Exit Sub
    If MsgBox("¿Confirma validar el presupuesto?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    GraboFicha
End Sub

Private Function ModificoLista() As Boolean
On Error GoTo errML
Dim miCol As New Collection
Dim iCont As Integer, iCont1 As Integer, bEsta As Boolean

    ModificoLista = False
    
    Cons = "Select * From ServicioRenglon Where SReServicio = " & tServicio.Tag _
        & " And SReTipoRenglon = " & TipoRenglonS.Cumplido
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        miCol.Add RsAux!SReMotivo & "C" & RsAux!SReCantidad
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Cargue lo anterior.
    If miCol.Count = 0 Then
        If vsMotivo.Rows > 1 Then
            ModificoLista = True: Exit Function
        Else
            Exit Function
        End If
    Else
        If miCol.Count <> vsMotivo.Rows - 1 Then
            'Quito o agrego.
            ModificoLista = True: Exit Function
        Else
            'Recorro la colección, la cantidad de elementos es la misma solo verifico que no este o que modifico la cantidad.
            With vsMotivo
                For iCont = 1 To .Rows - 1
                    bEsta = False
                    For iCont1 = 1 To miCol.Count
                        If Val(.Cell(flexcpData, iCont, 0)) = Val(Mid(miCol.Item(iCont1), 1, InStr(1, miCol.Item(iCont1), "C") - 1)) Then
                            If Val(.Cell(flexcpText, iCont, 0)) = Val(Mid(miCol.Item(iCont1), InStr(1, miCol.Item(iCont1), "C") + 1)) Then
                                bEsta = True
                                Exit For
                            End If
                        End If
                    Next iCont1
                    If Not bEsta Then
                        ModificoLista = True: Exit Function
                    End If
                Next iCont
            End With
        End If
    End If
    Exit Function
errML:
End Function
Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    
    If bAvisarStock Then
        If ModificoLista Then
            If MsgBox("Al modificar la lista de repuestos se harán movimientos de stock para los mismos." & vbCrLf _
                & "¿Confirma continuar?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                    Exit Function
            End If
        End If
    End If
    
    If Not clsGeneral.TextoValido(tComentarioP.Text) Then MsgBox "Ingreso alguna comilla simple en el comentario del servicio.", vbExclamation, "ATENCIÓN": Foco tComentarioP: Exit Function
    If tCosto.Text <> "" Then
        If Not IsNumeric(tCosto.Text) Then MsgBox "El costo debe ser numérico.", vbExclamation, "ATENCIÓN": Foco tCosto: Exit Function
        If cMoneda.ListIndex = -1 Then MsgBox "Seleccione una moneda.", vbExclamation, "ATENCIÓN": Foco cMoneda: Exit Function
    End If
    
    If cMonedaFinal.ListIndex = -1 Then MsgBox "Seleccione una moneda.", vbExclamation, "ATENCIÓN": Foco cMonedaFinal: Exit Function
    
    
    ValidoDatos = True
End Function

Private Sub AccionModificar()
Dim cAux As Currency
    
    tServicio.Text = tServicio.Tag
    CargoServicio
    If Val(LabCliente.Tag) > 0 Then loc_FindComentarios Val(LabCliente.Tag): loc_FindCuotaVencida Val(LabCliente.Tag)
    
    If tServicio.Text <> tServicio.Tag Or tServicio.Tag = "" Then Exit Sub
    MeBotones False, True, True
    
    MuestroCamposPresupuesto
    tServicio.Enabled = False
    
    If caCostoFinal.Tag = "0" Then
        BuscoCodigoEnCombo cMonedaFinal, paMonedaPesos
        caCostoFinal.Text = "0"
        
        If Not (Val(LabEstado.Tag) = EstadoP.SinCargo Or Val(LabCliente.Tag) = paClienteEmpresa) Then
        
            If Val(LabEstado.Tag) = EstadoP.FueraGarantia And Val(labEstadoServicio.Tag) <> EstadoS.Cumplido Then
                
                'Le doy un costo en base a los precios de la grilla.
                With vsMotivo
                    For I = 1 To .Rows - 1
                        If IsNumeric(.Cell(flexcpText, I, 2)) Then cAux = cAux + (CCur(.Cell(flexcpText, I, 2)) * Val(.Cell(flexcpValue, I, 0)))
                    Next I
                End With
                If cAux <> 0 Then caCostoFinal.Text = cAux: cAux = 0 Else cAux = -1
            Else
                caCostoFinal.Text = 0
            End If
            
        End If
    End If
    
    If labEstadoServicio.Tag = EstadoS.Cumplido Then
        If Val(LabCliente.Tag) <> paClienteEmpresa Then
            OcultoCamposPresupuesto
        Else
            Dim sFCosteo As String
            sFCosteo = ""
            Cons = "Select Max(CabMesCosteo) From CMCabezal"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                If Not IsNull(RsAux(0)) Then sFCosteo = RsAux(0)
            End If
            RsAux.Close
            
            Cons = "Select SerFCumplido From Servicio Where SerCodigo = " & tServicio.Text
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If RsAux!SerFCumplido < CDate(UltimoDia(CDate(sFCosteo)) + 1) Then
                RsAux.Close
                OcultoCamposPresupuesto
            Else
                RsAux.Close
                bAvisarStock = True
            End If
        End If
        cMoneda.Enabled = True: cMoneda.BackColor = vbWhite
        tCosto.Enabled = True: tCosto.BackColor = vbWhite
        tComentarioP.Enabled = True: tComentarioP.BackColor = vbWhite
        If caCostoFinal.Enabled Then Foco caCostoFinal
    Else
        If labFReparado.Caption <> "" Then bAvisarStock = True
        If caCostoFinal.Enabled Then
            Foco caCostoFinal
            If cAux = -1 And caCostoFinal.Text = 0 Then caCostoFinal.Clean
        Else
            Foco tPresupuesto
        End If
    End If
    
End Sub

Private Sub GraboFicha()
Dim IdLocalReparacion As Long
Dim RsArticulo As rdoResultset
Dim Aceptado As Boolean: Aceptado = False
Dim sCumplo As Boolean: sCumplo = False
Dim bQuitoAceptado As Boolean: bQuitoAceptado = False
    
    Screen.MousePointer = 11
    If LabEstado.Tag = EstadoP.SinCargo And LabFAceptado.Caption = "" Then
        If caCostoFinal.Text > 0 Then
           ' If MsgBox("El producto está sin cargo. ¿Desea dar como aceptado el presupuesto?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then Aceptado = True
        Else
            Aceptado = True
        End If
    ElseIf LabFAceptado.Caption = "" Then
        If Val(LabCliente.Tag) = paClienteEmpresa And Val(labEstadoServicio.Tag) = EstadoS.Taller Then
            If caCostoFinal.Text > 0 Then
              '  If MsgBox("El producto es de la empresa. ¿Desea dar como aceptado el presupuesto?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then Aceptado = True
            Else
                Aceptado = True
            End If
        ElseIf caCostoFinal.Text = 0 Then
            Aceptado = True
        End If
    End If
    
    If LabEstado.Tag <> EstadoP.SinCargo And caCostoFinal.Text = 0 And lReclamo.Caption = "" Then
        If MsgBox("Este servicio debería presupuestado. ¿Está seguro que el costo Cero?", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then
            Exit Sub
        End If
        
    End If
    
    If Not Aceptado Then
        Cons = "Select * From Servicio Where SerCodigo = " & tServicio.Tag
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If caCostoFinal.Text > 0 Then
            If Not IsNull(RsAux("SerCostoFinal")) Then
                bQuitoAceptado = (caCostoFinal.Text > RsAux("SerCostoFinal") And LabFAceptado.Caption <> "")
            End If
        End If
        RsAux.Close
        If bQuitoAceptado Then MsgBox "Se eliminará la condición de presupuesto aceptado al aumentar el costo.", vbInformation, "Atención"
    End If
    
    FechaDelServidor
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrVA
    Cons = "Select * From Servicio Where SerCodigo = " & tServicio.Tag
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close: cBase.RollbackTrans
        MsgBox "Otra terminal pudó eliminar el servicio verifique.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    Else
        If RsAux!SerModificacion = CDate(labIngreso.Tag) Then
            IdLocalReparacion = RsAux!SerLocalReparacion
            RsAux.Edit
            RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
            
            If cMonedaFinal.ListIndex = -1 Then
                RsAux!SerMoneda = paMonedaPesos
            Else
                RsAux!SerMoneda = cMonedaFinal.ItemData(cMonedaFinal.ListIndex)
            End If
            
            RsAux!SerCostoFinal = caCostoFinal.Text
            If sCumplo Then RsAux!SerEstadoServicio = EstadoS.Cumplido
            RsAux.Update: RsAux.Close
        Else
            RsAux.Close: cBase.RollbackTrans
            MsgBox "No podrá almacenar los datos debido a que otra terminal modifico los datos, verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    Cons = "Select * From Taller Where TalServicio = " & tServicio.Tag
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close: cBase.RollbackTrans
        MsgBox "Otra terminal pudó eliminar los datos de taller.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    Else
        If RsAux!TalModificacion = CDate(LabFAceptado.Tag) Then
            'Veo si ya fue reparado.
            RsAux.Edit
            If Aceptado Or sCumplo Then
                RsAux!TalFAceptacion = Format(gFechaServidor, sqlFormatoFH)
                RsAux!TalAceptado = 1
            ElseIf bQuitoAceptado Then
                RsAux!TalFAceptacion = Null
                RsAux!TalAceptado = 0
            End If
            RsAux!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
            If cMoneda.ListIndex = -1 Then
                RsAux!TalMonedaCosto = paMonedaPesos
                RsAux!TalCostoTecnico = 0
            Else
                RsAux!TalMonedaCosto = cMoneda.ItemData(cMoneda.ListIndex)
                RsAux!TalCostoTecnico = CCur(tCosto.Text)
            End If
            If Not IsNull(RsAux!TalComentario) Then tComentarioP.Tag = modStart.f_GetEventos(RsAux!TalComentario) Else tComentarioP.Tag = ""
            If Trim(tComentarioP.Text) = "" And Trim(tComentarioP.Tag) = "" Then
                RsAux!TalComentario = Null
            Else
                RsAux!TalComentario = Trim(tComentarioP.Tag) & Trim(tComentarioP.Text)
            End If
            RsAux.Update: RsAux.Close
        Else
            RsAux.Close: cBase.RollbackTrans
            MsgBox "No podrá almacenar los datos debido a que otra terminal modifico los datos, verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    
    If labFReparado.Caption <> "" Then ' Or (labFReparado.Caption = "" And Val(LabCliente.Tag) = paClienteEmpresa And vsMotivo.Enabled) Then
        
        'Or (labFReparado.Caption = "" And Val(LabCliente.Tag) = paClienteEmpresa)
        
        'Estos ya hicieron movimientos físicos.
        Cons = "Select * From ServicioRenglon, Articulo Where SReServicio = " & tServicio.Tag & " And SReTipoRenglon = " & TipoRenglonS.Cumplido & " And SReMotivo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        'En esta primera solo doy de baja a los movimientos físicos ya hechos.
        Dim sEncontre As Boolean
        Do While Not RsAux.EOF
            sEncontre = False
            With vsMotivo
                For I = 1 To .Rows - 1
                    If RsAux!SReMotivo = CLng(.Cell(flexcpData, I, 0)) Then
                        sEncontre = True
                        If CLng(.Cell(flexcpData, I, 1)) <> paTipoArticuloServicio Then
                            If RsAux!SReCantidad > CCur(.Cell(flexcpText, I, 0)) Then
                                'Doy alta a los  movimientos y alta al local.
                                MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, IdLocalReparacion, CLng(.Cell(flexcpData, I, 0)), RsAux!SReCantidad - CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, 1, TipoDocumento.Servicio, tServicio.Tag
                                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, IdLocalReparacion, CLng(.Cell(flexcpData, I, 0)), RsAux!SReCantidad - CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, 1
                                MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, RsAux!SReCantidad - CCur(.Cell(flexcpText, I, 0)), 1
                                
                            ElseIf RsAux!SReCantidad < CCur(.Cell(flexcpText, I, 0)) Then
                                'Doy baja en los movimientos físicos y baja en el local.
                                MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, IdLocalReparacion, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)) - RsAux!SReCantidad, paEstadoArticuloEntrega, -1, TipoDocumento.Servicio, tServicio.Tag
                                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, IdLocalReparacion, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)) - RsAux!SReCantidad, paEstadoArticuloEntrega, -1
                                MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpText, I, 0)) - RsAux!SReCantidad, -1
                            End If
                            'Este ya estaba. si la cantidad es la misma le pongo que es del tipo servicio
                            'para no hacer el movimiento abajo.
                            .Cell(flexcpData, I, 1) = paTipoArticuloServicio
                        End If
                    End If
                    If sEncontre Then Exit For
                Next I
            End With
            If Not sEncontre Then
                'Me borro este artículo de la lista.
                'Veo si el artículo es del tipo servicio, sino le doy la baja al movimiento.
                If RsAux!ArtTipo <> paTipoArticuloServicio Then
                    MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, IdLocalReparacion, RsAux!ArtID, RsAux!SReCantidad, paEstadoArticuloEntrega, 1, TipoDocumento.Servicio, tServicio.Tag
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, IdLocalReparacion, RsAux!ArtID, RsAux!SReCantidad, paEstadoArticuloEntrega, 1
                    MarcoMovimientoStockTotal RsAux!ArtID, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, RsAux!SReCantidad, 1
                End If
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close
        With vsMotivo
            For I = 1 To .Rows - 1
                If CLng(.Cell(flexcpData, I, 1)) <> paTipoArticuloServicio Then
                    MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, IdLocalReparacion, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, -1, TipoDocumento.Servicio, tServicio.Tag
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, IdLocalReparacion, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, -1
                    MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpText, I, 0)), -1
                End If
            Next I
        End With
    End If
    Cons = "Delete ServicioRenglon Where SReServicio = " & tServicio.Tag & " And SReTipoRenglon = " & TipoRenglonS.Cumplido
    cBase.Execute (Cons)
    With vsMotivo
        For I = 1 To .Rows - 1
            Cons = "Insert Into ServicioRenglon (SReServicio, SReTipoRenglon, SReMotivo, SReCantidad, SReTotal) Values (" _
                & tServicio.Tag & ", " & TipoRenglonS.Cumplido & ", " & Val(.Cell(flexcpData, I, 0)) & ", " & Val(.Cell(flexcpText, I, 0)) & ", " & CCur(.Cell(flexcpText, I, 2)) & ")"
            cBase.Execute (Cons)
        Next I
    End With
    
    cBase.CommitTrans
    On Error Resume Next
    Screen.MousePointer = 0
    CargoServicio
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Error al iniciar la transacción, reintente.", Err.Description
    Screen.MousePointer = 0: Exit Sub
ErrVA:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub

Private Sub vsHistoria_DblClick()
    Call bHistoria_Click
End Sub

Private Sub vsMotivo_DblClick()
Dim aValor As Currency
    If vsMotivo.BackColor = Inactivo Then Exit Sub
    FechaDelServidor
    With vsMotivo
        If .Row > 0 Then
            If CCur(.Cell(flexcpValue, .Row, 2)) > 0 Then
                If MsgBox("¿Desea actualizar el precio del repuesto?", vbQuestion + vbYesNo, "ACTUALIZAR PRECIO") = vbNo Then Exit Sub
            End If
            aValor = RetornoPrecioVigenteEnPesos(.Cell(flexcpData, .Row, 0))
            If aValor = 0 Then aValor = RetornoPrecioDolarEnPesos(.Cell(flexcpData, .Row, 0))
            If aValor > 0 Then
                labCostoGrilla.Caption = CCur(labCostoGrilla.Caption) - CCur(.Cell(flexcpValue, .Row, 2))
                .Cell(flexcpText, .Row, 2) = Format(aValor, FormatoMonedaP)
                labCostoGrilla.Caption = Format(CCur(labCostoGrilla.Caption) + CCur(.Cell(flexcpValue, .Row, 2)), FormatoMonedaP)
                Exit Sub
            End If
            If MsgBox("No existe un precio vigente para este artículo." & vbCr & "¿Desea ingresarlo?", vbQuestion + vbYesNo, "INGRESO DE PRECIO") = vbNo Then Exit Sub
            EjecutarApp App.Path & "\Articulos.exe", "P" & .Cell(flexcpData, .Row, 0), True
            FechaDelServidor
            aValor = RetornoPrecioVigenteEnPesos(.Cell(flexcpData, .Row, 0))
            If aValor = 0 Then aValor = RetornoPrecioDolarEnPesos(.Cell(flexcpData, .Row, 0))
            If aValor > 0 Then
                labCostoGrilla.Caption = CCur(labCostoGrilla.Caption) - CCur(.Cell(flexcpValue, .Row, 2))
                .Cell(flexcpText, .Row, 2) = Format(aValor, FormatoMonedaP)
                labCostoGrilla.Caption = Format(CCur(labCostoGrilla.Caption) + CCur(.Cell(flexcpValue, .Row, 2)), FormatoMonedaP)
            End If
        End If
    End With
End Sub

Private Sub vsMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsMotivo.BackColor = Inactivo Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn: Foco cMoneda
        Case vbKeyDelete: If vsMotivo.Row > 0 Then vsMotivo.RemoveItem vsMotivo.Row: PongoTotalEnEtiqueta
    End Select
End Sub

Private Sub CargoRenglones()
Dim RsSR As rdoResultset
Dim aValor As Long
Dim curPrecio As Currency

    vsMotivo.Rows = 1
    Cons = "Select * From ServicioRenglon, Articulo Where SReServicio = " & tServicio.Tag _
        & " And SReTipoRenglon = " & TipoRenglonS.Cumplido & " And SReMotivo = ArtID"
    Set RsSR = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsSR.EOF
        With vsMotivo
            .AddItem RsSR!SReCantidad
            aValor = RsSR!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            aValor = RsSR!ArtTipo: .Cell(flexcpData, .Rows - 1, 1) = aValor
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsSR!ArtCodigo, "(#,000,000)") & " " & Trim(RsSR!ArtNombre)
            If labCostoGrilla.Caption = "" Then labCostoGrilla.Caption = "0.00"
            
            labCostoGrilla.Caption = Format(CCur(labCostoGrilla.Caption) + RsSR!SReTotal, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsSR!SReTotal, FormatoMonedaP)
        End With
        RsSR.MoveNext
    Loop
    RsSR.Close
    
    If Val(labEstadoServicio.Tag) <> EstadoS.Anulado And Val(labEstadoServicio.Tag) <> EstadoS.Cumplido Then
        
        'Verifico si cambio los precios de cada artículo.
        With vsMotivo
            For I = 1 To .Rows - 1
                curPrecio = RetornoPrecioVigenteEnPesos(.Cell(flexcpData, I, 0))
                If curPrecio = 0 Then curPrecio = RetornoPrecioDolarEnPesos(.Cell(flexcpData, I, 0))
                If curPrecio <> 0 And Val(.Cell(flexcpText, I, 2)) = 0 Then
                    .Cell(flexcpText, I, 2) = Format(curPrecio, FormatoMonedaP)
                End If
            Next I
        End With
        PongoTotalEnEtiqueta
    End If
End Sub

Private Sub CargoHistoria(idProducto As Long, vsGrilla As vsFlexGrid)
Dim IdServicio As Long: IdServicio = 0
Dim aComentario As String
    
    On Error GoTo ErrCH
    
    If Val(vsGrilla.Tag) = idProducto Then Exit Sub Else vsGrilla.Tag = idProducto
    
    Screen.MousePointer = 11
    vsGrilla.Rows = 1
        
    Cons = "Select * From Servicio" _
            & " Left Outer Join ServicioRenglon ON SReTipoRenglon = " & TipoRenglonS.Cumplido _
                                                                    & " And SerCodigo = SReServicio" _
            & " Left Outer Join Articulo ON SReMotivo = ArtID " _
        & " Where SerProducto = " & idProducto _
        & " And SerEstadoServicio In (" & EstadoS.Cumplido & ", " & EstadoS.Anulado & ")" _
        & " Order By SerCodigo DESC"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        With vsGrilla
            .AddItem ""
            If RsAux!SerEstadoServicio = EstadoS.Anulado Then .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!SerFCumplido, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 2) = EstadoProducto(RsAux!SerEstadoProducto)
            If Not IsNull(RsAux!SerMoneda) And Not IsNull(RsAux!SerCostoFinal) Then .Cell(flexcpText, .Rows - 1, 3) = BuscoSignoMoneda(RsAux!SerMoneda) & " " & Format(RsAux!SerCostoFinal, FormatoMonedaP)
            If Not IsNull(RsAux!SerComentarioR) Then aComentario = Trim(RsAux!SerComentarioR) Else aComentario = ""
            IdServicio = RsAux!SerCodigo
            Do While Not RsAux.EOF
                If Not IsNull(RsAux!ArtNombre) Then
                    If Trim(.Cell(flexcpText, .Rows - 1, 1)) <> "" Then .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & ", "
                    .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & Trim(RsAux!ArtNombre)
                End If
                IdServicio = RsAux!SerCodigo
                RsAux.MoveNext
                If RsAux.EOF Then Exit Do
                If IdServicio <> RsAux!SerCodigo Then Exit Do
            Loop
            If Trim(.Cell(flexcpText, .Rows - 1, 1)) <> "" Then .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & Chr(13)
            .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & aComentario
        End With
    Loop
    RsAux.Close
    With vsGrilla
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 1, , False
    End With
    Screen.MousePointer = 0
    Exit Sub
ErrCH:
    clsGeneral.OcurrioError "Ocurrio un error al buscar la historia del producto.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoOtrosProducto()
Dim aValor As Long

    If vsProducto.Rows > 1 Or Val(LabCliente.Tag) = 0 Then Exit Sub
    On Error GoTo ErrCOP
    Screen.MousePointer = 11
    If Val(LabCliente.Tag) = paClienteEmpresa Then
'        Cons = Cons & " And ProFModificacion >= '" & Format(Date - 1, "mm/dd/yyyy 00:00:00") & "'"
        Cons = "Select Top 25 * "
    Else
        Cons = "Select * "
    End If
    Cons = Cons & "From Producto, Articulo Where ProCliente = " & LabCliente.Tag _
        & " And ProCodigo <> " & LabProducto.Tag _
        & " And ProArticulo = ArtID "
    
    If Val(LabCliente.Tag) = paClienteEmpresa Then Cons = Cons & " Order by ProFModificacion Desc"
    
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        With vsProducto
            .AddItem Format(RsAux!ProCodigo, "#,000") & " " & Trim(RsAux!ArtNombre)
            
            aValor = RsAux!ProCodigo
            .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            'Veo si tengo algún reporte abierto.
            .Cell(flexcpData, .Rows - 1, 1) = TieneReporteAbierto(RsAux!ProCodigo)
            If Val(.Cell(flexcpData, .Rows - 1, 1)) > 0 Then .Cell(flexcpPicture, .Rows - 1, 0) = Image1.ListImages("servicio").ExtractIcon
            
            .Cell(flexcpText, .Rows - 1, 1) = EstadoProducto(CalculoEstadoProducto(CInt(aValor)), True)
            
            If Not IsNull(RsAux!ProCompra) Then .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ProCompra, "dd/mm/yyyy")
            'Saco garantia-----------------------------------------------------------------------------------------------------------------------
            .Cell(flexcpText, .Rows - 1, 3) = RetornoGarantia(RsAux!ArtID)
            '--------------------------------------------------------------------------------------------------------------------------------------
            .Cell(flexcpData, .Rows - 1, 4) = TieneReporteAbierto(RsAux!ProCodigo)
            If Not IsNull(RsAux!ProNroSerie) Then .Cell(flexcpText, .Rows - 1, 4) = RsAux!ProNroSerie
            If Not IsNull(RsAux!ProFacturaS) Then .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!ProFacturaS) & " "
            If Not IsNull(RsAux!ProFacturaN) Then .Cell(flexcpText, .Rows - 1, 5) = .Cell(flexcpText, .Rows - 1, 5) & Trim(RsAux!ProFacturaN)
        End With

        RsAux.MoveNext
    Loop
    RsAux.Close
    If vsProducto.Rows > 1 Then CargoHistoria vsProducto.Cell(flexcpData, 1, 0), vsHistoria2
    Screen.MousePointer = 0
    Exit Sub
    
ErrCOP:
    clsGeneral.OcurrioError "Ocurrio un error al intentar cargar los productos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function TieneReporteAbierto(idProducto As Long) As Long
On Error GoTo ErrTRA
Dim RsSA As rdoResultset
    Screen.MousePointer = 11
    TieneReporteAbierto = 0
    Cons = "Select * From Servicio Where SerProducto = " & idProducto _
        & " And SerEstadoServicio Not IN (" & EstadoS.Anulado & ", " & EstadoS.Cumplido & ")"
    Set RsSA = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsSA.EOF Then TieneReporteAbierto = RsSA!SerCodigo
    RsSA.Close
    Screen.MousePointer = 0
    Exit Function
ErrTRA:
    clsGeneral.OcurrioError "Ocurrio un error al verificar si existe algún servicio abierto.", Trim(Err.Description)
    Screen.MousePointer = 0
End Function

Private Sub vsProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsProducto.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn: CargoHistoria Val(vsProducto.Cell(flexcpData, vsProducto.Row, 0)), vsHistoria2
        Case vbKeyAdd: If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) > 0 Then EjecutarApp App.Path & "\Seguimiento de Servicios.exe", vsProducto.Cell(flexcpData, vsProducto.Row, 4)
    End Select
End Sub

Private Sub vsProducto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 2 And vsProducto.Rows > 1 Then PopupMenu MnuAccesos
End Sub

Private Sub vsProducto_RowColChange()
    If vsProducto.Row > 0 Then CargoHistoria Val(vsProducto.Cell(flexcpData, vsProducto.Row, 0)), vsHistoria2
End Sub

Private Sub ConsultaAyuda()

Dim strComentario As String: strComentario = vbNullString
Dim strArticulos As String: strArticulos = vbNullString
Dim lngSuma As Long: lngSuma = 0

    If chRepuesto.Value = 0 Then
        For I = 1 To vsMotivo.Rows - 1
            If strArticulos = "" Then strArticulos = CLng(vsMotivo.Cell(flexcpData, I, 0)) Else strArticulos = strArticulos & ", " & CLng(vsMotivo.Cell(flexcpData, I, 0))
        Next I
    ElseIf chRepuesto.Value = 1 Then
        Cons = "Select Sum(SReCantidad * ArtCodigo) From ServicioRenglon, Articulo" _
            & " Where SReServicio = " & Val(tServicio.Tag) _
            & " And SReTipoRenglon = " & TipoRenglonS.Cumplido & " And SReMotivo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not IsNull(RsAux(0)) Then lngSuma = RsAux(0)
        RsAux.Close
    End If
    
    If chComentario.Value = 1 Then
        strComentario = InputBox("Ingrese un comentario a buscar.", "Bùsqueda por comentario de Técnico.")
        If strComentario = "" Then
            If MsgBox("No ingreso un parámetro de comentario." & Chr(13) & "¿Desea continuar?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
        End If
    End If

    'Cargo los servicios que tengan el mismo tipo de artículo.
    Cons = "Select Distinct(SerCodigo), Servicio = SerCodigo, rtrim(cast(datepart(dd,SerFecha) as char)) + '/' + rtrim(cast(datepart(mm, SerFecha) as char)) + '/' +  rtrim(cast(datepart(yy, SerFecha) as char)) as Fecha,Producto = ArtNombre, 'Costo Final' = SerCostoFinal, 'Mano de Obra' = TalCostoTecnico, Comentario = TalComentario " _
        & " From Servicio"
                
    If chRepuesto.Value = 0 And strArticulos <> "" Then Cons = Cons & ", ServicioRenglon"
        
    Cons = Cons & " , Producto, Articulo, Taller" _
        & " Where SerCodigo <> " & Val(tServicio.Tag) _
        & " And SerLocalReparacion = " & Val(labIDLocalReparacion.Tag) _
        & " And SerEstadoProducto = " & EstadoP.FueraGarantia & " And SerCostoFinal > 0" _
        & " And SerCodigo = TalServicio And SerProducto = ProCodigo And ProArticulo = ArtID"
        
    If chRepuesto.Value = 0 And strArticulos <> "" Then
        Cons = Cons & " And SerCodigo = SReServicio " _
                & " And SReTipoRenglon = " & TipoRenglonS.Cumplido _
                & " And SReMotivo IN (" & strArticulos & ")"
    End If
    
    If chArticulo.Value = 0 Then
        Cons = Cons & " And ArtTipo = (Select ArtTipo From Producto, Articulo" _
                                        & " Where ProCodigo = " & Val(LabProducto.Tag) & " And ProArticulo = ArtID)"
    Else
        Cons = Cons & " And ArtID = (Select ProArticulo From Producto " _
                                & " Where ProCodigo = " & Val(LabProducto.Tag) & " And ProArticulo = ArtID)"
    End If
    
    If chRepuesto.Value = 1 Then
        Cons = Cons & " And SerCodigo IN (Select SReServicio " _
            & " From ServicioRenglon, Articulo Where SReServicio <> " & Val(tServicio.Tag) _
            & " And SReTipoRenglon = " & TipoRenglonS.Cumplido & " And SReMotivo = ArtID Group by SReServicio " _
            & " Having Sum(SReCantidad * ArtCodigo) = " & lngSuma & ")"
    End If
    If strComentario <> "" Then Cons = Cons & " And TalComentario like '%" & Replace(strComentario, " ", "%") & "%'"
    
    Cons = Cons & "Order by SerCodigo DESC"
    InvocoConsulta Cons

End Sub

Private Sub InvocoConsulta(StrConsulta As String)
Dim bCosto As Boolean
    On Error GoTo ErrCPT
    If Me.ActiveControl.Name = "caCostoFinal" Then
        bCosto = True
    End If
    Dim objLista As New clsListadeAyuda
    If objLista.ActivarAyuda(cBase, Cons, 9600, 1, "Lista") > 0 Then
        EjecutarApp App.Path & "\Seguimiento de Servicios", objLista.RetornoDatoSeleccionado(0)
    End If
    Set objLista = Nothing
    On Error Resume Next
    If bCosto Then caCostoFinal.SetFocus
    Exit Sub
ErrCPT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al intentar acceder a la ayuda.", Err.Description
End Sub

Private Sub PongoTotalEnEtiqueta()
On Error GoTo ErrPTEE
    Screen.MousePointer = 11
    labCostoGrilla.Caption = "0"
    With vsMotivo
        For I = 1 To .Rows - 1
            labCostoGrilla.Caption = Format(CCur(labCostoGrilla.Caption) + (Val(.Cell(flexcpValue, I, 0)) * Val(.Cell(flexcpValue, I, 2))), FormatoMonedaP)
        Next
    End With
    Screen.MousePointer = 0
    Exit Sub
ErrPTEE:
    clsGeneral.OcurrioError "Ocurrio un error al intentar calcular el total de los repuestos.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub AccionAyuda()
On Error GoTo errHelp
    Screen.MousePointer = 11
    Dim aFile As String
    Cons = "Select * from Aplicacion Where AplNombre = '" & Trim(App.Title) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux!AplHelp) Then aFile = Trim(RsAux!AplHelp)
    RsAux.Close
    If aFile <> "" Then EjecutarApp aFile
    Screen.MousePointer = 0
    Exit Sub
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub
Public Sub loc_FindComentarios(idCliente As Long)
Dim RsCom As rdoResultset
Dim bHay As Boolean

    On Error GoTo errMenu
    Screen.MousePointer = 11
    bHay = False
    
    Cons = "Select * From Comentario " _
            & " Where ComCliente = " & idCliente & " And ComTipo IN (" & prmTipoComentario & ")"
            
    Set RsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsCom.EOF Then bHay = True
    RsCom.Close
    If Not bHay Then Screen.MousePointer = 0: Exit Sub
    
    Dim objC As New clsCliente
    objC.Comentarios idCliente:=idCliente
    Set objC = Nothing
    Me.Refresh
    Screen.MousePointer = 0
    Exit Sub
    
errMenu:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al fomulario de comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_FindCuotaVencida(ByVal idCliente As Long)
Dim rsFCV As rdoResultset

    If idCliente = paClienteEmpresa Or InStr(1, paCliCuoNoVen, "," & idCliente & ",", vbTextCompare) > 0 Then Exit Sub
    
    Cons = "Select Top 1 *  From Documento, Credito " _
        & "Where DocCliente = " & idCliente _
        & " And CreProximoVto < '" & Format(Date - 30, "mm/dd/yyyy") & "'" _
        & " And DocTipo = 2 And CreSaldoFactura > 0 And DocAnulado = 0 And DocCodigo = CreFactura"

    Set rsFCV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsFCV.EOF Then
        rsFCV.Close
        If MsgBox("El cliente tiene cuotas vencidas." & vbCr & "¿Desea visualizarlas?", vbQuestion + vbYesNo, "CUOTAS VENCIDAS") = vbYes Then
            Call bVisualizacion_Click
        End If
    Else
        rsFCV.Close
    End If

End Sub
