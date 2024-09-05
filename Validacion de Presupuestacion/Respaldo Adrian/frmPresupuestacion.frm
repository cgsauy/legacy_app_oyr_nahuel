VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACOMBO.OCX"
Begin VB.Form frmPresupuestacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validación de Presupuestos"
   ClientHeight    =   6075
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPresupuestacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicProductos 
      Height          =   1275
      Left            =   3840
      ScaleHeight     =   1215
      ScaleWidth      =   2055
      TabIndex        =   51
      Top             =   3900
      Width           =   2115
      Begin VSFlex6DAOCtl.vsFlexGrid vsProducto 
         Height          =   1035
         Left            =   60
         TabIndex        =   52
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1826
         _ConvInfo       =   1
         Appearance      =   1
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
         GridLines       =   1
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
         TabIndex        =   53
         Top             =   1140
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1402
         _ConvInfo       =   1
         Appearance      =   1
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
         GridLines       =   1
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
      Height          =   2115
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   6495
      TabIndex        =   42
      Top             =   3000
      Width           =   6555
      Begin VB.TextBox tCostoFinal 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   15
         TabIndex        =   13
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox tCosto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1260
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
         Height          =   285
         Left            =   960
         MaxLength       =   80
         TabIndex        =   15
         Top             =   1980
         Width           =   5415
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   4680
         TabIndex        =   9
         Top             =   1260
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
         Appearance      =   1
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
         GridLines       =   1
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
         Left            =   4680
         TabIndex        =   12
         Top             =   1620
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
      Begin VB.Label labFReparado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1/1/2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5220
         TabIndex        =   50
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label11 
         Caption         =   "Reparado:"
         Height          =   195
         Left            =   3780
         TabIndex        =   49
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Técnico:"
         Height          =   195
         Left            =   3780
         TabIndex        =   48
         Top             =   660
         Width           =   735
      End
      Begin VB.Label labTecnico 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coquito"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5220
         TabIndex        =   47
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label labFPresupuesto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-May 2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5220
         TabIndex        =   46
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label Label9 
         Caption         =   "Presupuestado:"
         Height          =   195
         Left            =   3780
         TabIndex        =   45
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "C&osto Final:"
         Height          =   255
         Left            =   3780
         TabIndex        =   11
         Top             =   1620
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
         Left            =   5220
         TabIndex        =   44
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label labTitAceptado 
         Caption         =   "Aceptado:"
         Height          =   195
         Left            =   3780
         TabIndex        =   43
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "&Mano Obra:"
         Height          =   255
         Left            =   3780
         TabIndex        =   8
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label labMotivo 
         Caption         =   "&Artículo: [F12 a Presupuesto]"
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "Co&mentario:"
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   1980
         Width           =   855
      End
   End
   Begin VB.PictureBox picHistoria 
      Height          =   735
      Left            =   3780
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   41
      Top             =   3000
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
         Appearance      =   1
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
         GridLines       =   1
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
      TabIndex        =   27
      Top             =   5820
      Width           =   6795
      _ExtentX        =   11986
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
      Height          =   2475
      Left            =   60
      TabIndex        =   22
      Top             =   60
      Width           =   6615
      Begin VB.TextBox tComentarioS 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   1920
         Width           =   6375
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
      Begin VB.Label labFactura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   40
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label LabEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   6060
         TabIndex        =   37
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label LabGarantia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12 Meses"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5040
         TabIndex        =   36
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Garantía:"
         Height          =   195
         Left            =   4320
         TabIndex        =   35
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label LabFCompra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-May 2000"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "F/Compra:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Factura:"
         Height          =   255
         Left            =   2340
         TabIndex        =   16
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label labIngreso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-May 2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5040
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Ingreso:"
         Height          =   195
         Left            =   4320
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label labTelefono 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   30
         Top             =   900
         Width           =   5535
      End
      Begin VB.Label Label4 
         Caption         =   "Teléfonos:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   26
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label LabCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WALTER ADRIAN OCCHIUZZI MARTINEZ"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label LabProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   23
         Top             =   1260
         Width           =   5055
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   435
      Left            =   60
      ScaleHeight     =   375
      ScaleWidth      =   6555
      TabIndex        =   21
      Top             =   5340
      Width           =   6615
      Begin VB.CommandButton bProducto 
         Height          =   310
         Left            =   1620
         Picture         =   "frmPresupuestacion.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "Ficha de Producto. [Ctrl+P]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bHistoria 
         Height          =   310
         Left            =   1260
         Picture         =   "frmPresupuestacion.frx":0BDC
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Historia. [Ctrl+H]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   780
         Picture         =   "frmPresupuestacion.frx":14A6
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Cancelar. [Ctrl+C]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bModificar 
         Height          =   310
         Left            =   60
         Picture         =   "frmPresupuestacion.frx":15A8
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Modificar. [Ctrl+M]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bSalir 
         Height          =   310
         Left            =   6180
         Picture         =   "frmPresupuestacion.frx":16F2
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Salir. [Ctrl+X]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   420
         Picture         =   "frmPresupuestacion.frx":17F4
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Grabar [Ctrl+G]."
         Top             =   60
         Width           =   310
      End
   End
   Begin ComctlLib.TabStrip TabValidacion 
      Height          =   2655
      Left            =   60
      TabIndex        =   2
      Top             =   2640
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4683
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
   Begin ComctlLib.ImageList Image1 
      Left            =   6300
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPresupuestacion.frx":18F6
            Key             =   "historia"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPresupuestacion.frx":1C10
            Key             =   "Valido"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPresupuestacion.frx":1F2A
            Key             =   "servicio"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPresupuestacion.frx":2244
            Key             =   "producto"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuAccesos 
      Caption         =   "Accesos"
      Visible         =   0   'False
      Begin VB.Menu MnuAccSeguimiento 
         Caption         =   "Seguimiento de Servicios"
      End
   End
End
Attribute VB_Name = "frmPresupuestacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private aTexto As String

Private Sub bCancelar_Click()
    CargoServicio
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bHistoria_Click()
    If Val(LabProducto.Tag) > 0 Then EjecutarApp pathApp & "\Historia Servicio" & " " & LabProducto.Tag
End Sub

Private Sub bModificar_Click()
    AccionModificar
End Sub

Private Sub bProducto_Click()
    If Val(LabCliente.Tag) > 0 Then EjecutarApp pathApp & "\Productos" & " " & LabCliente.Tag
End Sub

Private Sub bSalir_Click()
    Unload Me
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
    If KeyCode = vbKeyReturn Then Foco tCostoFinal
End Sub
Private Sub cMonedaFinal_LostFocus()
    cMonedaFinal.SelStart = 0
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    
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
    
    bHistoria.Picture = Image1.ListImages("historia").ExtractIcon
    bProducto.Picture = Image1.ListImages("producto").ExtractIcon
    
    Set TabValidacion.ImageList = Image1
    TabValidacion.Tabs(1).Image = Image1.ListImages("historia").Index
    TabValidacion.Tabs(2).Image = Image1.ListImages("Valido").Index
    TabValidacion.Tabs(3).Image = Image1.ListImages("producto").Index
    
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda
    CargoCombo Cons, cMonedaFinal
    
    LimpioCamposFicha
    LimpioCamposPresupuesto
    OcultoCamposPresupuesto
    
    labMotivo.Caption = "&Presupuesto: [F12 a Artículo]": labMotivo.Tag = 1
    MeBotones False, False, False
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyX: Unload Me
            Case vbKeyM: If bModificar.Enabled Then AccionModificar
            Case vbKeyC: If bCancelar.Enabled Then CargoServicio
            Case vbKeyG: If bGrabar.Enabled Then AccionGrabar
            Case vbKeyP: If Val(LabCliente.Tag) > 0 Then EjecutarApp pathApp & "\Productos" & " " & LabCliente.Tag
            Case vbKeyH: If Val(LabProducto.Tag) > 0 Then EjecutarApp pathApp & "\Historia Servicio" & " " & LabProducto.Tag
        End Select
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
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
        .FormatString = "Q|Motivos|>Costo"
        .ColWidth(0) = 250: .ColWidth(1) = 2400
        .Redraw = True
    End With
    With vsHistoria
        .Rows = 1
        .WordWrap = True
        .FormatString = "Fecha|Motivos|Estado|>Importe|"
        .ColWidth(0) = 750: .ColWidth(1) = 3200: .ColWidth(3) = 1300: .ColWidth(5) = 10
        .ColAlignment(0) = flexAlignLeftTop
        .ColAlignment(2) = flexAlignLeftTop
        .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(4) = flexAlignRightTop
        .ColAlignment(5) = flexAlignLeftTop
    End With
    With vsHistoria2
        .Rows = 1
        .WordWrap = True
        .FormatString = "Fecha|Motivos|Estado|>Importe|"
        .ColWidth(0) = 750: .ColWidth(1) = 3400: .ColWidth(3) = 1300: .ColWidth(5) = 10
        .ColAlignment(0) = flexAlignLeftTop
        .ColAlignment(2) = flexAlignLeftTop
        .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(4) = flexAlignRightTop
        .ColAlignment(5) = flexAlignLeftTop
    End With
    With vsProducto
        .Rows = 1
        .Cols = 1
        .ExtendLastCol = True
        .FormatString = "Artículo|Estado|>F.Compra|Garantía|N° Serie|Factura"
        .ColWidth(0) = 3000: .ColWidth(2) = 1000: .ColWidth(4) = 800
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
    If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) = 0 Then Exit Sub
    If vsProducto.Row > 0 Then EjecutarApp App.Path & "\Seguimiento de Servicios", vsProducto.Cell(flexcpData, vsProducto.Row, 4)
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

Private Sub tCostoFinal_GotFocus()
    With tCostoFinal
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tCostoFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentarioP
End Sub
Private Sub tCostoFinal_LostFocus()
    If IsNumeric(tCostoFinal.Text) Then tCostoFinal.Text = Format(tCostoFinal.Text, FormatoMonedaP) Else tCostoFinal.Text = ""
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
                labMotivo.Caption = "&Presupuesto: [F12 a Artículo]": labMotivo.Tag = 1
                tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
            Else
                labMotivo.Caption = "&Artículo: [F12 a Presupuesto]": labMotivo.Tag = 0
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
    MeBotones False, False, False
    Cons = "Select * From Servicio, Producto, Articulo " _
        & " Where SerCodigo = " & Val(tServicio.Text) _
        & " And SerProducto = ProCodigo And ProArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No existe un servicio con ese código.", vbInformation, "ATENCIÓN"
    Else
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
                cMonedaFinal.Text = "": tCostoFinal.Text = ""
                OcultoCamposPresupuesto
            End If
            RsAux.Close
        End If
    End If
    
End Sub
Private Sub CargoDatosServicio()
    
    tServicio.Tag = RsAux!SerCodigo
    labEstadoServicio.Caption = EstadoServicio(RsAux!SerEstadoServicio)
    labEstadoServicio.Tag = RsAux!SerEstadoServicio
    
    labIngreso.Caption = Format(RsAux!SerFecha, FormatoFP)
    labIngreso.Tag = RsAux!SerModificacion
    
    If Not IsNull(RsAux!SerCostoFinal) And Not IsNull(RsAux!SerMoneda) Then tCostoFinal.Text = RsAux!SerCostoFinal: BuscoCodigoEnCombo cMonedaFinal, RsAux!SerMoneda
    
    CargoDatosCliente RsAux!ProCliente
    
    LabProducto.Caption = "(" & Format(RsAux!ProCodigo, "#,000") & ") " & Trim(RsAux!ArtNombre)
    LabProducto.Tag = RsAux!ProCodigo
    LabFCompra.Tag = RsAux!ProFModificacion
    If Not IsNull(RsAux!ProFacturaS) Then labFactura.Caption = Trim(RsAux!ProFacturaS)
    If Not IsNull(RsAux!ProFacturaN) Then labFactura.Caption = Trim(labFactura.Caption) & " " & Trim(RsAux!ProFacturaN)
    If Not IsNull(RsAux!ProCompra) Then LabFCompra.Caption = Format(RsAux!ProCompra, "dd/mm/yyyy")
    
    LabEstado.Caption = EstadoProducto(RsAux!SerEstadoProducto, True)
    LabEstado.Tag = RsAux!SerEstadoProducto
    LabGarantia.Caption = RetornoGarantia(RsAux!ArtID)
    
    If Not IsNull(RsAux!SerComentario) Then tComentarioS.Text = Trim(RsAux!SerComentario)
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
End Sub
Private Sub LimpioCamposPresupuesto()
    
    InicializoGrilla
    LabFAceptado.BackColor = Inactivo: LabFAceptado.ForeColor = vbBlack
    LabFAceptado.Caption = "": LabFAceptado.Tag = ""
    labTitAceptado.Caption = "Aceptado:"
    labFPresupuesto.Caption = ""
    labFReparado.Caption = ""
    labTecnico.Caption = ""
    tPresupuesto.Text = ""
    tCantidad.Text = ""
    cMoneda.Text = ""
    tCosto.Text = ""
    cMonedaFinal.Text = ""
    tCostoFinal.Text = ""
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
        Select Case RsCli!CliTipo
        
            Case TipoCliente.Cliente
                If Not IsNull(RsCli!CliCiRuc) Then LabCliente.Caption = "(" & clsGeneral.RetornoFormatoCedula(RsCli!CliCiRuc) & ")"
                LabCliente.Caption = LabCliente.Caption & " " & Trim(Trim(Format(RsCli!CPeNombre1, "#")) & " " & Trim(Format(RsCli!CPeNombre2, "#"))) & ", " & Trim(Trim(Format(RsCli!CPeApellido1, "#")) & " " & Trim(Format(RsCli!CPeApellido2, "#")))
            Case TipoCliente.Empresa
                If Not IsNull(RsCli!CliCiRuc) Then LabCliente.Caption = "(" & Trim(RsCli!CliCiRuc) & ")"
                If Not IsNull(RsCli!CEmNombre) Then LabCliente.Caption = LabCliente.Caption & " " & Trim(RsCli!CEmFantasia)
                If Not IsNull(RsCli!CEmFantasia) Then LabCliente.Caption = LabCliente.Caption & " (" & Trim(RsCli!CEmFantasia) & ")"
        End Select
        LabCliente.Tag = RsCli!CliCodigo
    End If
    RsCli.Close
    labTelefono.Caption = TelefonoATexto(idCliente)     'Telefonos
End Sub
Private Sub OcultoCamposPresupuesto()
    vsMotivo.BackColor = Inactivo
    tPresupuesto.Enabled = False: tPresupuesto.BackColor = Inactivo
    tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tCosto.Enabled = False: tCosto.BackColor = Inactivo
    cMonedaFinal.Enabled = False: cMonedaFinal.BackColor = Inactivo
    tCostoFinal.Enabled = False: tCostoFinal.BackColor = Inactivo
    tComentarioP.Enabled = False: tComentarioP.BackColor = Inactivo
End Sub
Private Sub MuestroCamposPresupuesto()
    vsMotivo.BackColor = Blanco
    tPresupuesto.Enabled = True: tPresupuesto.BackColor = Blanco
    tCantidad.Enabled = True: tCantidad.BackColor = Blanco
    cMoneda.Enabled = True: cMoneda.BackColor = Blanco
    tCosto.Enabled = True: tCosto.BackColor = Blanco
    cMonedaFinal.Enabled = True: cMonedaFinal.BackColor = Blanco
    tCostoFinal.Enabled = True: tCostoFinal.BackColor = Blanco
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
        If Not RsAux!TalAceptado Then LabFAceptado.BackColor = rojo: LabFAceptado.ForeColor = vbWhite: labTitAceptado.Caption = "No Aceptado:"
    End If
    If Not IsNull(RsAux!TalMonedaCosto) Then BuscoCodigoEnCombo cMoneda, RsAux!TalMonedaCosto
    If Not IsNull(RsAux!TalCostoTecnico) Then tCosto.Text = Format(RsAux!TalCostoTecnico, FormatoMonedaP)
    If Not IsNull(RsAux!TalComentario) Then tComentarioP.Text = Trim(RsAux!TalComentario)
    If Not IsNull(RsAux!TalTecnico) Then labTecnico.Caption = BuscoUsuario(RsAux!TalTecnico, True)
    If Not IsNull(RsAux!TalFReparado) Then labFReparado.Caption = Format(RsAux!TalFReparado, FormatoFP)
    If Not IsNull(RsAux!TalFPresupuesto) Then MeBotones True, False, False
    CargoRenglones
End Sub
Private Sub GraboIngresoDeTraslado(IdServicio As Long, IdUsuario As Integer)
Dim Msg As String
Dim RsSer As rdoResultset
    
    FechaDelServidor
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRB
    
    Cons = "Select * From Servicio Where SerCodigo = " & IdServicio
    Set RsSer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsSer.EOF Then
        Msg = "Otra terminal elimino el servicio."
        RsSer.Close: RsSer.Edit 'Provoco error.
    Else
        If RsSer!SerModificacion = CDate(labIngreso.Tag) Then
            RsSer.Edit
            RsSer!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
            RsSer.Update
            RsSer.Close
        Else
            Msg = "Otra terminal modifico el servicio."
            RsSer.Close: RsSer.Edit 'Provoco error.
        End If
    End If
    
    Cons = "Update Taller Set TalModificacion = '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
        & ", TalFIngresoRecepcion = '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
        & ", TalUsuario = " & IdUsuario _
        & " Where TalServicio = " & IdServicio
    cBase.Execute (Cons)
    
    cBase.CommitTrans
    MeBotones True, False, False
    Screen.MousePointer = 0
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrRB:
    Resume ErrVA
ErrVA:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información." & Chr(13) & Msg, Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub MeBotones(Modif As Boolean, Grabo As Boolean, Cance As Boolean)
    bModificar.Enabled = Modif
    bGrabar.Enabled = Grabo
    bCancelar.Enabled = Cance
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
        & " Where PreNombre Like '" & tPresupuesto.Text & "%' And PreEsPresupuesto = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un presupuesto con ese nombre.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            aValor = RsAux!ID
            AgregoMotivo aValor, True
            tPresupuesto.Text = ""
        Else
            RsAux.Close
            Dim objLista As New clsListadeAyuda
            objLista.ActivoListaAyuda Cons, False, txtConexion, 4000
            aValor = objLista.ValorSeleccionado
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
        & " Where ArtNombre Like '" & tPresupuesto.Text & "%'"
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
            objLista.ActivoListaAyuda Cons, False, txtConexion, 4000
            aValor = objLista.ValorSeleccionado
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
                    If Not IsNull(RsAux!PViPrecio) Then .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PViPrecio, FormatoMonedaP) Else .Cell(flexcpText, .Rows - 1, 2) = "0.00"
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
                    If Not IsNull(RsAux!PViPrecio) Then .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PViPrecio, FormatoMonedaP) Else .Cell(flexcpText, .Rows - 1, 2) = "0.00"
                End If
            End With
        End If
        RsAux.Close
        tPresupuesto.Text = "": tCantidad.Text = "": Foco tPresupuesto
    End If
    
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

Private Sub AccionGrabar()

    If Not ValidoDatos Then Exit Sub
    If MsgBox("¿Confirma validar el presupuesto?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    GraboFicha

End Sub

Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    If Not clsGeneral.TextoValido(tComentarioP.Text) Then MsgBox "Ingreso alguna comilla simple en el comentario del servicio.", vbExclamation, "ATENCIÓN": Foco tComentarioP: Exit Function
    If tCosto.Text <> "" Then
        If Not IsNumeric(tCosto.Text) Then MsgBox "El costo debe ser numérico.", vbExclamation, "ATENCIÓN": Foco tCosto: Exit Function
        If cMoneda.ListIndex = -1 Then MsgBox "Seleccione una moneda.", vbExclamation, "ATENCIÓN": Foco cMoneda: Exit Function
    End If
    If tCostoFinal.Text <> "" Then
        If Not IsNumeric(tCostoFinal.Text) Then MsgBox "El costo final debe ser numérico.", vbExclamation, "ATENCIÓN": Foco tCostoFinal: Exit Function
        If cMonedaFinal.ListIndex = -1 Then MsgBox "Seleccione una moneda.", vbExclamation, "ATENCIÓN": Foco cMonedaFinal: Exit Function
    End If
    ValidoDatos = True
End Function

Private Sub AccionModificar()
    
    tServicio.Text = tServicio.Tag
    CargoServicio
    
    If tServicio.Text <> tServicio.Tag Or tServicio.Tag = "" Then Exit Sub
    MeBotones False, True, True
    
    MuestroCamposPresupuesto
    tServicio.Enabled = False
    
    If tCostoFinal.Text = "" Then
        BuscoCodigoEnCombo cMonedaFinal, paMonedaPesos
        tCostoFinal.Text = "0.00"
        If Val(LabEstado.Tag) = EstadoP.FueraGarantia Then
            'Le doy un costo en base a los precios de la grilla.
            With vsMotivo
                For I = 1 To .Rows - 1
                    If IsNumeric(.Cell(flexcpText, I, 2)) Then tCostoFinal.Text = CCur(tCostoFinal.Text) + CCur(.Cell(flexcpText, I, 2))
                Next I
                tCostoFinal.Text = Format(tCostoFinal.Text, FormatoMonedaP)
            End With
        End If
    End If
    If labEstadoServicio.Tag = EstadoS.Cumplido Then
        OcultoCamposPresupuesto
        cMoneda.Enabled = True: cMoneda.BackColor = vbWhite
        tCosto.Enabled = True: tCosto.BackColor = vbWhite
        tComentarioP.Enabled = True: tComentarioP.BackColor = vbWhite
        
    Else
        Foco tPresupuesto
    End If
    
End Sub

Private Sub GraboFicha()
Dim IdLocalReparacion As Long
Dim RsArticulo As rdoResultset
Dim Aceptado As Boolean: Aceptado = False
    
    Screen.MousePointer = 11
    If LabEstado.Tag = EstadoP.SinCargo And LabFAceptado.Caption = "" Then
        If MsgBox("El producto está sin cargo. ¿Desea dar como aceptado el presupuesto?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then Aceptado = True
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
                RsAux!SerCostoFinal = 0
            Else
                RsAux!SerMoneda = cMonedaFinal.ItemData(cMonedaFinal.ListIndex)
                RsAux!SerCostoFinal = CCur(tCostoFinal.Text)
            End If
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
            If Aceptado Then
                RsAux!TalFAceptacion = Format(gFechaServidor, sqlFormatoFH)
                RsAux!TalAceptado = 1
            End If
            RsAux!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
            If cMoneda.ListIndex = -1 Then
                RsAux!TalMonedaCosto = paMonedaPesos
                RsAux!TalCostoTecnico = 0
            Else
                RsAux!TalMonedaCosto = cMoneda.ItemData(cMoneda.ListIndex)
                RsAux!TalCostoTecnico = CCur(tCosto.Text)
            End If
            If Trim(tComentarioP.Text) = "" Then RsAux!TalComentario = Null Else RsAux!TalComentario = Trim(tComentarioP.Text)
            RsAux.Update: RsAux.Close
        Else
            RsAux.Close: cBase.RollbackTrans
            MsgBox "No podrá almacenar los datos debido a que otra terminal modifico los datos, verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    If labFReparado.Caption <> "" Then
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
                            ElseIf RsAux!SReCantidad < CCur(.Cell(flexcpText, I, 0)) Then
                                'Doy baja en los movimientos físicos y baja en el local.
                                MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, IdLocalReparacion, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)) - RsAux!SReCantidad, paEstadoArticuloEntrega, -1, TipoDocumento.Servicio, tServicio.Tag
                                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, IdLocalReparacion, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)) - RsAux!SReCantidad, paEstadoArticuloEntrega, -1
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
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción, reintente.", Err.Description
    Screen.MousePointer = 0: Exit Sub
ErrVA:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub

Private Sub vsMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsMotivo.BackColor = Inactivo Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn: Foco cMoneda
        Case vbKeyDelete: If vsMotivo.Row > 0 Then vsMotivo.RemoveItem vsMotivo.Row
    End Select
End Sub

Private Sub CargoRenglones()
Dim RsSR As rdoResultset
Dim aValor As Long
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
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsSR!SReTotal, FormatoMonedaP)
        End With
        RsSR.MoveNext
    Loop
    RsSR.Close
End Sub

Private Sub AccionReparar()
Dim RsSR As rdoResultset

    Screen.MousePointer = 11
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
            RsAux.Edit
            RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
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
            RsAux!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
            RsAux!TalFReparado = Format(gFechaServidor, sqlFormatoFH)
            RsAux.Update: RsAux.Close
        Else
            RsAux.Close: cBase.RollbackTrans
            MsgBox "No podrá almacenar los datos debido a que otra terminal modifico los datos, verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    With vsMotivo
        For I = 1 To .Rows - 1
            If CLng(.Cell(flexcpData, I, 1)) <> paTipoArticuloServicio Then
                MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, -1, TipoDocumento.Servicio, tServicio.Tag
                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, -1
            End If
        Next I
    End With
    cBase.CommitTrans
    On Error Resume Next
    Screen.MousePointer = 0
    CargoServicio
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción, reintente.", Err.Description
    Screen.MousePointer = 0: Exit Sub
ErrVA:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Sub
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
    Cons = "Select * From Producto, Articulo Where ProCliente = " & LabCliente.Tag _
        & " And ProCodigo <> " & LabProducto.Tag _
        & " And ProArticulo = ArtID"
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
        Case vbKeyAdd: If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) > 0 Then EjecutarApp App.Path & "\Seguimiento de Servicios", vsProducto.Cell(flexcpData, vsProducto.Row, 4)
    End Select
End Sub

Private Sub vsProducto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And vsProducto.Rows > 1 Then PopupMenu MnuAccesos
End Sub

Private Sub vsProducto_RowColChange()
    If vsProducto.Row > 0 Then CargoHistoria Val(vsProducto.Cell(flexcpData, vsProducto.Row, 0)), vsHistoria2
End Sub
