VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmDeFactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Factura"
   ClientHeight    =   5610
   ClientLeft      =   2550
   ClientTop       =   2730
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDeFactura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7965
   Begin VB.PictureBox Picture1 
      Height          =   3555
      Index           =   1
      Left            =   4500
      ScaleHeight     =   3495
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   2220
      Width           =   1575
      Begin VSFlex6DAOCtl.vsFlexGrid lEnvio 
         Height          =   1215
         Left            =   60
         TabIndex        =   33
         Top             =   60
         Width           =   7455
         _ExtentX        =   13150
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   10
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
      Begin VB.Label lECosto 
         BackStyle       =   0  'Transparent
         Caption         =   "Costo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lEEntregado 
         BackStyle       =   0  'Transparent
         Caption         =   "Agencia"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   5040
         TabIndex        =   31
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Entregado:"
         Height          =   255
         Left            =   4155
         TabIndex        =   30
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lEFPago 
         BackStyle       =   0  'Transparent
         Caption         =   "Pago"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2760
         TabIndex        =   29
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pago:"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lEUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Carlitos"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lEVaCon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comentarios:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6240
         TabIndex        =   25
         Top             =   2565
         Width           =   960
      End
      Begin VB.Label lEComentario 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentarios"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2565
         Width           =   6015
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   4155
         TabIndex        =   23
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lEAgencia 
         BackStyle       =   0  'Transparent
         Caption         =   "Agencia"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   5040
         TabIndex        =   22
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Piso:"
         Height          =   255
         Left            =   4155
         TabIndex        =   21
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lEPiso 
         BackStyle       =   0  'Transparent
         Caption         =   "Piso"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Costo:"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Flete:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lEFlete 
         BackStyle       =   0  'Transparent
         Caption         =   "Flete"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Camión:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lECamion 
         BackStyle       =   0  'Transparent
         Caption         =   "Camión"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrega:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lEFecha 
         BackStyle       =   0  'Transparent
         Caption         =   "Lun 14-Ene-1998  0000-9000"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         Height          =   255
         Left            =   4155
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lEEstado 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5040
         TabIndex        =   10
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lEDireccion 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   1320
         Width           =   6375
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   1515
         Left            =   60
         Top             =   1310
         Width           =   7455
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Index           =   3
      Left            =   3420
      ScaleHeight     =   3075
      ScaleWidth      =   555
      TabIndex        =   53
      Top             =   2940
      Width           =   615
      Begin VSFlex6DAOCtl.vsFlexGrid vsComentario 
         Height          =   1515
         Left            =   60
         TabIndex        =   54
         Top             =   60
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   2672
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
         Rows            =   10
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
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Index           =   2
      Left            =   2400
      ScaleHeight     =   3195
      ScaleWidth      =   555
      TabIndex        =   50
      Top             =   2940
      Width           =   615
      Begin VSFlex6DAOCtl.vsFlexGrid lSuceso 
         Height          =   1515
         Left            =   60
         TabIndex        =   51
         Top             =   60
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   2672
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
         Rows            =   10
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Documento"
      ForeColor       =   &H00000080&
      Height          =   1875
      Left            =   120
      TabIndex        =   35
      Top             =   360
      Width           =   7695
      Begin VB.TextBox tNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   1
         Top             =   250
         Width           =   915
      End
      Begin VB.TextBox tFRetiro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5880
         TabIndex        =   4
         Top             =   1500
         Width           =   855
      End
      Begin VB.CommandButton bGrabar 
         Height          =   320
         Left            =   6840
         Picture         =   "frmDeFactura.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1485
         Width           =   340
      End
      Begin VB.CommandButton bModificar 
         Height          =   320
         Left            =   6840
         Picture         =   "frmDeFactura.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1485
         Width           =   340
      End
      Begin VB.CommandButton bCancelar 
         Height          =   320
         Left            =   7200
         Picture         =   "frmDeFactura.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1485
         Width           =   340
      End
      Begin VB.CommandButton bEliminar 
         Height          =   320
         Left            =   7200
         Picture         =   "frmDeFactura.frx":0610
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1485
         Width           =   340
      End
      Begin AACombo99.AACombo cPendiente 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1500
         Width           =   3735
         _ExtentX        =   6588
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Docu&mento:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lDoc 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   2160
         TabIndex        =   56
         Top             =   255
         Width           =   5415
      End
      Begin VB.Label lAnulada 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "DOCUMENTO ANULADO !!"
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
         Height          =   225
         Left            =   4920
         TabIndex        =   55
         Top             =   1230
         Width           =   2655
      End
      Begin VB.Label lComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "14-Ene-1998"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   1140
         TabIndex        =   52
         Top             =   960
         Width           =   6435
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Emisión:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Usr:"
         Height          =   255
         Left            =   6285
         TabIndex        =   48
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
         Height          =   255
         Left            =   2580
         TabIndex        =   46
         Top             =   660
         Width           =   735
      End
      Begin VB.Label lFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "14-Ene-1998"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   1140
         TabIndex        =   45
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "USR"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   6660
         TabIndex        =   44
         Top             =   660
         Width           =   915
      End
      Begin VB.Label lImporte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "$"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   3240
         TabIndex        =   43
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label lDFRetira 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Retira desde 00/00/00 al:"
         Height          =   195
         Left            =   3975
         TabIndex        =   3
         Top             =   1530
         Width           =   1845
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento Pendiente"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label lVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "0"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   5460
         TabIndex        =   41
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Vend.:"
         Height          =   255
         Left            =   4920
         TabIndex        =   40
         Top             =   660
         Width           =   555
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2715
      Index           =   0
      Left            =   180
      ScaleHeight     =   2655
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   2940
      Width           =   1275
      Begin VSFlex6DAOCtl.vsFlexGrid lArticulo 
         Height          =   1515
         Left            =   60
         TabIndex        =   32
         Top             =   60
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2672
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
         Rows            =   10
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
   End
   Begin ComctlLib.TabStrip Tab1 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1085
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "A &Retirar"
            Key             =   "ARetirar"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lTitular 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label17"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   34
      Top             =   30
      Width           =   7695
   End
   Begin VB.Menu MnuARetirar 
      Caption         =   "&Detalles"
      Begin VB.Menu MnuVerTitulo 
         Caption         =   "Menú Documento"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuVerL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerDetalle 
         Caption         =   "Detalle de la &Operación"
      End
      Begin VB.Menu MnuVerPago 
         Caption         =   "Detalle de &Pagos"
      End
      Begin VB.Menu MnuVerDeuda 
         Caption         =   "Deuda en &Cheques"
      End
      Begin VB.Menu MnuVerL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerEntrega 
         Caption         =   "&Hora y Quién entregó"
      End
      Begin VB.Menu MnuInstalacion 
         Caption         =   "Ver Instalación"
      End
      Begin VB.Menu MnuVerOperaciones 
         Caption         =   "&Visualización de Operaciones"
      End
      Begin VB.Menu MnuInComentarios 
         Caption         =   "Ingresar Comentarios"
      End
      Begin VB.Menu MnuVerEnvio 
         Caption         =   "Ver Envío de Mercadería"
      End
   End
   Begin VB.Menu MnuEnvio 
      Caption         =   "&Envíos"
      Visible         =   0   'False
      Begin VB.Menu MnuIrTitulo 
         Caption         =   "Menú Envíos"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuIrL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIrAEnvio 
         Caption         =   "Visualizar &Envío"
      End
      Begin VB.Menu MnuIrACamion 
         Caption         =   "Datos del &Camión"
      End
   End
   Begin VB.Menu MnuModificar 
      Caption         =   "&Modificar"
      Begin VB.Menu MnuMoComentarios 
         Caption         =   "&Comentarios"
      End
   End
   Begin VB.Menu MnuDocs 
      Caption         =   "Docs. &Relacionados"
      Begin VB.Menu MnuD 
         Caption         =   "idDoc"
         Index           =   0
      End
   End
   Begin VB.Menu MnuVolver 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalir 
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmDeFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsXX As rdoResultset

Dim prmIdDocumento As Long   'Id del documento seleccionado
Dim prmTipoD As Integer          'Tipo de Documento Seleccionado
Dim prmIdCliente As Long         'Id del cliente del documento
Dim prmIdFactura As Long        'id de la factura asociada al documento (caso de notas o recibos de pagos)

Dim Fletes As String
Dim aTexto As String

Private Const prmSucesoVarios = 99

Private Sub bCancelar_Click()
    
    If cPendiente.BackColor = Inactivo Then Exit Sub
    
    Screen.MousePointer = 11
    On Error Resume Next
    bModificar.ZOrder 0
    bEliminar.ZOrder 0
        
    cPendiente.Enabled = False
    tFRetiro.Enabled = False
    cPendiente.BackColor = Inactivo
    tFRetiro.BackColor = Inactivo
    BuscoCodigoEnCombo cPendiente, CLng(cPendiente.Tag)
    tFRetiro.Text = tFRetiro.Tag
    Screen.MousePointer = 0
    
End Sub

Private Sub bEliminar_Click()

    If MsgBox("Confirma eliminar el estado pendiente al documento.", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Exit Sub
    
    On Error GoTo errGrabar
    Screen.MousePointer = 11
    cons = "Update Documento Set DocPendiente = NULL Where DocCodigo = " & prmIdDocumento
    cBase.Execute cons
    cons = "Update Envio Set EnvHabilitado = 1 Where EnvDocumento = " & prmIdDocumento
    cBase.Execute cons
    bEliminar.Enabled = False
    cPendiente.Text = ""
    cPendiente.Tag = ""
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al eliminar."
End Sub

Private Sub bGrabar_Click()
    
    If cPendiente.BackColor = Inactivo Then Exit Sub
        
    'Valido los datos para grabar----------------------------------------------------------------------------
    'If cPendiente.ListIndex = -1 Then
    '    MsgBox "Debe ingresar el código de pendiente.", vbExclamation, "ATENCIÓN"
    '    Foco cPendiente: Exit Sub
    'End If
    
    If Not IsDate(tFRetiro.Text) Then
        MsgBox "La fecha de retiro ingresada no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tFRetiro: Exit Sub
    Else
        If CDate(tFRetiro.Text) < Date Then
            MsgBox "La fecha de retiro no debe ser menor al día de hoy.", vbExclamation, "ATENCIÓN"
            Foco tFRetiro: Exit Sub
        End If
    End If
    
    If IsDate(lDFRetira.Tag) Then
        
        If CDate(tFRetiro.Text) < CDate(lDFRetira.Tag) Then
            MsgBox "La fecha de retiro ingresada no debe ser menor a la fecha desde.", vbExclamation, "ATENCIÓN"
            Foco tFRetiro: Exit Sub
        End If

        If DateDiff("d", CDate(lDFRetira.Tag), CDate(tFRetiro.Text)) > 59 Then
            MsgBox "La diferencia entre las fechas de retiro no debe ser mayor a 60 días.", vbExclamation, "ATENCIÓN"
            Foco tFRetiro: Exit Sub
        End If
        
    End If
    '---------------------------------------------------------------------------------------------------------------
    
    If MsgBox("Confirma grabar los valores ingresados al documento.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    On Error GoTo errGrabar
    Screen.MousePointer = 11
    
    Dim pdatRetira As Date
    pdatRetira = CDate(tFRetiro.Text)
    
    If IsDate(lDFRetira.Tag) Then
        pdatRetira = Format(pdatRetira, "yyyy/mm/dd") & " 01:" & DateDiff("d", CDate(lDFRetira.Tag), pdatRetira)
    End If
        
    If cPendiente.ListIndex <> -1 Then
        cons = "Update Documento Set DocPendiente = " & cPendiente.ItemData(cPendiente.ListIndex) _
                & ", DocFRetira = '" & Format(pdatRetira, sqlFormatoFH) & "'" _
                & " Where DocCodigo = " & prmIdDocumento

    Else
         cons = "Update Documento Set DocFRetira = '" & Format(pdatRetira, sqlFormatoFH) & "'" _
                & " Where DocCodigo = " & prmIdDocumento
    End If
    cBase.Execute cons
    
    bModificar.ZOrder 0: bEliminar.ZOrder 0
        
    If cPendiente.ListIndex <> -1 Then cPendiente.Tag = cPendiente.ItemData(cPendiente.ListIndex)
    tFRetiro.Text = Format(tFRetiro.Text, "d-Mmm yy"): tFRetiro.Tag = tFRetiro.Text
    
    cPendiente.Enabled = False: cPendiente.BackColor = Inactivo
    tFRetiro.Enabled = False: tFRetiro.BackColor = Inactivo
    
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    clsGeneral.OcurrioError "Ocurrió un error al grabar.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bModificar_Click()

    bModificar.ZOrder 1
    bEliminar.ZOrder 1
    
    cPendiente.Enabled = True
    tFRetiro.Enabled = True
    cPendiente.BackColor = Blanco
    tFRetiro.BackColor = Blanco
    
    Foco cPendiente
    
End Sub

Private Sub cPendiente_GotFocus()
    cPendiente.SelStart = 0: cPendiente.SelLength = Len(cPendiente.Text)
End Sub

Private Sub cPendiente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFRetiro
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape: Unload Me
    End Select
    
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    MnuDocs.Enabled = False
    'Me.Left = (Screen.Width - Me.Width) / 2: Me.Top = (Screen.Height - Me.Height) / 2
    ObtengoSeteoForm Me
    If Me.Height <> 6315 Then Me.Height = 6315
    
    Fletes = CargoArticulosDeFlete
       
    InicializoGrillas
    LimpioFichaEnvio
    LimpioFichaDocumento
    
    'SETEO LOS OBJETOS EN PANTALLA----------------------------------------------------------
    Tab1.Left = 120
    Tab1.Height = Me.ScaleHeight - Tab1.Top - 100 '3300
    Tab1.Width = 7695
    For I = 0 To Picture1.Count - 1
        Picture1(I).Top = Tab1.Top + 350
        Picture1(I).Left = Tab1.Left + 40
        Picture1(I).Height = Tab1.Height + 40 - 440
        Picture1(I).Width = Tab1.Width - 170
        Picture1(I).BorderStyle = 0
    Next
    lArticulo.Top = 60: lArticulo.Left = 60
    lArticulo.Width = Picture1(0).Width - 60
    lArticulo.Height = Picture1(0).Height - 100
    lSuceso.Top = lArticulo.Top: lSuceso.Width = lArticulo.Width: lSuceso.Height = lArticulo.Height: lSuceso.Left = lArticulo.Left
    vsComentario.Top = lArticulo.Top: vsComentario.Width = lArticulo.Width: vsComentario.Height = lArticulo.Height: vsComentario.Left = lArticulo.Left
    Picture1(0).ZOrder 0
    '----------------------------------------------------------------------------------------------------
    
    'Cargo los codigos de pendiente----------------------------------------------------------------
    cons = "Select PEnCodigo, PEnNombre From PendienteEntrega Order by PEnNombre"
    CargoCombo cons, cPendiente, ""
    '----------------------------------------------------------------------------------------------------
    
    prmIdDocumento = 0
    'Linea de Comandos    -------------------------------------
    If Trim(Command()) <> "" Then
        aTexto = Trim(Command())
    End If
    '---------------------------------------------------------------
    
    If Val(aTexto) <> 0 Then
        Screen.MousePointer = 11
        CargoDocumento Val(aTexto)
        Screen.MousePointer = 0
    End If
    
End Sub

Private Sub TabsEnvios(Documento As Long)
    
    On Error GoTo errCargar
'    cons = "Select EnvCodigo From Envio" & _
           " Where EnvTipo = " & TipoEnvio.Entrega & _
           " And (EnvDocumento = " & Documento & _
             " OR EnvDocumento IN (SELECT RDoRemito FROM RemitoDocumento Where RDoDocumento = " & Documento & ") )"
           
'    cons = "Select EnvCodigo From Envio" _
           & " Where EnvTipo = " & TipoEnvio.Entrega _
           & " And EnvDocumento = " & Documento
           
    cons = "Select EnvCodigo From Envio" & _
           " Where EnvTipo = " & TipoEnvio.Entrega & _
           " And EnvDocumento = " & Documento & _
                " UNION ALL " & _
            "Select EnvCodigo From Envio" & _
           " Where EnvTipo = " & TipoEnvio.Entrega & _
           " And EnvDocumento IN (SELECT RDoRemito FROM RemitoDocumento Where RDoDocumento = " & Documento & ")"
           
                     
    Set RsXX = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsXX.EOF
        Tab1.Tabs.Add pvkey:="A" & CStr(RsXX!EnvCodigo), pvCaption:="Envío (" & CStr(RsXX!EnvCodigo) & ")"
        RsXX.MoveNext
    Loop
    RsXX.Close
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los envíos", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub CargoDatosEnvio(Codigo As Long)
    
    On Error GoTo errCargar
    LimpioFichaEnvio
    cons = "Select Envio.*, AgeNombre, CamCodigo, CamNombre, TFlDescripcion, MonSigno " _
           & " From Envio, Camion, Agencia, TipoFlete, Moneda" _
           & " Where EnvCodigo = " & Codigo _
           & " And EnvCamion *= CamCodigo " _
           & " And EnvAgencia *= AgeCodigo " _
           & " And EnvTipoFlete *= TFlCodigo " _
           & " And EnvMoneda = MonCodigo"
           
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsAux.EOF Then
        lEDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, rsAux!EnvDireccion, Localidad:=True, Zona:=True)
        
        If Not IsNull(rsAux!EnvFechaPrometida) Then
            lEFecha.Caption = Format(rsAux!EnvFechaPrometida, "Ddd d Mmm-YY")
            If Not IsNull(rsAux!EnvRangoHora) Then lEFecha.Caption = lEFecha.Caption & "   " & Trim(rsAux!EnvRangoHora)
        End If
        
        If Not IsNull(rsAux!EnvMoneda) Then lECosto.Caption = Trim(rsAux!MonSigno)
        If Not IsNull(rsAux!EnvValorFlete) Then lECosto.Caption = lECosto.Caption & " " & Format(rsAux!EnvValorFlete, FormatoMonedaP)
        
        If Not IsNull(rsAux!EnvValorPiso) Then lEPiso.Caption = Format(rsAux!EnvValorPiso, FormatoMonedaP)
        
        If Not IsNull(rsAux!EnvComentario) Then lEComentario.Caption = Format(rsAux!EnvComentario, FormatoMonedaP)
        
        If Not IsNull(rsAux!EnvEstado) Then lEEstado.Caption = RetornoEstadoEnvio(rsAux!EnvEstado)
        If Not IsNull(rsAux!EnvFechaEntregado) Then lEEntregado.Caption = Format(rsAux!EnvFechaEntregado, "Ddd d-Mmm yy")
        
        lECamion.Tag = 0
        If Not IsNull(rsAux!CamNombre) Then
            lECamion.Caption = Trim(rsAux!CamNombre)
            lECamion.Tag = rsAux!CamCodigo
        End If
        
        If Not IsNull(rsAux!AgeNombre) Then lEAgencia.Caption = Trim(rsAux!AgeNombre)
        If Not IsNull(rsAux!TFlDescripcion) Then lEFlete.Caption = Trim(rsAux!TFlDescripcion)
        
        If Not IsNull(rsAux!EnvFormaPago) Then lEFPago.Caption = RetornoPagoEnvio(rsAux!EnvFormaPago)
        
        If Not IsNull(rsAux!EnvUsuario) Then lEUsuario.Caption = z_BuscoUsuario(rsAux!EnvUsuario, Identificacion:=True)
        
            
        If Not IsNull(rsAux!EnvVaCon) Then      'Cargo campos del va Con
            lEVaCon.Caption = "Va con "
            If rsAux!EnvVaCon > 0 Then
                lEVaCon.Caption = lEVaCon.Caption & "Nº " & rsAux!EnvVaCon
            Else
                Dim rs2 As rdoResultset
                cons = "Select isNull(Count(*), 0) From Envio Where EnvVaCon = " & Codigo
                Set rs2 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                
                If Not rs2.EOF Then
                    If rs2(0) = 1 Then
                        rs2.Close
                        cons = "Select * From Envio Where EnvVaCon = " & Codigo
                        Set rs2 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                        lEVaCon.Caption = lEVaCon.Caption & "Nº " & rsAux!EnvCodigo
                    Else
                        lEVaCon.Caption = lEVaCon.Caption & rs2(0) & " envíos."
                    End If
                End If
                rs2.Close
                
            End If
            
        End If
    End If
    rsAux.Close
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar el envío seleccionado", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoArticulosEnvio(Envio As Long)

    On Error GoTo errCargar
    lEnvio.Rows = 1: lEnvio.Refresh
    Dim pstrArticulo As String
    'cons = "Select * From RenglonEnvio, Articulo" _
           & " Where REvEnvio = " & Envio _
           & " And REvArticulo = ArtID "
    
    cons = "Select RenglonEnvio.*, ArtCodigo, ArtNombre, CodTexto Detalle " & _
            " From Envio" & _
                " INNER JOIN RenglonEnvio ON REvEnvio = EnvCodigo " & _
                " INNER JOIN Articulo ON REvArticulo = ArtId " & _
                " LEFT OUTER JOIN ArticuloEspecifico ON EnvDocumento = AEsDocumento And AEsTipoDocumento = 1 And REvArticulo = AEsArticulo" & _
                " LEFT OUTER JOIN Codigos ON AEsTipo = CodId AND CodCual = 128" & _
            " Where REvEnvio = " & Envio
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsAux.EOF
        With lEnvio
        
            pstrArticulo = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
            If Not IsNull(rsAux!Detalle) Then pstrArticulo = pstrArticulo & " (" & Trim(rsAux!Detalle) & ")"
                
            .AddItem pstrArticulo

            .Cell(flexcpText, .Rows - 1, 1) = rsAux!REvCantidad
            .Cell(flexcpText, .Rows - 1, 2) = rsAux!REvAEntregar
            If rsAux!REvAEntregar <> 0 Then
                .Cell(flexcpBackColor, .Rows - 1, 2) = Colores.RojoClaro
                .Cell(flexcpForeColor, .Rows - 1, 2) = Colores.Blanco
                .Cell(flexcpFontBold, .Rows - 1, 2) = True
            End If
            If Not IsNull(rsAux!REvComentario) Then .Cell(flexcpText, .Rows - 1, 0) = Trim(.Cell(flexcpText, .Rows - 1, 0)) & " - " & Trim(rsAux!REvComentario)
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos del envío seleccionado", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoArticulosARetirar(Documento As Long)
    
Dim aCantidad As Long

    On Error GoTo errCargar
    Screen.MousePointer = 11
    
    MnuD(0).Visible = True
    For I = 1 To MnuD.UBound
        Unload MnuD(I)
    Next
    
    aCantidad = 0
    
    Dim pstrArticulo As String

    'Cargo los Articulos A Retirar con la Factura--------------------------------------------
'    cons = "Select DocCodigo, DocTipo, DocSerie, DocNumero, Renglon.*, ArtCodigo, ArtNombre From Documento, Renglon, Articulo" _
            & " Where DocCodigo = " & Documento _
            & " And RenDocumento = DocCodigo" _
            & " And RenArticulo = ArtId"
            
    cons = "Select DocCodigo, DocTipo, DocSerie, DocNumero, Renglon.*, ArtCodigo, ArtNombre, IsNull(CodTexto, RenDescripcion) Detalle " & _
            " From Documento" & _
                " INNER JOIN Renglon ON RenDocumento = DocCodigo " & _
                " INNER JOIN Articulo ON RenArticulo = ArtId " & _
                " LEFT OUTER JOIN ArticuloEspecifico ON RenDocumento = AEsDocumento And AEsTipoDocumento = 1 And RenArticulo = AEsArticulo" & _
                " LEFT OUTER JOIN Codigos ON AEsTipo = CodId AND CodCual = 128" & _
            " Where DocCodigo = " & Documento
            
    Set RsXX = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsXX.EOF Then
        AddMenuDoc RsXX!DocCodigo, RetornoNombreDocumento(RsXX!DocTipo) & " " & Trim(RsXX!DocSerie) & "-" & RsXX!DocNumero
    End If
    
    
    Do While Not RsXX.EOF
        'If InStr(Fletes, RsXX!RenArticulo & ",") = 0 Then
            With lArticulo
                pstrArticulo = Format(RsXX!ArtCodigo, "(#,000,000)") & " " & Trim(RsXX!ArtNombre)
                If Not IsNull(RsXX!Detalle) Then pstrArticulo = pstrArticulo & " (" & Trim(RsXX!Detalle) & ")"
                
                .AddItem pstrArticulo

                .Cell(flexcpText, .Rows - 1, 1) = RsXX!RenCantidad
                .Cell(flexcpText, .Rows - 1, 2) = Format(RsXX!RenPrecio, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 3) = "FAC " & Trim(RsXX!DocSerie) & Trim(RsXX!DocNumero)
                .Cell(flexcpText, .Rows - 1, 4) = RsXX!RenCantidad - RsXX!RenARetirar
                .Cell(flexcpText, .Rows - 1, 5) = RsXX!RenARetirar
                If RsXX!RenARetirar <> 0 Then
                    If InStr("," & Fletes, "," & RsXX!RenArticulo & ",") = 0 Then .Cell(flexcpBackColor, .Rows - 1, 5) = Colores.RojoClaro: .Cell(flexcpForeColor, .Rows - 1, 5) = Colores.Blanco: .Cell(flexcpFontBold, .Rows - 1, 5) = True
                End If
            End With
            aCantidad = aCantidad + RsXX!RenARetirar
        'End If
        
        RsXX.MoveNext
    Loop
    RsXX.Close
    '-----------------------------------------------------------------------------------------------
    
    'Cargo los Articulos A Retirar con Remitos------------------------------------------------
    'cons = "Select * From Remito, RenglonRemito, Articulo" _
            & " Where RemDocumento = " & Documento _
            & " And RemCodigo = RReRemito " _
            & " And RReArticulo = ArtId"
            
    cons = "Select RemCodigo, RenglonRemito.*, ArtCodigo, ArtNombre, CodTexto Detalle " & _
            " From Remito" & _
                " INNER JOIN RenglonRemito ON RemCodigo = RReRemito " & _
                " INNER JOIN Articulo ON RReArticulo = ArtId " & _
                " LEFT OUTER JOIN ArticuloEspecifico ON RemDocumento = AEsDocumento And AEsTipoDocumento = 1 And RReArticulo = AEsArticulo" & _
                " LEFT OUTER JOIN Codigos ON AEsTipo = CodId AND CodCual = 128" & _
            " Where RemDocumento = " & Documento
            
    Set RsXX = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsXX.EOF
    
        'If InStr(Fletes, RsXX!RReArticulo & ",") = 0 Then
            With lArticulo
            
                pstrArticulo = Format(RsXX!ArtCodigo, "(#,000,000)") & " " & Trim(RsXX!ArtNombre)
                If Not IsNull(RsXX!Detalle) Then pstrArticulo = pstrArticulo & " (" & Trim(RsXX!Detalle) & ")"
            
                .AddItem pstrArticulo
                    
                .Cell(flexcpText, .Rows - 1, 1) = RsXX!RReCantidad
                '.Cell(flexcpText, .Rows - 1, 2) = Format(RsXX!RenPrecio, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 3) = "REM " & RsXX!RemCodigo
                .Cell(flexcpText, .Rows - 1, 4) = RsXX!RReCantidad - RsXX!RReAEntregar
                .Cell(flexcpText, .Rows - 1, 5) = RsXX!RReAEntregar
                If RsXX!RReAEntregar <> 0 Then .Cell(flexcpBackColor, .Rows - 1, 5) = Colores.RojoClaro: .Cell(flexcpForeColor, .Rows - 1, 5) = Colores.Blanco: .Cell(flexcpFontBold, .Rows - 1, 5) = True
                aCantidad = aCantidad + RsXX!RReAEntregar
            End With
        'End If
        
        RsXX.MoveNext
    Loop
    RsXX.Close
    '-----------------------------------------------------------------------------------------------
    
    'Cargo los Articulos en Nota----------------------------------------------------------------
    'cons = "Select  * From Nota, Documento, Renglon, Articulo" _
           & " Where NotFactura = " & Documento _
           & " And NotNota = DocCodigo " _
           & " And DocCodigo = RenDocumento" _
           & " And RenArticulo = ArtID " _
           & " And DocAnulado = 0"
           
    cons = "Select DocCodigo, DocTipo, DocSerie, DocNumero, Renglon.*, ArtCodigo, ArtNombre, IsNull(CodTexto, RenDescripcion) Detalle " & _
            " From Nota" & _
                " INNER JOIN Documento ON NotNota = DocCodigo " & _
                " INNER JOIN Renglon ON RenDocumento = DocCodigo " & _
                " INNER JOIN Articulo ON RenArticulo = ArtId " & _
                " LEFT OUTER JOIN ArticuloEspecifico ON RenDocumento = AEsDocumento And AEsTipoDocumento = 1 And RenArticulo = AEsArticulo" & _
                " LEFT OUTER JOIN Codigos ON AEsTipo = CodId AND CodCual = 128" & _
            " Where NotFactura = " & Documento & _
            " And DocAnulado = 0"
           
    Set RsXX = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsXX.EOF
        
        AddMenuDoc RsXX!DocCodigo, RetornoNombreDocumento(RsXX!DocTipo) & " " & Trim(RsXX!DocSerie) & "-" & RsXX!DocNumero
        
        With lArticulo
            pstrArticulo = Format(RsXX!ArtCodigo, "(#,000,000)") & " " & Trim(RsXX!ArtNombre)
            If Not IsNull(RsXX!Detalle) Then pstrArticulo = pstrArticulo & " (" & Trim(RsXX!Detalle) & ")"
        
            .AddItem pstrArticulo
                
            .Cell(flexcpText, .Rows - 1, 1) = RsXX!RenCantidad
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsXX!RenPrecio, "#,##0.00")
            Select Case RsXX!DocTipo
                Case TipoDocumento.NotaCredito: .Cell(flexcpText, .Rows - 1, 3) = "NCR " & Trim(RsXX!DocSerie) & Trim(RsXX!DocNumero)
                Case TipoDocumento.NotaDevolucion: .Cell(flexcpText, .Rows - 1, 3) = "NDE " & Trim(RsXX!DocSerie) & Trim(RsXX!DocNumero)
                Case TipoDocumento.NotaEspecial: .Cell(flexcpText, .Rows - 1, 3) = "NES " & Trim(RsXX!DocSerie) & Trim(RsXX!DocNumero)
            End Select
            .Cell(flexcpText, .Rows - 1, 4) = "********"
            .Cell(flexcpText, .Rows - 1, 5) = "******"
        End With
        RsXX.MoveNext
    Loop
    RsXX.Close
    '-----------------------------------------------------------------------------------------------
    
    If aCantidad > 0 Then bModificar.Enabled = True
    
    'Si no Hay Articulos en la lista, cargo los de la factura sin filtrar envios---------------
    If lArticulo.Rows = 1 Then
        'cons = "Select Renglon.*, ArtCodigo, ArtNombre From Renglon, Articulo" _
                & " Where RenDocumento = " & Documento _
                & " And RenArticulo = ArtId"
                
        cons = "Select Renglon.*, ArtCodigo, ArtNombre, IsNull(CodTexto, RenDescripcion) Detalle " & _
                " From Renglon" & _
                    " INNER JOIN Articulo ON RenArticulo = ArtId " & _
                    " LEFT OUTER JOIN ArticuloEspecifico ON RenDocumento = AEsDocumento And AEsTipoDocumento = 1 And RenArticulo = AEsArticulo" & _
                    " LEFT OUTER JOIN Codigos ON AEsTipo = CodId AND CodCual = 128" & _
                " Where RenDocumento = " & Documento
                
        Set RsXX = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsXX.EOF
        
            With lArticulo
                pstrArticulo = Format(RsXX!ArtCodigo, "(#,000,000)") & " " & Trim(RsXX!ArtNombre)
                If Not IsNull(RsXX!Detalle) Then pstrArticulo = pstrArticulo & " (" & Trim(RsXX!Detalle) & ")"
                .AddItem pstrArticulo

                .Cell(flexcpText, .Rows - 1, 1) = RsXX!RenCantidad
                .Cell(flexcpText, .Rows - 1, 2) = Format(RsXX!RenPrecio, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 3) = "FACTURA"
                .Cell(flexcpText, .Rows - 1, 4) = RsXX!RenCantidad - RsXX!RenARetirar
                .Cell(flexcpText, .Rows - 1, 5) = RsXX!RenARetirar
                If RsXX!RenARetirar <> 0 Then .Cell(flexcpBackColor, .Rows - 1, 5) = Colores.RojoClaro: .Cell(flexcpForeColor, .Rows - 1, 5) = Colores.Blanco: .Cell(flexcpFontBold, .Rows - 1, 5) = True
            End With
            RsXX.MoveNext
        Loop
        RsXX.Close
    End If
    '-----------------------------------------------------------------------------------------------
    Screen.MousePointer = 0

    
    'tFRetiro.Visible = bModificar.Enabled
'    If Not bModificar.Enabled Then lDFRetira.Caption = "Retira: " & tFRetiro.Text

    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los artículos a retirar.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub LimpioFichaEnvio()

    lEnvio.Rows = 1
    lEDireccion.Caption = "N/D"
    lEFecha.Caption = "N/D"
    lEEstado.Caption = "N/D"
    lECamion.Caption = "N/D"
    lEAgencia.Caption = "N/D"
    lEFlete.Caption = "N/D"
    lECosto.Caption = "N/D"
    lEPiso.Caption = "N/D"
    lEUsuario.Caption = "N/D"
    lEComentario.Caption = ""
    lEEntregado.Caption = "N/D"
    lEFPago.Caption = "N/D"
    lEVaCon.Caption = ""
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing

End Sub

Private Sub lDFRetira_Click()
    Foco tFRetiro
End Sub

Private Sub lArticulo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If lArticulo.Rows = 1 Then Exit Sub
    If prmIdDocumento = 0 Or prmIdCliente = 0 Then Exit Sub
    
    If Button = vbRightButton Then PopupMenu MnuARetirar, , , , MnuVerTitulo
    
End Sub

Private Sub lEnvio_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then PopupMenu MnuEnvio, , , , MnuIrTitulo
    
End Sub

Private Sub MnuD_Click(Index As Integer)
    CargoDocumento Val(MnuD(Index).Tag)
End Sub

Private Sub MnuInComentarios_Click()
    If prmIdDocumento = 0 Or prmIdCliente = 0 Then Exit Sub
    On Error GoTo errCom
    Screen.MousePointer = 11
    
    Dim objCom As New clsCliente
    objCom.ComentariosNuevo prmIdCliente, CStr("/D" & prmIdDocumento)
    Me.Refresh
    Set objCom = Nothing

    Screen.MousePointer = 0
    Exit Sub

errCom:
    clsGeneral.OcurrioError "Error al activar comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuInstalacion_Click()
    EjecutarApp prmPathApp & "Instalaciones.exe", "DOC:" & CStr(prmIdFactura)
End Sub

Private Sub MnuIrACamion_Click()

Dim aTitulo As String

    If Val(lECamion.Tag) = 0 Then Exit Sub
    On Error GoTo errCargar
    'Cargo los datos del camionero para visualizar
    cons = "Select * from Camion Where CamCodigo = " & lECamion.Tag
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    aTitulo = "Ficha del Camión: " & rsAux!CamNombre
    
    aTexto = "Empresa: "
    If Not IsNull(rsAux!CamEmpresa) Then aTexto = aTexto & Trim(rsAux!CamEmpresa) Else aTexto = aTexto & "S/D"
    aTexto = aTexto & Chr(vbKeyReturn)
    
    aTexto = aTexto & "Teléfono: "
    If Not IsNull(rsAux!CamTelefono) Then aTexto = aTexto & Trim(rsAux!CamTelefono) Else aTexto = aTexto & "S/D"
    aTexto = aTexto & Chr(vbKeyReturn)
    
    aTexto = aTexto & "Volumen: "
    If Not IsNull(rsAux!CamVolumen) Then aTexto = aTexto & Format(rsAux!CamVolumen, FormatoMonedaP) Else aTexto = aTexto & "S/D"
    aTexto = aTexto & Chr(vbKeyReturn)
    
    aTexto = aTexto & "Descripción: "
    If Not IsNull(rsAux!CamDescripcion) Then aTexto = aTexto & Trim(rsAux!CamDescripcion) Else aTexto = aTexto & "S/D"
    
    rsAux.Close
    Screen.MousePointer = 0
    MsgBox aTexto, vbInformation, aTitulo
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del camión."
End Sub

Private Sub MnuIrAEnvio_Click()
    On Error GoTo errEnvio
    If Tab1.Tabs.Count = 1 Then Exit Sub
    
    Dim aEnvio As Long
    aEnvio = Val(Mid(Tab1.SelectedItem.Key, 2, Len(Tab1.SelectedItem.Key)))
    If aEnvio = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    Dim objE As New clsEnvio
    objE.InvocoEnvio aEnvio, gPathListados
    Set objE = Nothing
    Screen.MousePointer = 0
    Exit Sub

errEnvio:
    clsGeneral.OcurrioError "Ocurrió un error al invocar el envío.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuMoComentarios_Click()
    
    On Error GoTo errCli
    If prmIdDocumento = 0 Then Exit Sub
    
    Dim bVaSuceso As Boolean
    
    'If Not miConexion.AccesoAlMenu("Movimientos Fisicos") Then
    '    MsgBox "Ud. no tiene permisos para modificar los comentarios.", vbExclamation, "No hay Permisos"
    '    Exit Sub
    'End If
    
    Dim sucUsuarioS As Long, sucUsuarioA As Long, sucDefensa As String
        
    bVaSuceso = Not miConexion.AccesoAlMenu("Movimientos Fisicos")
    
    If bVaSuceso Then       'Pido los datos del Suceso              ----------------------------------------------------
        
        Dim objSuceso As New clsSuceso
        objSuceso.TipoSuceso = prmSucesoVarios
        objSuceso.ActivoFormulario paCodigoDeUsuario, "Cambio de Comentario", cBase
        
        Me.Refresh
        sucUsuarioS = objSuceso.RetornoValor(Usuario:=True)
        sucDefensa = objSuceso.RetornoValor(Defensa:=True)
        sucUsuarioA = objSuceso.Autoriza
        Set objSuceso = Nothing
        
        If sucUsuarioS = 0 Or Trim(sucDefensa) = "" Then Exit Sub  'Abortó el ingreso del suceso
    End If                          '----------------------------------------------------------------------------------------------
    
    
    Dim mNewComm As String, mOldComm As String
    mOldComm = Trim(lComentario.Caption)
    
    mNewComm = "Ingrese el nuevo comentario para el documento" & Chr(vbKeyReturn) & _
                 "Para eliminarlo ingrese un '*' (asterisco)"
    mNewComm = InputBox(mNewComm, "Modificar Comentario", mOldComm)
    
    mNewComm = Trim(mNewComm)
    If mNewComm = "" Then Exit Sub
    
    Screen.MousePointer = 11             'Grabo Cambios ------------------------------------------------------------------------------------------
    
    If mNewComm = "*" Then
        cons = "Update Documento Set DocComentario = Null Where DocCodigo = " & prmIdDocumento
        cBase.Execute cons
        lComentario.Caption = ""
    Else
        If mNewComm <> mOldComm Then
            cons = "Update Documento Set DocComentario = '" & mNewComm & "' Where DocCodigo = " & prmIdDocumento
            cBase.Execute cons
            lComentario.Caption = " " & mNewComm
        End If
    End If
    
    If bVaSuceso And mNewComm <> mOldComm Then
        Dim mTexto As String
        mTexto = "Modifica Comentario del Documento"
        If mOldComm <> "" Then sucDefensa = mOldComm & " >DF: " & sucDefensa
        
        clsGeneral.RegistroSucesoAutorizado cBase, Now, prmSucesoVarios, paCodigoDeTerminal, sucUsuarioS, prmIdDocumento, _
                                 Descripcion:=mTexto, Defensa:=Trim(sucDefensa), _
                                 idCliente:=prmIdCliente, idAutoriza:=sucUsuarioA
    End If
    
    Screen.MousePointer = 0
    Exit Sub
errCli:
    clsGeneral.OcurrioError "Error al modificar el comentario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub MnuVerDetalle_Click()
    EjecutarApp prmPathApp & "Detalle de Operaciones", CStr(prmIdFactura)
End Sub

Private Sub MnuVerDeuda_Click()
    EjecutarApp prmPathApp & "Deuda en cheques", CStr(prmIdCliente)
End Sub

Private Sub MnuVerEntrega_Click()
    ListaDeEntregas
End Sub

Private Sub MnuVerEnvio_Click()
On Error GoTo errEnvio
    If prmIdDocumento = 0 Then Exit Sub
    If prmTipoD <> TipoDocumento.Contado And prmTipoD <> TipoDocumento.Credito Then Exit Sub
    
    Screen.MousePointer = 11
    Dim objE As New clsEnvio
    objE.InvocoEnvioXDocumento prmIdDocumento, prmTipoD, gPathListados
    Set objE = Nothing
    Screen.MousePointer = 0
    Exit Sub

errEnvio:
    clsGeneral.OcurrioError "Ocurrió un error al invocar el envío.", Err.Description
    Screen.MousePointer = 0

End Sub

Private Sub MnuVerOperaciones_Click()
    
    If prmIdCliente <> 0 Then EjecutarApp prmPathApp & "Visualizacion de Operaciones", CStr(prmIdCliente)
    
End Sub

Private Sub MnuVerPago_Click()
    EjecutarApp prmPathApp & "Detalle de pagos", CStr(prmIdFactura)
End Sub

Private Sub Tab1_Click()
    
    On Error Resume Next
    If Tab1.SelectedItem.Key = "sucesos" Then
        Picture1(2).ZOrder 0
        Me.Refresh: Exit Sub
    End If
    
    If Tab1.SelectedItem.Key = "comentarios" Then
        Picture1(3).ZOrder 0
        Me.Refresh: Exit Sub
    End If
    
    If Tab1.SelectedItem.Index > 1 Then
        Picture1(1).ZOrder 0
        Screen.MousePointer = 11
        Me.Refresh
        
        'Cargo los datos del Envio seleccionado
        LimpioFichaEnvio
        Picture1(1).Refresh
        CargoDatosEnvio Mid(Tab1.SelectedItem.Key, 2, Len(Tab1.SelectedItem.Key))
        CargoArticulosEnvio Mid(Tab1.SelectedItem.Key, 2, Len(Tab1.SelectedItem.Key))
        Screen.MousePointer = 0
        
    Else
        On Error Resume Next
        Picture1(0).ZOrder 0
        Me.Refresh
    End If
    
End Sub

Private Sub tFRetiro_GotFocus()

    tFRetiro.SelStart = 0
    tFRetiro.SelLength = Len(tFRetiro.Text)
    
End Sub

Private Sub tFRetiro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If bGrabar.Enabled Then bGrabar.SetFocus
End Sub

Private Sub tNumero_Change()
    lDoc.Caption = "": lDoc.Tag = 0
End Sub

Private Sub tNumero_GotFocus()
    tNumero.SelStart = 0: tNumero.SelLength = (Len(tNumero.Text))
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(lDoc.Tag) <> 0 Then Exit Sub
        If Trim(tNumero.Text) = "" Then Exit Sub
        
        LimpioFichaDocumento
        prmIdDocumento = 0
   
        Dim mDSerie As String, mDNumero As Long
        Dim adQ As Integer, adCodigo As Long, adTexto As String
        
        adTexto = Trim(tNumero.Text)
        If InStr(adTexto, "-") <> 0 Then
            mDSerie = Mid(adTexto, 1, InStr(adTexto, "-") - 1)
            mDNumero = Val(Mid(adTexto, InStr(adTexto, "-") + 1))
        Else
            mDSerie = Mid(adTexto, 1, 1)
            mDNumero = Val(Mid(adTexto, 2))
        End If
        tNumero.Text = UCase(mDSerie) & "-" & mDNumero
        
        Screen.MousePointer = 11
        adQ = 0: adTexto = ""
        
        cons = "Select DocCodigo, DocFecha as Fecha, DocSerie as Serie, Convert(char(7),DocNumero) as Numero " & _
                   " From Documento " & _
                   " Where DocSerie = '" & mDSerie & "'" & _
                   " And DocNumero = " & mDNumero & _
                   " And DocTipo IN (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ", " & TipoDocumento.NotaCredito & ", " & _
                                                   TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ", " & TipoDocumento.ReciboDePago & ")"
                                                   
            
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            adCodigo = rsAux!DocCodigo
            adQ = 1
            rsAux.MoveNext: If Not rsAux.EOF Then adQ = 2
        End If
        rsAux.Close
        
        Select Case adQ
            Case 2
                Dim miLDocs As New clsListadeAyuda
                adCodigo = miLDocs.ActivarAyuda(cBase, cons, 4100, 1)
                Me.Refresh
                If adCodigo > 0 Then adCodigo = miLDocs.RetornoDatoSeleccionado(0)
                Set miLDocs = Nothing
                
        End Select
        
        If adCodigo > 0 Then
            lDoc.Tag = adCodigo: lDoc.Caption = adTexto
            CargoDocumento adCodigo
        Else
            lDoc.Caption = " No Existe !!"
        End If
        
        Call tNumero_GotFocus
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errDoc:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function z_TextoDocumento(Tipo As Integer, Serie As String, Numero As Long) As String

    Select Case Tipo
        Case 1: z_TextoDocumento = "Ctdo. "
        Case 2: z_TextoDocumento = "Créd. "
        Case 3: z_TextoDocumento = "N/Dev. "
        Case 4: z_TextoDocumento = "N/Créd. "
        Case 5: z_TextoDocumento = "Recibo "
        Case 10: z_TextoDocumento = "N/Esp. "
    End Select
    
    z_TextoDocumento = z_TextoDocumento & Trim(Serie) & "-" & Numero

End Function

Private Sub LimpioFichaDocumento()

    lFecha.Caption = "" ' N/D"
    lImporte.Caption = "" ' N/D"
    lUsuario.Caption = "" 'N/D"
    lVendedor.Caption = "" ' N/D"
    lComentario.Caption = "" ' N/D"
    cPendiente.Text = ""
    tFRetiro.Text = ""
    lDFRetira.Caption = "&Retira:"
    
    lTitular.Caption = "N/D"
    lAnulada.Visible = False
    Me.Refresh
    
    lArticulo.Rows = 1: lArticulo.Refresh
    
    Do While Tab1.Tabs.Count > 1
        Tab1.Tabs.Remove Tab1.Tabs.Count
    Loop
    Picture1(0).ZOrder 0
    Tab1.Refresh: Picture1(0).Refresh
    
    cPendiente.Tag = ""
    tFRetiro.Tag = ""
    cPendiente.Enabled = False
    tFRetiro.Enabled = False
    cPendiente.BackColor = Inactivo
    tFRetiro.BackColor = Inactivo
    
    bModificar.ZOrder 0: bEliminar.ZOrder 0
    bModificar.Enabled = False: bEliminar.Enabled = False
     
    MnuVerDetalle.Enabled = False: MnuVerDeuda.Enabled = False: MnuVerPago.Enabled = False
    lDFRetira.Caption = "&Retira:"
End Sub

Private Sub CargoCliente(Cliente As Long)
    
    On Error GoTo errCliente
    cons = "Select CliCiRuc, CliTipo, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
           & " From Cliente, CPersona " _
           & " Where CliCodigo = " & Cliente _
           & " And CliCodigo = CPeCliente " _
                                                & " UNION " _
           & " Select CliCiRuc, CliTipo, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
           & " From Cliente, CEmpresa " _
           & " Where CliCodigo = " & Cliente _
           & " And CliCodigo = CEmCliente"

    Set RsXX = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    lTitular.Caption = ""
    If Not RsXX.EOF Then
        If RsXX!CliTipo = TipoCliente.Cliente Then
            If Not IsNull(RsXX!CliCIRuc) Then lTitular.Caption = clsGeneral.RetornoFormatoCedula(RsXX!CliCIRuc) & " - "
        Else
            If Not IsNull(RsXX!CliCIRuc) Then lTitular.Caption = clsGeneral.RetornoFormatoRuc(RsXX!CliCIRuc) & " - "
        End If
    End If
    lTitular.Caption = lTitular.Caption & Trim(RsXX!Nombre)
    RsXX.Close
    Exit Sub
    
errCliente:
    clsGeneral.OcurrioError "Error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoSucesos(Documento As Long)

    On Error GoTo errSuceso
    If Documento = 0 Then Exit Sub
    cons = "Select * from Suceso, TipoSuceso Where SucDocumento = " & Documento & " And SucTipo *= TSuCodigoSistema"
    Set RsXX = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    lSuceso.Rows = 1
    If Not RsXX.EOF Then
        Tab1.Tabs.Add pvCaption:="S&ucesos", pvkey:="sucesos"
        Do While Not RsXX.EOF
            With lSuceso
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Format(RsXX!SucFecha, "dd/mm/yy hh:mm")
                If Not IsNull(RsXX!TSuNombre) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(RsXX!TSuNombre)
                .Cell(flexcpText, .Rows - 1, 2) = z_BuscoUsuario(RsXX!sucUsuario, Identificacion:=True)
                If Not IsNull(RsXX!SucDescripcion) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsXX!SucDescripcion)
                If Not IsNull(RsXX!sucDefensa) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(RsXX!sucDefensa)
                
            End With
            RsXX.MoveNext
        Loop
    End If
    RsXX.Close
    Exit Sub

errSuceso:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los sucesos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoComentarios(Documento As Long, Optional Doc2 As Long = -1)

    On Error GoTo errSuceso
    If Documento = 0 And Doc2 = -1 Then Exit Sub
    
    cons = "Select * from Comentario, Usuario " & _
               " Where ComCliente = " & prmIdCliente & _
               " And ComDocumento In (" & Documento & ", " & Doc2 & ")" & _
               " And ComUsuario = UsuCodigo " & _
               " Order by ComFecha Desc"
    Set RsXX = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    vsComentario.Rows = 1
    If Not RsXX.EOF Then
        Tab1.Tabs.Add pvCaption:="&Comentarios", pvkey:="comentarios"
        Do While Not RsXX.EOF
            With vsComentario
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Format(RsXX!ComFecha, "dd/mm/yy hh:mm")
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsXX!ComComentario)
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsXX!UsuIdentificacion)
                
            End With
            RsXX.MoveNext
        Loop
        
        vsComentario.AutoSizeMode = flexAutoSizeRowHeight
        vsComentario.AutoSize 1, , False
    End If
    RsXX.Close
    Exit Sub

errSuceso:
    clsGeneral.OcurrioError "Error al cargar los comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub InicializoGrillas()

    With lArticulo
        .Rows = 1: .Cols = 1
        .FormatString = "<Artículo|>Q|>Importe x1|<Documento|>Entregados|<A Entregar"
        .ColWidth(0) = 3700: .ColWidth(1) = 500: .ColWidth(2) = 1100: .ColWidth(3) = 1200: .ColWidth(4) = 0
        .WordWrap = False: .ExtendLastCol = True
    End With

    With lEnvio
        .Rows = 1: .Cols = 1
        .FormatString = "<Artículo|>A Enviar|>Por Entregar"
        .ColWidth(0) = 5400
        .WordWrap = False: .ExtendLastCol = True
    End With

    With lSuceso
        .Rows = 1: .Cols = 1
        .FormatString = "<Fecha|Tipo|Usuario|Descripción|Defensa|"
        .ColWidth(0) = 1200: .ColWidth(1) = 1050: .ColWidth(2) = 950: .ColWidth(3) = 3000: .ColWidth(4) = 3600
        .WordWrap = False: .ExtendLastCol = True
    End With
    
     With vsComentario
        .Rows = 1: .Cols = 1
        .FormatString = "<Fecha|<Comentario|<Usuario"
        .ColWidth(0) = 1300: .ColWidth(1) = 4870: .ColWidth(2) = 1100
        .WordWrap = True: .ExtendLastCol = True
        .ColAlignment(0) = flexAlignLeftTop: .ColAlignment(1) = flexAlignLeftTop: .ColAlignment(2) = flexAlignLeftTop
    End With
    
End Sub

Private Sub ListaDeEntregas()

    If prmIdDocumento = 0 Then Exit Sub
    On Error GoTo errAyuda
    
    Screen.MousePointer = 11
    'Cargo los remitos para el Documento ---------------------------------------------------------------
    Dim mRemitos As String: mRemitos = ""
    cons = "Select RemCodigo From Remito Where RemDocumento = " & prmIdDocumento
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        mRemitos = rsAux!RemCodigo
        rsAux.MoveNext
        If Not rsAux.EOF Then mRemitos = mRemitos & ", "
    Loop
    rsAux.Close
    '------------------------------------------------------------------------------------------------------------
    
    'Consulta: Hora y quien entregó los arts de un documento
    cons = "SELECT MSFCodigo as Movimiento, MSFFecha AS Fecha, LocNombre as Local, ArtNombre as 'Artículo', MSFCantidad AS Q, EsMAbreviacion As Estado, MSFDocumento As Documento, UsuIdentificacion As Usuario " _
           & " FROM MovimientoStockFisico, Articulo, Usuario, Local, EstadoMercaderia" _
           & " Where MSFLocal = LocCodigo" _
           & " And MSFArticulo = ArtId" _
           & " And MSFUsuario = UsuCodigo" _
           & " And MSFEstado = EsMCodigo"
    
    If mRemitos = "" Then
        cons = cons & " And MSFDocumento = " & prmIdDocumento
        If prmTipoD <> 0 Then cons = cons & " And MSFTipoDocumento =" & prmTipoD
    Else
        cons = cons & " And ( (MSFDocumento = " & prmIdDocumento
        If prmTipoD <> 0 Then cons = cons & " And MSFTipoDocumento =" & prmTipoD
        cons = cons & " ) OR ( MSFDocumento in ( " & mRemitos & ") And MSFTipoDocumento =" & TipoDocumento.Remito & ") )"
    End If
    
    Dim objLista As New clsListadeAyuda
    objLista.ActivoListaAyudaSQL cBase, cons
    Set objLista = Nothing
    Screen.MousePointer = 0
    Exit Sub

errAyuda:
    clsGeneral.OcurrioError "Error al activar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub CargoDocumento(idDocumento As Long)
On Error GoTo errCargar

    Screen.MousePointer = 11
    LimpioFichaDocumento
    prmIdDocumento = 0
    prmTipoD = 0
    prmIdCliente = 0
    prmIdFactura = 0
    
    
'SELECT (Case DatePart(hh, DocFRetira)
'     When 1 Then DocFRetira - DatePart(n, DocFRetira)
'     Else DocFRetira
'    End) Desde, DocFRetira  Hasta
    
    cons = "Select Documento.*, MonSigno, SucAbreviacion, (Case DatePart(hh, DocFRetira)" & _
                                                        " When 1 Then DocFRetira - DatePart(n, DocFRetira) " & _
                                                        " Else DocFecha " & _
                                                        " End) RetiraDesde, ISnull(DocFRetira, DocFecha) RetiraHasta" & _
            " From Documento, Moneda, Sucursal" & _
            " Where DocCodigo = " & idDocumento & _
            " And DocMoneda = MonCodigo" & _
            " And DocSucursal = SucCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsAux.EOF Then
        tNumero.Text = UCase(rsAux!DocSerie) & "-" & rsAux!DocNumero
        prmIdDocumento = rsAux!DocCodigo
        prmTipoD = rsAux!DocTipo
        prmIdCliente = rsAux!DocCliente
        
        lDoc.Caption = z_TextoDocumento(prmTipoD, Trim(rsAux!DocSerie), rsAux!DocNumero)
        If Not IsNull(rsAux!SucAbreviacion) Then lDoc.Caption = lDoc.Caption & " (" & Trim(rsAux!SucAbreviacion) & ")"
        
        lFecha.Caption = " " & Format(rsAux!DocFecha, "d/mm/yy hh:mm")
        lUsuario.Caption = " " & z_BuscoUsuario(rsAux!DocUsuario, Identificacion:=True)
        If Not IsNull(rsAux!DocVendedor) Then lVendedor.Caption = " " & z_BuscoUsuario(rsAux!DocVendedor, Identificacion:=True) Else: lVendedor.Caption = ""
    
        lComentario.Caption = ""
        If Not IsNull(rsAux!DocComentario) Then lComentario.Caption = " " & Trim(rsAux!DocComentario)
    
        lImporte.Tag = rsAux!DocMoneda
        lImporte.Caption = " " & Trim(rsAux!MonSigno) & " " & Format(rsAux!DocTotal, FormatoMonedaP)
        
        If rsAux!DocAnulado Then
            MsgBox "El documento seleccionado ha sido anulado.", vbInformation, "Documento Anulado"
            lAnulada.Visible = True
        End If
    
        If Not IsNull(rsAux!DocPendiente) Then
            BuscoCodigoEnCombo cPendiente, rsAux!DocPendiente
            cPendiente.Tag = rsAux!DocPendiente
            bModificar.Enabled = True: bEliminar.Enabled = True
        End If
    
        'If Not IsNull(rsAux!DocFRetira) Then
        '    tFRetiro.Text = Format(rsAux!DocFRetira, "d-Mmm yy")
        '    tFRetiro.Tag = Format(rsAux!DocFRetira, "d-Mmm yy")
        'End If
        
        lDFRetira.Tag = ""
        tFRetiro.Tag = ""
        If Not IsNull(rsAux!RetiraDesde) Then
            lDFRetira.Caption = "&Retira desde " & Format(rsAux!RetiraDesde, "d-Mmm yy") & " al"
            lDFRetira.Tag = Format(rsAux!RetiraDesde, "dd/mm/yyyy")
        End If
        If Not IsNull(rsAux!RetiraHasta) Then
            tFRetiro.Text = Format(rsAux!RetiraHasta, "d-Mmm yy")
            tFRetiro.Tag = Format(rsAux!RetiraHasta, "d-Mmm yy")
        End If
        If Not IsNull(rsAux!RetiraDesde) And Not IsNull(rsAux!RetiraHasta) Then
            If Format(rsAux!RetiraDesde, "dd/mm/yyyy") = Format(rsAux!RetiraHasta, "dd/mm/yyyy") Then lDFRetira.Caption = "&Retira:"
        End If
        
        Me.Refresh
    
        MnuVerDetalle.Enabled = True: MnuVerDeuda.Enabled = True
        If prmTipoD = TipoDocumento.Contado Then MnuVerPago.Enabled = False Else: MnuVerPago.Enabled = True
        
    End If
    rsAux.Close
    
    If prmIdDocumento = 0 Then
        Screen.MousePointer = 0
        MsgBox "No existe un documento para las características ingresadas.", vbExclamation, "No Hay Datos"
    End If
    
    If prmIdDocumento <> 0 Then CargoDemasDatos
    
    Screen.MousePointer = 0
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDemasDatos()

    CargoCliente prmIdCliente
        
    Select Case prmTipoD
        Case TipoDocumento.Contado, TipoDocumento.Credito
                prmIdFactura = prmIdDocumento
                CargoArticulosARetirar prmIdDocumento
                
                TabsEnvios prmIdDocumento
                Tab1.Tabs(1).Selected = True: DoEvents
                CargoSucesos prmIdDocumento
                CargoComentarios prmIdDocumento
            
        Case TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial
        
                cons = "Select * from Nota Where NotNota = " & prmIdDocumento
                Set RsXX = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsXX.EOF Then
                    prmIdFactura = RsXX!NotFactura
                    If Not IsNull(RsXX!NotSalidaCaja) Then lComentario.Caption = Trim(lComentario.Caption) & " - Devuelve: " & Format(RsXX!NotSalidaCaja, FormatoMonedaP)
                End If
                RsXX.Close
                    
                If prmIdFactura <> 0 Then
                    CargoArticulosARetirar prmIdFactura
                    TabsEnvios prmIdFactura
                    Tab1.Tabs(1).Selected = True: DoEvents
                    CargoSucesos prmIdFactura
                    CargoComentarios prmIdFactura, prmIdDocumento
                Else
                    If prmTipoD = TipoDocumento.NotaEspecial Then
                        CargoArticulosARetirar prmIdDocumento
                        DoEvents
                        CargoSucesos prmIdDocumento
                        CargoComentarios prmIdDocumento
                    End If
                    MsgBox "No se encontró la relación factura-nota.", vbExclamation, "Falta Relación"
                End If
                
                bModificar.Enabled = False: bEliminar.Enabled = False
            
        Case TipoDocumento.ReciboDePago
                
                cons = "Select * from DocumentoPago Where DPaDocQSalda = " & prmIdDocumento
                Set RsXX = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                Do While Not RsXX.EOF
                    If prmIdFactura = 0 Or RsXX!DPaAmortizacion <> 0 Or RsXX!DPaMora <> 0 Then prmIdFactura = RsXX!DPaDocASaldar
                    RsXX.MoveNext
                Loop
                RsXX.Close
        
                CargoArticulosARetirar prmIdFactura
                
                TabsEnvios prmIdFactura
                Tab1.Tabs(1).Selected = True: DoEvents
                CargoSucesos prmIdFactura
                CargoComentarios prmIdFactura
                
                bModificar.Enabled = False: bEliminar.Enabled = False
    End Select
    
    If MnuD.Count > 1 Then
        MnuD(0).Visible = False
        MnuDocs.Enabled = True
    Else
        MnuDocs.Enabled = False
    End If
    
    
End Sub


Private Function AddMenuDoc(mIdDoc As Long, mTexto As String)
On Error GoTo errMnu
    For I = 1 To MnuD.UBound
        If Val(MnuD(I).Tag) = mIdDoc Then Exit Function
    Next
    
    Load MnuD(MnuD.UBound + 1)
    With MnuD(MnuD.UBound)
        .Caption = mTexto
        .Tag = mIdDoc
    End With

errMnu:
End Function

