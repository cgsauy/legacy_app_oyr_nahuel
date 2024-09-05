VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{1292AE18-2B08-4CE3-9F79-9CB06F26AB54}#1.7#0"; "orEMails.ocx"
Begin VB.Form frmResolucion 
   Caption         =   "Estudio de Solicitudes"
   ClientHeight    =   8280
   ClientLeft      =   3120
   ClientTop       =   2385
   ClientWidth     =   12045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResolucion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   12045
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Index           =   2
      Left            =   660
      ScaleHeight     =   5115
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   2340
      Width           =   675
      Begin VSFlex6DAOCtl.vsFlexGrid lEmpleoT 
         Height          =   1215
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   9375
         _ExtentX        =   16536
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
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
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
      Begin VSFlex6DAOCtl.vsFlexGrid lReferenciaT 
         Height          =   1215
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   9375
         _ExtentX        =   16536
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
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
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
      Begin VSFlex6DAOCtl.vsFlexGrid lComentarioT 
         Height          =   1215
         Left            =   120
         TabIndex        =   47
         Top             =   2880
         Width           =   9375
         _ExtentX        =   16536
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
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
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
   End
   Begin VB.PictureBox Picture1 
      Height          =   5355
      Index           =   7
      Left            =   4920
      ScaleHeight     =   5295
      ScaleWidth      =   10095
      TabIndex        =   96
      Top             =   5040
      Width           =   10155
      Begin VB.ComboBox cCondicionR 
         Height          =   315
         ItemData        =   "frmResolucion.frx":08CA
         Left            =   2460
         List            =   "frmResolucion.frx":08CC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Frame frmCondicion 
         Caption         =   "Valores Condicionales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1395
         Left            =   1380
         TabIndex        =   97
         Top             =   240
         Width           =   8775
         Begin VB.CheckBox cRDefinitiva 
            Appearance      =   0  'Flat
            Caption         =   "Resolución Definitiva"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6840
            TabIndex        =   150
            Top             =   40
            Width           =   1815
         End
         Begin VB.TextBox tRMonto 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2280
            TabIndex        =   107
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox tSMonto 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5280
            TabIndex        =   104
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox tPEntrega 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7380
            TabIndex        =   111
            Top             =   600
            Width           =   1335
         End
         Begin MSMask.MaskEdBox tCValor1 
            Height          =   315
            Left            =   2760
            TabIndex        =   114
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox tCValor2 
            Height          =   315
            Left            =   5280
            TabIndex        =   116
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin AACombo99.AACombo cRMoneda 
            Height          =   315
            Left            =   1440
            TabIndex        =   106
            Top             =   600
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
            Text            =   ""
         End
         Begin AACombo99.AACombo cPCuota 
            Height          =   315
            Left            =   5280
            TabIndex        =   109
            Top             =   600
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
         Begin MSMask.MaskEdBox tGarantia 
            Height          =   315
            Left            =   1440
            TabIndex        =   102
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            PromptInclude   =   0   'False
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#.###.###-#"
            PromptChar      =   "_"
         End
         Begin VB.Label lCValor2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Valor2:"
            Height          =   255
            Left            =   4080
            TabIndex        =   115
            Top             =   1020
            Width           =   1095
         End
         Begin VB.Label lCValor1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Valor1:"
            Height          =   255
            Left            =   1440
            TabIndex        =   113
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "Recibo Sueldo ..."
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Garantía ..."
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "Cambio de Plan ..."
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3840
            TabIndex        =   108
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega:"
            Height          =   255
            Left            =   6660
            TabIndex        =   110
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Solicitud ..."
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3840
            TabIndex        =   103
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "Comprobante ..."
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   1020
            Width           =   1335
         End
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2460
         MaxLength       =   100
         TabIndex        =   119
         Top             =   2040
         Width           =   7695
      End
      Begin VSFlex6DAOCtl.vsFlexGrid lCondicion 
         Height          =   1395
         Left            =   1380
         TabIndex        =   117
         Top             =   2520
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2461
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
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Resolución:"
         Height          =   255
         Left            =   1440
         TabIndex        =   100
         Top             =   1740
         Width           =   975
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000005&
         X1              =   60
         X2              =   1200
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   60
         X2              =   1200
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         X1              =   60
         X2              =   1200
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   60
         X2              =   1200
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Image imgCondRes 
         Height          =   480
         Index           =   8
         Left            =   60
         MouseIcon       =   "frmResolucion.frx":08CE
         MousePointer    =   99  'Custom
         Picture         =   "frmResolucion.frx":0BD8
         Top             =   3180
         Width           =   480
      End
      Begin VB.Image imgCondRes 
         Height          =   480
         Index           =   7
         Left            =   660
         MouseIcon       =   "frmResolucion.frx":11FA
         MousePointer    =   99  'Custom
         Picture         =   "frmResolucion.frx":1504
         Top             =   2640
         Width           =   480
      End
      Begin VB.Image imgCondRes 
         Height          =   480
         Index           =   6
         Left            =   60
         MouseIcon       =   "frmResolucion.frx":1B89
         MousePointer    =   99  'Custom
         Picture         =   "frmResolucion.frx":1E93
         Top             =   2640
         Width           =   480
      End
      Begin VB.Image imgCondRes 
         Height          =   480
         Index           =   5
         Left            =   60
         MouseIcon       =   "frmResolucion.frx":23DD
         MousePointer    =   99  'Custom
         Picture         =   "frmResolucion.frx":26E7
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image imgCondRes 
         Height          =   480
         Index           =   4
         Left            =   660
         MouseIcon       =   "frmResolucion.frx":2CD0
         MousePointer    =   99  'Custom
         Picture         =   "frmResolucion.frx":2FDA
         Top             =   1500
         Width           =   480
      End
      Begin VB.Image imgCondRes 
         Height          =   480
         Index           =   3
         Left            =   60
         MouseIcon       =   "frmResolucion.frx":36A1
         MousePointer    =   99  'Custom
         Picture         =   "frmResolucion.frx":39AB
         Top             =   1500
         Width           =   480
      End
      Begin VB.Image imgCondRes 
         Height          =   480
         Index           =   2
         Left            =   660
         MouseIcon       =   "frmResolucion.frx":3FFF
         MousePointer    =   99  'Custom
         Picture         =   "frmResolucion.frx":4309
         Top             =   780
         Width           =   480
      End
      Begin VB.Image imgCondRes 
         Height          =   480
         Index           =   1
         Left            =   60
         MouseIcon       =   "frmResolucion.frx":48C1
         MousePointer    =   99  'Custom
         Picture         =   "frmResolucion.frx":4BCB
         Top             =   780
         Width           =   480
      End
      Begin VB.Image imgCondRes 
         Height          =   480
         Index           =   0
         Left            =   60
         MouseIcon       =   "frmResolucion.frx":529D
         MousePointer    =   99  'Custom
         Picture         =   "frmResolucion.frx":55A7
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lRComentario 
         BackStyle       =   0  'Transparent
         Caption         =   "Co&mentarios:"
         Height          =   255
         Left            =   1440
         TabIndex        =   118
         Top             =   2100
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5595
      Index           =   5
      Left            =   2160
      ScaleHeight     =   5535
      ScaleWidth      =   10935
      TabIndex        =   5
      Top             =   4200
      Width           =   10995
      Begin VB.CommandButton bCRefrescar 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   7020
         TabIndex        =   27
         Top             =   30
         Width           =   1095
      End
      Begin VB.PictureBox pClearing 
         Height          =   1215
         Index           =   0
         Left            =   480
         ScaleHeight     =   1155
         ScaleWidth      =   1155
         TabIndex        =   42
         Top             =   2160
         Width           =   1215
         Begin VSFlex6DAOCtl.vsFlexGrid lCSolicitud 
            Height          =   735
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   1296
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
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
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
      End
      Begin VB.CommandButton bClearing 
         Caption         =   "Realizar &..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   8220
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmResolucion.frx":5BEA
         TabIndex        =   28
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin MSComctlLib.TabStrip TabClearing 
         Height          =   2595
         Left            =   120
         TabIndex        =   41
         Top             =   1800
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4577
         TabWidthStyle   =   2
         TabFixedWidth   =   4851
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Solicitudes Reali&zadas"
               Key             =   "solicitudes "
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Antecedentes "
               Key             =   "antecedentes"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "C&heques Sin Fondo"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TabStrip TabCPersona 
         Height          =   360
         Left            =   75
         TabIndex        =   51
         Top             =   45
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         TabWidthStyle   =   2
         Style           =   1
         TabFixedWidth   =   2028
         TabFixedHeight  =   616
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Titular"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Garantía"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Cónyuge"
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
      Begin VB.Label lClCosto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "$ 40"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9420
         TabIndex        =   151
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   1425
      End
      Begin VB.Label lClInfoMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         Caption         =   "Ch.s/Fondo && Obs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   148
         Top             =   1380
         Width           =   1665
      End
      Begin VB.Label lClFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "12/12/71"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4980
         TabIndex        =   37
         Top             =   75
         UseMnemonic     =   0   'False
         Width           =   1920
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Alquiler:"
         Height          =   255
         Left            =   6000
         TabIndex        =   40
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lClAlquiler 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "U$S 12,000.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6960
         TabIndex        =   39
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label72 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Último Clearing:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3720
         TabIndex        =   38
         Top             =   45
         Width           =   3225
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   1395
         Width           =   705
      End
      Begin VB.Label lClDireccionE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   35
         Top             =   1395
         UseMnemonic     =   0   'False
         Width           =   6735
      End
      Begin VB.Label Label68 
         BackStyle       =   0  'Transparent
         Caption         =   "E. Civil:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lClECivil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   360
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "F. Nacimiento:"
         Height          =   255
         Left            =   4320
         TabIndex        =   32
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lClFNacimiento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "12/12/71"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5400
         TabIndex        =   31
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "Antigüedad:"
         Height          =   255
         Left            =   6000
         TabIndex        =   30
         Top             =   1185
         Width           =   945
      End
      Begin VB.Label lClAntiguedad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6960
         TabIndex        =   29
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   360
      End
      Begin VB.Label lClTelefono 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Preferencial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7800
         TabIndex        =   26
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   960
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         Height          =   255
         Left            =   6960
         TabIndex        =   25
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lClEmpleo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Empleado Público"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "Empleos:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1185
         Width           =   735
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Cónyuge:"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lClConyuge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3.709.385-6"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   1695
      End
      Begin VB.Label lClDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Niagara 2345"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   4815
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   945
         Width           =   705
      End
      Begin VB.Label lClNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   510
         UseMnemonic     =   0   'False
         Width           =   7695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   510
         Width           =   855
      End
      Begin VB.Shape shpClearing 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   1215
         Left            =   120
         Top             =   480
         Width           =   8775
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5295
      Index           =   0
      Left            =   1560
      ScaleHeight     =   5235
      ScaleWidth      =   9435
      TabIndex        =   4
      Top             =   3420
      Width           =   9495
      Begin VB.OptionButton oVtaTel 
         Caption         =   "&Vtas. Telefónicas"
         Height          =   255
         Left            =   7140
         TabIndex        =   12
         Top             =   0
         Width           =   1995
      End
      Begin MSComctlLib.TabStrip TabPersona 
         Height          =   495
         Left            =   75
         TabIndex        =   7
         Top             =   45
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   873
         TabWidthStyle   =   2
         Style           =   1
         TabFixedWidth   =   2028
         TabFixedHeight  =   616
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Titular"
               Key             =   "titular"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Garantía"
               Key             =   "garantia"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Cónyuge"
               Key             =   "conyuge"
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
      Begin VB.OptionButton oArticulo 
         Caption         =   "&Artículos "
         Height          =   255
         Left            =   5400
         TabIndex        =   11
         Top             =   240
         Width           =   1755
      End
      Begin VB.OptionButton oSolicitud 
         Caption         =   "&Solicitudes"
         Height          =   255
         Left            =   5400
         TabIndex        =   10
         Top             =   0
         Width           =   1755
      End
      Begin VB.OptionButton oContado 
         Caption         =   "C&ontado"
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton oCredito 
         Caption         =   "Cré&dito"
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   0
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VSFlex6DAOCtl.vsFlexGrid lOperacion 
         Height          =   735
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1296
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&lista"
         Height          =   255
         Left            =   180
         TabIndex        =   49
         Top             =   1800
         Width           =   285
      End
      Begin VB.Label lTotalComprado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Deuda Pendiente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   4500
         UseMnemonic     =   0   'False
         Width           =   4455
      End
      Begin VB.Label lDeuda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Deuda Pendiente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   4500
         UseMnemonic     =   0   'False
         Width           =   1395
      End
      Begin VB.Label lPieOp 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   4500
         Width           =   8760
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Index           =   1
      Left            =   180
      ScaleHeight     =   5115
      ScaleWidth      =   11175
      TabIndex        =   69
      Top             =   5280
      Width           =   11235
      Begin VB.TextBox tCodigo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   120
         Width           =   1095
      End
      Begin VSFlex6DAOCtl.vsFlexGrid lArticulo 
         Height          =   1335
         Left            =   120
         TabIndex        =   141
         Top             =   2100
         Width           =   4035
         _ExtentX        =   7117
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
      Begin VB.Label lMoneda2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2940
         TabIndex        =   142
         Top             =   660
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Solicitud Nº:"
         Height          =   255
         Left            =   240
         TabIndex        =   95
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lMoneda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7800
         TabIndex        =   94
         Top             =   660
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Solicitud:"
         Height          =   255
         Left            =   2880
         TabIndex        =   93
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "12/12/98"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4440
         TabIndex        =   92
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1455
      End
      Begin VB.Label lUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Efectivo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6720
         TabIndex        =   91
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   1575
      End
      Begin VB.Label lPago 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque Diferido"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4440
         TabIndex        =   90
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   6060
         TabIndex        =   89
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago:"
         Height          =   255
         Left            =   3120
         TabIndex        =   88
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Entrega Diferida:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   87
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label lFEntrega 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2040
         TabIndex        =   86
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label lSEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6720
         TabIndex        =   85
         Top             =   900
         UseMnemonic     =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrega Efectiva:"
         Height          =   255
         Left            =   5400
         TabIndex        =   84
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Financiado:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   83
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lFMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1800
         TabIndex        =   82
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lFCuota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2040
         TabIndex        =   81
         Top             =   1140
         UseMnemonic     =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuotas:"
         Height          =   255
         Left            =   480
         TabIndex        =   80
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label43 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Financiado:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   79
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label lSFinanciado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6720
         TabIndex        =   78
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label Label45 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Total:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   77
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lSMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6360
         TabIndex        =   76
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label47 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Financiación                                         "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   75
         Top             =   660
         Width           =   2775
      End
      Begin VB.Label Label48 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Totales                                        "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   74
         Top             =   660
         Width           =   2655
      End
      Begin VB.Line Line1 
         X1              =   480
         X2              =   3240
         Y1              =   1400
         Y2              =   1400
      End
      Begin VB.Line Line2 
         X1              =   5400
         X2              =   8040
         Y1              =   1400
         Y2              =   1400
      End
      Begin VB.Label lComentario 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Efectivo"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   73
         Top             =   1755
         UseMnemonic     =   0   'False
         Width           =   8295
      End
      Begin VB.Label lSucursal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "12/12/98"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1200
         TabIndex        =   72
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal:"
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpSolicitud 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   1635
         Left            =   120
         Top             =   60
         Width           =   8295
      End
   End
   Begin orEMails.ctrEMails cEMailT 
      Height          =   315
      Left            =   840
      TabIndex        =   149
      Top             =   450
      Width           =   585
      _ExtentX        =   582
      _ExtentY        =   556
      BackColor       =   12648447
      ForeColor       =   0
      Modalidad       =   2
   End
   Begin MSComctlLib.Toolbar tbTool 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   147
      Top             =   0
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "img2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "autor"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "resolverysig"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "resolver"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "soltar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "llamara"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "plantillas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sepsalir"
            Style           =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   5115
      Index           =   3
      Left            =   5100
      ScaleHeight     =   5055
      ScaleWidth      =   1515
      TabIndex        =   98
      Top             =   2640
      Width           =   1575
      Begin VB.ComboBox cDireccionG 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   144
         Top             =   790
         Width           =   1635
      End
      Begin VSFlex6DAOCtl.vsFlexGrid lEmpleoG 
         Height          =   735
         Left            =   120
         TabIndex        =   138
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1296
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
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
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
      Begin VSFlex6DAOCtl.vsFlexGrid lReferenciaG 
         Height          =   735
         Left            =   120
         TabIndex        =   139
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1296
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
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
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
      Begin VSFlex6DAOCtl.vsFlexGrid lComentarioG 
         Height          =   735
         Left            =   120
         TabIndex        =   140
         Top             =   3000
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1296
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
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
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
      Begin VB.Image bDireccionG 
         Height          =   285
         Left            =   8740
         Picture         =   "frmResolucion.frx":5CFA
         Stretch         =   -1  'True
         Top             =   795
         Width           =   345
      End
      Begin VB.Label lCategoriaG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Preferencial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3240
         TabIndex        =   137
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría:"
         Height          =   255
         Left            =   2400
         TabIndex        =   136
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil:"
         Height          =   255
         Left            =   4440
         TabIndex        =   135
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lECivilG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lECivil"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5400
         TabIndex        =   134
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Edad:"
         Height          =   255
         Left            =   4920
         TabIndex        =   133
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lEdadG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "22 (5-Mar-1976)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5400
         TabIndex        =   132
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1680
      End
      Begin VB.Label lOcupacionG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Empleado Público"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   131
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Ocupación:"
         Height          =   255
         Left            =   240
         TabIndex        =   130
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblInfoCIConyuge 
         BackStyle       =   0  'Transparent
         Caption         =   "CI:"
         Height          =   255
         Left            =   240
         TabIndex        =   129
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lCIG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3.709.385-6"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   128
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1335
      End
      Begin VB.Label lDireccionG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Niagara 2345"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1920
         TabIndex        =   127
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   6975
      End
      Begin VB.Label lNDireccionG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   240
         TabIndex        =   126
         Top             =   840
         Width           =   705
      End
      Begin VB.Label lGarantia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   125
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   124
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lNSolicitudG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "12345678911"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7740
         TabIndex        =   123
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   960
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Sol.:"
         Height          =   255
         Left            =   7380
         TabIndex        =   122
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lTituloG 
         BackStyle       =   0  'Transparent
         Caption         =   "Títulos (2)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7380
         TabIndex        =   121
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lChequesG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Opera c/cheques"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7380
         TabIndex        =   120
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1440
      End
      Begin VB.Shape shpGarantia 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   1110
         Left            =   120
         Top             =   60
         Width           =   8775
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Index           =   6
      Left            =   6900
      ScaleHeight     =   5115
      ScaleWidth      =   495
      TabIndex        =   145
      Top             =   2520
      Width           =   555
      Begin VSFlex6DAOCtl.vsFlexGrid vsRelacion 
         Height          =   1215
         Left            =   120
         TabIndex        =   146
         Top             =   120
         Width           =   9375
         _ExtentX        =   16536
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
   End
   Begin VB.Timer tmClearing 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10020
      Top             =   540
   End
   Begin VB.ComboBox cDireccion 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   143
      Top             =   1350
      Width           =   1635
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   1095
      Left            =   300
      TabIndex        =   6
      Top             =   2100
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   1931
      TabWidthStyle   =   2
      TabFixedWidth   =   2293
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Operaciones"
            Key             =   "operaciones"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Solicitud"
            Key             =   "solicitud"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Titular"
            Key             =   "titular"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Garantía"
            Key             =   "garantia"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cón&yuge"
            Key             =   "conyuge"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Clearing"
            Key             =   "clearing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Relac&iones"
            Key             =   "relaciones"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Resolver"
            Key             =   "resolver"
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
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   44
      Top             =   8025
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15584
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img2 
      Left            =   9420
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":5EE4
            Key             =   "cerrar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":61FE
            Key             =   "resolver"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":6518
            Key             =   "plantillas"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":6832
            Key             =   "llamar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":6B4C
            Key             =   "auto"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":6E66
            Key             =   "soltar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":7180
            Key             =   "resolver2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9420
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":749A
            Key             =   "Si"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":77B4
            Key             =   "No"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":7ACE
            Key             =   "Alerta"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":7DE8
            Key             =   "Gestor"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":8102
            Key             =   "Perdida"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":841C
            Key             =   "Clearing"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":8736
            Key             =   "Nota"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":8A50
            Key             =   "clearing1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":8D6A
            Key             =   "clearing0"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":9084
            Key             =   "clearing6"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":939E
            Key             =   "clearing4"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":96B8
            Key             =   "clearing3"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":99D2
            Key             =   "operaciones"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":9AE4
            Key             =   "titular"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":9BF6
            Key             =   "solicitudes"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":9D08
            Key             =   "antecedentes"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":9E1A
            Key             =   "cheques"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":9F2C
            Key             =   "conyuge"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":A036
            Key             =   "conyugeno"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":A148
            Key             =   "resolver"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":A4DA
            Key             =   "garantia"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":A7F4
            Key             =   "garantiano"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":AB0E
            Key             =   "solicitud"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":AE28
            Key             =   "noconyuge"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":B142
            Key             =   "relaciones"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResolucion.frx":B45C
            Key             =   "relacionesno"
         EndProperty
      EndProperty
   End
   Begin VB.Label lFinanciacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Niagara 2345"
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
      Left            =   240
      TabIndex        =   67
      Top             =   1695
      UseMnemonic     =   0   'False
      Width           =   7455
   End
   Begin VB.Label lCiRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3.709.385-6"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   66
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblInfoCI 
      BackStyle       =   0  'Transparent
      Caption         =   "CI / RUC:"
      Height          =   255
      Left            =   240
      TabIndex        =   65
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lEdad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "22 (5-Mar-1976)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5520
      TabIndex        =   64
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Edad:"
      Height          =   255
      Left            =   4980
      TabIndex        =   63
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Categoría:"
      Height          =   255
      Left            =   2400
      TabIndex        =   62
      Top             =   720
      Width           =   960
   End
   Begin VB.Label lCategoria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Preferencial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3240
      TabIndex        =   61
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   1680
   End
   Begin VB.Label lTelefono 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Casa 9242557; Celular 099405236Casa 9242557; Celular 099405236Casa 9242557; Celular 099405236"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1080
      TabIndex        =   60
      Top             =   1140
      UseMnemonic     =   0   'False
      Width           =   5595
   End
   Begin VB.Label Label56 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfonos:"
      Height          =   255
      Left            =   240
      TabIndex        =   59
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label57 
      BackStyle       =   0  'Transparent
      Caption         =   "Sol.:"
      Height          =   255
      Left            =   8220
      TabIndex        =   58
      Top             =   1410
      Width           =   360
   End
   Begin VB.Label lNSolicitud 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "12345678911"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   8580
      TabIndex        =   57
      Top             =   1410
      UseMnemonic     =   0   'False
      Width           =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ocupación:"
      Height          =   255
      Left            =   240
      TabIndex        =   56
      Top             =   930
      Width           =   855
   End
   Begin VB.Label lOcupacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Empleado Público"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   55
      Top             =   930
      UseMnemonic     =   0   'False
      Width           =   3015
   End
   Begin VB.Label lECivil 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lECivil"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5520
      TabIndex        =   54
      Top             =   930
      UseMnemonic     =   0   'False
      Width           =   1560
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Estado Civil:"
      Height          =   255
      Left            =   4500
      TabIndex        =   53
      Top             =   930
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Titular:"
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   500
      Width           =   555
   End
   Begin VB.Image bDireccionT 
      Height          =   285
      Left            =   7740
      Picture         =   "frmResolucion.frx":B776
      Stretch         =   -1  'True
      Top             =   1380
      Width           =   345
   End
   Begin VB.Label lCheques 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Opera c/cheques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6840
      TabIndex        =   43
      Top             =   1170
      UseMnemonic     =   0   'False
      Width           =   1260
   End
   Begin VB.Label lTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Títulos (2)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6600
      TabIndex        =   13
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lDireccion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Niagara 2345"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1920
      TabIndex        =   2
      Top             =   1395
      UseMnemonic     =   0   'False
      Width           =   960
   End
   Begin VB.Label lNDireccion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1395
      Width           =   705
   End
   Begin VB.Label lTitular 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   495
      UseMnemonic     =   0   'False
      Width           =   5115
   End
   Begin VB.Label lAFinanciar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Detalle a Financiar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   1680
      Width           =   9135
   End
   Begin VB.Shape shpTitular 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1515
      Left            =   120
      Top             =   435
      Width           =   9135
   End
   Begin VB.Menu MnuTitular 
      Caption         =   "Titular"
      Visible         =   0   'False
      Begin VB.Menu MnuTTitular 
         Caption         =   "Menú Titular"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuTL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTFicha 
         Caption         =   "&Ficha del Cliente"
      End
      Begin VB.Menu MnuTEmpleo 
         Caption         =   "&Empleos"
      End
      Begin VB.Menu MnuTReferencia 
         Caption         =   "&Referencias"
      End
      Begin VB.Menu MnuTTitulo 
         Caption         =   "&Títulos"
      End
      Begin VB.Menu MnuTComentario 
         Caption         =   "&Comentarios"
      End
      Begin VB.Menu MnuTL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTRefrescar 
         Caption         =   "Refrescar datos..."
      End
   End
   Begin VB.Menu MnuConyuge 
      Caption         =   "Cónyuge"
      Visible         =   0   'False
      Begin VB.Menu MnuTConyuge 
         Caption         =   "Menú Cónyuge"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuCL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCFicha 
         Caption         =   "&Ficha del Cliente"
      End
      Begin VB.Menu MnuCEmpleo 
         Caption         =   "&Empleos"
      End
      Begin VB.Menu MnuCReferencia 
         Caption         =   "&Referencias"
      End
      Begin VB.Menu MnuCTitulo 
         Caption         =   "&Títulos"
      End
      Begin VB.Menu MnuCComentario 
         Caption         =   "&Comentarios"
      End
      Begin VB.Menu MnuCL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCRefrescar 
         Caption         =   "Refrescar datos..."
      End
   End
   Begin VB.Menu MnuOpCredito 
      Caption         =   "OpCredito"
      Visible         =   0   'False
      Begin VB.Menu MnuOpCrTitulo 
         Caption         =   "Operaciones a Crédito"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuOpCrL0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTOpCredito 
         Caption         =   "Filtrar Operaciones"
         Begin VB.Menu MnuOpCrVigente 
            Caption         =   "&Vigentes"
         End
         Begin VB.Menu MnuOpCrCancelada 
            Caption         =   "&Canceladas"
         End
         Begin VB.Menu MnuOpCrTitular 
            Caption         =   "&Titular"
         End
         Begin VB.Menu MnuOpCrGarantia 
            Caption         =   "&Garantía"
         End
         Begin VB.Menu MnuOpCrL1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuOpCrEliminar 
            Caption         =   "&Eliminar Filtros"
         End
      End
      Begin VB.Menu MnuOL0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpCrDetalleFa 
         Caption         =   "Detalle de &Factura"
      End
      Begin VB.Menu MnuOpCrDetalleOp 
         Caption         =   "Detalle de la &Operación"
      End
      Begin VB.Menu MnuOpCrDetallePa 
         Caption         =   "Detalle de &Pagos"
      End
      Begin VB.Menu MnuOpCrDeudaCh 
         Caption         =   "Deuda en &Cheques"
      End
      Begin VB.Menu MnuOL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpCrCancelar 
         Caption         =   "Cancelar"
      End
   End
   Begin VB.Menu MnuOpContado 
      Caption         =   "OpContado"
      Visible         =   0   'False
      Begin VB.Menu MnuOpCoTitulo 
         Caption         =   "Operaciones Contado"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuOpCoL0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpCoDetalleFa 
         Caption         =   "Detalle de &Factura"
      End
      Begin VB.Menu MnuOpCoDetalleOp 
         Caption         =   "Detalle de la &Operación"
      End
      Begin VB.Menu MnuCoL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpCoCancelar 
         Caption         =   "Cancelar"
      End
   End
   Begin VB.Menu MnuOpVtaT 
      Caption         =   "OpVtaT"
      Visible         =   0   'False
      Begin VB.Menu MnuOpVtaTitulo 
         Caption         =   "Ventas Telefónicas"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuOpVtaL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpVtaDF 
         Caption         =   "Detalle de &Factura"
      End
      Begin VB.Menu MnuOpVtaDO 
         Caption         =   "Detalle de la &Operación"
      End
      Begin VB.Menu MnuOpVtaL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpVtaCancelar 
         Caption         =   "Cancelar"
      End
   End
   Begin VB.Menu MnuOpDireccion 
      Caption         =   "OpDireccion"
      Visible         =   0   'False
      Begin VB.Menu MnuOpDiTitulo 
         Caption         =   "Menú Dirección"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuOpDiL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpDiConfirmar 
         Caption         =   "Confirmar Dirección"
      End
      Begin VB.Menu MnuOpDiModificar 
         Caption         =   "Modificar Dirección"
      End
   End
   Begin VB.Menu MnuPlantilla 
      Caption         =   "MnuPlantilla"
      Visible         =   0   'False
      Begin VB.Menu MnuPlX 
         Caption         =   "MnuPlX"
         Index           =   0
      End
   End
   Begin VB.Menu MnuTelefono 
      Caption         =   "MnuTelefono"
      Visible         =   0   'False
      Begin VB.Menu MnuTLlamoDe 
         Caption         =   "Llamó de ..."
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuTelL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTX 
         Caption         =   "MnuTX"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMCondicion 
      Caption         =   "mnuMCondicion"
      Visible         =   0   'False
      Begin VB.Menu mnuCondicionX 
         Caption         =   "mnuCondicion"
         Index           =   0
      End
      Begin VB.Menu mnuMCX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMCX2 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "frmResolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MC_RES_AUTHASTA = 12

Const Fichas = 7                    'Cantidad de Fichas
Const PorPagina = 20             'Cantidad de Operaciones por Página

Dim RsCli As rdoResultset       'Resultset para cargar el Cliente seleccionado
Dim RsSol As rdoResultset

Dim gIdSolicitud As Long
Dim gCliente As Long, gConyuge As Long, gGarantia As Long, gTipoCliente As Integer

Dim aTexto As String
Dim Articulos As String
Dim aDoc As Long

Dim rsMon As rdoResultset   'Para sacar la Moneda y TC de la Moneda Fija
Dim aMonedaAnteriorC As Long, aMonedaAnteriorT As String       'Codigo y Texto para Hacer menos consultas

'Filtros de operaciones crédito
Dim bFVigentes As Boolean, bFCanceladas As Boolean
Dim bFTitular As Boolean, bFGarante As Boolean

Dim bDevuelta As Boolean

Dim auxEncabezado As String
Dim mValor As Long
Dim aAccionDireccion As Integer     'Para ver q boton de direccion se apretó    (1 Titular, 2 Conyuge)

Dim gAnalizaU As Long
Dim dStartTime As Date

Dim TipoMnu As Integer
Enum MnuCG
    Conyuge = 1
    Garantia = 2
End Enum

Dim WithEvents objClearing As clsClearing
Attribute objClearing.VB_VarHelpID = -1

Public prmVistoPor As Long

Public Property Get prmSolicitud() As Long
    prmSolicitud = gIdSolicitud
End Property
Public Property Let prmSolicitud(Codigo As Long)
    gIdSolicitud = Codigo
End Property

Private Function val_ValidoResolucion(Optional darMensaje As Boolean = True) As Boolean
Dim datos_OK As Boolean

    datos_OK = True
    
    'Valido que se hayan ingresado los parámetros condicionales
    If tGarantia.Enabled And Val(tGarantia.Tag) = 0 Then datos_OK = False
    If tSMonto.Enabled And Not IsNumeric(tSMonto.Text) Then datos_OK = False
    If cRMoneda.Enabled And (cRMoneda.ListIndex = -1 Or Not IsNumeric(tRMonto.Text)) Then datos_OK = False
    If cPCuota.Enabled And cPCuota.ListIndex = -1 Then datos_OK = False
    If tPEntrega.Enabled And Not IsNumeric(tPEntrega.Text) Then datos_OK = False
    
    'Formato de los valores de Comprobante------------------------------------------------------
    If tCValor1.Enabled Then
        If Val(tCValor1.Tag) = 0 And UCase(Trim(lCValor1.Tag)) = "CEDULA" Then
            datos_OK = False
        Else
            If FormatoGrabarReferencia(tCValor1.Text, lCValor1.Tag) = "" Then datos_OK = False
        End If
    End If
    
    If tCValor2.Enabled Then
        If Val(tCValor2.Tag) = 0 And UCase(Trim(lCValor2.Tag)) = "CEDULA" Then
            datos_OK = False
        Else
            If FormatoGrabarReferencia(tCValor2.Text, lCValor2.Tag) = "" Then datos_OK = False
        End If
    End If
    '-------------------------------------------------------------------------------------------------------
    
    val_ValidoResolucion = datos_OK
    If Not darMensaje Then Exit Function
    
    If Not datos_OK Then
        MsgBox "Los datos ingresados no son correctos o están incompletos.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If tSMonto.Enabled And IsNumeric(tSMonto.Text) Then
        If CCur(lSFinanciado.Caption) >= CCur(tSMonto.Text) Then
            fnc_SugerirFinanciacion CCur(tSMonto.Text)
        End If
    End If
    zfn_TextoCondicion
    
End Function


Private Function fnc_SugerirFinanciacion(mImporteFinan As Currency) As Boolean
On Error GoTo errFnc
    Screen.MousePointer = 11
    
    frmHelpMonto.prmIDSolicitud = gIdSolicitud
    frmHelpMonto.prmTotalFinanciado = mImporteFinan
    frmHelpMonto.Show vbModal, Me

    If frmHelpMonto.prmSugerirPlanes <> "" Then
        mImporteFinan = frmHelpMonto.prmTotalFinanciado
        tSMonto.Text = mImporteFinan
        'tComentario.Tag = tComentario.Tag & mImporteFinan & " (" & frmHelpMonto.prmSugerirPlanes & ")"
        If CStr(mImporteFinan) = CStr(frmHelpMonto.prmSugerirPlanes) Then
            tComentario.Tag = lRComentario.Tag & "$" & mImporteFinan
        Else
            tComentario.Tag = lRComentario.Tag & "$" & mImporteFinan & " (" & frmHelpMonto.prmSugerirPlanes & ")"
        End If
    End If
    
errFnc:
    Screen.MousePointer = 0
End Function

Private Sub bClearing_Click()

Dim aTexto As String

    If gCliente = 0 Then Exit Sub
    
    Select Case LCase(TabCPersona.SelectedItem.Key)
        Case "titular"
            'Si titular es empresa y está seleccionado el tab de titular (1)
            If gTipoCliente = TipoCliente.Empresa Then
                MsgBox "No se peuden realizar clearing a empresas", vbInformation, "ATENCIÓN"
                Exit Sub
            End If
        
            aTexto = "el Titular."
        
        Case "garantia"
            If gGarantia = 0 Then
                MsgBox "No hay una garantía seleccionada para realizar el clearing.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            aTexto = "la Garantía."
        
        Case "conyuge"
            If gConyuge = 0 Then
                MsgBox "No hay un cónyuge seleccionado para realizar el clearing.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            aTexto = "el Cónyuge."
    End Select
    
    If IsDate(lClFecha.Caption) Then
        If lClCosto.ForeColor = vbRed Then
            If MsgBox("Confirma realizar nuevo clearing para " & aTexto, vbQuestion + vbYesNo, "REALIZAR CLEARING") = vbNo Then Exit Sub
        End If
    End If
    On Error GoTo errClearing
    
    Screen.MousePointer = 11
    
    Select Case LCase(TabCPersona.SelectedItem.Key)
        Case "titular": objClearing.PedirClearing cBase, gCliente, "S", gIdSolicitud
        Case "garantia": objClearing.PedirClearing cBase, gGarantia, "G", gIdSolicitud
        Case "conyuge": objClearing.PedirClearing cBase, gConyuge, "G", gIdSolicitud
    End Select
    
    Screen.MousePointer = 0
    Exit Sub
errClearing:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al invocar la aplicación para realizar el clearing.", Err.Description
End Sub

Private Sub bCRefrescar_Click()
    
    Select Case LCase(TabCPersona.SelectedItem.Key)
        Case "titular": CargoDatosClienteClearing gCliente
        Case "garantia": CargoDatosClienteClearing gGarantia
        Case "conyuge": CargoDatosClienteClearing gConyuge
    End Select
    
End Sub

Private Sub bDireccionG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    aAccionDireccion = 3
    mValor = Val(Picture1(3).Tag)
    If mValor <> 0 Then PopupMenu MnuOpDireccion, , , , MnuOpDiTitulo
End Sub

Private Sub acc_LlamarA()

    On Error GoTo errLlamarA
    If gCliente = 0 Then Exit Sub
    Dim bOK As Boolean: bOK = False
    
    If Trim(tCodigo.Text) <> "" Then '-----------------------------------------------------------------
        If MsgBox("Confirma 'Llamar a...' " & Trim(lUsuario.Caption) & "?.", vbQuestion + vbYesNo, "LLamar a...") = vbNo Then Exit Sub
        
        Screen.MousePointer = 11
        Cons = "Select * from Solicitud Where SolCodigo = " & Trim(tCodigo.Text)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If IsNull(RsAux!SolFResolucion) And (RsAux!SolEstado = EstadoSolicitud.Pendiente Or RsAux!SolEstado = EstadoSolicitud.ParaRetomar) Then
                RsAux.Edit
                RsAux!SolProceso = TipoResolucionSolicitud.LlamarA
                RsAux.Update
                bOK = True
            Else
                MsgBox "La solicitud ya ha sido resuelta por otro usuario.", vbExclamation, "Solicitud Resuelta"
            End If
        
        Else
            MsgBox "La solicitud seleccionada no existe !!.", vbExclamation, "Posible Error !!"
        End If
        RsAux.Close
    End If
    '-------------------------------------------------------------------------------------------------------
    If bOK Then prmVistoPor = paCodigoDeUsuario
    Screen.MousePointer = 0
    Exit Sub
    
errLlamarA:
    clsGeneral.OcurrioError "Error al grabar los datos para llamar a ...", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function acc_MnuPlantillas()
    If gCliente = 0 Then Exit Function
    PopupMenu MnuPlantilla, , tbTool.Buttons("plantillas").Left, tbTool.Top + tbTool.Height
End Function

Private Function acc_PreguntaGrabarSolicitud(ByVal Titulo As String, ByVal Pregunta As String, ByRef resolucion As clsResolucionAccesible) As VbMsgBoxResult
    acc_PreguntaGrabarSolicitud = vbNo
    Set resolucion = Nothing
    With New frmPregunta
        .Pregunta = Pregunta
        .Titulo = Titulo
        .ShowDialog
        If .DialogResult = vbYes Then
            Set resolucion = .ResolucionAccesible
            acc_PreguntaGrabarSolicitud = vbYes
        End If
    End With
End Function

Private Function acc_ResolverSolicitud(irASiguiente As Boolean)

    'Valido las condiciones a para resolver solicitud
    If gCliente = 0 Then Exit Function
    
    If Val(mnuMCondicion.Tag) = MC_RES_AUTHASTA Then
        If Trim(tSMonto.Text) = "" Or Not IsNumeric(tSMonto.Text) Then
            'Voy a resolver x condicion ST Sí
            mnuMCondicion.Tag = 0
        Else
            If CCur(tSMonto.Text) > CCur(lSFinanciado.Caption) Then
                If MsgBox("El máximo autorizado supera el monto financiado." & Chr(13) & "¿Desea preautorizar el monto y grabar por resolución estandar?", vbQuestion + vbYesNo, "Grabar solicitud") = vbYes Then
                    'Grabo en Tabla de Autorizaciones y Resuelvo x condicion STD
                    GrabarPreAutorizacion CCur(tSMonto.Text)
                    mnuMCondicion.Tag = 0
                Else
                    If Not val_ValidoResolucion(False) Then Exit Function
                End If
            Else
                If Not val_ValidoResolucion(False) Then Exit Function
            End If
        End If
    End If

    Dim oResAccesible As clsResolucionAccesible
    If Val(mnuMCondicion.Tag) = 0 Then
        
        If paResolucionEstandar <> 0 Then   'Pregunto si la desea resolver con la condicion estándar.
            
            mnuMCondicion.Tag = paResolucionEstandar
            If acc_PreguntaGrabarSolicitud("Resolver solicitud", "¿Desea resolver la solicitud con la condición estándar?", oResAccesible) <> vbYes Then
            'If MsgBox("¿Desea resolver la solicitud con la condición estándar?", vbQuestion + vbYesNo, "Resolver Solicitud") = vbNo Then
                mnuMCondicion.Tag = ""
                Exit Function
            End If
            
            ZHabilitoCondiciones CLng(paResolucionEstandar)
            'If val_ValidoResolucion Then AccionGrabar bShowNext:=irASiguiente Else PierdoFoco Label49
            'La estándar es Sí  x lo tanto le saco la validación de las condiciones
            AccionGrabar oResAccesible, bShowNext:=irASiguiente
        Else
            MsgBox "Se debe ingresar el resultado de la solicitud en la ficha 'Resolver'.", vbExclamation, "ATENCIÓN"
        End If
        Exit Function
        
    Else
        If Trim(tComentario.Text) = "" Then tComentario.Text = lRComentario.Tag
    End If
    
    If acc_PreguntaGrabarSolicitud("Resolver solicitud", "Confirma resolver la solictiud con la siguiente condición." & vbCrLf & "Resolución: " & Trim(tComentario.Text), oResAccesible) <> vbYes Then Exit Function
            
'    If MsgBox("Confirma resolver la solictiud con la siguiente condición." & vbCrLf & _
'                   "Resolución: " & Trim(tComentario.Text), vbQuestion + vbYesNo, "Resolver Solicitud") = vbNo Then Exit Function
                   
    AccionGrabar oResAccesible, bShowNext:=irASiguiente
        
End Function
Private Sub AccionGrabarPorPrecondicion(ByVal numero As Long)
    
    Screen.MousePointer = 11
    On Error GoTo errorBT
    '[dbo].[prg_AprobarSolicitudPreAutorizada] @solicitud int, @nroresolucion tinyint, @usuario smallint
    cBase.Execute "EXEC prg_AprobarSolicitudPreAutorizada " & gIdSolicitud & ", " & numero & ", " & paCodigoDeUsuario
    
    LimpioFicha True
    tCodigo.Text = "": gIdSolicitud = 0: gCliente = 0: gGarantia = 0: gConyuge = 0
    If FormActivo("frmLista") Then
        
        If frmLista.WindowState = vbMinimized Then frmLista.WindowState = vbNormal
        If frmLista.Visible Then frmLista.SetFocus
        
    End If
    Screen.MousePointer = 0
    Exit Sub

errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al intentar resolver la solicitud.", Err.Description
    Exit Sub

End Sub


Private Sub AccionGrabar(ByVal ResoluccionAccesible As clsResolucionAccesible, Optional bShowNext As Boolean = False)
Dim aMsgError As String
Dim rs1 As rdoResultset

    Screen.MousePointer = 11
    aMsgError = ""
    On Error GoTo errorBT
    
    Dim condicionRes As Byte
    condicionRes = 0
    
    Dim m_KeyResolucion As String, m_IDResolucion As Long, m_IDEstadoSolicitud As Integer, m_QuedaPendiente As Boolean
    m_IDResolucion = Val(mnuMCondicion.Tag)
    
    m_KeyResolucion = zfn_ArmoKeyCondicion(m_IDResolucion, condicionRes)
    m_IDEstadoSolicitud = Val(frmCondicion.Tag)
    If m_IDEstadoSolicitud = 0 Then
        MsgBox "El estado de la solicitud puede no ser correcto. " & vbCrLf & _
                    "Verifique que seleccionó una condición para resolver la solicitud.", vbExclamation, "Posible Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    m_QuedaPendiente = (cRDefinitiva.value = vbUnchecked)
    If prmAutorizaCredHasta <> -1 And (CCur(lFMonto.Caption) > prmAutorizaCredHasta) And condicionRes <> 2 Then m_QuedaPendiente = True
        
    FechaDelServidor
    
    'Selecciono la solicitud para ver si aún no ha sido resuelta
    Cons = "Select * from Solicitud Where SolCodigo = " & gIdSolicitud
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
    
        If RsAux!SolEstado <> EstadoSolicitud.Pendiente Then
            'Solicitud ya resuelta---------------------------------------------
            Screen.MousePointer = 0
            If Not IsNull(RsAux!SolUsuarioR) Then
                MsgBox "La solicitud ha sido resuelta por otro usuario (" & z_BuscoUsuario(RsAux!SolUsuarioR, Identificacion:=True) & ").", vbExclamation, "Solicitud Resuelta"
            Else
                MsgBox "La solicitud ha sido resuelta por otro usuario.", vbCritical, "ERROR"
            End If
        
        Else        'Solicitud SIGUE PENDIENTE
            
            'Valido si se hizo clearing HOY ---------------------------------------------------------------
            Dim bHizoClearing As Boolean: bHizoClearing = False
            Dim rsVCl As rdoResultset, miSql As String
            miSql = "Select * from Clearing " & _
                        " Where CleCliente = " & gCliente & _
                        " And CleFecha > '" & Format(gFechaServidor, "mm/dd/yyyy 00:00") & "'"
            Set rsVCl = cBase.OpenResultset(miSql, rdOpenDynamic, rdConcurValues)
            If Not rsVCl.EOF Then bHizoClearing = True
            rsVCl.Close
            '---------------------------------------------------------------------------------------------------
            
            cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
            On Error GoTo errorET
            RsAux.Requery
            
            If RsAux!SolEstado <> EstadoSolicitud.Pendiente Then        'Valido con RS Bloqueado
                aMsgError = "La solicitud ha sido resuelta por otro usuario."
                RsAux.Close
                GoTo errorET: Exit Sub
            End If  '----------------------------------------------------------------------------------------
            
            'Cambio el estado de la Solicitud
            RsAux.Edit
            RsAux!SolProceso = TipoResolucionSolicitud.Manual
            
            RsAux!SolEstado = IIf(m_QuedaPendiente, EstadoSolicitud.Pendiente, m_IDEstadoSolicitud)
            RsAux!SolFResolucion = IIf(m_QuedaPendiente, Null, Format(gFechaServidor, sqlFormatoFH))
            RsAux!SolUsuarioR = IIf(m_QuedaPendiente, Null, paCodigoDeUsuario)
            RsAux!SolCondicionR = IIf(m_QuedaPendiente, Null, m_IDResolucion)
        
            If bHizoClearing Then RsAux!SolSeHizoClearing = 1
            
            Dim tmSeg As Long
            If Not IsNull(RsAux!SolTiempoResolucion) Then tmSeg = RsAux!SolTiempoResolucion Else tmSeg = 0
            tmSeg = tmSeg + DateDiff("s", dStartTime, gFechaServidor)
            If tmSeg > 32000 Then tmSeg = 32000
            RsAux!SolTiempoResolucion = tmSeg
            RsAux.Update
            
            
            Dim m_Numero As Byte
            
            Cons = "Select Top 1 * from SolicitudResolucion Where ResSolicitud = " & gIdSolicitud & " Order by ResNumero DESC"
            Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If Not rs1.EOF Then m_Numero = rs1!ResNumero Else m_Numero = 0
            m_Numero = m_Numero + 1
            
            rs1.AddNew
            rs1!ResSolicitud = gIdSolicitud
            rs1!ResNumero = m_Numero
            rs1!ResCondicion = m_IDResolucion
            rs1!ResTexto = Trim(IIf(m_KeyResolucion <> "", m_KeyResolucion, Null))
            rs1!ResComentario = IIf(Trim(tComentario.Text) <> "", Trim(tComentario.Text), Null)
            rs1!ResFecha = Format(gFechaServidor, sqlFormatoFH)
            rs1!ResUsuario = paCodigoDeUsuario
            rs1.Update
            
            rs1.Close
            
            If Not ResoluccionAccesible Is Nothing Then
                Cons = ""
                If ResoluccionAccesible.ID > 0 Then
                    Cons = "UPDATE Codigos SET CodValor1 = IsNull(CodValor1, 0) + 1 WHERE CodCual = 166 AND CodID = " & ResoluccionAccesible.ID
                ElseIf ResoluccionAccesible.Texto <> "" Then
                    Cons = "INSERT INTO Codigos (CodCual, CodId, CodTexto, CodValor1) values (166, (Select IsNull(MAX(CodID), 0) From Codigos Where CodCual = 166)+1, '" & ResoluccionAccesible.Texto & "', 1)"
                End If
                If Cons <> "" Then cBase.Execute Cons
            End If
            
            cBase.CommitTrans   'FINALIZO TRANSACCION -------------------------------------------
            RsAux.Requery
            
            frmLista.signalR_RefrescoSolicitudesResueltas
            frmLista.signalR_RefrescoSolicitudes
            
            LimpioFicha True
            tCodigo.Text = "": gIdSolicitud = 0: gCliente = 0: gGarantia = 0: gConyuge = 0
            
            If FormActivo("frmLista") Then
                
                
                Dim bDo As Boolean: bDo = True
                If bShowNext Then
                    If frmLista.acc_SiguienteAsunto(mTipoAsunto:=Asuntos.solicitudes) = 1 Then bDo = False
                End If
                If bDo Then
                    If frmLista.WindowState = vbMinimized Then frmLista.WindowState = vbNormal
                    If frmLista.Visible Then frmLista.SetFocus
                End If
                
            End If
            
        End If
        
    Else
        Screen.MousePointer = 0
        MsgBox "La solicitud ha sido eliminada. Verifique la lista de solicitudes pendientes.", vbCritical, "ERROR"
    End If
    
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub

errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    If Trim(aMsgError) = "" Then aMsgError = "No se ha podido realizar la transacción. Reintente la operación."
    clsGeneral.OcurrioError aMsgError, Err.Description
    Exit Sub
End Sub

Private Function acc_SoltarSolicitud()

    On Error GoTo errActualizar
    If gCliente = 0 Then Exit Function
    Dim bOK As Boolean: bOK = False
    
    If Trim(tCodigo.Text) <> "" Then '-----------------------------------------------------------------
        If MsgBox("Confirma liberar la solicitud y marcarla para reestudiar ?.", vbQuestion + vbYesNo, "Soltar Para Retomar") = vbNo Then Exit Function
        Screen.MousePointer = 11
        
        FechaDelServidor
        
        Cons = "Select * from Solicitud Where SolCodigo = " & Trim(tCodigo.Text)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If IsNull(RsAux!SolFResolucion) And (RsAux!SolEstado = EstadoSolicitud.Pendiente Or RsAux!SolEstado = EstadoSolicitud.ParaRetomar) Then
                RsAux.Edit
                RsAux!SolUsuarioR = paCodigoDeUsuario
                RsAux!SolEstado = EstadoSolicitud.ParaRetomar
                
                Dim tmSeg As Long
                If Not IsNull(RsAux!SolTiempoResolucion) Then tmSeg = RsAux!SolTiempoResolucion Else tmSeg = 0
                tmSeg = tmSeg + DateDiff("s", dStartTime, gFechaServidor)
                If tmSeg > 32000 Then tmSeg = 32000
                RsAux!SolTiempoResolucion = tmSeg
            
                RsAux.Update
                bOK = True
            Else
                MsgBox "La solicitud ya ha sido resuelta por otro usuario.", vbExclamation, "Solicitud Resuelta"
            End If
        
        Else
            MsgBox "La solicitud seleccionada no existe !!.", vbExclamation, "Posible Error !!"
        End If
        RsAux.Close
    End If
    '-------------------------------------------------------------------------------------------------------
    
    If bOK Then
        LimpioFicha True
        tCodigo.Text = "": gIdSolicitud = 0: gCliente = 0: gGarantia = 0: gConyuge = 0
        frmLista.signalR_RefrescoSolicitudes
        If FormActivo("frmLista") Then
            If frmLista.WindowState = vbMinimized Then frmLista.WindowState = vbNormal
            If frmLista.Visible Then frmLista.SetFocus
        End If
    End If
    
    Screen.MousePointer = 0
    Exit Function
    
errActualizar:
    clsGeneral.OcurrioError "Error al liberar la solicitud.", Err.Description
    Screen.MousePointer = 0
End Function


Private Sub cCondicionR_Click()
    
    If Me.ActiveControl.Name = "cCondicionR" Then
        ZDeshabilitoCondiciones
    End If
          
End Sub

Private Sub cCondicionR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(mnuMCondicion.Tag) <> Val(cCondicionR.ItemData(cCondicionR.ListIndex)) Then
            mnuMCondicion.Tag = Val(cCondicionR.ItemData(cCondicionR.ListIndex))
            ZHabilitoCondiciones Val(mnuMCondicion.Tag)
        
            PierdoFoco cCondicionR 'FocotComentario.SetFocus
        Else
            If val_ValidoResolucion(darMensaje:=False) Then tComentario.SetFocus Else PierdoFoco cCondicionR
        End If
    End If

End Sub

Private Sub cCondicionR_LostFocus()
    On Error Resume Next
    Call cCondicionR_KeyPress(vbKeyReturn)
End Sub

Private Sub cDireccion_Click()
On Error GoTo errCargar

    If cDireccion.ListIndex <> -1 Then
        Screen.MousePointer = 11
        lDireccion.Caption = ""
        lDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, cDireccion.ItemData(cDireccion.ListIndex), Departamento:=True, Localidad:=True, Zona:=True, ConfyVD:=True)
        If InStr(lDireccion.Caption, "(Cf.)") <> 0 Then lDireccion.ForeColor = Colores.osVerde Else lDireccion.ForeColor = Colores.Azul
        Screen.MousePointer = 0
    End If

errCargar:
    Screen.MousePointer = 0
End Sub

Private Sub cDireccionG_Click()
On Error GoTo errCargar

    If cDireccionG.ListIndex <> -1 Then
        Screen.MousePointer = 11
        lDireccionG.Caption = ""
        lDireccionG.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, cDireccionG.ItemData(cDireccionG.ListIndex), Departamento:=True, Localidad:=True, Zona:=True, ConfyVD:=True)
        If InStr(lDireccionG.Caption, "(Cf.)") <> 0 Then lDireccionG.ForeColor = Colores.osVerde Else lDireccionG.ForeColor = Colores.Azul
        Screen.MousePointer = 0
    End If

errCargar:
    Screen.MousePointer = 0
End Sub

Private Sub cPCuota_Change()
    tPEntrega.Text = ""
    tPEntrega.Enabled = False: tPEntrega.BackColor = vbButtonFace
End Sub

Private Sub cPCuota_Click()
    tPEntrega.Text = ""
    tPEntrega.Enabled = False: tPEntrega.BackColor = vbButtonFace
End Sub

Private Sub cPCuota_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        'Valido si el plan tiene entrega para habilitar
        If cPCuota.ListIndex = -1 Then Exit Sub
        
        Cons = "Select * from TipoCuota Where TCuCodigo = " & cPCuota.ItemData(cPCuota.ListIndex)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If RsAux.EOF Then RsAux.Close: Exit Sub
                
        If Not IsNull(RsAux!TCuVencimientoE) Then tPEntrega.Enabled = True: tPEntrega.BackColor = Obligatorio
        RsAux.Close
        
        If tPEntrega.Enabled Then Foco tPEntrega Else PierdoFoco cPCuota
    End If
    
End Sub

Private Sub cRMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tRMonto
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
        
        Case vbKeyF2:
                If gCliente = 0 Then Exit Sub
                AccionMenuFicha gCliente, gTipoCliente
        
        Case vbKeyF12:
                If gCliente <> 0 Then EjecutarApp pathApp & "Visualizacion de operaciones.exe", CStr(gCliente)
                
        Case Else
                If Shift = vbCtrlMask Then
                    Select Case KeyCode
                        Case vbKeyR: acc_ResolverSolicitud irASiguiente:=True
                        Case vbKeyS: acc_SoltarSolicitud
                        Case vbKeyX: Unload Me
                        Case vbKeyL: acc_LlamarA
                        Case vbKeyP: acc_MnuPlantillas
                    End Select
                End If
    End Select
    
End Sub

Private Sub Form_Load()
On Error Resume Next

    Set objClearing = New clsClearing
    
    cEMailT.OpenControl cBase
    cEMailT.IDUsuario = paCodigoDeUsuario
    
    ObtengoSeteoForm Me

    '---------------------------------------------------------------------------------------------------
    With tbTool
        .Buttons("autor").Image = img2.ListImages("auto").Index
        .Buttons("autor").ToolTipText = "Ver Resolución Automática."
        .Buttons("resolverysig").Image = img2.ListImages("resolver").Index
        .Buttons("resolverysig").ToolTipText = "Resolver y Siguiente (Ctrl+R)."
        .Buttons("resolver").Image = img2.ListImages("resolver2").Index
        .Buttons("resolver").ToolTipText = "Resolver Solicitud."
        
        .Buttons("soltar").Image = img2.ListImages("soltar").Index
        .Buttons("soltar").ToolTipText = "Soltar para Retomar (Ctrl+S)."
        .Buttons("llamara").Image = img2.ListImages("llamar").Index
        .Buttons("llamara").ToolTipText = "Llamar a ...(Ctrl+L)."
        
        .Buttons("plantillas").Image = img2.ListImages("plantillas").Index
        .Buttons("plantillas").ToolTipText = "Plantillas Interactivas (Ctrl+P)."
        
        .Buttons("salir").Image = img2.ListImages("cerrar").Index
        .Buttons("salir").ToolTipText = "Cancelar Resolución (Ctrl+X).."
    End With
    
    '---------------------------------------------------------------------------------------------------
    ImageList1.UseMaskColor = True
    Set Tab1.ImageList = ImageList1
    Tab1.Tabs("operaciones").Image = ImageList1.ListImages("operaciones").Index
    Tab1.Tabs("solicitud").Image = ImageList1.ListImages("solicitud").Index
    Tab1.Tabs("titular").Image = ImageList1.ListImages("titular").Index
    Tab1.Tabs("clearing").Image = ImageList1.ListImages("clearing0").Index
    Tab1.Tabs("resolver").Image = ImageList1.ListImages("resolver").Index
    Tab1.Tabs("relaciones").Image = ImageList1.ListImages("relacionesno").Index
    Tab1.Tabs("garantia").Image = ImageList1.ListImages("garantiano").Index
    
    Set TabPersona.ImageList = ImageList1
    TabPersona.Tabs("titular").Image = ImageList1.ListImages("titular").Index
        
    Set TabClearing.ImageList = ImageList1
        
    With TabCPersona
        Set .ImageList = ImageList1
        .Tabs.Clear
        .Tabs.Add pvKey:="titular", pvCaption:="Titular", pvImage:=ImageList1.ListImages("titular").Index
    End With
    '---------------------------------------------------------------------------------------------------
    
    FechaDelServidor
    Call Form_Resize        'Para evitar ajustar los obj. cdo esta visible
    
    'Filtros en Falso, cargo todas
    bFVigentes = False: bFCanceladas = False: bFTitular = False: bFGarante = False
    InicializoGrillas
    
    Picture1(0).ZOrder 0: pClearing(0).ZOrder 0
    
    auxEncabezado = ""
    EncabezadoOperaciones Credito:=True
    '----------------------------------------------------------------------------------------------------
    frmResolucion.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    frmResolucion.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)

    On Error GoTo errXX
    
    'Cargo los TAGs de caga ícono con los items del menú    ------------------------------------------------------
    Dim xItm As Integer, rsX As rdoResultset
    
    Cons = "Select ConCodigo, ConNombre, ConGrupoIcono From CondicionResolucion " & _
               " Where (ConNoHabilitado Is Null Or ConNoHabilitado = 0) " & _
               " Order by ConGrupoIcono, ConGrupoOrden, ConNombre"
    Set rsX = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsX.EOF
        If Not IsNull(rsX!ConGrupoIcono) Then xItm = rsX!ConGrupoIcono - 1 Else xItm = 99
        If xItm > imgCondRes.UBound Then xItm = imgCondRes.UBound
        
        Cons = imgCondRes(xItm).Tag
        imgCondRes(xItm).Tag = Cons & IIf(Cons = "", "", "·") & Trim(rsX!ConCodigo) & "|" & Trim(rsX!ConNombre)
        
        cCondicionR.AddItem Trim(rsX!ConNombre)
        cCondicionR.ItemData(cCondicionR.NewIndex) = rsX!ConCodigo
         
        rsX.MoveNext
    Loop
    rsX.Close
    '-----------------------------------------------------------------------------------------------------------------------
    
    If gIdSolicitud <> 0 Then
        'Dado que se están presentando casos donde 2 usuarios tienen la misma solicitud consulto para ver si sigue siendo
        'el mismo usuario que la tomo en 1er lugar.
        SiguienteSolicitud
        frmLista.signalR_RefrescoSolicitudes
        If gIdSolicitud = 0 Then
            Unload Me
        Else
            BuscoSolicitud gIdSolicitud
        End If
    Else
        LimpioFicha
    End If
    Exit Sub
    
errXX:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar las condiciones de resolución.", Err.Description
End Sub

Sub SiguienteSolicitud()
On Error GoTo errSS
Dim rsS As rdoResultset
    Cons = "EXEC prg_ProximaSolicitudAResolver " & paCodigoDeUsuario & ", " & gIdSolicitud
    
    
    Cons = "SELECT TOP 1 (Case When SolCodigo = " & gIdSolicitud & " Then 0 Else SolCodigo End) NroSolicitud " & _
            "From Solicitud Where SolFecha > CONVERT(DateTime, Floor(CONVERT(Float, getdate()))) " & _
            "AND (SolUsuarioR Is Null OR (SolEstado = 4 And SolUsuarioR IS NOT NULL) " & _
            "OR (SolEstado = 0 And SolUsuarioR = " & paCodigoDeUsuario & ") ) " & _
            "ORDER BY (Case When SolCodigo = " & gIdSolicitud & " Then 0 Else SolCodigo End)"

    
    Set rsS = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsS.EOF Then
        'Si es la que yo le envíe
        If 0 = rsS("NroSolicitud") Then
            'Es la misma sigo de largo
        ElseIf gIdSolicitud <> rsS("NroSolicitud") Then
        
            If MsgBox("La solicitud a la cual accedió está siendo analizada por otro usuario." & vbCrLf & vbCrLf & "¿Desea visualizar dicha solicitud o acceder a la primera disponible?" & vbCrLf & vbCrLf _
                & "Presione: " & vbCrLf & vbTab & " <SI> visualizar esta solicitud de todas formas" & vbCrLf & vbTab _
                & "<NO> para visualizar la primer solicitud disponible", vbQuestion + vbYesNo + vbDefaultButton2, "Solicitud ocupada") = vbYes Then
                'No hago nada dejo que siga tal cual está.
            Else
                gIdSolicitud = rsS("NroSolicitud")
                'La marco como tomada.
                Dim usuRet As Long
                Select Case fnc_BlockearSolicitud(gIdSolicitud, usuRet)
                    Case 0  'OTRO USUARIO
                         If usuRet <> paCodigoDeUsuario Then
                             Screen.MousePointer = 0
                             If MsgBox("La solicitud se está analizando por otro usuario." & vbCrLf & _
                                            "Desea visualizarla.", vbExclamation + vbYesNo, "Solicitud en Proceso") = vbNo Then
                                            gIdSolicitud = 0
                                            Exit Sub
                                        End If
                         End If
                    
                     Case -1 'ERROR o FUE RESUELTA
                         Screen.MousePointer = 0
                         gIdSolicitud = 0
                         MsgBox "Posiblemente la solicitud ya fue resuelta (o ha sido eliminada).", vbExclamation, "Cambiaron los Datos"
                         Exit Sub
                End Select  '----------------------------------------------------------------------------------------
            End If
        Else
            'es la misma no hago nada.
        End If
    Else
        MsgBox "La solicitud ya fue resuelta o no hay solicitudes disponibles.", vbExclamation, "ATENCIÖN"
        gIdSolicitud = 0
    End If
    rsS.Close
    Exit Sub
errSS:
    clsGeneral.OcurrioError "Error al validar si la solicitud fue tomada por otro usuario.", Err.Description
End Sub

Public Sub BuscoSolicitud(Codigo As Long)

    Screen.MousePointer = 11
    dStartTime = gFechaServidor
    Dim aIDCliente As Long: aIDCliente = gCliente
    LimpioFicha True
    gCliente = aIDCliente
    
    Cons = "Select * from Solicitud, Moneda " _
            & " Where SolCodigo = " & Codigo _
            & " And SolMoneda = MonCodigo"
    Set RsSol = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsSol.EOF Then
        tCodigo.Text = RsSol!SolCodigo
        
        CargoDatosCliente RsSol!SolCliente
        
        lMoneda.Caption = Trim(RsSol!MonSigno): lMoneda.Tag = RsSol!SolMoneda
        lMoneda2.Caption = lMoneda.Caption

        'Cargo Datos de Financiacion---------------------------------
        CargoFinanciacion Codigo
        If RsSol!SolFormaPago = TipoPagoSolicitud.ChequeDiferido Then lFinanciacion.Caption = "C/Ch " & Trim(lFinanciacion.Caption)

        If Not IsNull(RsSol!SolGarantia) Then gGarantia = RsSol!SolGarantia
        
        lFecha.Caption = Format(RsSol!SolFecha, "d-Mmm-yyyy hh:mm")
        
        If RsSol!SolFormaPago = TipoPagoSolicitud.Efectivo Then lPago.Caption = "Efectivo" Else lPago.Caption = "Cheque Dif."
        
        Cons = "Select * from Sucursal Where SucCodigo = " & RsSol!SolSucursal
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then lSucursal.Caption = Trim(RsAux!SucAbreviacion)
        RsAux.Close
        
        If Not IsNull(RsSol!SolComentarioS) Then lComentario.Caption = Trim(RsSol!SolComentarioS): lAFinanciar.Caption = Trim(RsSol!SolComentarioS) & "  "

        lUsuario.Caption = z_BuscoUsuario(RsSol!SolUsuarioS, Identificacion:=True)
        
        gCliente = RsSol!SolCliente
        If Not IsNull(RsSol!SolUsuarioR) Then gAnalizaU = RsSol!SolUsuarioR
       
        'Comentario de Solicitid Devuelta
         bDevuelta = False
        If Not IsNull(RsSol!SolDevuelta) Then
            If RsSol!SolDevuelta Then bDevuelta = True
            If RsSol!SolDevuelta And Not IsNull(RsSol!SolComentarioS) Then
                Me.Caption = Me.Caption & " (" & Trim(RsSol!SolComentarioS) & ")"
            Else
                Me.Caption = "Estudio de Solicitudes"
            End If
        End If
        '--------------------------------------------
        'If Not IsNull(RsSol!SolResolAuto) Then lResAuto.Tag = RsSol!SolResolAuto
        Me.Refresh
    End If
    RsSol.Close

    If InStr(paCatsDistribuidor, "," & lCategoria.Tag & ",") = 0 Then       'Cargo los datos en OPERACIONES (si no es Distribuidor)
        CargoCantidadOperaciones gCliente
        If oCredito.Tag <> 0 Then
            CargoOpCredito gCliente
            Me.Refresh
            CargoDeudaCliente gCliente
        Else
            If oContado.Tag <> 0 Then oContado.value = True
        End If
    Else
        Tab1.Tabs("solicitud").Selected = True
    End If
        
    'Cargo datos de las resoluciones anteriores     ----------------------------------------------------------------------------
    Cons = "Select SolicitudResolucion.*, UsuIdentificacion, ConNombre " & _
                " From SolicitudResolucion " & _
                    " Left Outer Join Usuario ON ResUsuario = UsuCodigo " & _
                    " Left Outer Join CondicionResolucion ON ResCondicion = ConCodigo " & _
            " Where ResSolicitud = " & Codigo & _
            " Order by ResNumero DESC"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With lCondicion
            .AddItem "" ' |^Fecha|<Visto por ...|<Resolución
            .Cell(flexcpText, .Rows - 1, 0) = RsAux!ResNumero: .Cell(flexcpFontBold, .Rows - 1, 0) = True
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ResFecha, "d/mm hh:mm")
            If Not IsNull(RsAux!UsuIdentificacion) Then
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!UsuIdentificacion)
            Else
                .Cell(flexcpText, .Rows - 1, 2) = "Res.Autom."
            End If
            If Not IsNull(RsAux!ResComentario) Then
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ResComentario)
            Else
                If Not IsNull(RsAux!ConNombre) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ConNombre)
            End If
            If Not IsNull(RsAux("ResTexto")) Then .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & ".  " & Trim(RsAux("ResTexto"))
            
            Dim lID As Long
            lID = RsAux("ResNumero"): .Cell(flexcpData, .Rows - 1, 1) = lID
            lID = RsAux("ResCondicion"): .Cell(flexcpData, .Rows - 1, 2) = lID
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    '---------------------------------------------------------------------------------------------------------------------------------
            
    If gGarantia <> 0 And gTipoCliente = TipoCliente.Empresa Then
        Cons = "Select * from CPersona Where CPeCliente = " & gGarantia
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux!CPeConyuge) Then gConyuge = RsAux!CPeConyuge
            If Not IsNull(RsAux!CPeEstadoCivil) Then lECivilG.Tag = RsAux!CPeEstadoCivil
        End If
        RsAux.Close
    End If
    HabilitoTabs
    
    If lCondicion.Rows > lCondicion.FixedRows Then
        Tab1.Tabs("resolver").Selected = True
    End If
    
    Screen.MousePointer = 0
    If gCliente <> 0 Then
        ProcesoActivacionAutomatica
        Me.Show: DoEvents
        BuscoComentariosAlerta gCliente
        
        code_ProcesoScript "APE01", gCliente, Codigo
        
        'X defecto pongo Autorizado Hasta   -----------------------------------
        tComentario.Text = "": tComentario.Tag = ""
        lRComentario.Tag = ""
        BuscoCodigoEnCombo cCondicionR, MC_RES_AUTHASTA 'Autorizado hasta
        ZDeshabilitoCondiciones
        mnuMCondicion.Tag = MC_RES_AUTHASTA
        ZHabilitoCondiciones MC_RES_AUTHASTA
        '----------------------------------------------------------------------
    End If
    
    Exit Sub
    
errSolicitud:
    clsGeneral.OcurrioError "Error al cargar los datos de la solicitud.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoComentariosAlerta(aIDCliente As Long)

    On Error GoTo errBCA
    Dim rsCA As rdoResultset, bHay As Boolean
    
    Cons = " Select * from Comentario, TipoComentario" & _
               " Where ComCliente = " & aIDCliente & _
               " And ComTipo = TCoCodigo" & _
               " And TCoAccion IN ( " & Accion.Cuota & "," & Accion.Decision & "," & Accion.Alerta & ")"
    Set rsCA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsCA.EOF Then bHay = False Else bHay = True
    rsCA.Close

    If bHay Then
        Me.Show
        AccionMenuComentario aIDCliente
    End If
    
errBCA:
End Sub

Public Sub BuscoClienteSeleccionado(Codigo As Long)

Dim aCliente As Long
    Screen.MousePointer = 11
    LimpioFicha True
    aCliente = Codigo
    LimpioFicha True
    Codigo = aCliente: gCliente = Codigo
    CargoDatosCliente Codigo 'Cargo Datos del Cliente Seleccionado
    
    'Cargo los datos en OPERACIONES
    'Picture1(0).Tag = "OK"
    CargoCantidadOperaciones Codigo
    
    If oCredito.Tag <> 0 Then
        CargoOpCredito Codigo     'TITULAR
        Me.Refresh
        CargoDeudaCliente Codigo
    Else
        If oContado.Tag <> 0 Then oContado.value = True
    End If
    
    HabilitoTabs
    
    If Codigo <> 0 Then ProcesoActivacionAutomatica
        
    Screen.MousePointer = 0
    Exit Sub
    
errSolicitud:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de la solicitud.", Err.Description
End Sub

Private Sub CargoDatosCliente(Titular As Long)

    On Error GoTo errCliente
    cDireccion.Clear: cDireccion.BackColor = Colores.Obligatorio: cDireccion.Tag = 0
    
    'Cargo Datos Tabla Cliente----------------------------------------------------------------------
'    Cons = "Select * from Cliente, CategoriaCliente, EmailDireccion" _
'           & " Where CliCodigo = " & Titular _
'           & " And CliCategoria *= CClCodigo" _
'           & " And CliCodigo *= EMDIdCliente"
           
    Cons = "SELECT * " _
        & "From Cliente " _
        & "LEFT OUTER JOIN CategoriaCliente ON CLiCategoria = CClCodigo " _
        & "LEFT OUTER JOIN EMailDireccion ON CliCodigo = EMDIdCliente " _
        & "WHERE CliCodigo = " & Titular
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If Not IsNull(RsAux!CliCiRuc) Then      'CI o RUC
        PresentoPaisDelDocumentoCliente RsAux("CliPaisDelDocumento"), RsAux("CliCiRuc"), lblInfoCI, lCiRuc
'        Select Case RsAux!CliTipo
'            Case TipoCliente.Cliente: lCiRuc.Caption = clsGeneral.RetornoFormatoCedula(RsAux!CliCiRuc)
'            Case TipoCliente.Empresa: lCiRuc.Caption = Trim(RsAux!CliCiRuc)
'        End Select
    Else
        PresentoPaisDelDocumentoCliente RsAux("CliPaisDelDocumento"), "", lblInfoCI, lCiRuc
        'lCiRuc.Caption = "N/D"
    End If
    
    'Direccion
    If Not IsNull(RsAux!CliDireccion) Then
        cDireccion.AddItem "Dirección Principal": cDireccion.ItemData(cDireccion.NewIndex) = RsAux!CliDireccion
        cDireccion.Tag = RsAux!CliDireccion
    End If
    
    If Not IsNull(RsAux!CliAlta) Then
        If Format(RsAux!CliAlta, "dd/mm/yyyy") = Format(gFechaServidor, "dd/mm/yyyy") Then
            shpTitular.BackColor = Colores.clVerde
        Else
            shpTitular.BackColor = Colores.Obligatorio
        End If
        shpTitular.Refresh
    End If
   
    If Not IsNull(RsAux!CliCategoria) Then
        If RsAux!CliCategoria = paCatCliFallecido Then shpTitular.BackColor = Colores.Gris
        If RsAux!CliCategoria <> paCategoriaCliente Then lCategoria.ForeColor = vbRed
    End If
    cEMailT.BackColor = shpTitular.BackColor
    
    Dim bLlamoDe As Boolean
    lTelefono.Caption = TelefonoATexto(Titular, bLlamoDe)     'Telefonos
    If bLlamoDe Then lTelefono.Tag = 1 Else lTelefono.Tag = 0
    
    gTipoCliente = RsAux!CliTipo
    If Not IsNull(RsAux!CClNombre) Then lCategoria.Caption = Trim(RsAux!CClNombre)
    If Not IsNull(RsAux!CliCategoria) Then lCategoria.Tag = RsAux!CliCategoria Else lCategoria.Tag = 0
    
    If Not IsNull(RsAux!CliSolicitud) Then lNSolicitud.Caption = Trim(RsAux!CliSolicitud)
    If Not IsNull(RsAux!CliCheque) Then If UCase(RsAux!CliCheque) = "S" Then lCheques.Caption = "Opera c/cheques"
    
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------
    
    cEMailT.CargarDatos Titular
    
    If Val(cDireccion.Tag) <> 0 Then BuscoCodigoEnCombo cDireccion, Val(cDireccion.Tag)
    CargoDireccionesAuxiliares cDireccion, Titular
    
    'Cargo Datos Tabla CPersona o CEmpresa------------------------------------------------------
    If gTipoCliente = TipoCliente.Cliente Then
'        Cons = "Select * from CPersona, Ocupacion, EstadoCivil, InquilinoPropietario " _
'               & " Where CPeCliente = " & Titular _
'               & " And CPeOcupacion *= OcuCodigo  " _
'               & " And CPeEstadoCivil *= ECiCodigo" _
'               & " And CPePropietario *= IPrCodigo"
'
        Cons = "SELECT * FROM CPersona " _
            & "LEFT OUTER JOIN Ocupacion ON CPeOcupacion = OcuCodigo " _
            & "LEFT OUTER JOIN EstadoCivil ON CPeEstadoCivil = ECiCodigo " _
            & "LEFT OUTER JOIN InquilinoPropietario  ON CPePropietario = IPrCodigo " _
            & "WHERE CPeCliente =" & Titular
               
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        
        lTitular.Caption = ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
        If Not IsNull(RsAux!CPeFNacimiento) Then
            Dim edad As Byte
            If DateAdd("yyyy", DateDiff("yyyy", RsAux("CPeFNacimiento"), Date), RsAux("CPeFNacimiento")) > Date Then
                edad = DateDiff("yyyy", RsAux("CPeFNacimiento"), Date) - 1
            Else
                edad = DateDiff("yyyy", RsAux("CPeFNacimiento"), Date)
            End If
            lEdad.Caption = edad & Format(RsAux!CPeFNacimiento, " (d-Mmm-yyyy)") '((Date - RsAux!CPeFNacimiento) \ 365) & Format(RsAux!CPeFNacimiento, " (d-Mmm-yyyy)")
            'If ((Date - RsAux!CPeFNacimiento) \ 365) > paMayorEdad1 Then lEdad.ForeColor = vbRed
            If edad > paMayorEdad1 Then lEdad.ForeColor = vbRed
        End If
        If Not IsNull(RsAux!OcuNombre) Then lOcupacion.Caption = Trim(RsAux!OcuNombre)
        
        lECivil.Tag = 0
        If Not IsNull(RsAux!ECiNombre) Then
            lECivil.Caption = Trim(RsAux!ECiNombre)
            lECivil.Tag = RsAux!ECiCodigo
        End If
        
        If Not IsNull(RsAux!CPeConyuge) Then gConyuge = RsAux!CPeConyuge
        
        If Not IsNull(RsAux!IPrNombre) And Val(cDireccion.Tag) <> 0 Then
            If cDireccion.ListCount > 0 Then cDireccion.List(0) = Trim(RsAux!IPrNombre)
            lNDireccion.Caption = Trim(RsAux!IPrNombre) & ":"
        End If
        
        lDireccion.Refresh
        RsAux.Close
    
    Else    'Empresa
'        Cons = "Select * from CEmpresa, Ramo " _
'               & " Where CEmCliente = " & Titular _
'               & " And CEmRamo *= RamCodigo"
               
        Cons = "SELECT CEmFantasia, CEmNombre, RamNombre " & _
                "FROM CEmpresa LEFT OUTER JOIN Ramo ON CEmRamo = RamCodigo " _
               & "WHERE CEmCliente = " & Titular
               
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        aTexto = Trim(RsAux!CEmFantasia)
        If Not IsNull(RsAux!CEmNombre) Then aTexto = aTexto & " (" & Trim(RsAux!CEmNombre) & ")"
        lTitular.Caption = aTexto
        If Not IsNull(RsAux!RamNombre) Then lOcupacion.Caption = Trim(RsAux!RamNombre)
        
        RsAux.Close
        
    End If
    '------------------------------------------------------------------------------------------------------
    
    If cDireccion.ListCount <= 1 Then
        If Val(cDireccion.Tag) = 0 Then lNDireccion.Caption = Trim(cDireccion.Text) & ":"
        lNDireccion.Refresh
        lDireccion.Left = lNDireccion.Left + lNDireccion.Width + 40
    Else
        lDireccion.Left = cDireccion.Left + cDireccion.Width + 40
    End If
    
    lTitulo.Caption = CargoTitulos(Titular)
    lTitulo.Refresh
    
    'Icono de relaciones
    If gTipoCliente = TipoCliente.Cliente Then
        Cons = "Select * from PersonaRelacion Where PReClienteDe = " & Titular
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then Tab1.Tabs("relaciones").Image = ImageList1.ListImages("relaciones").Index
        RsAux.Close
    End If
    
    Exit Sub
    
errCliente:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente."
End Sub

Private Sub CargoDatosClienteRelaciones(Cliente As Long, Optional Full As Boolean = False)

    If Cliente = 0 Then Exit Sub
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    
    Dim mValor As Long, mClientes As String
    Dim prmAccesoFuncionariosEmp As Boolean
    
    vsRelacion.Rows = vsRelacion.FixedRows
    
    If gTipoCliente = 1 Then
        Cons = "Select RelOrden, RelCodigo, RelNombre, Cliente.*, CPersona.* " & _
                    " From PersonaRelacion, CPersona, Cliente, Relaciones " & _
                    " Where PReClienteDe = " & Cliente & _
                    " And PReClienteEs = CPeCliente " & _
                    " And PReRelacion = RelCodigo" & _
                    " And CPeCliente = CliCodigo"
                    
        If gConyuge <> 0 Then
            Cons = Cons & " UNION " & _
                    " Select -1 as RelOrden, 0 as RelCodigo, 'Cónyuge' as RelNombre, Cliente.*, CPersona.* " & _
                    " From CPersona, Cliente" & _
                    " Where CliCodigo = " & gConyuge & _
                    " And CPeCliente = CliCodigo"
        End If
        
        Cons = Cons & " Order by RelOrden"
    Else
        prmAccesoFuncionariosEmp = miConexion.AccesoAlMenu("Funcionarios de Empresas")
        
        Cons = "Select 0 as RelCodigo, OcuNombre as RelNombre, Cliente.*, CPersona.* " & _
                    " From Cliente, CPersona, Ocupacion, Empleo" & _
                    " Where EmpCEmpresa = " & Cliente & _
                    " And EmpCliente = CPeCliente" & _
                    " And CPeCliente = CliCodigo" & _
                    " And EmpOcupacion = OcuCodigo"
        If Not Full Then Cons = Cons & " And OcuCodigo IN (" & prmOcupacionesEmp & ")"
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF And gTipoCliente = 2 Then Tab1.Tabs("relaciones").Image = ImageList1.ListImages("relaciones").Index
    
    Do While Not RsAux.EOF
        With vsRelacion
            .AddItem ""
            aTexto = Trim(RsAux!RelNombre)
            If RsAux!RelCodigo = paRelPadre Then
                Select Case RsAux!CPeSexo
                    Case "M": aTexto = "Padre"
                    Case "F": aTexto = "Madre"
                    Case Else: aTexto = RsAux!RelNombre
                End Select
            End If
            .Cell(flexcpText, .Rows - 1, 0) = Trim(aTexto)
            
            If Not IsNull(RsAux!CliCiRuc) Then .Cell(flexcpText, .Rows - 1, 1) = clsGeneral.RetornoFormatoCedula(RsAux!CliCiRuc)
            mValor = RsAux!CliCodigo: .Cell(flexcpData, .Rows - 1, 1) = mValor
            mClientes = mClientes & mValor
            
            .Cell(flexcpText, .Rows - 1, 2) = ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
            
            If Not IsNull(RsAux!CPeFNacimiento) Then .Cell(flexcpText, .Rows - 1, 3) = ((Date - RsAux!CPeFNacimiento) \ 365) '& Format(rsAux!CPeFNacimiento, " (d-Mmm-yyyy)")
            
            .Cell(flexcpBackColor, .Rows - 1, 4, , 5) = RGB(230, 230, 250)
        End With
        RsAux.MoveNext
        
        If Not RsAux.EOF Then mClientes = mClientes & ","
    Loop
    RsAux.Close
    
    Me.Refresh
    
    If Not Full And gTipoCliente = 2 And prmAccesoFuncionariosEmp Then
        vsRelacion.AddItem ""
        vsRelacion.Cell(flexcpText, vsRelacion.Rows - 1, 2) = "(cargar todos los empleados)"
        vsRelacion.Cell(flexcpData, vsRelacion.Rows - 1, 1) = "-1"
        vsRelacion.Cell(flexcpBackColor, vsRelacion.Rows - 1, 0, , vsRelacion.Cols - 1) = vsRelacion.GridColorFixed
        vsRelacion.Cell(flexcpForeColor, vsRelacion.Rows - 1, 0, , vsRelacion.Cols - 1) = vbWhite
        vsRelacion.Cell(flexcpFontBold, vsRelacion.Rows - 1, 0, , vsRelacion.Cols - 1) = True
    End If
    
    If Trim(mClientes) = "" Then Screen.MousePointer = 0: Exit Sub
    rel_CargoOtrosDatos 0
    
    Screen.MousePointer = 0
    Exit Sub

errCargar:
    clsGeneral.OcurrioError "Error al cargar las relaciones.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function rel_CargoOtrosDatos(mTipoDato As Integer)
'0- Todos, 1-Filas de Totales,  2-Filas de Vigentes, 3-Comentarios Negativos

Dim mClientes As String, idX As Integer

    mClientes = ""
    With vsRelacion
        For idX = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpData, idX, 1)) <> 0 Then
                mClientes = mClientes & IIf(mClientes = "", Val(.Cell(flexcpData, idX, 1)), "," & Val(.Cell(flexcpData, idX, 1)))
            End If
        Next
    End With
    If Trim(mClientes) = "" Then Screen.MousePointer = 0: Exit Function
    
    If mTipoDato = 0 Or mTipoDato = 2 Then    'CASO 2: Cargo los saldos de operaciones para c/u de las relaciones -------------------------
        
        Cons = "Select  DocCliente, Count(*) as Q, Min(CreProximoVto) as Vto, Sum(CreSaldoFactura) as Saldo" & _
                    " From Credito, Documento" & _
                    " Where CreFactura = DocCodigo " & _
                    " And DocCliente In (" & Trim(mClientes) & ")" & _
                    " And DocTipo = " & TipoDocumento.Credito & _
                    " And DocAnulado = 0 " & _
                    " And CreSaldoFactura > 0 " & _
                    " Group by DocCliente"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            With vsRelacion
            For I = 1 To .Rows - 1
                If .Cell(flexcpData, I, 1) = RsAux!DocCliente Then
                    .Cell(flexcpText, I, 6) = RsAux!Q
                    .Cell(flexcpAlignment, I, 6) = flexAlignRightCenter
                    If Not IsNull(RsAux!Saldo) Then
                        If RsAux!Saldo > 0 Then
                            .Cell(flexcpText, I, 8) = Format(RsAux!Saldo, "#,##0.00")
                            .Cell(flexcpAlignment, I, 8) = flexAlignRightCenter
                            
                            If Not IsNull(RsAux!Vto) Then
                                .Cell(flexcpText, I, 7) = Format(RsAux!Vto, "dd/mm/yyyy")
                                
                                Select Case RsAux!Vto - gFechaServidor
                                    Case Is < 0
                                        If Abs(RsAux!Vto - gFechaServidor) > paToleranciaMora Then
                                            .Cell(flexcpBackColor, I, 7) = Colores.RojoClaro
                                            .Cell(flexcpForeColor, I, 7) = Colores.Blanco
                                        Else
                                            .Cell(flexcpBackColor, I, 7) = Colores.Obligatorio
                                        End If
                                End Select
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        vsRelacion.Refresh
    End If
    
    If mTipoDato = 0 Or mTipoDato = 1 Then  'CASO 1: Cargo las Columnas de Totales -------------------------------------------
        Cons = "Select  DocCliente, Count(*) as Q, Sum(DocTotal) as Total" & _
                    " From Documento" & _
                    " Where DocCliente In (" & Trim(mClientes) & ")" & _
                    " And DocTipo = " & TipoDocumento.Credito & _
                    " And DocAnulado = 0 " & _
                    " Group by DocCliente"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            With vsRelacion
            For I = 1 To .Rows - 1
                If .Cell(flexcpData, I, 1) = RsAux!DocCliente Then
                    .Cell(flexcpText, I, 4) = RsAux!Q
                    .Cell(flexcpAlignment, I, 4) = flexAlignRightCenter
                    
                    .Cell(flexcpText, I, 5) = Format(RsAux!Total, "#,##0.00")
                    .Cell(flexcpAlignment, I, 5) = flexAlignRightCenter
                    Exit For
                End If
            Next
            End With
        
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        vsRelacion.Refresh
    End If
    
    If mTipoDato = 0 Or mTipoDato = 3 Then  'CASO 3: Cargo comentarios Negativos -------------------------------------------
        Cons = "Select ComCliente from Comentario, TipoComentario" & _
                    " Where ComCliente In (" & Trim(mClientes) & ")" & _
                    " And ComTipo = TCoCodigo And TCoClase = 3"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            With vsRelacion
            For I = 1 To .Rows - 1
                If .Cell(flexcpData, I, 1) = RsAux!ComCliente Then
                    .Cell(flexcpText, I, .Cols - 1) = "Ver comentarios ..."
                    .Cell(flexcpForeColor, I, .Cols - 1) = Colores.RojoClaro
                    .Cell(flexcpAlignment, I, .Cols - 1) = flexAlignLeftCenter
                    Exit For
                End If
            Next
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        vsRelacion.Refresh
    End If
End Function

Private Sub CargoEmpleos(Cliente As Long, Optional Titular As Boolean = False, Optional Conyuge As Boolean = False, Optional Garantia As Boolean = False)

Dim aRsEmp As rdoResultset
Dim mValor As Long

    Screen.MousePointer = 11
    
    On Error GoTo errEmpleos
'    Cons = "Select Empleo.*, CEmpresa.*, OcuNombre, TExAbreviacion" _
'           & " From Empleo, CEmpresa, Ocupacion, TipoExhibido" _
'           & " Where EmpCliente = " & Cliente _
'           & " And EmpCEmpresa *= CEmCliente" _
'           & " And EmpOcupacion = OcuCodigo" _
'           & " And EmpTipoExhibido *= TExCodigo" & " Order by EmpCodigo DESC"
    
    Cons = "SELECT Empleo.*, Empresa = RTrim(IsNull(CEmNombre, rTrim(CPeNombre1) + ' '  + rTrim(CPeApellido1))), OcuNombre, TExAbreviacion " & _
        "FROM Empleo " & _
        "LEFT OUTER JOIN CEmpresa ON EmpCEmpresa = CEmCliente " & _
        "LEFT OUTER JOIN CPersona ON EmpCEmpresa = CPeCliente " & _
        "INNER JOIN Ocupacion ON EmpOcupacion = OcuCodigo " & _
        "LEFT OUTER JOIN TipoExhibido ON EmpTipoExhibido = TExCodigo " & _
        "WHERE EmpCliente = " & Cliente & _
        " ORDER BY EmpCodigo DESC"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    'Agrego en la lista de empleos segun la señal booleana
    Dim aVs As vsFlexGrid
    If Titular Then Set aVs = lEmpleoT
    If Conyuge Or Garantia Then Set aVs = lEmpleoG
    
    aVs.Rows = 1
    Do While Not RsAux.EOF
        With aVs
            
            aTexto = "S/D"
'            If Not IsNull(RsAux!EmpCEmpresa) Then
'                If Not IsNull(RsAux!CEmFantasia) Then aTexto = Trim(RsAux!CEmFantasia) Else If Not IsNull(RsAux("CEmNombre")) Then aTexto = Trim(RsAux!CEmNombre)
'            End If
            .AddItem RsAux("Empresa")
            mValor = RsAux!EmpCodigo: .Cell(flexcpData, .Rows - 1, 0) = mValor
            
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!EmpFechaIngreso, "mm/yy")
            
            aTexto = Trim(RsAux!OcuNombre)
            If Not IsNull(RsAux!EmpCargo) Then If Trim(RsAux!EmpCargo) <> "" Then aTexto = aTexto & " (" & Trim(RsAux!EmpCargo) & ")"
            .Cell(flexcpText, .Rows - 1, 2) = aTexto
            
            'Saco los datos de tablas auxiliares del empleo---------------------------------------------------------------------
'            Cons = "Select TInNombre, MonSigno, VEmNombre" _
'                    & " From Empleo, TipoIngreso, Moneda, VigenciaEmpleo" _
'                    & " Where EmpCodigo = " & RsAux!EmpCodigo _
'                    & " And EmpTipoIngreso *= TInCodigo" _
'                    & " And EmpMoneda *= MonCodigo" _
'                    & " And EmpVigencia *= VEmCodigo"
                    
                    
            Cons = "SELECT TInNombre, MonSigno, VEmNombre " & _
                    "FROM Empleo " & _
                    "LEFT OUTER JOIN TipoIngreso ON EmpTipoIngreso = TInCodigo " & _
                    "LEFT OUTER JOIN Moneda ON EmpMoneda = MonCodigo " & _
                    "LEFT OUTER JOIN VigenciaEmpleo ON EmpVigencia = VEmCodigo " & _
                    "WHERE EmpCodigo = " & RsAux!EmpCodigo
                    
            Set aRsEmp = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            aTexto = ""
            If Not IsNull(aRsEmp!MonSigno) Then aTexto = Trim(aRsEmp!MonSigno)
            If Not IsNull(RsAux!EmpLiquido) Then aTexto = aTexto & " Líq." & Format(RsAux!EmpLiquido, "#,##0") & " "
            If Not IsNull(RsAux!EmpNominal) Then aTexto = aTexto & " Nom." & Format(RsAux!EmpNominal, "#,##0")
            
            If IsNull(RsAux!EmpLiquido) Or IsNull(RsAux!EmpNominal) Then
                If Not IsNull(aRsEmp!TInNombre) Then aTexto = aTexto & " " & Trim(aRsEmp!TInNombre)
                If Not IsNull(RsAux!EmpIngreso) Then aTexto = aTexto & " " & Format(RsAux!EmpIngreso, "#,##0")
            End If
            If Not IsNull(aRsEmp!VEmNombre) Then aTexto = aTexto & " (" & Trim(aRsEmp!VEmNombre) & ")"
            
            .Cell(flexcpText, .Rows - 1, 3) = aTexto
            aRsEmp.Close
            '----------------------------------------------------------------------------------------------------------------------------
            
            If Not IsNull(RsAux!EmpTipoExhibido) Then
                aTexto = Trim(RsAux!TExAbreviacion)
                If Not IsNull(RsAux!EmpExhibido) Then
                    aTexto = aTexto & " " & Format(RsAux!EmpExhibido, "dd/mm/yy")
                    If Abs(DateDiff("d", RsAux!EmpExhibido, gFechaServidor)) < 45 Then  'Es reciente
                        .Cell(flexcpBackColor, .Rows - 1, 4) = Colores.Obligatorio
                        If Not IsNull(RsAux!EmpFModificacion) Then
                            If Format(RsAux!EmpFModificacion, "dd/mm/yyyy") = Format(gFechaServidor, "dd/mm/yyyy") Then .Cell(flexcpBackColor, .Rows - 1, 4) = Colores.clVerde
                        End If
                    End If
                End If
                .Cell(flexcpText, .Rows - 1, 4) = aTexto
            End If
            
            If Not IsNull(RsAux!EmpVigencia) Then
                If RsAux!EmpVigencia = paempNoTrabajaMas Then .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
                If RsAux!EmpVigencia = paempSeguroParo Then .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris
            End If
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Set aVs = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errEmpleos:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los empleos.", Err.Description
End Sub

Private Function CargoTitulos(Cliente As Long)

    On Error GoTo errTitulos
    CargoTitulos = ""
    
    Cons = "Select Count(*) From Titulo" _
           & " Where TitCliente = " & Cliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux(0)) Then
            If RsAux(0) > 0 Then CargoTitulos = "Títulos (" & RsAux(0) & ")"
        End If
    End If
    RsAux.Close
    Exit Function
    
errTitulos:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los títulos."
End Function

Private Sub CargoReferencias(Cliente As Long, Optional Titular As Boolean = False, Optional Conyuge As Boolean = False, Optional Garantia As Boolean = False)
Dim mValor As Long

    On Error GoTo errReferencias
    Screen.MousePointer = 11
'    Cons = "Select ReferenciaCliente.*, RefDescripcion, Comprobante.*, MonSigno" _
'           & " From ReferenciaCliente, Referencia, Comprobante, Moneda" _
'           & " Where RClCPersona = " & Cliente _
'           & " And RClReferencia = RefCodigo" _
'           & " And RClComprobante = ComCodigo" _
'           & " And RClMoneda *= MonCodigo" & " Order by RClCodigo DESC"
           
    Cons = "SELECT ReferenciaCliente.*, RefDescripcion, Comprobante.*, MonSigno " & _
        "FROM ReferenciaCliente INNER JOIN Referencia ON RClReferencia = RefCodigo " & _
        "INNER JOIN Comprobante ON RClComprobante = ComCodigo " & _
        "LEFT OUTER JOIN Moneda  ON RCLMoneda = MonCodigo " & _
        "WHERE RClCPersona = " & Cliente & _
        "ORDER BY RClCodigo DESC"
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    'Agrego en la lista de empleos segun la señal booleana
    Dim aVs As vsFlexGrid
    If Titular Then Set aVs = lReferenciaT
    If Garantia Or Conyuge Then Set aVs = lReferenciaG
    
    aVs.Rows = 1
    Do While Not RsAux.EOF
        With aVs
        
            .AddItem Trim(RsAux!RefDescripcion)
            mValor = RsAux!RClCodigo: .Cell(flexcpData, .Rows - 1, 0) = mValor
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ComNombre)
            
            aTexto = Trim(RsAux!ComNombreValor1) & ": " & FormatoReferencia(RsAux!RClValor1, RsAux!ComFormatoV1)
            If Not IsNull(RsAux!RClValor2) Then aTexto = aTexto & "; " & Trim(RsAux!ComNombreValor2) & ": " & FormatoReferencia(RsAux!RClValor2, RsAux!ComFormatoV2)
            If Not IsNull(RsAux!MonSigno) Then aTexto = "(" & Trim(RsAux!MonSigno) & ") " & aTexto
            .Cell(flexcpText, .Rows - 1, 2) = aTexto
            
            If Not IsNull(RsAux!RClExhibido) Then
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!RClExhibido, "dd/mm/yy")
                If Abs(DateDiff("d", RsAux!RClExhibido, gFechaServidor)) < 45 Then
                    .Cell(flexcpBackColor, .Rows - 1, 3) = Colores.Obligatorio
                    If Not IsNull(RsAux!RClFModificacion) Then
                        If Format(RsAux!RClFModificacion, "dd/mm/yyyy") = Format(gFechaServidor, "dd/mm/yyyy") Then .Cell(flexcpBackColor, .Rows - 1, 3) = Colores.clVerde
                    End If
                End If
            End If
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Set aVs = Nothing
    Screen.MousePointer = 0
    Exit Sub

errReferencias:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar las referencias."
End Sub


Private Sub CargoComentarios(Cliente As Long, Optional Titular As Boolean = False, Optional Conyuge As Boolean = False, Optional Garantia As Boolean = False)

Dim mValor As Long

    Screen.MousePointer = 11
    On Error GoTo errComentario

    'Agrego en la lista de empleos segun la señal booleana
    Dim aVs As vsFlexGrid
    If Titular Then Set aVs = lComentarioT
    If Garantia Or Conyuge Then Set aVs = lComentarioG
    
    Dim aColor  As Long
    aVs.Rows = 1
    
    Cons = "Select * from Comentario Left Outer Join Documento On ComDocumento = DocCodigo, TipoComentario, Usuario" _
            & " Where ComCliente = " & Cliente _
            & " And ComTipo = TCoCodigo" _
            & " And ComUsuario = UsuCodigo"

    Cons = Cons & " Order by ComFecha DESC"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    With aVs
    Do While Not RsAux.EOF
        .AddItem Format(RsAux!ComFecha, "dd/mm/yy hh:mm")
        mValor = RsAux!ComCodigo: .Cell(flexcpData, .Rows - 1, 0) = mValor
        
        .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ComComentario)
        .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!TCoNombre)
        
        .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!UsuIdentificacion)
        
        aColor = vbWindowBackground
        If Not IsNull(RsAux!TCoBColor) Then aColor = RsAux!TCoBColor
        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = aColor
        
        aColor = vbWindowText
        If Not IsNull(RsAux!TCoFColor) Then aColor = RsAux!TCoFColor
        .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = aColor
        
        If Not IsNull(RsAux!DocCodigo) Then
            .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!DocSerie) & " " & RsAux!DocNumero
            mValor = RsAux!DocCodigo: .Cell(flexcpData, .Rows - 1, 4) = mValor
        End If
        
        RsAux.MoveNext
    Loop
    End With
    Set aVs = Nothing
    RsAux.Close
    Exit Sub
    
errComentario:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los comentarios del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoCantidadOperaciones(Cliente As Long)
    
    oCredito.Tag = 0: oContado.Tag = 0: oSolicitud.Tag = 0
    oArticulo.Tag = "": oArticulo.Caption = "&Artículos"
    oSolicitud.ForeColor = vbBlack
    If Cliente = 0 Then
        oContado.Caption = "C&ontado (N/D)"
        oCredito.Caption = "Cré&dito (N/D)"
        oSolicitud.Caption = "&Solicitudes (N/D)"
        oVtaTel.Caption = "&Vtas. Telefónicas (N/D)"
        Exit Sub
    Else
        oContado.Caption = "C&ontado (0)"
        oCredito.Caption = "Cré&dito (0)"
        oSolicitud.Caption = "&Solicitudes (0)"
        oVtaTel.Caption = "&Vtas. Telefónicas (0)"
    End If
    
    On Error GoTo errCantidad
    Screen.MousePointer = 11
    'Cargo Cantidad de Operaciones Credito Y Contado-------------------------------------------
    Cons = "Select Count(*), DocTipo From Documento (index = iClienteTipo)" _
           & " Where DocCliente = " & Cliente _
           & " And DocTipo IN (" & TipoDocumento.Credito & "," & TipoDocumento.Contado & "," & TipoDocumento.NotaEspecial & ")" _
           & " And DocAnulado = 0 Group by DocTipo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Dim aQCtdo As Long: aQCtdo = 0
    Do While Not RsAux.EOF
        If Not IsNull(RsAux(0)) Then
            Select Case RsAux!DocTipo
                Case TipoDocumento.Credito: oCredito.Caption = "Cré&dito (" & RsAux(0) & ")": oCredito.Tag = RsAux(0)
                
                Case TipoDocumento.Contado, TipoDocumento.NotaEspecial: aQCtdo = aQCtdo + RsAux(0)
            End Select
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    oContado.Caption = "C&ontado (" & aQCtdo & ")": oContado.Tag = aQCtdo
    
    'Si tiene 0 como titular veo si es garante
    If oCredito.Tag = 0 Then
        Cons = "Select Count(*) " _
                & " From Credito (index = iGarantia), Documento" _
                & " Where CreGarantia = " & Cliente _
                & " And CreFactura = DocCodigo" _
                & " And DocAnulado = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then oCredito.Tag = RsAux(0)
        RsAux.Close
    End If
    '---------------------------------------------------------------------------------------------
        
    'Cargo Cantidad de Solicitudes----------------------------------------------------------
    Cons = "Select Count(*), Datediff(d, MAX(SolFecha), Getdate()) From Solicitud" _
           & " Where SolCliente = " & Cliente & " AND SolCodigo <> " & prmSolicitud
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux(0)) Then
            oSolicitud.Caption = "&Solicitudes (" & RsAux(0) & ")"
            oSolicitud.Tag = RsAux(0)
            
            If RsAux(1) > 0 And RsAux(1) < 30 Then oSolicitud.ForeColor = &H40C0&
        End If
    End If
    RsAux.Close
    '---------------------------------------------------------------------------------------------

    'Cargo Cantidad de Solicitudes----------------------------------------------------------
    Cons = "Select Count(*) From VentaTelefonica" _
           & " Where VTeCliente = " & Cliente & " And VTeAnulado is null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux(0)) Then
            oVtaTel.Caption = "&Vtas. Telefónicas (" & RsAux(0) & ")"
            oVtaTel.Tag = RsAux(0)
        End If
    End If
    RsAux.Close
    '---------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    Exit Sub

errCantidad:
    clsGeneral.OcurrioError "Error al procesar la cantidad de operaciones.", Err.Description
End Sub

Private Sub CargoOpContado(Cliente As Long, Optional bNext As Boolean = False)

    If Not bNext Or lOperacion.Tag = "" Then
        If bNext And lOperacion.Tag = "" Then Exit Sub
        lOperacion.Rows = 1: lOperacion.Refresh
        lTotalComprado.Caption = "": lTotalComprado.Tag = 0
    End If

    If Cliente = 0 Or oContado.Tag = 0 Then Exit Sub
    
    On Error GoTo errContado
    Screen.MousePointer = 11
    lOperacion.Redraw = False

Dim mSQLTCs As String, colDesdeTCs As String
Dim mIDSDocsCtdo As String

    mSQLTCs = "": mIDSDocsCtdo = ""
    colDesdeTCs = (lOperacion.Rows - lOperacion.FixedRows)
    aMonedaAnteriorC = 0: I = 0
        
    'cons = "Select Top " & PorPagina + 1 & " DocCodigo, DocTipo, DocMoneda, DocFecha, DocSerie, DocNumero, DocTotal," & _
                            " IsNull(Count(ArtID), 0) as Renglones, Sum(RenCantidad) QArticulos, " & _
                            " Min(ArtNombre) as ArtNombre, " & _
                            " IsNull(Count (NotFactura), 0) as QNotas" & _
            " From Documento(Index = iClienteTipo) " & _
                            " Left Outer Join Nota On NotFactura = DocCodigo, " & _
                    " Renglon Left Outer Join  Articulo On RenArticulo = ArtID " & _
                                                        " And ArtTipo <> " & paTipoArticuloServicio & _
           " Where DocCliente = " & Cliente & _
           " And DocTipo  IN ( " & TipoDocumento.Contado & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")" & _
           " And DocAnulado = 0" & _
           " And DocCodigo = RenDocumento"
    
    Cons = "Select Top " & PorPagina + 1 & " DocCodigo, DocTipo, DocMoneda, DocFecha, DocSerie, DocNumero, DocTotal," & _
                            " IsNull(Count(ArtID), 0) as Renglones, Sum(Case When ArtID IS NULL THEN 0 Else RenCantidad end) as QArticulos,  " & _
                            " Min(ArtNombre) as ArtNombre " & _
            " From Documento(Index = iClienteTipo), " & _
                    " Renglon Left Outer Join  Articulo On RenArticulo = ArtID " & _
                                                        " And ArtTipo <> " & paTipoArticuloServicio & _
           " Where DocCliente = " & Cliente & _
           " And DocTipo  IN ( " & TipoDocumento.Contado & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")" & _
           " And DocAnulado = 0" & _
           " And DocCodigo = RenDocumento"
           
    If Trim(lOperacion.Tag) <> "" Then Cons = Cons & " And DocFecha < '" & Format(lOperacion.Tag, sqlFormatoFH) & "'"
    
    Cons = Cons & " Group By DocCodigo, DocTipo, DocMoneda, DocFecha, DocSerie, DocNumero, DocTotal" & _
                        " Order by DocFecha Desc"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF And I < PorPagina
        
        With lOperacion
            .Tag = RsAux!DocFecha
        
            .AddItem Format(RsAux!DocFecha, "dd/mm/yy")
            mValor = RsAux!DocCodigo: .Cell(flexcpData, .Rows - 1, 0) = mValor  '0= Codigo del documento
            mValor = RsAux!DocTipo: .Cell(flexcpData, .Rows - 1, 1) = mValor      '1= Tipo del documento
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!DocTotal, "#,##0.00")
        
            'Saco la Moneda
            If aMonedaAnteriorC <> RsAux!DocMoneda Then
                Cons = "Select * from Moneda Where MonCodigo = " & RsAux!DocMoneda
                Set rsMon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                .Cell(flexcpText, .Rows - 1, 3) = Trim(rsMon!MonSigno)
                aMonedaAnteriorT = Trim(rsMon!MonSigno)
                aMonedaAnteriorC = rsMon!MonCodigo
                rsMon.Close
            Else
                .Cell(flexcpText, .Rows - 1, 3) = Trim(aMonedaAnteriorT)
            End If
        
            If paMonedaFija <> 0 And RsAux!DocMoneda <> paMonedaFija Then
                '.Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!DocTotal * TasadeCambio(rsAux!DocMoneda, paMonedaFija, rsAux!DocFecha), "#,##0.00")
                '.Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!DocTotal, "#,##0.00")
                If Trim(mSQLTCs) <> "" Then mSQLTCs = mSQLTCs & " UNION ALL "
                mSQLTCs = mSQLTCs & _
                            "Select " & RsAux!DocCodigo & " as ID, TCaComprador from TasaCambio" & _
                            " Where TCaFecha = (Select MAX(TCaFecha) from TasaCambio " & _
                                          " Where TCaFecha < '" & Format(RsAux!DocFecha, "mm/dd/yyyy 23:59") & "'" _
                                          & " And TCaOriginal = " & RsAux!DocMoneda & " And TCaDestino = " & paMonedaFija & ")" & _
                            " And TCaOriginal = " & RsAux!DocMoneda & " And TCaDestino = " & paMonedaFija
                
            Else
                .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!DocTotal, "#,##0.00")
            End If
            
            'Si es Nota le coloco icono "Nota"
            If RsAux!DocTipo = TipoDocumento.NotaDevolucion Or RsAux!DocTipo = TipoDocumento.NotaEspecial Then
                .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("Nota").ExtractIcon
            End If
            
            If RsAux!DocTipo = TipoDocumento.Contado Then 'And Not IsNull(rsAux!QNotas) Then
                If Trim(mIDSDocsCtdo) <> "" Then mIDSDocsCtdo = mIDSDocsCtdo & ","
                mIDSDocsCtdo = mIDSDocsCtdo & RsAux!DocCodigo
                
                'If (rsAux!QNotas) > 1 Then      '19/07/04 No valido mas si está Anulada !!!!
                '    .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("Alerta").ExtractIcon
                '    .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1)
                '    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio
                '    .Cell(flexcpText, .Rows - 1, 2) = rsAux!QNotas
                'End If
            End If
            '-------------------------------------------------------------------------------------------------------------------------
            
            aTexto = ""
            Select Case RsAux!Renglones
                Case 0: aTexto = "Articulo/s de Servicio"
                Case 1: aTexto = IIf(RsAux!QArticulos > 1, RsAux!QArticulos & " ", "") & Trim(RsAux!ArtNombre)
                Case Else
                    aTexto = RsAux!QArticulos & " Artículos"
                    If RsAux!Renglones <> RsAux!QArticulos Then aTexto = aTexto & " (" & RsAux!Renglones & " diferentes)"
            End Select
            .Cell(flexcpText, .Rows - 1, 6) = Trim(aTexto)
                           
        End With
        RsAux.MoveNext
        I = I + 1
    Loop
    If RsAux.EOF Then lOperacion.Tag = ""
    RsAux.Close
    
    On Error Resume Next
    lOperacion.Redraw = True
    lOperacion.Select lOperacion.Rows - I, 0
    lOperacion.TopRow = lOperacion.Row
    lOperacion.Refresh
    
    Dim idX As Integer
    '2) Cargo las Tasas de cambios para c/u de los documentos   -------------------------------------------------------------
    If mSQLTCs <> "" Then
        Set RsAux = cBase.OpenResultset(mSQLTCs, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            With lOperacion
                For idX = (colDesdeTCs + .FixedRows) To .Rows - 1
                    If RsAux!ID = .Cell(flexcpData, idX, 0) Then
                        .Cell(flexcpText, idX, 5) = Format(.Cell(flexcpValue, idX, 4) * Format(RsAux!TCaComprador, "#.000"), "#,##0.00")
                        Exit For
                    End If
                Next
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If

    '3) Cargo las Q de Notas Para las facturas ----------------------------------------------------------------------------------------------------------------------------------
    If mIDSDocsCtdo <> "" Then
        Cons = "Select NotFactura, Count(*) as QNotas From Nota, Documento " & _
                   " Where NotNota = DocCodigo And NotFactura In (" & mIDSDocsCtdo & ")" & _
                   " And DocAnulado = 0 " & _
                   " Group by NotFactura Having Count(*) > 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            With lOperacion
                For idX = (colDesdeTCs + .FixedRows) To .Rows - 1
                    If RsAux!NotFactura = .Cell(flexcpData, idX, 0) Then
                        .Cell(flexcpPicture, idX, 0) = ImageList1.ListImages("Alerta").ExtractIcon
                        .Cell(flexcpText, idX, 1) = .Cell(flexcpText, idX, 1)
                        .Cell(flexcpBackColor, idX, 0, , .Cols - 1) = Colores.Obligatorio
                        .Cell(flexcpText, idX, 2) = RsAux!QNotas
                        Exit For
                    End If
                Next
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
        
    If lOperacion.Tag = "" Then
        lTotalComprado.Caption = lOperacion.Rows - 1 & " de " & lOperacion.Rows - 1 & "  Contados"
    Else
        lTotalComprado.Caption = lOperacion.Rows - 1 & " de ..." & "  Contados"
    End If

    'Recorro para sacar el Total de las compras de éstas operaciones    ------------------------------------------------------------------------------------
    Dim mTotalOps As Currency: mTotalOps = 0
    With lOperacion
        For idX = (colDesdeTCs + .FixedRows) To .Rows - 1
            If .Cell(flexcpData, idX, 1) = TipoDocumento.NotaDevolucion Or .Cell(flexcpData, idX, 1) = TipoDocumento.NotaEspecial Then
                mTotalOps = mTotalOps - .Cell(flexcpValue, idX, 5)
            Else
                mTotalOps = mTotalOps + .Cell(flexcpValue, idX, 5)
            End If
        Next
    End With
    If mTotalOps <> 0 Then
        If Val(lTotalComprado.Tag) <> 0 Then mTotalOps = CCur(lTotalComprado.Tag) + mTotalOps
        lTotalComprado.Tag = mTotalOps
        lTotalComprado.Caption = lTotalComprado.Caption & " (" & Trim(paMonedaFijaTexto) & " " & Format(mTotalOps, "#,##0.00") & ")"
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Screen.MousePointer = 0
    Exit Sub
    
errContado:
    clsGeneral.OcurrioError "Error al cargar las operaciones contado.", Err.Description
    lOperacion.Redraw = True
    Screen.MousePointer = 0
End Sub


Private Sub CargoOpCredito(Cliente As Long, Optional bNext As Boolean = False)
On Error Resume Next

    If Not bNext Or lOperacion.Tag = "" Then
        If bNext And lOperacion.Tag = "" Then Exit Sub
        lOperacion.Rows = 1: lOperacion.Refresh
        lTotalComprado.Caption = "": lTotalComprado.Tag = 0
    End If
    
    If Cliente = 0 Or oCredito.Tag = 0 Then Exit Sub
    
Dim RsAux1 As rdoResultset
Dim mValor As Long, aIcono As String
Dim mSQLTCs As String, colDesdeTCs As Integer

    On Error GoTo errCargar
    Screen.MousePointer = 11
    
    mSQLTCs = ""
    colDesdeTCs = (lOperacion.Rows - lOperacion.FixedRows)
    aMonedaAnteriorC = 0
    I = 0
    
    lOperacion.Redraw = False
    ArmoConsultaDeCreditos Cliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF And I < PorPagina
        I = I + 1
        With lOperacion
            .Tag = RsAux!DocFecha
            .AddItem ""
            mValor = RsAux!DocCodigo: .Cell(flexcpData, .Rows - 1, 0) = mValor  'Guardo el codigo del documento
            .Cell(flexcpBackColor, .Rows - 1, 0) = Colores.Azul: .Cell(flexcpForeColor, .Rows - 1, 0) = Colores.Blanco
            
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!DocFecha, "dd/mm/yy")
        
            If RsAux!CreSaldoFactura > 0 Then                    'Facturas pendientes
                aIcono = ""
                Select Case RsAux!CreTipo
                    Case TipoCredito.Normal
                        Select Case RsAux!CreProximoVto - gFechaServidor        'Pongo iconos a Facturas Vigentes
                            Case Is < 0
                                If Abs(RsAux!CreProximoVto - gFechaServidor) > paToleranciaMora Then
                                    aIcono = "No"
                                    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.RojoClaro: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
                                Else
                                    aIcono = "Alerta"
                                    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio
                                End If
                            Case Is >= 0: aIcono = "Si"
                        End Select
                
                    Case TipoCredito.Gestor: aIcono = "Gestor"
                    Case TipoCredito.Incobrable: aIcono = "Perdida"
                    Case TipoCredito.Clearing: aIcono = "Clearing"
                End Select
                If aIcono <> "" Then .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages(aIcono).ExtractIcon
                
            Else                'Iconos de Puntaje 0..9
                If IsNull(RsAux!CrePuntaje) Then
                    .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("Nota").ExtractIcon
                Else
                    Select Case RsAux!CrePuntaje
                        Case 0: aIcono = "0"
                        Case 1: aIcono = "1"
                        Case 2: aIcono = "2"
                        Case 3: aIcono = "3"
                        Case 4: aIcono = "4"
                        Case 5: aIcono = "5"
                        Case 6: aIcono = "6"
                        Case 7: aIcono = "7"
                        Case 8: aIcono = "8"
                        Case Else: aIcono = "9"
                    End Select
                    .Cell(flexcpText, .Rows - 1, 0) = aIcono
                    .Cell(flexcpFontBold, .Rows - 1, 0) = True
                End If
            End If
            
            If RsAux!CreFormaPago = TipoPagoSolicitud.ChequeDiferido Then aTexto = "CD " & Trim(RsAux!TCuAbreviacion) Else aTexto = Trim(RsAux!TCuAbreviacion)
            .Cell(flexcpText, .Rows - 1, 2) = aTexto
            
            aTexto = "0/"
            If Not IsNull(RsAux!CreVaCuota) Then aTexto = Trim(RsAux!CreVaCuota) & "/"
            aTexto = aTexto & Trim(RsAux!CreDeCuota)
            .Cell(flexcpText, .Rows - 1, 3) = Trim(aTexto)
            
            .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!CreCumplimiento)
            
            If Trim(mSQLTCs) <> "" Then mSQLTCs = mSQLTCs & " UNION ALL "
            mSQLTCs = mSQLTCs & _
                            "Select " & RsAux!DocCodigo & " as ID, TCaComprador from TasaCambio" & _
                            " Where TCaFecha = (Select MAX(TCaFecha) from TasaCambio " & _
                                          " Where TCaFecha < '" & Format(RsAux!DocFecha, "mm/dd/yyyy 23:59") & "'" _
                                          & " And TCaOriginal = " & RsAux!DocMoneda & " And TCaDestino = " & paMonedaFija & ")" & _
                            " And TCaOriginal = " & RsAux!DocMoneda & " And TCaDestino = " & paMonedaFija
                            
            '(Valores en M/N p/dps hacer SQL c/ TC) -> Monto Total y Valor Cuota con T/C a la fecha del documento
            .Cell(flexcpText, .Rows - 1, 11) = Format(RsAux!DocTotal, "#,##0.00")
            If Not IsNull(RsAux!CreValorCuota) Then .Cell(flexcpText, .Rows - 1, 12) = Format(RsAux!CreValorCuota, "#,##0.00")
            
            'Saldo de la Factura con T/C al día de hoy
            If RsAux!CreSaldoFactura > 0 Then
                .Cell(flexcpText, .Rows - 1, 13) = Format(RsAux!CreSaldoFactura * TasadeCambio(RsAux!DocMoneda, paMonedaFija, Now), "#,##0.00")
            Else
                .Cell(flexcpText, .Rows - 1, 13) = "0.00"
            End If
                        
            If Not IsNull(RsAux!CreGarantia) Then
                If RsAux!CreGarantia = Cliente Then      'Si Gtia es el cliente q' cargo    --> Garante
                    aTexto = "Garante"
                    If (RsAux!CreCliente = gConyuge And Cliente = gCliente) Or (RsAux!CreCliente = gCliente And Cliente = gConyuge) Then aTexto = "F. Ambos"
                    mValor = RsAux!CreCliente
                    
                Else
                    'Si la garantia es el conyuge --> F. Ambos
                    If (RsAux!CreGarantia = gConyuge And Cliente = gCliente And gTipoCliente = TipoCliente.Cliente) Or (RsAux!CreGarantia = gCliente And Cliente = gConyuge) Then
                        aTexto = "F. Ambos"
                    Else
                        Cons = "Select CliCiRuc, RelCodigo, RelNombre, CPeSexo from " & _
                                   " CPersona, Cliente Left Outer Join PersonaRelacion On CliCodigo = PReClienteEs And PReClienteDe =" & Cliente & _
                                                                " Left Outer Join Relaciones On PReRelacion = RelCodigo " & _
                                   " Where CliCodigo = " & RsAux!CreGarantia & _
                                   " And CliCodigo = CPeCliente"
                        Set RsAux1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                        If Not RsAux1.EOF Then
                            If Not IsNull(RsAux1!RelNombre) Then
                                aTexto = Trim(RsAux1!RelNombre)
                                If RsAux1!RelCodigo = paRelPadre Then
                                    Select Case RsAux1!CPeSexo
                                        Case "M": aTexto = "Padre"
                                        Case "F": aTexto = "Madre"
                                    End Select
                                End If
                            Else
                                aTexto = "CG sin C.I."
                                If Not IsNull(RsAux1!CliCiRuc) Then aTexto = clsGeneral.RetornoFormatoCedula(RsAux1!CliCiRuc)
                            End If
                        Else
                            aTexto = "#Eliminado"
                        End If
                        RsAux1.Close
                    End If
                    mValor = RsAux!CreGarantia
                    
                End If
            Else
                aTexto = "": mValor = 0
            End If
            .Cell(flexcpText, .Rows - 1, 8) = aTexto
            .Cell(flexcpData, .Rows - 1, 8) = mValor
            
            If Not IsNull(RsAux!CreUltimoPago) Then .Cell(flexcpText, .Rows - 1, 9) = Format(RsAux!CreUltimoPago, "dd/mm/yy")
        
            'Saco la Moneda original de la Factura
            If aMonedaAnteriorC <> RsAux!DocMoneda Then
                Cons = "Select * from Moneda Where MonCodigo = " & RsAux!DocMoneda
                Set RsAux1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                .Cell(flexcpText, .Rows - 1, 10) = Trim(RsAux1!MonSigno)
                aMonedaAnteriorT = Trim(RsAux1!MonSigno)
                aMonedaAnteriorC = RsAux1!MonCodigo
                RsAux1.Close
            Else
                .Cell(flexcpText, .Rows - 1, 10) = Trim(aMonedaAnteriorT)
            End If
            
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!DocTotal, "#,##0.00")
            If Not IsNull(RsAux!CreValorCuota) Then .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!CreValorCuota, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!CreSaldoFactura, "#,##0.00")
            
            .Cell(flexcpText, .Rows - 1, 14) = Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero
            .Cell(flexcpData, .Rows - 1, 14) = IIf(IsNull(RsAux!CrePuntaje) And Not (RsAux!CreSaldoFactura > 0), 1, 0)  '--> marco si tiene Nota asignada
        End With
        
        RsAux.MoveNext
    Loop
    If RsAux.EOF Then lOperacion.Tag = ""
    RsAux.Close
    
    '2) Cargo las Tasas de cambios para c/u de los documentos   -------------------------------------------------------------
    Dim idX As Integer
    If mSQLTCs <> "" Then
        Set RsAux = cBase.OpenResultset(mSQLTCs, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            With lOperacion
                For idX = colDesdeTCs + .FixedRows To .Rows - 1
                    If RsAux!ID = .Cell(flexcpData, idX, 0) Then
                        .Cell(flexcpText, idX, 11) = Format(.Cell(flexcpValue, idX, 11) * Format(RsAux!TCaComprador, "#.000"), "#,##0.00")
                        .Cell(flexcpText, idX, 12) = Format(.Cell(flexcpValue, idX, 12) * Format(RsAux!TCaComprador, "#.000"), "#,##0.00")
                        Exit For
                    End If
                Next
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    
    On Error Resume Next
    lOperacion.Redraw = True
    lOperacion.Select lOperacion.Rows - I, 0
    lOperacion.TopRow = lOperacion.Row
    
    'Recorro para sacar el Total de las compras de éstas operaciones    ------------------------------------------------------------------------------------
    Dim mTotalOps As Currency: mTotalOps = 0
    With lOperacion
        For idX = colDesdeTCs + .FixedRows To .Rows - 1
            mTotalOps = mTotalOps + .Cell(flexcpValue, idX, 11)
        Next
    End With
    
    If Val(lTotalComprado.Tag) = 0 Then
        lTotalComprado.Tag = mTotalOps
    Else
        mTotalOps = mTotalOps + CCur(lTotalComprado.Tag)
        lTotalComprado.Tag = mTotalOps
    End If
    
    lTotalComprado.Caption = lOperacion.Rows - 1 & " de " & IIf(lOperacion.Tag = "", lOperacion.Rows - 1, "...") & " Créditos."
    lTotalComprado.Caption = Trim(lTotalComprado.Caption) & " (" & Trim(paMonedaFijaTexto) & " " & Format(mTotalOps, "#,##0.00") & ")"
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    lOperacion.Redraw = True
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar las operaciones de crédito."
End Sub

Private Sub CargoDeudaCliente(Cliente As Long)

    lDeuda.Caption = ""
    If Cliente = 0 Then Exit Sub
    
    Dim aStr As String, curDolares As Currency, aDolares As Currency, aPesos As Currency
    
    aStr = MontoAdeudado(Cliente, curDolares)
    aDolares = curDolares
    If Trim(aStr) = "" Then aStr = " $ 0"
    lDeuda.Caption = "Deuda:" & aStr

    Select Case Cliente
        Case gCliente
            If gConyuge <> 0 Then
                aStr = MontoAdeudado(gConyuge, curDolares)
                aDolares = aDolares + curDolares
                If Trim(aStr) <> "" Then
                    lDeuda.Caption = Trim(lDeuda.Caption) & "  Cóny.:" & aStr
                
                    aPesos = aDolares / TasadeCambio(paMonedaPesos, paMonedaFija, Now)
                    lDeuda.Caption = Trim(lDeuda.Caption) & "  = Ambos: " & Format(aPesos, "#,##0")
                End If
            End If
        
        Case gConyuge
            If gCliente <> 0 Then
                aStr = MontoAdeudado(gCliente, curDolares)
                aDolares = aDolares + curDolares
                If Trim(aStr) <> "" Then
                    lDeuda.Caption = Trim(lDeuda.Caption) & "  Cóny.:" & aStr
                
                    aPesos = aDolares / TasadeCambio(paMonedaPesos, paMonedaFija, Now)
                    lDeuda.Caption = Trim(lDeuda.Caption) & " = Ambos: " & Format(aPesos, "#,##0")
                End If
            End If
    End Select
    
    If aDolares <> 0 Then lDeuda.Caption = Trim(lDeuda.Caption) & " (" & paMonedaFijaTexto & " " & Format(aDolares, "#,##0") & ")"
    
End Sub

Private Function MontoAdeudado(Cliente As Long, retDolares As Currency) As String

    MontoAdeudado = ""
    retDolares = 0
    'Saco la Deuda de creditos no cancelados para cada moneda
    Cons = "Select MonCodigo, MonSigno, Sum(CreSaldoFactura) Saldo From Documento, Credito, Moneda" _
           & " Where DocCliente = " & Cliente _
           & " And DocTipo = " & TipoDocumento.Credito _
           & " And DocAnulado = 0" _
           & " And CreSaldoFactura > 0" _
           & " And DocCodigo = CreFactura" _
           & " And DocMoneda = MonCodigo" _
           & " Group by MonCodigo, MonSigno"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        MontoAdeudado = MontoAdeudado & " " & Trim(RsAux!MonSigno) & " " & Format(RsAux!Saldo, "#,##0")
        retDolares = retDolares + (Format(RsAux!Saldo, "#,##0") * TasadeCambio(RsAux!MonCodigo, paMonedaFija, Now))
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Saco la Deuda de en cheques diferidos----------------------------------------------------------------------
    Cons = "Select MonCodigo, MonSigno, Sum(CDiImporte) Saldo From ChequeDiferido, Moneda" _
           & " Where CDiCliente = " & Cliente _
           & " And CDiDepositado Is Null And CDiEliminado Is Null And CDiVencimiento Is Not Null" _
           & " And CDiMoneda = MonCodigo" _
           & " Group by MonCodigo, MonSigno"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        MontoAdeudado = MontoAdeudado & " Ch." & Trim(RsAux!MonSigno) & " " & Format(RsAux!Saldo, "#,##0")
        retDolares = retDolares + (Format(RsAux!Saldo, "#,##0") * TasadeCambio(RsAux!MonCodigo, paMonedaFija, Now))
        RsAux.MoveNext
    Loop
    RsAux.Close
        
End Function

Private Sub NEWCargoDeudaCliente(Cliente As Long)

    lDeuda.Caption = ""
    If Cliente = 0 Then Exit Sub
    
    Dim aStr As String, curDolares As Currency, aDolares As Currency, aPesos As Currency
    Dim tmpT As Long, tmpC As Long, tmpCaso As Byte
    
    Select Case Cliente
        Case gCliente: tmpC = gConyuge: tmpT = gCliente: tmpCaso = 1
        Case gConyuge: tmpC = gCliente: tmpT = gConyuge: tmpCaso = 1
        Case gGarantia: tmpT = gGarantia: tmpC = 0: tmpCaso = 3
    End Select
    
    aStr = NEWMontoAdeudado(tmpT, tmpC, curDolares, tmpCaso)
    aDolares = curDolares
    If Trim(aStr) = "" Then aStr = " $ 0"
    lDeuda.Caption = "Deuda:" & aStr

    If tmpC <> 0 Then
        aStr = NEWMontoAdeudado(tmpT, tmpC, curDolares, 2)
        aDolares = aDolares + curDolares
        If Trim(aStr) <> "" Then
            lDeuda.Caption = Trim(lDeuda.Caption) & "  Cóny.:" & aStr
        
            aPesos = aDolares / TasadeCambio(paMonedaPesos, paMonedaFija, Now)
            lDeuda.Caption = Trim(lDeuda.Caption) & "  = Ambos: " & Format(aPesos, "#,##0")
        End If
    End If
    
    If aDolares <> 0 Then lDeuda.Caption = Trim(lDeuda.Caption) & " (" & paMonedaFijaTexto & " " & Format(aDolares, "#,##0") & ")"
    
End Sub

'Private Function MontoAdeudado(Cliente As Long, retDolares As Currency) As String
Private Function NEWMontoAdeudado(xIDTitular As Long, xIDConyuge As Long, retDolares As Currency, xCasoT_C_G As Byte) As String
'xCasoT_C_G ->  caso Titular , Conyuge, Garantia

    NEWMontoAdeudado = ""
    retDolares = 0
    'Saco la Deuda de creditos no cancelados para cada moneda
    Cons = "Select MonCodigo, MonSigno, Sum(CreSaldoFactura) Saldo From Documento, Credito, Moneda"
    
    Select Case xCasoT_C_G
        Case 1
                Cons = Cons & " Where ( DocCliente = " & xIDTitular & " OR CreGarantia = " & xIDTitular & ") "
        Case 2
                Cons = Cons & " Where ( ( DocCliente = " & xIDConyuge & " AND (CreGarantia <> " & xIDTitular & " OR CreGarantia Is Null) ) " & _
                                    " OR ( DocCliente <> " & xIDTitular & " AND CreGarantia = " & xIDConyuge & ") )"
        Case 3
                Cons = Cons & " Where ( DocCliente = " & xIDTitular & " OR CreGarantia = " & xIDTitular & ") "
    End Select
           
     Cons = Cons _
           & " And DocTipo = " & TipoDocumento.Credito _
           & " And DocAnulado = 0" _
           & " And CreSaldoFactura > 0" _
           & " And DocCodigo = CreFactura" _
           & " And DocMoneda = MonCodigo" _
           & " Group by MonCodigo, MonSigno"
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        NEWMontoAdeudado = NEWMontoAdeudado & " " & Trim(RsAux!MonSigno) & " " & Format(RsAux!Saldo, "#,##0")
        retDolares = retDolares + (Format(RsAux!Saldo, "#,##0") * TasadeCambio(RsAux!MonCodigo, paMonedaFija, Now))
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Saco la Deuda de en cheques diferidos----------------------------------------------------------------------
    Cons = "Select MonCodigo, MonSigno, Sum(CDiImporte) Saldo From ChequeDiferido, Moneda" _
           & " Where CDiCliente = " & Cliente _
           & " And CDiDepositado Is Null And CDiEliminado Is Null And CDiVencimiento Is Not Null" _
           & " And CDiMoneda = MonCodigo" _
           & " Group by MonCodigo, MonSigno"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        NEWMontoAdeudado = NEWMontoAdeudado & " Ch." & Trim(RsAux!MonSigno) & " " & Format(RsAux!Saldo, "#,##0")
        retDolares = retDolares + (Format(RsAux!Saldo, "#,##0") * TasadeCambio(RsAux!MonCodigo, paMonedaFija, Now))
        RsAux.MoveNext
    Loop
    RsAux.Close
        
End Function

Private Sub ArmoConsultaDeCreditos(Cliente As Long)

    Dim mSelect As String
    mSelect = " DocCodigo, DocFecha, DocNumero, DocSerie, DocMoneda, DocTotal, " & _
                   " CreSaldoFactura, CreTipo, CreFormaPago, CreValorCuota, CreProximoVto, CrePuntaje, CreVaCuota, CreDeCuota, CreGarantia, CreCliente, CreUltimoPago," & _
                   " TCuAbreviacion , TCuVencimientoE, TCuVencimientoC, TCuDistancia, dbo.Cumplimiento(CreCumplimiento, CreProximoVto, GETDATE()) as CreCumplimiento"
                   
    If bFTitular Then
        Cons = "Select Top " & PorPagina + 1 & mSelect & " From Documento (index = iClienteTipo), Credito, TipoCuota" _
               & " Where DocCliente = " & Cliente _
               & " And DocTipo = " & TipoDocumento.Credito _
               & " And DocAnulado = 0" _
               & " And DocCodigo = CreFactura" _
               & " And CreTipoCuota = TCuCodigo"
        If bFVigentes Then Cons = Cons & " And CreSaldoFactura > 0"
        If bFCanceladas Then Cons = Cons & " And CreSaldoFactura = 0"
        If Trim(lOperacion.Tag) <> "" Then Cons = Cons & " And DocFecha < '" & Format(lOperacion.Tag, sqlFormatoFH) & "'"
        Cons = Cons & " Order by DocFecha Desc"
        Exit Sub
    End If
    
    If bFGarante Then
        Cons = " Select Top " & PorPagina + 1 & mSelect & " From Credito (index = iGarantia), Documento , TipoCuota" _
                & " Where CreGarantia = " & Cliente _
                & " And CreFactura = DocCodigo" _
                & " And CreTipoCuota = TCuCodigo" _
                & " And DocTipo = " & TipoDocumento.Credito _
                & " And DocAnulado = 0"

        If bFVigentes Then Cons = Cons & " And CreSaldoFactura > 0"
        If bFCanceladas Then Cons = Cons & " And CreSaldoFactura = 0"
        If Trim(lOperacion.Tag) <> "" Then Cons = Cons & " And DocFecha < '" & Format(lOperacion.Tag, sqlFormatoFH) & "'"
        Cons = Cons & " Order by DocFecha Desc"
        Exit Sub
    End If
        
    Cons = "Select Top " & PorPagina + 1 & mSelect & " From Documento, Credito, TipoCuota" _
           & " Where (CreCliente = " & Cliente & " Or CreGarantia = " & Cliente & ") " _
           & " And DocTipo = " & TipoDocumento.Credito _
           & " And DocAnulado = 0" _
           & " And DocCodigo = CreFactura" _
           & " And CreTipoCuota = TCuCodigo"
    If bFVigentes Then Cons = Cons & " And CreSaldoFactura > 0"
    If bFCanceladas Then Cons = Cons & " And CreSaldoFactura = 0"
    If Trim(lOperacion.Tag) <> "" Then Cons = Cons & " And DocFecha < '" & Format(lOperacion.Tag, sqlFormatoFH) & "'"
    
    Cons = Cons & " Order by DocFecha Desc"
    
End Sub

Private Sub CargoOpSolicitud(Cliente As Long, Optional bNext As Boolean = False)
    
    If Not bNext Or lOperacion.Tag = "" Then
        If bNext And lOperacion.Tag = "" Then Exit Sub
        lOperacion.Rows = 1: lOperacion.Refresh
    End If
    If Cliente = 0 Or oSolicitud.Tag = 0 Then Exit Sub

Dim aMonto As Currency
Dim aFecha As Date
Dim aMoneda As Long
    
    On Error GoTo errCargar
    I = 0
    Screen.MousePointer = 11
    lOperacion.Redraw = False
    
    Cons = "Select Top " & PorPagina + 1 & " SolCodigo, SolTipo, SolEstado, SolFecha, SolMoneda, ResComentario, RSoDocumento, TCuCodigo, TCuAbreviacion, TCuCantidad, " & _
                                                               " Sum(RSoValorCuota) as RSoValorCuota, Sum(IsNull(RSoValorEntrega, 0)) as RSoValorEntrega" & _
                " From Solicitud(Index = iCliente)" & _
                        " Left Outer Join SolicitudResolucion ON SolCodigo = ResSolicitud, " & _
                        " RenglonSolicitud , TipoCuota" & _
                " Where SolCliente = " & Cliente
        
   If Trim(lOperacion.Tag) <> "" Then Cons = Cons & " And SolFecha < '" & Format(lOperacion.Tag, sqlFormatoFH) & "'"
   
    Cons = Cons & _
            " And SolCodigo = RSoSolicitud" & _
            " And RSoTipoCuota = TCuCodigo" & _
            " And ResNumero = (Select Max(ResNumero) From SolicitudResolucion Where SolCodigo = ResSolicitud)" & _
            " Group by SolCodigo, SolTipo, SolEstado, SolFecha, SolMoneda, ResComentario, RSoDocumento, TCuCodigo, TCuAbreviacion, TCuCantidad " & _
            " Order by SolFecha DESC"
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF And I < PorPagina
        I = I + 1
        With lOperacion
            .Tag = RsAux!SolFecha
            
            .AddItem RsAux!SolCodigo
            
            Select Case RsAux!SolEstado
                Case EstadoSolicitud.Rechazada: .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.RojoClaro: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
                Case EstadoSolicitud.Condicional: .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio
            End Select
        
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!SolFecha, "dd/mm/yy")
            Select Case RsAux!SolTipo
                Case TipoSolicitud.AlMostrador: .Cell(flexcpText, .Rows - 1, 2) = "Mos."
                Case TipoSolicitud.Reserva: .Cell(flexcpText, .Rows - 1, 2) = "Tel."
                Case TipoSolicitud.Servicio: .Cell(flexcpText, .Rows - 1, 2) = "Ser"
            End Select
            
            .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!TCuAbreviacion)
        
            If Not IsNull(RsAux!RSoDocumento) Then .Cell(flexcpText, .Rows - 1, 5) = "Si" Else .Cell(flexcpText, .Rows - 1, 5) = "No"
            If Not IsNull(RsAux!ResComentario) Then .Cell(flexcpText, .Rows - 1, 6) = Trim(RsAux!ResComentario)
            
            aMoneda = RsAux!SolMoneda
            aFecha = RsAux!SolFecha
            aMonto = 0
            
            If IsNull(RsAux!RSoValorEntrega) Then   '------------------------------------------------------
                aMonto = aMonto + RsAux!RSoValorCuota * RsAux!TCuCantidad
            Else
                aMonto = aMonto + RsAux!RSoValorEntrega + (RsAux!RSoValorCuota * RsAux!TCuCantidad)
            End If
            
            If paMonedaFija <> 0 Then
                .Cell(flexcpText, .Rows - 1, 3) = Format(aMonto * TasadeCambio(aMoneda, paMonedaFija, aFecha), "#,##0.00")
            Else
                .Cell(flexcpText, .Rows - 1, 3) = Format(aMonto, "#,##0.00")
            End If
            
            RsAux.MoveNext
        End With
    Loop
    
    If RsAux.EOF Then lOperacion.Tag = "" 'Else lOperacion.AddItem ("Mas .........")
    RsAux.Close
    
    On Error Resume Next
    lOperacion.Redraw = True
    lOperacion.Select lOperacion.Rows - I, 0
    lOperacion.TopRow = lOperacion.Row
    
    If lOperacion.Tag = "" Then
        lTotalComprado.Caption = lOperacion.Rows - 1 & " de " & lOperacion.Rows - 1 & "  Solicitudes."
    Else
        lTotalComprado.Caption = lOperacion.Rows - 1 & " de ..." & "  Solicitudes."
    End If

    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    lOperacion.Redraw = True
    clsGeneral.OcurrioError "Error al cargar las solicitudes realizadas.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoOpArticulo(Cliente As Long, Optional bNext As Boolean = False)
    
    If Trim(Articulos) = "" Then Articulos = CargoArticulosDeFlete
    
    If Not bNext Or lOperacion.Tag = "" Then
        If bNext And lOperacion.Tag = "" Then Exit Sub
        lOperacion.Rows = 1: lOperacion.Refresh
    End If
    If Cliente = 0 Then Exit Sub
    
    On Error GoTo errCargar
    'Cuento las cantidades de los articulos -----------------------------------------------------------------------
    If Trim(oArticulo.Tag) = "" Then
        Dim aQArt As Long: aQArt = 0
        
        Cons = "Select  DocTipo, Sum(RenCantidad) as Q" _
               & " From Documento (index = iClienteTipo), Renglon" _
               & " Where DocCliente = " & Cliente _
               & " And DocTipo IN ( " & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaCredito & ", " & TipoDocumento.NotaEspecial & ")" _
               & " And DocAnulado = 0" _
               & " And DocCodigo = RenDocumento" _
               & " And RenArticulo Not IN (" & Mid(Articulos, 1, Len(Articulos) - 1) & ")" _
               & " Group by DocTipo"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            Select Case RsAux!DocTipo
                Case TipoDocumento.Contado, TipoDocumento.Credito: aQArt = aQArt + RsAux!Q
                Case Else: aQArt = aQArt - RsAux!Q
            End Select
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        oArticulo.Tag = aQArt: oArticulo.Caption = "&Artículos (" & aQArt & ")"
    End If
    '------------------------------------------------------------------------------------------------------------------
    
    Dim mValor As Long
    lOperacion.Redraw = False
    Screen.MousePointer = 11
    I = 0
    Cons = "Select TOP " & PorPagina + 1 & " DocCodigo, DocFecha, DocTipo, DocNumero, DocSerie, ArtID, ArtNombre, RenCantidad" _
           & " From Documento (index = iClienteTipo), Renglon, Articulo" _
           & " Where DocCliente = " & Cliente _
           & " And DocTipo IN ( " & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaCredito & ", " & TipoDocumento.NotaEspecial & ")" _
           & " And DocAnulado = 0" _
           & " And DocCodigo = RenDocumento" _
           & " And RenArticulo = ArtID" _
           & " And RenArticulo Not IN (" & Mid(Articulos, 1, Len(Articulos) - 1) & ")"
           
    If Trim(lOperacion.Tag) <> "" Then
        Dim arrKey() As String
        arrKey = Split(Trim(lOperacion.Tag), "|")
        Cons = Cons & " And DocFecha <= '" & Format(arrKey(0), sqlFormatoFH) & "'" & _
                            " And RenArticulo > " & arrKey(1)
                            
    End If
    Cons = Cons & " Order by DocFecha DESC, ArtID ASC"
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF And I < PorPagina
        With lOperacion
            .Tag = RsAux!DocFecha & "|" & RsAux!ArtID
        
            I = I + 1
            .AddItem Format(RsAux!DocFecha, "dd/mm/yy")
        
            'Veo si es Devolucion para ponerle el icono
            mValor = -1
            Select Case RsAux!DocTipo
                Case TipoDocumento.Contado: .Cell(flexcpText, .Rows - 1, 1) = "Contado": mValor = TipoDocumento.Contado
                
                Case TipoDocumento.NotaDevolucion
                    .Cell(flexcpText, .Rows - 1, 1) = "Contado"
                    .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("Nota").ExtractIcon
                    
                Case TipoDocumento.Credito: .Cell(flexcpText, .Rows - 1, 1) = "Crédito": mValor = TipoDocumento.Credito
                    
                Case TipoDocumento.NotaCredito
                    .Cell(flexcpText, .Rows - 1, 1) = "Crédito"
                    .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("Nota").ExtractIcon
                
                Case TipoDocumento.NotaEspecial
                    .Cell(flexcpText, .Rows - 1, 1) = "N.Especial"
                    .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("Nota").ExtractIcon
            End Select
            
            .Cell(flexcpData, .Rows - 1, 1) = mValor
            mValor = RsAux!DocCodigo: .Cell(flexcpData, .Rows - 1, 0) = mValor
            
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!DocSerie) & RsAux!DocNumero
            
            If RsAux!RenCantidad = 1 Then
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ArtNombre)
            Else
                .Cell(flexcpText, .Rows - 1, 3) = RsAux!RenCantidad & " " & Trim(RsAux!ArtNombre)
            End If
            RsAux.MoveNext
        End With
    Loop
    If RsAux.EOF Then lOperacion.Tag = "" 'Else lOperacion.AddItem "Mas ........."
    
    RsAux.Close
    
    On Error Resume Next
    lOperacion.Redraw = True
    lOperacion.Select lOperacion.Rows - I, 0
    lOperacion.TopRow = lOperacion.Row
    
    If lOperacion.Tag = "" Then
        lTotalComprado.Caption = lOperacion.Rows - 1 & " de " & lOperacion.Rows - 1 & "  Artículos."
    Else
        lTotalComprado.Caption = lOperacion.Rows - 1 & " de ..." & "  Artículos."
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    lOperacion.Redraw = True
    clsGeneral.OcurrioError "Error al cargar los artículos comprados.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoOpVtaTelefonica(Cliente As Long, Optional bNext As Boolean = False)
 
    If Not bNext Or lOperacion.Tag = "" Then
        If bNext And lOperacion.Tag = "" Then Exit Sub
        lOperacion.Rows = 1: lOperacion.Refresh
    End If
    If Cliente = 0 Then Exit Sub
    
    On Error GoTo errVta
    Screen.MousePointer = 11
    lOperacion.Redraw = False

Dim aTotalComprado As Currency

    aMonedaAnteriorC = 0
    aTotalComprado = 0
    I = 0
    
    Cons = "Select VentaTelefonica.*, RVTCantidad, ArtNombre, DocNumero, DocSerie, DocAnulado" _
           & " From VentaTelefonica Left Outer Join Documento On VTeDocumento = DocCodigo, RenglonVtaTelefonica, Articulo" _
           & " Where VTeCliente = " & Cliente _
           & " And VTeTipo  = " & TipoDocumento.ContadoDomicilio _
           & " And VTeCodigo = RVTVentaTelefonica" _
           & " And RVTArticulo = ArtID"
    
    If Trim(lOperacion.Tag) <> "" Then Cons = Cons & " And VTeFechaLlamado < '" & Format(lOperacion.Tag, sqlFormatoFH) & "'"
    Cons = Cons & " Order by VTeFechaLlamado Desc"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF And I < PorPagina
        I = I + 1
        ' "<ID|Llamada|Facturada|>Importe|<Artículos|<Comentarios|Pendiente|Tel. Llamada"
        With lOperacion
            .Tag = RsAux!VTeFechaLlamado
        
            .AddItem CStr(RsAux!VTeCodigo)
            mValor = RsAux!VTeCodigo: .Cell(flexcpData, .Rows - 1, 0) = mValor  'Guardo el codigo del documento
            
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!VTeFechaLlamado, "dd/mm/yy hh:mm")
            
            If Not IsNull(RsAux!DocSerie) Then
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!DocSerie) & " " & RsAux!DocNumero
                mValor = RsAux!VTeDocumento: .Cell(flexcpData, .Rows - 1, 2) = mValor  'Guardo el codigo del documento
            End If
            
            If Not IsNull(RsAux!DocAnulado) Then If RsAux!DocAnulado = True Then .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
        
            'Saco la Moneda
            If aMonedaAnteriorC <> RsAux!VTeMoneda Then
                Cons = "Select * from Moneda Where MonCodigo = " & RsAux!VTeMoneda
                Set rsMon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                aMonedaAnteriorT = Trim(rsMon!MonSigno)
                aMonedaAnteriorC = rsMon!MonCodigo
                rsMon.Close
            End If
        
            If paMonedaFija <> 0 And RsAux!VTeMoneda <> paMonedaFija Then
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!VTeTotal * TasadeCambio(RsAux!VTeMoneda, paMonedaFija, RsAux!VTeFechaLlamado), "#,##0.00")
            Else
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!VTeTotal, "#,##0.00")
            End If
            If IsNull(RsAux!VTeAnulado) And (IsNull(RsAux!DocAnulado) Or RsAux!DocAnulado = False) Then
                aTotalComprado = aTotalComprado + .Cell(flexcpValue, .Rows - 1, 3)
            End If
            .Cell(flexcpText, .Rows - 1, 3) = "(" & aMonedaAnteriorT & ") " & .Cell(flexcpText, .Rows - 1, 3)
            '-------------------------------------------------------------------------------------------------------------------------
            
            If Not IsNull(RsAux!VTeComentario) Then .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!VTeComentario)
            If Not IsNull(RsAux!VTeTelefonoLlamada) Then .Cell(flexcpText, .Rows - 1, 7) = Trim(RsAux!VTeTelefonoLlamada)
            'If Not IsNull(rsAux!VTePendiente) Then
            '    cons = "Select * from PendienteEntrega Where PEnCodigo = " & rsAux!VTePendiente
            '    Set rsMon = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            '    If Not rsMon.EOF Then .Cell(flexcpText, .Rows - 1, 6) = Trim(rsMon!PEnNombre)
            '    rsMon.Close
            '    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio
            'End If
            
            If Not IsNull(RsAux!VTeAnulado) Then
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
                .Cell(flexcpText, .Rows - 1, 9) = Format(RsAux!VTeAnulado, "dd/mm/yy hh:mm")
            End If
                        
            .Cell(flexcpText, .Rows - 1, 8) = z_BuscoUsuario(RsAux!VTeUsuario, Identificacion:=True)
            'Concateno los articulos
            aDoc = RsAux!VTeCodigo
            aTexto = ""
            Do While RsAux!VTeCodigo = aDoc
                If RsAux!RVTCantidad = 1 Then
                    aTexto = aTexto & Trim(RsAux!ArtNombre) & ", "
                Else
                    aTexto = aTexto & RsAux!RVTCantidad & " " & Trim(RsAux!ArtNombre) & ", "
                End If
                
                RsAux.MoveNext
                If RsAux.EOF Then Exit Do
            Loop
            .Cell(flexcpText, .Rows - 1, 4) = Mid(aTexto, 1, Len(aTexto) - 2)
        End With
    Loop
    If RsAux.EOF Then lOperacion.Tag = "" 'Else lOperacion.AddItem "Mas ........."
    RsAux.Close
    
    'If aTotalComprado <> 0 Then lTotalComprado.Caption = "Operaciones x " & Trim(paMonedaFijaTexto) & " " & Format(aTotalComprado, "#,##0.00")
    On Error Resume Next
    lOperacion.Redraw = True
    lOperacion.Select lOperacion.Rows - I, 0
    lOperacion.TopRow = lOperacion.Row
    
    If lOperacion.Tag = "" Then
        lTotalComprado.Caption = lOperacion.Rows - 1 & " de " & lOperacion.Rows - 1 & "  Ventas"
    Else
        lTotalComprado.Caption = lOperacion.Rows - 1 & " de ..." & "  Ventas"
    End If

    If aTotalComprado <> 0 Then
        If Val(lTotalComprado.Tag) <> 0 Then aTotalComprado = CCur(lTotalComprado.Tag) + aTotalComprado
        lTotalComprado.Tag = aTotalComprado
        lTotalComprado.Caption = lTotalComprado.Caption & " (" & Trim(paMonedaFijaTexto) & " " & Format(aTotalComprado, "#,##0.00") & ")"
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errVta:
    lOperacion.Redraw = True
    clsGeneral.OcurrioError "Ocurrió un error al cargar las operaciones contado.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub LimpioFicha(Optional ClienteACero As Boolean = False)
On Error Resume Next
    
    Tab1.Tabs("relaciones").Image = ImageList1.ListImages("relacionesno").Index
    shpTitular.BackColor = Colores.Inactivo
    
    If ClienteACero Then gCliente = 0: Me.Caption = "Estudio de Solicitudes"
    gConyuge = 0: gGarantia = 0
    Articulos = ""
    
    'Tag del las pictures en ""---------------------
    For I = 0 To Fichas: Picture1(I).Tag = "": Next
    Picture1(0).ZOrder 0

    'Captions de las Tabs--------------------------
    Tab1.Tabs("garantia").Caption = "&Garantía"
    Tab1.Tabs("conyuge").Caption = "Cón&yuge"
    
    With TabPersona
        .Tabs.Clear
        .Tabs.Add pvKey:="titular", pvCaption:="Titular", pvImage:=ImageList1.ListImages("titular").Index
        .Tabs("titular").Selected = True
    End With
    
    'Datos del Cliente-------------------------------
    lTitular.Caption = "N/D"
    cEMailT.ClearObjects: cEMailT.BackColor = shpTitular.BackColor
    
    PresentoPaisDelDocumentoCliente 1, "", lblInfoCI, lCiRuc
    
    lCategoria.Caption = "N/D": lCategoria.ForeColor = Colores.Azul
    lEdad.Caption = "N/D": lEdad.ForeColor = Colores.Azul
    lOcupacion.Caption = "N/D"
    lDireccion.Caption = "N/D": lDireccion.Tag = 0
    lECivil.Caption = "N/D"
    lTitulo.Caption = ""
    lTelefono.Caption = "N/D"
    lNSolicitud.Caption = "N/D"
    lCheques.Caption = ""
    cDireccion.Clear: cDireccion.BackColor = Colores.Obligatorio: cDireccion.Visible = False: cDireccion.Tag = 0
    lNDireccion.Caption = "Dirección Principal:"
    lDireccion.Left = lNDireccion.Left + lNDireccion.Width + 40
    
    'Datos de Operaciones--------------------------
    EncabezadoOperaciones True, False, False, False
    lOperacion.Rows = 1
    oCredito.value = True
    TabPersona.Tabs("titular").Selected = True
    Tab1.Tabs("operaciones").Selected = True
    
    'Datos de la Solicitud---------------------------
    lFecha.Caption = "N/D"
    lPago.Caption = "N/D"
    lSucursal.Caption = "N/D"
    lMoneda.Caption = "": lMoneda2.Caption = ""
    lArticulo.Rows = 1
    lComentario.Caption = "N/D": lAFinanciar.Caption = ""
    lUsuario.Caption = "N/D"
    lArticulo.Rows = 1
    
    lFEntrega.Caption = "0.00": lFCuota.Caption = "0.00": lFMonto.Caption = "0.00"
    lSEfectivo.Caption = "0.00": lSFinanciado.Caption = "0.00": lSMonto.Caption = "0.00"
    lFinanciacion.Caption = "N/D"
    
    'Informacion del Cliente------------------------
    lEmpleoT.Rows = 1: lComentarioT.Rows = 1: lReferenciaT.Rows = 1
           
    LimpioFichaConGar
    
    'Limpio los campos de la ficha clearing-------------------------------------------
    shpClearing.BackColor = Colores.Gris
    lClFecha.Caption = ""
    lClNombre.Caption = ""
    lClConyuge.Caption = ""
    lClTelefono.Caption = ""
    lClDireccion.Caption = ""
    lClFNacimiento.Caption = ""
    lClEmpleo.Caption = ""
    lClECivil.Caption = ""
    lClDireccionE.Caption = ""
    lClAntiguedad.Caption = ""
    lClAlquiler.Caption = ""
    
    lCSolicitud.Rows = 1
    
    TabClearing.Tabs.Clear
    
    With TabCPersona
        .Tabs.Clear
        .Tabs.Add pvKey:="titular", pvCaption:="Titular", pvImage:=ImageList1.ListImages("titular").Index
        .Tabs("titular").Selected = True
    End With

    ZDeshabilitoCondiciones
    
    lCondicion.Rows = 1
    tComentario.Text = ""
    
    frmCondicion.Tag = ""
    cCondicionR.ListIndex = -1
    
    vsRelacion.Rows = 1
    
End Sub

Private Sub LimpioFichaConGar()

    'Datos de la Garantia Y Conyuge------------------------------
    shpGarantia.BackColor = Colores.Inactivo
    
    lGarantia.Caption = "N/D"
    'lCIG.Caption = "N/D"
    PresentoPaisDelDocumentoCliente 1, "", lblInfoCIConyuge, lCIG
    
    lCategoriaG.Caption = "N/D": lCategoriaG.ForeColor = Colores.Azul
    lEdadG.Caption = "N/D": lEdadG.ForeColor = Colores.Azul
    lOcupacionG.Caption = "N/D"
    lDireccionG.Caption = "N/D": lDireccionG.Tag = 0
    lECivilG.Caption = "N/D"
    
    lEmpleoG.Rows = 1: lReferenciaG.Rows = 1: lComentarioG.Rows = 1
    lTituloG.Caption = ""
    lNSolicitudG.Caption = "N/D"
    lChequesG.Caption = ""
    cDireccionG.Clear: cDireccionG.BackColor = Colores.Obligatorio: cDireccionG.Visible = False: cDireccionG.Tag = 0
    lNDireccionG.Caption = "Dirección Principal:"
    lDireccionG.Left = lNDireccionG.Left + lNDireccionG.Width + 40

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error GoTo errActualizar
    Screen.MousePointer = 11
    'Si la Solicitud no se resolvio le quito la marca de Analizando
    If Trim(tCodigo.Text) <> "" And IsNumeric(tCodigo.Text) Then '-----------------------------------------------------------------
        FechaDelServidor
        
        Cons = "Select * from Solicitud Where SolCodigo = " & Trim(tCodigo.Text)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If IsNull(RsAux!SolFResolucion) Then
                If RsAux!SolUsuarioR = paCodigoDeUsuario And RsAux!SolEstado = EstadoSolicitud.Pendiente Then 'USO SolUsuarioR para marca de analizando
                    RsAux.Edit
                    If prmVistoPor <> 0 Then
                        RsAux!SolUsuarioR = prmVistoPor
                        RsAux!SolEstado = EstadoSolicitud.ParaRetomar
                    Else
                        RsAux!SolUsuarioR = Null
                    End If
                    
                    Dim tmSeg As Long
                    If Not IsNull(RsAux!SolTiempoResolucion) Then tmSeg = RsAux!SolTiempoResolucion Else tmSeg = 0
                    tmSeg = tmSeg + DateDiff("s", dStartTime, gFechaServidor)
                    If tmSeg > 32000 Then tmSeg = 32000
                    RsAux!SolTiempoResolucion = tmSeg
            
                    RsAux.Update
                End If
            Else
                MsgBox "La solicitud ya ha sido resuelta por otro usuario.", vbExclamation, "Solicitud Resuelta"
            End If
        End If
        RsAux.Close
        
        On Error Resume Next
        frmLista.signalR_RefrescoSolicitudes
    End If
    '-------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    Exit Sub
    
errActualizar:
    clsGeneral.OcurrioError "Error al quitar la marca de Analizando.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
Dim aAuxiliar As Currency
    
    On Error Resume Next
            
    'SETEO LOS OBJETOS EN PANTALLA----------------------------------------------------------
    Tab1.Top = shpTitular.Top + shpTitular.Height + 40
    Tab1.Left = 120
    Tab1.Height = Me.ScaleHeight - Tab1.Top - Status.Height - 50
    Tab1.Width = Me.ScaleWidth - 240
    shpTitular.Width = Tab1.Width: lAFinanciar.Width = shpTitular.Width '+ 20
    
    lFinanciacion.Width = lAFinanciar.Width / 2
    
    'Fichas del Tab Principal-----------------------------------------------------
    For I = 0 To Fichas
        Picture1(I).Top = Tab1.Top + 350
        Picture1(I).Left = Tab1.Left + 40
        Picture1(I).Height = Tab1.Height - 400
        Picture1(I).Width = Tab1.Width - 80
        Picture1(I).BorderStyle = 0
    Next
    
    'Ficha de Operaciones---------------------------------------------------------
    lPieOp.Top = Picture1(0).Height - lPieOp.Height
    lPieOp.Left = 120: lPieOp.Width = Picture1(0).Width - (lPieOp.Left * 2):
    lTotalComprado.Top = lPieOp.Top: lDeuda.Top = lPieOp.Top
    lTotalComprado.Left = lPieOp.Left
    lTotalComprado.Left = lPieOp.Left + lPieOp.Width - (lTotalComprado.Width) - 80
    
    lOperacion.Top = 550
    lOperacion.Width = Picture1(0).Width - (lOperacion.Left * 2)
    lOperacion.Height = Picture1(0).Height - lOperacion.Top - lDeuda.Height - 20
    '----------------------------------------------------------------------------------
    
    'Ficha de Solicitud--------------------------------------------------------------
    lArticulo.Width = Picture1(0).Width - (lArticulo.Left * 2)
    lArticulo.Height = Picture1(0).Height - lArticulo.Top - 20
    shpSolicitud.Width = lArticulo.Width
    lComentario.Width = lArticulo.Width
    '----------------------------------------------------------------------------------
    
    'Ficha del Titular----------------------------------------------------------------
    aAuxiliar = Picture1(0).Height - 160 - (40 * 2)
    With lEmpleoT
        .Top = 40: .Width = Picture1(0).ScaleWidth - (.Left * 2)
        .Height = aAuxiliar / 3
        .ColWidth(2) = .ClientWidth - (.ColWidth(0) + .ColWidth(1) + .ColWidth(3) + .ColWidth(4) + 80)
    End With
    With lReferenciaT
        .Top = lEmpleoT.Top + lEmpleoT.Height + 40
        .Width = lEmpleoT.Width
        .Height = lEmpleoT.Height
        .ColWidth(2) = .ClientWidth - (.ColWidth(0) + .ColWidth(1) + .ColWidth(3) + 80)
    End With
    With lComentarioT
        .Top = lReferenciaT.Top + lReferenciaT.Height + 40
        .Width = lEmpleoT.Width
        .Height = lEmpleoT.Height
        .ColWidth(1) = .ClientWidth - (.ColWidth(0) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + 80)
    End With
    
    'Ficha de la Garantia Y Conyuge----------------------------------------------------------------
    shpGarantia.Width = Picture1(0).Width - (lEmpleoG.Left * 2)
    aAuxiliar = Picture1(0).Height - (shpGarantia.Top + shpGarantia.Height) - 160 - (40 * 2)
    
    With lEmpleoG
        .Top = shpGarantia.Top + shpGarantia.Height + 40
        .Width = Picture1(0).Width - (lEmpleoG.Left * 2)
        .Height = aAuxiliar / 3
        .ColWidth(2) = .ClientWidth - (.ColWidth(0) + .ColWidth(1) + .ColWidth(3) + .ColWidth(4) + 80)
    End With
    With lReferenciaG
        .Top = lEmpleoG.Top + lEmpleoG.Height + 40
        .Width = lEmpleoG.Width
        .Height = lEmpleoG.Height
        .ColWidth(2) = .ClientWidth - (.ColWidth(0) + .ColWidth(1) + .ColWidth(3) + 80)
    End With
    With lComentarioG
        .Top = lReferenciaG.Top + lReferenciaG.Height + 40
        .Width = lEmpleoG.Width
        .Height = lEmpleoG.Height
        .ColWidth(1) = .ClientWidth - (.ColWidth(0) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + 80)
    End With
    '----------------------------------------------------------------------------------
    
    'Clearing-------------------------------------------------------------------------
    TabClearing.Top = 1800
    TabClearing.Height = Tab1.Height - TabClearing.Top - 440
    TabClearing.Width = Tab1.Width - 320
    For I = 0 To pClearing.Count - 1
        pClearing(I).Top = TabClearing.Top + 350
        pClearing(I).Left = TabClearing.Left + 40
        pClearing(I).Height = TabClearing.Height - 400
        pClearing(I).Width = TabClearing.Width - 80
        pClearing(I).BorderStyle = 0
    Next
    
    'Listas del Clearing ---------------------------------------------------------
    shpClearing.Width = Tab1.Width - 320
    lCSolicitud.Width = pClearing(0).Width - 40
    lCSolicitud.Height = pClearing(0).Height - 40
    '--------------------------------------------------------------------------------
    
    'Ficha de Resolucion----------------------------------------------------------
    lEmpleoT.Top = 40
    lCondicion.Width = frmCondicion.Width
    lCondicion.Height = Picture1(0).Height - lCondicion.Top - tComentario.Height - 140
    '--------------------------------------------------------------------------------
    
    With vsRelacion     'Relaciones
        .Top = 40: .Left = 120
        .Width = Picture1(0).ScaleWidth - (.Left * 2)
        .Height = Picture1(0).ScaleHeight - (.Top * 2)
    End With
    
    tbTool.Buttons("sepsalir").Width = Me.ScaleWidth - (tbTool.Buttons("plantillas").Left + tbTool.ButtonWidth) - tbTool.ButtonWidth * 1.5
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    Set objClearing = Nothing
    
End Sub

Private Sub bDireccionT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    aAccionDireccion = 1
    If gCliente <> 0 Then PopupMenu MnuOpDireccion, , , , MnuOpDiTitulo
End Sub

Private Sub imgCondRes_Click(Index As Integer)
On Error GoTo errMnuR
    If imgCondRes(Index).Tag = "" Then Exit Sub
    
    Dim iX As Integer, mItems() As String, mVal() As String

    For iX = mnuCondicionX.UBound To (mnuCondicionX.LBound + 1) Step -1
        Unload mnuCondicionX(iX)
    Next
        
    mItems = Split(imgCondRes(Index).Tag, "·")
    For iX = LBound(mItems) To UBound(mItems)
        If iX > 0 Then Load mnuCondicionX(iX)
        mVal = Split(mItems(iX), "|")
        mnuCondicionX(iX).Tag = mVal(0)
        mnuCondicionX(iX).Caption = mVal(1)
    Next
    PopupMenu mnuMCondicion
    
    Exit Sub
errMnuR:
    clsGeneral.OcurrioError "Error al cargar el menú con las condiciones.", Err.Description
End Sub

Private Sub lArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errDelete

    If KeyCode = vbKeyDelete Then
        If gIdSolicitud = 0 Then Exit Sub
        
        If Not (lArticulo.Row >= lArticulo.FixedRows) Then Exit Sub
        If (lArticulo.Rows - lArticulo.FixedRows - 1) = 1 Then Exit Sub
        
        Dim mIDArticuloToDel As Long, mIDTipoCuota As Integer
        mIDArticuloToDel = Val(lArticulo.Cell(flexcpData, lArticulo.Row, 1))
        mIDTipoCuota = Val(lArticulo.Cell(flexcpData, lArticulo.Row, 4))
         
        If mIDArticuloToDel = 0 Then Exit Sub
        
        If MsgBox("Quitar el artículo “" & lArticulo.Cell(flexcpText, lArticulo.Row, 1) & "” de la solicitud." & vbCrLf & vbCrLf & _
                    "Si elimina el renglón, se va cambiar la solicitud original y no es posible volver atrás." & vbCrLf & _
                    "¿Confirma eliminar el renglón de la solicitud?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Renglón") = vbNo Then Exit Sub
        
        Screen.MousePointer = 11
        
        Cons = "Select * from RenglonSolicitud " & _
            " Where RSoSolicitud = " & gIdSolicitud & _
            " And RSoTipoCuota = " & mIDTipoCuota & _
            " And RSoArticulo = " & mIDArticuloToDel
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then RsAux.Delete
        RsAux.Close
                    
        CargoFinanciacion gIdSolicitud
        
        Screen.MousePointer = 0
    End If
    Exit Sub

errDelete:
    clsGeneral.OcurrioError "Error al eliminar el artículo de la solicitud.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub lCiRuc_DblClick()
    If gCliente <> 0 Then EjecutarApp pathApp & "Visualizacion de operaciones.exe", CStr(gCliente)
End Sub

Private Sub lComentarioG_DblClick()
    
    On Error Resume Next
    mValor = Val(Picture1(3).Tag)
    If mValor <> 0 Then
        
        If lComentarioG.Col <> lComentarioG.Cols - 1 Then
            AccionMenuComentario mValor
        Else
            Dim aIdDoc As Long
            aIdDoc = Val(lComentarioG.Cell(flexcpData, lComentarioG.Row, lComentarioG.Cols - 1))
            If aIdDoc <> 0 Then
                EjecutarApp pathApp & "Detalle de Factura.exe", CStr(aIdDoc)
            Else
                AccionMenuComentario mValor
            End If
        End If
        
    End If
    
End Sub

Private Sub lComentarioG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        mValor = Val(Picture1(3).Tag)
        If mValor = 0 Then Exit Sub
        
        If mValor = gGarantia Then
            TipoMnu = MnuCG.Garantia: MnuTConyuge.Caption = "Menú Garantía"
            PopupMenu MnuConyuge, , , , MnuTConyuge
        End If
        
        If mValor = gConyuge Then
            TipoMnu = MnuCG.Conyuge: MnuTConyuge.Caption = "Menú Conyuge"
            PopupMenu MnuConyuge, , , , MnuTConyuge
        End If
    End If

End Sub

Private Sub lComentarioT_DblClick()
    
    If gCliente <> 0 Then
        On Error Resume Next
        
        If lComentarioT.Col <> lComentarioT.Cols - 1 Then
            AccionMenuComentario gCliente
        Else
            Dim aIdDoc As Long
            aIdDoc = Val(lComentarioT.Cell(flexcpData, lComentarioT.Row, lComentarioT.Cols - 1))
            If aIdDoc <> 0 Then
                EjecutarApp pathApp & "Detalle de Factura.exe", CStr(aIdDoc)
            Else
                AccionMenuComentario gCliente
            End If
        End If
    End If
    
End Sub

Private Sub lComentarioT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And gCliente <> 0 Then MenuBDClienteTitular
End Sub

Private Sub lCondicion_DblClick()
    
    'Resuelvo con la condición de esta fila.
    On Error GoTo errCD
    If lCondicion.Row <= 0 Then Exit Sub
    
    If Val(lCondicion.Cell(flexcpData, lCondicion.Row, 1)) > 0 Then
        
        If MsgBox("¿Confirma aprobar la solicitud con la condición " & lCondicion.Cell(flexcpText, lCondicion.Row, 3) & "?", vbQuestion + vbYesNo, "Aprobar Solicitud") = vbYes Then
            
            AccionGrabarPorPrecondicion Val(lCondicion.Cell(flexcpData, lCondicion.Row, 2))
        
        End If
    
    End If
    
    Exit Sub
errCD:
End Sub

Private Sub lCSolicitud_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If lCSolicitud.Rows = 1 Then Exit Sub
        If Not IsNumeric(lCSolicitud.Cell(flexcpText, lCSolicitud.Row, 0)) Then Exit Sub
        BuscoEmpresaAfiliada lCSolicitud.Cell(flexcpText, lCSolicitud.Row, 0)
    End If
    
End Sub

Private Sub BuscoEmpresaAfiliada(Afiliado As String)
        
    On Error GoTo errAfiliado
    'Busco los datos de la empresa por numero de afiliado
    Cons = "Select * from CEmpresa Where CEmAfiliado = " & Afiliado
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Cons = "Empresa: " & Trim(RsAux!CEmFantasia)
        If Not IsNull(RsAux!CEmNombre) Then Cons = Cons & " (" & Trim(RsAux!CEmNombre) & ")"
        MsgBox Cons, vbInformation, "Detalle del Afiliado: " & Afiliado
    Else
        MsgBox "No existe registro de empresas afiliadas.", vbInformation, "Detalle del Afiliado: " & Afiliado
    End If
    RsAux.Close
    Exit Sub

errAfiliado:
    clsGeneral.OcurrioError "Ocurrió un error al buscar la empresa afiliada.", Err.Description
End Sub

Private Sub lDireccion_DblClick()
    ConfirmarDireccion lDireccion, cDireccion.ItemData(cDireccion.ListIndex)
End Sub

Private Sub lDireccionG_DblClick()
    ConfirmarDireccion lDireccionG, cDireccionG.ItemData(cDireccionG.ListIndex)
End Sub

Private Sub lEmpleoG_DblClick()
    mValor = Val(Picture1(3).Tag)
    If mValor <> 0 Then AccionMenuEmpleo mValor
End Sub

Private Sub lEmpleoG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        mValor = Val(Picture1(3).Tag)
        If mValor = 0 Then Exit Sub
        
        If mValor = gGarantia Then
            TipoMnu = MnuCG.Garantia
            MnuTConyuge.Caption = "Menú Garantia"
            PopupMenu MnuConyuge, , , , MnuTConyuge
        End If
    
        If mValor = gConyuge Then
            TipoMnu = MnuCG.Conyuge: MnuTConyuge.Caption = "Menú Cónyuge"
            PopupMenu MnuConyuge, , , , MnuTConyuge
        End If
    End If
    
End Sub

Private Sub lEmpleoT_DblClick()
    If gCliente <> 0 Then AccionMenuEmpleo gCliente
End Sub

Private Sub lEmpleoT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And gCliente <> 0 Then MenuBDClienteTitular
End Sub

Private Sub MenuBDClienteTitular()

    If gTipoCliente = TipoCliente.Empresa Then MnuTEmpleo.Enabled = False Else MnuTEmpleo.Enabled = True
    PopupMenu MnuTitular, , , , MnuTTitular
    
End Sub

Private Sub lGarantia_DblClick()
    
    mValor = Val(Picture1(3).Tag)
    If mValor = 0 Then Exit Sub
    
    AccionMenuFicha mValor, TipoCliente.Cliente
    
End Sub

Private Sub lOperacion_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 8 Or Not oCredito.value Then Cancel = True
End Sub

Private Sub lOperacion_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    EjecutarApp pathApp & "Visualizacion de operaciones", CStr(lOperacion.Cell(flexcpData, Row, Col))
End Sub

Private Sub lOperacion_DblClick()
    On Error GoTo errDC
    If oContado.value And oContado.Tag <> 0 Then EjecutarDetalleFactura
    If oCredito.value And oCredito.Tag <> 0 Then EjecutarDetalleOperaciones bIrANota:=(lOperacion.Cell(flexcpData, lOperacion.Row, 14) = 1)
    
    If oArticulo.value Then
        If lOperacion.Rows > 1 Then
            'data 0=idDocumento, 1 TipoDocumento
            Select Case lOperacion.Cell(flexcpData, lOperacion.Row, 1)
                Case TipoDocumento.Contado: EjecutarDetalleFactura
                Case TipoDocumento.Credito: EjecutarDetalleOperaciones
            End Select
        End If
    End If
    
    
    If oSolicitud.value And Val(oSolicitud.Tag) <> 0 Then
        EjecutarApp pathApp & "Visualizacion de Solicitudes", CStr(lOperacion.Cell(flexcpValue, lOperacion.Row, 0))
    End If

errDC:
End Sub

Private Sub lOperacion_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo errOp
    Select Case KeyCode
        Case vbKeyAdd, vbKeyPageDown      'Tecla + para avanzar de página-----------------------------------------
            
            If KeyCode = vbKeyPageDown Then
                If lOperacion.Rows - 1 <> lOperacion.Row Then
                    lOperacion.Select lOperacion.Rows - 1, 0
                    lOperacion.TopRow = lOperacion.Rows - 1
                    Exit Sub
                End If
                KeyCode = 0
            End If
        
            Dim auxIDCliente As Long
            Select Case LCase(TabPersona.SelectedItem.Key)
                Case "titular": auxIDCliente = gCliente
                Case "garantia": auxIDCliente = gGarantia
                Case "conyuge": auxIDCliente = gConyuge
            End Select
            
            If oContado.value And oContado.Tag <> 0 Then CargoOpContado auxIDCliente, bNext:=True
            If oCredito.value And oCredito.Tag <> 0 Then CargoOpCredito auxIDCliente, bNext:=True
            If oSolicitud.value And oSolicitud.Tag <> 0 Then CargoOpSolicitud auxIDCliente, bNext:=True
            If oArticulo.value Then CargoOpArticulo auxIDCliente, bNext:=True
            If oVtaTel.value Then CargoOpVtaTelefonica auxIDCliente, bNext:=True
            
        Case 93            'Boton Derecho------------------------------------------------------------------------------------------------
                BotonDerechoOperaciones True
                
        Case vbKeyReturn
                If oContado.value And oContado.Tag <> 0 Then EjecutarDetalleOperaciones
                If oCredito.value And oCredito.Tag <> 0 Then EjecutarDetalleOperaciones bIrANota:=(lOperacion.Cell(flexcpData, lOperacion.Row, 14) = 1)
             
    End Select
    Exit Sub
errOp:
End Sub

Private Sub lOperacion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then BotonDerechoOperaciones
End Sub

Private Sub BotonDerechoOperaciones(Optional coordBoton As Boolean = False)

    On Error GoTo errOp
    
    Dim aDocumento As Long, aX As Long, aY As Long
    
    If coordBoton Then
        aY = (lOperacion.Row * lOperacion.RowHeight(0)) - lOperacion.RowHeight(0) + Tab1.Top + Picture1(0).Top - lOperacion.Top
        aX = 1000
    End If
    
    aDocumento = lOperacion.Cell(flexcpData, lOperacion.Row, 0)
    Dim Estado As Boolean
    If aDocumento = 0 Then Estado = False Else Estado = True
    
    'CREDITOS------------------------------------------------------------------------------------------------------------------------------------
    If oCredito.value And oCredito.Tag <> 0 Then
        MnuOpCrDetalleOp.Enabled = Estado: MnuOpCrDetallePa.Enabled = Estado
        MnuOpCrDeudaCh.Enabled = Estado: MnuOpCrDetalleFa.Enabled = Estado
        If coordBoton Then PopupMenu MnuOpCredito, , aX, aY, MnuOpCrTitulo Else PopupMenu MnuOpCredito, DefaultMenu:=MnuOpCrTitulo
    End If
    
    'CONTADOS-----------------------------------------------------------------------------------------------------------------------------------
    If oContado.value And oContado.Tag <> 0 Then
        If lOperacion.Rows = 1 Then Exit Sub
        MnuOpCoDetalleFa.Enabled = Estado: MnuOpCoDetalleOp.Enabled = Estado
        If coordBoton Then PopupMenu MnuOpContado, , aX, aY, MnuOpCoTitulo Else PopupMenu MnuOpContado, DefaultMenu:=MnuOpCoTitulo
    End If
    
    'VTAS TELEFONICAS--------------------------------------------------------------------------------------------------------------------------
    If oVtaTel.value And oVtaTel.Tag <> 0 Then
        If lOperacion.Rows = 1 Then Exit Sub
        MnuOpVtaDF.Enabled = Estado: MnuOpVtaDO.Enabled = Estado
        If coordBoton Then PopupMenu MnuOpVtaT, , aX, aY, MnuOpVtaTitulo Else PopupMenu MnuOpVtaT, DefaultMenu:=MnuOpVtaTitulo
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------------------
    Exit Sub

errOp:
End Sub


Private Sub lReferenciaG_DblClick()
    mValor = Val(Picture1(3).Tag)
    If mValor <> 0 Then AccionMenuReferencia mValor
End Sub

Private Sub lReferenciaG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        mValor = Val(Picture1(3).Tag)
        If mValor = 0 Then Exit Sub
        
        If mValor = gGarantia Then
            TipoMnu = MnuCG.Garantia: MnuTConyuge.Caption = "Menú Garantía"
            PopupMenu MnuConyuge, , , , MnuTConyuge
        End If
    
        If mValor = gConyuge Then
            TipoMnu = MnuCG.Conyuge: MnuTConyuge.Caption = "Menú Cónyuge"
            PopupMenu MnuConyuge, , , , MnuTConyuge
        End If
    End If
    
End Sub

Private Sub lReferenciaT_DblClick()
    If gCliente <> 0 Then AccionMenuReferencia gCliente
End Sub

Private Sub lReferenciaT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton And gCliente <> 0 Then MenuBDClienteTitular

End Sub

Private Sub lTelefono_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Val(lTelefono.Tag) = 0 Then Exit Sub
    If Button = vbRightButton Then AccionMenuLlamoDe gCliente
End Sub

Private Sub AccionMenuLlamoDe(idCliente As Long)
    On Error GoTo errMnuLlamo
    Dim I As Integer
    MnuTX(0).Visible = True
    MnuTX(0).Tag = idCliente
    For I = 1 To MnuTX.UBound
        Unload MnuTX(I)
    Next
    
    Cons = "Select * from TipoTelefono " & _
              " Where TTeCodigo Not In (Select TelTipo from Telefono Where TelCliente = " & idCliente & ")" & _
              " Order by TTeNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        I = MnuTX.UBound + 1
        Load MnuTX(I)
        With MnuTX(I)
            .Caption = Trim(RsAux!TTeNombre)
            .Tag = RsAux!TTeCodigo
            .Visible = True
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If MnuTX.UBound > 0 Then
        MnuTX(0).Visible = False
        PopupMenu MnuTelefono, DefaultMenu:=MnuTLlamoDe
    End If

errMnuLlamo:
End Sub

Private Sub lTitular_DblClick()
    If gCliente = 0 Then Exit Sub
    AccionMenuFicha gCliente, gTipoCliente
End Sub

Private Sub lTitulo_DblClick()
    If gCliente = 0 Then Exit Sub
    If Trim(lTitulo.Caption) <> "" Then AccionMenuTitulo gCliente
End Sub

Private Sub lTituloG_Click()
    mValor = Val(Picture1(3).Tag)
    If mValor = 0 Then Exit Sub
    If Trim(lTituloG.Caption) <> "" Then AccionMenuTitulo mValor
End Sub

Private Sub MnuCComentario_Click()
    Select Case TipoMnu
        Case MnuCG.Conyuge: AccionMenuComentario gConyuge
        Case MnuCG.Garantia: AccionMenuComentario gGarantia
    End Select
End Sub

Private Sub MnuCEmpleo_Click()
    Select Case TipoMnu
        Case MnuCG.Conyuge: AccionMenuEmpleo gConyuge
        Case MnuCG.Garantia: AccionMenuEmpleo gGarantia
    End Select
End Sub

Private Sub MnuCFicha_Click()
    
    Select Case TipoMnu
        Case MnuCG.Conyuge
                If gConyuge = 0 Then Exit Sub
                AccionMenuFicha gConyuge, TipoCliente.Cliente
        Case MnuCG.Garantia
                If gGarantia = 0 Then Exit Sub
                AccionMenuFicha gGarantia, TipoCliente.Cliente
    End Select
                                            
End Sub

Private Sub mnuCondicionX_Click(Index As Integer)
    
    If gCliente = 0 Then Exit Sub
    If Val(mnuCondicionX(Index).Tag) = 0 Then Exit Sub
    
    tComentario.Text = "": tComentario.Tag = ""
    lRComentario.Tag = ""
    
    BuscoCodigoEnCombo cCondicionR, Val(mnuCondicionX(Index).Tag)

    ZDeshabilitoCondiciones
    mnuMCondicion.Tag = Val(mnuCondicionX(Index).Tag)
    ZHabilitoCondiciones Val(mnuCondicionX(Index).Tag)

End Sub

Private Sub MnuCReferencia_Click()
    Select Case TipoMnu
        Case MnuCG.Conyuge: AccionMenuReferencia gConyuge
        Case MnuCG.Garantia: AccionMenuReferencia gGarantia
    End Select
End Sub

Private Sub MnuCRefrescar_Click()
    On Error Resume Next
    Screen.MousePointer = 11
    
    Select Case TipoMnu
        Case MnuCG.Conyuge
                If gConyuge <> 0 Then
                    CargoEmpleos gConyuge, Conyuge:=True
                    CargoReferencias gConyuge, Conyuge:=True
                    CargoComentarios gConyuge, Conyuge:=True
                End If
        Case MnuCG.Garantia
                If gGarantia <> 0 Then
                    CargoEmpleos gGarantia, Garantia:=True
                    CargoReferencias gGarantia, Garantia:=True
                    CargoComentarios gGarantia, Garantia:=True
                End If
    End Select
    
    Screen.MousePointer = 0
End Sub

Private Sub MnuCTitulo_Click()
    Select Case TipoMnu
        Case MnuCG.Conyuge: AccionMenuTitulo gConyuge
        Case MnuCG.Garantia: AccionMenuTitulo gGarantia
    End Select
End Sub

Private Sub MnuOpCoDetalleFa_Click()
    EjecutarDetalleFactura
End Sub

Private Sub MnuOpCoDetalleOp_Click()
    EjecutarDetalleOperaciones
End Sub

Private Sub MnuOpCrCancelada_Click()

    bFCanceladas = True
    bFVigentes = False
    MnuOpCrVigente.Checked = False
    MnuOpCrCancelada.Checked = True
    lOperacion.Tag = ""
    
    Select Case LCase(TabPersona.SelectedItem.Key)
        Case "titular": CargoOpCredito gCliente
        Case "garantia": CargoOpCredito gGarantia
        Case "conyuge": CargoOpCredito gConyuge
    End Select

End Sub

Private Sub MnuOpCrDetalleFa_Click()
    EjecutarDetalleFactura
End Sub

Private Sub MnuOpCrDetalleOp_Click()
    EjecutarDetalleOperaciones
End Sub

Private Sub MnuOpCrDetallePa_Click()
    EjecutarDetallePagos
End Sub

Private Sub EjecutarDetallePagos()
    
    If lOperacion.Rows = 0 Then Exit Sub
    
    Dim aDocumento As Long: aDocumento = 0
    aDocumento = lOperacion.Cell(flexcpData, lOperacion.Row, 0)
    If aDocumento = 0 Then Exit Sub
    
    EjecutarApp pathApp & "Detalle de pagos", CStr(aDocumento)
    
End Sub

Private Sub EjecutarDetalleOperaciones(Optional Col As Integer = 0, Optional bIrANota As Boolean = False)
    
    If lOperacion.Rows = 0 Then Exit Sub
    
    Dim aDocumento As Long: aDocumento = 0
    aDocumento = lOperacion.Cell(flexcpData, lOperacion.Row, Col)
    If aDocumento = 0 Then Exit Sub
    
    'Si el icono es una Nota y estoy visualizando opcreditos voy al detalle de factura de la NOTA !!
    If bIrANota Then
        Cons = "Select Top 1 NotNota from Nota, Documento " & _
                    " Where NotFactura = " & aDocumento & _
                    " And NotNota = DocCodigo And DocAnulado = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux!NotNota) Then aDocumento = RsAux!NotNota
        End If
        RsAux.Close
        EjecutarApp pathApp & "Detalle de Factura.exe", CStr(aDocumento)
        
    Else
        EjecutarApp pathApp & "Detalle de operaciones", CStr(aDocumento)
    End If
    
End Sub

Private Sub EjecutarDetalleFactura(Optional Col As Integer = 0)
    
    If lOperacion.Rows = 0 Then Exit Sub
    
    Dim aDocumento As Long: aDocumento = 0
    aDocumento = lOperacion.Cell(flexcpData, lOperacion.Row, Col)
    If aDocumento = 0 Then Exit Sub
    
    EjecutarApp pathApp & "Detalle de factura", CStr(aDocumento)
    
End Sub

Private Sub MnuOpCrDeudaCh_Click()
    
    If lOperacion.Rows = 0 Then Exit Sub
    
    If oCredito.value Then
        Dim auxIDCliente As Long
        
        Select Case LCase(TabPersona.SelectedItem.Key)
            Case "titular": auxIDCliente = gCliente
            Case "garantia": auxIDCliente = gGarantia
            Case "conyuge": auxIDCliente = gConyuge
        End Select
        EjecutarApp pathApp & "Deuda en cheques", CStr(auxIDCliente)
    End If

End Sub

Private Sub MnuOpCrEliminar_Click()

    bFVigentes = False: bFCanceladas = False: bFTitular = False: bFGarante = False
    lOperacion.Tag = ""
    
    MnuOpCrVigente.Checked = False: MnuOpCrCancelada.Checked = False: MnuOpCrGarantia.Checked = False: MnuOpCrTitular.Checked = False
    
    Dim auxIDCliente As Long
    Select Case LCase(TabPersona.SelectedItem.Key)
        Case "titular": auxIDCliente = gCliente
        Case "garantia": auxIDCliente = gGarantia
        Case "conyuge": auxIDCliente = gConyuge
    End Select
        
    CargoOpCredito auxIDCliente
    
End Sub

Private Sub MnuOpCrGarantia_Click()
    
    bFTitular = False: bFGarante = True
    MnuOpCrTitular.Checked = False: MnuOpCrGarantia.Checked = True

    lOperacion.Tag = ""
    Dim auxIDCliente As Long
    Select Case LCase(TabPersona.SelectedItem.Key)
        Case "titular": auxIDCliente = gCliente
        Case "garantia": auxIDCliente = gGarantia
        Case "conyuge": auxIDCliente = gConyuge
    End Select
    
    CargoOpCredito auxIDCliente
    
End Sub

Private Sub MnuOpCrTitular_Click()

    bFTitular = True: bFGarante = False
    MnuOpCrTitular.Checked = True: MnuOpCrGarantia.Checked = False
    
    lOperacion.Tag = ""
    
    Dim auxIDCliente As Long
    Select Case LCase(TabPersona.SelectedItem.Key)
        Case "titular": auxIDCliente = gCliente
        Case "garantia": auxIDCliente = gGarantia
        Case "conyuge": auxIDCliente = gConyuge
    End Select
    CargoOpCredito auxIDCliente
    
End Sub

Private Sub MnuOpCrVigente_Click()

    bFVigentes = True: bFCanceladas = False
    MnuOpCrVigente.Checked = True: MnuOpCrCancelada.Checked = False
    
    lOperacion.Tag = ""
    
    Dim auxIDCliente As Long
    Select Case LCase(TabPersona.SelectedItem.Key)
        Case "titular": auxIDCliente = gCliente
        Case "garantia": auxIDCliente = gGarantia
        Case "conyuge": auxIDCliente = gConyuge
    End Select
    CargoOpCredito auxIDCliente
    
End Sub

Private Sub MnuOpDiConfirmar_Click()
    
    Select Case aAccionDireccion
        Case 1: ConfirmarDireccion lDireccion, cDireccion.ItemData(cDireccion.ListIndex)
        Case 3: ConfirmarDireccion lDireccionG, cDireccionG.ItemData(cDireccionG.ListIndex)
    End Select
    
End Sub

Private Sub MnuOpDiModificar_Click()
    
    On Error GoTo errDir
    Screen.MousePointer = 11
    Dim aCodDir As Long, aDirEdicion As Long
    Dim bEsPrincipal As Boolean, aIdCoG As Long
    
    Select Case aAccionDireccion
        Case 1: aDirEdicion = cDireccion.ItemData(cDireccion.ListIndex)
                    If aDirEdicion = Val(cDireccion.Tag) Then bEsPrincipal = True Else bEsPrincipal = False
                    
        Case 3: aDirEdicion = cDireccionG.ItemData(cDireccionG.ListIndex)
                    If aDirEdicion = Val(cDireccionG.Tag) Then bEsPrincipal = True Else bEsPrincipal = False
                    
                    aIdCoG = Val(Picture1(3).Tag)
                    If aIdCoG = 0 Then Exit Sub
    End Select
    
    Dim objDir As New clsDireccion          '----------------------------------------------------------------------------------------------------!!!!
    Select Case aAccionDireccion
        Case 1:
                    If bEsPrincipal Then
                        objDir.ActivoFormularioDireccion cBase, aDirEdicion, gCliente, "Cliente", "CliDireccion", "CliCodigo", gCliente
                    Else
                        objDir.ActivoFormularioDireccion cBase, aDirEdicion, gCliente
                    End If
        
        Case 3:
                    If bEsPrincipal Then
                        objDir.ActivoFormularioDireccion cBase, aDirEdicion, aIdCoG, "Cliente", "CliDireccion", "CliCodigo", aIdCoG
                    Else
                        objDir.ActivoFormularioDireccion cBase, aDirEdicion, aIdCoG
                    End If
    End Select
    Me.Refresh
    aCodDir = objDir.CodigoDeDireccion
    Set objDir = Nothing                        '----------------------------------------------------------------------------------------------------!!!!
    
    Select Case aAccionDireccion
        Case 1:
                    lDireccion.Caption = ""
                    If aCodDir = 0 Then         'Direccion Eliminada
                        cDireccion.RemoveItem cDireccion.ListIndex
                        If bEsPrincipal Then
                            cDireccion.Tag = 0
                        Else
                            Cons = "Delete DireccionAuxiliar Where DAuCliente = " & gCliente & " And DAuDireccion =" & aDirEdicion
                            cBase.Execute Cons
                        End If
                        
                        If cDireccion.ListCount > 0 Then cDireccion.Text = cDireccion.List(0)
                    Else
                        lDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, aCodDir, Departamento:=True, Localidad:=True, Zona:=True, ConfyVD:=True)
                        If InStr(lDireccion.Caption, "(Cf.)") <> 0 Then lDireccion.ForeColor = Colores.osVerde Else lDireccion.ForeColor = Colores.Azul
                    End If
                        
        Case 3:
                    lDireccionG.Caption = ""
                    If aCodDir = 0 Then         'Direccion Eliminada
                        cDireccionG.RemoveItem cDireccionG.ListIndex
                        If bEsPrincipal Then
                            cDireccionG.Tag = 0
                        Else
                            Cons = "Delete DireccionAuxiliar Where DAuCliente = " & aIdCoG & " And DAuDireccion = " & aDirEdicion
                            cBase.Execute Cons
                        End If
                        
                        If cDireccionG.ListCount > 0 Then cDireccionG.Text = cDireccionG.List(0) Else cDireccionG.BackColor = Colores.Obligatorio
                    Else
                        lDireccionG.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, aCodDir, Departamento:=True, Localidad:=True, Zona:=True, ConfyVD:=True)
                        If InStr(lDireccionG.Caption, "(Cf.)") <> 0 Then lDireccionG.ForeColor = Colores.osVerde Else lDireccionG.ForeColor = Colores.Azul
                    End If
                    
    End Select
    Screen.MousePointer = 0
    
    Exit Sub
errDir:
    clsGeneral.OcurrioError "Ocurrió un error al editar la dirección", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub MnuOpVtaDF_Click()
    EjecutarDetalleFactura Col:=2
End Sub

Private Sub MnuOpVtaDO_Click()
    EjecutarDetalleOperaciones Col:=2
End Sub

Private Sub MnuPlX_Click(Index As Integer)
    
    If Cliente <> 0 Then EjecutarApp App.Path & "\appExploreMsg.exe ", Val(MnuPlX(Index).Tag) & ":" & (gCliente)
        
End Sub

Private Sub MnuTComentario_Click()
    AccionMenuComentario gCliente
End Sub

Private Sub MnuTEmpleo_Click()
    AccionMenuEmpleo gCliente
End Sub

Private Sub MnuTFicha_Click()
    If gCliente = 0 Then Exit Sub
    AccionMenuFicha gCliente, gTipoCliente
End Sub

Private Sub MnuTReferencia_Click()
    AccionMenuReferencia gCliente
End Sub

Private Sub MnuTRefrescar_Click()
    
    On Error Resume Next
    Screen.MousePointer = 11
    If gTipoCliente = TipoCliente.Cliente Then CargoEmpleos gCliente, Titular:=True
    CargoReferencias gCliente, Titular:=True
    CargoComentarios gCliente, Titular:=True
    Screen.MousePointer = 0
    
End Sub

Private Sub MnuTTitulo_Click()
    AccionMenuTitulo gCliente
End Sub

Private Sub MnuTX_Click(Index As Integer)

    If MnuTX(0).Tag = 0 Then Exit Sub
    On Error GoTo errActualizar
    Screen.MousePointer = 11
    'Actualizo el Tipo de Telefono del Llamo de
    Cons = "Update Telefono Set TelTipo = " & MnuTX(Index).Tag & _
               " Where TelCliente = " & Val(MnuTX(0).Tag) & _
               " And TelTipo = " & paTipoTelefonoLlamoDe
    cBase.Execute Cons
    
    lTelefono.Caption = TelefonoATexto(CLng(MnuTX(0).Tag))
    lTelefono.Tag = 0
    Screen.MousePointer = 0
    Exit Sub

errActualizar:
    clsGeneral.OcurrioError "Error al actualizar el tipo de teléfono.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub oArticulo_Click()

    EncabezadoOperaciones False, False, False, True
    lOperacion.Tag = ""
    
    Dim auxIDCliente As Long
    Select Case LCase(TabPersona.SelectedItem.Key)
        Case "titular": auxIDCliente = gCliente
        Case "garantia": auxIDCliente = gGarantia
        Case "conyuge": auxIDCliente = gConyuge
    End Select
    CargoOpArticulo auxIDCliente
    
End Sub

Private Sub oArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And lOperacion.Rows > lOperacion.FixedRows Then lOperacion.SetFocus
End Sub

Private Sub objClearing_RespuestaPedido(ByVal idEvento As Integer, ByVal strMsg As String, ByVal lngCliente As Long)
    
    On Error Resume Next
    If lngCliente = gCliente Or lngCliente = gConyuge Or lngCliente = gGarantia Then
        tmClearing.Tag = idEvento & ":" & lngCliente & ":" & Trim(strMsg)
        tmClearing.Enabled = True
    End If
    
End Sub

Private Sub oContado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And lOperacion.Rows > lOperacion.FixedRows Then lOperacion.SetFocus
End Sub

Private Sub oCredito_Click()
   
    EncabezadoOperaciones True, False, False, False
    lOperacion.Tag = ""
     
    Dim auxIDCliente As Long
    Select Case LCase(TabPersona.SelectedItem.Key)
        Case "titular": auxIDCliente = gCliente
        Case "garantia": auxIDCliente = gGarantia
        Case "conyuge": auxIDCliente = gConyuge
    End Select
    CargoOpCredito auxIDCliente
    CargoDeudaCliente auxIDCliente
    
End Sub

Private Sub oContado_Click()

    EncabezadoOperaciones False, True, False, False
        
    lOperacion.Tag = ""

    Dim auxIDCliente As Long
    Select Case LCase(TabPersona.SelectedItem.Key)
        Case "titular": auxIDCliente = gCliente
        Case "garantia": auxIDCliente = gGarantia
        Case "conyuge": auxIDCliente = gConyuge
    End Select
    CargoOpContado auxIDCliente
    
End Sub

Private Sub oCredito_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And lOperacion.Rows > lOperacion.FixedRows Then lOperacion.SetFocus
End Sub


Private Sub oSolicitud_Click()

    EncabezadoOperaciones False, False, True, False
    
    lOperacion.Tag = ""
        
    Dim auxIDCliente As Long
    Select Case LCase(TabPersona.SelectedItem.Key)
        Case "titular": auxIDCliente = gCliente
        Case "garantia": auxIDCliente = gGarantia
        Case "conyuge": auxIDCliente = gConyuge
    End Select
    
    CargoOpSolicitud auxIDCliente
    
End Sub

Private Sub oSolicitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And lOperacion.Rows > lOperacion.FixedRows Then lOperacion.SetFocus
End Sub

Private Sub oVtaTel_Click()
    EncabezadoOperaciones VtaTel:=True
    lOperacion.Tag = ""
    
    Dim auxIDCliente As Long
    Select Case LCase(TabPersona.SelectedItem.Key)
        Case "titular": auxIDCliente = gCliente
        Case "garantia": auxIDCliente = gGarantia
        Case "conyuge": auxIDCliente = gConyuge
    End Select
    CargoOpVtaTelefonica auxIDCliente
    
End Sub

Private Sub Tab1_Click()

Dim miIdx As Integer
    
    miIdx = Tab1.SelectedItem.Index - 1
    If miIdx = 4 Then miIdx = 3
    If Val(Tab1.Tag) = 0 Then Tab1.Tag = 1
    
    Picture1(miIdx).ZOrder 0
    
    Me.Refresh
    
    If gCliente = 0 Then Exit Sub
    
    Select Case Tab1.SelectedItem.Key
        Case "operaciones"
            Tab1.Tag = Tab1.SelectedItem.Index
            If Picture1(miIdx).Tag <> "" Then Exit Sub
            Screen.MousePointer = 11
            Picture1(miIdx).Tag = "OK"
            Screen.MousePointer = 0
        
        Case "solicitud": Tab1.Tag = Tab1.SelectedItem.Index
            
        Case "titular"
            Tab1.Tag = Tab1.SelectedItem.Index
            If Picture1(miIdx).Tag <> "" Then Exit Sub
            Screen.MousePointer = 11
            
            Picture1(Tab1.SelectedItem.Index - 1).Tag = "OK"
            If gTipoCliente = TipoCliente.Cliente Then CargoEmpleos gCliente, Titular:=True
            CargoReferencias gCliente, Titular:=True
            CargoComentarios gCliente, Titular:=True
            
            Screen.MousePointer = 0
        
        Case "garantia" 'Informacion de la Garantia-----------------------------------------------
            If gGarantia = 0 Then
                Tab1.Tabs(Val(Tab1.Tag)).Selected = True
            Else
                Tab1.Tag = Tab1.SelectedItem.Index
                If Val(Picture1(miIdx).Tag) = gGarantia Then Exit Sub
                LimpioFichaConGar
                Screen.MousePointer = 11
                Picture1(miIdx).Tag = gGarantia
                CargoDatosGarantia gGarantia
                Screen.MousePointer = 0
            End If
            
        Case "conyuge" 'Informacion del Conyuge-------------------------------------------------
            If gConyuge = 0 Then
                Tab1.Tabs(Val(Tab1.Tag)).Selected = True
            Else
                Tab1.Tag = Tab1.SelectedItem.Index
                If Val(Picture1(miIdx).Tag) = gConyuge Then Exit Sub
                Screen.MousePointer = 11
                LimpioFichaConGar
                Picture1(Tab1.SelectedItem.Index - 2).Tag = gConyuge
                CargoDatosGarantia gConyuge
                Screen.MousePointer = 0
            End If
            
        Case "clearing" 'Informacion del Clearing-------------------------------------------------
            Tab1.Tag = Tab1.SelectedItem.Index
            If Picture1(miIdx).Tag <> "" Then Exit Sub
            Screen.MousePointer = 11
            Picture1(Tab1.SelectedItem.Index - 1).Tag = "OK"
            CargoDatosClienteClearing gCliente
            Screen.MousePointer = 0
            
        Case "resolver" 'Informacion del Resolucion-----------------------------------------------
            Tab1.Tag = Tab1.SelectedItem.Index
            If Picture1(miIdx).Tag <> "" Then Exit Sub
            Picture1(Tab1.SelectedItem.Index - 1).Tag = "OK"
            If cRMoneda.ListCount <> 0 Then Exit Sub
            Screen.MousePointer = 11
            Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"   'Cargo las Monedas
            CargoCombo Cons, cRMoneda, ""
            'Cargo los Planes (Tipos de Cuotas)
            Cons = "Select TCuCodigo, TCuAbreviacion From TipoCuota" _
                    & " Where TCuCodigo <> " & paTipoCuotaContado _
                    & " Order by TCuAbreviacion"
            CargoCombo Cons, cPCuota, ""
            Screen.MousePointer = 0
            
        Case "relaciones" 'Relaciones de Parentezco-------------------------------------------------
            Tab1.Tag = Tab1.SelectedItem.Index
            If Picture1(miIdx).Tag <> "" Then Exit Sub
            Screen.MousePointer = 11
            Picture1(Tab1.SelectedItem.Index - 1).Tag = "OK"
            CargoDatosClienteRelaciones gCliente
            Screen.MousePointer = 0
    End Select
    
    On Error Resume Next
    Select Case LCase(Tab1.SelectedItem.Key)
        Case "operaciones": lOperacion.SetFocus
        Case "titular": lEmpleoT.SetFocus
        Case "garantia": lEmpleoG.SetFocus
        Case "clearing": lCSolicitud.SetFocus
        Case "resolver"
    End Select

End Sub

Private Sub CargoDatosGarantia(Codigo As Long)

    On Error GoTo errConyuge
    
    Dim pintTipoCliente As Integer
    
    cDireccionG.Clear: cDireccionG.BackColor = Colores.Obligatorio: cDireccionG.Tag = 0
        
    'Cargo Datos Tabla Cliente----------------------------------------------------------------------
'    Cons = "Select * from Cliente, CategoriaCliente" _
'           & " Where CliCodigo = " & Codigo _
'           & " And CliCategoria *= CClCodigo"

    Cons = "Select * FROM Cliente LEFT OUTER JOIN CategoriaCliente ON CliCategoria = CClCodigo" _
           & " Where CliCodigo = " & Codigo
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    pintTipoCliente = RsAux!CliTipo
    If Not IsNull(RsAux!CliCiRuc) Then
        PresentoPaisDelDocumentoCliente RsAux("CliPaisDelDocumento"), RsAux!CliCiRuc, lblInfoCIConyuge, lCIG
        'lCIG.Caption = clsGeneral.RetornoFormatoCedula(RsAux!CliCiRuc)
    Else
        PresentoPaisDelDocumentoCliente RsAux("CliPaisDelDocumento"), "", lblInfoCIConyuge, lCIG
    End If
    
    If Not IsNull(RsAux!CliDireccion) Then
        cDireccionG.AddItem "Dirección Principal": cDireccionG.ItemData(cDireccionG.NewIndex) = RsAux!CliDireccion
        cDireccionG.Tag = RsAux!CliDireccion
    End If
    
    If Not IsNull(RsAux!CliAlta) Then
        If Format(RsAux!CliAlta, "dd/mm/yyyy") = Format(gFechaServidor, "dd/mm/yyyy") Then
            shpGarantia.BackColor = Colores.clVerde
        Else
            shpGarantia.BackColor = Colores.Obligatorio
        End If
    Else
        shpGarantia.BackColor = Colores.Obligatorio
    End If
    shpGarantia.Refresh
    
    If Not IsNull(RsAux!CliCategoria) Then
        If RsAux!CliCategoria = paCatCliFallecido Then shpGarantia.BackColor = Colores.Gris
        If RsAux!CliCategoria <> paCategoriaCliente Then lCategoriaG.ForeColor = vbRed
    End If
    
    If Not IsNull(RsAux!CClNombre) Then lCategoriaG.Caption = Trim(RsAux!CClNombre)
    If Not IsNull(RsAux!CliSolicitud) Then lNSolicitudG.Caption = Trim(RsAux!CliSolicitud)
    If Not IsNull(RsAux!CliCheque) Then If UCase(RsAux!CliCheque) = "S" Then lChequesG.Caption = "Opera c/cheques"
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------
    If Val(cDireccionG.Tag) <> 0 Then BuscoCodigoEnCombo cDireccionG, Val(cDireccionG.Tag)
    CargoDireccionesAuxiliares cDireccionG, Codigo
    
    'Cargo Datos Tabla CPersona o CEmpresa------------------------------------------------------
    If pintTipoCliente = 1 Then
        
'        Cons = "Select * from CPersona, Ocupacion, EstadoCivil, InquilinoPropietario " _
'               & " Where CPeCliente = " & Codigo _
'               & " And CPeOcupacion *= OcuCodigo" _
'               & " And CPeEstadoCivil *= ECiCodigo" _
'               & " And CPePropietario *= IPrCodigo"
               
        Cons = "SELECT * FROM CPersona " _
            & "LEFT OUTER JOIN Ocupacion ON CPeOcupacion = OcuCodigo " _
            & "LEFT OUTER JOIN EstadoCivil ON CPeEstadoCivil = ECiCodigo " _
            & "LEFT OUTER JOIN InquilinoPropietario  ON CPePropietario = IPrCodigo " _
            & "WHERE CPeCliente = " & Codigo
               
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            lGarantia.Caption = ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
            If Not IsNull(RsAux!CPeFNacimiento) Then
                lEdadG.Caption = ((Date - RsAux!CPeFNacimiento) \ 365) & Format(RsAux!CPeFNacimiento, " (d-Mmm-yyyy)")
                If ((Date - RsAux!CPeFNacimiento) \ 365) > paMayorEdad1 Then lEdadG.ForeColor = vbRed
            End If
            If Not IsNull(RsAux!OcuNombre) Then lOcupacionG.Caption = Trim(RsAux!OcuNombre)
            If Not IsNull(RsAux!ECiNombre) Then lECivilG.Caption = Trim(RsAux!ECiNombre)
            
            If Not IsNull(RsAux!IPrNombre) And Val(cDireccionG.Tag) <> 0 Then
                If cDireccionG.ListCount > 0 Then cDireccionG.List(0) = Trim(RsAux!IPrNombre)
                lNDireccionG.Caption = Trim(RsAux!IPrNombre) & ":"
            End If
        End If
        RsAux.Close
    Else
'        Cons = "Select * from CEmpresa, Ramo " _
'                & " Where CEmCliente = " & Codigo _
'                & " And CEmRamo *= RamCodigo"
                
        Cons = "SELECT CEmFantasia, CEmNombre, RamNombre " & _
                "FROM CEmpresa LEFT OUTER JOIN Ramo ON CEmRamo = RamCodigo " _
               & "WHERE CEmCliente = " & Codigo
                
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            aTexto = Trim(RsAux!CEmFantasia)
            If Not IsNull(RsAux!CEmNombre) Then aTexto = aTexto & " (" & Trim(RsAux!CEmNombre) & ")"
            lGarantia.Caption = aTexto
            If Not IsNull(RsAux!RamNombre) Then lOcupacion.Caption = Trim(RsAux!RamNombre)
        End If
        RsAux.Close
    End If
    
    If cDireccionG.ListCount <= 1 Then
        If Val(cDireccionG.Tag) = 0 Then lNDireccionG.Caption = Trim(cDireccionG.Text) & ":"
        lNDireccionG.Refresh
        lDireccionG.Left = lNDireccionG.Left + lNDireccionG.Width + 40
    Else
        lDireccionG.Left = cDireccionG.Left + cDireccionG.Width + 40
    End If
    
    lTituloG.Caption = CargoTitulos(Codigo)
    
    CargoEmpleos Codigo, Garantia:=True
    CargoReferencias Codigo, Garantia:=True
    CargoComentarios Codigo, Garantia:=True
    Exit Sub

errConyuge:
    clsGeneral.OcurrioError "Error al cargar los datos de la garantía.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosClienteClearing(Cliente As Long)

    'Limpio los campos de la ficha clearing-------------------------------------------
    shpClearing.BackColor = Colores.Gris
    lClFecha.Caption = ""
    lClNombre.Caption = ""
    lClConyuge.Caption = ""
    lClTelefono.Caption = ""
    lClDireccion.Caption = ""
    lClFNacimiento.Caption = ""
    lClEmpleo.Caption = ""
    lClECivil.Caption = ""
    lClDireccionE.Caption = ""
    lClAntiguedad.Caption = ""
    lClAlquiler.Caption = ""
    
    lClInfoMsg.Visible = False: lClInfoMsg.Caption = ""
    lCSolicitud.Rows = lCSolicitud.FixedRows
    lClCosto.Caption = ""
    
    TabClearing.Tabs.Clear
    TabClearing.Tabs.Add pvKey:="solicitudes", pvCaption:="Solicitudes Reali&zadas", pvImage:=ImageList1.ListImages("solicitudes").Index
        
    If Cliente = 0 Then Exit Sub
    
    Dim mTabSelected As String
    
    If Cliente = 0 Then
        Tab1.Tabs("clearing").Image = ImageList1.ListImages("clearing0").Index
        Exit Sub
    End If
    Dim mValorC As Integer: mValorC = 0
    
    Screen.MousePointer = 11
    On Error GoTo errClearing
    Dim dUltimoClearing As Date, bHayDatos As Boolean
    dUltimoClearing = CDate("01/01/1900"): bHayDatos = False
    
    'Cargo Datos Tabla Clearing----------------------------------------------------------------------
    Cons = "Select * from Clearing Left Outer Join CalificaClearing On CleCalificacion = CClCodigo Where CleCliente= " & Cliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        bHayDatos = True
        shpClearing.BackColor = &HC0E0FF
        
        dUltimoClearing = RsAux!CleFecha
        lClFecha.Caption = Format(RsAux!CleFecha, "d-Mmm yyyy hh:mm")
        
        lClNombre.Caption = Trim(RsAux!CleApellidos) & ", " & Trim(RsAux!CleNombre)
        If Not IsNull(RsAux!CleConyuge) Then lClConyuge.Caption = Trim(RsAux!CleConyuge)
        If Not IsNull(RsAux!CleTelefono) Then lClTelefono.Caption = Trim(RsAux!CleTelefono)
        If Not IsNull(RsAux!CleDireccion) Then lClDireccion.Caption = Trim(RsAux!CleDireccion)
        If Not IsNull(RsAux!CleFNacimiento) Then lClFNacimiento.Caption = Format(RsAux!CleFNacimiento, "d-Mmm yyyy")
        If Not IsNull(RsAux!CleECivil) Then lClECivil.Caption = Trim(RsAux!CleECivil)
        If Not IsNull(RsAux!CleEmpleo) Then lClEmpleo.Caption = Trim(RsAux!CleEmpleo)
        If Not IsNull(RsAux!CleDireccionEmpleo) Then lClDireccionE.Caption = Trim(RsAux!CleDireccionEmpleo)
        If Not IsNull(RsAux!CleAntiguedad) Then lClAntiguedad.Caption = Trim(RsAux!CleAntiguedad)
        
        aTexto = ""
        If Not IsNull(RsAux!ClePropInq) Then aTexto = Trim(RsAux!ClePropInq)
        If Not IsNull(RsAux!CleMonedaAlqCuota) Then aTexto = aTexto & " " & Trim(RsAux!CleMonedaAlqCuota)
        If Not IsNull(RsAux!CleMontoAlqCuota) Then aTexto = aTexto & " " & Format(RsAux!CleMontoAlqCuota, FormatoMonedaP)
        If Trim(aTexto) <> "" Then lClAlquiler.Caption = Trim(aTexto)
        
        mValorC = 1
        If Not IsNull(RsAux!CClValor) Then
            Select Case RsAux!CClValor
                Case 0, 1, 2: mValorC = 1
                Case 3: mValorC = 3
                Case 4, 5: mValorC = 4
                Case 6: mValorC = 6
            End Select
        End If
    End If
    RsAux.Close
    
    'Porcentaje = Costo Clearing / (Tot.Financ - Contado)
    'Para presentarlo en el tab del clearing: $ 0 (0.0%)      ==>  $ Costo Clearing (Porcentaje de incidencia)
    
    '(Tot.Financ - Contado) --> los saco de la lista de articulos
    Dim pcurValor As Currency
    pcurValor = lArticulo.Cell(flexcpValue, lArticulo.Rows - 1, 6) - lArticulo.Cell(flexcpValue, lArticulo.Rows - 1, 2)

    Cons = "select dbo.PrecioClearing('" & Format(dUltimoClearing, "yyyy/mm/dd") & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        lClCosto.Caption = "$ " & RsAux(0) & " " & Format(RsAux(0) / pcurValor, "(0.0%)")
        lClCosto.ForeColor = IIf((RsAux(0) / pcurValor * 100) > prmAlertaPorcCostoClearing, vbRed, vbBlue)
    End If
    RsAux.Close
    Screen.MousePointer = 0
    
        
    If Not bHayDatos Then
        TabClearing.Tabs("solicitudes").Selected = True
        Exit Sub
    End If
    
    Tab1.Tabs("clearing").Image = ImageList1.ListImages("clearing" & CStr(mValorC)).Index
    '-----------------------------------------------------------------------------------------------------
    
    '1) Cargo Datos Tabla ClearingSolicitud ---------------------------------------------------------------------------------------------------
    Cons = "Select * from ClearingSolicitud Where CSoCliente= " & Cliente _
            & " Order by CSoFecha Desc"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    I = 0
    Do While Not RsAux.EOF
        If RsAux!CSoFecha > Now - 180 Then I = I + 1    'Ultimos 6M
        RsAux.MoveNext
    Loop
    RsAux.Close
    TabClearing.Tabs("solicitudes").Caption = Trim(TabClearing.Tabs("solicitudes").Caption) & " (6m: " & I & ")"
    mTabSelected = "solicitudes"
    
    '2) Cargo Datos Tabla ClearingAntecedente -----------------------------------------------------------------------------------------------
    I = 0
    Cons = "Select Count(*) from ClearingAntecedente Where CAnCliente= " & Cliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then I = RsAux(0)
    RsAux.Close
    
    If I > 0 Then
        TabClearing.Tabs.Add pvKey:="antecedentes", pvCaption:="&Antecedentes (" & I & ")", pvImage:=ImageList1.ListImages("antecedentes").Index
        mTabSelected = "antecedentes"
    End If
    
    '3) Cargo Datos Tabla ClearingCheques   ------------------------------------------------------------------------------------------------------
    I = 0
    Cons = "Select Count(*) from ClearingCheque Where CChCliente= " & Cliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then I = RsAux(0)
    RsAux.Close
    
    If I > 0 Then
        TabClearing.Tabs.Add pvKey:="cheques", pvCaption:="C&heques Sin Fondo (" & I & ")", pvImage:=ImageList1.ListImages("cheques").Index
        lClInfoMsg.Caption = lClInfoMsg.Caption & IIf(lClInfoMsg.Caption = "", "Ch.s/Fondo", " && Ch.s/Fondo")
    End If
    
    '4) Cargo Datos Tabla ClearingComentarios------------------------------------------------------------------------------------------------------
    I = 0
    Cons = "Select Count(*) from ClearingComentario Where CCoCliente= " & Cliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then I = RsAux(0)
    RsAux.Close
    
    If I > 0 Then
        TabClearing.Tabs.Add pvKey:="comentarios", pvCaption:="Observaciones (" & I & ")", pvImage:=ImageList1.ListImages("Gestor").Index
        lClInfoMsg.Caption = lClInfoMsg.Caption & IIf(lClInfoMsg.Caption = "", "Obs.", " && Obs.")
    End If
    
    TabClearing.Tabs(mTabSelected).Selected = True
    lClInfoMsg.Visible = (Trim(lClInfoMsg.Caption) <> "")

    Screen.MousePointer = 0
    Exit Sub

errClearing:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del Clearing.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Tab1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Tab1.Tabs("solicitud").Selected Then
        If KeyCode = 93 And gCliente <> 0 Then MenuBDClienteTitular
    End If
    
End Sub

Private Sub Tab1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Tab1.Tabs("solicitud").Selected Then
        If Button = vbRightButton And gCliente <> 0 Then MenuBDClienteTitular
    End If

End Sub

Private Sub TabClearing_Click()
    
    Dim mQueCargo As Integer
    Select Case LCase(Trim(TabClearing.SelectedItem.Key))
        Case "solicitudes": mQueCargo = 1
        Case "antecedentes": mQueCargo = 2
        Case "cheques": mQueCargo = 3
        Case "comentarios": mQueCargo = 4
    End Select
    
    Select Case LCase(TabCPersona.SelectedItem.Key)
        Case "titular": CargoListaClearing mQueCargo, gCliente
        Case "garantia": CargoListaClearing mQueCargo, gGarantia
        Case "conyuge": CargoListaClearing mQueCargo, gConyuge
    End Select
    
End Sub

Private Sub TabClearing_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then lCSolicitud.SetFocus
    
End Sub

Private Sub TabCPersona_Click()
    
    Select Case LCase(TabCPersona.SelectedItem.Key)
        Case "titular": CargoDatosClienteClearing gCliente  'Titular
        Case "garantia": CargoDatosClienteClearing gGarantia      'Conyuge
        Case "conyuge": CargoDatosClienteClearing gConyuge      'Conyuge
    End Select

End Sub

Private Sub TabPersona_Click()

    lDeuda.Caption = ""
    lTotalComprado.Caption = ""
    lOperacion.Tag = ""
        
    Select Case LCase(TabPersona.SelectedItem.Key)
        Case "titular"
            CargoCantidadOperaciones gCliente
            If oCredito.value Then CargoOpCredito gCliente: CargoDeudaCliente gCliente
            If oContado.value Then CargoOpContado gCliente
            If oArticulo.value Then CargoOpArticulo gCliente
            If oSolicitud.value Then CargoOpSolicitud gCliente
            If oVtaTel.value Then CargoOpVtaTelefonica gCliente
        
        Case "garantia"
            CargoCantidadOperaciones gGarantia
            If oCredito.value Then CargoOpCredito gGarantia: CargoDeudaCliente gGarantia
            If oContado.value Then CargoOpContado gGarantia
            If oArticulo.value Then CargoOpArticulo gGarantia
            If oSolicitud.value Then CargoOpSolicitud gGarantia
            If oVtaTel.value Then CargoOpVtaTelefonica gGarantia
            
        Case "conyuge"
            CargoCantidadOperaciones gConyuge
            If oCredito.value Then CargoOpCredito gConyuge: CargoDeudaCliente gConyuge
            If oContado.value Then CargoOpContado gConyuge
            If oArticulo.value Then CargoOpArticulo gConyuge
            If oSolicitud.value Then CargoOpSolicitud gConyuge
            If oVtaTel.value Then CargoOpVtaTelefonica gConyuge
    End Select
    
End Sub

Private Sub HabilitoTabs()

Dim imgConyuge As String

    imgConyuge = "conyugeno"
    If gConyuge <> 0 Then
        Dim m_ECivil As Long
        m_ECivil = Val(lECivil.Tag)
        If gTipoCliente = TipoCliente.Empresa Then m_ECivil = Val(lECivilG.Tag)
        If m_ECivil = paECivilConyuge Then imgConyuge = "conyuge" Else imgConyuge = "noconyuge"
    End If
    
    Tab1.Tabs("conyuge").Image = ImageList1.ListImages(imgConyuge).Index
    Tab1.Tabs("garantia").Image = ImageList1.ListImages(IIf(gGarantia = 0, "garantiano", "garantia")).Index
    Tab1.Refresh
    
    With TabPersona
        .Tabs.Clear
        .Tabs.Add 1, pvKey:="titular", pvCaption:="Titular", pvImage:=ImageList1.ListImages("titular").Index
        If gGarantia > 0 Then .Tabs.Add pvKey:="garantia", pvCaption:="Garantía", pvImage:=ImageList1.ListImages("garantia").Index
        If gConyuge > 0 Then .Tabs.Add pvKey:="conyuge", pvCaption:="Cónyuge", pvImage:=ImageList1.ListImages(imgConyuge).Index
    End With
    
    With TabCPersona
        .Tabs.Clear
        .Tabs.Add 1, pvKey:="titular", pvCaption:="Titular", pvImage:=ImageList1.ListImages("titular").Index
        If gGarantia > 0 Then .Tabs.Add pvKey:="garantia", pvCaption:="Garantía", pvImage:=ImageList1.ListImages("garantia").Index
        If gConyuge > 0 Then .Tabs.Add pvKey:="conyuge", pvCaption:="Cónyuge", pvImage:=ImageList1.ListImages(imgConyuge).Index
    End With
    
End Sub

Private Sub EncabezadoOperaciones(Optional Credito As Boolean = False, Optional Contado As Boolean = False, _
                                                     Optional Solicitud As Boolean = False, Optional Articulo As Boolean = False, Optional VtaTel As Boolean = False)
    
    lDeuda.Caption = "": lTotalComprado.Caption = ""
    
    With lOperacion
    .Rows = 1
    
    If Credito And auxEncabezado <> "CR" Then
        auxEncabezado = "CR": .Cols = 1
        '.FormatString = "^|Fecha|Plan|^Va/De|<Cumplimiento|" _
                             & ">Importe " & paMonedaFijaTexto & "|>Cuota " & paMonedaFijaTexto & "|>Saldo " & paMonedaFijaTexto & "|" _
                             & "Garantía|^Último Pago|Mon|>Importe|>Cuota|>Saldo|Factura"
                             
        .FormatString = "^|Fecha|Plan|^Va/De|<Cumplimiento|>Importe|>Cuota|>Saldo|" _
                             & "Garantía|^Último Pago|Mon|" _
                             & ">Importe " & paMonedaFijaTexto & "|>Cuota " & paMonedaFijaTexto & "|>Saldo " & paMonedaFijaTexto & "|" _
                             & " Factura"
                             
        .ColWidth(0) = 255
        .ColWidth(1) = 900: .ColWidth(2) = 700: .ColWidth(3) = 550: .ColWidth(4) = 1250
        .ColWidth(5) = 1100: .ColWidth(6) = 1100: .ColWidth(7) = 950
        .ColWidth(8) = 1000: .ColWidth(10) = 400: .ColWidth(11) = 1100: .ColWidth(12) = 1100: .ColWidth(13) = 1100: .ColWidth(14) = 700
        
        .ColComboList(8) = "..."
        lOperacion.Refresh
    End If
    
    If Contado And auxEncabezado <> "CO" Then
        auxEncabezado = "CO": .Cols = 1
'        .FormatString = ">Fecha|Documento|Mon|>Importe|>Importe " & paMonedaFijaTexto & "|Artículos"
'        .ColWidth(0) = 1000: .ColWidth(1) = 900: .ColWidth(2) = 400: .ColWidth(3) = 1100: .ColWidth(4) = 1100
        
        .FormatString = ">Fecha|Documento|^Notas|Mon|>Importe|>Importe " & paMonedaFijaTexto & "|Artículos"
        .ColWidth(0) = 1000: .ColWidth(1) = 950: .ColWidth(2) = 700: .ColWidth(3) = 400: .ColWidth(4) = 1100: .ColWidth(5) = 1100

        lOperacion.Refresh
    End If
    
    If Solicitud And auxEncabezado <> "SO" Then
        auxEncabezado = "SO": .Cols = 1
        .FormatString = "<Código|Fecha|Tipo|>Monto " & paMonedaFijaTexto & "|Financiación|Facturada|Comentarios"
        .ColWidth(0) = 800: .ColWidth(1) = 900: .ColWidth(2) = 400: .ColWidth(3) = 1200
        lOperacion.Refresh
    End If
    
    If Articulo And auxEncabezado <> "AR" Then
        auxEncabezado = "AR": .Cols = 1
        .FormatString = ">Fecha|Operación|<Documento|Artículos"
        .ColWidth(0) = 1000
        lOperacion.Refresh
    End If
    
    If VtaTel And auxEncabezado <> "VR" Then
        auxEncabezado = "VR": .Cols = 1
        
        .FormatString = "<ID|Hora Llamada|Facturada|>Importe U$S|<Artículos|<Comentarios|Pendiente|Tel. Llamada|Usuario|Fecha Anulada"
        .ColWidth(0) = 650: .ColWidth(1) = 1250: .ColWidth(3) = 1400: .ColWidth(4) = 3200
        .ColWidth(5) = 2200: .ColWidth(6) = 1400: .ColWidth(7) = 1200: .ColWidth(8) = 850
        lOperacion.Refresh
    End If

    End With
    
End Sub

Private Sub AccionMenuFicha(Cliente As Long, tipo As Integer)
    
    On Error GoTo errObj
    Screen.MousePointer = 11
    
    Dim objCliente As New clsCliente
    If tipo = TipoCliente.Cliente Then objCliente.Personas idCliente:=Cliente
    If tipo = TipoCliente.Empresa Then objCliente.Empresas idCliente:=Cliente
    Me.Refresh
    Set objCliente = Nothing
    
    'Refresco Datos---------------------------------------------------------------------------------
    Select Case Cliente
        Case gCliente
                Screen.MousePointer = 11
                CargoDatosCliente Cliente       'Al cargar el cliente me refresca el gConyuge
                Screen.MousePointer = 11
                'If gConyuge <> 0 Then CargoDatosGarantia gConyuge
                HabilitoTabs
                
        Case gConyuge: CargoDatosGarantia Cliente
        Case gGarantia: CargoDatosGarantia Cliente
    End Select
    '----------------------------------------------------------------------------------------------------
    
    Screen.MousePointer = 0
    Exit Sub

errObj:
    clsGeneral.OcurrioError "Ocurrió un error al activar la aplicación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionMenuTitulo(Cliente As Long)
    On Error GoTo errObj
    Screen.MousePointer = 11
    
    Dim objCliente As New clsCliente
    objCliente.Titulos Cliente
    Me.Refresh
    Set objCliente = Nothing
    Screen.MousePointer = 0
    Exit Sub

errObj:
    clsGeneral.OcurrioError "Ocurrió un error al activar la aplicación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionMenuReferencia(Cliente As Long)
    
    On Error GoTo errObj
    Screen.MousePointer = 11
    
    Dim objCliente As New clsCliente
    objCliente.Referencias Cliente
    Me.Refresh
    Set objCliente = Nothing
    Screen.MousePointer = 0
    Exit Sub

errObj:
    clsGeneral.OcurrioError "Ocurrió un error al activar la aplicación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionMenuEmpleo(Cliente As Long)
    
    On Error GoTo errObj
    Screen.MousePointer = 11
    
    Dim objCliente As New clsCliente
    objCliente.Empleos Cliente
    Me.Refresh
    Set objCliente = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errObj:
    clsGeneral.OcurrioError "Ocurrió un error al activar la aplicación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionMenuComentario(Cliente As Long)
    On Error GoTo errObj
    Screen.MousePointer = 11
    
    Dim objCliente As New clsCliente
    objCliente.Comentarios Cliente
    Me.Refresh
    Set objCliente = Nothing
    Screen.MousePointer = 0
    Exit Sub

errObj:
    clsGeneral.OcurrioError "Ocurrió un error al activar la aplicación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub ConfirmarDireccion(aControl As Label, idDireccion As Long)

    If idDireccion = 0 Then Exit Sub
    Dim aResp As Integer
    aResp = MsgBox("Que desea realizar con la dirección del cliente." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
            & "Si - Confirmar dirección." & Chr(vbKeyReturn) _
            & "No - Eliminar confirmación de dirección." & Chr(vbKeyReturn) _
            & "Cancelar - Cancela la operación", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Confirmar Dirección")
    
    If aResp = vbCancel Then Exit Sub
    
    On Error GoTo errConfirmar
    Screen.MousePointer = 11
    If aResp = vbYes Then
        Cons = "Update Direccion Set DirConfirmada = 1 Where DirCodigo = " & idDireccion
    Else
        Cons = "Update Direccion Set DirConfirmada = 0 Where DirCodigo = " & idDireccion
    End If
    cBase.Execute Cons
    
    aControl.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, idDireccion, Departamento:=True, Localidad:=True, Zona:=True, ConfyVD:=True)
    If InStr(aControl.Caption, "(Cf.)") <> 0 Then aControl.ForeColor = Colores.osVerde Else aControl.ForeColor = Colores.Azul
    Screen.MousePointer = 0
    
    Exit Sub
errConfirmar:
    clsGeneral.OcurrioError "Error al confirmar la dirección del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Function ArmoNombre(Ape1 As String, Ape2 As String, Nom1 As String, Nom2 As String) As String

    ArmoNombre = Trim(Ape1) & " " & Trim(Ape2)
    ArmoNombre = Trim(ArmoNombre) & ", " & Trim(Nom1) & " " & Trim(Nom2)
    
End Function

Private Sub InicializoGrillas()

    With lEmpleoT
        .Rows = 1: .Cols = 1
        .FormatString = "<Empresa|^Ingresó|<Ocupación (Cargo)|<Tipo de Ingreso|<Exhibió"
        .ColWidth(0) = 2600: .ColWidth(2) = 1700:: .ColWidth(3) = 2400: .ColWidth(4) = 2000
        .WordWrap = False: .MergeCells = flexMergeSpill: .ExtendLastCol = True
    End With

    With lEmpleoG
        .Rows = 1: .Cols = 1
        .FormatString = "<Empresa|^Ingresó|<Ocupación (Cargo)|<Tipo de Ingreso|<Exhibió"
        .ColWidth(0) = 2600: .ColWidth(2) = 1700: .ColWidth(3) = 2400: .ColWidth(4) = 2000
        .WordWrap = False: .MergeCells = flexMergeSpill: .ExtendLastCol = True
    End With
    
    With lReferenciaT
        .Rows = 1: .Cols = 1
        .FormatString = "<Referencia|<Comprobante|<Valores|<Exhibido"
        .ColWidth(0) = 2660: .ColWidth(1) = 1500: .ColWidth(2) = 3450: .ColWidth(3) = 1400
        .WordWrap = False: .MergeCells = flexMergeSpill: .ExtendLastCol = True
    End With

    With lReferenciaG
        .Rows = 1: .Cols = 1
        .FormatString = "<Referencia|<Comprobante|<Valores|<Exhibido"
        .ColWidth(0) = 2660: .ColWidth(1) = 1500: .ColWidth(2) = 3450: .ColWidth(3) = 1400
        .WordWrap = False
        .MergeCells = flexMergeSpill: .ExtendLastCol = True
    End With
    
    With lComentarioT
        .Cols = 1: .Rows = 1: .ExtendLastCol = True
        .FormatString = "Fecha|Comentario|Tipo|Usuario|Documento"
        .WordWrap = False
        .ColWidth(0) = 800: .ColWidth(1) = 4800: .ColWidth(2) = 1200: .ColWidth(3) = 1000
        .ColDataType(0) = flexDTDate
    End With
    
    With lComentarioG
        .Cols = 1: .Rows = 1: .ExtendLastCol = True
        .FormatString = "Fecha|Comentario|Tipo|Usuario|Documento"
        .WordWrap = False
        .ColWidth(0) = 800: .ColWidth(1) = 4800: .ColWidth(2) = 1200: .ColWidth(3) = 1000
        .ColDataType(0) = flexDTDate
    End With
            
    With lOperacion
        .Cols = 1: .Rows = 1: .ExtendLastCol = True: .WordWrap = False
    End With
    
    'Articulos de la solicitud
    With lArticulo
        .Cols = 1: .Rows = 1: .ExtendLastCol = True
        .FormatString = ">Q|Artículo|>Contado|>Entrega|Financiación|>Cuota|>Tot.Financ.|"
        .WordWrap = False
        .ColWidth(0) = 450: .ColWidth(1) = 3200: .ColWidth(2) = 1200: .ColWidth(3) = 1200: .ColWidth(4) = 1200: .ColWidth(5) = 1300
        .MergeCells = flexMergeSpill
    End With
    
    With lCondicion
        .Cols = 1: .Rows = 1:
        .ExtendLastCol = True: .WordWrap = False
        .FormatString = "|^Hora|<Visto por ...|<Resolución"
        .ColWidth(0) = 300: .ColWidth(1) = 1000: .ColWidth(2) = 1300: .ColWidth(3) = 6000
        .RowHeight(0) = 260
    End With
       
    With vsRelacion
        .Rows = 1: .Cols = 1
        .FormatString = "<Relación|Documento|<Nombre|^Edad|^Q Cr.T|^Monto Tot.|^Q Cr.V|^Próx. Vto.|^Saldo Pte|"
        .ColWidth(0) = 1100: .ColWidth(1) = 1100: .ColWidth(2) = 3100: .ColWidth(3) = 500
        .ColWidth(4) = 700: .ColWidth(5) = 1150: .ColWidth(6) = 700: .ColWidth(7) = 1000: .ColWidth(8) = 1150
        .WordWrap = False: .ExtendLastCol = True
    End With
    
On Error GoTo errStart
    For I = 1 To MnuPlX.UBound
        Unload MnuPlX(I)
    Next
    
    I = 1
    Cons = "Select PlaCodigo, PlaNombre from Plantilla Where PlaCodigo IN (" & paPlantillasVOpe & ") Order by PlaNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        Load MnuPlX(I)
        With MnuPlX(I)
            .Tag = CStr(RsAux!PlaCodigo)
            .Caption = Trim(RsAux!PlaNombre)
            .Visible = True
        End With
        I = I + 1
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    MnuPlX(0).Visible = Not (MnuPlX.UBound > 0)
    tbTool.Buttons("plantillas").Enabled = (MnuPlX.UBound > 0)
    
errStart:
End Sub

Private Sub ProcesoActivacionAutomatica()
    On Error GoTo errPAA
    
    'Icono de la ficha Clearing--------------------------------------------------------------------------------------------------------------
    Dim mValorC As Integer
    mValorC = 0
    Cons = "Select * from Clearing Left Outer Join CalificaClearing On CleCalificacion = CClCodigo" & _
             " Where CleCliente = " & gCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        mValorC = 1
        If Not IsNull(RsAux!CClValor) Then
            Select Case RsAux!CClValor
                Case 0, 1, 2: mValorC = 1
                Case 3: mValorC = 3
                Case 4, 5: mValorC = 4
                Case 6: mValorC = 6
            End Select
        End If
    End If
    RsAux.Close
    Tab1.Tabs("clearing").Image = ImageList1.ListImages("clearing" & CStr(mValorC)).Index
    '------------------------------------------------------------------------------------------------------------------------------------------------
    If bDevuelta Then
        Tab1.Tabs("resolver").Selected = True
        Exit Sub
    End If
    
    Dim RsCle As rdoResultset
    Dim sActivar As Boolean: sActivar = False
    'Busco si hay antecedentes para el titular de la operacion
    Cons = "Select * from ClearingAntecedente " _
           & " Where CAnCliente = " & gCliente _
           & " And ((CAnSaldo > 0)" _
                        & " OR " _
                    & "( Datediff(year, CAnUltimoPago, getdate()) < " & paAnosAntecedentes & "))"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then sActivar = True
    RsAux.Close
    
    If sActivar Then Tab1.Tabs("clearing").Selected = True
    
errPAA:
End Sub

Private Sub CargoFinanciacion(Solicitud As Long)

Dim aEntrega As Currency, aEntregaF As Currency
Dim aCuotaF As Currency, aCuota As Currency
Dim aTotal As Currency
Dim mValorData As Long

    On Error GoTo errRenglon
    aEntrega = 0: aEntregaF = 0: aCuota = 0: aCuotaF = 0: aTotal = 0
    lArticulo.Rows = 1
            
    Cons = "Select RenglonSolicitud.*, Articulo.*, TipoCuota.*, IsNull(PCoPrecio, 0) as PrecioCtdo, ISNull(AEsNombre, ArtNombre) as NomArticulo" & _
            " From RenglonSolicitud " & _
                "INNER JOIN Articulo ON RSoArticulo = ArtID " & _
                "INNER JOIN TipoCuota ON RSoTipoCuota = TCuCodigo " & _
                "LEFT OUTER JOIN ArticuloEspecifico ON RSoSolicitud = AEsDocumento And AEsTipoDocumento = 2 And RSoArticulo = AEsArticulo " & _
                "LEFT OUTER JOIN PrecioContado ON ArtID = PCoArticulo And PCoMoneda = " & Val(lMoneda.Tag) & _
            " Where RSoSolicitud = " & Solicitud
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aTexto = ""
    Do While Not RsAux.EOF
        If RsAux!TCuVencimientoC <> 0 Then  'Cuota no la paga en el momento
            'Cuota Financiada
            aCuotaF = aCuotaF + RsAux!TCuCantidad * RsAux!RSoValorCuota
            aTexto = aTexto & RsAux!TCuCantidad & " x " & Format(RsAux!RSoValorCuota, "#,##0.00") & " + "
            
        Else    'Paga Cuota en el Momento
            aCuotaF = aCuotaF + ((RsAux!TCuCantidad - 1) * RsAux!RSoValorCuota)
            aTexto = aTexto & RsAux!TCuCantidad - 1 & " x " & Format(RsAux!RSoValorCuota, "#,##0.00") & " + "
            aCuota = aCuota + RsAux!RSoValorCuota
        End If
        
        If Not IsNull(RsAux!RSoValorEntrega) Then   'Si hay entrega veo si es financiada o no
            If RsAux!TCuVencimientoE = 0 Then  'Entrega en el momento
                aEntrega = aEntrega + RsAux!RSoValorEntrega
            Else
                aEntregaF = aEntregaF + RsAux!RSoValorEntrega      'Entrega Financiada
            End If
        End If
        
        '">Q|Artículo|>Contado|>Entrega|Financiación|>Cuota|>Tot.Financ.|"
        With lArticulo
            .AddItem CStr(RsAux!RSoCantidad)
            If Not IsNull(RsAux!ArtCodigo) Then .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "(#,000,000)")
            .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & " " & Trim(RsAux!NomArticulo)
            mValorData = RsAux!RSoArticulo: .Cell(flexcpData, .Rows - 1, 1) = mValorData
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PrecioCtdo, "#,##0.00")
                        
            If Not IsNull(RsAux!RSoValorEntrega) Then .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!RSoValorEntrega, "#,##0.00")
            
            .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!TCuAbreviacion)
            mValorData = RsAux!RSoTipoCuota: .Cell(flexcpData, .Rows - 1, 4) = mValorData
            
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!RSoValorCuota, "#,##0.00")
            
            If Not IsNull(RsAux!RSoValorEntrega) Then
                .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!RSoValorEntrega + RsAux!RSoValorCuota * RsAux!TCuCantidad, "#,##0.00")
            Else
                .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!RSoValorCuota * RsAux!TCuCantidad, "#,##0.00")
            End If
            
            aTotal = aTotal + .Cell(flexcpValue, .Rows - 1, 6)
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close

    With lArticulo
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 2, , .BackColorFixed, .ForeColorFixed, True, "Totales"
        '.Subtotal flexSTSum, -1, 2, , Colores.Rojo, Colores.Blanco, True, "Totales"
        .Subtotal flexSTSum, -1, 3: .Subtotal flexSTSum, -1, 5: .Subtotal flexSTSum, -1, 6
    End With
        
    'Cargo Texto de financiacion-----------------------------------------------------------------------------------
    If aEntregaF <> 0 Then
        lFinanciacion.Caption = Format(aEntregaF, "#,##0.00") & " + " & Mid(aTexto, 1, Len(aTexto) - 3)
    Else
        lFinanciacion.Caption = Mid(aTexto, 1, Len(aTexto) - 3)
    End If
    
    aTexto = Trim(lFinanciacion.Caption)
    lFinanciacion.Caption = "(" & Trim(lMoneda.Caption) & " " & Format((aCuotaF + aEntregaF), "#,##0.00") & ") " & aTexto
    '-------------------------------------------------------------------------------------------------------------------
    
    lFEntrega.Caption = Format(aEntregaF, "#,##0.00")
    lFCuota.Caption = Format(aCuotaF, "#,##0.00")
    lFMonto.Caption = Format(aCuotaF + aEntregaF, "#,##0.00")
    
    lSEfectivo.Caption = Format(aEntrega + aCuota, "#,##0.00")
    lSFinanciado.Caption = Format(aCuotaF + aEntregaF, "#,##0.00")
    lSMonto.Caption = Format(aCuotaF + aEntregaF + aCuota + aEntrega, "#,##0.00")
    
    cRDefinitiva.value = vbChecked: cRDefinitiva.Enabled = True
    If prmAutorizaCredHasta <> -1 Then
'        cRDefinitiva.Value = IIf(Not (CCur(lSMonto.Caption) > prmAutorizaCredHasta), vbChecked, vbUnchecked)
'        cRDefinitiva.Enabled = Not (CCur(lSMonto.Caption) > prmAutorizaCredHasta)
            '30/7/2010 cambie a pedido de carlos sólo tomar el monto de la financiación pendiente.
        cRDefinitiva.value = IIf(Not (CCur(lFMonto.Caption) > prmAutorizaCredHasta), vbChecked, vbUnchecked)
        cRDefinitiva.Enabled = Not (CCur(lFMonto.Caption) > prmAutorizaCredHasta)
    End If
    
    Exit Sub
errRenglon:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de financiación.", Err.Description
End Sub

Private Sub ZHabilitoCondiciones(Condicion As Long)

Dim sHayCondicion As Boolean

    On Error GoTo errCondicion
    
    Cons = "Select * from CondicionResolucion Where ConCodigo = " & Condicion
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    frmCondicion.Tag = RsAux!ConValor   'Estado Condicion .. de la Solicitud
    
    sHayCondicion = False
        
    'Guardo en el Tag de cCondicion la Avreviacion para ponerla en el comentario
    tComentario.Tag = Trim(RsAux!ConAbreviacion)
    lRComentario.Tag = Trim(RsAux!ConAbreviacion)
    
    If RsAux!ConValor = EstadoSolicitud.Condicional Then    'La solicitud es CONDICIONAL (Pido Valores)
        If RsAux!ConGarantia Then
            tGarantia.Enabled = True: tGarantia.BackColor = Obligatorio: tGarantia.SetFocus
            sHayCondicion = True
        End If
        
        If RsAux!ConImporte Then
            tSMonto.Enabled = True: tSMonto.BackColor = Obligatorio: tSMonto.SetFocus
            sHayCondicion = True
        End If
        
        If RsAux!ConReciboSueldo Then
            cRMoneda.Enabled = True: cRMoneda.BackColor = Obligatorio
            tRMonto.Enabled = True: tRMonto.BackColor = Obligatorio
            sHayCondicion = True: cRMoneda.SetFocus
        End If
        
        If RsAux!ConCambioPlan Then
            cPCuota.Enabled = True: cPCuota.BackColor = Obligatorio: cPCuota.SetFocus
            sHayCondicion = True
        End If
        
        If Not IsNull(RsAux!ConComprobante) Then
            'Hay que ver que comprobante es para las máscaras
            Dim RsAux2 As rdoResultset
            Cons = "Select * from Comprobante Where ComCodigo = " & RsAux!ConComprobante
            Set RsAux2 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux2.EOF Then
                lCValor1.Tag = Trim(RsAux2!ComFormatoV1)
                lCValor1.Caption = Trim(RsAux2!ComNombreValor1)
                
                If UCase(Trim(lCValor1.Tag)) = "CEDULA" Then tCValor1.Mask = "#.###.###-#"
                
                tCValor1.Enabled = True: tCValor1.BackColor = Obligatorio
                
                If Not IsNull(RsAux2!ComNombreValor2) Then
                    lCValor2.Tag = Trim(RsAux2!ComFormatoV2)
                    lCValor2.Caption = Trim(RsAux2!ComNombreValor2)
                    
                    If UCase(Trim(lCValor2.Tag)) = "CEDULA" Then tCValor2.Mask = "#.###.###-#"

                    tCValor2.Enabled = True: tCValor2.BackColor = Obligatorio
                End If
                tCValor1.SetFocus
                sHayCondicion = True
            End If
            RsAux2.Close
        End If
        
'        If Not sHayCondicion Then   'Es condicional pero no se pide ningun campo
'            zfn_TextoCondicion zfn_ArmoKeyCondicion(Condicion)
'        End If
        
    'Else    'NO ES Condicional
    '    zfn_TextoCondicion ""
    End If
    RsAux.Close
        
    If Not sHayCondicion Then
        zfn_TextoCondicion
        On Error Resume Next
        If tComentario.Text <> "" Then
            tComentario.Text = Trim(tComentario.Text) & " "
            tComentario.SelStart = Len(tComentario.Text)
            If Me.ActiveControl.Name <> "cCondicionR" Then tComentario.SetFocus Else cCondicionR.SetFocus
        End If
    End If
    
    Exit Sub
    
errCondicion:
    clsGeneral.OcurrioError "Ocurrió un error al procesar la condición", Err.Description
End Sub


'-------------------------------------------------------------------------------------------------------------------
'   Arma el Renglon con el Texto de la resolucion. Cada opción va separado entre paréntesis, para
'   saber cuando empieza y termina un valor: Ej. (FCO)(FGA10562)
'
'   ABREVIACIONES:
'       FCO         -> Firmando con el cónyuge
'       FGA         -> Garantía predeterminada........FGA y codigo de cliente garante  (FGA10562)
'       GPR         -> Garantía propietaria
'       PRO         -> Titular propietario
'       CPL          -> Cambio de Plan...........CPL y codigo de plan (CPL8) ó (CPL1E1500) - con entrega
'       CHD         -> Pago con cheque diferido
'       MON        -> Por un valor máximo a XXXX ....... MON5600
'       COM        -> Exhibir Comprobante  ... COM y cod. comp V16000V28000 (Valor uno y/o valor 2 - montos mínimos)
'       RSU         -> Recibo de Sueldo RSU , cód. Moneda , Monto .... RSU1M6500  RSU $ Monto:6500
'       CLE          -> Con Clearing
'       CDE         -> Cancelar deuda pendiente
'       COP         -> Cancelar operaciones pendiente
'-------------------------------------------------------------------------------------------------------------------
Private Function zfn_ArmoKeyCondicion(CodCondicion As Long, ByRef CondicionSINOCondicional As Byte) As String

'CondicionSINOCondicional = 1 = si, 2 = no , 3 = condicional
Dim aCodigo As Long
Dim rsRes As rdoResultset

    aTexto = ""
    
'    FUNCTION [dbo].[TextoCondicionResolucionSolicitud]
'(
'    @Condicion smallint, @Garantia int = null,
'    @CambioPlan smallint = null, @CambioPlanEntrega money = null,
'    @MontoSolicitud money = null, @MonedaRecSueldo smallint = null, @MontoRecSueldo money = null,
'    @idClienteV1 int = null, @CIV1 varchar(20) = null,
'    @idClienteV2 int = null, @CIV2 varchar(20) = null
    
    Cons = "SELECT dbo.TextoCondicionResolucionSolicitud(" & CodCondicion & ", " _
            & IIf(Val(tGarantia.Tag) > 0, Val(tGarantia.Tag), "Null") & ", "
    
    If cPCuota.ListIndex >= 0 Then
        Cons = Cons & cPCuota.ItemData(cPCuota.ListIndex)
        If Trim(tPEntrega.Text) = "" Then Cons = Cons & "E" & CCur(tPEntrega.Text)
    Else
        Cons = Cons & "Null"
    End If
    
    Cons = Cons & ", "
    If Trim(tSMonto.Text) <> "" Then Cons = Cons & CCur(tSMonto.Text) Else Cons = Cons & "Null"
    
    Cons = Cons & ", "
    If Trim(tSMonto.Text) <> "" Then Cons = Cons & CCur(tSMonto.Text) Else Cons = Cons & "Null"
    
    'cRMoneda.ItemData(cRMoneda.ListIndex) & "M" & Trim(tRMonto.Text)
    If cRMoneda.ListIndex >= 0 And Trim(tRMonto.Text) <> "" Then
        Cons = Cons & ", " & cRMoneda.ItemData(cRMoneda.ListIndex) & ", " & CCur(tRMonto.Text)
    Else
        Cons = Cons & ", Null, Null"
    End If
    
    If Val(tCValor1.Tag) <> 0 Then
        Cons = Cons & ", " & Trim(tCValor1.Tag) & ", Null"
    ElseIf tCValor1.Text <> "" Then
        Cons = Cons & ", Null, '" & FormatoReferencia(tCValor1.Text, lCValor1.Tag) & "'"
    Else
        Cons = Cons & ", NULL, NULL"
    End If
    
    If Val(tCValor2.Tag) <> 0 Then
        Cons = Cons & ", " & Trim(tCValor2.Tag) & ", Null"
    ElseIf tCValor2.Text <> "" Then
        Cons = Cons & ", Null, '" & FormatoReferencia(tCValor2.Text, lCValor2.Tag) & "'"
    Else
        Cons = Cons & ", NULL, NULL"
    End If
    Cons = Cons & ")"
    Set rsRes = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsRes.EOF Then
        If Not IsNull(rsRes(0)) Then aTexto = rsRes(0)
    End If
    rsRes.Close
    zfn_ArmoKeyCondicion = aTexto
    Exit Function
    
            
    Cons = "Select * from CondicionResolucion Where ConCodigo = " & CodCondicion
    Set rsRes = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    CondicionSINOCondicional = rsRes("ConValor")
    
    If rsRes!ConConyuge Then aTexto = aTexto & "(FCO)"
    If rsRes!ConPropiedadG Then aTexto = aTexto & "(GPR)"
    If rsRes!ConChequeD Then aTexto = aTexto & "(CHD)"
    If rsRes!ConPropiedadT Then aTexto = aTexto & "(PRO)"
    If rsRes!ConClearing Then aTexto = aTexto & "(CLE)"
    If rsRes!ConCancelarDeuda Then aTexto = aTexto & "(CDE)"
    If rsRes!ConCancelarOperacion Then aTexto = aTexto & "(COP)"
    
    'Garantía Predeterminada
    If rsRes!ConGarantia Then
        aTexto = aTexto & "(FGA" & tGarantia.Tag & ")"
    End If
    
    'Cambio de Plan
    If rsRes!ConCambioPlan Then
        If Trim(tPEntrega.Text) = "" Then
            aTexto = aTexto & "(CPL" & cPCuota.ItemData(cPCuota.ListIndex) & ")"
        Else
            aTexto = aTexto & "(CPL" & cPCuota.ItemData(cPCuota.ListIndex) & "E" & Trim(tPEntrega.Text) & ")"
        End If
    End If
    
    
    
    'Importe de Solicitud
    If rsRes!ConImporte Then aTexto = aTexto & "(MON" & Trim(tSMonto.Text) & ")"
        
    'Recibo de Sueldo
    If rsRes!ConReciboSueldo Then aTexto = aTexto & "(RSU" & cRMoneda.ItemData(cRMoneda.ListIndex) & "M" & Trim(tRMonto.Text) & ")"
        
    'Presentar Comprobantes     COMcodigo....V1 ...V2
    If Not IsNull(rsRes!ConComprobante) Then
        aTexto = aTexto & "(COM" & rsRes!ConComprobante
        If tCValor1.Tag <> 0 Then
            aTexto = aTexto & "V1" & Trim(tCValor1.Tag)
        Else
            aTexto = aTexto & "V1" & FormatoReferencia(tCValor1.Text, lCValor1.Tag)
        End If
        
        If tCValor2.Enabled Then
            If tCValor2.Tag <> 0 Then
                aTexto = aTexto & "V2" & Trim(tCValor2.Tag)
            Else
                aTexto = aTexto & "V2" & FormatoReferencia(tCValor2.Text, lCValor2.Tag)
            End If
        End If
        aTexto = aTexto & ")"
    End If
    
    rsRes.Close
    
    zfn_ArmoKeyCondicion = aTexto
    
End Function

Private Function zfn_TextoCondicion()

Dim mTXT1 As String
    
    On Error GoTo errAgregar
    mTXT1 = " ("
    
    If tGarantia.Text <> "" Then mTXT1 = mTXT1 & Trim(tGarantia.FormattedText) & "; "
    'If Trim(tSMonto.Text) <> "" Then mTXT1 = mTXT1 & Format(tSMonto.Text, "#,##0.00") & "; "
    If Trim(cRMoneda.Text) <> "" Then mTXT1 = mTXT1 & Trim(cRMoneda.Text) & " " & Format(tRMonto.Text, "#,##0.00") & "; "
    
    If Trim(cPCuota.Text) <> "" Then
        mTXT1 = mTXT1 & Trim(cPCuota.Text)
        If Trim(tPEntrega.Text) <> "" Then
            mTXT1 = mTXT1 & " E: " & Format(tPEntrega.Text, "#,##0.00") & "; "
        Else
            mTXT1 = mTXT1 & "; "
        End If
    End If
    
    If Trim(tCValor1.Text) <> "" Then mTXT1 = mTXT1 & Trim(tCValor1.Text) & "; "
    If Trim(tCValor2.Text) <> "" Then mTXT1 = mTXT1 & Trim(tCValor2.Text) & "; "
    
    If mTXT1 <> " (" Then mTXT1 = Mid(mTXT1, 1, Len(mTXT1) - 2) & ")" Else mTXT1 = ""
        
    'tComentario.Text = Trim(tComentario.Text)
    'tComentario.Text = tComentario.Text & IIf(tComentario.Text = "", "", ", ") & Trim(tComentario.Tag) & mTXT1
    tComentario.Text = Trim(tComentario.Tag) & mTXT1
    Exit Function

errAgregar:
    clsGeneral.OcurrioError "Ocurrió un error al agregar la condición. Reintente.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub ZDeshabilitoCondiciones()

    mnuMCondicion.Tag = ""
    
    tGarantia.Text = "": tGarantia.Enabled = False: tGarantia.BackColor = vbButtonFace
    tSMonto.Text = "": tSMonto.Enabled = False: tSMonto.BackColor = vbButtonFace
    cRMoneda.Text = "": cRMoneda.Enabled = False: cRMoneda.BackColor = vbButtonFace
    tRMonto.Text = "": tRMonto.Enabled = False: tRMonto.BackColor = vbButtonFace
    cPCuota.Text = "": cPCuota.Enabled = False: cPCuota.BackColor = vbButtonFace
    tPEntrega.Text = "": tPEntrega.Enabled = False: tPEntrega.BackColor = vbButtonFace
    
    tCValor1.Mask = "": tCValor1.Text = "": tCValor1.Enabled = False: tCValor1.BackColor = vbButtonFace
    tCValor2.Text = "": tCValor2.Mask = "": tCValor2.Enabled = False: tCValor2.BackColor = vbButtonFace
    
    lCValor1.Caption = "Valor1:": lCValor1.Tag = "":
    lCValor2.Caption = "Valor2:": lCValor2.Tag = ""
            
End Sub

Private Sub tbTool_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case LCase(Button.Key)
        Case "autor": acc_AccionAutoEvaluar
        Case "resolverysig": acc_ResolverSolicitud irASiguiente:=True
        Case "resolver": acc_ResolverSolicitud irASiguiente:=False
        
        Case "soltar": acc_SoltarSolicitud
        Case "llamara": acc_LlamarA
        Case "plantillas": acc_MnuPlantillas
                
        Case "salir": Unload Me
    
    End Select
        
End Sub

Private Sub tComentario_GotFocus()
    On Error Resume Next
    If tComentario.Text <> "" Then
        tComentario.Text = Trim(tComentario.Text) & " "
        tComentario.SelStart = Len(tComentario.Text)
    End If
End Sub

Private Sub tComentario_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errFnc

    If KeyCode = vbKeyF1 Then
        
        dic_LoadDiccionario
        
        Cons = "Select DicIncorrecto as 'Abreviación', DicCorrecto as 'Sustituir por ...' " & _
                  " from Diccionario " & _
                  " Where DicTipo = 4 Order by DicIncorrecto"
                  
        Dim mHlp As New clsListadeAyuda
        mHlp.ActivarAyuda cBase, Cons, 4000, , "Lista de Abreviaciones"
        Set mHlp = Nothing
        
    End If
    Exit Sub

errFnc:
    clsGeneral.OcurrioError "Error al activar la lista de abreviaciones.", Err.Description
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        acc_ResolverSolicitud irASiguiente:=True
        Exit Sub
    End If
    
    On Error GoTo errKD
    If KeyAscii = vbKeySpace Or Chr(KeyAscii) = "." Or Chr(KeyAscii) = "," Then
        Dim iSS As Integer
        iSS = tComentario.SelStart
        tComentario.Text = dic_PulirTexto(tComentario.Text, iSS)
        tComentario.SelStart = iSS
    End If
        
errKD:

End Sub

Private Sub tCValor1_Change()
    tCValor2.Tag = 0
End Sub

Private Sub tCValor1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tCValor1.Tag = 0
        'Valido la Cédula ingresada----------
        If UCase(Trim(lCValor1.Tag)) = "CEDULA" Then
            Screen.MousePointer = 11
            If Trim(tCValor1.Text) <> "_.___.___-_" Then
                If Not clsGeneral.CedulaValida(clsGeneral.QuitoFormatoCedula(tCValor1.Text)) Then
                    Screen.MousePointer = 0
                    MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
            End If
            'Busco el Cliente -----------------------
            If Trim(tCValor1.Text) <> "_.___.___-_" Then
                Cons = "Select CliCodigo from Cliente Where CliCiRuc = '" & clsGeneral.QuitoFormatoCedula(tCValor1) & "'"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux.EOF Then
                    tCValor1.Tag = RsAux!CliCodigo
                Else
                    Screen.MousePointer = 0
                    MsgBox "No existe un cliente para el documento ingresado.", vbExclamation, "ATENCIÓN"
                End If
                RsAux.Close
            End If
        End If
        Screen.MousePointer = 0
        
        If tCValor2.Enabled Then
            Foco tCValor2
        Else
            PierdoFoco tCValor1
        End If
    End If

End Sub

Private Sub tCValor2_Change()
    tCValor2.Tag = 0
End Sub

Private Sub tCValor2_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        tCValor2.Tag = 0
        'Valido la Cédula ingresada----------
        If UCase(Trim(lCValor2.Tag)) = "CEDULA" Then
            Screen.MousePointer = 11
            If Trim(tCValor2.Text) <> "_.___.___-_" Then
                If Not clsGeneral.CedulaValida(clsGeneral.QuitoFormatoCedula(tCValor2.Text)) Then
                    Screen.MousePointer = 0
                    MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
            End If
            'Busco el Cliente -----------------------
            If Trim(tCValor2.Text) <> "_.___.___-_" Then
                Cons = "Select CliCodigo from Cliente Where CliCiRuc = '" & clsGeneral.QuitoFormatoCedula(tCValor2.Text) & "'"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux.EOF Then
                    tCValor2.Tag = RsAux!CliCodigo
                Else
                    Screen.MousePointer = 0
                    MsgBox "No existe un cliente para el documento ingresado.", vbExclamation, "ATENCIÓN"
                End If
                RsAux.Close
            End If
        End If
        Screen.MousePointer = 0
        
        If val_ValidoResolucion Then tComentario.SetFocus
        
    End If
        
End Sub

Private Sub tGarantia_Change()
    tGarantia.Tag = 0
End Sub

Private Sub tGarantia_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        tGarantia.Tag = 0
        'Valido la Cédula ingresada----------
        Screen.MousePointer = 11
        If Trim(tGarantia.Text) <> "" Then
            If Not clsGeneral.CedulaValida(tGarantia.Text) Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        'Busco el Cliente -----------------------
        If Trim(tGarantia.Text) <> "" Then
            Cons = "Select CliCodigo from Cliente Where CliCiRuc = '" & Trim(tGarantia.Text) & "'"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                tGarantia.Tag = RsAux!CliCodigo
            Else
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para el documento ingresado.", vbExclamation, "ATENCIÓN"
            End If
            RsAux.Close
        End If

        If Val(tGarantia.Tag) <> 0 Then PierdoFoco tGarantia
        Screen.MousePointer = 0
    End If
        
    'If KeyAscii = vbKeyEscape Then Call bCCancelar_Click
    
End Sub

Private Sub PierdoFoco(C1 As Control)

    On Error Resume Next
    If tGarantia.TabIndex > C1.TabIndex Then If tGarantia.Enabled Then tGarantia.SetFocus: Exit Sub
    If tSMonto.TabIndex > C1.TabIndex Then If tSMonto.Enabled Then tSMonto.SetFocus: Exit Sub
    
    If cRMoneda.TabIndex > C1.TabIndex Then If cRMoneda.Enabled Then cRMoneda.SetFocus: Exit Sub
    If tRMonto.TabIndex > C1.TabIndex Then If tRMonto.Enabled Then tRMonto.SetFocus: Exit Sub
    
    If cPCuota.TabIndex > C1.TabIndex Then If cPCuota.Enabled Then cPCuota.SetFocus: Exit Sub
    If tPEntrega.TabIndex > C1.TabIndex Then If tPEntrega.Enabled Then tPEntrega.SetFocus: Exit Sub
    
    If tCValor1.TabIndex > C1.TabIndex Then If tCValor1.Enabled Then tCValor1.SetFocus: Exit Sub
    If tCValor2.TabIndex > C1.TabIndex Then If tCValor2.Enabled Then tCValor2.SetFocus: Exit Sub
    
    If val_ValidoResolucion Then tComentario.SetFocus
    
End Sub

Private Sub tmClearing_Timer()
    '2- OK por CI
    '3- Ok por Nombres
    '4- Sinonimos por CI
    '5- Sinonimos Nombres
    '6- 2 fichas vinculadas.
    On Error GoTo errClearing
    tmClearing.Enabled = False
    
    Dim aEvento As Integer, aClienteC As Long, aTag As String, aNombreCC As String, aMensajeC As String
    Dim mIdxTab As String
    
    aTag = tmClearing.Tag
    aEvento = Trim(Mid(aTag, 1, InStr(aTag, ":") - 1))
    
    aTag = Trim(Mid(aTag, InStr(aTag, ":") + 1))
    aClienteC = Trim(Mid(aTag, 1, InStr(aTag, ":") - 1))
    aMensajeC = Trim(Mid(aTag, InStr(aTag, ":") + 1))
    
    Select Case aClienteC
        Case gCliente: aNombreCC = Trim(lTitular.Caption): mIdxTab = "titular"
        Case gConyuge: aNombreCC = Trim(lGarantia.Caption): mIdxTab = "conyuge"
        Case gGarantia: aNombreCC = Trim(lGarantia.Caption): mIdxTab = "garantia"
    End Select
    
    Tab1.Tabs("clearing").Selected = True
    TabCPersona.Tabs(mIdxTab).Selected = True
    Exit Sub
    
errClearing:
    clsGeneral.OcurrioError "Error al cargar los datos del clearing.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tPEntrega_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PierdoFoco tPEntrega
End Sub

Private Sub tPEntrega_LostFocus()
    If IsNumeric(tPEntrega.Text) Then tPEntrega.Text = Format(tPEntrega.Text, FormatoMonedaP)
End Sub

Private Sub tRMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PierdoFoco tRMonto
End Sub

Private Sub tRMonto_LostFocus()
    If IsNumeric(tRMonto.Text) Then tRMonto.Text = Format(tRMonto.Text, FormatoMonedaP)
End Sub

Private Sub tSMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PierdoFoco tSMonto
        cCondicionR_KeyPress vbKeyReturn
    End If
End Sub

Private Sub tSMonto_LostFocus()
    If IsNumeric(tSMonto.Text) Then tSMonto.Text = Format(tSMonto.Text, FormatoMonedaP)
End Sub


Private Sub CargoDireccionesAuxiliares(aCombo As ComboBox, aIDCliente As Long)

    On Error GoTo errCDA
    Dim rsDA As rdoResultset
    'Direcciones Auxiliares-----------------------------------------------------------------------
    Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & aIDCliente
    Set rsDA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDA.EOF Then
        Do While Not rsDA.EOF
            aCombo.AddItem Trim(rsDA!DAuNombre)
            aCombo.ItemData(aCombo.NewIndex) = rsDA!DAuDireccion
            rsDA.MoveNext
        Loop
        
        If aCombo.ListCount > 1 Then aCombo.BackColor = Colores.Blanco
        If Val(aCombo.Tag) = 0 And aCombo.ListCount > 0 Then aCombo.Text = aCombo.List(0)
    End If
    rsDA.Close
    
    If aCombo.ListCount > 1 Then aCombo.Visible = True Else aCombo.Visible = False
    aCombo.Refresh
    
errCDA:
End Sub

Private Sub vsRelacion_DblClick()
On Error GoTo errActivo
    If vsRelacion.Rows = 1 Then Exit Sub
    
    Dim mIDAuxiliar As Long
    mIDAuxiliar = vsRelacion.Cell(flexcpData, vsRelacion.Row, 1)
    
    If mIDAuxiliar = -1 Then
        CargoDatosClienteRelaciones gCliente, Full:=True
    Else
        EjecutarApp pathApp & "Visualizacion de operaciones", CStr(mIDAuxiliar)
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errActivo:
    clsGeneral.OcurrioError "Error al activar el formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub acc_AccionAutoEvaluar()
    On Error GoTo errEvaluar
    If gIdSolicitud = 0 Then Exit Sub
    
    Dim mOBJ As New clsResolAuto
    mOBJ.VerCaminoAlResolver cBase, gIdSolicitud
    Set mOBJ = Nothing
    
    Screen.MousePointer = 0
    Exit Sub
    
errEvaluar:
    clsGeneral.OcurrioError "Error al evaluar la condición", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function code_ProcesoScript(mEvento As String, mIDCliente As Long, mIDSolicitud As Long) As String
On Error GoTo errCode
    
    If prmPlantillaPuente = 0 Then Exit Function
    Screen.MousePointer = 11
    code_ProcesoScript = ""
    
    Dim mFmt As Integer, mResult As String, mParams As String
    Dim objCode As New clsPlantillaI
    
    mFmt = 1
    
    mParams = "EVE=" & mEvento & "|" & _
                      "CLI=" & mIDCliente & "|" & _
                      "SOL=" & mIDSolicitud
                      
    If objCode.ProcesoPlantillaInteractiva(cBase, prmPlantillaPuente, mFmt, mResult, "", mParams, False) Then
        code_ProcesoScript = mResult
    End If
    
    Set objCode = Nothing
    Screen.MousePointer = 0
    Exit Function

errCode:
    clsGeneral.OcurrioError "Error al procesar código externo.", Err.Description
    Screen.MousePointer = 0
End Function


Private Sub CargoListaClearing(miCaso As Integer, miCliente As Long)

    On Error GoTo errCargarL
    'Tabls del Clearing    -------------------------------------------------------------------------------------------------
    With lCSolicitud
        .Cols = 1: .Rows = 1: .ExtendLastCol = True
        .WordWrap = False
        
        Select Case miCaso
            Case 1
                .FormatString = "Afiliado|Empresa|Fecha|Tipo|>Importe|<Cuotas"
                .ColWidth(0) = 750: .ColWidth(1) = 3000: .ColWidth(2) = 950: .ColWidth(3) = 550: .ColWidth(4) = 1300
            Case 2
                .FormatString = "Afiliado|Empresa|Fecha|Tipo|>Importe|<Cuotas|>Saldo|<U.Pago"
                .ColWidth(0) = 750: .ColWidth(1) = 3000: .ColWidth(2) = 950: .ColWidth(3) = 550: .ColWidth(4) = 1300: .ColWidth(5) = 650: .ColWidth(6) = 1300
            Case 3
                .FormatString = "Afiliado|Tipo|Nº de Cuenta|Banco|<Nº Cheque|>Importe|Emisión|<Vencimiento"
                .ColWidth(0) = 750: .ColWidth(1) = 550: .ColWidth(2) = 1100: .ColWidth(3) = 1500: .ColWidth(4) = 1100: .ColWidth(5) = 1100: .ColWidth(6) = 950
            Case 4
                .FormatString = "Fecha|<Observaciones"
                .ColWidth(0) = 950
        End Select
    End With
    
    If miCliente = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    'Cargo Datos Tabla ClearingSolicitud------------------------------------------------------------------
    If miCaso = 1 Then
        Cons = "Select * from ClearingSolicitud inner join ClearingAfiliados on CSoAfiliado = CAfNumero Where CSoCliente= " & miCliente _
                & " Order by CSoFecha Desc"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            With lCSolicitud
                .AddItem Trim(RsAux!CSoAfiliado)
                
                If Not IsNull(RsAux!CAfNombre) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!CAfNombre)
                
                .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!CSoFecha, "dd/mm/yyyy")
                
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!CSoTipo)
                If Not IsNull(RsAux!CSoMonto) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!CSoMoneda) & " " & Format(RsAux!CSoMonto, FormatoMonedaP)
                
                If Not IsNull(RsAux!CSoCuota) Then .Cell(flexcpText, .Rows - 1, 5) = RsAux!CSoCuota
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    '-----------------------------------------------------------------------------------------------------
    
    'Cargo Datos Tabla ClearingAntecedente------------------------------------------------------------------
    If miCaso = 2 Then
        Cons = "Select * from ClearingAntecedente INNER JOIN ClearingAfiliados on CAnAfiliado = CAfNumero Where CAnCliente= " & miCliente _
                & " Order by CAnFecha Desc"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            With lCSolicitud
                .AddItem Trim(RsAux!CAnAfiliado)
            
                If Not IsNull(RsAux!CAfNombre) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!CAfNombre)
                .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!CAnFecha, "dd/mm/yy")
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!CAnTipo)
                
                If Not IsNull(RsAux!CAnMonto) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!CAnMoneda) & " " & Format(RsAux!CAnMonto, FormatoMonedaP)
                
                If Not IsNull(RsAux!CAnCuota) Then .Cell(flexcpText, .Rows - 1, 5) = RsAux!CAnCuota
                
                If RsAux!CAnSaldo = 0 Then
                    .Cell(flexcpText, .Rows - 1, 6) = "*******"
                    If RsAux!CAnEnClearing = 0 Then
                        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris
                    Else
                        If DateDiff("yyyy", RsAux!CAnFecha, Date) <= paAnosAntecedentes Then .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
                    End If
                Else
                    .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!CAnSaldo, FormatoMonedaP)
                    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.RojoClaro: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
                End If
                    
                If Not IsNull(RsAux!CAnUltimoPago) Then .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!CAnUltimoPago, "mm/yy")
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    '-----------------------------------------------------------------------------------------------------
    
    'Cargo Datos Tabla ClearingCheques------------------------------------------------------------
    If miCaso = 3 Then
        Cons = "Select * from ClearingCheque Where CChCliente= " & miCliente _
                & " Order by CChCodigo Desc"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            With lCSolicitud
                .AddItem Trim(RsAux!CChAfiliado)
                If Not IsNull(RsAux!CChPropioEndozado) Then .Cell(flexcpText, .Rows - 1, 1) = RsAux!CChPropioEndozado
                If Not IsNull(RsAux!CChNroCuenta) Then .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!CChNroCuenta)
                
                If Not IsNull(RsAux!CChBanco) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!CChBanco)
                        
                If Not IsNull(RsAux!CChSerie) Then .Cell(flexcpText, .Rows - 1, 4) = RsAux!CChSerie
                If Not IsNull(RsAux!CChNroCheque) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(.Cell(flexcpText, .Rows - 1, 4)) & " " & Trim(RsAux!CChNroCheque)
                
                If Not IsNull(RsAux!CChMoneda) Then .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!CChMoneda)
                If Not IsNull(RsAux!CChImporte) Then .Cell(flexcpText, .Rows - 1, 5) = .Cell(flexcpText, .Rows - 1, 5) & " " & Format(RsAux!CChImporte, FormatoMonedaP)
                
                If Not IsNull(RsAux!CChEmision) Then .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!CChEmision, "dd/mm/yy")
                If Not IsNull(RsAux!CChVencimiento) Then .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!CChVencimiento, "dd/mm/yy")
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    '-----------------------------------------------------------------------------------------------------
    
    'Cargo Datos Tabla ClearingComentario------------------------------------------------------------
    If miCaso = 4 Then
        Cons = "Select * from ClearingComentario Where CCoCliente= " & miCliente _
                & " Order by CCoFecha Desc"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            With lCSolicitud
                .AddItem Format(RsAux!CCoFecha, "dd/mm/yyyy")
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!CCoTexto)
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCargarL:
    clsGeneral.OcurrioError "Error al cargar los datos del clearing.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function GrabarPreAutorizacion(Importe As Currency)
On Error GoTo errGPre
Dim rs1 As rdoResultset, mSQL As String

    mSQL = "Select * from CreditoPreAutorizado " & _
            " Where CPACliente = " & gCliente & _
            " And CPAFecha = '" & Format(Date, "yyyy/mm/dd") & "'"
            
    Set rs1 = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If rs1.EOF Then rs1.AddNew Else rs1.Edit
    
    rs1!CPACliente = gCliente
    rs1!CPAFecha = Date
    rs1!CPAVence = DateAdd("d", prmValidezLimiteCredito, Date)
    rs1!CPAUsuario = paCodigoDeUsuario
    rs1!CPALimite = Importe
    'rs1!CPAPrevioClearing = 0

    rs1.Update
    rs1.Close
    Exit Function
    
errGPre:
    clsGeneral.OcurrioError "Error al grabar Crédito PreAutorizado.", Err.Description
    Screen.MousePointer = 0
End Function

Public Sub PresentoPaisDelDocumentoCliente(ByVal paisDelDoc As Integer, ByVal cirut As String, ByRef lblInfo As Label, ByRef lblCIRUT As Label)
    
    lblInfo.ForeColor = vbBlack
    lblInfo.FontBold = False
    
    Dim TipoDoc As clsTipoDocumento
    Set TipoDoc = mTiposDocumentos.ObtenerTipoDocumento(paisDelDoc)
    If TipoDoc Is Nothing Then Exit Sub
    Select Case TipoDoc.ID
        Case 1
            lblInfo.Caption = "Cédula"
            If cirut = "" Then
                lblCIRUT.Caption = "N/D"
            Else
                lblCIRUT.Caption = clsGeneral.RetornoFormatoCedula(cirut)
            End If
        
        Case 2
            lblInfo.Caption = "RUT"
            If cirut = "" Then
                lblCIRUT.Caption = "N/D"
            Else
                lblCIRUT.Caption = clsGeneral.RetornoFormatoRuc(cirut)
            End If
        
        Case Else
            lblInfo.Caption = TipoDoc.Abreviacion
            lblCIRUT.Caption = cirut
            lblInfo.ForeColor = &H40C0&
            lblInfo.FontBold = True
    End Select
End Sub



