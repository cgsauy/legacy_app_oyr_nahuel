VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form MaArticulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos"
   ClientHeight    =   5940
   ClientLeft      =   2670
   ClientTop       =   3435
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaArt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7500
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   1
      Left            =   240
      ScaleHeight     =   4215
      ScaleWidth      =   6975
      TabIndex        =   78
      Top             =   1260
      Width           =   6975
      Begin VB.Frame Frame7 
         Caption         =   "Listas de Precios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1755
         Left            =   0
         TabIndex        =   87
         Top             =   2400
         Width           =   6975
         Begin AACombo99.AACombo cLista 
            Height          =   315
            Left            =   1080
            TabIndex        =   47
            Top             =   300
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   556
            BackColor       =   12648447
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
         Begin VB.TextBox tEVidriera 
            Height          =   285
            Left            =   6120
            MaxLength       =   3
            TabIndex        =   51
            Top             =   660
            Width           =   735
         End
         Begin VB.TextBox tENormal 
            Height          =   285
            Left            =   6120
            MaxLength       =   3
            TabIndex        =   49
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox tALargo 
            Appearance      =   0  'Flat
            Height          =   645
            Left            =   1080
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   55
            Top             =   1020
            Width           =   5775
         End
         Begin VB.TextBox tACorto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   53
            Top             =   660
            Width           =   3915
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Et. Vidriera:"
            Height          =   255
            Left            =   5100
            TabIndex        =   50
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Et. Normales:"
            Height          =   255
            Left            =   5100
            TabIndex        =   48
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Arg. Lar&go:"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Arg. Cor&to:"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Li&sta:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2415
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Width           =   6975
         Begin VB.CheckBox cNroSerie 
            Alignment       =   1  'Right Justify
            Caption         =   "Con Número de Serie:"
            Height          =   255
            Left            =   4140
            TabIndex        =   39
            Top             =   1380
            Width           =   1935
         End
         Begin VB.CheckBox cWebSPrecio 
            Alignment       =   1  'Right Justify
            Caption         =   "We&b Sin Precio"
            Height          =   255
            Left            =   5040
            TabIndex        =   36
            Top             =   1080
            Width           =   1755
         End
         Begin VB.CheckBox cAMercaderia 
            Alignment       =   1  'Right Justify
            Caption         =   "A Mercadería Gral."
            Height          =   195
            Left            =   5040
            TabIndex        =   34
            Top             =   600
            Width           =   1755
         End
         Begin VB.CheckBox cEnUso 
            Alignment       =   1  'Right Justify
            Caption         =   "En &Uso"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   915
         End
         Begin AACombo99.AACombo cLocal 
            Height          =   315
            Left            =   1320
            TabIndex        =   38
            Top             =   1320
            Width           =   2475
            _ExtentX        =   4366
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
         Begin AACombo99.AACombo cIVA 
            Height          =   315
            Left            =   1320
            TabIndex        =   27
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            BackColor       =   12648447
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
         Begin AACombo99.AACombo cCategoria 
            Height          =   315
            Left            =   1320
            TabIndex        =   41
            Top             =   1680
            Width           =   1575
            _ExtentX        =   2778
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
         Begin VB.TextBox tBarCode 
            Height          =   285
            Left            =   4140
            MaxLength       =   30
            TabIndex        =   29
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox tComentarioF 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   150
            TabIndex        =   45
            Top             =   2025
            Width           =   5535
         End
         Begin VB.TextBox tCompra 
            Height          =   285
            Left            =   5880
            MaxLength       =   4
            TabIndex        =   43
            Top             =   1665
            Width           =   975
         End
         Begin VB.CheckBox cHabilitado 
            Alignment       =   1  'Right Justify
            Caption         =   "&Habilitado para Venta"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1935
         End
         Begin VB.CheckBox cExclusivo 
            Alignment       =   1  'Right Justify
            Caption         =   "Venta &Exclusiva"
            Height          =   255
            Left            =   2820
            TabIndex        =   33
            Top             =   840
            Width           =   1515
         End
         Begin VB.CheckBox cInterior 
            Alignment       =   1  'Right Justify
            Caption         =   "Venta al I&nterior"
            Height          =   255
            Left            =   2820
            TabIndex        =   32
            Top             =   600
            Width           =   1515
         End
         Begin VB.CheckBox cArtEnWeb 
            Alignment       =   1  'Right Justify
            Caption         =   "Artículo en &Web"
            Height          =   255
            Left            =   5040
            TabIndex        =   35
            Top             =   840
            Width           =   1755
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Local de &Retiro:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1335
            Width           =   1335
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Código de &Barras:"
            Height          =   255
            Left            =   2820
            TabIndex        =   28
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "C&omentarios:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   2025
            Width           =   975
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad &Mínima de Compra:"
            Height          =   255
            Left            =   3720
            TabIndex        =   42
            Top             =   1665
            Width           =   2175
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Categ. de &Dto:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1665
            Width           =   1335
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de I.&V.A.:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   0
      Left            =   240
      ScaleHeight     =   4215
      ScaleWidth      =   6975
      TabIndex        =   77
      Top             =   1260
      Width           =   6975
      Begin VB.Frame Frame5 
         Caption         =   "Grupos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2175
         Left            =   3000
         TabIndex        =   83
         Top             =   1920
         Width           =   3975
         Begin AACombo99.AACombo cGrupo 
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Top             =   240
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
         Begin MSComctlLib.ListView lGrupo 
            Height          =   1455
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dimensiones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2175
         Left            =   0
         TabIndex        =   81
         Top             =   1920
         Width           =   2775
         Begin VB.TextBox tAlto 
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   17
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox tProfundidad 
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   19
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox tFrente 
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox tPeso 
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   21
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox tVolumen 
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   23
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "A&lto (cm):"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "P&rofundidad (cm):"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "F&rente (cm):"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Pe&so (kg):"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "V&olumen (Lt):"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1815
         Left            =   0
         TabIndex        =   80
         Top             =   0
         Width           =   6975
         Begin VB.TextBox tProveedor 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   960
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1320
            Width           =   3135
         End
         Begin AACombo99.AACombo cMarca 
            Height          =   315
            Left            =   960
            TabIndex        =   5
            Top             =   600
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            BackColor       =   12648447
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
         Begin AACombo99.AACombo cTipo 
            Height          =   315
            Left            =   4320
            TabIndex        =   3
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            BackColor       =   12648447
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
         Begin AACombo99.AACombo cGarantia 
            Height          =   315
            Left            =   4920
            TabIndex        =   13
            Top             =   1320
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
         Begin VB.TextBox tNombre 
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            MaxLength       =   40
            TabIndex        =   9
            Top             =   960
            Width           =   5895
         End
         Begin VB.TextBox tDescripcion 
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4320
            MaxLength       =   20
            TabIndex        =   7
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox tCodigo 
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            MaxLength       =   7
            TabIndex        =   1
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Modificado:"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4080
            TabIndex        =   86
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lModificacion 
            Alignment       =   2  'Center
            Caption         =   "Mar 12-Oct 1998 12:43"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4920
            TabIndex        =   85
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "&Garantía:"
            Height          =   255
            Left            =   4200
            TabIndex        =   12
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "&Proveedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "&Nombre:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "&Descripción:"
            Height          =   255
            Left            =   3360
            TabIndex        =   6
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "&Marca:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "&Tipo:"
            Height          =   255
            Left            =   3360
            TabIndex        =   2
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "&Código:"
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Index           =   2
      Left            =   240
      ScaleHeight     =   4335
      ScaleWidth      =   6975
      TabIndex        =   79
      Top             =   1200
      Width           =   6975
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   0
         TabIndex        =   91
         Top             =   3120
         Width           =   6975
         Begin VB.TextBox tArtRelacionadoCon 
            Height          =   285
            Left            =   120
            TabIndex        =   72
            Text            =   "Text1"
            Top             =   480
            Width           =   3135
         End
         Begin VSFlex6DAOCtl.vsFlexGrid vsRelacionadoCon 
            Height          =   1045
            Left            =   3300
            TabIndex        =   73
            Top             =   120
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   1843
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
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   1
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
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Relacionado con:"
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   1395
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Datos de Importación"
         Height          =   2955
         Left            =   0
         TabIndex        =   84
         Top             =   120
         Width           =   6975
         Begin VB.TextBox tProvExt 
            Height          =   285
            Left            =   1920
            TabIndex        =   58
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox tComentarioI 
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   70
            Top             =   2520
            Width           =   6555
         End
         Begin VB.TextBox tCodigoFabrica 
            Height          =   285
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   68
            Top             =   2040
            Width           =   3375
         End
         Begin AACombo99.AACombo cPuerto 
            Height          =   315
            Left            =   1920
            TabIndex        =   64
            Top             =   1320
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
         Begin AACombo99.AACombo cContenedor 
            Height          =   315
            Left            =   1920
            TabIndex        =   60
            Top             =   960
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
         Begin VB.TextBox tDemora 
            Height          =   285
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   66
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox tCantContenedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5880
            MaxLength       =   5
            TabIndex        =   62
            Top             =   960
            Width           =   735
         End
         Begin VB.CheckBox cImportada 
            Alignment       =   1  'Right Justify
            Caption         =   "&Mercadería Importada"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "C&omentarios:"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lProvExt 
            Caption         =   "Proveedor del exterior:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "&Demora de Fabricación:"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Código de Fabrica:"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "&Tipo de Contenedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "&Puerto de Embarque:"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "&Cantidad por Contenedor:"
            Height          =   255
            Left            =   3840
            TabIndex        =   61
            Top             =   960
            Width           =   2055
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir. [Ctrl+X]"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "caracteristica"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "datosweb"
            Object.ToolTipText     =   "Datos Web. [Ctrl+W]"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refrescar"
            Object.ToolTipText     =   "Refrescar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   4100
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   75
      Top             =   5685
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11509
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1640
            MinWidth        =   2
            Text            =   "F1 - Ayuda "
            TextSave        =   "F1 - Ayuda "
         EndProperty
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaArt.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaArt.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaArt.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaArt.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaArt.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaArt.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaArt.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaArt.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaArt.frx":10E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaArt.frx":13FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaArt.frx":1716
            Key             =   "datosweb"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   76
      Top             =   780
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ficha del &Artículo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Datos de &Facturación"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Datos de &Importación"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lUsuMod 
      BackStyle       =   0  'Transparent
      Caption         =   "Modificado por:  Adrián "
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5220
      TabIndex        =   90
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lFalta 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha alta:  12- Oct de 1999 19:30:35"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1740
      TabIndex        =   89
      Top             =   480
      Width           =   2835
   End
   Begin VB.Label lUsuAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Alta:  Adrián "
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   88
      Top             =   480
      Width           =   1455
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
      Begin VB.Menu MnuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefrescar 
         Caption         =   "&Refrescar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MnuPrecio 
         Caption         =   "Ingreso de &Precios"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuOpCaracteristicas 
         Caption         =   "Características de Artículos"
      End
      Begin VB.Menu MnuOpDatosWeb 
         Caption         =   "Datos Web"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu MnuMantenimientos 
      Caption         =   "&Mantenimientos"
      Begin VB.Menu MnuMaGrupo 
         Caption         =   "&Grupo"
      End
      Begin VB.Menu MnuMaMarca 
         Caption         =   "M&arca"
      End
      Begin VB.Menu MnuMaTipo 
         Caption         =   "&Tipo"
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
Attribute VB_Name = "MaArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sNuevo As Boolean, sModificar As Boolean, sAyuda As Boolean
Private RsArticulo As rdoResultset

Private frm1Campo As New MaUnCampo

Private aFecha As Date      'Para controlar modificaciones multiusuario

Private Function fnc_ShowTipo(ByVal bModal As Boolean) As Boolean
On Error GoTo ErrAcceso
    If Not miConexion.AccesoAlMenu("Mantenimiento de Tipo") Then
        MsgBox "Ud no tiene permiso para acceder al formulario de Tipos de Artículos.", vbCritical, "ATENCIÓN"
        Exit Function
    End If
    Screen.MousePointer = 11
    EjecutarApp App.Path & "\tipoarticulo.exe", , bModal
    fnc_ShowTipo = True
'    MaTipo.pSeleccionado = 0
 '   MaTipo.pTipoLlamado = TipoLlamado.Visualizacion
  '  MaTipo.Show vbModeless, Me
    Exit Function
ErrAcceso:
    clsGeneral.OcurrioError "Error al acceder a tipos.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function fnc_ShowGrupo(ByVal bModal As Boolean) As Boolean
On Error GoTo ErrAcceso
    If Not miConexion.AccesoAlMenu("Mantenimiento de Grupo") Then
        MsgBox "Ud no tiene permiso para acceder al formulario de Tipos de Artículos.", vbCritical, "ATENCIÓN"
        Exit Function
    End If
    Screen.MousePointer = 11
    EjecutarApp App.Path & "\gruposarticulos.exe", , bModal
    fnc_ShowGrupo = True
    Exit Function
ErrAcceso:
    clsGeneral.OcurrioError "Error al acceder a grupos.", Err.Description
    Screen.MousePointer = 0
End Function



Private Sub cAMercaderia_GotFocus()
    Status.Panels(1).Text = "Indique si el artículo va a mercadería general."
End Sub

Private Sub cAMercaderia_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If cArtEnWeb.Visible Then cArtEnWeb.SetFocus Else cWebSPrecio.SetFocus
    End If
End Sub

Private Sub cArtEnWeb_GotFocus()
    Status.Panels(1).Text = "Indique si el artículo está publicado en la Web."
End Sub

Private Sub cArtEnWeb_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cWebSPrecio.SetFocus
End Sub

Private Sub cContenedor_GotFocus()
    cContenedor.SelStart = 0
    cContenedor.SelLength = Len(cContenedor.Text)
    Status.Panels(1).Text = "Seleccione el tipo de contenedor."
End Sub
Private Sub cContenedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tCantContenedor
End Sub
Private Sub cContenedor_LostFocus()
    cContenedor.SelLength = 0
    Status.Panels(1).Text = ""
End Sub

Private Sub cEnUso_GotFocus()
    Status.Panels(1).Text = "Indique si el artículo esta en uso."
End Sub
Private Sub cEnUso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cHabilitado.Enabled Then cHabilitado.SetFocus
End Sub
Private Sub cEnUso_LostFocus()
        Status.Panels(1).Text = ""
End Sub

Private Sub cExclusivo_GotFocus()
    Status.Panels(1).Text = "Indique si el artículo es venta exclusiva."
End Sub

Private Sub cExclusivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cAMercaderia.SetFocus
End Sub

Private Sub cGarantia_GotFocus()
   
    cGarantia.SelStart = 0
    cGarantia.SelLength = Len(cGarantia.Text)
    Status.Panels(1).Text = "Seleccione el tipo de garantía para el artículo."
    
End Sub

Private Sub cGarantia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFrente
End Sub
Private Sub cGarantia_LostFocus()
    cGarantia.SelLength = 0
    Status.Panels(1).Text = ""
End Sub
Private Sub cGrupo_GotFocus()
    With cGrupo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.Panels(1).Text = "Seleccione un grupo al cual pertenece el artículo."
End Sub

Private Sub cGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrKP
    If KeyCode = vbKeyReturn Then
        If cGrupo.ListIndex <> -1 Then
            If lGrupo.ListItems.Count > 0 Then
                For I = 1 To lGrupo.ListItems.Count
                    If lGrupo.ListItems(I).Key = "A" & Str(cGrupo.ItemData(cGrupo.ListIndex)) Then
                        MsgBox "El grupo seleccionado ya ha sido ingresado.", vbExclamation, "ATENCIÓN"
                        Foco cGrupo
                        Exit Sub
                    End If
                Next I
            End If
            Set itmx = lGrupo.ListItems.Add(, "A" & Str(cGrupo.ItemData(cGrupo.ListIndex)), Trim(cGrupo.Text))
            cGrupo.ListIndex = -1
        Else
            If Trim(cGrupo.Text) <> "" And (sNuevo Or sModificar) Then
                'Ingreso de un Nuevo Grupo de Artículo-------------------------------------------------------------------
                If MsgBox("El grupo ingresado no existe. Desea ingresarlo ?", vbQuestion + vbOKCancel, "ATENCIÓN") = vbOK Then
                    Screen.MousePointer = 11
                    fnc_ShowGrupo True
                    Cons = "Select GruCodigo, GruNombre From Grupo Order by GruNombre"
                    CargoCombo Cons, cGrupo, ""
                    'AccederAGrupo TipoLlamado.IngresoNuevo
                    Screen.MousePointer = 0
                End If  '--------------------------------------------------------------------------------------------------------
            Else
                If KeyCode = vbKeyReturn And Trim(cGrupo.Text) = "" Then Foco lGrupo
            End If
        End If
    End If
    Exit Sub
ErrKP:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cGrupo_LostFocus()
    Status.Panels(1).Text = ""
    cGrupo.SelLength = 0
End Sub
Private Sub cHabilitado_GotFocus()
    Status.Panels(1).Text = "Indique si el artículo está habilitado para la venta."
End Sub

Private Sub cHabilitado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cInterior
End Sub


Private Sub cImportada_Click()

    If Not sNuevo And Not sModificar Then Exit Sub
    
    If cImportada.Value = 0 Then
        cContenedor.Enabled = False
        tProvExt.Enabled = False
        cPuerto.Enabled = False
        tDemora.Enabled = False
        tCodigoFabrica.Enabled = False: tCodigoFabrica.BackColor = Inactivo
        tComentarioI.Enabled = False
        tCantContenedor.Enabled = False
        cContenedor.BackColor = Inactivo
        tProvExt.BackColor = Inactivo
        cPuerto.BackColor = Inactivo
        tDemora.BackColor = Inactivo
        tComentarioI.BackColor = Inactivo
        tCantContenedor.BackColor = Inactivo
    Else
        cContenedor.Enabled = True
        tProvExt.Enabled = True
        cPuerto.Enabled = True
        tDemora.Enabled = True
        tComentarioI.Enabled = True
        tCodigoFabrica.Enabled = True: tCodigoFabrica.BackColor = Blanco
        tCantContenedor.Enabled = True
        cContenedor.BackColor = Blanco
        tProvExt.BackColor = Blanco
        cPuerto.BackColor = Blanco
        tDemora.BackColor = Blanco
        tComentarioI.BackColor = Blanco
        tCantContenedor.BackColor = Blanco
    End If
    
End Sub

Private Sub cImportada_GotFocus()

    Status.Panels(1).Text = "Indique si la mercadería es importada."
    
End Sub

Private Sub cImportada_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cContenedor.Enabled Then
        If tProvExt.Enabled And tProvExt.Visible Then
            Foco tProvExt
        Else
            Foco cContenedor
        End If
    End If
End Sub
Private Sub cInterior_GotFocus()
    Status.Panels(1).Text = "Indique si el artículo se vende al interior."
End Sub
Private Sub cInterior_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cExclusivo
End Sub
Private Sub cIva_GotFocus()
    cIVA.SelStart = 0
    cIVA.SelLength = Len(cIVA.Text)
    Status.Panels(1).Text = "Seleccione el tipo de I.V.A. para facturar el artículo."
End Sub
Private Sub cIva_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tBarCode
End Sub
Private Sub cIva_LostFocus()
    Status.Panels(1).Text = ""
    cIVA.SelLength = 0
End Sub

Private Sub cLista_GotFocus()
    cLista.SelStart = 0
    cLista.SelLength = Len(cLista.Text)
    Status.Panels(1).Text = "Seleccione la lista de precios para el artículo."
End Sub
Private Sub cLista_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tENormal
End Sub
Private Sub cLista_LostFocus()
    Status.Panels(1).Text = ""
    cLista.SelLength = 0
End Sub

Private Sub cMarca_GotFocus()
    cMarca.SelStart = 0
    cMarca.SelLength = Len(cMarca.Text)
    Status.Panels(1).Text = "Seleccione la marca del artículo."
    Status.Panels(2).Enabled = True
End Sub

Private Sub cMarca_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrKD
    If KeyCode = vbKeyF1 And cMarca.ListIndex <> -1 And Not sNuevo And Not sModificar Then
        Cons = "Select ArtId, Nombre = ArtNombre from Articulo" _
                & " Where ArtMarca = " & cMarca.ItemData(cMarca.ListIndex) _
                & " Order by ArtNombre"
        AyudaArticulo Cons
    End If
    
    If KeyCode = vbKeyReturn Then
        If cMarca.ListIndex = -1 And Trim(cMarca.Text) <> "" And (sNuevo Or sModificar) Then
            'Ingreso de una Nueva Marca de Artículo-------------------------------------------------------------------
            If MsgBox("La marca ingresada no existe. Desea ingresarla ?", vbQuestion + vbOKCancel, "ATENCIÓN") = vbOK Then
                Screen.MousePointer = 11
                AccederAMarca TipoLlamado.IngresoNuevo
                Screen.MousePointer = 0
                If cMarca.ListIndex > -1 Then Foco tDescripcion
            End If  '--------------------------------------------------------------------------------------------------------
        Else
            If tDescripcion.Enabled Then
                Foco tDescripcion
            Else
                Foco tNombre
            End If
        End If
    End If
    Exit Sub
ErrKD:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cMarca_LostFocus()
    cMarca.SelStart = 0
    Status.Panels(2).Enabled = False
    Status.Panels(1).Text = ""
End Sub

Private Sub cPuerto_GotFocus()

    cPuerto.SelStart = 0
    cPuerto.SelLength = Len(cPuerto.Text)
    Status.Panels(1).Text = "Seleccione el puerto de embarque de la mercadedría."
    
End Sub
Private Sub cPuerto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDemora
End Sub
Private Sub cPuerto_LostFocus()
    Status.Panels(1).Text = ""
    cPuerto.SelLength = 0
End Sub

Private Sub cTipo_GotFocus()

    cTipo.SelStart = 0
    cTipo.SelLength = Len(cTipo.Text)
    
    Status.Panels(1).Text = "Seleccione el tipo del artículo."
    Status.Panels(2).Enabled = True

End Sub

Private Sub cTipo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrKD
    If KeyCode = vbKeyF1 And cTipo.ListIndex <> -1 And Not sNuevo And Not sModificar Then
        Cons = "Select ArtId, Nombre = ArtNombre from Articulo" _
                & " Where ArtTipo = " & cTipo.ItemData(cTipo.ListIndex) _
                & " Order by ArtTipo"
        AyudaArticulo Cons
    End If
    
    If KeyCode = vbKeyReturn Then
        If cTipo.ListIndex = -1 And Trim(cTipo.Text) <> "" And (sNuevo Or sModificar) Then
            'Ingreso de un Nuevo Tipo de Artículo-------------------------------------------------------------------
            If MsgBox("El tipo ingresado no existe. Desea ingresarlo ?", vbQuestion + vbOKCancel, "ATENCIÓN") = vbOK Then
                Screen.MousePointer = 11
                If fnc_ShowTipo(True) Then
'                MaTipo.pSeleccionado = 0
 '               MaTipo.pTipoLlamado = TipoLlamado.IngresoNuevo
  '              MaTipo.Show vbModal, Me
                
'                If MaTipo.pSeleccionado <> 0 Then
                    'Cargo los TIPOS DE ARTICULOS
                    Cons = "Select TipCodigo, TipNombre From Tipo Order by TipNombre"
                    CargoCombo Cons, cTipo, ""
'                    BuscoCodigoEnCombo cTipo, MaTipo.pSeleccionado
'                End If
                End If
                Me.Refresh
                Screen.MousePointer = 0
            End If  '--------------------------------------------------------------------------------------------------------
        Else
            Foco cMarca
        End If
    End If
 
    Exit Sub
ErrKD:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Trim(Err.Description)
End Sub

Private Sub cTipo_LostFocus()
    cTipo.SelLength = 0
    Status.Panels(1).Text = ""
    Status.Panels(2).Enabled = False
End Sub

Private Sub cWebSPrecio_GotFocus()
    Status.Panels(1).Text = "Indique si el artículo se va a publicar sin precio."
End Sub

Private Sub cWebSPrecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cLocal.SetFocus
End Sub

Private Sub Form_Activate()
    If Not sAyuda Then RsArticulo.Requery
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad

    cArtEnWeb.Visible = miConexion.AccesoAlMenu("Datos_Web")
    ObtengoSeteoForm Me
    Me.Height = 6600
    MnuPrecio.Enabled = False
    Botones True, False, False, False, False, Toolbar1, Me
    
    SetearLView lvValores.FullRow, lGrupo
    
    sNuevo = False: sModificar = False
    
    InicializoCombos
    
    Cons = "Select * FROM Articulo Where ArtCodigo = " & 0
    Set RsArticulo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    LimpioCampos
    DeshabilitoIngreso
    Picture1(0).ZOrder 0
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al ingresar al formulario.", Err.Description
    DeshabilitoIngreso
    Cons = "Select * FROM Articulo Where ArtCodigo = " & 0
    Set RsArticulo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.SimpleText = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If sNuevo Or sModificar Then
        If MsgBox("Ud. realizó modificaciones en la ficha y no ha grabado." & Chr(13) & _
            "Desea almacenar la información ingresada.", vbYesNo + vbQuestion, "ATENCIÓN") = vbYes Then
            
            AccionGrabar
            
            If sNuevo Or sModificar Then
                Cancel = True
                Exit Sub
            End If
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    GuardoSeteoForm Me
    RsArticulo.Close
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    Forms(Forms.Count - 2).SetFocus
    End
    
End Sub

Private Sub Label1_Click()
    Foco tCodigo
End Sub

Private Sub Label10_Click()
    Foco tPeso
End Sub

Private Sub Label11_Click()
    Foco tAlto
End Sub

Private Sub Label12_Click()
    Foco tProveedor
End Sub

Private Sub Label13_Click()
    Foco cCategoria
End Sub

Private Sub Label14_Click()
    Foco cGarantia
End Sub

Private Sub Label15_Click()
    Foco cIVA
End Sub

Private Sub Label16_Click()
    Foco tVolumen
End Sub

Private Sub Label17_Click()
    Foco tBarCode
End Sub

Private Sub Label19_Click()
    Foco tDemora
End Sub
Private Sub Label2_Click()
    Foco cTipo
End Sub
Private Sub Label20_Click()
    Foco cContenedor
End Sub
Private Sub Label21_Click()
    Foco tComentarioI
End Sub
Private Sub Label22_Click()
    Foco cPuerto
End Sub
Private Sub Label23_Click()
    Foco tCantContenedor
End Sub
Private Sub Label25_Click()
    Foco cLocal
End Sub

Private Sub Label29_Click()
    Foco tArtRelacionadoCon
End Sub

Private Sub Label3_Click()
    Foco cMarca
End Sub
Private Sub Label4_Click()
    Foco tDescripcion
End Sub
Private Sub Label5_Click()
    Foco tNombre
End Sub
Private Sub Label6_Click()
    Foco tComentarioF
End Sub
Private Sub Label7_Click()
    Foco tProfundidad
End Sub
Private Sub Label8_Click()
    Foco tCompra
End Sub
Private Sub Label9_Click()
   Foco tFrente
End Sub

Private Sub lGrupo_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not sNuevo And Not sModificar Then Exit Sub
    If KeyCode = vbKeyDelete And lGrupo.ListItems.Count > 0 Then
        lGrupo.ListItems.Remove (lGrupo.SelectedItem.Key)
    End If
    If KeyCode = vbKeyReturn Then
        Tab1.Tabs(2).Selected = True
        Picture1(1).ZOrder 0
         Foco cIVA
    End If

End Sub

Private Sub lProvExt_Click()
    Foco tProvExt
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
Private Sub MnuMaGrupo_Click()
    fnc_ShowGrupo False
End Sub
Private Sub MnuMaMarca_Click()
    AccederAMarca TipoLlamado.Visualizacion
End Sub
Private Sub MnuMaTipo_Click()

    fnc_ShowTipo False

End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub
Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuOpCaracteristicas_Click()
    If cTipo.ListIndex > -1 Then
        EjecutarApp App.Path & "\Caracteristicas.exe", CStr(cTipo.ItemData(cTipo.ListIndex))
    Else
        EjecutarApp App.Path & "\Caracteristicas.exe"
    End If
End Sub

Private Sub MnuOpDatosWeb_Click()
    AccionDatosWeb
End Sub

Private Sub MnuPrecio_Click()
On Error GoTo errPrecio
    Screen.MousePointer = 11
    EjecutarApp paPathApp & "Precio_Articulo.exe", RsArticulo!ArtID
    Exit Sub
errPrecio:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los precios del artículo.", Err.Description
End Sub

Private Sub MnuRefrescar_Click()
    AccionRefrescar
End Sub
Private Sub MnuVolver_Click()
    Unload Me
End Sub

Sub AccionNuevo()
    EjecutarApp App.Path & "\wiz_articulo.exe", IIf(Val(tCodigo.Tag) > 0, "p" & tCodigo.Tag, "")
End Sub

Private Sub LimpioCampos()
    
    'Ficha de Datos Articulo--------------------------
    tCodigo.Text = "": tCodigo.Tag = ""
    cTipo.Text = ""
    cMarca.Text = ""
    tNombre.Text = ""
    tDescripcion.Text = ""
    cGarantia.Text = ""
    tProveedor.Text = ""
    
    tVolumen.Text = ""
    tAlto.Text = ""
    tProfundidad.Text = ""
    tPeso.Text = ""
    tFrente.Text = ""
    
    cGrupo.Text = ""
    lGrupo.ListItems.Clear
    
    'Ficha de Facturacion-----------------------------
    cIVA.Text = ""
    cEnUso.Value = 0
    cHabilitado.Value = 0
    cExclusivo.Value = 0
    cInterior.Value = 0
    cArtEnWeb.Value = 0
    cWebSPrecio.Value = 0
    cAMercaderia.Value = 0
       
    tComentarioF.Text = ""
    cCategoria.Text = ""
    tCompra.Text = ""
    tBarCode.Text = ""
    cLocal.Text = ""
    
    lUsuAlta.Caption = "Alta: " & String(12, "-")
    lFalta.Caption = "Fecha alta: " & Space(2) & String(20, "-")
    lUsuMod.Caption = "Modificado por: " & String(12, "-")
    lModificacion.Caption = String(20, "-")
    cLista.Text = "": tENormal.Text = "": tEVidriera.Text = "": tACorto.Text = "": tALargo.Text = ""
    
    'Ficha de Importacion-----------------------------
    cImportada.Value = 0
    tProvExt.Text = "": tProvExt.Tag = ""
    cContenedor.Text = ""
    cPuerto.Text = ""
    tDemora.Text = ""
    tCodigoFabrica.Text = ""
    tComentarioI.Text = ""
    tCantContenedor.Text = ""
    
    vsRelacionadoCon.Rows = 0
    tArtRelacionadoCon.Text = ""
    
End Sub

Sub AccionModificar()

    'Prendo Señal que es modificaci{on.
    sModificar = True
    
    'Habilito y Desabilito Botones.
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    MnuPrecio.Enabled = False
    MnuMaGrupo.Enabled = False
    MnuMaMarca.Enabled = False
    MnuMaMarca.Enabled = False
    
    HabilitoIngreso
    
    Screen.MousePointer = 11
    CargoDatosArticulo
    Screen.MousePointer = 0
    Foco tCodigo

End Sub

Sub AccionGrabar()

Dim aCodigo As Long, sDefensa As String, UIDSuceso As Long

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
        Screen.MousePointer = 11
        
        If Not sNuevo Then
            RsArticulo.Requery
            
            If aFecha <> RsArticulo!ArtModificado Then
                Screen.MousePointer = 0
                MsgBox "La ficha ha sido modificada por otro usuario. Verifique los datos antes de grabar.", vbExclamation, "ATENCIÓN"
                AccionCancelar
                Exit Sub
            End If
            
            UIDSuceso = 0
            'Veo si lleva suceso.
            If RsArticulo!ArtTipo <> cTipo.ItemData(cTipo.ListIndex) Then
                'Verifico si alguno de los dos es del tipo servicio.
                If RsArticulo!ArtTipo = paTipoArticuloServicio Or cTipo.ItemData(cTipo.ListIndex) = paTipoArticuloServicio Then
                    'Suceso
                    Dim objSuceso As New clsSuceso
                    objSuceso.ActivoFormulario miConexion.UsuarioLogueado(True), "Cambio de Tipo de Artículo", cBase
                    UIDSuceso = objSuceso.RetornoValor(True, False)
                    sDefensa = objSuceso.RetornoValor(False, True)
                    If UIDSuceso = 0 Then
                        Exit Sub
                    End If
                    Dim sDescSuceso As String
                    If RsArticulo!ArtTipo = paTipoArticuloServicio Then
                        sDescSuceso = "De Tipo Servicio a " & cTipo.Text
                    Else
                        sDescSuceso = "De Otro (id = " & RsArticulo!ArtTipo & ") a Tipo Servicio"
                    End If
                End If
            End If
            
            On Error GoTo ErrBT
            cBase.BeginTrans    'COMIENZO LA TRANSACCION--------------------------
            On Error GoTo ErrET
            
            aCodigo = RsArticulo!ArtID
            RsArticulo.Requery
            
            RsArticulo.Edit
            CargoCamposBDArticulo
            RsArticulo.Update
            
            CargoCamposBDArticuloFacturacion aCodigo
            
            CargoCamposBDGrupos aCodigo
            
            If cImportada.Value = 1 Then
                CargoCamposBDArticuloImportacion aCodigo
            Else
                'Verifico para borrar (por si hay algún registro)----------------------------------
                Cons = "Select * from ArticuloImportacion Where AImArticulo = " & aCodigo
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    RsAux.Delete
                End If
                RsAux.Close '-------------------------------------------------------------------------
            End If
        
            CargoCamposWeb aCodigo
            GuardoRelacionadoCon aCodigo
            
            If UIDSuceso > 0 Then
                clsGeneral.RegistroSuceso cBase, Now, TipoSuceso.CambioTipoArticuloServicio, paCodigoDeTerminal, UIDSuceso, 0, RsArticulo!ArtID, sDescSuceso, sDefensa
            End If
            sModificar = False
            cBase.CommitTrans   'FIN DE LA TRANSACCION-------------------------------
            RsArticulo.Requery
        End If
        lUsuMod.Caption = "Modificado por: " & BuscoUsuario(miConexion.UsuarioLogueado(True), True)
        lModificacion.Caption = Format(Now, "Ddd d-Mmm yyyy hh:mm")
        Call Botones(True, True, True, False, False, Toolbar1, Me)
        MnuPrecio.Enabled = True
        MnuMaGrupo.Enabled = True
        MnuMaMarca.Enabled = True
        MnuMaMarca.Enabled = True
        DeshabilitoIngreso
        
        Tab1.Tabs(1).Selected = True
        Picture1(0).ZOrder 0
        Foco tCodigo
    
        Screen.MousePointer = 0
        
    End If
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "No se pudo iniciar la transacción.", Err.Description
    RsArticulo.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
ErrET:
    Resume ErrRoll
ErrRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se pudo almacenar la ficha del artículo, reintente."
    RsArticulo.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Sub AccionEliminar()

    If MsgBox("Confrima eliminar el artículo: " & Trim(RsArticulo!ArtNombre), vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        Screen.MousePointer = 11
        
        If ExisteEnCosteo(RsArticulo!ArtID) Then Screen.MousePointer = 0: Exit Sub
        
        On Error GoTo ErrBT
        cBase.BeginTrans    'COMIENZO LA TRANSACCION--------------------------
        On Error GoTo ErrET
        
        RsArticulo.Requery
        
        'Elimino Tabla ArticuloImportacion
        Cons = "Select * from ArticuloImportacion Where AImArticulo = " & RsArticulo!ArtID
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Delete
        End If
        RsAux.Close
        
        'Elimino Tabla ArticuloFacturacion
        Cons = "Select * from ArticuloFacturacion Where AFaArticulo = " & RsArticulo!ArtID
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Delete
        End If
        RsAux.Close
        
        'Elimino Relacion ArticuloGrupo
        Cons = "Delete ArticuloGrupo Where AGrArticulo = " & RsArticulo!ArtID
        cBase.Execute (Cons)
        
        'Elimino art. relacionadoCon
        Cons = "Delete RelacionArticulo Where RArArticulo = " & RsArticulo!ArtID
        cBase.Execute (Cons)
        
        RsArticulo.Delete
        
        cBase.CommitTrans    'FINALIZO LA TRANSACCION--------------------------
        
        RsArticulo.Requery
                
        LimpioCampos
        Call Botones(True, False, False, False, False, Toolbar1, Me)
        MnuPrecio.Enabled = False
        Foco tCodigo
        Screen.MousePointer = 0
    End If
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "No se pudo iniciar la transacción."
    RsArticulo.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
ErrET:
    Resume ErrRoll
ErrRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se pudo eliminar la ficha del artículo."
    RsArticulo.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Sub AccionCancelar()

    Screen.MousePointer = 11
    LimpioCampos
    
    If sNuevo Then
        Call Botones(True, False, False, False, False, Toolbar1, Me)
        MnuPrecio.Enabled = False
    Else
        Call Botones(True, True, True, False, False, Toolbar1, Me)
        MnuPrecio.Enabled = True
        CargoDatosArticulo
    End If
    MnuMaGrupo.Enabled = True
    MnuMaMarca.Enabled = True
    MnuMaMarca.Enabled = True
    sNuevo = False
    sModificar = False
    
    DeshabilitoIngreso
    Foco tCodigo
    Screen.MousePointer = 0
    
End Sub

Private Sub Tab1_Click()

    If Tab1.SelectedItem.Index = 1 Then
        Picture1(Tab1.SelectedItem.Index - 1).ZOrder 0
        Me.Refresh
    ElseIf Tab1.SelectedItem.Index = 2 Then
        If miConexion.AccesoAlMenu("Datos de Facturacion") Then
            Picture1(Tab1.SelectedItem.Index - 1).ZOrder 0
            Me.Refresh
        Else
            Tab1.Tabs(1).Selected = True
            Picture1(0).ZOrder 0
            Me.Refresh
            MsgBox "Ud no tiene permisos para acceder a esta ficha.", vbExclamation, "ATENCIÓN"
        End If
    ElseIf Tab1.SelectedItem.Index = 3 Then
        If miConexion.AccesoAlMenu("Datos de Importacion") Then
            tProvExt.Visible = True
        Else
            tProvExt.Visible = False
        End If
        lProvExt.Visible = tProvExt.Visible
        Picture1(Tab1.SelectedItem.Index - 1).ZOrder 0
        Me.Refresh
    End If

End Sub

Private Sub tACorto_GotFocus()
Status.Panels(1).Text = ""
End Sub

Private Sub tACorto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tALargo
End Sub

Private Sub tALargo_GotFocus()
Status.Panels(1).Text = ""
End Sub

Private Sub tALargo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 0 And (sNuevo Or sModificar) Then
        Tab1.Tabs(3).Selected = True
        Picture1(2).ZOrder 0
        cImportada.SetFocus
    End If
End Sub

Private Sub tAlto_GotFocus()

    tAlto.SelStart = 0
    tAlto.SelLength = Len(tAlto.Text)
    
    Status.Panels(1).Text = "Ingrese la medida del alto (en centímetros)."
    
End Sub

Private Sub tAlto_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Foco tProfundidad
    
End Sub

Private Sub tAlto_LostFocus()
    
    tAlto.Text = Format(tAlto.Text, FormatoMonedaP)
    
End Sub

Private Sub cCategoria_GotFocus()
    cCategoria.SelStart = 0
    cCategoria.SelLength = Len(cCategoria.Text)
    Status.Panels(1).Text = "Ingrese la categoría del artículo (para descuentos)."
End Sub
Private Sub cCategoria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tCompra
End Sub
Private Sub cCategoria_LostFocus()
    cCategoria.SelLength = 0
    Status.Panels(1).Text = ""
End Sub

Private Sub tArtRelacionadoCon_Change()
    If Val(tArtRelacionadoCon.Tag) <> 0 Then tArtRelacionadoCon.Tag = ""
End Sub

Private Sub tArtRelacionadoCon_GotFocus()
    With tArtRelacionadoCon
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tArtRelacionadoCon_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tArtRelacionadoCon.Text) <> "" Then
            If Val(tArtRelacionadoCon.Tag) <> 0 Then
                InsertoArtRelacionado
                tArtRelacionadoCon.Text = ""
            Else
                BuscoArticuloRelacionado tArtRelacionadoCon
            End If
        Else
            AccionGrabar
        End If
    End If
    Exit Sub

End Sub

Private Sub tBarCode_GotFocus()
    tBarCode.SelStart = 0
    tBarCode.SelLength = Len(tBarCode.Text)
    Status.Panels(1).Text = "Ingrese el código de barras del artículo."
End Sub

Private Sub tBarCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cEnUso.SetFocus
End Sub

Private Sub tCantContenedor_GotFocus()

    tCantContenedor.SelStart = 0
    tCantContenedor.SelLength = Len(tCantContenedor.Text)

    Status.Panels(1).Text = "Ingrese la cantidad de artículos por contenedor."
    
End Sub

Private Sub tCantContenedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Foco cPuerto
    
End Sub

Private Sub tCodigo_Change()
    If Not sNuevo And Not sModificar And Val(tCodigo.Tag) > 0 Then LimpioCampos
End Sub

Private Sub tCodigo_GotFocus()

    tCodigo.SelStart = 0
    tCodigo.SelLength = Len(tCodigo)

    Status.Panels(1).Text = "Ingrese el código del artículo."
    
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)

Dim aCodigo As Long

    If KeyAscii = vbKeyReturn Then
        If Not sModificar And Not sNuevo Then
            If IsNumeric(tCodigo) Then
                Cons = "Select * from Articulo Where ArtCodigo = " & tCodigo.Text
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    aCodigo = RsAux!ArtID
                    RsAux.Close
                    BuscoArticulo aCodigo
                Else
                    Botones True, False, False, False, False, Toolbar1, Me
                    MnuPrecio.Enabled = False
                    RsAux.Close
                    MsgBox "No existe un artículo para el código ingresado.", vbExclamation, "ATENCIÓN"
                End If
            End If
        Else
            Foco cTipo
        End If
    End If
    
End Sub

Private Sub tCodigoFabrica_GotFocus()
    With tCodigoFabrica
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCodigoFabrica_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentarioI
End Sub

Private Sub tCodigoFabrica_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub tComentarioF_GotFocus()

    tComentarioF.SelStart = 0
    tComentarioF.SelLength = Len(tComentarioF.Text)
    Status.Panels(1).Text = "Ingrese un comentario de facturación para el artículo."
    
End Sub

Private Sub tComentarioF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cLista
End Sub

Private Sub tComentarioI_GotFocus()
    tComentarioI.SelStart = 0
    tComentarioI.SelLength = Len(tComentarioI.Text)
    Status.Panels(1).Text = "Ingrese un comentario de importación."
End Sub

Private Sub tComentarioI_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Foco tArtRelacionadoCon
    
End Sub

Private Sub tCompra_GotFocus()

    tCompra.SelStart = 0
    tCompra.SelLength = Len(tCompra.Text)
    
    Status.Panels(1).Text = "Ingrese la cantidad mínima de compra para aplicar el descuento."
    
End Sub

Private Sub tCompra_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Foco tComentarioF
    
End Sub

Private Sub tDemora_GotFocus()

    tDemora.SelStart = 0
    tDemora.SelLength = Len(tDemora.Text)
    Status.Panels(1).Text = "Ingrese la demora de fabricación (en días)."
    
End Sub

Private Sub tDemora_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Foco tCodigoFabrica
    
End Sub

Private Sub TDescripcion_GotFocus()

    tDescripcion.SelStart = 0
    tDescripcion.SelLength = Len(tDescripcion)

    Status.Panels(1).Text = "Ingrese la descripción del artículo."
    
End Sub

Private Sub TDescripcion_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        On Error GoTo errTexto
        Dim aNombre As String

        If (sNuevo Or sModificar) And cTipo.ListIndex <> -1 And cMarca.ListIndex <> -1 Then
            
            aNombre = ""
            If UCase(Trim(cTipo.Text)) <> "N/D" Then
                Cons = "Select * from Tipo Where TipCodigo = " & cTipo.ItemData(cTipo.ListIndex)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not IsNull(RsAux!TipAbreviacion) Then aNombre = Trim(RsAux!TipAbreviacion) & " " Else: aNombre = Trim(cTipo.Text) & " "
                RsAux.Close
            End If
            If UCase(Trim(cMarca.Text)) <> "N/D" Then aNombre = aNombre & Trim(cMarca.Text) & " "
            If UCase(Trim(tDescripcion.Text)) <> "N/D" Then aNombre = aNombre & Trim(tDescripcion)
        
            If Trim(tNombre.Text) <> "" And Trim(tNombre.Text) <> Trim(aNombre) Then
                If MsgBox("Desea modificar el nombre del artículo.", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
            End If
            tNombre = Trim(aNombre)
        
        End If

        Foco tNombre
    End If
    Exit Sub
    
errTexto:
    clsGeneral.OcurrioError "Ocurrió un error al armar la descripción. " & Err.Description
End Sub

Private Sub tENormal_GotFocus()
    Status.Panels(1).Text = "Cantidad de etiquetas normales por defecto."
End Sub

Private Sub tENormal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tEVidriera
End Sub

Private Sub tEVidriera_GotFocus()
Status.Panels(1).Text = "Cantidad de etiquetas vidriera por defecto."
End Sub

Private Sub tEVidriera_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tACorto
End Sub

Private Sub tFrente_GotFocus()
    
    tFrente.SelStart = 0
    tFrente.SelLength = Len(tFrente.Text)
    
    Status.Panels(1).Text = "Ingrese la medida del frente (en centímetros)."

End Sub

Private Sub tFrente_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Foco tAlto
    
End Sub

Private Sub tFrente_LostFocus()

    tFrente.Text = Format(tFrente.Text, FormatoMonedaP)
    
End Sub

Private Sub tNombre_GotFocus()

    tNombre.SelStart = 0
    tNombre.SelLength = Len(tNombre.Text)

    Status.Panels(2).Enabled = True
    Status.Panels(1).Text = "Ingrese el nombre completo del artículo."
    
End Sub

Private Sub tNombre_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 And Trim(tNombre.Text) <> "" And Not sNuevo And Not sModificar Then
        Cons = "Select ArtId, Nombre = ArtNombre from Articulo" _
            & " Where ArtNombre LIKE '" & clsGeneral.Replace(Trim(tNombre.Text), " ", "%") & "%'" _
            & " Order By ArtNombre"
        AyudaArticulo Cons
        tNombre.SelStart = 0: tNombre.SelLength = Len(tNombre.Text)
    End If

End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tNombre.Text) <> "" And Not sNuevo And Not sModificar Then
            Cons = "Select ArtId, Nombre = ArtNombre from Articulo" _
                & " Where ArtNombre LIKE '" & clsGeneral.Replace(tNombre.Text, " ", "%") & "%'" _
                & " Order By ArtNombre"
            AyudaArticulo Cons
            tNombre.SelStart = 0: tNombre.SelLength = Len(tNombre.Text)
        Else
            If tProveedor.Enabled Then
                Foco tProveedor
            Else
                Foco tCodigo
            End If
        End If
    End If

End Sub

Private Sub tNombre_LostFocus()

    Status.Panels(2).Enabled = False
        
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        
        Case "caracteristica"
            If cTipo.ListIndex > -1 Then
                EjecutarApp App.Path & "\Caracteristicas.exe", CStr(cTipo.ItemData(cTipo.ListIndex))
            Else
                EjecutarApp App.Path & "\Caracteristicas.exe"
            End If

        Case "nuevo": AccionNuevo
        
        Case "modificar": AccionModificar
        
        Case "eliminar": AccionEliminar
        
        Case "grabar": AccionGrabar
        
        Case "cancelar": AccionCancelar
        
        Case "refrescar": AccionRefrescar
            
        Case "datosweb": AccionDatosWeb
            
        Case "salir"
            Unload Me
            
    End Select

End Sub

Sub CargoDatosArticulo()

    On Error GoTo errCargar
    aFecha = RsArticulo!ArtModificado
    
    'CARGO DATOS DE TABLA Articulo -------------------
    tCodigo.Text = RsArticulo!ArtCodigo
    tCodigo.Tag = RsArticulo!ArtCodigo
    
    
    tNombre = Trim(RsArticulo!ArtNombre)
    tDescripcion = Trim(RsArticulo!ArtDescripcion)
    
    BuscoCodigoEnCombo cTipo, RsArticulo!ArtTipo
    BuscoCodigoEnCombo cMarca, RsArticulo!ArtMarca

    If Not IsNull(RsArticulo("ArtProveedor")) Then loc_FindProveedor RsArticulo("ArtProveedor")
    
    If RsArticulo!ArtEnUso Then cEnUso.Value = 1 Else cEnUso.Value = 0
    If Not IsNull(RsArticulo!ArtHabilitado) Then If RsArticulo!ArtHabilitado = "S" Then cHabilitado.Value = 1
    
    If Not IsNull(RsArticulo!ArtVolumen) Then tVolumen = Format(RsArticulo!ArtVolumen, "##,##0.00")
    lModificacion.Caption = Format(RsArticulo!ArtModificado, "Ddd d-Mmm yyyy hh:mm")
    
    If Not IsNull(RsArticulo!ArtAlta) Then lFalta.Caption = "Fecha alta: " & Format(RsArticulo!ArtAlta, "Ddd d-Mmm yyyy hh:mm")
    If Not IsNull(RsArticulo!ArtUsuAlta) Then
        lUsuAlta.Caption = "Alta: " & BuscoUsuario(RsArticulo!ArtUsuAlta, True)
    End If
    If Not IsNull(RsArticulo!ArtUsuModificacion) Then
        lUsuMod.Caption = "Modificado por: " & BuscoUsuario(RsArticulo!ArtUsuModificacion, True)
    End If
    
    'Codigo de Barras y Local Retira x Defecto
    If Not IsNull(RsArticulo!ArtBarCode) Then tBarCode.Text = Trim(RsArticulo!ArtBarCode)
    If Not IsNull(RsArticulo!ArtLocalRetira) Then BuscoCodigoEnCombo cLocal, RsArticulo!ArtLocalRetira
    If RsArticulo!ArtAMercaderia Then cAMercaderia.Value = 1 Else cAMercaderia.Value = vbUnchecked
    If RsArticulo!ArtEnWeb Then cArtEnWeb.Value = 1 Else cArtEnWeb.Value = vbUnchecked
    
    If RsArticulo!ArtNroSerie Then cNroSerie.Value = vbChecked Else cNroSerie.Value = vbUnchecked
    
    cImportada.Value = 0
    If RsArticulo!ArtSeImporta Then cImportada.Value = 1
    
    'CARGO DATOS DE TABLA ArticuloFacturacion -------------------
    Cons = "Select * from ArticuloFacturacion Where AFaArticulo = " & RsArticulo!ArtID
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!AFaGarantia) Then BuscoCodigoEnCombo cGarantia, RsAux!AFaGarantia
        
        BuscoCodigoEnCombo cIVA, RsAux!AFaIva
        
        If RsAux!AFaExclusivo Then cExclusivo.Value = 1
        If RsAux!AFaInterior Then cInterior.Value = 1
        'If RsAux!AFaComentario Then cComentario.Value = 1
        
        If Not IsNull(RsAux!AFaFrente) Then tFrente.Text = Format(RsAux!AFaFrente, "#,##0.00")
        If Not IsNull(RsAux!AFaAlto) Then tAlto.Text = Format(RsAux!AFaAlto, "#,##0.00")
        If Not IsNull(RsAux!AFaProfundidad) Then tProfundidad.Text = Format(RsAux!AFaProfundidad, "#,##0.00")
        If Not IsNull(RsAux!AFaPeso) Then tPeso.Text = Format(RsAux!AFaPeso, "#,##0.00")
        
        If Not IsNull(RsAux!AFaCategoriaD) Then BuscoCodigoEnCombo cCategoria, RsAux!AFaCategoriaD
        If Not IsNull(RsAux!AFaCantidadD) Then tCompra.Text = RsAux!AFaCantidadD
        If Not IsNull(RsAux!AFaComentarioA) Then tComentarioF.Text = Trim(RsAux!AFaComentarioA)
        
        If Not IsNull(RsAux!AFaLista) Then BuscoCodigoEnCombo cLista, RsAux!AFaLista
        If Not IsNull(RsAux!AFaEtNormales) Then tENormal.Text = RsAux!AFaEtNormales
        If Not IsNull(RsAux!AFaEtVidriera) Then tEVidriera.Text = RsAux!AFaEtVidriera
        
        If Not IsNull(RsAux!AFaArgumCorto) Then tACorto.Text = Trim(RsAux!AFaArgumCorto)
        If Not IsNull(RsAux!AFaArgumLargo) Then tALargo.Text = Trim(RsAux!AFaArgumLargo)
        
    End If
    RsAux.Close
    
    'Cargo Datos de la tabla ArticuloGrupo---------------------------------------
    lGrupo.ListItems.Clear
    Cons = "Select * from ArticuloGrupo, Grupo" _
            & " Where AGrArticulo = " & RsArticulo!ArtID _
            & " And GruCodigo = AGrGrupo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        Set itmx = lGrupo.ListItems.Add(, "A" & Str(RsAux!AGrGrupo), Trim(RsAux!GruNombre))
        RsAux.MoveNext
    Loop
    RsAux.Close
        
    'Cargo Datos de la tabla ArticuloImportacion---------------------------------------
    If RsArticulo!ArtSeImporta Then
        Cons = "Select * from ArticuloImportacion Where AImArticulo = " & RsArticulo!ArtID
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            
            If Not IsNull(RsAux!AImCantContenedor) Then tCantContenedor.Text = RsAux!AImCantContenedor
            If Not IsNull(RsAux!AImDemoraFabricacion) Then tDemora.Text = RsAux!AImDemoraFabricacion
            If Not IsNull(RsAux!AImComentario) Then tComentarioI.Text = Trim(RsAux!AImComentario)
            If Not IsNull(RsAux!AImProveedor) Then loc_FindProvExterior RsAux!AImProveedor
            If Not IsNull(RsAux!AImContenedor) Then BuscoCodigoEnCombo cContenedor, RsAux!AImContenedor
            If Not IsNull(RsAux!AImPuertoEmbarque) Then BuscoCodigoEnCombo cPuerto, RsAux!AImPuertoEmbarque
            If Not IsNull(RsAux!AImCodigoFabrica) Then tCodigoFabrica.Text = Trim(RsAux!AImCodigoFabrica)
        End If
        RsAux.Close
    End If
    
    'CARGO DATOS DE TABLA Wee -------------------
    Cons = "Select * from ArticuloWebPAge Where AWPArticulo = " & RsArticulo!ArtID
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux!AWPSinPrecio Then cWebSPrecio.Value = vbChecked
    End If
    RsAux.Close
    '--------------------------------------------------------------------------------------------
    
    vsRelacionadoCon.Rows = 0
    Dim idAR As Long
    Cons = "Select * from RelacionArticulo, Articulo Where RArArticulo = " & RsArticulo!ArtID _
        & " And RArRelacionadoCon = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsRelacionadoCon
            If Not IsNull(RsAux!ArtCodigo) Then
                .AddItem "(" & Format(RsAux!ArtCodigo) & ") " & Trim(RsAux!ArtNombre)
            Else
                .AddItem Trim(RsAux!ArtNombre)
            End If
            idAR = RsAux!ArtID
            .Cell(flexcpData, .Rows - 1, 0) = idAR
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al intentar cargar los datos del artículo.", Err.Description
    Screen.MousePointer = vbDefault
End Sub

Sub BuscoArticulo(Codigo As Long)

    On Error GoTo errBuscar
    
    Screen.MousePointer = 11
    RsArticulo.Close
    Cons = "Select * from Articulo Where ArtId = " & Codigo
    Set RsArticulo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    LimpioCampos
        
    If RsArticulo.EOF Then
        Screen.MousePointer = 0
        MsgBox "El artículo seleccionado no existe o ha sido eliminado.", vbExclamation, "ATENCIÓN"
        Call Botones(True, False, False, False, False, Toolbar1, Me)
        MnuPrecio.Enabled = False

    Else
        Call Botones(True, True, True, False, False, Toolbar1, Me)
        MnuPrecio.Enabled = True

        CargoDatosArticulo
    End If
    Screen.MousePointer = 0
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Ha ocurrido un error al buscar el artículo.", Err.Description
    Screen.MousePointer = vbDefault
End Sub

Function ValidoCampos()

    ValidoCampos = False
    
    'FICHA ARTICULO ------------------------------------------------------------------------
    If Trim(tCodigo) = "" Or Not IsNumeric(tCodigo) Then
        MsgBox "El código ingresado no es correcto.", vbExclamation, "ATENCION"
        Foco tCodigo
        Exit Function
    End If
    
    If cTipo.ListIndex = -1 Or cMarca.ListIndex = -1 Or Trim(tNombre.Text) = "" Or Trim(tDescripcion.Text) = "" Or Val(tProveedor.Tag) <= 0 Then
        MsgBox "La ficha del artículo está incompleta. Ingrese los datos obligatorios.", vbExclamation, "ATENCION"
        Foco cTipo
        Exit Function
    End If
    
    If Not clsGeneral.TextoValido(tNombre.Text) Or Not clsGeneral.TextoValido(tDescripcion.Text) Then
        MsgBox "Se ingresaron caracteres no válidos, verifique.", vbExclamation, "ATENCION"
        Foco tNombre
        Exit Function
    End If
    
    If Trim(cGarantia.Text) <> "" And cGarantia.ListIndex = -1 Then
        MsgBox "La garantía ingresada no es correcta.", vbExclamation, "ATENCION"
        Foco cGarantia
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------
    
    'FICHA ARTICULO -----------------------------------------------------------------------------
    If cIVA.ListIndex = -1 Then
        MsgBox "La ficha del los datos de venta está incompleta. Ingrese los datos obligatorios.", vbExclamation, "ATENCION"
        Foco cIVA
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------
    
    'FICHA DIMENSIONES ------------------------------------------------------------------------
    If (Trim(tAlto.Text) <> "" And Not IsNumeric(tAlto.Text)) Or (Trim(tPeso.Text) <> "" And Not IsNumeric(tPeso.Text)) Or _
    (Trim(tProfundidad.Text) <> "" And Not IsNumeric(tProfundidad.Text)) Or (Trim(tVolumen.Text) <> "" And Not IsNumeric(tVolumen.Text)) Or _
    (Trim(tFrente.Text) <> "" And Not IsNumeric(tFrente.Text)) Then
        MsgBox "Los datos ingresados en la ficha dimensiones no son correctos.", vbExclamation, "ATENCION"
        Foco tFrente
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------
    
    'FICHA DESCUENTOS -------------------------------------------------------------------------
    If (Trim(cCategoria.Text) <> "" And cCategoria.ListIndex = -1) Or (Trim(tCompra.Text) <> "" And Not IsNumeric(tCompra.Text)) Then
        MsgBox "Los datos ingresados en la ficha descuentos no son correctos.", vbExclamation, "ATENCION"
        cCategoria.SetFocus
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------
        
    'FICHA IMPORTACIONES----------------------------------------------------------------------
    If (Trim(tCantContenedor.Text) <> "" And Not IsNumeric(tCantContenedor)) Or (Trim(tDemora.Text) <> "" And Not IsNumeric(tDemora.Text)) Then
        MsgBox "Los datos ingresados en la ficha de importación no son correctos.", vbExclamation, "ATENCION"
        cImportada.SetFocus
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------
    
    'Valido que no exista otro articulo con el mismo codigo de barras
    If Trim(tBarCode.Text) <> "" Then
        Cons = "Select * from Articulo Where ArtBarCode = '" & Trim(tBarCode.Text) & "'"
        If sModificar Then Cons = Cons & " And ArtId <> " & RsArticulo!ArtID
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            MsgBox "Existen artículo para el código de barras ingresado. Verifique.", vbExclamation, "ATENCIÓN"
            RsAux.Close
            Exit Function
        End If
        RsAux.Close
    End If
    '-------------------------------------------------------------------------------------------------
    
    If miConexion.UsuarioLogueado(True) = 0 Then
        miConexion.AccesoAlMenu (App.Title)
        If miConexion.UsuarioLogueado(True) = 0 Then
            MsgBox "Para grabar debe loguearse al sistema.", vbCritical, "ATENCIÓN"
            Exit Function
        End If
    End If
    
    If sNuevo Then
        If cTipo.ItemData(cTipo.ListIndex) = paTipoArticuloServicio Then
            MsgBox "El tipo de artículo seleccionado NO AFECTARA EL STOCK.", vbExclamation, "ATENCIÓN"
        End If
    ElseIf RsArticulo!ArtTipo <> cTipo.ItemData(cTipo.ListIndex) Then
        'Verifico si alguno de los dos es del tipo servicio.
        If RsArticulo!ArtTipo = paTipoArticuloServicio Then
            MsgBox "Ud. modificó el tipo de artículo, el nuevo tipo AFECTA EL STOCK.", vbInformation, "ATENCIÓN"
        ElseIf cTipo.ItemData(cTipo.ListIndex) = paTipoArticuloServicio Then
            MsgBox "Ud. modificó el tipo de artículo, el nuevo tipo DEJARÁ DE AFECTAR EL STOCK.", vbInformation, "ATENCIÓN"
        End If
    End If
    ValidoCampos = True
    
End Function

Private Sub CargoCamposBDArticulo()
    
    RsArticulo!ArtCodigo = CLng(tCodigo)
    RsArticulo!ArtTipo = cTipo.ItemData(cTipo.ListIndex)
    RsArticulo!ArtMarca = cMarca.ItemData(cMarca.ListIndex)
    RsArticulo!ArtProveedor = tProveedor.Tag
    
    RsArticulo!ArtNombre = Trim(tNombre.Text)
    RsArticulo!ArtDescripcion = Trim(tDescripcion.Text)
    
    If Trim(tVolumen.Text) <> "" Then RsArticulo!ArtVolumen = CCur(tVolumen.Text) Else: RsArticulo!ArtVolumen = Null
    If cEnUso.Value = 0 Then RsArticulo!ArtEnUso = 0 Else RsArticulo!ArtEnUso = 1
    If cHabilitado.Value Then RsArticulo!ArtHabilitado = "S" Else: RsArticulo!ArtHabilitado = Null
    If cAMercaderia.Value = 1 Then RsArticulo!ArtAMercaderia = 1 Else RsArticulo!ArtAMercaderia = 0
    
    If cArtEnWeb.Visible Then
        If cArtEnWeb.Value = 1 Then RsArticulo!ArtEnWeb = 1 Else RsArticulo!ArtEnWeb = 0
    End If
    
    If cImportada.Value = 0 Then RsArticulo!ArtSeImporta = False Else: RsArticulo!ArtSeImporta = True
    
    If cLocal.ListIndex <> -1 Then RsArticulo!ArtLocalRetira = cLocal.ItemData(cLocal.ListIndex) Else: RsArticulo!ArtLocalRetira = Null
    If Trim(tBarCode.Text) <> "" Then RsArticulo!ArtBarCode = Trim(tBarCode.Text) Else: RsArticulo!ArtBarCode = Null
    
    If cNroSerie.Value = vbChecked Then RsArticulo!ArtNroSerie = 1 Else RsArticulo!ArtNroSerie = 0
    
    RsArticulo!ArtUsuModificacion = miConexion.UsuarioLogueado(True)
    RsArticulo!ArtModificado = Format(Now, sqlFormatoFH)
    
End Sub

Private Sub CargoCamposBDArticuloFacturacion(Articulo As Long)
    
    Cons = "Select * from ArticuloFacturacion Where AFaArticulo = " & Articulo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then RsAux.AddNew Else RsAux.Edit
    
    RsAux!AFaArticulo = Articulo
    
    If cGarantia.ListIndex <> -1 Then RsAux!AFaGarantia = cGarantia.ItemData(cGarantia.ListIndex) Else RsAux!AFaGarantia = Null
    
    RsAux!AFaIva = cIVA.ItemData(cIVA.ListIndex)
    
    If cExclusivo.Value = 1 Then RsAux!AFaExclusivo = True Else RsAux!AFaExclusivo = False
    If cInterior.Value = 1 Then RsAux!AFaInterior = True Else RsAux!AFaInterior = False
    'If cComentario.Value = 1 Then RsAux!AFaComentario = True Else RsAux!AFaComentario = False
    RsAux!AFaComentario = 0
    
    If Trim(tFrente.Text) <> "" Then RsAux!AFaFrente = CCur(tFrente.Text) Else RsAux!AFaFrente = Null
    If Trim(tAlto.Text) <> "" Then RsAux!AFaAlto = CCur(tAlto.Text) Else RsAux!AFaAlto = Null
    If Trim(tProfundidad.Text) <> "" Then RsAux!AFaProfundidad = CCur(tProfundidad.Text) Else RsAux!AFaProfundidad = Null
    If Trim(tPeso.Text) <> "" Then RsAux!AFaPeso = CCur(tPeso.Text) Else RsAux!AFaPeso = Null
    
    If cCategoria.ListIndex <> -1 Then RsAux!AFaCategoriaD = cCategoria.ItemData(cCategoria.ListIndex) Else RsAux!AFaCategoriaD = Null
    If Trim(tCompra.Text) <> "" Then RsAux!AFaCantidadD = CLng(tCompra.Text) Else RsAux!AFaCantidadD = Null
    If Trim(tComentarioF.Text) <> "" Then RsAux!AFaComentarioA = Trim(tComentarioF.Text) Else RsAux!AFaComentarioA = Null
    
    If cLista.ListIndex <> -1 Then RsAux!AFaLista = cLista.ItemData(cLista.ListIndex) Else RsAux!AFaLista = Null
    If Trim(tENormal.Text) <> "" Then RsAux!AFaEtNormales = CInt(tENormal.Text) Else RsAux!AFaEtNormales = Null
    If Trim(tEVidriera.Text) <> "" Then RsAux!AFaEtVidriera = CInt(tEVidriera.Text) Else RsAux!AFaEtVidriera = Null
    
    If Trim(tACorto.Text) <> "" Then RsAux!AFaArgumCorto = Trim(tACorto.Text) Else RsAux!AFaArgumCorto = Null
    If Trim(tALargo.Text) <> "" Then RsAux!AFaArgumLargo = Trim(tALargo.Text) Else RsAux!AFaArgumLargo = Null
    
    RsAux.Update
    RsAux.Close
    
End Sub

Private Sub CargoCamposBDArticuloImportacion(Codigo As Long)

    Cons = "Select * from ArticuloImportacion Where AImArticulo = " & Codigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.AddNew
    Else
        RsAux.Edit
    End If
    
    RsAux!AImArticulo = Codigo
    If Val(tProvExt.Tag) > 0 Then
        RsAux!AImProveedor = Val(tProvExt.Tag)
    Else
        RsAux!AImProveedor = Null
    End If
    
    If cContenedor.ListIndex <> -1 Then
        RsAux!AImContenedor = cContenedor.ItemData(cContenedor.ListIndex)
    Else
        RsAux!AImContenedor = Null
    End If
    
    If Trim(tCantContenedor) <> "" Then
        RsAux!AImCantContenedor = CInt(tCantContenedor.Text)
    Else
        RsAux!AImCantContenedor = Null
    End If
    
    If cPuerto.ListIndex <> -1 Then
        RsAux!AImPuertoEmbarque = cPuerto.ItemData(cPuerto.ListIndex)
    Else
        RsAux!AImPuertoEmbarque = Null
    End If
    
    If Trim(tComentarioI.Text) <> "" Then
        RsAux!AImComentario = Trim(tComentarioI.Text)
    Else
        RsAux!AImComentario = Null
    End If
    
    If Trim(tDemora.Text) <> "" Then
        RsAux!AImDemoraFabricacion = CInt(tDemora.Text)
    Else
        RsAux!AImDemoraFabricacion = Null
    End If
    
    If Trim(tCodigoFabrica.Text) <> "" Then RsAux!AImCodigoFabrica = tCodigoFabrica.Text Else RsAux!AImCodigoFabrica = Null
    
    RsAux.Update
    RsAux.Close
    
End Sub



Private Sub tPeso_GotFocus()

    tPeso.SelStart = 0
    tPeso.SelLength = Len(tPeso.Text)
    
    Status.Panels(1).Text = "Ingrese el peso del artículo (en kilogramos)."
    
End Sub
Private Sub tPeso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tVolumen
End Sub
Private Sub tPeso_LostFocus()
    tPeso.Text = Format(tPeso.Text, FormatoMonedaP)
End Sub

Private Sub tProfundidad_GotFocus()

    tProfundidad.SelStart = 0
    tProfundidad.SelLength = Len(tProfundidad.Text)
    
    Status.Panels(1).Text = "Ingrese la profundidad (en centímetros)."
    
End Sub

Private Sub tProfundidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Foco tPeso
    
End Sub

Private Sub tProfundidad_LostFocus()

    tProfundidad.Text = Format(tProfundidad.Text, FormatoMonedaP)
    
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = ""
End Sub

Private Sub tProveedor_GotFocus()
    With tProveedor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) > 0 Or RTrim(tProveedor.Text) = "" Then
            Foco cGarantia
        Else
            loc_FindProveedor 0, tProveedor.Text, True
            If Val(tProveedor.Tag) > 0 Then Foco cGarantia
        End If
    End If
End Sub

Private Sub tProvExt_Change()
    tProvExt.Tag = ""
End Sub

Private Sub tProvExt_GotFocus()
    With tProvExt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tProvExt_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Val(tProvExt.Tag) > 0 Or RTrim(tProvExt.Text) = "" Then
            Foco cContenedor
        Else
            loc_FindProvExterior 0, tProvExt.Text, True
            If Val(tProvExt.Tag) > 0 Then Foco cContenedor
        End If
    End If
End Sub

Private Sub tVolumen_GotFocus()

    tVolumen.SelStart = 0
    tVolumen.SelLength = Len(tVolumen)
    
    Status.Panels(1).Text = "Ingrese el volumen del artículo (en litros)."
    
End Sub

Private Sub tVolumen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cGrupo
End Sub
Private Sub tVolumen_LostFocus()
    tVolumen.Text = Format(tVolumen.Text, FormatoMonedaP)
End Sub

Private Sub AyudaArticulo(Consulta As String)
On Error GoTo ErrIA
    Screen.MousePointer = 11
    Dim objAyuda As New clsListadeAyuda
    sAyuda = True
    If objAyuda.ActivarAyuda(cBase, Consulta, 3000, 1, "Ayuda de Artículo") > 0 Then
        sAyuda = False
        BuscoArticulo objAyuda.RetornoDatoSeleccionado(0)
    Else
        sAyuda = False
        Botones True, False, False, False, False, Toolbar1, Me
        MnuPrecio.Enabled = False
    End If
    Set objAyuda = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrIA:
    Screen.MousePointer = 0
    MnuPrecio.Enabled = False
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos."
    Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub AccionRefrescar()

    On Error GoTo errRefrescar
    
    Screen.MousePointer = 11
    
    'Cargo los TIPOS DE ARTICULOS
    Cons = "Select TipCodigo, TipNombre From Tipo Order by TipNombre"
    CargoCombo Cons, cTipo, cTipo.Text
    
    'Cargo las MARCAS
    Cons = "Select MarCodigo, MarNombre From Marca Order by MarNombre"
    CargoCombo Cons, cMarca, cMarca.Text
    
    'Cargo os TIPOS DE IVA
    Cons = "Select IvaCodigo, IvaAbreviacion From TipoIva Order by IvaAbreviacion"
    CargoCombo Cons, cIVA, cIVA.Text
    
     'Cargo las Categorias de Descuento (Articulos)
    Cons = "Select CArCodigo, CArNombre From CategoriaArticulo Order by CArNombre"
    CargoCombo Cons, cCategoria, cCategoria.Text
    
    'Cargo las GARANTIAS
    Cons = "Select GarCodigo, GarNombre From Garantia Order by GarNombre"
    CargoCombo Cons, cGarantia, cGarantia.Text
    
    RsArticulo.Requery
    
    If Not RsArticulo.EOF And sModificar Then
        LimpioCampos
        CargoDatosArticulo
    Else
        If Not sNuevo Then
            Botones True, False, False, False, False, Toolbar1, Me
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errRefrescar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al refrescar los datos.", Err.Description
End Sub

Private Sub DeshabilitoIngreso()

    'Datos Articulo----------------------------------------
    tDescripcion.BackColor = Inactivo
    cGarantia.BackColor = Inactivo
    tProveedor.BackColor = Inactivo
    tDescripcion.Enabled = False
    cGarantia.Enabled = False
    tProveedor.Enabled = False
    
    tVolumen.BackColor = Inactivo
    tAlto.BackColor = Inactivo
    tProfundidad.BackColor = Inactivo
    tPeso.BackColor = Inactivo
    tFrente.BackColor = Inactivo
    tVolumen.Enabled = False
    tAlto.Enabled = False
    tProfundidad.Enabled = False
    tPeso.Enabled = False
    tFrente.Enabled = False
    
    cGrupo.Enabled = False
    cGrupo.BackColor = Inactivo
    lGrupo.BackColor = Inactivo
    
    'Ficha de Facturacion----------------------------------
    cIVA.BackColor = Inactivo
    cIVA.Enabled = False
    cEnUso.Enabled = False
    cHabilitado.Enabled = False
    cExclusivo.Enabled = False
    cInterior.Enabled = False
    cArtEnWeb.Enabled = False
    cWebSPrecio.Enabled = False
    cAMercaderia.Enabled = False
    
    cNroSerie.Enabled = False
    cCategoria.BackColor = Inactivo
    tCompra.BackColor = Inactivo
    tComentarioF.BackColor = Inactivo
    tBarCode.BackColor = Inactivo
    cLocal.BackColor = Inactivo
    cCategoria.Enabled = False
    tCompra.Enabled = False
    tComentarioF.Enabled = False
    tBarCode.Enabled = False
    cLocal.Enabled = False
    
    cLista.Enabled = False
    cLista.BackColor = Inactivo
    tENormal.Enabled = False
    tENormal.BackColor = Inactivo
    tEVidriera.Enabled = False
    tEVidriera.BackColor = Inactivo
    tACorto.Enabled = False
    tACorto.BackColor = Inactivo
    tALargo.BackColor = Inactivo
    
    'Ficha de Importacion---------------------------------------
    cImportada.Value = 0
    cImportada.Enabled = False
    cContenedor.Enabled = False
    tProvExt.Enabled = False
    cPuerto.Enabled = False
    tDemora.Enabled = False
    tCodigoFabrica.Enabled = False: tCodigoFabrica.BackColor = Inactivo
    tComentarioI.Enabled = False
    tCantContenedor.Enabled = False
    cContenedor.BackColor = Inactivo
    tProvExt.BackColor = Inactivo
    cPuerto.BackColor = Inactivo
    tDemora.BackColor = Inactivo
    tComentarioI.BackColor = Inactivo
    tCantContenedor.BackColor = Inactivo
    
    vsRelacionadoCon.Enabled = False: vsRelacionadoCon.BackColor = Inactivo
    tArtRelacionadoCon.Enabled = False: tArtRelacionadoCon.BackColor = Inactivo
    
End Sub

Sub HabilitoIngreso()
    
    'Ficha Artículo------------------------------------------
    tDescripcion.BackColor = Obligatorio: tDescripcion.Enabled = True
    cGarantia.BackColor = Blanco: cGarantia.Enabled = True
    tProveedor.BackColor = Obligatorio: tProveedor.Enabled = True
    
    tVolumen.BackColor = Blanco: tVolumen.Enabled = True
    tAlto.BackColor = Blanco: tAlto.Enabled = True
    tProfundidad.BackColor = Blanco: tProfundidad.Enabled = True
    tPeso.BackColor = Blanco: tPeso.Enabled = True
    tFrente.BackColor = Blanco: tFrente.Enabled = True
    
    cGrupo.Enabled = True: cGrupo.BackColor = Blanco
    lGrupo.BackColor = Blanco
    
    'Ficha Facturacion--------------------------------------
    cIVA.BackColor = Obligatorio: cIVA.Enabled = True
    cEnUso.Enabled = True
    cHabilitado.Enabled = True
    cExclusivo.Enabled = True
    cInterior.Enabled = True
    cArtEnWeb.Enabled = True: cWebSPrecio.Enabled = True
    'cComentario.Enabled = True
    cAMercaderia.Enabled = True
    
    cNroSerie.Enabled = True
    
    cCategoria.BackColor = Blanco: cCategoria.Enabled = True
    tCompra.BackColor = Blanco: tCompra.Enabled = True
    tComentarioF.BackColor = Blanco: tComentarioF.Enabled = True
    tBarCode.BackColor = Blanco: tBarCode.Enabled = True
    cLocal.BackColor = Blanco: cLocal.Enabled = True

    cLista.Enabled = True: cLista.BackColor = Blanco
    tENormal.Enabled = True: tENormal.BackColor = Blanco
    tEVidriera.Enabled = True: tEVidriera.BackColor = Blanco
    tACorto.Enabled = True: tACorto.BackColor = Blanco
    tALargo.BackColor = Blanco
    
    vsRelacionadoCon.Enabled = True: vsRelacionadoCon.BackColor = Blanco
    tArtRelacionadoCon.Enabled = True: tArtRelacionadoCon.BackColor = Blanco
    
    'Ficha de Importacion-------------------------
    cImportada.Enabled = True
    
    
End Sub

Private Sub CargoCamposBDGrupos(Codigo As Long)

    Cons = "Delete ArticuloGrupo Where AGrArticulo = " & Codigo
    cBase.Execute (Cons)
    
    For I = 1 To lGrupo.ListItems.Count
        Cons = "INSERT INTO ArticuloGrupo (AGrArticulo, AGrGrupo)" _
            & " Values(" & Codigo & ", " & Right(lGrupo.ListItems(I).Key, Len(lGrupo.ListItems(I).Key) - 1) & ")"
        cBase.Execute (Cons)
    Next I

End Sub

Private Sub cLocal_GotFocus()
    cLocal.SelStart = 0
    cLocal.SelLength = Len(cLocal.Text)
    Status.Panels(1).Text = "Seleccione local dónde se retira el artículo."
End Sub
Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cEnUso.Enabled Then cCategoria.SetFocus
End Sub
Private Sub cLocal_LostFocus()
    Status.Panels(1).Text = ""
    cLocal.SelLength = 0
End Sub

Private Sub AccederAMarca(Tipo As Integer)
On Error GoTo ErrAcceso
    
    If Not miConexion.AccesoAlMenu("Mantenimiento de Marcas") Then
        MsgBox "Ud no tiene permiso para acceder al formulario de Marcas de Artículos.", vbCritical, "ATENCIÓN"
        Exit Sub
    End If
    
    Set frm1Campo = New MaUnCampo
    frm1Campo.pSeleccionado = 0
    frm1Campo.pCampoCodigo = "MarCodigo"
    frm1Campo.pCampoNombre = "MarNombre"
    frm1Campo.pTipoLlamado = Tipo
    frm1Campo.pTabla = "MARCA"
    frm1Campo.Caption = "Marcas"
    frm1Campo.tDescripcion.Text = Trim(cMarca.Text)
    If Tipo = TipoLlamado.IngresoNuevo Then
        frm1Campo.Show vbModal, Me
    Else
        frm1Campo.Show vbModeless, Me
    End If
    
    Dim Sele As Long
    Sele = frm1Campo.pSeleccionado
    'Cargo las Marcas de Articulos
    Cons = "Select MarCodigo, MarNombre From Marca Order by MarNombre"
    CargoCombo Cons, cMarca, ""

    If frm1Campo.pSeleccionado <> 0 Then BuscoCodigoEnCombo cMarca, Sele
    
    Set frm1Campo = Nothing
    cMarca.ListIndex = cMarca.ListIndex
    Me.Refresh
    
    Exit Sub

ErrAcceso:
    clsGeneral.OcurrioError "Ocurrio un error al acceder al formulario de Marcas.", Err.Description
End Sub

Private Function ExisteEnCosteo(idArticulo As Long) As Boolean
Dim RsCM As rdoResultset
On Error GoTo ErrEC
    
    ExisteEnCosteo = True
    
    Cons = "Select * From CMCompra Where CoMArticulo = " & idArticulo
    Set RsCM = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsCM.EOF Then
        RsCM.Close: MsgBox "El artículo esta en las tablas de existencias del costeo, no podrá eliminarlo.", vbCritical, "ATENCIÓN"
        Exit Function
    End If
    RsCM.Close
    
    Cons = "Select * From CMVenta Where VenArticulo = " & idArticulo
    Set RsCM = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsCM.EOF Then
        RsCM.Close: MsgBox "El artículo esta en la tabla de ventas rebotadas del costeo, no podrá eliminarlo.", vbCritical, "ATENCIÓN"
        Exit Function
    End If
    RsCM.Close
    
    Cons = "Select * From CMCosteo Where CosArticulo = " & idArticulo
    Set RsCM = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsCM.EOF Then
        RsCM.Close: MsgBox "El artículo esta en la tabla COSTEO, no podrá eliminarlo.", vbCritical, "ATENCIÓN"
        Exit Function
    End If
    RsCM.Close
    
    ExisteEnCosteo = False
    Exit Function
ErrEC:
    clsGeneral.OcurrioError "Ocurrió un error al validar información de costeo.", Trim(Err.Description)
End Function

Private Sub CargoCamposWeb(Articulo As Long)
Dim rsWeb As rdoResultset

    Cons = "Select * from ArticuloWebPage Where AWPArticulo = " & Articulo
    Set rsWeb = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsWeb.EOF Then
        rsWeb.Edit
        If cWebSPrecio.Value = vbChecked Then rsWeb!AWPSinPrecio = 1 Else rsWeb!AWPSinPrecio = 0
        rsWeb.Update
    Else
        If cWebSPrecio.Value = vbChecked Then
            rsWeb.AddNew
            rsWeb!AWPArticulo = Articulo
            rsWeb!AWPSinPrecio = 1
            rsWeb.Update
        End If
    End If
    rsWeb.Close
    
End Sub
Private Sub AccionDatosWeb()
On Error GoTo errADW
    If RsArticulo.EOF Then
        EjecutarApp App.Path & "\datos_web.exe"
    Else
        EjecutarApp App.Path & "\datos_web.exe", RsArticulo!ArtID
    End If
    Exit Sub
errADW:
    Screen.MousePointer = 0
    MsgBox "Ocurrió el siguiente error al intentar acceder a datos web: " & Err.Description, vbExclamation, "ATENCIÓN"
End Sub

Private Sub BuscoArticuloRelacionado(mControl As TextBox)

    On Error GoTo errBuscaG
    Screen.MousePointer = 11
    
    Cons = "Select ArtID, ArtCodigo as Codigo, ArtNombre  as Nombre from Articulo "
    If IsNumeric(mControl.Text) Then
        Cons = Cons & " Where ArtCodigo = " & Val(mControl.Text)
    Else
        Cons = Cons & "Where ArtNombre like '" & Replace(Trim(mControl.Text), " ", "%") & "%'"
    End If
    Cons = Cons & " Order by ArtNombre"
    
    Dim aQ As Integer, aIDArticulo As Long, aTexto As String
    aQ = 0: aIDArticulo = 0
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        aQ = 1
        aIDArticulo = RsAux!ArtID: aTexto = Format(RsAux!Codigo, "(#,000,000)") & " " & Trim(RsAux!Nombre)
        RsAux.MoveNext: If Not RsAux.EOF Then aQ = 2
    End If
    RsAux.Close
        
    Select Case aQ
        Case 0: MsgBox "No hay datos que coincidan con el texto ingersado.", vbExclamation, "No hay datos"
        
        Case 2:
                Dim miLista As New clsListadeAyuda
                aIDArticulo = miLista.ActivarAyuda(cBase, Cons, 4000, 1, "Lista de Articulos")
                Me.Refresh
                If aIDArticulo > 0 Then
                    aIDArticulo = miLista.RetornoDatoSeleccionado(0)
                    
                    aTexto = Format(miLista.RetornoDatoSeleccionado(1), "(#,000,000)") & " "
                    aTexto = aTexto & miLista.RetornoDatoSeleccionado(2)
                End If
                Set miLista = Nothing
    End Select
        
    If aIDArticulo > 0 Then
        mControl.Text = aTexto
        mControl.Tag = aIDArticulo
        Foco mControl
    End If
    
    Screen.MousePointer = 0
   
    Exit Sub
errBuscaG:
    clsGeneral.OcurrioError "Error al buscar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub InsertoArtRelacionado()
Dim iCont As Integer
    For iCont = 0 To vsRelacionadoCon.Rows - 1
        If Val(vsRelacionadoCon.Cell(flexcpData, iCont, 0)) = Val(tArtRelacionadoCon.Tag) Then
            MsgBox "Ese artículo ya esta ingresado.", vbInformation, "ATENCIÓN"
            Exit Sub
        End If
    Next
    vsRelacionadoCon.AddItem tArtRelacionadoCon.Text
    vsRelacionadoCon.Cell(flexcpData, vsRelacionadoCon.Rows - 1, 0) = tArtRelacionadoCon.Tag
End Sub

Private Sub GuardoRelacionadoCon(ByVal idArtR As Long)
Dim iCont As Integer
Dim rsARC As rdoResultset

    Cons = "Delete RelacionArticulo Where RArArticulo = " & idArtR
    cBase.Execute (Cons)
    Cons = "Select * From RelacionArticulo Where RArArticulo = " & idArtR
    Set rsARC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    For iCont = 0 To vsRelacionadoCon.Rows - 1
        rsARC.AddNew
        rsARC!RArArticulo = idArtR
        rsARC!RArRelacionadoCon = Val(vsRelacionadoCon.Cell(flexcpData, iCont, 0))
        rsARC.Update
    Next iCont
    rsARC.Close
    
End Sub

Private Sub vsRelacionadoCon_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsRelacionadoCon.Rows = 0 Then Exit Sub
    If KeyCode = vbKeyDelete Then
        vsRelacionadoCon.RemoveItem vsRelacionadoCon.Row
    End If
End Sub

Private Sub InicializoCombos()

    'Cargo los TIPOS DE ARTICULOS
    Cons = " Select '1' as Tipo, TipCodigo as Codigo, TipNombre as Nombre From Tipo " & " UNION ALL " & _
               " Select '5' as Tipo, GruCodigo as Codigo, GruNombre as Nombre From Grupo " & " UNION ALL " & _
               " Select '6' as Tipo, IvaCodigo as Codigo, IvaAbreviacion as Nombre From TipoIva " & " UNION ALL " & _
               " Select '7' as Tipo, CArCodigo as Codigo, CArNombre as Nombre From CategoriaArticulo " & " UNION ALL " & _
               " Select '8' as Tipo, GarCodigo as Codigo, GarNombre collate SQL_Latin1_General_Cp1251_CS_AS as Nombre From Garantia" & " UNION ALL " & _
               " Select '9' as Tipo, SucCodigo as Codigo, SucAbreviacion collate SQL_Latin1_General_Cp1251_CS_AS as Nombre From Sucursal " & " UNION ALL " & _
               " Select '10' as Tipo, LDPCodigo as Codigo, LDPDescripcion as Nombre From ListasDePrecios " & " UNION ALL " & _
               " Select '11' as Tipo, ConCodigo as Codigo, ConNombre as Nombre From Contenedor" & " UNION ALL " & _
               " Select '12' as Tipo, CiuCodigo as Codigo, CiuNombre as Nombre From Ciudad" & _
               " Order by Nombre"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case Val(RsAux!Tipo)
            Case 1: cTipo.AddItem Trim(RsAux!Nombre): cTipo.ItemData(cTipo.NewIndex) = RsAux!Codigo
            Case 5: cGrupo.AddItem Trim(RsAux!Nombre): cGrupo.ItemData(cGrupo.NewIndex) = RsAux!Codigo
            Case 6: cIVA.AddItem Trim(RsAux!Nombre): cIVA.ItemData(cIVA.NewIndex) = RsAux!Codigo
            Case 7: cCategoria.AddItem Trim(RsAux!Nombre): cCategoria.ItemData(cCategoria.NewIndex) = RsAux!Codigo
            Case 8: cGarantia.AddItem Trim(RsAux!Nombre): cGarantia.ItemData(cGarantia.NewIndex) = RsAux!Codigo
            Case 9: cLocal.AddItem Trim(RsAux!Nombre): cLocal.ItemData(cLocal.NewIndex) = RsAux!Codigo
            Case 10: cLista.AddItem Trim(RsAux!Nombre): cLista.ItemData(cLista.NewIndex) = RsAux!Codigo
            Case 11: cContenedor.AddItem Trim(RsAux!Nombre): cContenedor.ItemData(cContenedor.NewIndex) = RsAux!Codigo
            Case 12: cPuerto.AddItem Trim(RsAux!Nombre): cPuerto.ItemData(cPuerto.NewIndex) = RsAux!Codigo
        End Select
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    '16-10-06   Cambio x collation
    Cons = " Select MarCodigo as Codigo, MarNombre as Nombre From Marca Order by Nombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        cMarca.AddItem Trim(RsAux!Nombre): cMarca.ItemData(cMarca.NewIndex) = RsAux!Codigo
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    
End Sub

Private Sub loc_FindProvExterior(Optional lID As Long = 0, Optional sName As String = "", Optional bDisplayMsg As Boolean = False)
Dim rsPE As rdoResultset

    Cons = "Select  PExCodigo as Codigo, PExNombre as Nombre From ProveedorExterior "
    If lID > 0 Then
        Cons = Cons & " Where PExCodigo = " & lID
    Else
        Cons = Cons & " Where PExNombre like '" & Replace(sName, " ", "%") & "%'" & _
                    "Order by PExNombre"
    End If

    Set rsPE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsPE.EOF Then
        If lID > 0 Then
            tProvExt.Text = Trim(rsPE("Nombre"))
            tProvExt.Tag = rsPE("Codigo")
            rsPE.Close
        Else
            rsPE.MoveNext
            If rsPE.EOF Then
                rsPE.MoveFirst
                tProvExt.Text = Trim(rsPE("Nombre"))
                tProvExt.Tag = rsPE("Codigo")
                rsPE.Close
            Else
                rsPE.Close
                'Invoco lista de ayuda.
                Dim objHelp As New clsListadeAyuda
                If objHelp.ActivarAyuda(cBase, Cons, 6000, 1, "Proveedores del Exterior") > 0 Then
                    tProvExt.Text = objHelp.RetornoDatoSeleccionado(1)
                    tProvExt.Tag = objHelp.RetornoDatoSeleccionado(0)
                End If
                Set objHelp = Nothing
            End If
        End If
    Else
        rsPE.Close
        If bDisplayMsg Then
            MsgBox "No se encontro un proveedor del exterior para el dato ingresado.", vbExclamation, "ATENCIÓN"
        End If
    End If

End Sub

Private Sub loc_FindProveedor(Optional lID As Long = 0, Optional sName As String = "", Optional bDisplayMsg As Boolean = False)

Dim rsPE As rdoResultset

    Cons = "Select  PMeCodigo as Codigo, PMeNombre as Nombre From ProveedorMercaderia "
    If lID > 0 Then
        Cons = Cons & " Where PMeCodigo = " & lID
    Else
        Cons = Cons & " Where PMeNombre like '" & Replace(sName, " ", "%") & "%'" & _
                    "Order by PMeNombre"
    End If

    Set rsPE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsPE.EOF Then
        If lID > 0 Then
            tProveedor.Text = Trim(rsPE("Nombre"))
            tProveedor.Tag = rsPE("Codigo")
            rsPE.Close
        Else
            rsPE.MoveNext
            If rsPE.EOF Then
                rsPE.MoveFirst
                tProveedor.Text = Trim(rsPE("Nombre"))
                tProveedor.Tag = rsPE("Codigo")
                rsPE.Close
            Else
                rsPE.Close
                'Invoco lista de ayuda.
                Dim objHelp As New clsListadeAyuda
                If objHelp.ActivarAyuda(cBase, Cons, 6000, 1, "Proveedores de Mercadería") > 0 Then
                    tProveedor.Text = objHelp.RetornoDatoSeleccionado(1)
                    tProveedor.Tag = objHelp.RetornoDatoSeleccionado(0)
                End If
                Set objHelp = Nothing
            End If
        End If
    Else
        rsPE.Close
        If bDisplayMsg Then
            MsgBox "No se encontro un proveedor para el dato ingresado.", vbExclamation, "ATENCIÓN"
        End If
    End If

End Sub

