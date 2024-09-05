VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmRecibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos en Dólares"
   ClientHeight    =   4755
   ClientLeft      =   2505
   ClientTop       =   3240
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   9030
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9030
      _ExtentX        =   15928
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
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
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
            Object.Width           =   400
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "pagos"
            Object.ToolTipText     =   "Con Qué paga"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "dolar"
            Object.ToolTipText     =   "Tasas de cambio"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Comprobante"
      ForeColor       =   &H00000080&
      Height          =   915
      Left            =   60
      TabIndex        =   25
      Top             =   480
      Width           =   8895
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5100
         MaxLength       =   40
         TabIndex        =   5
         Top             =   240
         Width           =   3675
      End
      Begin VB.TextBox tID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tTCDolar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5100
         MaxLength       =   6
         TabIndex        =   12
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox tIOriginal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   10
         Text            =   "1,000,000.00"
         Top             =   540
         Width           =   1155
      End
      Begin VB.TextBox tNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1360
         MaxLength       =   9
         TabIndex        =   8
         Text            =   "000000000"
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3120
         MaxLength       =   12
         TabIndex        =   3
         Text            =   "00/00/0000"
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox tSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "BB"
         Top             =   540
         Width           =   320
      End
      Begin VB.Label lPesos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12/12/2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7440
         TabIndex        =   38
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Pesos:"
         Height          =   255
         Left            =   6900
         TabIndex        =   37
         Top             =   585
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "I&d Compra:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
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
         Left            =   5880
         TabIndex        =   27
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T/&C:"
         Height          =   255
         Left            =   4680
         TabIndex        =   11
         Top             =   585
         Width           =   375
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Dó&lares:"
         Height          =   255
         Left            =   2460
         TabIndex        =   9
         Top             =   585
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   4260
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   2460
         TabIndex        =   2
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Número:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   585
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Facturas que se Pagan"
      ForeColor       =   &H00000080&
      Height          =   3000
      Left            =   60
      TabIndex        =   24
      Top             =   1440
      Width           =   8895
      Begin VB.TextBox tFPaga 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3300
         MaxLength       =   15
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton bAgregar 
         Caption         =   "&Agregar"
         Height          =   315
         Left            =   4740
         TabIndex        =   18
         Top             =   225
         Width           =   855
      End
      Begin VB.TextBox tFSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "BB"
         Top             =   240
         Width           =   320
      End
      Begin VB.TextBox tFNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1365
         MaxLength       =   9
         TabIndex        =   15
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   21
         Top             =   2640
         Width           =   7575
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
         Height          =   1725
         Left            =   120
         TabIndex        =   19
         Top             =   900
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3043
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
         FocusRect       =   2
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
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
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T/C:"
         Height          =   255
         Left            =   7080
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lFTC 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12/12/2000"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Pesos:"
         Height          =   255
         Left            =   5100
         TabIndex        =   34
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lFImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,000,000.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5640
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lFSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,000,000.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   31
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pa&ga:"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lFFecha 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12/12/2000"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   30
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lFImporteD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,000,000.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3300
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   7080
         TabIndex        =   28
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fac&tura:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Dólares:"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario&s:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   975
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   4500
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7805
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":0A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":0D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":0E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":11B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":14CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":17E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":1B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":1E1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecibos.frx":2134
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuExit 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean
Dim gIDComprobante As Long
Dim gFModificacion As Date

Dim RsCom As rdoResultset

Dim bIngresarPagos As Boolean                   'Señal para ingresar los pagos

'Parametros de la disponibilidad x defecto para Ingreso de pagos
Dim aMonedaDisponibilidad As Integer
Dim bEsBancaria As Boolean
Dim txtDisponibilidad As String

Private Sub bAgregar_Click()

    If Val(tFNumero.Tag) = 0 Then
        MsgBox "Ingrese los datos la factura a pagar con el recibo.", vbExclamation, "ATENCIÓN"
        Foco tFSerie: Exit Sub
    End If
    
    If Not IsNumeric(tFPaga.Text) Then
        MsgBox "Ingrese el importe a pagar para la factura seleccionada.", vbExclamation, "ATENCIÓN"
        Foco tFPaga: Exit Sub
    End If
    If CCur(tFPaga.Text) = 0 Then
        MsgBox "El importe a pagar para la factura seleccionada debe ser mayor a cero.", vbExclamation, "ATENCIÓN"
        Foco tFPaga: Exit Sub
    End If
    
    If HayPago + CCur(tFPaga.Text) > CCur(tIOriginal.Text) Then
        MsgBox "El importe ingresado (acumulado de pagos) es mayor que el importe original del recibo.", vbExclamation, "ATENCIÓN"
        Foco tIOriginal: Exit Sub
    End If
    
    If CCur(tFPaga.Text) > CCur(lFSaldo.Caption) Then
        MsgBox "El importe ingresado (para saldar la factura) es mayor que el saldo de la factura.", vbExclamation, "ATENCIÓN"
        Foco tFPaga: Exit Sub
    End If
    
    On Error GoTo errAgregar
    'Verifico si la factura está en la lista----------------------------------------
    For I = 1 To vsLista.Rows - 1
        If vsLista.Cell(flexcpValue, I, 4) = Val(tFNumero.Tag) Then
            MsgBox "La factura seleccionada ya está ingresada. Verifique la lista de facturas pagas.", vbInformation, "ATENCIÓN"
            Exit Sub
        End If
    Next
        
    'Agrego la factura a la lista de facturas pagas----------------------------------------
    Screen.MousePointer = 11
    With vsLista
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = Trim(tFSerie.Text) & " " & Trim(tFNumero.Text)
        .Cell(flexcpText, .Rows - 1, 1) = lFFecha.Caption
        
        '.Cell(flexcpText, .Rows - 1, 2) = lFImporteD.Caption
        '.Cell(flexcpText, .Rows - 1, 3) = lFImporteP.Caption
        
        .Cell(flexcpText, .Rows - 1, 2) = lFSaldo.Caption
        .Cell(flexcpText, .Rows - 1, 3) = Format(.Cell(flexcpValue, .Rows - 1, 2) * CCur(lFTC.Caption), FormatoMonedaP)
        
        .Cell(flexcpText, .Rows - 1, 4) = Format(tFNumero.Tag, "#,##0")
        .Cell(flexcpText, .Rows - 1, 5) = Format(tFPaga.Text, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 6) = lFSaldo.Caption
        .Cell(flexcpText, .Rows - 1, 7) = lFTC.Caption
        
        .Cell(flexcpBackColor, .Rows - 1, 2, , 3) = Colores.Inactivo
    End With
    
    'Busco las diferencias de cambio para la compra ingresada.------------------------
    Cons = "Select * from Compra Where ComDCDe = " & Val(tFNumero.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsLista
        .AddItem ""
            
            If Not IsNull(RsAux!ComSerie) Then .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!ComSerie) & " "
            If Not IsNull(RsAux!ComNumero) Then .Cell(flexcpText, .Rows - 1, 0) = .Cell(flexcpText, .Rows - 1, 0) & Trim(RsAux!ComNumero)
            
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ComFecha, "dd/mm/yyyy")
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ComImporte, FormatoMonedaP)
            
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!ComCodigo, "#,##0")
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!ComImporte, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!ComImporte, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!ComTC, "#.000")
            .Cell(flexcpBackColor, .Rows - 1, 2, , 3) = Colores.Inactivo
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    tFSerie.Text = "": tFNumero.Text = "": tFPaga.Text = ""
    If CCur(tIOriginal.Text) = HayPago Then Foco tComentario Else Foco tFSerie
    
    Screen.MousePointer = 0     '-------------------------------------------------------------------
    Exit Sub
    
errAgregar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al agregar la factura a la lista.", Err.Description
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    
    ObtengoSeteoForm Me, Me.Left, Me.Top, Me.Width, Me.Height
    
    sNuevo = False: sModificar = False
    InicializoGrillas
    
    DeshabilitoIngreso
    LimpioFicha
    
    FechaDelServidor
    
    If Trim(Command()) <> "" Then CargoCamposDesdeBD Val(Command())
    If gIDComprobante <> 0 Then Botones True, True, True, False, False, Toolbar1, Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
End Sub

Private Sub Label1_Click()
    Foco tFSerie
End Sub

Private Sub Label10_Click()
    Foco tComentario
End Sub


Private Sub Label11_Click()
    Foco tID
End Sub

Private Sub Label2_Click()
    Foco tFPaga
End Sub

Private Sub Label3_Click()
    Foco tProveedor
End Sub

Private Sub Label4_Click()
    Foco tFecha
End Sub

Private Sub Label5_Click()
    Foco tSerie
End Sub

Private Sub Label9_Click()
    Foco tTCDolar
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuExit_Click()
    Unload Me
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

Sub AccionNuevo(Optional DesdeNuevo As Boolean = False)

    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso

    If DesdeNuevo Then
        tNumero.Text = "": tSerie.Text = "": tIOriginal.Text = ""
        tFSerie.Text = "": tFNumero.Text = ""
        tComentario.Text = ""
        
        vsLista.Rows = 1
        Foco tSerie
    Else
        LimpioFicha
        Foco tFecha
    End If
        
    'vsLista.Editable = True
    gIDComprobante = 0
    
End Sub

Private Sub AccionModificar()

    On Error Resume Next
    Screen.MousePointer = 11
    If Not ValidoDatosMovimientos(gIDComprobante) Then Exit Sub
        
    LimpioFicha
    CargoCamposDesdeBD gIDComprobante
    If gIDComprobante = 0 Then Screen.MousePointer = 0: Exit Sub
    
    sModificar = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso
    Foco tFecha
    Screen.MousePointer = 0
    
End Sub
Private Sub AccionGrabar()

Dim aError As String: aError = ""
Dim bNuevoIngreso As Boolean: bNuevoIngreso = False

    bIngresarPagos = False
    Screen.MousePointer = 11
    If Not ValidoCampos Then Screen.MousePointer = 0: Exit Sub
    If Not ValidoDocumento Then Screen.MousePointer = 0: Exit Sub
    
    Screen.MousePointer = 0
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    
    FechaDelServidor
    If sNuevo Then
        bNuevoIngreso = True
        Dim aCompra As Long
        
        On Error GoTo errorBT
        cBase.BeginTrans    'COMIENZO TRANSACCION----------------------------------------------!!!!!!!!!!!!!!!!!!!!!
        On Error GoTo errorET
        
        'Cargo tabla: Compra----------------------------------------------------------------
        Cons = "Select * from Compra Where ComCodigo = 0"
        Set RsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsCom.AddNew
        CargoCamposBDComprobante
        RsCom.Update: RsCom.Close
        
        Cons = "Select Max(ComCodigo) from Compra"      'Saco el ID del Recibo
        Set RsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        aCompra = RsCom(0)
        RsCom.Close
        
        CargoCamposBDCompraPago aCompra     'Cargo tabla: CompraPago
        GraboBDGastosSubRubro aCompra
        
        ActualizoDivisaPaga
        
        cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
    
        bIngresarPagos = True
        
    Else                                    'Modificar----
    
        Screen.MousePointer = 11
        On Error GoTo errorBT
                
        cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
        On Error GoTo errorET
        
        aCompra = gIDComprobante
        
         'Cargo tabla: Compra----------------------------------------------------------------
        Cons = "Select * from Compra Where ComCodigo = " & gIDComprobante
        Set RsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If gFModificacion <> RsCom!ComFModificacion Then
            aError = "El comprobante ha sido modificado recientemente por otro usuario. Vuelva a cargar los datos"
            GoTo errorET: Exit Sub
        End If
        RsCom.Edit
        CargoCamposBDComprobante
        RsCom.Update: RsCom.Close
        
        'Elimino tabla: Pagos y actualizo los saldos de las compras
        EliminoPagos aCompra
        
        'Cargo tabla: CompraPago
        CargoCamposBDCompraPago aCompra
        
        GraboBDGastosSubRubro aCompra, True
        
        cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
        
    End If
    
    sNuevo = False: sModificar = False
    gIDComprobante = aCompra: gFModificacion = gFechaServidor
    DeshabilitoIngreso
    Botones True, True, True, False, False, Toolbar1, Me
    Foco tFecha
    
    If bIngresarPagos Then AccionIngresoPagos aCompra
    If bNuevoIngreso Then AccionNuevo True
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
    If Trim(aError) = "" Then aError = "No se ha podido inicializar la transacción. Reintente la operación."
    Screen.MousePointer = 0
    clsGeneral.OcurrioError aError, Err.Description
End Sub

Private Sub GraboBDGastosSubRubro(IdCompra As Long, Optional Elimino As Boolean = False)

Dim AlSubrubro As Long

    If Elimino Then
        Cons = "Delete GastoSubrubro Where GSrIDCompra = " & IdCompra
        cBase.Execute Cons
    End If
    
    
    AlSubrubro = paSubrubroAcreedoresVarios
    'Busco si el proveedor tiene un subrubro para asignarlo, sino queda al SR AcreedoresVarios
    Cons = " Select * from EmpresaDato" & _
               " Where EDaCodigo = " & Val(tProveedor.Tag) & _
               " And EDaTipoEmpresa = " & TipoEmpresa.Cliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux!EDaSRubroContable) Then AlSubrubro = RsAux!EDaSRubroContable
    RsAux.Close
    
    Cons = "Select * from GastoSubrubro Where GSrIDCompra = " & IdCompra
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    RsAux.AddNew
    RsAux!GSrIDCompra = IdCompra
    RsAux!GSrIDSubrubro = AlSubrubro
    RsAux!GSrImporte = CCur(tIOriginal.Text)
    RsAux.Update: RsAux.Close
    
End Sub


Private Sub EliminoPagos(aRecibo As Long)

    'Actualizo los saldos de las facturas
    Cons = "Select * from CompraPago Where CPaDocQSalda = " & aRecibo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        Cons = "Update Compra Set ComSaldo = ComSaldo + " & RsAux!CPaAmortizacion & ", " _
                                            & " ComFModificacion = '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
                & " Where ComCodigo = " & RsAux!CPaDocASaldar
                                            
        cBase.Execute Cons
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Elimino los pagos ingresados
    Cons = "Delete CompraPago Where CPaDocQSalda = " & aRecibo
    cBase.Execute Cons
    
End Sub

Private Sub ActualizoDivisaPaga()

    Dim rs1 As rdoResultset
    With vsLista
    
    For I = 1 To .Rows - 1
        If .Cell(flexcpValue, I, 6) - .Cell(flexcpValue, I, 5) = 0 And IsNumeric(.Cell(flexcpText, I, 2)) Then
            Cons = "Select * from GastoImportacion " _
                   & " Where GImIDCompra = " & .Cell(flexcpValue, I, 4) _
                   & " And GImIDSubrubro = " & paSubrubroDivisa _
                   & " And GImNivelFolder = " & Folder.cFEmbarque
            Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rs1.EOF Then
                Cons = " Update Embarque Set EmbDivisaPaga = 1 Where EmbID = " & rs1!GImFolder
                cBase.Execute Cons
            End If
            rs1.Close
        End If
    Next
    End With

End Sub

Private Sub AccionEliminar()
Dim aError As String
    
    If Not ValidoDatosMovimientos(gIDComprobante) Then Exit Sub
    
    If MsgBox("Confirma eliminar el recibo de pago seleccionado", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        Screen.MousePointer = 11
        
        On Error GoTo errorBT
        cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
        On Error GoTo errorET
        
        Cons = "Select * from Compra Where ComCodigo = " & gIDComprobante
        Set RsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If gFModificacion <> RsCom!ComFModificacion Then
            aError = "El comprobante ha sido modificado recientemente por otro usuario. Vuelva a cargar los datos"
            GoTo errorET: Exit Sub
        End If
        
        'Elimino tabla: Registro de pagos
        EliminoPagos gIDComprobante
        
        'Elimino relacion GastosSubrubro------
        Cons = "Delete GastoSubrubro Where GSrIDCompra = " & gIDComprobante
        cBase.Execute Cons

        RsCom.Delete: RsCom.Close
        
        'Actualizo Divisas Como Impagas------------------------------------------------------
        Dim rs1 As rdoResultset
        With vsLista
        For I = 1 To .Rows - 1
            If IsNumeric(.Cell(flexcpText, I, 2)) Then
                Cons = "Select * from GastoImportacion " _
                       & " Where GImIDCompra = " & .Cell(flexcpValue, I, 4) _
                       & " And GImIDSubrubro = " & paSubrubroDivisa _
                       & " And GImNivelFolder = " & Folder.cFEmbarque
                Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rs1.EOF Then
                    Cons = " Update Embarque Set EmbDivisaPaga = 0 Where EmbID = " & rs1!GImFolder
                    cBase.Execute Cons
                End If
                rs1.Close
            End If
        Next
        End With
        
        
        cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
        
        LimpioFicha
        DeshabilitoIngreso
        Botones True, False, False, False, False, Toolbar1, Me
        gIDComprobante = 0
        Screen.MousePointer = 0
    End If
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
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación.", Err.Description
End Sub

Private Sub AccionCancelar()

    LimpioFicha
    If sModificar Then
        Botones True, True, True, False, False, Toolbar1, Me
        CargoCamposDesdeBD gIDComprobante
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    
    DeshabilitoIngreso
    sNuevo = False: sModificar = False
    Foco tFecha

End Sub

Private Sub CargoCamposBDComprobante()

    RsCom!ComTipoDocumento = TipoDocumento.CompraReciboDePago
    RsCom!ComFecha = Format(tFecha.Text, sqlFormatoF)
    RsCom!ComProveedor = Val(tProveedor.Tag)
    RsCom!ComMoneda = paMonedaDolar
    
    If Trim(tSerie.Text) <> "" Then RsCom!ComSerie = Trim(tSerie.Text) Else RsCom!ComSerie = Null
    If Trim(tNumero.Text) <> "" Then RsCom!ComNumero = tNumero.Text Else RsCom!ComNumero = Null
    
    RsCom!ComImporte = CCur(tIOriginal.Text)
    RsCom!ComTC = CCur(tTCDolar.Text)
    
    If Trim(tComentario.Text) <> "" Then RsCom!ComComentario = Trim(tComentario.Text) Else RsCom!ComComentario = Null
    
    RsCom!ComFModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsCom!ComSaldo = 0
    
End Sub

Private Sub CargoCamposBDCompraPago(aRecibo As Long)

    With vsLista
    
    For I = 1 To .Rows - 1
        'Achico el saldo con lo que se paga
        Cons = " Update Compra Set " _
                & " ComSaldo = ComSaldo - " & .Cell(flexcpValue, I, 5) & ", " _
                & " ComFModificacion = '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
                & " Where ComCodigo = " & .Cell(flexcpValue, I, 4)
        cBase.Execute Cons
        
        'Grabo relacion Tabla: CompraPago
        Cons = "Insert into CompraPago (CPaDocASaldar, CPaDocQSalda, CPaAmortizacion) " _
                & "Values (" & .Cell(flexcpValue, I, 4) & ", " & aRecibo & ", " & .Cell(flexcpValue, I, 5) & ")"
        cBase.Execute Cons
    Next
    End With
    
End Sub

Private Sub CargoCamposDesdeBD(IdCompra As Long)

Dim aValor As Long

    Screen.MousePointer = 11
    On Error GoTo errCargar
    'Cargo los datos desde la tabla COMPRA-----------------------------------------------------------------------------------------
    Cons = "Select * from Compra Where ComCodigo = " & IdCompra & " And ComTipoDocumento = " & TipoDocumento.CompraReciboDePago
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe registro de un recibo de pago para el id ingresado.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: gIDComprobante = 0: Botones True, False, False, False, False, Toolbar1, Me: Exit Sub
    End If
        
    gIDComprobante = RsAux!ComCodigo
    gFModificacion = RsAux!ComFModificacion
    
    tID.Text = Format(RsAux!ComCodigo, "#,##0")
    tFecha.Text = Format(RsAux!ComFecha, FormatoFP)
    
    Dim rs1 As rdoResultset
    Cons = "Select * from ProveedorCliente Where PClCodigo = " & RsAux!ComProveedor
    Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rs1.EOF Then
        tProveedor.Text = Trim(rs1!PClFantasia)
        tProveedor.Tag = RsAux!ComProveedor
    End If
    rs1.Close
    
    If Not IsNull(RsAux!ComSerie) Then tSerie.Text = Trim(RsAux!ComSerie)
    If Not IsNull(RsAux!ComNumero) Then tNumero.Text = RsAux!ComNumero
    
    tIOriginal.Text = Format(RsAux!ComImporte, FormatoMonedaP)
    If Not IsNull(RsAux!ComTC) Then If RsAux!ComTC <> 1 Then tTCDolar.Text = Format(RsAux!ComTC, "0.000")
    
    If Not IsNull(RsAux!ComComentario) Then tComentario.Text = Trim(RsAux!ComComentario)
    RsAux.Close
    lPesos.Caption = Format(CCur(tIOriginal.Text) * CCur(tTCDolar.Text), FormatoMonedaP)
    
    'Cargo los datos las facturas pagas-----------------------------------------------------------------------------------------
    Cons = "Select * from CompraPago, Compra" _
           & " Where CPaDocQSalda = " & IdCompra _
           & " And CPaDocASaldar = ComCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsLista
            .AddItem ""
            If Not IsNull(RsAux!ComSerie) Then .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!ComSerie) & " "
            .Cell(flexcpText, .Rows - 1, 0) = .Cell(flexcpText, .Rows - 1, 0) & RsAux!ComNumero
            
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ComFecha, "dd/mm/yyyy")
            
            If RsAux!ComMoneda = paMonedaDolar Then
'                If Not IsNull(RsAux!ComIva) Then
'                    .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ComImporte + RsAux!ComIva, FormatoMonedaP)
'                Else
'                    .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ComImporte, FormatoMonedaP)
'                End If
                .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!CPaAmortizacion, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 3) = Format(.Cell(flexcpValue, .Rows - 1, 2) * RsAux!ComTC, FormatoMonedaP)
            Else
                'If Not IsNull(RsAux!ComIva) Then
                '    .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ComImporte + RsAux!ComIva, FormatoMonedaP)
                'Else
                '    .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ComImporte, FormatoMonedaP)
                'End If
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!CPaAmortizacion, FormatoMonedaP)
            End If
            
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!ComCodigo, "#,##0")
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!CPaAmortizacion, FormatoMonedaP)
            
            .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!ComTC, "#0.000")
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Screen.MousePointer = 0
    Exit Sub
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del comprobante.", Err.Description
End Sub

Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0: tComentario.SelLength = Len(tComentario.Text)
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub


Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And Not sNuevo And Not sModificar Then AccionListaDeAyuda
    If KeyCode = vbKeyDown Then tFecha.Text = Format(Now, FormatoFP)
End Sub

Private Sub AccionListaDeAyuda()

    On Error GoTo errAyuda
    
    If Not IsDate(tFecha.Text) And Val(tProveedor.Tag) = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    Dim aLista As New clsListadeAyuda
    Dim aSeleccionado As Long: aSeleccionado = 0
    
    Cons = " Select ID_Compra = ComCodigo, Fecha = ComFecha, Proveedor = PClFantasia, Comprobante = ComSerie + Convert(char(10), ComNumero), Moneda = MonSigno , Importe = ComImporte, Comentarios = ComComentario" _
            & " from Compra, ProveedorCliente, Moneda" _
            & " Where ComProveedor = PClCodigo" _
            & " And ComMoneda = MonCodigo And ComTipoDocumento = " & TipoDocumento.CompraReciboDePago
            
    If IsDate(tFecha.Text) Then Cons = Cons & " And ComFecha >= '" & Format(tFecha.Text, sqlFormatoF) & "'"
    If Val(tProveedor.Tag) <> 0 Then Cons = Cons & " And ComProveedor = " & Val(tProveedor.Tag)
    
    Cons = Cons & " Order by ComFecha DESC"
    
    aLista.ActivoListaAyudaSQL Cons, miConexion.TextoConexion(logImportaciones)
    Me.Refresh
    
    If IsNumeric(aLista.ItemSeleccionadoSQL) Then aSeleccionado = CLng(aLista.ItemSeleccionadoSQL)
    Set aLista = Nothing
    
    If aSeleccionado <> 0 Then LimpioFicha: CargoCamposDesdeBD aSeleccionado
    If gIDComprobante <> 0 Then Botones True, True, True, False, False, Toolbar1, Me
    
    Screen.MousePointer = 0
    Exit Sub
        
errAyuda:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al activar la lista de ayuda.", Err.Description
End Sub

Private Sub tFNumero_Change()
    tFNumero.Tag = 0
    lFImporteD.Caption = "": lFImporteP.Caption = "": lFFecha.Caption = "": lFSaldo.Caption = "": lFTC.Caption = ""
End Sub

Private Sub tFNumero_GotFocus()
    Status.Panels(4).Text = "F1- Lista de facturas pendientes."
End Sub

Private Sub tFNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call tFSerie_KeyDown(vbKeyF1, False)
End Sub

Private Sub tFNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) = 0 Then MsgBox "Ingrese el proveedor del recibo de pago.", vbExclamation, "ATENCIÓN": Foco tProveedor: Exit Sub
        If Not IsNumeric(tIOriginal.Text) Then MsgBox "Ingrese el importe del recibo de pago.", vbExclamation, "ATENCIÓN": Foco tProveedor: Exit Sub
        
        If Val(tFNumero.Tag) <> 0 Then
            'tFPaga.Text = Format(CCur(tIOriginal.Text) - HayPago, FormatoMonedaP)
            'Foco tFPaga: Exit Sub
            bAgregar.SetFocus: Exit Sub
        End If
        
        If Not IsNumeric(tFNumero.Text) Then MsgBox "Ingrese el número de la factura que se paga.", vbExclamation, "ATENCIÓN": Foco tFNumero: Exit Sub
        
        On Error GoTo errBusco
        Screen.MousePointer = 11
        'Busco la factura
        
        Cons = "Select ID_Compra = ComCodigo, Fecha = ComFecha, Serie = ComSerie, ComNumero 'Número', ComSaldo as 'Saldo a Pagar', Comentarios = ComComentario from Compra " _
                & " Where ComProveedor = " & CLng(tProveedor.Tag) _
                & " And ComTipoDocumento = " & TipoDocumento.CompraCredito _
                & " And ComNumero = " & Trim(tFNumero.Text) _
                & " And ComMoneda = " & paMonedaDolar
        If Trim(tFSerie.Text) <> "" Then Cons = Cons & " And ComSerie = '" & Trim(tFSerie.Text) & "'"

        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Dim aFSeleccionada As Long

        If Not RsAux.EOF Then
            aFSeleccionada = RsAux(0)
            'Valido si hay mas de una factura-------------------------------------------------------------
            RsAux.MoveNext
            If Not RsAux.EOF Then
                RsAux.Close
                aFSeleccionada = ListaDeFacturas(Cons)
                If aFSeleccionada = 0 Then Screen.MousePointer = 0: Exit Sub
            Else
                RsAux.Close
            End If
            
            CargoDatosFactura aFSeleccionada
            
        Else
            MsgBox "No existe una factura para el proveedor, moneda y número de documento ingresado.", vbInformation, "ATENCIÓN"
            RsAux.Close
        End If
        
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errBusco:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar la factura ingresada.", Err.Description
End Sub

Private Sub CargoDatosFactura(Codigo As Long)

    Cons = "Select * from Compra Where ComCodigo = " & Codigo & " And ComTipoDocumento = " & TipoDocumento.CompraCredito
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "El comprobante seleccionado no es del tipo Crédito. Verifique.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    Else
        If paMonedaDolar <> RsAux!ComMoneda Then
            RsAux.Close
            MsgBox "La moneda del comprobante seleccionado es distinta a la del recibo de pago. Verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    '----------------------------------------------------------------------------------------------------

    If Not IsNull(RsAux!ComSerie) Then tFSerie.Text = Trim(RsAux!ComSerie)
    If Not IsNull(RsAux!ComNumero) Then tFNumero.Text = RsAux!ComNumero
    tFNumero.Tag = RsAux!ComCodigo
    
    If Not IsNull(RsAux!ComTC) Then lFTC.Caption = Format(RsAux!ComTC, "#,##0.000") Else lFTC.Caption = "1.000"
    If Not IsNull(RsAux!ComIva) Then
        lFImporteD.Caption = Format(RsAux!ComImporte + RsAux!ComIva, FormatoMonedaP)
    Else
        lFImporteD.Caption = Format(RsAux!ComImporte, FormatoMonedaP)
    End If
    lFImporteP.Caption = Format(CCur(lFImporteD.Caption) * CCur(lFTC.Caption), FormatoMonedaP)
    
    lFFecha.Caption = Format(RsAux!ComFecha, "dd/mm/yyyy")
        
    If Not IsNull(RsAux!ComSaldo) Then
        lFSaldo.Caption = Format(RsAux!ComSaldo, FormatoMonedaP)
        If RsAux!ComSaldo = 0 Then
            MsgBox "El saldo de la factura seleccionada es cero. Verifique los pagos de ésta factura antes de continuar.", vbExclamation, "Saldo Cero"
        Else
            
            'Verifico si hay ingreso de vencimientos para cargar el valor de una cta.
            Dim rs1 As rdoResultset
            Dim hayVencimientos As Boolean: hayVencimientos = False
            Dim hayPagos As Integer: hayPagos = 0
            Dim aImporteAP As Currency: aImporteAP = 0
            
            Cons = "Select * from CompraVencimiento Where CVeIDCompra = " & Codigo
            Set rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not rs1.EOF Then hayVencimientos = True
            rs1.Close
            
            If hayVencimientos Then
                Cons = "Select Count(*) from CompraPago Where CPaDocASaldar = " & Codigo
                Set rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not rs1.EOF Then If Not IsNull(rs1(0)) Then hayPagos = rs1(0)
                rs1.Close
                
                Cons = "Select * from CompraVencimiento Where CVeIDCompra = " & Codigo
                Set rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                Dim aRow As Integer: aRow = 1
                Do While Not rs1.EOF
                    If hayPagos < aRow Then
                        aImporteAP = rs1!CVeImporte
                        Exit Do
                    Else
                        aRow = aRow + 1
                    End If
                    rs1.MoveNext
                Loop
                rs1.Close
                
                If aImporteAP > 0 Then tFPaga.Text = Format(aImporteAP, FormatoMonedaP)
            End If
            
            If aImporteAP = 0 Then
                'If CCur(tIOriginal.Text) - HayPago > RsAux!ComSaldo Then
                    tFPaga.Text = Format(RsAux!ComSaldo, FormatoMonedaP)
                'Else
                '    tFPaga.Text = Format(CCur(tIOriginal.Text) - HayPago, FormatoMonedaP)
                'End If
            End If
            
        End If
    End If
    RsAux.Close
    Foco bAgregar
            
End Sub

Private Sub tFNumero_LostFocus()
    Status.Panels(4).Text = ""
End Sub

Private Sub tFPaga_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsNumeric(tFPaga.Text) Then bAgregar.SetFocus
End Sub

Private Sub tFPaga_LostFocus()
    On Error Resume Next
    If IsNumeric(tFPaga.Text) Then tFPaga.Text = Format(tFPaga.Text, FormatoMonedaP)
End Sub

Private Sub tFSerie_Change()
    tFNumero.Tag = 0
    lFImporteD.Caption = "": lFImporteP.Caption = "": lFFecha.Caption = "": lFSaldo.Caption = "": lFTC.Caption = ""
End Sub

Private Sub tFSerie_GotFocus()
    tFSerie.SelStart = 0: tFSerie.SelLength = Len(tFSerie.Text)
    Status.Panels(4).Text = "F1- Lista de facturas pendientes."
End Sub

Private Sub tFSerie_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        If Val(tProveedor.Tag) = 0 Then MsgBox "Ingrese el proveedor del recibo de pago.", vbExclamation, "ATENCIÓN": Foco tProveedor: Exit Sub
        If Not IsNumeric(tIOriginal.Text) Then MsgBox "Ingrese el importe del recibo de pago.", vbExclamation, "ATENCIÓN": Foco tProveedor: Exit Sub
        
        On Error GoTo errBusco
        Screen.MousePointer = 11
        
        'Cons = "Select ID_Compra = ComCodigo, Fecha = ComFecha, Serie = ComSerie, ComNumero 'Número', Importe = ComImporte, ComIva 'I.V.A', Saldo = ComSaldo, Comentarios = ComComentario from Compra "
        Cons = "Select ID_Compra = ComCodigo, Fecha = ComFecha, Serie = ComSerie, ComNumero 'Número', ComSaldo as 'Saldo a Pagar', Comentarios = ComComentario from Compra " _
                & " Where ComProveedor = " & CLng(tProveedor.Tag) _
                & " And ComTipoDocumento = " & TipoDocumento.CompraCredito _
                & " And ComMoneda = " & paMonedaDolar _
                & " And ComSaldo > 0 "
        
        Dim aFSeleccionada As Long
        aFSeleccionada = ListaDeFacturas(Cons)
        If aFSeleccionada = 0 Then Screen.MousePointer = 0: Exit Sub
        CargoDatosFactura aFSeleccionada
        
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errBusco:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar las facturas del proveedor.", Err.Description
End Sub

Private Sub tFSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then If Trim(tFSerie.Text) = "" And vsLista.Rows > 1 Then Foco tComentario Else Foco tFNumero
End Sub

Private Sub tFSerie_LostFocus()
    Status.Panels(4).Text = ""
End Sub

Private Sub tID_Change()
    If tID.Enabled Then Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub tID_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tID.Text) = "" Then Foco tFecha: Exit Sub
        If Not IsNumeric(tID.Text) Then MsgBox "El id ingresado no es correcto. Verifique.", vbExclamation, "ATENCIÓN": Exit Sub
        gIDComprobante = CLng(tID.Text)
        LimpioFicha
        CargoCamposDesdeBD gIDComprobante
        If gIDComprobante <> 0 Then Botones True, True, True, False, False, Toolbar1, Me
    End If
    
End Sub

Private Sub tNumero_GotFocus()
    tNumero.SelStart = 0: tNumero.SelLength = Len(tNumero.Text)
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If vsLista.Rows = 1 Then tComentario.Text = "Pago Imp. " & Trim(tSerie.Text) & " " & Trim(tNumero.Text)
        Foco tIOriginal
    End If
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0
    tFecha.SelLength = Len(tFecha.Text)
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If tProveedor.Enabled Then Foco tProveedor Else Foco tSerie
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, FormatoFP) Else tFecha.Text = ""
End Sub

Private Sub tIOriginal_GotFocus()
    tIOriginal.SelStart = 0
    tIOriginal.SelLength = Len(tIOriginal.Text)
End Sub

Private Sub tIOriginal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tFecha.Text) Then
            Dim aFechaTC As String: aFechaTC = ""
            tTCDolar.ToolTipText = "": lTC.Caption = ""
            
            'TC del ultimo dia del mes anterior
            tTCDolar.Text = TasadeCambio(paMonedaDolar, paMonedaPesos, UltimoDia(DateAdd("m", -1, CDate(tFecha.Text))), aFechaTC)
            lTC.Caption = aFechaTC
        End If
        If IsNumeric(tTCDolar.Text) And IsNumeric(tIOriginal.Text) Then lPesos.Caption = Format(CCur(tIOriginal.Text) * CCur(tTCDolar.Text), FormatoMonedaP)
        Foco tTCDolar
    End If
End Sub

Private Sub tIOriginal_LostFocus()

    If Not IsNumeric(tIOriginal.Text) Then tIOriginal.Text = ""
    tIOriginal.Text = Format(tIOriginal.Text, "##,##0.00")
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        
        Case "pagos": EjecutarApp App.Path & "\Con Que Paga", CStr(gIDComprobante)
        Case "dolar": EjecutarApp App.Path & "\Tasa de Cambio"
        
        Case "salir": Unload Me
    End Select

End Sub

Private Sub AccionIngresoPagos(Optional Compra As Long = 0)
Dim aRetorno As Integer

    'Veo si están cargados los valores de la disponibilidad por defecto------------------------------
    On Error GoTo err1
    If aMonedaDisponibilidad = 0 Then
        Screen.MousePointer = 11
        Cons = "Select * from Disponibilidad Where DisID = " & paDisponibilidad
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            txtDisponibilidad = Trim(RsAux!DisNombre)
            aMonedaDisponibilidad = RsAux!DisMoneda
            If Not IsNull(RsAux!DisSucursal) Then bEsBancaria = True Else bEsBancaria = False
        End If
        RsAux.Close
        Screen.MousePointer = 0
    End If  '--------------------------------------------------------------------------------------------------
    
    If Not bEsBancaria And aMonedaDisponibilidad = paMonedaDolar Then
        aRetorno = MsgBox("Desea ingresar todo el pago con la disponibilidad <<" & txtDisponibilidad & ">>" & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                                 & "Si- Graba movimiento automático." & Chr(vbKeyReturn) _
                                 & "No- Abre 'Con Que Paga'.", vbQuestion + vbYesNoCancel, "Ingreso del Pago")
        If aRetorno = vbCancel Then Exit Sub
        Screen.MousePointer = 11
        If aRetorno = vbNo Then
            EjecutarApp App.Path & "\Con Que Paga", Str(gIDComprobante)
        Else
            'Ingreso el pago automático----------------------------------------------------------------------------------------------------
            On Error GoTo errorBT
            cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
            On Error GoTo errorET
            MovimientoDeCaja paMDPagoDeCompra, CDate(tFecha.Text), paDisponibilidad, aMonedaDisponibilidad, CCur(tIOriginal.Text), Trim(tComentario.Text), True, Compra
            cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------------------
        End If
        
    Else
        If MsgBox("Desea ingresar con qué paga el comprobante.", vbQuestion + vbYesNo, "Ingreso del Pago ") = vbNo Then Exit Sub
        EjecutarApp App.Path & "\Con Que Paga", Str(gIDComprobante)
    End If
    Screen.MousePointer = 0
    Exit Sub

err1:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar los datos de la disponbilidad por defecto.", Err.Description
    Exit Sub
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans: Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
End Sub

Private Sub DeshabilitoIngreso()

    tID.BackColor = Blanco: tID.Enabled = True
    tFecha.BackColor = Blanco
    tProveedor.Enabled = True: tProveedor.BackColor = Blanco
    tSerie.Enabled = False: tSerie.BackColor = Inactivo
    tNumero.Enabled = False: tNumero.BackColor = Inactivo
    
    tIOriginal.Enabled = False: tIOriginal.BackColor = Inactivo
    tTCDolar.Enabled = False: tTCDolar.BackColor = Inactivo
    
    tFSerie.Enabled = False: tFSerie.BackColor = Inactivo
    tFNumero.Enabled = False: tFNumero.BackColor = Inactivo
    tFPaga.Enabled = False: tFPaga.BackColor = Inactivo
    
    tComentario.Enabled = False: tComentario.BackColor = Inactivo
    vsLista.BackColor = Inactivo
    
    bAgregar.Enabled = False
    
End Sub

Private Sub HabilitoIngreso()
    
    tID.BackColor = Inactivo: tID.Enabled = False
    
    tFecha.BackColor = Obligatorio
    
    If Not sModificar Then
        tProveedor.BackColor = Obligatorio
    Else
        tProveedor.BackColor = Inactivo
        tProveedor.Enabled = False
    End If
    
    tSerie.Enabled = True: tSerie.BackColor = Blanco
    tNumero.Enabled = True: tNumero.BackColor = Obligatorio
    
    tIOriginal.Enabled = True: tIOriginal.BackColor = Obligatorio
    tTCDolar.Enabled = True: tTCDolar.BackColor = Obligatorio
    
    tFSerie.Enabled = True: tFSerie.BackColor = Blanco
    tFNumero.Enabled = True: tFNumero.BackColor = Blanco
    'tFPaga.Enabled = True: tFPaga.BackColor = Blanco
    
    tComentario.Enabled = True: tComentario.BackColor = Blanco
    vsLista.BackColor = Blanco
    
    bAgregar.Enabled = True
    
End Sub

Private Sub LimpioFicha()
    
    tID.Text = ""
    tFecha.Text = ""
    tProveedor.Text = ""
    tIOriginal.Text = ""
    tSerie.Text = "": tNumero.Text = ""
    
    tTCDolar.Text = ""
    lTC.Caption = "": lPesos.Caption = ""
    
    tFSerie.Text = "": tFNumero.Text = "": tFPaga.Text = ""
    lFImporteD.Caption = "": lFImporteP.Caption = "": lFFecha.Caption = "": lFSaldo.Caption = "": lFTC.Caption = ""
    
    vsLista.Rows = 1
    tComentario.Text = ""
    
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = 0
End Sub

Private Sub tProveedor_GotFocus()
    tProveedor.SelStart = 0: tProveedor.SelLength = Len(tProveedor.Text)
End Sub

Private Sub tProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And Val(tProveedor.Tag) <> 0 Then AccionListaDeAyuda
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then If tID.Enabled Then Foco tID Else Foco tSerie: Exit Sub
        Screen.MousePointer = 11
        Dim aQ As Long, aIdProveedor As Long, aTexto As String
        
        aQ = 0
        Cons = "Select PClCodigo, PClFantasia, PClNombre from ProveedorCliente " _
                & " Where PClNombre like '" & Trim(tProveedor.Text) & "%' Or PClFantasia like '" & Trim(tProveedor.Text) & "%'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aQ = 1: aIdProveedor = RsAux!PClCodigo: aTexto = Trim(RsAux!PClFantasia)
            RsAux.MoveNext: If Not RsAux.EOF Then aQ = 2
        End If
        RsAux.Close
        
        Select Case aQ
            Case 0:
                    MsgBox "No existe una empresa para el con el nombre ingresado.", vbExclamation, "No existe Empresa"
            
            Case 1:
                    tProveedor.Text = aTexto
                    tProveedor.Tag = aIdProveedor
                    If (sNuevo Or sModificar) Then Foco tSerie
        
            Case 2:
                    Dim aLista As New clsListadeAyuda
                    aLista.ActivoListaAyuda Cons, False, miConexion.TextoConexion(logImportaciones), 5500
                    If aLista.ValorSeleccionado <> 0 Then
                        tProveedor.Text = Trim(aLista.ItemSeleccionado)
                        tProveedor.Tag = aLista.ValorSeleccionado
                        
                        If (sNuevo Or sModificar) Then Foco tSerie
                    Else
                        tProveedor.Text = ""
                    End If
                    Set aLista = Nothing
        End Select
        Screen.MousePointer = 0
    End If
    Exit Sub
    Screen.MousePointer = 0

errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar la lista de ayuda.", Err.Description
End Sub

Private Sub tSerie_GotFocus()
    tSerie.SelStart = 0: tSerie.SelLength = Len(tSerie.Text)
End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then Foco tNumero
End Sub

Private Sub tTCDolar_Change()
    lTC.Caption = "manual"
    lPesos.Caption = ""
End Sub

Private Sub tTCDolar_GotFocus()
    tTCDolar.SelStart = 0: tTCDolar.SelLength = Len(tTCDolar.Text)
End Sub

Private Sub tTCDolar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And tTCDolar.Enabled Then
        If IsNumeric(tTCDolar.Text) And IsNumeric(tIOriginal.Text) Then lPesos.Caption = Format(CCur(tIOriginal.Text) * CCur(tTCDolar.Text), FormatoMonedaP)
        Foco tFSerie
    End If
    
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsLista
        .Rows = 1: .Cols = 1
        .Editable = False
        .FormatString = "<Numero|<Fecha|>Saldo Dólares|>Saldo Pesos|>Id_Compra|>Paga|>Saldo Actual|>TC|"
        .ExtendLastCol = True
        .WordWrap = True
        .ColWidth(0) = 950: .ColWidth(1) = 950: .ColWidth(2) = 1400: .ColWidth(3) = 1500: .ColWidth(5) = 1400: .ColWidth(6) = 1400: .ColWidth(7) = 950
        .ColDataType(2) = flexDTCurrency: .ColDataType(5) = flexDTCurrency
        .ColHidden(5) = True ': .ColHidden(6) = True
        .ExtendLastCol = True
        .AllowUserResizing = flexResizeColumns
    End With
    
End Sub


Private Sub vsLista_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        If vsLista.RowSel < 1 Then Exit Sub
        If Not sNuevo And Not sModificar Then Exit Sub
        On Error Resume Next
        vsLista.RemoveItem vsLista.RowSel
    End If
    
End Sub

Private Function ValidoCampos() As Boolean

Dim aTotal As Currency
    
    On Error GoTo errValido
    ValidoCampos = False
    
    If Not IsDate(tFecha.Text) Then
        MsgBox "La fecha ingresada para el registro del comprobante no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    If Val(tProveedor.Tag) = 0 Then
        MsgBox "Debe seleccionar el proveedor del pago.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Function
    End If
    
    If Not IsNumeric(tNumero.Text) Then
        MsgBox "Debe ingresar la numeración del comprobante.", vbExclamation, "ATENCIÓN"
        Foco tNumero: Exit Function
    End If
    
    If Not IsNumeric(tIOriginal.Text) Then
        MsgBox "Debe ingresar el importe total del recibo de pago.", vbExclamation, "ATENCIÓN"
        Foco tIOriginal: Exit Function
    End If
    
    If Not IsNumeric(tTCDolar.Text) Then
        MsgBox "Debe ingresar el valor del dólar para la fecha ingresada (tasa de cambio).", vbExclamation, "ATENCIÓN"
        Foco tNumero: Exit Function
    End If
    
    If Not IsNumeric(lPesos.Caption) Then
        MsgBox "Debe ingresar el valor del dólar y el importe para registrar el saldo en pesos .", vbExclamation, "ATENCIÓN"
        Foco tIOriginal: Exit Function
    End If
    
    If vsLista.Rows = 1 Then
        MsgBox "Debe ingresar las facturas a pagar con el recibo.", vbExclamation, "ATENCIÓN"
        Foco tFSerie: Exit Function
    End If
       
    'Valido importe del las facturas contra el importe original
    If HayPago <> CCur(tIOriginal.Text) Then
        MsgBox "El importe del comprobante (" & tIOriginal.Text & ") no coincide con la suma de los gastos (" & Format(HayPago, FormatoMonedaP) & ").", vbExclamation, "ATENCIÓN"
        Foco tIOriginal: Exit Function
    End If
    
    Dim aPesos As Currency, aDif As Currency
    aPesos = 0: aDif = 0
    With vsLista
        For I = 1 To .Rows - 1: aPesos = aPesos + .Cell(flexcpValue, I, 3): Next
    End With
    aDif = CCur(lPesos.Caption) - aPesos
    If aDif <> 0 Then
        If aDif < 10 And aDif > -10 Then
            If MsgBox("Existe una diferencia en pesos de " & Format(aDif, FormatoMonedaP) & Chr(vbKeyReturn) & "La diferencia se pudo provocar en la emisión de notas, al corregir las diferencias de cambio ya generadas (proceso automático en embarques)." & Chr(vbKeyReturn) & "Si ud. desea ingonar esta diferencia y continuar con el pago presione SI.", vbExclamation + vbYesNo + vbDefaultButton2, "Diferencia en TC") = vbNo Then
                Foco tIOriginal: Exit Function
            End If
        Else
            MsgBox "El importe del comprobante en pesos (" & lPesos.Caption & ") no coincide con la suma de los gastos (" & Format(aPesos, FormatoMonedaP) & ")." & Chr(vbKeyReturn) & "Verifique la tasa de cambio ingresada.", vbExclamation, "ATENCIÓN"
            Foco tIOriginal: Exit Function
        End If
    End If
     
    ValidoCampos = True
    Exit Function

errValido:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al validar los datos.", Err.Description
End Function

Private Function ValidoDocumento() As Boolean
    
    On Error Resume Next
    ValidoDocumento = False
    
    Cons = "Select * from Compra Where ComCodigo <> " & gIDComprobante
           
    If Trim(tNumero.Text) <> "" Then Cons = Cons & " And ComNumero = " & Trim(tNumero.Text)
    
    Cons = Cons & " And ComProveedor = " & Val(tProveedor.Tag) _
                       & " And ComMoneda = " & paMonedaDolar _
                       & " And ComImporte = " & CCur(tIOriginal.Text) _
                       & " And ComTipoDocumento = " & TipoDocumento.CompraReciboDePago
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Screen.MousePointer = 0
        If MsgBox("Ya existen recibos registrados con el mismo documento y proveedor." & Chr(vbKeyReturn) _
            & "Fecha: " & Format(RsAux!ComFecha, "d-mmm yyyy") & Chr(vbKeyReturn) _
            & "Importe: " & Format(RsAux!ComImporte, "##,##0.00") & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
            & "Desea proseguir con el ingreso del gasto.", vbInformation + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then
                RsAux.Close: Exit Function
        End If
    End If
    RsAux.Close
    Screen.MousePointer = 0
    ValidoDocumento = True
                   
End Function

Private Function ValidoDatosMovimientos(aRecibo As Long) As Boolean

    On Error GoTo errValidar
    ValidoDatosMovimientos = True
    
    'Valido los campos de la tabla vencimiento-------------------------------------------------------------------------------------------
    Cons = "Select * from MovimientoDisponibilidad Where MDiIdCompra = " & aRecibo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        ValidoDatosMovimientos = False
        Screen.MousePointer = 0
        MsgBox "Hay movimientos de disponibilidades ingresados para el comprobante." & Chr(vbKeyReturn) & "Para continuar con la acción debe eliminarlos.", vbInformation, "Movimientos de Disponibilidades"
    End If
    RsAux.Close
    Exit Function

errValidar:
    clsGeneral.OcurrioError "Ocurrió un error al validar movimientos de disponibilidades.", Err.Description
End Function


Private Function HayPago() As Currency
    Dim aRetorno As Currency
    
    aRetorno = 0
    With vsLista
        For I = 1 To .Rows - 1
            If IsNumeric(.Cell(flexcpText, I, 2)) Then aRetorno = aRetorno + .Cell(flexcpValue, I, 5)
        Next
    End With
    HayPago = aRetorno
    
End Function

Private Function ListaDeFacturas(Consulta As String) As Long

    On Error GoTo errAyuda
    ListaDeFacturas = 0
    
    Dim aLista As New clsListadeAyuda
    Dim aSeleccionado As Long: aSeleccionado = 0
    
    'Cons = " Select ID_Compra = ComCodigo, Fecha = ComFecha, Proveedor = PClFantasia, Comprobante = ComSerie + Convert(char(10), ComNumero), Moneda = MonSigno , Importe = ComImporte, Comentarios = ComComentario" _
            & " from Compra, ProveedorCliente, Moneda" _
            & " Where ComProveedor = PClCodigo" _
            & " And ComMoneda = MonCodigo"
            
    aLista.ActivoListaAyudaSQL Consulta, miConexion.TextoConexion(logImportaciones)
    Me.Refresh
    
    If IsNumeric(aLista.ItemSeleccionadoSQL) Then aSeleccionado = CLng(aLista.ItemSeleccionadoSQL)
    Set aLista = Nothing
    
    ListaDeFacturas = aSeleccionado
    Screen.MousePointer = 0
    Exit Function
        
errAyuda:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al activar la lista de ayuda.", Err.Description
End Function


