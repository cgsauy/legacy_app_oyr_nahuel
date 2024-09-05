VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmListado 
   Caption         =   "Asientos Varios"
   ClientHeight    =   7590
   ClientLeft      =   1800
   ClientTop       =   1935
   ClientWidth     =   11880
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
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   Begin VB.PictureBox picTab 
      Height          =   2715
      Index           =   1
      Left            =   2580
      ScaleHeight     =   2655
      ScaleWidth      =   4455
      TabIndex        =   24
      Top             =   3540
      Width           =   4515
      Begin VSFlex6DAOCtl.vsFlexGrid vsReval 
         Height          =   2775
         Left            =   360
         TabIndex        =   27
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   4895
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
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
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
   Begin VB.PictureBox picTab 
      Height          =   4215
      Index           =   0
      Left            =   1140
      ScaleHeight     =   4155
      ScaleWidth      =   1395
      TabIndex        =   23
      Top             =   2340
      Width           =   1455
      Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
         Height          =   2775
         Left            =   480
         TabIndex        =   25
         Top             =   120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4895
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
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsMovimiento 
         Height          =   2415
         Left            =   1500
         TabIndex        =   26
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4260
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
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
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
   Begin MSComctlLib.TabStrip tabAsientos 
      Height          =   915
      Left            =   420
      TabIndex        =   22
      Top             =   2880
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   1614
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Asientos Comunes"
            Key             =   "comun"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Revaluación de Ventas"
            Key             =   "revaluacion"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   1995
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   3519
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
      Zoom            =   100
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   11475
      TabIndex        =   19
      Top             =   6720
      Width           =   11535
      Begin VB.CommandButton bExportar 
         Height          =   310
         Left            =   5220
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Exportar"
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":074C
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":084E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmListado.frx":0D80
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmListado.frx":0E6A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin MSComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6060
         TabIndex        =   20
         Top             =   120
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   7335
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "bd"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12753
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
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
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   11175
      Begin VB.CheckBox opNetear 
         Caption         =   "&Netear Asientos"
         Height          =   195
         Left            =   4500
         TabIndex        =   21
         Top             =   280
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox tFHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Text            =   "28/12/2000"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "28/12/2000"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   255
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   11280
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":0F54
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":126E
            Key             =   "help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1380
            Key             =   "consultar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":14DA
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1634
            Key             =   "print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":178E
            Key             =   "limpiar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":18A0
            Key             =   "vista2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":19FA
            Key             =   "vista1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1B54
            Key             =   "move2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1CAE
            Key             =   "move3"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1E08
            Key             =   "move4"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1F62
            Key             =   "move1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":20BC
            Key             =   "configprint"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog ctlDlg 
      Left            =   8940
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

'Data Asientos de Ventas -------------------------
Private Type typRenglon
    NombreRubro As String
    Importe As Currency
End Type

Dim arrAsientos() As typRenglon
Dim arrAsientosCtas() As typRenglon     'Para cargar los datos de la cobranza de ctas.
'-------------------------------------------------------

'Data Tasas de Cambio   --------------------------
Dim prmMonedaTC As Long
Private Type typTC
    Fecha As String
    Valor As Currency
End Type

Dim arrTC() As typTC
Dim arrTCMA() As Currency
'-------------------------------------------------------

Private Sub AccionLimpiar()
    tFecha.Text = "": tFHasta.Text = ""
    vsConsulta.Rows = 1: vsMovimiento.Rows = 1
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    vsConsulta.SetFocus
    DoEvents
    AccionConsultar
End Sub

Private Sub bExportar_Click()
        
On Error GoTo errCancel
    
    With ctlDlg
        .CancelError = True
        
        .FileName = "AsientosVarios"
        
        .Filter = "Libro de Microsoft Exel|*.xls|" & _
                     "Texto (delimitado por tabulaciones)|*.txt|" & "Texto (delimitado por comas)|*.txt"
        
        .ShowSave
        
        'Confirma exportar el contenido de la lista al archivo:
        If MsgBox("Confirma exportar el contenido de la lista al archivo: " & .FileName, vbQuestion + vbYesNo) = vbYes Then
        
            On Error GoTo errSaving
            Screen.MousePointer = 11
            Me.Refresh
            DoEvents
            
            Dim mSSetting As SaveLoadSettings
            
            Select Case .FilterIndex
                Case 1, 2: mSSetting = flexFileTabText
                Case 3: mSSetting = flexFileCommaText
            End Select
            
            vsConsulta.SaveGrid .FileName, mSSetting, True
            
            Screen.MousePointer = 0
        End If
        
    End With
    
errCancel:
    Screen.MousePointer = 0
    Exit Sub
errSaving:
     clsGeneral.OcurrioError "Error al exportar el contenido de la lista.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bImprimir_Click()
    AccionImprimir True
End Sub
Private Sub bNoFiltros_Click()
    AccionLimpiar
End Sub

Private Sub bPrimero_Click()
    IrAPagina vsListado, 1
End Sub

Private Sub bSiguiente_Click()
    IrAPagina vsListado, vsListado.PreviewPage + 1
End Sub

Private Sub bUltima_Click()
    IrAPagina vsListado, vsListado.PageCount
End Sub

Private Sub bZMas_Click()
    Zoom vsListado, vsListado.Zoom + 5
End Sub

Private Sub bZMenos_Click()
    Zoom vsListado, vsListado.Zoom - 5
End Sub

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub


Private Sub chVista_Click()
    
    If chVista.Value = 0 Then
        'vsConsulta.ZOrder 0
        vsListado.Visible = False
    Else
        AccionImprimir
        vsListado.ZOrder 0
        vsListado.Visible = True
    End If

End Sub

Private Sub Label2_Click()
    Foco tFHasta
End Sub

Private Sub opNetear_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub tabAsientos_Click()
    
    Select Case LCase(tabAsientos.SelectedItem.Key)
        Case "comun": picTab(0).ZOrder 0
        Case "revaluacion": picTab(1).ZOrder 0
    End Select
    
End Sub

Private Sub tFecha_GotFocus()
    With tFecha: .SelStart = 0: .SelLength = Len(.Text): End With
    Ayuda "Ingrese una fecha de compra."
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFHasta
End Sub
Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then
        tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
        If Not IsDate(tFHasta.Text) Then tFHasta.Text = tFecha.Text
    End If
    Ayuda ""
End Sub

Private Sub Label5_Click()
    Foco tFecha
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrLoad
    StartMe
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    pbProgreso.Value = 0
    
    InicializoGrillas
    AccionLimpiar
    bCargarImpresion = True
    
    CargoConstantesSubrubros
    
    vsListado.Orientation = orPortrait
    vsListado.Visible = False
    tabAsientos.Tabs("comun").Selected = True
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    vsListado.MarginLeft = 300
    vsListado.MarginRight = 300
    vsListado.PhysicalPage = True
    
    With vsConsulta
        .ColSel = 0
        .ColSort(0) = flexSortStringAscending
        .Sort = flexSortUseColSort
        
        .Cols = 1: .Rows = 1:
        .FormatString = "Rubro Orden|<Rubro|>Debe $|>Debe M/E|>Haber $|>Haber M/E|>T/C|"
            
        .WordWrap = False
        .ColWidth(0) = 0
        .ColWidth(1) = 3300: .ColWidth(2) = 1300: .ColWidth(3) = 1300: .ColWidth(4) = 1300: .ColWidth(5) = 1300
         .ColWidth(6) = 1400
    End With
    
    With vsMovimiento
        .Cols = 1: .Rows = 1:
        .FormatString = "<Movimiento|<Concepto|>Debe $|>Haber $|"
            
        .WordWrap = False
        .ColWidth(0) = 3500: .ColWidth(1) = 4000: .ColWidth(2) = 1300: .ColWidth(3) = 1300
    End With
    
    With vsReval
        .ColSel = 0
        .ColSort(0) = flexSortStringAscending
        .Sort = flexSortUseColSort
        
        .Cols = 1: .Rows = 1:
        .FormatString = "Rubro Orden|<Rubro|>Debe $|>Debe M/E|>Haber $|>Haber M/E|"
            
        .WordWrap = False
        .ColWidth(0) = 0
        .ColWidth(1) = 3500: .ColWidth(2) = 1300: .ColWidth(3) = 1300: .ColWidth(4) = 1300: .ColWidth(5) = 1300
    End With
      
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            
            Case vbKeyE: AccionConsultar
            
            Case vbKeyP: IrAPagina vsListado, 1
            Case vbKeyA: IrAPagina vsListado, vsListado.PreviewPage - 1
            Case vbKeyS: IrAPagina vsListado, vsListado.PreviewPage + 1
            Case vbKeyU: IrAPagina vsListado, vsListado.PageCount
            
            Case vbKeyAdd: Zoom vsListado, vsListado.Zoom + 5
            Case vbKeySubtract: Zoom vsListado, vsListado.Zoom - 5
            
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir True
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyC: AccionConfigurar
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11
    
    fFiltros.Width = Me.ScaleWidth - (fFiltros.Left * 2)

    vsListado.Top = fFiltros.Top + fFiltros.Height + 60
    vsListado.Left = fFiltros.Left
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    vsListado.Width = fFiltros.Width
    
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 150
    
    With tabAsientos
        .Left = vsListado.Left
        .Width = vsListado.Width
        .Height = vsListado.Height
        .Top = vsListado.Top
    End With
    
    For I = picTab.LBound To picTab.UBound
        picTab(I).Top = tabAsientos.ClientTop
        picTab(I).Left = tabAsientos.ClientLeft
        picTab(I).Width = tabAsientos.ClientWidth
        picTab(I).Height = tabAsientos.ClientHeight
        picTab(I).BorderStyle = 0
    Next
    
    
    Dim Altura As Currency
    Altura = picTab(0).Height
    
    vsConsulta.Top = picTab(0).ScaleTop ' vsListado.Top
    vsConsulta.Width = picTab(0).ScaleWidth ' vsListado.Width
    vsConsulta.Left = picTab(0).ScaleLeft ' vsListado.Left
    If vsMovimiento.Visible Then vsConsulta.Height = (Altura / 5) * 4 Else vsConsulta.Height = Altura
    
    vsMovimiento.Width = vsConsulta.Width
    vsMovimiento.Left = vsConsulta.Left
    vsMovimiento.Top = vsConsulta.Top + vsConsulta.Height + 40
    vsMovimiento.Height = (Altura / 5) - 40
    
    With vsReval
        .Top = picTab(0).ScaleTop
        .Width = picTab(0).ScaleWidth
        .Left = picTab(0).ScaleLeft
        .Height = picTab(0).ScaleHeight
    End With
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub AccionConsultar()

    If Not ValidoCampos Then Exit Sub
    prmMonedaTC = 0
    
    Select Case LCase(tabAsientos.SelectedItem.Key)
        Case "comun"
                If MsgBox("Confirma realizar la consulta de 'Asientos Varios' ?", vbQuestion + vbYesNo, "Consultar Asientos Varios") = vbNo Then Exit Sub
                AccionConsultarComunes
            
        Case "revaluacion"
                If MsgBox("Confirma realizar la consulta de 'Revaluación de Ventas' ?", vbQuestion + vbYesNo, "Consultar Revaluación de Ventas") = vbNo Then Exit Sub
                 AccionConsultarReval
    End Select
    
End Sub

Private Sub AccionConsultarComunes()
Dim aSR As Long, aTipoM As Long
Dim aSROrden As String

    On Error GoTo errConsultar
    Screen.MousePointer = 11
    bCargarImpresion = True
    vsConsulta.Rows = 1: vsMovimiento.Rows = 1
    vsConsulta.Refresh: vsMovimiento.Refresh
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    cons = " Select Count(Distinct(MDiTipo)) from MovimientoDisponibilidad " _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiIDCompra Is Null "
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If rsAux(0) = 0 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            rsAux.Close: Screen.MousePointer = 0: vsConsulta.Rows = 1: Exit Sub
        End If
        pbProgreso.Max = rsAux(0)
    End If
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------------------
    
    Ayuda "Generando Asientos por Movimientos de Disponibilidades..."
    
    cons = " Select TMDNombre, MDiTipo, DH = 'Haber', Importe = Sum(MDrImportePesos), IOriginal = Sum(MDrHaber)" _
            & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiID = MDRIDMovimiento " _
            & " And MDiIDCompra Is Null " _
            & " And MDiTipo = TMDCodigo " _
            & " And MDRHaber <> Null " _
            & " Group by TMDNombre, MDiTipo"
    
    cons = cons & " UNION ALL "
    
    cons = cons & " Select TMDNombre, MDiTipo, DH = 'Debe', Importe = Sum(MDrImportePesos), IOriginal = Sum(MDrHaber)" _
                        & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = TMDCodigo " _
                        & " And MDRDebe <> Null " _
                        & " Group by TMDNombre, MDiTipo"
            
    cons = cons & " Order by TMDNombre "
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close
        Ayuda ""
        Screen.MousePointer = 0: Exit Sub
    End If
    
    vsConsulta.Redraw = False
    aTipoM = 0
    pbProgreso.Value = 0
    Do While Not rsAux.EOF
        aSR = 0
        With vsConsulta
            
            Select Case rsAux!MDiTipo
                Case paMCNotaCredito, paMCAnulacion: aSR = paSubrubroDeudoresPorVenta
                Case paMCChequeDiferido: aSR = paSubrubroCDAlCobro
                Case paMCVtaTelefonica: aSR = paSubrubroVtasTelACobrar
                Case paMCLiquidacionCamionero: aSR = paSubrubroCobranzaVtasTel
                Case paMCSenias: aSR = paSubrubroSeniasRecibidas
                
                Case Else
                    'Consulto en el movimiento si es del tipo tranferencia
                    cons = "Select * from TipoMovDisponibilidad Where TMDCodigo = " & rsAux!MDiTipo
                    Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    If Not rs1.EOF Then
                        If Not IsNull(rs1!TMDTransferencia) Then If rs1!TMDTransferencia = 1 Then aSR = -1
                    End If
                    rs1.Close
            End Select
            
            If aSR <> 0 Then
            
                If aTipoM <> rsAux!MDiTipo Then
                    pbProgreso.Value = pbProgreso.Value + 1
                    aTipoM = rsAux!MDiTipo
                    
                    .AddItem aTipoM 'Nombre del Movimiento
                    .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!TMDNombre)
                    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
                    
                    CargoConceptos aTipoM   'Hay que cargar contra que conceptos van
                    
                End If
                
                If aSR <> -1 Then       'NO ES TRANSFERENCIA
                    .AddItem ""
                    aTexto = RetornoConstanteSubrubro(aSR)
                    .Cell(flexcpText, .Rows - 1, 0) = aTipoM
                    .Cell(flexcpText, .Rows - 1, 1) = aTexto
                    
                    Select Case LCase(rsAux!DH)     'Van al reves por los tipos de movimientos
                        Case "debe"
                                    .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Importe, FormatoMonedaP)
                                    'If rsAux!IOriginal <> rsAux!Importe Then .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!IOriginal, FormatoMonedaP)
                        Case "haber":
                                    .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Importe, FormatoMonedaP)
                                    'If rsAux!IOriginal <> rsAux!Importe Then .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!IOriginal, FormatoMonedaP)
                    End Select
                End If
            Else
                If aTipoM <> rsAux!MDiTipo Then
                    pbProgreso.Value = pbProgreso.Value + 1
                    aTipoM = rsAux!MDiTipo
                    If aTipoM <> paMCIngresosOperativos Then CargoListaMovimiento aTipoM
                End If
            End If
            
            rsAux.MoveNext
        End With
    Loop
    rsAux.Close
        
    With vsConsulta
        If vsConsulta.Rows > 1 Then
            .Select 1, 0, 1, 1
            .Sort = flexSortGenericDescending
        End If
    End With
    
    EliminoAsientosDobles   'Recorro para eliminar asientos dobles--------------------------------
    
    CargoOtrosAsientos
    
    AsientoVentasContado
    AsientoVentasCredito
    AsientoMorosidadesCuotas
    AsientoNotasDevolucion
    AsientoNotasCredito

    CargoTasasDeCambio
    If vsMovimiento.Rows > 1 Then vsMovimiento.Visible = True Else vsMovimiento.Visible = False
    Call Form_Resize
    
    vsConsulta.Redraw = True
    pbProgreso.Value = 0
    Ayuda ""
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True: Screen.MousePointer = 0
    PastoSQL cons
End Sub


Private Sub CargoConceptos(aIDTipoM As Long, Optional Transferencia As Boolean = False)

    cons = " Select SRuCodigo, SRuNombre, SRuID, DH = 'Haber', Importe = Sum(MDrImportePesos), IOriginal = Sum(MDRHaber)"
    cons = cons & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, Disponibilidad, SubRubro " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = " & aIDTipoM _
                        & " And MDRIDDisponibilidad = DisID " _
                        & " And DisIDSubrubro = SRuID" _
                        & " And MDRHaber <> Null " _
                        & " Group by SRuCodigo, SRuNombre, SRuID"
    
    cons = cons & " UNION ALL "
    
    cons = cons & " Select SRuCodigo, SRuNombre, SRuID, DH = 'Debe', Importe = Sum(MDrImportePesos), IOriginal = Sum(MDRDebe) "
    cons = cons & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, Disponibilidad, SubRubro " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = " & aIDTipoM _
                        & " And MDRIDDisponibilidad = DisID " _
                        & " And DisIDSubrubro = SRuID" _
                        & " And MDRDebe <> Null " _
                        & " Group by SRuCodigo, SRuNombre, SRuID"
            
    Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rs1.EOF
        
        With vsConsulta
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = aIDTipoM
             aTexto = rs1!SRuID: .Cell(flexcpData, .Rows - 1, 1) = aTexto   '8/5/2003
             aTexto = Format(rs1!SRuCodigo, "000000000") & " " & Trim(rs1!SRuNombre)
            .Cell(flexcpText, .Rows - 1, 1) = aTexto
            
            
            Select Case LCase(rs1!DH)
                Case "debe"
                    .Cell(flexcpText, .Rows - 1, 2) = Format(rs1!Importe, FormatoMonedaP)
                    If rs1!IOriginal <> rs1!Importe Then .Cell(flexcpText, .Rows - 1, 3) = Format(rs1!IOriginal, FormatoMonedaP)
                
                Case "haber"
                    .Cell(flexcpText, .Rows - 1, 4) = Format(rs1!Importe, FormatoMonedaP)
                    If rs1!IOriginal <> rs1!Importe Then .Cell(flexcpText, .Rows - 1, 5) = Format(rs1!IOriginal, FormatoMonedaP)
            End Select
            
            'If rs1!Importe <> rs1!IOriginal Then .Cell(flexcpText, .Rows - 1, 6) = Format(rs1!Importe / rs1!IOriginal, "#,##0.000")
        End With
        
        rs1.MoveNext
    Loop
    rs1.Close
    
End Sub

Private Function AsientoVentasContado()

Dim rsSuc As rdoResultset
Dim rsData As rdoResultset
Dim mSubRubro As String
Dim idx As Integer

Dim mSucursal As Long
Dim mIDDisponibilidad As Long
Dim mTotal As Currency, mIva As Currency, mCofis As Currency
Dim auxTotal As Currency, auxIva As Currency, auxCofis As Currency
Dim mTotalME As Currency, mIvaME As Currency, mCofisME As Currency
Dim auxTC As Currency

Dim mIDMoneda As Long, mNameMoneda As String
Dim rsMon As rdoResultset

    ReDim arrAsientos(0)
    idx = 0
    
    cons = "Select * from Moneda Where MonCoeficienteMora Is Not Null "
    Set rsMon = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsMon.EOF
        mNameMoneda = Trim(rsMon!MonNombre)
        mIDMoneda = rsMon!MonCodigo
        mTotal = 0: mIva = 0: mCofis = 0
        mTotalME = 0: mIvaME = 0: mCofisME = 0
        
        '1)  Consulto Todas las Sucursales POR MONEDA   --------------------------------------
        If mIDMoneda = paMonedaPesos Then
            cons = "Select * from Sucursal Where SucDisponibilidad Is Not Null"
        Else
            cons = "Select * from Sucursal Where SucDisponibilidadME Is Not Null"
        End If
        cons = cons & " And SucICoNombre Is Not Null"
    
        Set rsSuc = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not rsSuc.EOF
            mSucursal = rsSuc!SucCodigo
            Ayuda "Cargando Ventas Contado " & mNameMoneda & " (" & Trim(rsSuc!SucAbreviacion) & ")..."
            
            '2) CONSULTO VENTAS POR SUCURSAL    -----------------------------------------------
            cons = "Select Datepart(dd, DocFecha) as DocFecha, Sum(DocTotal) Total, Sum(DocIva) Iva, Sum(DocCofis) Cofis " _
                    & " From Documento (index = iTipoFechaSucursalMoneda) " _
                    & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                    & " And DocTipo = " & TipoDocumento.Contado _
                    & " And DocAnulado = 0" _
                    & " And DocMoneda = " & mIDMoneda _
                    & " And DocSucursal = " & mSucursal _
                    & " Group by Datepart(dd, DocFecha)"
                    
            Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rsAux.EOF Then
            
                If mIDMoneda <> paMonedaPesos Then
                    meCargoTasasDeCambio mIDMoneda
                    Ayuda "Cargando Ventas Contado " & mNameMoneda & " (" & Trim(rsSuc!SucAbreviacion) & ")..."
                End If
                
                mIDDisponibilidad = dis_DisponibilidadPara(mSucursal, mIDMoneda)
                mSubRubro = meSubroParaDisponibilidad(mIDDisponibilidad, mNameMoneda)
                
                Do While Not rsAux.EOF
                
                    If Not IsNull(rsAux!Total) Then
                        If rsAux!Total <> 0 Then
                    
                            auxTotal = 0: auxIva = 0: auxCofis = 0
                            auxTC = 1
                            If mIDMoneda <> paMonedaPesos Then auxTC = meTCaFecha(CDate(rsAux!DocFecha))
                            
                            If Not IsNull(rsAux!Cofis) Then auxCofis = Format(rsAux!Cofis * auxTC, FormatoMonedaP)
                            auxTotal = Format(rsAux!Total * auxTC, FormatoMonedaP)
                            auxIva = Format(rsAux!Iva * auxTC, FormatoMonedaP)
                            
                            mCofis = mCofis + auxCofis
                            mTotal = mTotal + auxTotal
                            mIva = mIva + auxIva
                            
                            If mIDMoneda <> paMonedaPesos Then
                                If Not IsNull(rsAux!Cofis) Then mCofisME = mCofisME + Format(rsAux!Cofis, FormatoMonedaP)
                                mTotalME = mTotalME + Format(rsAux!Total, FormatoMonedaP)
                                mIvaME = mIvaME + Format(rsAux!Iva, FormatoMonedaP)
                            End If
                            
                            'Si el Subrubro está en el array --> Acumulo los datos, si no lo agrego --------------------------------------
                            meAddarrAsientos mSubRubro, auxTotal
                        End If
                    End If
                    rsAux.MoveNext
                Loop
            End If
            rsAux.Close
            '2) FIN CONSULTA POR SUCURSAL   ------------------------------------------------------
            
            rsSuc.MoveNext
        Loop
        rsSuc.Close
        '1) FIN consulta de Sucursales POR MONEDA   ------------------------------------------
        
        If mTotal <> 0 Then     'ASIENTO DE VENTAS CONTADO EN M/E
            With vsConsulta
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = "Ventas Contado (" & mNameMoneda & ")"
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
                
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
                .Cell(flexcpText, .Rows - 1, 2) = Format(mTotal, FormatoMonedaP)
                If mIDMoneda <> paMonedaPesos Then .Cell(flexcpText, .Rows - 1, 3) = Format(mTotalME, FormatoMonedaP)
                
                .AddItem ""
                If mIDMoneda = paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentas)
                Else
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentasME)
                    .Cell(flexcpText, .Rows - 1, 5) = Format(mTotalME - mIvaME - mCofisME, FormatoMonedaP)
                End If
                .Cell(flexcpText, .Rows - 1, 4) = Format(mTotal - mIva - mCofis, FormatoMonedaP)
                
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubRubroCofis)
                .Cell(flexcpText, .Rows - 1, 4) = Format(mCofis, FormatoMonedaP)
                If mIDMoneda <> paMonedaPesos Then .Cell(flexcpText, .Rows - 1, 5) = Format(mCofisME, FormatoMonedaP)
                
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
                .Cell(flexcpText, .Rows - 1, 4) = Format(mIva, FormatoMonedaP)
                If mIDMoneda <> paMonedaPesos Then .Cell(flexcpText, .Rows - 1, 5) = Format(mIvaME, FormatoMonedaP)
            End With
        End If
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------
        rsMon.MoveNext
    Loop
    rsMon.Close
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
    With vsConsulta         'ASIENTO RESUMEN DE VENTAS CONTADO (TODAS LAS MONEDAS CONTRA VENTAS)
        If Trim(arrAsientos(0).NombreRubro) <> "" Then
            Dim mTotalAsiento As Currency
            mTotalAsiento = 0
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = "Ventas Contado"
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            For I = LBound(arrAsientos) To UBound(arrAsientos)
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = arrAsientos(I).NombreRubro
                .Cell(flexcpText, .Rows - 1, 2) = Format(arrAsientos(I).Importe, FormatoMonedaP)
                mTotalAsiento = mTotalAsiento + arrAsientos(I).Importe
            Next
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
            .Cell(flexcpText, .Rows - 1, 4) = Format(mTotalAsiento, FormatoMonedaP)
        End If
    End With
    
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
End Function

Private Sub AsientoVentasCredito()

Dim mTotal As Currency, mIva As Currency, mCofis As Currency
Dim auxTotal As Currency, auxIva As Currency, auxCofis As Currency
Dim auxTC As Currency
Dim mTotalME As Currency, mIvaME As Currency, mCofisME As Currency
    
Dim rsMon As rdoResultset
Dim mName As String, mCodigo As Long
    
    cons = "Select * from Moneda Where MonCoeficienteMora Is Not Null "
    Set rsMon = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsMon.EOF
        mName = Trim(rsMon!MonNombre)
        mCodigo = rsMon!MonCodigo
        
        mTotal = 0: mIva = 0: mCofis = 0
        mTotalME = 0: mIvaME = 0: mCofisME = 0
        Ayuda "Cargando Ventas Crédito " & mName & "..."
        
        'ASIENTO DE VENTAS CREDITO x MONEDA --------------------------------------------------------------------------------------------------------
        cons = "Select Datepart(dd, DocFecha) as DocFecha, Sum(DocTotal) Total, Sum(DocIva) IVA, Sum(DocCofis) Cofis " _
                & " From Documento (index = iTipoFechaSucursalMoneda) " _
                & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                & " And DocTipo = " & TipoDocumento.Credito _
                & " And DocAnulado = 0" _
                & " And DocMoneda = " & mCodigo _
                & " Group by Datepart(dd, DocFecha)"

        Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        If Not rsAux.EOF Then
            
            If mCodigo <> paMonedaPesos Then
                meCargoTasasDeCambio mCodigo
                Ayuda "Cargando Ventas Crédito " & mName & "..."
            End If
            
            Do While Not rsAux.EOF
                If Not IsNull(rsAux!Total) Then
                    If rsAux!Total <> 0 Then
            
                        auxTC = 1
                        auxTotal = 0: auxIva = 0: auxCofis = 0
                    
                        If mCodigo <> paMonedaPesos Then
                            auxTC = meTCaFecha(CDate(rsAux!DocFecha))
                        End If
                    
                        If Not IsNull(rsAux!Cofis) Then auxCofis = rsAux!Cofis * auxTC
                        'auxTotal = Format(rsAux!Total * auxTC, FormatoMonedaP)
                        'auxIva = Format(rsAux!Iva * auxTC, FormatoMonedaP)
                        auxTotal = rsAux!Total * auxTC
                        auxIva = rsAux!Iva * auxTC
                        
                        mCofis = mCofis + auxCofis
                        mTotal = mTotal + auxTotal
                        mIva = mIva + auxIva
                        
                        If mCodigo <> paMonedaPesos Then
                            If Not IsNull(rsAux!Cofis) Then mCofisME = mCofisME + Format(rsAux!Cofis, FormatoMonedaP)
                            mTotalME = mTotalME + Format(rsAux!Total, FormatoMonedaP)
                            mIvaME = mIvaME + Format(rsAux!Iva, FormatoMonedaP)
                        End If
                        
                    End If
                End If
                rsAux.MoveNext
            Loop
            
        End If
        rsAux.Close
        
        'ASIENTO DE VENTAS CREDITO ------------------------------------------------------------------------------------------------------------------
        If mTotal <> 0 Then
            With vsConsulta
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = "Ventas Crédito (" & mName & ")"
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
                
                .AddItem ""
                If mCodigo = paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVenta)
                Else
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVentaME)
                    .Cell(flexcpText, .Rows - 1, 5) = Format(mTotalME, FormatoMonedaP)
                End If
                .Cell(flexcpText, .Rows - 1, 2) = Format(mTotal, FormatoMonedaP)
            
                .AddItem ""
                If mCodigo = paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentas)
                Else
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentasME)
                    .Cell(flexcpText, .Rows - 1, 5) = Format(mTotalME - mIvaME - mCofisME, FormatoMonedaP)
                End If
                .Cell(flexcpText, .Rows - 1, 4) = Format(mTotal - mIva - mCofis, FormatoMonedaP)
                
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubRubroCofis)
                .Cell(flexcpText, .Rows - 1, 4) = Format(mCofis, FormatoMonedaP)
                If mCodigo <> paMonedaPesos Then .Cell(flexcpText, .Rows - 1, 5) = Format(mCofisME, FormatoMonedaP)
                
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
                .Cell(flexcpText, .Rows - 1, 4) = Format(mIva, FormatoMonedaP)
                If mCodigo <> paMonedaPesos Then .Cell(flexcpText, .Rows - 1, 5) = Format(mIvaME, FormatoMonedaP)
            End With
        End If
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
        rsMon.MoveNext
    Loop
    rsMon.Close
    
End Sub

Private Sub AsientoNotasDevolucion()

Dim rsMon As rdoResultset, rsSuc As rdoResultset

Dim mIDMoneda As Long, mNameMoneda As String
Dim mSucursal As Long, mIDDisponibilidad As Long
Dim mSubRubro As String

Dim auxTC As Currency
Dim mTotalMoneda As Currency, mIvaMoneda As Currency, mCofisMoneda As Currency
Dim mTotalSucursal As Currency
Dim mFechaF As Date

    ReDim arrAsientos(0)
           
    cons = "Select * from Moneda Where MonCoeficienteMora Is Not Null"
    Set rsMon = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsMon.EOF
        mNameMoneda = Trim(rsMon!MonNombre)
        mIDMoneda = rsMon!MonCodigo
        mTotalMoneda = 0: mIvaMoneda = 0: mCofisMoneda = 0
        
        '1)  Consulto Todas las Sucursales POR MONEDA   --------------------------------------
        If mIDMoneda = paMonedaPesos Then
            cons = "Select * from Sucursal Where SucDisponibilidad Is Not Null"
        Else
            cons = "Select * from Sucursal Where SucDisponibilidadME Is Not Null"
        End If
        cons = cons & " And (SucDNDevolucion Is Not Null Or SucDNEspecial Is Not Null)"
        
        Set rsSuc = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not rsSuc.EOF
            mTotalSucursal = 0
            mSucursal = rsSuc!SucCodigo
            mIDDisponibilidad = dis_DisponibilidadPara(mSucursal, mIDMoneda)
            mSubRubro = meSubroParaDisponibilidad(mIDDisponibilidad, mNameMoneda)
            
            If mIDMoneda = paMonedaPesos Then
                cons = "Select Sum(DocTotal) Total, Sum(DocIva) IVA, Sum(DocCofis) Cofis " _
                        & " From Documento (index = iTipoFechaSucursalMoneda) " _
                        & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                        & " And DocTipo IN (" & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")" _
                        & " And DocAnulado = 0" _
                        & " And DocMoneda = " & mIDMoneda _
                        & " And DocSucursal = " & mSucursal
            Else
                
                cons = " Select Documento.DocFecha as DocFecha, Documento.DocTotal as Total, Documento.DocIva as Iva, Documento.DocCofis as Cofis, " & _
                                     " Factura.DocFecha as FFactura" & _
                           " From Documento(Index = iTipoFechaSucursalMoneda)" & _
                                    " Left Outer Join Nota On Documento.DocCodigo = NotNota " & _
                                    " Left Outer Join Documento Factura On Factura.DocCodigo = NotFactura " & _
                           " Where Documento.DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" & _
                           " And Documento.DocTipo IN (" & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")" & _
                           " And Documento.DocAnulado = 0" & _
                           " And Documento.DocMoneda = " & mIDMoneda & _
                           " And Documento.DocSucursal = " & mSucursal
            End If
            Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rsAux.EOF Then
                Ayuda "Cargando Notas Devolución " & mNameMoneda & " (" & Trim(rsSuc!SucAbreviacion) & ")..."
                                
                Do While Not rsAux.EOF
                    auxTC = 1
                    If mIDMoneda <> paMonedaPesos Then
                        If Not IsNull(rsAux!FFactura) Then
                            mFechaF = rsAux!FFactura - 1
                        Else
                            mFechaF = rsAux!DocFecha - 1
                        End If
                        auxTC = TasadeCambio(CInt(mIDMoneda), paMonedaPesos, mFechaF, TipoTC:=prmTCVentasME)
                    End If
                    
                    If Not IsNull(rsAux!Total) Then
                        mTotalSucursal = mTotalSucursal + (rsAux!Total * auxTC)
                        mTotalMoneda = mTotalMoneda + (rsAux!Total * auxTC)
                    End If
                    If Not IsNull(rsAux!Iva) Then mIvaMoneda = mIvaMoneda + (rsAux!Iva * auxTC)
                    If Not IsNull(rsAux!Cofis) Then mCofisMoneda = mCofisMoneda + (rsAux!Cofis * auxTC)
                    
                    rsAux.MoveNext
                Loop
            End If
            rsAux.Close
            '---------------------------------------------------------------------------------------------------------------------------------
                
            If mTotalSucursal <> 0 Then meAddarrAsientos mSubRubro, mTotalSucursal
            rsSuc.MoveNext
        Loop
        rsSuc.Close
        
        If mTotalMoneda <> 0 Then
            With vsConsulta
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = "Notas de Devolución y Especiales (" & mNameMoneda & ")"
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
                
                .AddItem ""
                If mIDMoneda = paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentas)
                Else
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentasME)
                End If
                .Cell(flexcpText, .Rows - 1, 2) = Format(mTotalMoneda - mIvaMoneda - mCofisMoneda, FormatoMonedaP)
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubRubroCofis)
                .Cell(flexcpText, .Rows - 1, 2) = Format(mCofisMoneda, FormatoMonedaP)
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
                .Cell(flexcpText, .Rows - 1, 2) = Format(mIvaMoneda, FormatoMonedaP)
                        
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
                .Cell(flexcpText, .Rows - 1, 4) = Format(mTotalMoneda, FormatoMonedaP)
            End With
        End If
        
        rsMon.MoveNext
    Loop
    rsMon.Close
    
    'Contra Asiento de NOTAS DE DEV. Y ESPECIALES   ------------------------------------------------------
    If Trim(arrAsientos(0).NombreRubro) <> "" Then
    With vsConsulta
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = "Notas de Devolución y Notas Especiales"
        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
        
        Dim mTotalAsiento As Currency: mTotalAsiento = 0
        For I = LBound(arrAsientos) To UBound(arrAsientos)
            mTotalAsiento = mTotalAsiento + arrAsientos(I).Importe
        Next
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
        .Cell(flexcpText, .Rows - 1, 2) = Format(mTotalAsiento, FormatoMonedaP)
        
        For I = LBound(arrAsientos) To UBound(arrAsientos)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = arrAsientos(I).NombreRubro
            .Cell(flexcpText, .Rows - 1, 4) = Format(arrAsientos(I).Importe, FormatoMonedaP)
        Next
    End With
    End If
    
End Sub

Private Sub AsientoNotasCredito()

Dim rsMon As rdoResultset, rsSuc As rdoResultset

Dim mIDMoneda As Long, mNameMoneda As String
Dim mSucursal As Long, mIDDisponibilidad As Long
Dim mSubRubro As String

Dim auxTC As Currency
Dim mTotalMoneda As Currency, mIvaMoneda As Currency, mCofisMoneda As Currency
Dim mTotalSucursal As Currency
Dim mFechaF As Date

    cons = "Select * from Moneda Where MonCoeficienteMora Is Not Null"
    Set rsMon = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsMon.EOF
        mNameMoneda = Trim(rsMon!MonNombre)
        mIDMoneda = rsMon!MonCodigo
        mTotalMoneda = 0: mIvaMoneda = 0: mCofisMoneda = 0
        
        If mIDMoneda = paMonedaPesos Then
            
            cons = "Select Sum(DocTotal) Total, Sum(DocIva) IVA,  Sum(DocCofis) Cofis " _
                    & " From documento (index = iTipoFechaSucursalMoneda) " _
                    & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                    & " And DocTipo = " & TipoDocumento.NotaCredito _
                    & " And DocAnulado = 0" _
                    & " And DocMoneda = " & mIDMoneda
                    
        Else
            cons = "Select DocFecha, DocTotal as Total, DocIva as Iva, DocCofis as Cofis, CreUltimoPago" & _
                       " From Documento(Index = iTipoFechaSucursalMoneda) " & _
                                " Left Outer Join Nota On DocCodigo = NotNota " & _
                                " Left Outer Join Credito On NotFactura = CreFactura " & _
                       " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" & _
                       " And DocTipo = " & TipoDocumento.NotaCredito & _
                       " And DocAnulado = 0" & _
                       " And DocMoneda = " & mIDMoneda
        End If
        
        Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        If Not rsAux.EOF Then
            Ayuda "Cargando Notas Crédito " & mNameMoneda & " ..."
                                
            Do While Not rsAux.EOF
                auxTC = 1
                If mIDMoneda <> paMonedaPesos Then
                    If Not IsNull(rsAux!CreUltimoPago) Then
                        mFechaF = rsAux!CreUltimoPago - 1
                    Else
                        mFechaF = rsAux!DocFecha - 1
                    End If
                    auxTC = TasadeCambio(CInt(mIDMoneda), paMonedaPesos, mFechaF, TipoTC:=prmTCVentasME)
                End If
                
                If Not IsNull(rsAux!Total) Then mTotalMoneda = mTotalMoneda + (rsAux!Total * auxTC)
                If Not IsNull(rsAux!Iva) Then mIvaMoneda = mIvaMoneda + (rsAux!Iva * auxTC)
                If Not IsNull(rsAux!Cofis) Then mCofisMoneda = mCofisMoneda + (rsAux!Cofis * auxTC)
                
                rsAux.MoveNext
            Loop
        End If
        rsAux.Close
        '---------------------------------------------------------------------------------------------------------------------------------
                
        If mTotalMoneda <> 0 Then
            With vsConsulta
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = "Notas de Crédito (" & mNameMoneda & ")"
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
                
                .AddItem ""
                If mIDMoneda = paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentas)
                Else
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentasME)
                End If
                .Cell(flexcpText, .Rows - 1, 2) = Format(mTotalMoneda - mIvaMoneda - mCofisMoneda, FormatoMonedaP)
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubRubroCofis)
                .Cell(flexcpText, .Rows - 1, 2) = Format(mCofisMoneda, FormatoMonedaP)
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
                .Cell(flexcpText, .Rows - 1, 2) = Format(mIvaMoneda, FormatoMonedaP)
                        
                .AddItem ""
                If mIDMoneda = paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVenta)
                Else
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVentaME)
                End If
                .Cell(flexcpText, .Rows - 1, 4) = Format(mTotalMoneda, FormatoMonedaP)
                
            End With
        End If
        
        rsMon.MoveNext
    Loop
    rsMon.Close
    
End Sub

Private Sub AsientoMorosidadesCuotas()

Dim rsMon As rdoResultset, rsSuc As rdoResultset

Dim mIDMoneda As Long, mNameMoneda As String
Dim mSucursal As Long
Dim mIDDisponibilidad As Long
Dim mSubRubro As String

Dim mMoraSucursal As Currency, mMoraMoneda As Currency
Dim mIvaMorasTotal As Currency
Dim auxTC As Currency

Dim mCtasxSucursal As Currency, mCtasMoneda As Currency
Dim mSenasxMoneda As Currency, mSenasxMonedaME As Currency

    mIvaMorasTotal = 0
    ReDim arrAsientos(0)
           
    cons = "Select * from Moneda Where MonCoeficienteMora Is Not Null"
    Set rsMon = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsMon.EOF
        mNameMoneda = Trim(rsMon!MonNombre)
        mIDMoneda = rsMon!MonCodigo
        
        mMoraMoneda = 0
        mCtasMoneda = 0: mSenasxMoneda = 0: mSenasxMonedaME = 0
        ReDim arrAsientosCtas(0)
        
        '1)  Consulto Todas las Sucursales POR MONEDA   --------------------------------------
        If mIDMoneda = paMonedaPesos Then
            cons = "Select * from Sucursal Where SucDisponibilidad Is Not Null"
        Else
            cons = "Select * from Sucursal Where SucDisponibilidadME Is Not Null"
        End If
        cons = cons & " And SucDRecibo Is Not Null"
        
        Set rsSuc = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not rsSuc.EOF
            
            mSucursal = rsSuc!SucCodigo
            mIDDisponibilidad = dis_DisponibilidadPara(mSucursal, mIDMoneda)
            mSubRubro = meSubroParaDisponibilidad(mIDDisponibilidad, mNameMoneda)
            
            mMoraSucursal = 0
            mCtasxSucursal = 0
            
            'Saco lo que se cobró por Mora ---------------------------------------------------------------------------
            cons = "Select Datepart(dd, DocFecha) as DocFecha, Sum(DPaMora) as Total, 0 as IVA" _
                    & " From Documento (index = iTipoFechaSucursalMoneda), DocumentoPago " _
                    & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                    & " And DocTipo = " & TipoDocumento.ReciboDePago _
                    & " And DocAnulado = 0 And DocCodigo = DPaDocQSalda and DPaMora <> 0" _
                    & " And DocSucursal = " & mSucursal & " And DocMoneda = " & mIDMoneda _
                    & " Group by Datepart(dd, DocFecha)"
            
            cons = cons & " UNION ALL " & _
                        " Select Datepart(dd, DocFecha) as DocFecha,  0 as Total, Sum(DocIVA) AS IVA " & _
            " From Documento(Index = iTipoFechaSucursalMoneda) " & _
            " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" & _
            " And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.NotaDebito & _
            " And DocSucursal = " & mSucursal & " And DocMoneda = " & mIDMoneda & _
            " Group by Datepart(dd, DocFecha)"

    
            Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rsAux.EOF Then
    
                If mIDMoneda <> paMonedaPesos Then meCargoTasasDeCambio mIDMoneda
                Ayuda "Cargando Morosidades en " & mNameMoneda & " (" & Trim(rsSuc!SucAbreviacion) & ")..."
                
                Do While Not rsAux.EOF
                    auxTC = 1
                    If mIDMoneda <> paMonedaPesos Then auxTC = meTCaFecha(CDate(rsAux!DocFecha))
                    If Not IsNull(rsAux!Total) Then
                        mMoraSucursal = mMoraSucursal + (rsAux!Total * auxTC)
                        mCtasMoneda = mCtasMoneda - rsAux!Total
                    End If
                    If Not IsNull(rsAux!Iva) Then mIvaMorasTotal = mIvaMorasTotal + (rsAux!Iva * auxTC)

                    rsAux.MoveNext
                Loop
            End If
            rsAux.Close
            
             'A la Mora, Le sumo la cobranza de creditos a Perdida (Amortizacion --> La Mora ya esta)
            cons = "Select Datepart(dd, DocFecha) as DocFecha, Sum(DPAAmortizacion) Suma " _
                    & " From Documento (index = iTipoFechaSucursalMoneda), DocumentoPago, Credito " _
                    & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                    & " And DocTipo = " & TipoDocumento.ReciboDePago _
                    & " And DocAnulado = 0" _
                    & " And DocCodigo = DPaDocQSalda And DPaDocASaldar = CreFactura" _
                    & " And CreTipo = " & TipoCredito.Incobrable _
                    & " And DocSucursal = " & mSucursal _
                    & " And DocMoneda = " & mIDMoneda _
                    & " Group by Datepart(dd, DocFecha)"
            Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rsAux.EOF Then
                
                If mIDMoneda <> paMonedaPesos Then meCargoTasasDeCambio mIDMoneda
                Ayuda "Cargando Cob. Créditos A Pérdida " & mNameMoneda & " (" & Trim(rsSuc!SucAbreviacion) & ")..."
                
                Do While Not rsAux.EOF
                    auxTC = 1
                    If mIDMoneda <> paMonedaPesos Then auxTC = meTCaFecha(CDate(rsAux!DocFecha))
                    If Not IsNull(rsAux!Suma) Then
                        mMoraSucursal = mMoraSucursal + (rsAux!Suma * auxTC)
                        mCtasMoneda = mCtasMoneda - rsAux!Suma
                    End If
                    rsAux.MoveNext
                Loop
            End If
            rsAux.Close
            '---------------------------------------------------------------------------------------------------------------------------------
            
            If mMoraSucursal <> 0 Then meAddarrAsientos mSubRubro, mMoraSucursal
            
            mMoraMoneda = mMoraMoneda + mMoraSucursal
            
            'Saco el importe por la Cobraza de Cuotas de la Sucursal        --------------------------------------------------
            cons = "Select Datepart(dd, DocFecha) as DocFecha, Sum(DocTotal) as Total, Sum(DocIva) as Iva " _
                    & " From Documento (index = iTipoFechaSucursalMoneda)" _
                    & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                    & " And DocTipo = " & TipoDocumento.ReciboDePago _
                    & " And DocAnulado = 0" _
                    & " And DocMoneda = " & mIDMoneda _
                    & " And DocSucursal = " & mSucursal _
                    & " Group by Datepart(dd, DocFecha)"
            Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rsAux.EOF Then
                
                If mIDMoneda <> paMonedaPesos Then meCargoTasasDeCambio mIDMoneda
                Ayuda "Cargando Cobranza de Cuotas " & mNameMoneda & " (" & Trim(rsSuc!SucAbreviacion) & ")..."
                                
                Do While Not rsAux.EOF
                    auxTC = 1
                    If mIDMoneda <> paMonedaPesos Then auxTC = meTCaFecha(CDate(rsAux!DocFecha))
                    If Not IsNull(rsAux!Total) Then
                        mCtasxSucursal = mCtasxSucursal + (rsAux!Total * auxTC)
                        mCtasMoneda = mCtasMoneda + rsAux!Total
                    End If
                    
                    If Not IsNull(rsAux!Iva) Then mIvaMorasTotal = mIvaMorasTotal + (rsAux!Iva * auxTC)
                    
                    rsAux.MoveNext
                Loop
            End If
            rsAux.Close
            
            If mCtasxSucursal <> 0 Then
                mCtasxSucursal = mCtasxSucursal - mMoraSucursal
                meAddarrAsientosCtas mSubRubro, mCtasxSucursal
            End If
            '---------------------------------------------------------------------------------------------------------------------------------
            
            'Saco el importe cobrado por cuotas que son SEÑAS RECIBIDAS
            Dim aSenias As Currency: aSenias = 0
            cons = "Select Datepart(dd, DocFecha) as DocFecha, Sum(DocTotal) as Total,  Sum(DocIva) as Iva " _
                    & " From Documento (index = iTipoFechaSucursalMoneda) " _
                                        & " Left Outer Join DocumentoPago On DocCodigo = DPaDocQSalda " _
                    & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                    & " And DocTipo = " & TipoDocumento.ReciboDePago _
                    & " And DocAnulado = 0 And DPaDocQSalda is null And DPaDocASaldar is null" _
                    & " And DocMoneda = " & mIDMoneda _
                    & " And DocSucursal = " & mSucursal _
                    & " Group by Datepart(dd, DocFecha)"
            Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rsAux.EOF Then
                If mIDMoneda <> paMonedaPesos Then meCargoTasasDeCambio mIDMoneda
                Ayuda "Cargando Señas Recibidas " & mNameMoneda & " (" & Trim(rsSuc!SucAbreviacion) & ")..."
                
                Do While Not rsAux.EOF
                    auxTC = 1
                    If mIDMoneda <> paMonedaPesos Then auxTC = meTCaFecha(CDate(rsAux!DocFecha))
                    If Not IsNull(rsAux!Total) Then
                        mSenasxMoneda = mSenasxMoneda + (rsAux!Total * auxTC)
                        mSenasxMonedaME = mSenasxMonedaME + rsAux!Total
                    End If
                    
                    If Not IsNull(rsAux!Iva) Then mIvaMorasTotal = mIvaMorasTotal - (rsAux!Iva * auxTC)
                    
                    rsAux.MoveNext
                Loop
            End If
            rsAux.Close
            '---------------------------------------------------------------------------------------------------------------------------------
            
            rsSuc.MoveNext
        Loop
        rsSuc.Close
        
        If arrAsientosCtas(0).NombreRubro <> "" Then
        With vsConsulta     'ASIENTO DE COBRANZA DE CUOTAS POR MONEDA -----------------------------------------
        
            Dim mTotalAsientoCtas As Currency: mTotalAsientoCtas = 0
                       
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = "Cobranza de Cuotas (" & Trim(mNameMoneda) & ")"
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            For I = LBound(arrAsientosCtas) To UBound(arrAsientosCtas)
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = arrAsientosCtas(I).NombreRubro
                .Cell(flexcpText, .Rows - 1, 2) = Format(arrAsientosCtas(I).Importe, FormatoMonedaP)
                mTotalAsientoCtas = mTotalAsientoCtas + arrAsientosCtas(I).Importe
            Next
            
            If mTotalAsientoCtas - mSenasxMoneda > 0 Then   'hay pago de ctas
                .AddItem ""
                If mIDMoneda = paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVenta)
                Else
                    .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVentaME)
                End If
                .Cell(flexcpText, .Rows - 1, 4) = Format(mTotalAsientoCtas - mSenasxMoneda, FormatoMonedaP)
                If mIDMoneda <> paMonedaPesos Then .Cell(flexcpText, .Rows - 1, 5) = Format(mCtasMoneda - mSenasxMonedaME, FormatoMonedaP)
            End If
            
            If mSenasxMoneda <> 0 Then    'hay senias
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroSeniasRecibidas)
                .Cell(flexcpText, .Rows - 1, 4) = Format(mSenasxMoneda, FormatoMonedaP)
                If mIDMoneda <> paMonedaPesos Then .Cell(flexcpText, .Rows - 1, 5) = Format(mSenasxMonedaME, FormatoMonedaP)
            End If
        End With    '----------------------------------------------------------------------------------------------------------------
        End If
        rsMon.MoveNext
    Loop
    rsMon.Close
    
    If mIvaMorasTotal <> 0 Then
        Dim mTotalAsiento As Currency
        mTotalAsiento = 0
        'ASIENTO COBRANZA DE MOROSIDADES -----------------------
        With vsConsulta
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = "Cobranza de Morosidades"
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            For I = LBound(arrAsientos) To UBound(arrAsientos)
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = arrAsientos(I).NombreRubro
                .Cell(flexcpText, .Rows - 1, 2) = Format(arrAsientos(I).Importe, FormatoMonedaP)
                mTotalAsiento = mTotalAsiento + arrAsientos(I).Importe
            Next
        
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIngresosVarios)
            .Cell(flexcpText, .Rows - 1, 4) = Format(mTotalAsiento - mIvaMorasTotal, FormatoMonedaP)   'En recibos el IVA es sobre la MORA
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
            .Cell(flexcpText, .Rows - 1, 4) = Format(mIvaMorasTotal, FormatoMonedaP)
        End With
    End If

End Sub


Private Sub CargoOtrosAsientos()
    
    Dim aTC As Currency
            
    'ASIENTO DE CHEQUES DIF A PAGAR (LOS Q VENCIERON)------------------------------------------------------------------------------------------------
    'Cargo los Rubros de la Dispnibilidad con los Cheuqes Diferidos a Pagar (Van al reves por q son la contrapartida de los siguientes)
    Ayuda "Cargando Asientos de Cheques Diferidos a Pagar..."
    
    cons = "Select DisMoneda, SRuCodigo, SRuNombre, Importe = Sum(MDRImportePesos), DH = 'Haber', IOriginal = Sum(MDRHaber) " _
            & " From Cheque, MovimientoDisponibilidadRenglon, Disponibilidad, SubRubro" _
            & " Where CheID = MDRIDCheque " _
            & " And MDRIdDisponibilidad = DisID " _
            & " And DisIDSRCheque = SRuID " _
            & " And CheVencimiento Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDRHaber Is Not Null " _
            & " Group by DisMoneda, SRuCodigo, SRuNombre" _
                    & " UNION ALL " _
            & " Select DisMoneda, SRuCodigo, SRuNombre, Importe = Sum(MDRImportePesos), DH = 'Debe', IOriginal = Sum(MDRDebe)" _
            & " From Cheque, MovimientoDisponibilidadRenglon, Disponibilidad, SubRubro" _
            & " Where CheID = MDRIDCheque " _
            & " And MDRIdDisponibilidad = DisID " _
            & " And DisIDSRCheque = SRuID " _
            & " And CheVencimiento Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDRDebe Is Not Null " _
            & " Group by DisMoneda, SRuCodigo, SRuNombre"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        With vsConsulta
        .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Vto. de Cheques Diferidos a Pagar": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
        Do While Not rsAux.EOF
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!SRuCodigo, "000000000") & " " & Trim(rsAux!SRuNombre)
    
            'TC del ultimo dia del mes anterior
            aTC = TasadeCambio(rsAux!DisMoneda, paMonedaPesos, UltimoDia(DateAdd("m", -1, CDate(tFecha.Text))))
            
            Select Case LCase(rsAux!DH)
                Case "haber"
                    .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Importe, FormatoMonedaP)
                    If rsAux!IOriginal <> rsAux!Importe Then
                        .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!IOriginal, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!IOriginal * aTC, FormatoMonedaP)
                    End If
                
                Case "debe"
                    .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Importe, FormatoMonedaP)
                    If rsAux!IOriginal <> rsAux!Importe Then
                        .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!IOriginal, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!IOriginal * aTC, FormatoMonedaP)
                    End If
            End Select
            rsAux.MoveNext
        Loop
        End With
    End If
    rsAux.Close
    
    cons = "Select DisMoneda, SRuCodigo, SRuNombre, Importe = Sum(MDRImportePesos), DH = 'Haber', IOriginal = Sum(MDRHaber) " _
            & " From Cheque, MovimientoDisponibilidadRenglon, Disponibilidad, SubRubro" _
            & " Where CheID = MDRIDCheque " _
            & " And MDRIdDisponibilidad = DisID " _
            & " And DisIDSubrubro = SRuID " _
            & " And CheVencimiento Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDRHaber Is Not Null " _
            & " Group by DisMoneda, SRuCodigo, SRuNombre" _
                    & " UNION ALL " _
            & " Select DisMoneda, SRuCodigo, SRuNombre, Importe = Sum(MDRImportePesos), DH = 'Debe', IOriginal = Sum(MDRDebe)" _
            & " From Cheque, MovimientoDisponibilidadRenglon, Disponibilidad, SubRubro" _
            & " Where CheID = MDRIDCheque " _
            & " And MDRIdDisponibilidad = DisID " _
            & " And DisIDSubrubro = SRuID " _
            & " And CheVencimiento Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDRDebe Is Not Null " _
            & " Group by DisMoneda, SRuCodigo, SRuNombre"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        With vsConsulta
        Do While Not rsAux.EOF
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!SRuCodigo, "000000000") & " " & Trim(rsAux!SRuNombre)
            
            'TC del ultimo dia del mes anterior
            aTC = TasadeCambio(rsAux!DisMoneda, paMonedaPesos, UltimoDia(DateAdd("m", -1, CDate(tFecha.Text))))
            
            Select Case LCase(rsAux!DH)
                Case "debe"
                    .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Importe, FormatoMonedaP)
                    If rsAux!IOriginal <> rsAux!Importe Then
                        .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!IOriginal, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!IOriginal * aTC, FormatoMonedaP)
                    End If
                
                Case "haber"
                    .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Importe, FormatoMonedaP)
                    If rsAux!IOriginal <> rsAux!Importe Then
                        .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!IOriginal, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!IOriginal * aTC, FormatoMonedaP)
                    End If
            End Select
            rsAux.MoveNext
        Loop
        End With
    End If
    rsAux.Close
    
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub CargoListaMovimiento(aIDTipoM As Long)

    cons = " Select TMDNombre, MDiComentario, DH = 'Haber', Importe = Sum(MDrImportePesos)" _
            & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiTipo = " & aIDTipoM _
            & " And MDiID = MDRIDMovimiento " _
            & " And MDiIDCompra Is Null " _
            & " And MDiTipo = TMDCodigo " _
            & " And MDRHaber <> Null " _
            & " Group by TMDNombre, MDiComentario"
    
    cons = cons & " UNION ALL "
    
    cons = cons & " Select TMDNombre, MDiComentario, DH = 'Debe', Importe = Sum(MDrImportePesos)" _
                        & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiTipo = " & aIDTipoM _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = TMDCodigo " _
                        & " And MDRDebe <> Null " _
                        & " Group by TMDNombre, MDiComentario"
    
    Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rs1.EOF
        
        With vsMovimiento
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = rs1!TMDNombre
            If Not IsNull(rs1!MDiComentario) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(rs1!MDiComentario)
            
            Select Case LCase(rs1!DH)
                Case "debe": .Cell(flexcpText, .Rows - 1, 2) = Format(rs1!Importe, FormatoMonedaP)
                Case "haber": .Cell(flexcpText, .Rows - 1, 3) = Format(rs1!Importe, FormatoMonedaP)
            End Select
        End With
        
        rs1.MoveNext
    Loop
    rs1.Close
    
End Sub

Private Sub tFHasta_GotFocus()
    'With tFHasta: .SelStart = 0: .SelLength = Len(.Text): End With
    With tFHasta: .SelStart = 0: .SelLength = 2: End With
End Sub

Private Sub tFHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then opNetear.SetFocus
End Sub

Private Sub tFHasta_LostFocus()
    If IsDate(tFHasta.Text) Then tFHasta.Text = Format(tFHasta.Text, "dd/mm/yyyy")
End Sub

Private Sub vsConsulta_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)

    vsConsulta.Row = vsConsulta.MouseRow
    
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    If bCargarImpresion Then
        If vsConsulta.Rows = 1 And vsReval.Rows = 1 Then Screen.MousePointer = 0: Exit Sub
        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        EncabezadoListado vsListado, "Asientos Varios- Del " & Trim(tFecha.Text) & " al " & Trim(tFHasta.Text), False
        vsListado.FileName = "Listado Asientos Varios"
            
        If vsConsulta.Rows > 1 Then
            vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        End If
        
        If vsReval.Rows > 1 Then
            vsListado.Paragraph = "": vsListado.Paragraph = "Asientos de Revaluaciones Vtas en M/E"
            vsReval.ExtendLastCol = False: vsListado.RenderControl = vsReval.hwnd: vsReval.ExtendLastCol = True
        End If
        
        If vsMovimiento.Rows > 1 Then
            vsListado.NewPage
            vsListado.Paragraph = "": vsListado.Paragraph = "Movimientos sin Asientos"
            vsMovimiento.ExtendLastCol = False: vsListado.RenderControl = vsMovimiento.hwnd: vsMovimiento.ExtendLastCol = True
        End If
                
        vsListado.EndDoc
        bCargarImpresion = False
    End If
    
    If Imprimir Then
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    Screen.MousePointer = 0
    
    Exit Sub
    
errImprimir:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionConfigurar()
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
End Sub

Private Sub Ayuda(strTexto As String)
    Status.Panels(4).Text = strTexto
    Status.Refresh
End Sub

Private Sub vsConsulta_DblClick()
On Error GoTo errDblClick
    
    With vsConsulta
        If .Rows = 1 Then Exit Sub
        
        Dim mTipoM As String, mRubro As Long
        mRubro = 0
        
        mTipoM = Trim(.Cell(flexcpText, .Row, 0))
        mRubro = .Cell(flexcpData, .Row, 1)
        
        If mTipoM = "" Then Exit Sub
        
        
        mTipoM = Mid(mTipoM, 1, Len(mTipoM) - 1) 'El ultimo char es el orden para listar D/H
        
        Dim mFrm As New frmDetalleTM
        With mFrm
            .prmTipoMovimiento = mTipoM
            .prmIDRubro = mRubro
            .prmFecha1 = tFecha.Text
            .prmFecha2 = tFHasta.Text
            
            .Show , Me
        End With
        
        'Codigo Anterior para ir al Listado de Movimientos de Caja  -------------------------------------------------------------------------
        'Ir a Ver Movimientos de Caja por Tipo de Movimiento (x difs de TCs)
        'Dim mTipoM As String, mRubro As String
        'mTipoM = Trim(.Cell(flexcpText, .Row, 0))
        'mRubro = Trim(.Cell(flexcpData, .Row, 1))
        'If mTipoM = "" Or mRubro = "" Then Exit Sub
        
        'Saco la Disponibilidad que se lista con el Rubro XX    ----------------------------------------
        'cons = "Select DisID from Disponibilidad, SubRubro" & _
                " Where DisIDSubrubro = SRuID And SRuCodigo = '" & mRubro & "'"
                
        'Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        'If Not rsAux.EOF Then mRubro = rsAux!DisID Else mRubro = ""
        'rsAux.Close
        '------------------------------------------------------------------------------------------------------
        
        'mRubro = mRubro & ":" & mTipoM & ":" & Format(tFecha.Text, "dd/mm/yyyy") & ":" & Format(tFHasta.Text, "dd/mm/yyyy")
        'EjecutarApp prmPathApp & "Movimientos de Caja.exe", mRubro
        '-------------------------------------------------------------------------------------------------------------------------------------------------
    End With
    
errDblClick:
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    If Not IsDate(tFecha.Text) And IsDate(tFHasta.Text) Then
        MsgBox "Ingrese la fecha desde.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    If IsDate(tFecha.Text) And Not IsDate(tFHasta.Text) Then
        If Trim(tFHasta.Text) = "" Then
            tFHasta.Text = tFecha.Text
        Else
            MsgBox "La fecha hasta no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tFHasta: Exit Function
        End If
    End If
    If IsDate(tFecha.Text) And IsDate(tFHasta.Text) Then
        If CDate(tFecha.Text) > CDate(tFHasta.Text) Then
            MsgBox "Los rangos de fecha no son correctos.", vbExclamation, "ATENCIÓN"
            Foco tFecha: Exit Function
        End If
    Else
        Exit Function
    End If
        
    If Month(CDate(tFecha.Text)) <> Month(CDate(tFHasta.Text)) Or Year(CDate(tFecha.Text)) <> Year(CDate(tFHasta.Text)) Then
        MsgBox "Los asientos varios deben realizarse para un mismo mes." & vbCrLf & _
                    "Las fechas ingresadas no corresponde al mismo mes. Verifique", vbExclamation, "Error en Fechas"
        Foco tFecha: Exit Function
    End If
    
    ValidoCampos = True
    
End Function

Private Sub EliminoAsientosDobles()

    On Error GoTo errEliminar
    With vsConsulta
        If opNetear.Value = vbChecked Then
        For I = 1 To .Rows - 1
            If I + 1 <= .Rows - 1 Then
                If .Cell(flexcpText, I, 0) = .Cell(flexcpText, I + 1, 0) And (.Cell(flexcpText, I, 2) <> "" Or .Cell(flexcpText, I, 4) <> "") Then
                    If .Cell(flexcpText, I, 1) = .Cell(flexcpText, I + 1, 1) Then
                        
                        If .Cell(flexcpText, I, 2) <> "" Then
                            Select Case .Cell(flexcpValue, I, 2)
                                Case Is > .Cell(flexcpValue, I + 1, 4)
                                            .Cell(flexcpText, I, 2) = Format(.Cell(flexcpValue, I, 2) - .Cell(flexcpValue, I + 1, 4), FormatoMonedaP)
                                            If .Cell(flexcpText, I, 3) <> "" Then .Cell(flexcpText, I, 3) = Format(.Cell(flexcpValue, I, 3) - .Cell(flexcpValue, I + 1, 5), FormatoMonedaP)
                                            .RemoveItem I + 1
                                
                                Case Is < .Cell(flexcpValue, I + 1, 4)
                                            .Cell(flexcpText, I + 1, 4) = Format(.Cell(flexcpValue, I + 1, 4) - .Cell(flexcpValue, I, 2), FormatoMonedaP)
                                            If .Cell(flexcpText, I + 1, 5) <> "" Then .Cell(flexcpText, I + 1, 5) = Format(.Cell(flexcpValue, I + 1, 5) - .Cell(flexcpValue, I, 3), FormatoMonedaP)
                                            .RemoveItem I
                                
                                Case Is = .Cell(flexcpValue, I + 1, 4): .RemoveItem I: .RemoveItem I
                            End Select
                        Else
                            Select Case .Cell(flexcpValue, I, 4)
                                Case Is > .Cell(flexcpValue, I + 1, 2)
                                            .Cell(flexcpText, I, 4) = Format(.Cell(flexcpValue, I, 4) - .Cell(flexcpValue, I + 1, 2), FormatoMonedaP)
                                            If .Cell(flexcpText, I, 5) <> "" Then .Cell(flexcpText, I, 5) = Format(.Cell(flexcpValue, I, 5) - .Cell(flexcpValue, I + 1, 3), FormatoMonedaP)
                                            .RemoveItem I + 1
                                
                                Case Is < .Cell(flexcpValue, I + 1, 2)
                                            .Cell(flexcpText, I + 1, 2) = Format(.Cell(flexcpValue, I + 1, 2) - .Cell(flexcpValue, I, 4), FormatoMonedaP)
                                            If .Cell(flexcpText, I + 1, 3) <> "" Then .Cell(flexcpText, I + 1, 3) = Format(.Cell(flexcpValue, I + 1, 3) - .Cell(flexcpValue, I, 5), FormatoMonedaP)
                                            .RemoveItem I
                                
                                Case Is = .Cell(flexcpValue, I + 1, 2): .RemoveItem I: .RemoveItem I
                            End Select
                        End If
                    End If
                End If
            End If
        Next
        
        For I = 1 To .Rows - 1
            If I + 1 <= .Rows - 1 Then
                If .Cell(flexcpBackColor, I, 0) = Colores.Inactivo And .Cell(flexcpBackColor, I + 1, 0) = Colores.Inactivo Then
                    .RemoveItem I
                End If
            End If
        Next
        If .Cell(flexcpBackColor, .Rows - 1, 0) = Colores.Inactivo Then .RemoveItem .Rows - 1
        End If
        
        'Vuelvo a Ordenar para que me quede DEBE/HABER
        For I = 1 To .Rows - 1
            If .Cell(flexcpBackColor, I, 0) = Colores.Inactivo Then
                .Cell(flexcpText, I, 0) = .Cell(flexcpText, I, 0) & "2"
            Else
                If .Cell(flexcpText, I, 2) <> "" Then .Cell(flexcpText, I, 0) = .Cell(flexcpText, I, 0) & "2" Else .Cell(flexcpText, I, 0) = .Cell(flexcpText, I, 0) & "1"
            End If
        Next
        If vsConsulta.Rows > 1 Then
            .Select 1, 0, 1, 1
            .Sort = flexSortGenericDescending
        End If
    End With
    Exit Sub

errEliminar:
    clsGeneral.OcurrioError "Ocurrió un error al eliminar los asientos dobles.", Err.Description
End Sub

Private Function meCargoTasasDeCambio(idMoneda As Long)

    'Cargo array TCs contra pesos   --------------------------------------------------------
    If prmMonedaTC = idMoneda Then Exit Function
    prmMonedaTC = idMoneda
    Ayuda "Cargando Tasas de Cambio..."
    
    ReDim arrTC(0)
    Dim mQ As Integer, dDesde As Date
    
    dDesde = CDate(tFecha.Text) - 1
    mQ = DateDiff("d", dDesde, CDate(tFHasta.Text))
    
    Dim I As Integer
    Dim mTasa As Currency
    
    For I = 1 To mQ
        mTasa = TasadeCambio(CInt(idMoneda), paMonedaPesos, dDesde + I, TipoTC:=prmTCVentasME)
        If I <> 1 Then ReDim Preserve arrTC(UBound(arrTC) + 1)
        With arrTC(UBound(arrTC))
            .Fecha = DatePart("d", dDesde + I)
            .Valor = mTasa
        End With
    Next
    
End Function

Private Function meCargoTasasDeCambioMesAnterior(idMoneda As Long)

    'Cargo array TCs contra pesos   --------------------------------------------------------
    Ayuda "Cargando Tasas de Cambio Mes Anterior..."
    
    Dim mQ As Integer, dDesde As Date
    
    'dDesde = CDate("01/" & Month(CDate(tFecha.Text)) - 1 & "/" & Year(CDate(tFecha.Text)))
    dDesde = "01/" & Format(DateAdd("m", -1, CDate(tFecha.Text)), "mm/yyyy")
    mQ = Day(UltimoDia(dDesde))
    
    ReDim arrTCMA(mQ)
    
    dDesde = CDate(dDesde) - 1
        
    Dim I As Integer
    Dim mTasa As Currency
    
    For I = 1 To mQ
        mTasa = TasadeCambio(CInt(idMoneda), paMonedaPesos, dDesde + I, TipoTC:=prmTCVentasME)
        arrTCMA(I) = mTasa
    Next
    
End Function

Private Function meTCaFecha(mFecha As Integer) As Currency

    meTCaFecha = 1
    For I = LBound(arrTC) To UBound(arrTC)
        'If Format(arrTC(I).Fecha, "dd/mm/yyyy") = Format(mFecha, "dd/mm/yyyy") Then
        If Val(arrTC(I).Fecha) = Val(mFecha) Then
            meTCaFecha = arrTC(I).Valor
            Exit For
        End If
    Next
    
End Function

Private Function meSubroParaDisponibilidad(idDisp As Long, Optional mMoneda As String) As String

Dim rsData As rdoResultset

        'Saco el subrubro de la disponibilidad para hacer el asiento final --------------------------------------------
        meSubroParaDisponibilidad = "CAJA " & mMoneda
        
        cons = "Select * from Disponibilidad, Subrubro " _
               & " Where DisID = " & idDisp _
               & " And DisIDSubRubro = SRuId"
        Set rsData = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        If Not rsData.EOF Then
            meSubroParaDisponibilidad = Format(rsData!SRuCodigo, "000000000") & " " & Trim(rsData!SRuNombre)
        End If
        rsData.Close
        '------------------------------------------------------------------------------------------------------------------------------------
        
End Function

Private Function meAddarrAsientos(mTexto As String, mImporte As Currency)

    'Si el Subrubro está en el array --> Acumulo los datos, si no lo agrego --------------------------------------
    If arrAsientos(0).NombreRubro = "" And UBound(arrAsientos) = 0 Then
        With arrAsientos(0)
            .NombreRubro = mTexto
            .Importe = Format(mImporte, FormatoMonedaP)
        End With
    Else
        Dim bOK As Boolean: bOK = False
        Dim idx As Integer
        
        For idx = LBound(arrAsientos) To UBound(arrAsientos)
            If arrAsientos(idx).NombreRubro = mTexto Then
                arrAsientos(idx).Importe = arrAsientos(idx).Importe + Format(mImporte, FormatoMonedaP)
                bOK = True
                Exit For
            End If
        Next
        
        If Not bOK Then
            ReDim Preserve arrAsientos(UBound(arrAsientos) + 1)
            With arrAsientos(UBound(arrAsientos))
                .NombreRubro = mTexto
                .Importe = Format(mImporte, FormatoMonedaP)
            End With
        End If
    End If
    '--------------------------------------------------------------------------------------------------------------------------
End Function

Private Function meAddarrAsientosCtas(mTexto As String, mImporte As Currency)

    'Si el Subrubro está en el array --> Acumulo los datos, si no lo agrego --------------------------------------
    If arrAsientosCtas(0).NombreRubro = "" And UBound(arrAsientosCtas) = 0 Then
        With arrAsientosCtas(0)
            .NombreRubro = mTexto
            .Importe = Format(mImporte, FormatoMonedaP)
        End With
    Else
        Dim bOK As Boolean: bOK = False
        Dim idx As Integer
        
        For idx = LBound(arrAsientosCtas) To UBound(arrAsientosCtas)
            If arrAsientosCtas(idx).NombreRubro = mTexto Then
                arrAsientosCtas(idx).Importe = arrAsientosCtas(idx).Importe + Format(mImporte, FormatoMonedaP)
                bOK = True
                Exit For
            End If
        Next
        
        If Not bOK Then
            ReDim Preserve arrAsientosCtas(UBound(arrAsientosCtas) + 1)
            With arrAsientosCtas(UBound(arrAsientosCtas))
                .NombreRubro = mTexto
                .Importe = Format(mImporte, FormatoMonedaP)
            End With
        End If
    End If
    '--------------------------------------------------------------------------------------------------------------------------
End Function


Private Function PastoSQL(mSQL As String)
On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText mSQL
End Function

Private Sub AccionConsultarReval()

On Error GoTo errConsultar
Dim mTotalME As Currency
Dim mVentas As Currency, mIva As Currency, mCofis As Currency
Dim mFNeto As Currency, mFIva As Currency, mFCofis As Currency

Dim rsMon As rdoResultset
Dim mNameMoneda As String, mIDMoneda As Long
Dim myFecha As String
Dim mIndex As Integer

    Screen.MousePointer = 11
    bCargarImpresion = True
    
    vsReval.Rows = 1: vsReval.Refresh
    Dim aQDias As Integer
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    aQDias = DateDiff("d", CDate(tFecha.Text), CDate(tFHasta.Text)) + 1
    pbProgreso.Max = aQDias
    '-----------------------------------------------------------------------------------------------------------------
    
    cons = "Select * from Moneda Where MonCoeficienteMora Is Not Null And MonCodigo <> " & paMonedaPesos
    Set rsMon = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsMon.EOF
        mNameMoneda = Trim(rsMon!MonSigno)
        mIDMoneda = rsMon!MonCodigo
    
        meCargoTasasDeCambio mIDMoneda
        meCargoTasasDeCambioMesAnterior mIDMoneda
    
        pbProgreso.Value = 0
        mVentas = 0: mIva = 0: mCofis = 0
        mTotalME = 0
        pbProgreso.Value = 0
        
        For mIndex = 1 To aQDias
        
            myFecha = Day(CDate(tFecha.Text)) + (mIndex - 1) & "/" & Format(tFecha.Text, "mm/yyyy")
            Ayuda "Generando Revaluaciones " & mNameMoneda & " día " & Day(myFecha) & " ..."
        
            cons = "Select Factura.DocFecha as FacFecha, Factura.DocTotal as FacTotal, Factura.DocIVA as FacIva, Factura.DocCofis as FacCofis, " & _
                                " Recibo.DocCodigo as IDRecibo, Recibo.DocFecha as RecFPago, Pago.DPaDocASaldar as IDFactura, " & _
                                " ReciboA.DocFecha as RecFPagoA " & _
                    " From DocumentoPago as Pago " & _
                        " Left Outer Join DocumentoPago as Anterior " & _
                                    " On Pago.DPaDocASaldar = Anterior.DPaDocASaldar " & _
                                    " And (Anterior.DPaCuota = (Pago.DPaCuota - 1)   OR Anterior.DPaCuota = Pago.DPaCuota )" & _
                                    " And Anterior.DPaAmortizacion Is Not Null " & _
                                    " And Anterior.DPaDocQSalda <> Pago.DPaDocQSalda " & _
                                                " Left Outer Join Documento as ReciboA On Anterior.DPaDocQSalda = ReciboA.DocCodigo And ReciboA.DocAnulado = 0, " & _
                " Documento as Recibo (index = iTipoFechaSucursalMoneda), Documento as Factura " & _
                " Where Pago.DPaDocQSalda = Recibo.DocCodigo " & _
                " And Recibo.DocFecha Between " & Format(myFecha, "'mm/dd/yyyy'") & " AND " & Format(myFecha, "'mm/dd/yyyy 23:59'") & _
                " And Recibo.DocTipo = " & TipoDocumento.ReciboDePago & _
                " And Recibo.DocAnulado = 0 " & _
                " And Recibo.DocMoneda = " & mIDMoneda & _
                " And Pago.DPaAmortizacion Is Not Null" & _
                " And Pago.DPaDocASaldar = Factura.DocCodigo"
    
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            
            vsReval.Redraw = False
            
            Dim mPagoAnterior As Date, mPago As Date
            Dim mTC As Currency, mTCA As Currency
            
            Dim mIdRecibo As Long: mIdRecibo = 0
            
            Do While Not rsAux.EOF
                    If mIdRecibo = rsAux!IDRecibo Then GoTo nextPago
                    mIdRecibo = rsAux!IDRecibo
                    
                    If Not IsNull(rsAux!FacIva) Then mFIva = rsAux!FacIva Else mFIva = 0
                    If Not IsNull(rsAux!FacCofis) Then mFCofis = rsAux!FacCofis Else mFCofis = 0
                    
                    If Not IsNull(rsAux!FacTotal) Then
                        mFNeto = rsAux!FacTotal
                        mTotalME = mTotalME + rsAux!FacTotal
                    Else
                        mFNeto = 0
                    End If
                    
                    mFNeto = mFNeto - mFIva - mFCofis
                    
                    mPago = Format(rsAux!RecFPago, "dd/mm/yyyy")
                    If Not IsNull(rsAux!RecFPagoA) Then
                        mPagoAnterior = Format(rsAux!RecFPagoA, "dd/mm/yyyy")
                    Else
                        mPagoAnterior = Format(rsAux!FacFecha, "dd/mm/yyyy")
                    End If
                    
                    If mPago <> mPagoAnterior Then  'Revalúo
                        mTC = meTCaFecha(Day(mPago))
                        
                        'Si el pago anterior fue en el mes anterior saco TC del array
                        If Format(DateAdd("m", -1, mPago), "mm/yyyy") = Format(mPagoAnterior, "mm/yyyy") Then
                            mTCA = arrTCMA(Day(mPagoAnterior))
                        Else
                            mTCA = TasadeCambio(CInt(mIDMoneda), paMonedaPesos, mPagoAnterior, TipoTC:=prmTCVentasME)
                        End If
                        
                        If mTC <> mTCA Then
                            mVentas = mVentas + ((mFNeto * mTC) - (mFNeto * mTCA))
                            mIva = mIva + ((mFIva * mTC) - (mFIva * mTCA))
                            mCofis = mCofis + ((mFCofis * mTC) - (mFCofis * mTCA))
                        End If
                    End If
                    
nextPago:
                rsAux.MoveNext
            Loop
            
            rsAux.Close
            pbProgreso.Value = pbProgreso.Value + 1
        Next
        
        If mTotalME <> 0 Then
        With vsReval
            .AddItem ""
            
            .Cell(flexcpText, .Rows - 1, 1) = "Revaluac. Vtas en " & mNameMoneda & " al momento del cobro"
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVentaME)
            .Cell(flexcpText, .Rows - 1, 3) = Format(mTotalME, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 2) = Format(mVentas + mIva + mCofis, FormatoMonedaP)
        
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentasME)
            .Cell(flexcpText, .Rows - 1, 4) = Format(mVentas, FormatoMonedaP)
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubRubroCofis)
            .Cell(flexcpText, .Rows - 1, 4) = Format(mCofis, FormatoMonedaP)
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
            .Cell(flexcpText, .Rows - 1, 4) = Format(mIva, FormatoMonedaP)
        End With
        End If
        
        rsMon.MoveNext
    Loop
    rsMon.Close
    
    vsReval.Redraw = True
    pbProgreso.Value = 0
    Ayuda ""
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
    vsReval.Redraw = True: Screen.MousePointer = 0
    PastoSQL cons
End Sub


Private Function CargoTasasDeCambio()

On Error Resume Next

    With vsConsulta
        For I = 1 To .Rows - 1
            If .Cell(flexcpValue, I, 2) <> 0 And .Cell(flexcpValue, I, 3) <> 0 Then
                .Cell(flexcpText, I, 6) = Format(.Cell(flexcpValue, I, 2) / .Cell(flexcpValue, I, 3), "#,##0.0000000")
            End If
            
            If .Cell(flexcpValue, I, 4) <> 0 And .Cell(flexcpValue, I, 5) <> 0 Then
                .Cell(flexcpText, I, 6) = Format(.Cell(flexcpValue, I, 4) / .Cell(flexcpValue, I, 5), "#,##0.0000000")
            End If
        
        Next
    End With
    
End Function

Private Function StartMe()
On Error Resume Next
    With img1
        bConsultar.Picture = .ListImages("consultar").ExtractIcon
        bPrimero.Picture = .ListImages("move1").ExtractIcon
        bAnterior.Picture = .ListImages("move2").ExtractIcon
        bSiguiente.Picture = .ListImages("move3").ExtractIcon
        bUltima.Picture = .ListImages("move4").ExtractIcon
        
        bImprimir.Picture = .ListImages("print").ExtractIcon
        bConfigurar.Picture = .ListImages("configprint").ExtractIcon
        
        bNoFiltros.Picture = .ListImages("limpiar").ExtractIcon
        bCancelar.Picture = .ListImages("salir").ExtractIcon
        chVista.Picture = .ListImages("vista1").ExtractIcon
        chVista.DownPicture = .ListImages("vista2").ExtractIcon
        
    End With

End Function


