VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{B443E3A5-0B4D-4B43-B11D-47B68DC130D7}#1.5#0"; "orArticulo.ocx"
Begin VB.Form CoEmbarque 
   Caption         =   "Consulta de Embarques"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CoEmbarque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10920
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   34
      Top             =   4080
      Width           =   5175
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "CoEmbarque.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "CoEmbarque.frx":067C
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "CoEmbarque.frx":097E
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "CoEmbarque.frx":0CC0
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   4320
         Picture         =   "CoEmbarque.frx":0FC2
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   3600
         Picture         =   "CoEmbarque.frx":10C4
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3240
         Picture         =   "CoEmbarque.frx":148A
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
   End
   Begin ComctlLib.TabStrip tsVistas 
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Vista &1- General"
            Object.Tag             =   "Vista 1- General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Vista &2- Embarques"
            Object.Tag             =   "Vista 2- Embarques"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Vista &3- Arribos"
            Object.Tag             =   "Vista 3- Arribos"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   5220
      Width           =   10920
      _ExtentX        =   19262
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
            Object.Width           =   11033
            Object.Tag             =   ""
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
      Height          =   1695
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   10635
      Begin VB.TextBox tProveedor 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox tLC 
         Height          =   285
         Left            =   4740
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   960
         Width           =   1995
      End
      Begin VB.TextBox tGrupo 
         Height          =   285
         Left            =   4740
         TabIndex        =   9
         Top             =   240
         Width           =   1995
      End
      Begin VB.TextBox tTipo 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin prjFindArticulo.orArticulo tArticulo 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
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
      End
      Begin VB.TextBox tTop 
         Height          =   285
         Left            =   8340
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox tAgencia 
         Height          =   315
         Left            =   7920
         TabIndex        =   15
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox chCosteada 
         Alignment       =   1  'Right Justify
         Caption         =   "C&osteados"
         Height          =   255
         Left            =   8640
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox tVapor 
         Height          =   315
         Left            =   4740
         TabIndex        =   13
         Top             =   600
         Width           =   1995
      End
      Begin VB.CheckBox chEmbarcado 
         Caption         =   "E&mbarcados"
         Height          =   255
         Left            =   5340
         TabIndex        =   21
         Top             =   1320
         Value           =   2  'Grayed
         Width           =   1215
      End
      Begin VB.CheckBox chArribado 
         Caption         =   "Arri&bados"
         Height          =   255
         Left            =   4140
         TabIndex        =   20
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox tFactura 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox tFecha 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6660
         TabIndex        =   23
         Top             =   1320
         Width           =   1815
      End
      Begin AACombo99.AACombo cboPrioridad 
         Height          =   315
         Left            =   7920
         TabIndex        =   19
         Top             =   960
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
      Begin VB.Label Label10 
         Caption         =   "Prioridad:"
         Height          =   255
         Left            =   7080
         TabIndex        =   18
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "&Sólo los últimos:"
         Height          =   315
         Left            =   7080
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Age&ncia:"
         Height          =   255
         Left            =   7080
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "&L/C:"
         Height          =   255
         Left            =   4140
         TabIndex        =   16
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "&Factura:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "&Vapor:"
         Height          =   255
         Left            =   4140
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo:"
         Height          =   255
         Left            =   4140
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label label2 
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid lConsulta 
      Height          =   1575
      Index           =   0
      Left            =   0
      TabIndex        =   36
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
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
      HighLight       =   1
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
   Begin VSFlex6DAOCtl.vsFlexGrid lConsulta 
      Height          =   1575
      Index           =   1
      Left            =   1440
      TabIndex        =   37
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
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
      HighLight       =   1
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
   Begin VSFlex6DAOCtl.vsFlexGrid lConsulta 
      Height          =   1575
      Index           =   2
      Left            =   2880
      TabIndex        =   38
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
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
      HighLight       =   1
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
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   2895
      Left            =   240
      TabIndex        =   35
      Top             =   1920
      Visible         =   0   'False
      Width           =   8895
      _Version        =   196608
      _ExtentX        =   15690
      _ExtentY        =   5106
      _StockProps     =   229
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
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.Menu MnuMouse 
      Caption         =   "Mouse"
      Visible         =   0   'False
      Begin VB.Menu MnuEmbarque 
         Caption         =   "Ver Embarque"
      End
      Begin VB.Menu MnuL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBuscarSub 
         Caption         =   "Buscar Subcarpetas"
      End
      Begin VB.Menu MnuBuscarGas 
         Caption         =   "Buscar Gastos"
      End
      Begin VB.Menu MnuCosteo 
         Caption         =   "Evolución del Costeo"
      End
      Begin VB.Menu MnuL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCancelarMenu 
         Caption         =   "Cancelar Menú"
      End
   End
End
Attribute VB_Name = "CoEmbarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Modificaciones:
    ' 29-5-2001 En ArmoConsulta puse que la carpeta no este anulada.
    ' 17-2-2003 Agregue opción Top, el textbox + ajuste de consulta.

Const PorPantalla = 150
Dim PosLista As Integer

Dim RsConsulta As rdoResultset
Dim aSeleccionado As Long

Dim aFormato As String

Dim aFechas As String
Dim aTexto As String
Dim gTxtCarpeta As String

Private Sub AccionSiguiente()

    On Error GoTo ErrAA
    If Not bSiguiente.Enabled Then Exit Sub
    
    If Not RsConsulta.EOF Then
        bAnterior.Enabled = True: bPrimero.Enabled = True
        
        PosLista = PosLista + (lConsulta(0).Rows - 1)
        CargoLista
    Else
        MsgBox "Se ha llegado al final de la consulta, no hay más datos a desplegar.", vbInformation, "ATENCIÓN"
    End If
    
    lConsulta(tsVistas.SelectedItem.Index - 1).SetFocus
    Exit Sub
    
ErrAA:
    Screen.MousePointer = 0
End Sub

Private Sub AccionAnterior()

Dim UltimaPosicion As Long

    On Error GoTo ErrAA
    If Not bAnterior.Enabled Then Exit Sub
    
    If RsConsulta.EOF And lConsulta(0).Rows - 1 > 0 And RsConsulta.AbsolutePosition = -1 Then
        Screen.MousePointer = 11
        RsConsulta.MoveLast
        Screen.MousePointer = 0
        UltimaPosicion = PosLista + lConsulta(0).Rows
        RsConsulta.MoveNext
        Screen.MousePointer = 0
    Else
        UltimaPosicion = PosLista + lConsulta(0).Rows + 1
    End If
    
    If UltimaPosicion - (lConsulta(0).Rows - 1) - PorPantalla >= 1 Then

        If UltimaPosicion - (lConsulta(0).Rows - 1) - PorPantalla = 1 Then bAnterior.Enabled = False: bPrimero.Enabled = False
        bSiguiente.Enabled = True
        
        RsConsulta.Move UltimaPosicion - (lConsulta(0).Rows - 1) - PorPantalla, 1
        CargoLista
        PosLista = PosLista - (lConsulta(0).Rows - 1)
        Screen.MousePointer = 0
    Else
        MsgBox "Se ha llegado al principio de la consulta.", vbInformation, "ATENCIÓN"
    End If
    lConsulta(tsVistas.SelectedItem.Index - 1).SetFocus
    Exit Sub
    
ErrAA:
    Screen.MousePointer = 0
End Sub

Private Sub AccionPrimero()
    
    If Not bPrimero.Enabled Then Exit Sub
    
    PosLista = 0
    Screen.MousePointer = 11
    On Error Resume Next
    RsConsulta.MoveFirst
    On Error GoTo ErrAA
    CargoLista
    lConsulta(tsVistas.SelectedItem.Index - 1).SetFocus
    bSiguiente.Enabled = True
    bPrimero.Enabled = False: bAnterior.Enabled = False
    Screen.MousePointer = 0
    Exit Sub
    
ErrAA:
    Screen.MousePointer = 0
End Sub

Private Sub AccionLimpiar()

    tTipo.Text = ""
    tGrupo.Text = ""
    tArticulo.Text = ""
    tVapor.Text = ""
    tProveedor.Text = ""
    tLC.Text = ""
    
    tFecha.Text = "": tFecha.Enabled = True
    tFactura.Text = ""
    chArribado.Value = vbUnchecked
    chEmbarcado.Value = vbGrayed
    chCosteada.Value = vbGrayed

End Sub

Private Sub bAnterior_Click()
    AccionAnterior
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    If VerificoFiltros Then AccionConsultar
End Sub

Private Sub bImprimir_Click()
    AccionImprimir
End Sub

Private Sub bNoFiltros_Click()
    AccionLimpiar
End Sub

Private Sub bPrimero_Click()
    AccionPrimero
End Sub

Private Sub bSiguiente_Click()
    AccionSiguiente
End Sub

Private Sub cboPrioridad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then chArribado.SetFocus
End Sub

Private Sub chArribado_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbShiftMask Then EstadoCheck chArribado
End Sub

Private Sub chArribado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chEmbarcado.SetFocus
End Sub

Private Sub chArribado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then EstadoCheck chArribado
End Sub

Private Sub chCosteada_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbShiftMask Then EstadoCheck chCosteada
End Sub

Private Sub chCosteada_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub chCosteada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then EstadoCheck chCosteada
End Sub

Private Sub chEmbarcado_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbShiftMask Then EstadoCheck chEmbarcado
End Sub

Private Sub chEmbarcado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If tFecha.Enabled Then Foco tFecha Else Foco tTipo
End Sub

Private Sub chEmbarcado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then EstadoCheck chEmbarcado
End Sub
Private Sub Label9_Click()
    Foco tTop
End Sub

Private Sub lConsulta_Click(Index As Integer)
    With lConsulta(Index)
        If .MouseRow = 0 Then
            .ColSel = .MouseCol
            If .ColSort(lConsulta(Index).MouseCol) = flexSortGenericAscending Then
                .ColSort(.MouseCol) = flexSortGenericDescending
            Else
                .ColSort(.MouseCol) = flexSortGenericAscending
            End If
            .Sort = flexSortUseColSort
        End If
    End With
End Sub

Private Sub lConsulta_RowColChange(Index As Integer)
    On Error Resume Next
    Static inhere%
    inhere = False
    If inhere Then Exit Sub
    inhere = True
    
    'Synchronize seleccion en otras vistas
    For I = 0 To 2
        If I <> Index Then lConsulta(I).Select lConsulta(Index).Row, 0
    Next
    
    inhere = False
    
End Sub

Private Sub lConsulta_Scroll(Index As Integer)

    On Error Resume Next
    Static inhere%
    inhere = False
    If inhere Then Exit Sub
    inhere = True
    
    'Synchronize scroll en otras vistas
    For I = 0 To 2
        If I <> Index Then lConsulta(I).TopRow = lConsulta(Index).TopRow
    Next
    
    inhere = False
    
End Sub

Private Sub MnuBuscarGas_Click()
    On Error GoTo errApp
    Dim RetVal
    Screen.MousePointer = 11: Me.Refresh
    RetVal = Shell(App.Path & "\Consulta de Gastos " & gTxtCarpeta, 1)
    Screen.MousePointer = 0
    Exit Sub
errApp:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al ejecutar la aplicación. ", Err.Description
End Sub

Private Sub MnuBuscarSub_Click()
On Error GoTo errApp
    Dim RetVal
    Screen.MousePointer = 11: Me.Refresh
    RetVal = Shell(App.Path & "\Consulta Subcarpetas", 1)
    Screen.MousePointer = 0
    Exit Sub
errApp:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al ejecutar la aplicación. ", Err.Description
End Sub

Private Sub MnuCosteo_Click()
    On Error GoTo errApp
    Dim RetVal
    Screen.MousePointer = 11: Me.Refresh
    
    With lConsulta(tsVistas.SelectedItem.Index - 1)
        RetVal = Shell(App.Path & "\Costeo de Carpetas " & Folder.cFEmbarque & .Cell(flexcpData, .Row, 0), 1)
    End With
        
    Screen.MousePointer = 0
    Exit Sub
errApp:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al ejecutar la aplicación. ", Err.Description
End Sub

Private Sub MnuEmbarque_Click()
    On Error GoTo errApp
    Dim RetVal
    Screen.MousePointer = 11: Me.Refresh
    
    With lConsulta(tsVistas.SelectedItem.Index - 1)
        RetVal = Shell(App.Path & "\Embarque " & .Cell(flexcpData, .Row, 0), 1)
    End With
        
    Screen.MousePointer = 0
    Exit Sub
errApp:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al ejecutar la aplicación. ", Err.Description
End Sub

Private Sub tAgencia_Change()
    lConsulta(0).Rows = 1: lConsulta(1).Rows = 1: lConsulta(2).Rows = 1
    tAgencia.Tag = 0
End Sub

Private Sub tAgencia_GotFocus()
    With tAgencia
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrBA
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF1 Then
    
        If Val(tAgencia.Tag) = 0 And Trim(tAgencia.Text) <> "" Then
            Screen.MousePointer = 11
            Cons = "Select ATrCodigo, 'Agencia' = ATrNombre From AgenciaTransporte Where ATrNombre Like '" & fnc_GetQueryLike(tAgencia.Text) & "%'"
            loc_FindByName tAgencia, Cons, 1, "Agencias", 4000
            Screen.MousePointer = 0
        Else
             tLC.SetFocus
        End If
    End If
    Exit Sub
ErrBA:
    MsgBox "Ocurrio un error al buscar la agencia.", vbCritical, "ATENCIÓN"
    Screen.MousePointer = 0
End Sub

Private Sub tArticulo_Change()
    lConsulta(0).Rows = 1: lConsulta(1).Rows = 1: lConsulta(2).Rows = 1
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And (tArticulo.prm_ArtID > 0 Or tArticulo.Text = "") Then tProveedor.SetFocus
End Sub

Private Sub tFecha_Change()
    lConsulta(0).Rows = 1: lConsulta(1).Rows = 1: lConsulta(2).Rows = 1
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chCosteada.SetFocus
End Sub

Private Sub tGrupo_Change()
    lConsulta(0).Rows = 1: lConsulta(1).Rows = 1: lConsulta(2).Rows = 1
    tGrupo.Tag = ""
End Sub

Private Sub tGrupo_GotFocus()
    With tGrupo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tGrupo.Tag) > 0 Or tGrupo.Text = "" Then
            tTop.SetFocus
        Else
            loc_FindGrupo
        End If
    End If
End Sub

Private Sub tLC_Change()
    lConsulta(0).Rows = 1: lConsulta(1).Rows = 1: lConsulta(2).Rows = 1
    tLC.Tag = ""
End Sub

Private Sub tLC_GotFocus()
    With tLC
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tLC_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If tLC.Tag <> "" Or tLC.Text = "" Then
            cboPrioridad.SetFocus
        Else
            loc_FindLC
        End If
    End If
End Sub

Private Sub tProveedor_Change()
    lConsulta(0).Rows = 1: lConsulta(1).Rows = 1: lConsulta(2).Rows = 1
    tProveedor.Tag = ""
End Sub

Private Sub tProveedor_GotFocus()
    With tProveedor
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) > 0 Or tProveedor.Text = "" Then
            tFactura.SetFocus
        Else
            loc_FindProveedor
        End If
    End If
End Sub

Private Sub tTipo_Change()
    lConsulta(0).Rows = 1: lConsulta(1).Rows = 1: lConsulta(2).Rows = 1
    tTipo.Tag = ""
End Sub

Private Sub tTipo_GotFocus()
    With tTipo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tTipo.Tag) > 0 Or tTipo.Text = "" Then
            tArticulo.SetFocus
        Else
            loc_FindTipo
        End If
    End If
End Sub

Private Sub tTop_GotFocus()
    With tTop
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tTop_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tVapor.SetFocus
    End If
End Sub

Private Sub tTop_LostFocus()
    If Not IsNumeric(tTop.Text) Then tTop.Text = ""
End Sub

Private Sub tVapor_Change()
    lConsulta(0).Rows = 1: lConsulta(1).Rows = 1: lConsulta(2).Rows = 1
    tVapor.Tag = 0
End Sub

Private Sub tVapor_GotFocus()
    tVapor.SelStart = 0
    tVapor.SelLength = Len(tVapor.Text)
End Sub

Private Sub tVapor_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF1 Then
    
        If Val(tVapor.Tag) = 0 And Trim(tVapor.Text) <> "" Then
            Screen.MousePointer = 11
            Cons = "Select TraCodigo, 'Transporte/Vapor' = TraNombre From Transporte Where TraNombre Like '" & fnc_GetQueryLike(tVapor.Text) & "%'"
            loc_FindByName tVapor, Cons, 1, "Transporte/Vapor", 4000
            Screen.MousePointer = 0
        Else
             Foco tAgencia
        End If
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    bPrimero.Enabled = False: bSiguiente.Enabled = False: bAnterior.Enabled = False
    
    InicializoGrillas
    lConsulta(0).ZOrder 0
    
    CargoCombo "SELECT CodID, RTRIM(CodTexto) FROM Codigos WHERE CodCual = 165 ORDER BY CodTexto", cboPrioridad
    
    Cons = "Select * from ArticuloFolder Where AFoTipo = 0 And AFoCodigo = 0"
    Set RsConsulta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    AccionLimpiar
    With tArticulo
        Set .Connect = cBase
        .FindNombreEnUso = False
    End With
        
    
    ObtengoSeteoForm Me, WidthIni:=10215, HeightIni:=6825
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyE: AccionConsultar
            
            Case vbKeyP: AccionPrimero
            Case vbKeyA: AccionAnterior
            Case vbKeyS: AccionSiguiente
            
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    picBotones.BorderStyle = vbFlat
    picBotones.Top = Me.Height - picBotones.Height - 700
    
    tsVistas.Height = Me.Height - tsVistas.ClientTop - 400 - picBotones.Height
    tsVistas.Width = Me.Width - (tsVistas.ClientLeft * 2)
    
    lConsulta(0).Top = tsVistas.ClientTop + 50
    lConsulta(1).Top = lConsulta(0).Top: lConsulta(2).Top = lConsulta(0).Top
    
    lConsulta(0).Left = tsVistas.ClientLeft
    lConsulta(1).Left = lConsulta(0).Left: lConsulta(2).Left = lConsulta(0).Left
    
    lConsulta(0).Width = tsVistas.ClientWidth
    lConsulta(1).Width = lConsulta(0).Width: lConsulta(2).Width = lConsulta(0).Width
    
    lConsulta(0).Height = tsVistas.ClientHeight - 50
    lConsulta(1).Height = lConsulta(0).Height: lConsulta(2).Height = lConsulta(0).Height
    
    fFiltros.Width = tsVistas.Width
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    RsConsulta.Close
    CierroConexion
    
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    
    GuardoSeteoForm Me
    
End Sub

Private Sub Label1_Click()
    Foco tArticulo
End Sub

Private Sub Label2_Click()
    Foco tTipo
End Sub

Private Sub Label3_Click()
    Foco tGrupo
End Sub

Private Sub Label4_Click()
    Foco tProveedor
End Sub

Private Sub Label5_Click()
    Foco tVapor
End Sub

Private Sub Label6_Click()
    Foco tFactura
End Sub

Private Sub Label7_Click()
    Foco tLC
End Sub

Private Sub lConsulta_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo Seleccione
    'Set lConsulta(Index).SelectedItem = lConsulta(Index).HitTest(X, Y)
    gTxtCarpeta = lConsulta(Index).Cell(flexcpText, lConsulta(Index).Row, 0)
    gTxtCarpeta = Mid(gTxtCarpeta, 1, InStr(gTxtCarpeta, ".") - 1)
    If lConsulta(Index).Rows > 1 And Button = vbRightButton Then PopupMenu MnuMouse
    Exit Sub
    
Seleccione:
End Sub

Private Sub AccionConsultar()
    
    If Not VerificoFiltros Then Exit Sub
    On Error Resume Next: RsConsulta.Close
    On Error GoTo errConsultar
    
    Screen.MousePointer = 11
    PosLista = 0
    ArmoConsulta
    Set RsConsulta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    CargoLista
    bPrimero.Enabled = False: bAnterior.Enabled = False
    If RsConsulta.EOF Then
        bSiguiente.Enabled = False
    Else
        bSiguiente.Enabled = True
    End If
    
    Screen.MousePointer = 0
    Exit Sub

errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub ArmoConsulta()
    
    Cons = " From ArticuloFolder, " _
                                 & " Embarque left outer join Transporte on EmbTransporte = TraCodigo " _
                                        & " left outer join AgenciaTransporte on EmbAgencia = ATrCodigo, " _
                                 & " Carpeta left outer join BancoLocal on CarBcoEmisor = BLoCodigo, " _
                                 & " Articulo" _
            & " Where AFoTipo = " & Folder.cFEmbarque _
            & " And CarFAnulada Is Null And AFoCodigo = EmbID" _
            & " And AFoArticulo = ArtID " _
            & " And EmbCarpeta = CarID"
        
    If tArticulo.prm_ArtID > 0 Then Cons = Cons & " And AFoArticulo = " & tArticulo.prm_ArtID
    
    If Val(tTipo.Tag) > 0 Then Cons = Cons & " And ArtTipo =" & tTipo.Tag
    
    If Val(tProveedor.Tag) > 0 Then Cons = Cons & " And CarProveedor = " & Val(tProveedor.Tag)
    
    If tLC.Text <> "" Then Cons = Cons & " And CarCartaCredito = '" & tLC.Text & "'"
    
    If Trim(tFactura.Text) <> "" Then Cons = Cons & " And CarFactura LIKE '" & tFactura.Text & "%'"
    
    If Val(tAgencia.Tag) > 0 Then Cons = Cons & " And EmbAgencia = " & Val(tAgencia.Tag)
    
    If cboPrioridad.ListIndex > -1 Then Cons = Cons & " And EmbPrioridad = " & Val(cboPrioridad.ItemData(cboPrioridad.ListIndex))
    
    Select Case chCosteada.Value
        Case vbChecked: Cons = Cons & " And EmbCosteado = 1"
        Case vbUnchecked: Cons = Cons & " And EmbCosteado = 0"
    End Select
    
    ArmoConsultaFechas
    
    If Val(tVapor.Tag) <> 0 Then Cons = Cons & " And EmbTransporte = " & Val(tVapor.Tag)
    
    If Val(tGrupo.Tag) > 0 Then
        Cons = Cons & " And ArtID IN (Select AGrArticulo From ArticuloGrupo Where AGrGrupo = " & Val(tGrupo.Tag) & ")"
    End If
    
    If IsNumeric(tTop.Text) Then
        Cons = Cons & " And CarID IN(Select Top " & tTop.Text & " CarID " & Cons & " Group by CarID Order by CarID Desc)"
    End If
    Cons = "Select * " & Cons
    Cons = Cons & " Order by AFoCodigo DESC"
    
End Sub

Private Sub ArmoConsultaFechas()
    
    If aFechas <> "" Then       'Si hay fechas ingresadas
        
        Select Case chArribado.Value
            Case vbChecked: Cons = Cons & ConsultaDeFecha(" And", "EmbFArribo", aFechas)                    'Arribo --------------
            Case vbUnchecked: Cons = Cons & ConsultaDeFecha(" And Not ", "EmbFArribo", aFechas)          'NO Arribo --------------
        End Select
        
        Select Case chEmbarcado.Value
            Case vbChecked: Cons = Cons & ConsultaDeFecha(" And", "EmbFEmbarque", aFechas)                    'Embarcado --------------
            Case vbUnchecked: Cons = Cons & ConsultaDeFecha(" And Not ", "EmbFEmbarque", aFechas)          'NO Embarcado --------------
        End Select
        
    Else
        Select Case chArribado.Value
            Case vbChecked: Cons = Cons & " And EmbFArribo <> NULL"         'Arribo --------------
            Case vbUnchecked: Cons = Cons & " And EmbFArribo = NULL"        'NO Arribo --------------
        End Select
        
        Select Case chEmbarcado.Value
            Case vbChecked: Cons = Cons & " And EmbFEmbarque <> NULL"          'Embarcado --------------
            Case vbUnchecked: Cons = Cons & " And EmbFEmbarque = NULL"         'NO Embarcado --------------
        End Select
        
    End If
    

End Sub

Private Sub CargoLista()

Dim aValor As Long
Dim aMonedaID As Long, aMonedaTXT As String: aMonedaID = 0

    On Error GoTo ErrInesperado
    
    lConsulta(0).Rows = 1: lConsulta(1).Rows = 1: lConsulta(2).Rows = 1
    
    If RsConsulta.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Do While Not RsConsulta.EOF And lConsulta(0).Rows < PorPantalla
        'Lista (0)
        lConsulta(0).AddItem RsConsulta!CarCodigo & "." & RsConsulta!EmbCodigo
        aValor = RsConsulta!EmbId: lConsulta(0).Cell(flexcpData, lConsulta(0).Rows - 1, 0) = aValor
        
        lConsulta(0).Cell(flexcpText, lConsulta(0).Rows - 1, 1) = RsConsulta!AFoCantidad
        lConsulta(0).Cell(flexcpText, lConsulta(0).Rows - 1, 2) = Trim(RsConsulta!ArtNombre)
        
        'Lista (1)
        With lConsulta(1)
            .AddItem RsConsulta!CarCodigo & "." & RsConsulta!EmbCodigo
            .Cell(flexcpData, .Rows - 1, 0) = aValor
        
            .Cell(flexcpText, .Rows - 1, 1) = RsConsulta!AFoCantidad
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsConsulta!ArtNombre)
        End With
        
        'Lista (2)
        lConsulta(2).AddItem RsConsulta!CarCodigo & "." & RsConsulta!EmbCodigo
        lConsulta(2).Cell(flexcpData, lConsulta(2).Rows - 1, 0) = aValor
        
        lConsulta(2).Cell(flexcpText, lConsulta(2).Rows - 1, 1) = RsConsulta!AFoCantidad
        lConsulta(2).Cell(flexcpText, lConsulta(2).Rows - 1, 2) = Trim(RsConsulta!ArtNombre)
        
        'Lista 0 -Vista 1
        With lConsulta(0)
            If Not IsNull(RsConsulta!BLoNombre) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsConsulta!BLoNombre)
            If Not IsNull(RsConsulta!CarCartaCredito) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(RsConsulta!CarCartaCredito)
            If Not IsNull(RsConsulta!CarFactura) Then .Cell(flexcpText, .Rows - 1, 5) = Trim(RsConsulta!CarFactura)
            
            If Not IsNull(RsConsulta!EmbMoneda) Then
                If aMonedaID <> RsConsulta!EmbMoneda Then
                    Cons = "Select * from Moneda Where MonCodigo = " & RsConsulta!EmbMoneda
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    aMonedaTXT = ""
                    aMonedaID = RsConsulta!EmbMoneda
                    If Not RsAux.EOF Then If Not IsNull(RsAux!MonSigno) Then aMonedaTXT = Trim(RsAux!MonSigno)
                    RsAux.Close
                End If
                .Cell(flexcpText, .Rows - 1, 6) = aMonedaTXT & " " & Format(RsConsulta!AFoPUnitario, "##,##0.00")
            End If
        End With
        
        'Lista 1 - Vista 2
        With lConsulta(1)
            If Not IsNull(RsConsulta!ATrNombre) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsConsulta!ATrNombre)     'Agencia Transporte
            If Not IsNull(RsConsulta!TraNombre) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(RsConsulta!TraNombre)
            If Not IsNull(RsConsulta!EmbFEPrometido) Then .Cell(flexcpText, .Rows - 1, 5) = Trim(Format(RsConsulta!EmbFEPrometido, "dd/mm/yyyy"))
            If Not IsNull(RsConsulta!EmbFEmbarque) Then .Cell(flexcpText, .Rows - 1, 6) = Trim(Format(RsConsulta!EmbFEmbarque, "dd/mm/yyyy"))
            
            If Not IsNull(RsConsulta!EmbUltFechaEmbarque) Then .Cell(flexcpText, .Rows - 1, 8) = Trim(Format(RsConsulta!EmbUltFechaEmbarque, "dd/mm/yyyy"))
        End With
        
        If Not IsNull(RsConsulta!EmbFAPrometido) Then
            lConsulta(1).Cell(flexcpText, lConsulta(1).Rows - 1, 7) = Trim(Format(RsConsulta!EmbFAPrometido, "dd/mm/yyyy"))
            lConsulta(2).Cell(flexcpText, lConsulta(2).Rows - 1, 3) = Trim(Format(RsConsulta!EmbFAPrometido, "dd/mm/yyyy"))
            If Not IsNull(RsConsulta!EmbFEmbarque) Then lConsulta(2).Cell(flexcpText, lConsulta(2).Rows - 1, 9) = DateDiff("d", RsConsulta!EmbFEmbarque, RsConsulta!EmbFAPrometido)
        End If
        
        With lConsulta(2)
            If Not IsNull(RsConsulta!EmbFArribo) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(Format(RsConsulta!EmbFArribo, "dd/mm/yyyy"))
            If RsConsulta!EmbLocal = paLocalZF Then .Cell(flexcpText, .Rows - 1, 5) = "Si" Else .Cell(flexcpText, .Rows - 1, 5) = "No"
            If Not IsNull(RsConsulta!EmbFlete) Then .Cell(flexcpText, .Rows - 1, 6) = Format(RsConsulta!EmbFlete, "##,##0.00")
            If RsConsulta!EmbFletePago Then .Cell(flexcpText, .Rows - 1, 7) = "Si" Else .Cell(flexcpText, .Rows - 1, 7) = "No"
            If RsConsulta!EmbCosteado Then .Cell(flexcpText, .Rows - 1, 8) = "Si" Else .Cell(flexcpText, .Rows - 1, 8) = "No"
        End With
        
        RsConsulta.MoveNext
    Loop
    
    For I = 0 To 2
        lConsulta(I).SubtotalPosition = flexSTBelow
        lConsulta(I).Subtotal flexSTSum, -1, 1, "0", 128, vbWindowBackground, True, "Total"
    Next
    
    If RsConsulta.EOF Then bSiguiente.Enabled = False
    Screen.MousePointer = 0
    Exit Sub
    
ErrInesperado:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al realizar la consulta de embarques.", Err.Description
End Sub

Private Sub AccionImprimir()
    
    If lConsulta(0).Rows = 1 Then
        MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo errPrint
    Screen.MousePointer = 11
    
    With vsListado
    
        If Not .PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 0: Exit Sub
        
        .Preview = True
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
    
        EncabezadoListado vsListado, "Importaciones - Consulta de Embarques", False
        
        .FileName = "Consulta de Embarques"
        .FontSize = 8: .FontBold = False
        
        .Paragraph = tsVistas.SelectedItem.Tag
        .Paragraph = "Filtros: " & ArmoFormulaFiltros
        
        lConsulta(tsVistas.SelectedItem.Index - 1).ExtendLastCol = False
        .RenderControl = lConsulta(tsVistas.SelectedItem.Index - 1).hWnd
        lConsulta(tsVistas.SelectedItem.Index - 1).ExtendLastCol = True
        
        .EndDoc
        .PrintDoc
        
    End With
    
    Screen.MousePointer = 0
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al realizar la impresión. " & Trim(Err.Description)
End Sub

Private Sub tFactura_Change()
    lConsulta(0).Rows = 1: lConsulta(1).Rows = 1: lConsulta(2).Rows = 1
End Sub

Private Sub tFactura_GotFocus()
    tFactura.SelStart = 0
    tFactura.SelLength = Len(tFactura.Text)
End Sub

Private Sub tFactura_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then tGrupo.SetFocus

End Sub

Private Function VerificoFiltros() As Boolean

    VerificoFiltros = False
    
    If tTipo.Text <> "" And Val(tTipo.Tag) = 0 Then
        MsgBox "El tipo de artículo ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tTipo: Exit Function
    End If
    
    If tGrupo.Text <> "" And Val(tGrupo.Tag) = 0 Then
        MsgBox "El grupo de artículos ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tGrupo: Exit Function
    End If
    
    If tArticulo.Text <> "" And tArticulo.prm_ArtID = 0 Then
        MsgBox "El artículo seleccionado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tArticulo: Exit Function
    End If
    
    If tVapor.Text <> "" And Val(tVapor.Tag) = 0 Then
        MsgBox "Hay errores en el vapor seleccionado o no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tVapor: Exit Function
    End If
    
    If tAgencia.Text <> "" And Val(tAgencia.Tag) = 0 Then
        MsgBox "Hay errores en la Agencia seleccionada o no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tAgencia: Exit Function
    End If
    
    If tProveedor.Text <> "" And Val(tProveedor.Tag) = 0 Then
        MsgBox "Hay errores en el proveedor seleccionado o no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Function
    End If
    
    If tLC.Text <> "" And tLC.Tag = "" Then
        MsgBox "Hay errores en la carta de crédito seleccionada o no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tLC: Exit Function
    End If
    
    aFechas = ""
    If Trim(tFecha.Text) <> "" Then
        aFechas = ValidoPeriodoFechas(Trim(tFecha.Text))
        If aFechas = "" Then
            MsgBox "Hay errores en el formato de fechas ingresado, o no es correcto.", vbExclamation, "ATENCIÓN"
            Foco tFecha: Exit Function
        End If
        aFechas = Trim(tFecha.Text)
    End If
    
    If IsNumeric(tTop.Text) Then
        If Val(tTop.Text) < 1 Then
            MsgBox "Para ver los últimos embarques el valor tiene que ser mayor que cero.", vbExclamation, "ATENCIÓN"
            Foco tTop: Exit Function
        End If
    End If
    VerificoFiltros = True

End Function

Private Sub tsVistas_Click()
    lConsulta(tsVistas.SelectedItem.Index - 1).ZOrder 0
    Me.Refresh
End Sub

Private Function EstadoCheck(chBox As CheckBox)

    Select Case chBox.Value
        Case vbGrayed:  chBox.Value = vbChecked
        Case vbChecked:  chBox.Value = vbUnchecked
        Case vbUnchecked:  chBox.Value = vbGrayed
    End Select
    If chBox.Name = "chCosteada" Then Exit Function
    If chEmbarcado.Value = vbGrayed And chArribado = vbGrayed Then
        tFecha.BackColor = inactivo: tFecha.Enabled = False: tFecha.Text = ""
    Else
        tFecha.BackColor = vbWindowBackground: tFecha.Enabled = True
    End If
    
End Function

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With lConsulta(0)
        .Cols = 1: .Rows = 1:
        .FormatString = "<Folder|>Q|Artículos|Banco|LC|<Factura|>Precio|"
            
        .WordWrap = False
        .ColWidth(0) = 700: .ColWidth(1) = 500: .ColWidth(2) = 2400: .ColWidth(3) = 1800: .ColWidth(4) = 900: .ColWidth(5) = 900: .ColWidth(6) = 1400
        
    End With
      
    With lConsulta(1)
        .Cols = 1: .Rows = 1:
        .FormatString = "<Folder|>Q|Artículos|Agencia|Vapor|Embarca|Embarcó|Arriba|Últ. Embarque|"
            
        .WordWrap = False
        .ColWidth(0) = 700: .ColWidth(1) = 500: .ColWidth(2) = 2240: .ColWidth(3) = 1210: .ColWidth(4) = 950: .ColWidth(5) = 950: .ColWidth(6) = 950: .ColWidth(7) = 950
        .ColDataType(5) = flexDTDate: .ColDataType(6) = flexDTDate: .ColDataType(7) = flexDTDate: .ColDataType(8) = flexDTDate
    End With
    
    With lConsulta(2)
        .Cols = 1: .Rows = 1:
        .FormatString = "<Folder|>Q|Artículos|Arriba|Arribó|ZF|>Flete|^Pago|^Cost.|>TransitTime|"
            
        .WordWrap = False
        .ColWidth(0) = 700: .ColWidth(1) = 500: .ColWidth(2) = 2240: .ColWidth(3) = 950: .ColWidth(4) = 950: .ColWidth(5) = 400: .ColWidth(6) = 1400
        .ColWidth(7) = 500: .ColWidth(8) = 500: .ColWidth(9) = 1000
        .ColDataType(3) = flexDTDate
    End With
      
End Sub

Private Function ConsultaDeFecha(Condicion As String, Campo As String, CadenaFecha As String) As String
Dim aStr As String

    Condicion = Trim(Condicion)
    Campo = Trim(Campo)
    
    If IsDate(CadenaFecha) Then         'Igual a una Fecha
       aStr = " " & Condicion & " " & Campo _
              & " Between '" & Format(CadenaFecha, "mm/dd/yyyy") & "'" _
              & " And '" & Format(CadenaFecha, "mm/dd/yyyy 23:59") & "'"
    Else
        'Mayor o menor a una fecha
       If Mid(CadenaFecha, 1, 1) = ">" Or Mid(CadenaFecha, 1, 1) = "<" Then
            If Mid(CadenaFecha, 1, 1) = ">" Then
               aStr = " " & Condicion & " " & Campo & " > '" & Format(Mid(CadenaFecha, 2, 10), "mm/dd/yyyy") & "'"
            Else
               aStr = " " & Condicion & " " & Campo & " < '" & Format(Mid(CadenaFecha, 2, 10), "mm/dd/yyyy") & "'"
            End If
       Else
            'Entre una fecha y tal otra
            CadenaFecha = ValidoPeriodoFechas(CadenaFecha)
            aStr = " " & Condicion & " " & Campo _
                   & " Between '" & Format(Mid(CadenaFecha, 1, 10), "mm/dd/yyyy 00:00") & "'" _
                   & " And '" & Format(Mid(CadenaFecha, 11, 10), "mm/dd/yyyy 23:59") & "'"
       End If
    End If
    ConsultaDeFecha = aStr
    
End Function

Private Function ValidoPeriodoFechas(Cadena As String)
    
Dim aS1 As String
Dim aS2 As String

    ValidoPeriodoFechas = ""
    Cadena = UCase(Cadena)
    If IsDate(Cadena) Then
        ValidoPeriodoFechas = Format(Cadena, "dd/mm/yyyy")
        Exit Function
    End If
    
    If Mid(Cadena, 1, 1) = ">" Or Mid(Cadena, 1, 1) = "<" Then
        If IsDate(Mid(Cadena, 2, Len(Cadena))) Then
            ValidoPeriodoFechas = Mid(Cadena, 1, 1) & Format(Mid(Cadena, 2, Len(Cadena)), "dd/mm/yyyy")
            Exit Function
        End If
    End If
    
    If Mid(Cadena, 1, 1) = "E" Then
        If InStr(Cadena, "Y") <> 0 Then
            If IsDate(Mid(Cadena, 2, InStr(Cadena, "Y") - 2)) Then
                aS1 = Mid(Cadena, 2, InStr(Cadena, "Y") - 2)
                If IsDate(Mid(Cadena, InStr(Cadena, "Y") + 1, Len(Cadena))) Then
                    aS2 = Mid(Cadena, InStr(Cadena, "Y") + 1, Len(Cadena))
                    ValidoPeriodoFechas = Format(aS1, "dd/mm/yyyy") & Format(aS2, "dd/mm/yyyy")
                    Exit Function
                End If
            End If
        End If
    End If
    
End Function

Private Function ArmoFormulaFiltros() As String
Dim aRetorno As String

    On Error Resume Next
    aRetorno = ""
    
    If tArticulo.prm_ArtID <> 0 Then aRetorno = aRetorno & "Art.: " & Trim(tArticulo.Text) & ", "
    If Val(tTipo.Tag) > 0 Then aRetorno = aRetorno & "Tipo: " & Trim(tTipo.Text) & ", "
    If Val(tProveedor.Tag) > 0 Then aRetorno = aRetorno & "Proveedor: " & Trim(tProveedor.Text) & ", "
    If tLC.Tag <> "" Then aRetorno = aRetorno & "LC: " & Trim(tLC.Text) & ", "
    If Trim(tFactura.Text) <> "" Then aRetorno = aRetorno & "Fact.: " & Trim(tFactura.Text) & ", "
    If Val(tVapor.Tag) <> 0 Then aRetorno = aRetorno & "Transp.: " & tVapor.Text & ", "
    If Val(tGrupo.Tag) > 0 Then aRetorno = aRetorno & "Grupo: " & tGrupo.Text & ", "
    
    Select Case chCosteada.Value
        Case vbChecked: aRetorno = aRetorno & "Costeados, "
        Case vbUnchecked: aRetorno = aRetorno & "No Costeados, "
    End Select
    
    Select Case chArribado.Value
        Case vbChecked: aRetorno = aRetorno & "Arribados "
        Case vbUnchecked: aRetorno = aRetorno & "No Arribados "
    End Select
    
    Select Case chEmbarcado.Value
        Case vbChecked: aRetorno = aRetorno & "Embarcados "
        Case vbUnchecked: aRetorno = aRetorno & "No Embarcados "
    End Select
    
    If Trim(tFecha.Text) <> "" Then aRetorno = aRetorno & "(" & Trim(tFecha.Text) & "), "
    
    aRetorno = Mid(aRetorno, 1, Len(aRetorno) - 2)
    ArmoFormulaFiltros = aRetorno

End Function

Private Sub loc_FindTipo()
On Error GoTo errFT
    Screen.MousePointer = 11
    Cons = "Select TipCodigo, TipNombre as 'Nombre' From Tipo"
    If IsNumeric(tTipo.Text) Then
        Cons = Cons & " Where TipCodigo = " & Val(tTipo.Text)
    Else
        Cons = Cons & " Where TipNombre Like '" & fnc_GetQueryLike(tTipo.Text) & "%' Order by TipNombre"
    End If
    loc_FindByName tTipo, Cons, 1, "Tipos", 4500
    Screen.MousePointer = 0
    Exit Sub
errFT:
    clsGeneral.OcurrioError "Error al buscar.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_FindGrupo()
On Error GoTo errFT
    Screen.MousePointer = 11
    Cons = "Select GruCodigo, GruNombre as 'Nombre'  from Grupo "
    If IsNumeric(tTipo.Text) Then
        Cons = Cons & " Where GruCodigo = " & Val(tGrupo.Text)
    Else
        Cons = Cons & " Where GruNombre Like '" & fnc_GetQueryLike(tGrupo.Text) & "%'"
    End If
    loc_FindByName tGrupo, Cons, 1, "Grupos", 4500
    Screen.MousePointer = 0
    Exit Sub
errFT:
    clsGeneral.OcurrioError "Error al buscar.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_FindLC()
On Error GoTo errFT
    Screen.MousePointer = 11
    Cons = "Select count(*), CarCartaCredito as 'L/C'  From Carpeta Where CarCartaCredito Like '" & fnc_GetQueryLike(tLC.Text) & "%' Group By CarCartaCredito"
    loc_FindByName tLC, Cons, 1, "Carta de Crédito", 3000
    Screen.MousePointer = 0
    Exit Sub
errFT:
    clsGeneral.OcurrioError "Error al buscar.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_FindProveedor()
On Error GoTo errFT
    Screen.MousePointer = 11
    Cons = "Select PExCodigo, PExNombre as 'Proveedor' from ProveedorExterior Where PExNombre  Like '" & fnc_GetQueryLike(tProveedor.Text) & "%'" & _
                "Order by PExNombre"
    loc_FindByName tProveedor, Cons, 1, "Proveedores", 4000
    Screen.MousePointer = 0
    Exit Sub
errFT:
    clsGeneral.OcurrioError "Error al buscar.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function fnc_GetQueryLike(ByVal sName As String) As String
    sName = Replace(sName, "  ", " ", Compare:=vbTextCompare)
    fnc_GetQueryLike = Replace(sName, " ", "%", Compare:=vbTextCompare)
End Function

Private Sub loc_FindByName(tControl As TextBox, ByVal sQuery As String, ByVal iColOculta As Integer, ByVal sTitulo As String, Optional cWidth As Currency = 6000)
    Dim objLista As New clsListadeAyuda
    If objLista.ActivarAyuda(cBase, sQuery, cWidth, iColOculta, sTitulo) > 0 Then
        With tControl
            .Text = objLista.RetornoDatoSeleccionado(1)
            .Tag = objLista.RetornoDatoSeleccionado(0)
        End With
    End If
End Sub
