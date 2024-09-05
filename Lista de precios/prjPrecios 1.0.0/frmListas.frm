VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmListas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listas de Precios y Etiquetas"
   ClientHeight    =   4965
   ClientLeft      =   3450
   ClientTop       =   4110
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6240
   Begin VB.PictureBox picLista 
      Height          =   2715
      Index           =   1
      Left            =   780
      ScaleHeight     =   2655
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   2280
      Width           =   4575
      Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
         Height          =   975
         Left            =   60
         TabIndex        =   25
         Top             =   120
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1720
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
         BackColorBkg    =   -2147483643
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
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
   End
   Begin VB.PictureBox picLista 
      Height          =   3795
      Index           =   2
      Left            =   120
      ScaleHeight     =   3735
      ScaleWidth      =   5955
      TabIndex        =   27
      Top             =   960
      Width           =   6015
      Begin VB.CommandButton bPrintEtiqueta 
         Caption         =   "Imprimir"
         Height          =   315
         Left            =   4800
         TabIndex        =   20
         Top             =   2880
         Width           =   1035
      End
      Begin VB.CommandButton bFiltrarEtiqueta 
         Caption         =   "..."
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   3480
         Width           =   375
      End
      Begin VB.ComboBox cEtiquetaAImprimir 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton bAgregar 
         Caption         =   "A&gregar"
         Height          =   315
         Left            =   4800
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cQueEtiqueta 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.VScrollBar vscCantidad 
         Height          =   285
         Left            =   4080
         Min             =   1
         TabIndex        =   28
         Top             =   480
         Value           =   1
         Width           =   255
      End
      Begin VB.TextBox tCantidad 
         Height          =   285
         Left            =   3600
         TabIndex        =   15
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox tArticulo 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   120
         Width           =   3735
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsEtiquetaArt 
         Height          =   1935
         Left            =   60
         TabIndex        =   17
         Top             =   840
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3413
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
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
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
      Begin VB.Label lFiltro 
         Caption         =   "Agregar por &Filtros:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   5760
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label8 
         Caption         =   "Et&iqueta a imprimir:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "¿C&uales?:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "&Cantidad:"
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComctlLib.TabStrip tabLista 
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   660
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Público"
            Key             =   "definidas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Distribuidores"
            Key             =   "varias"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Etiquetas"
            Key             =   "etiquetas"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLista 
      Height          =   2055
      Index           =   0
      Left            =   240
      ScaleHeight     =   1995
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   1800
      Width           =   4455
      Begin VB.CommandButton bContado 
         Caption         =   "Imprimir"
         Height          =   315
         Left            =   3540
         Picture         =   "frmListas.frx":0442
         TabIndex        =   6
         Top             =   60
         Width           =   1095
      End
      Begin VB.ComboBox cCategoria 
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CommandButton bContadoDto 
         Caption         =   "Imprimir"
         Height          =   315
         Left            =   3540
         Picture         =   "frmListas.frx":0974
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría de Cliente:"
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Listas para Distribuidores (Precios Contado con Descuentos)"
         Height          =   255
         Left            =   60
         TabIndex        =   23
         Top             =   660
         Width           =   4815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Listas para Distribuidores (Precios Contado)"
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   120
         Width           =   3315
      End
   End
   Begin VB.TextBox tVigencia 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lSep 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      Height          =   15
      Left            =   120
      TabIndex        =   26
      Top             =   480
      Width           =   5895
   End
   Begin VB.Label lVigencia 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de &Vigencia:"
      Height          =   255
      Left            =   2820
      TabIndex        =   2
      Top             =   180
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de &Vigencia:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1755
   End
End
Attribute VB_Name = "frmListas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer

Enum TReporte
    Contado = 1
    ContadoConDto = 2
    AlPublico = 3
    EtiquetaNormal = 4
    EtiquetaConArgumento = 5
    EtiquetaSinArgumento = 6
End Enum

'Variables para Crystal Engine.---------------------------------
Private Result As Integer, JobSRep1 As Integer, JobSRep2 As Integer, jobnum As Integer
Private NombreFormula As String, CantForm As Integer, aTexto As String

Private prmVigencia As String

' <VB WATCH>
Const VBWMODULE = "frmListas"
' </VB WATCH>

Private Sub AccionImprimir(idReporte As Integer)
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>

2          If Not IsDate(tVigencia.Text) Then
3              MsgBox "La fecha de vigencia no es correcta.", vbExclamation, "Datos Incorrectos"
4              tVigencia.SetFocus
5              Exit Sub
6          End If

7          prmVigencia = Format(tVigencia.Text, "mm/dd/yyyy 23:59:59")

8          Select Case idReporte
               Case TReporte.Contado
9                              If InicializoReporteEImpresora("", 1, "lprListaContado.RPT") Then
10                                  Exit Sub
11                             End If
12                             rptListaContado
13                             If Not crCierroTrabajo(jobnum) Then
14                                  MsgBox crMsgErr
15                             End If

16             Case TReporte.ContadoConDto
17                             If InicializoReporteEImpresora("", 1, "lprListaConDescuentos.RPT") Then
18                                  Exit Sub
19                             End If
20                             rptListaContadoCategoria
21                             If Not crCierroTrabajo(jobnum) Then
22                                  MsgBox crMsgErr
23                             End If

24             Case TReporte.AlPublico
25                             rptListaAlPublico

26             Case TReporte.EtiquetaNormal
27                             If InicializoReporteEImpresora("", 1, "lprEtiquetaNormal.RPT") Then
28                                  Exit Sub
29                             End If
30                             rptImprimoEtiquetas "Etiquetas Normales"
31                             If Not crCierroTrabajo(jobnum) Then
32                                  MsgBox crMsgErr
33                             End If

34             Case TReporte.EtiquetaConArgumento
35                             If InicializoReporteEImpresora("", 1, "lprEtiquetaArgumento.RPT", 2) Then
36                                  Exit Sub
37                             End If
38                             rptImprimoEtiquetas " Etiquetas Vidriera con Argumento"
39                             If Not crCierroTrabajo(jobnum) Then
40                                  MsgBox crMsgErr
41                             End If

42             Case TReporte.EtiquetaSinArgumento
43                             If InicializoReporteEImpresora("", 1, "lprEtiquetaSinArgumento.RPT", 2) Then
44                                  Exit Sub
45                             End If
46                             rptImprimoEtiquetas " Etiquetas de Vidriera"
47                             If Not crCierroTrabajo(jobnum) Then
48                                  MsgBox crMsgErr
49                             End If
50         End Select

51         Screen.MousePointer = 0
' <VB WATCH>
52         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "AccionImprimir"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub bAgregar_Click()
' <VB WATCH>
53         On Error GoTo vbwErrHandler
' </VB WATCH>
54         If Val(tArticulo.Tag) > 0 And cQueEtiqueta.ListIndex <> -1 Then
55             If tCantidad.Enabled Then
56                 If Val(tCantidad.Text) = 0 Then
57                     MsgBox "La cantidad no puede ser cero.", vbExclamation, "ATENCIÓN"
58                     Exit Sub
59                 End If
60             End If
61             etiqueta_AgregoArticuloALista Val(tArticulo.Tag), Trim(tArticulo.Text), tCantidad.Text, cQueEtiqueta.ListIndex
62             LimpioIngresoArticuloEtiqueta
63             tArticulo.SetFocus
64         Else
65             MsgBox "Falta ingresar algún dato.", vbExclamation, "ATENCIÓN"
66         End If
' <VB WATCH>
67         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "bAgregar_Click"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub bContado_Click()
' <VB WATCH>
68         On Error GoTo vbwErrHandler
' </VB WATCH>
69         AccionImprimir TReporte.Contado
' <VB WATCH>
70         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "bContado_Click"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub bContadoDto_Click()
' <VB WATCH>
71         On Error GoTo vbwErrHandler
' </VB WATCH>

72         If cCategoria.ListIndex = -1 Then
73             MsgBox "Seleccione la categoría de cliente para sacar la lista de precios.", vbExclamation, "Falta Categoría de Cliente"
74             tVigencia.SetFocus
75             Exit Sub
76         End If

77         AccionImprimir TReporte.ContadoConDto

' <VB WATCH>
78         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "bContadoDto_Click"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub bFiltrarEtiqueta_Click()
' <VB WATCH>
79         On Error GoTo vbwErrHandler
' </VB WATCH>
80     On Error GoTo errBFE
81     Dim lEsta As Long
82     Dim iCant As Integer
83     Dim sNombre As String, lid As Long
84         frmFiltroEtiqueta.Show vbModal, Me
85         If frmFiltroEtiqueta.prmHayDatos Then
86             Screen.MousePointer = 11
87             For iCant = 1 To frmFiltroEtiqueta.prmCantResultado
88                 BuscoArticuloPorCodigo frmFiltroEtiqueta.prmIDResultado(iCant), sNombre, lid
89                 If lid > 0 Then
90                     For lEsta = 1 To vsEtiquetaArt.Rows - 1
91                         If Val(vsEtiquetaArt.Cell(flexcpData, lEsta, 0)) = lid Then
92                             lid = 0
93                             Exit For
94                         End If
95                     Next
96                 End If
97                 If lid > 0 Then
98                     etiqueta_AgregoArticuloALista lid, sNombre, frmFiltroEtiqueta.prmCantidad, frmFiltroEtiqueta.prmQueEtiqueta
99                 End If
100            Next
101            Screen.MousePointer = 0
102        End If
103        Set frmFiltroEtiqueta = Nothing
104        Screen.MousePointer = 0
105        Exit Sub
106    errBFE:
107        clsGeneral.OcurrioError "Ocurrió un error al filtrar.", Err.Description, "Error (filtraretiqueta)"
108        Screen.MousePointer = 0
' <VB WATCH>
109        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "bFiltrarEtiqueta_Click"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub bPrintEtiqueta_Click()
' <VB WATCH>
110        On Error GoTo vbwErrHandler
' </VB WATCH>

111        If vsEtiquetaArt.Rows = 1 Then
112            MsgBox "No hay artículos ingresados.", vbExclamation, "ATENCIÓN"
113            Exit Sub
114        End If
115        If cEtiquetaAImprimir.ListIndex = -1 Then
116            MsgBox "Seleccione el tipo de etiqueta que desea imprimir.", vbExclamation, "ATENCIÓN"
117            cEtiquetaAImprimir.SetFocus
118        Else
119            Screen.MousePointer = 11
120            etiqueta_MandoAImprimir
121            Screen.MousePointer = 0
122        End If
' <VB WATCH>
123        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "bPrintEtiqueta_Click"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub cEtiquetaAImprimir_KeyPress(KeyAscii As Integer)
' <VB WATCH>
124        On Error GoTo vbwErrHandler
' </VB WATCH>
125    On Error Resume Next
126        If KeyAscii = vbKeyReturn Then
127             bPrintEtiqueta.SetFocus
128        End If
' <VB WATCH>
129        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cEtiquetaAImprimir_KeyPress"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub cQueEtiqueta_Change()
' <VB WATCH>
130        On Error GoTo vbwErrHandler
' </VB WATCH>
131    On Error Resume Next
132        If cQueEtiqueta.ListIndex > -1 Then
133            If cQueEtiqueta.ListIndex = 2 Then
134                tCantidad.Enabled = False
135                tCantidad.BackColor = vbButtonFace
136            Else
137                tCantidad.Enabled = True
138                tCantidad.BackColor = vbWindowBackground
139            End If
140        Else
141            tCantidad.Enabled = False
142            tCantidad.BackColor = vbButtonFace
143        End If
144        vscCantidad.Enabled = tCantidad.Enabled
' <VB WATCH>
145        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cQueEtiqueta_Change"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub cQueEtiqueta_Click()
' <VB WATCH>
146        On Error GoTo vbwErrHandler
' </VB WATCH>
147    On Error Resume Next
148        If cQueEtiqueta.ListIndex > -1 Then
149            If cQueEtiqueta.ListIndex = 2 Then
150                tCantidad.Enabled = False
151                tCantidad.BackColor = vbButtonFace
152            Else
153                tCantidad.Enabled = True
154                tCantidad.BackColor = vbWindowBackground
155            End If
156        Else
157            tCantidad.Enabled = False
158            tCantidad.BackColor = vbButtonFace
159        End If
160        vscCantidad.Enabled = tCantidad.Enabled
' <VB WATCH>
161        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cQueEtiqueta_Click"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub cQueEtiqueta_KeyPress(KeyAscii As Integer)
' <VB WATCH>
162        On Error GoTo vbwErrHandler
' </VB WATCH>
163        If KeyAscii = vbKeyReturn Then
164            If Val(tArticulo.Tag) > 0 Then
165                If tCantidad.Enabled Then
166                    tCantidad.SetFocus
167                Else
168                    bAgregar.SetFocus
169                End If
170            Else
171                vsEtiquetaArt.SetFocus
172            End If
173        End If
' <VB WATCH>
174        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cQueEtiqueta_KeyPress"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub Form_Activate()
' <VB WATCH>
175        On Error GoTo vbwErrHandler
' </VB WATCH>
176        Screen.MousePointer = vbDefault
177        Me.Refresh
' <VB WATCH>
178        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Activate"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub Form_Load()
' <VB WATCH>
179        On Error GoTo vbwErrHandler
' </VB WATCH>

180        On Error Resume Next
181        Me.Top = (Screen.Height - Me.Height) / 2
182        Me.Left = (Screen.Width - Me.Width) / 2

183        crAbroEngine
184        tVigencia.Text = Format(Now, "dd/mm/yyyy")
185        lVigencia.Caption = Format(Now, "Ddd d/Mmm/yyyy")

186        Cons = "Select LDiCodigo, LDiNombre from ListasDistribuidores order by LDiNombre"
187        CargoCombo Cons, cCategoria

188        InicializoGrillas
189        InicializoObjetosEtiqueta

190        picLista(1).ZOrder 0
191        tabLista.SetFocus
' <VB WATCH>
192        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Load"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub Form_Resize()
' <VB WATCH>
193        On Error GoTo vbwErrHandler
' </VB WATCH>
194        On Error Resume Next
195        If Me.WindowState = vbMinimized Then
196             Exit Sub
197        End If

198        With tabLista
199            .Left = 60
200            .Width = Me.ScaleWidth - (.Left * 2)
201            .Top = tVigencia.Top + 600
202            .Height = Me.ScaleHeight - .Top - 60
203        End With

204        With lSep
205            .Left = tabLista.Left
206            .Width = tabLista.Width
207            .Top = tabLista.Top - 150
208        End With

209        For I = picLista.LBound To picLista.UBound
210            With picLista(I)
211                .Left = tabLista.ClientLeft
212                .Top = tabLista.ClientTop
213                .Width = tabLista.ClientWidth
214                .Height = tabLista.ClientHeight
215                .BorderStyle = 0
216            End With
217        Next

218        With vsLista
219            .Top = 60
220            .Left = 60
221            .Width = picLista(0).ScaleWidth - (.Left * 2)
222            .Height = picLista(0).ScaleHeight
223        End With

' <VB WATCH>
224        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Resize"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub Form_Unload(Cancel As Integer)
' <VB WATCH>
225        On Error GoTo vbwErrHandler
' </VB WATCH>
226        On Error Resume Next

227        Screen.MousePointer = 11
228        crCierroEngine
229        CierroConexion
230        Set clsGeneral = Nothing
231        Set miConexion = Nothing
232        Screen.MousePointer = 0

233        End

' <VB WATCH>
234        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Unload"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub rptListaContado()
' <VB WATCH>
235        On Error GoTo vbwErrHandler
' </VB WATCH>
236    On Error GoTo ErrCrystal
237    Dim I As Integer
238    Dim dRige As Date

239        Screen.MousePointer = 11

           'Saco la máxima fecha de vigencia para cargar valor Rige
240        dRige = Now

241        Cons = "Select Max(Precios.HPrVigencia) from HistoriaPrecio Precios, Articulo, ArticuloFacturacion " & _
                      " Where (Precios.HPrVigencia IN " & _
                           " (Select MAX(H.HPrVigencia)" & _
                                   " FROM HistoriaPrecio H" & _
                                   " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                                   " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                                   " And H.HPrMoneda = Precios.HPrMoneda " & _
                                   " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                                   " )) " & _
                       " And Precios.HPrArticulo = ArtID" & _
                       " And ArtId = AFaArticulo " & _
                       " And ArtEnUso = 1 And AFaInterior = 1 " & _
                       " And Precios.HPrTipoCuota = " & paTipoCuotaContado & _
                       " And Precios.HPrMoneda = " & paMonedaPesos & _
                       " And Precios.HPrHabilitado = 1"
242        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
243        If Not RsAux.EOF Then
244             dRige = RsAux(0)
245        End If
246        RsAux.Close
           '--------------------------------------------------------------------------------------------------------------------------------------------

           'Obtengo la cantidad de formulas que tiene el reporte.----------------------
247        CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
248        If CantForm = -1 Then
249             GoTo ErrCrystal
250        End If

           'Cargo Propiedades para el reporte Contado --------------------------------
251        For I = 0 To CantForm - 1
252            NombreFormula = crObtengoNombreFormula(jobnum, I)

253            Select Case LCase(NombreFormula)
                   Case ""
254                GoTo ErrCrystal
255                Case "prmvigencia"
256                Result = crSeteoFormula(jobnum%, NombreFormula, "'Rige desde el: " & Format(dRige, "dd/Mmm/yyyy") & "'")

257                Case Else
258                Result = 1
259            End Select
260            If Result = 0 Then
261                 GoTo ErrCrystal
262            End If
263        Next
           '--------------------------------------------------------------------------------------------------------------------------------------------

           'Seteo la Query del reporte-----------------------------------------------------------------
264        Cons = "Select Articulo.ArtCodigo, Articulo.ArtNombre, Articulo.ArtHabilitado," & _
                               " ArticuloFacturacion.AFaInterior, ArticuloFacturacion.AFaArgumCorto," & _
                               " Precios.HPrPrecio, Especie.EspNombre" & _
                      " From " & _
                               paBD & ".dbo.HistoriaPrecio Precios, " & _
                               paBD & ".dbo.Articulo, " & paBD & ".dbo.Tipo, " & paBD & ".dbo.ArticuloFacturacion, " & paBD & ".dbo.Especie" & _
                       " WHERE (HPrVigencia IN" & _
                                   " (Select MAX(H.HPrVigencia)" & _
                                   " FROM HistoriaPrecio H" & _
                                   " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                                   " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                                   " And H.HPrMoneda = Precios.HPrMoneda " & _
                                   " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                                   " )) " & _
                       " And Precios.HPrArticulo = ArtID" & _
                       " And ArtId = AFaArticulo " & _
                       " And ArtTipo = TipCodigo " & _
                       " And TipEspecie = EspCodigo" & _
                       " And ArtEnUso = 1 And AFaInterior = 1 " & _
                       " And Precios.HPrTipoCuota = " & paTipoCuotaContado & _
                       " And Precios.HPrMoneda = " & paMonedaPesos & _
                       " And Precios.HPrHabilitado = 1"

265                    Cons = Cons & " Order By Especie.EspNombre Asc, Articulo.ArtNombre Asc"

266        Cons = Trim(Cons) & Chr$(0)

267        If crSeteoSqlQuery(jobnum%, Cons) = 0 Then
268             GoTo ErrCrystal
269        End If
           '-------------------------------------------------------------------------------------------------------------------------------------

270        If crMandoAPantalla(jobnum, "Lista de Contados") = 0 Then
271             GoTo ErrCrystal
272        End If
           'If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
273        If Not crInicioImpresion(jobnum, True, False) Then
274             GoTo ErrCrystal
275        End If

       '    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal

       '    crEsperoCierreReportePantalla
276        Screen.MousePointer = 0
277        Exit Sub

278    ErrCrystal:
279        Screen.MousePointer = 0
280        clsGeneral.OcurrioError crMsgErr
281        On Error Resume Next
282        Screen.MousePointer = 11
283        crCierroSubReporte JobSRep1
284        Screen.MousePointer = 0
285        Exit Sub
' <VB WATCH>
286        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "rptListaContado"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub rptListaContadoCategoria()
' <VB WATCH>
287        On Error GoTo vbwErrHandler
' </VB WATCH>
288    On Error GoTo ErrCrystal
289    Dim dRige As Date
290    Dim paCategoriaCliente As Long
291    Dim paTipoCuota As Long

292    Dim bHay As Boolean

293        Screen.MousePointer = 11

294        Cons = "Select * from ListasDistribuidores Where LDICodigo = " & cCategoria.ItemData(cCategoria.ListIndex)
295        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
296        If Not RsAux.EOF Then
297            If Not IsNull(RsAux!LDiCatCliente) Then
298                 paCategoriaCliente = RsAux!LDiCatCliente
299            End If
300            If Not IsNull(RsAux!LDiTipoCuota) Then
301                 paTipoCuota = RsAux!LDiTipoCuota
302            End If
303        End If
304        RsAux.Close

           'Saco la máxima fecha de vigencia para cargar valor Rige
305        dRige = Now

306        Cons = "Select Max(Precios.HPrVigencia) from HistoriaPrecio Precios, Articulo, ArticuloFacturacion, CategoriaDescuento " & _
                      " Where (Precios.HPrVigencia IN " & _
                           " (Select MAX(H.HPrVigencia)" & _
                                   " FROM HistoriaPrecio H" & _
                                   " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                                   " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                                   " And H.HPrMoneda = Precios.HPrMoneda " & _
                                   " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                                   " )) " & _
                       " And Precios.HPrArticulo = ArtID" & _
                       " And ArtId = AFaArticulo " & _
                       " And ArtEnUso = 1 And AFaInterior = 1 " & _
                       " And Precios.HPrTipoCuota = " & paTipoCuotaContado & _
                       " And Precios.HPrMoneda = " & paMonedaPesos & _
                       " And Precios.HPrHabilitado = 1" & _
                       " And AFaCategoriaD = CDtCatArticulo" & _
                       " And CDtCatCliente = " & paCategoriaCliente & _
                       " And CDtCatPlazo = " & paTipoCuota
307        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
308        bHay = True
309        If Not RsAux.EOF Then
310            If Not IsNull(RsAux(0)) Then
311                 dRige = RsAux(0)
312            Else
313                 bHay = False
314            End If
315        End If
316        RsAux.Close
           '--------------------------------------------------------------------------------------------------------------------------------------------

317        If Not bHay Then
318            MsgBox "No hay precios vigentes para la categoría seleccionada.", vbExclamation, "No hay Datos"
319            Screen.MousePointer = 0
320            Exit Sub
321        End If

           'Obtengo la cantidad de formulas que tiene el reporte.----------------------
322        CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
323        If CantForm = -1 Then
324             GoTo ErrCrystal
325        End If

           'Cargo Propiedades para el reporte Contado --------------------------------
326        For I = 0 To CantForm - 1
327            NombreFormula = crObtengoNombreFormula(jobnum, I)

328            Select Case LCase(NombreFormula)
                   Case ""
329                GoTo ErrCrystal
330                Case "prmvigencia"
331                Result = crSeteoFormula(jobnum%, NombreFormula, "'Rige desde el: " & Format(dRige, "dd/Mmm/yyyy") & "'")
332                Case "prmlista"
333                Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(cCategoria.Text) & "'")

334                Case Else
335                Result = 1
336            End Select
337            If Result = 0 Then
338                 GoTo ErrCrystal
339            End If
340        Next
           '--------------------------------------------------------------------------------------------------------------------------------------------

           'Seteo la Query del reporte-----------------------------------------------------------------
341        Cons = "Select Articulo.ArtCodigo, Articulo.ArtNombre, Articulo.ArtHabilitado," & _
                               " ArticuloFacturacion.AFaInterior, ArticuloFacturacion.AFaArgumCorto," & _
                               " Precios.HPrPrecio, Especie.EspNombre" & _
                      " From " & _
                               paBD & ".dbo.HistoriaPrecio Precios, " & paBD & ".dbo.Articulo, " & _
                               paBD & ".dbo.Tipo, " & paBD & ".dbo.ArticuloFacturacion, " & paBD & ".dbo.Especie, " & paBD & ".dbo.CategoriaDescuento " & _
                       " WHERE (HPrVigencia IN" & _
                                   " (Select MAX(H.HPrVigencia)" & _
                                   " FROM HistoriaPrecio H" & _
                                   " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                                   " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                                   " And H.HPrMoneda = Precios.HPrMoneda " & _
                                   " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                                   " )) " & _
                       " And Precios.HPrArticulo = ArtID" & _
                       " And ArtId = AFaArticulo " & _
                       " And ArtTipo = TipCodigo " & _
                       " And TipEspecie = EspCodigo" & _
                       " And ArtEnUso = 1 And AFaInterior = 1 " & _
                       " And Precios.HPrTipoCuota = " & paTipoCuotaContado & _
                       " And Precios.HPrMoneda = " & paMonedaPesos & _
                       " And Precios.HPrHabilitado = 1" & _
                       " And AFaCategoriaD = CDtCatArticulo" & _
                       " And CDtCatCliente = " & paCategoriaCliente & _
                       " And CDtCatPlazo = " & paTipoCuota

                       '" Order By Especie.EspNombre Asc, Articulo.ArtNombre Asc"

342        Cons = Trim(Cons) & Chr$(0)

343        If crSeteoSqlQuery(jobnum%, Cons) = 0 Then
344             GoTo ErrCrystal
345        End If
           '-------------------------------------------------------------------------------------------------------------------------------------

346        If crMandoAPantalla(jobnum, "Lista Distribuidores") = 0 Then
347             GoTo ErrCrystal
348        End If
           'If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
349        If Not crInicioImpresion(jobnum, True, False) Then
350             GoTo ErrCrystal
351        End If

       '    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal

       '    crEsperoCierreReportePantalla
352        Screen.MousePointer = 0
353        Exit Sub

354    ErrCrystal:
355        Screen.MousePointer = 0
356        clsGeneral.OcurrioError crMsgErr
357        On Error Resume Next
358        Screen.MousePointer = 11
359        crCierroSubReporte JobSRep1
360        Screen.MousePointer = 0
361        Exit Sub
' <VB WATCH>
362        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "rptListaContadoCategoria"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Function InicializoReporteEImpresora(paNImpresora As String, paBImpresora As Integer, Reporte As String, Optional Orientacion As Integer = 1) As Boolean
' <VB WATCH>
363        On Error GoTo vbwErrHandler
' </VB WATCH>
364    On Error GoTo ErrCrystal

365        jobnum = crAbroReporte(prmPathListados & Reporte)
366        If jobnum = 0 Then
367             GoTo ErrCrystal
368        End If

           'Configuro la Impresora
           'If Trim(Printer.DeviceName) <> Trim(paNImpresora) Then SeteoImpresoraPorDefecto paNImpresora
369        If Not crSeteoImpresora(jobnum, Printer, paBImpresora, Orientacion) Then
370             GoTo ErrCrystal
371        End If
372        InicializoReporteEImpresora = False
373        Exit Function

374    ErrCrystal:
375        InicializoReporteEImpresora = True
376        Screen.MousePointer = 0
377        clsGeneral.OcurrioError crMsgErr
378        On Error Resume Next
379        Screen.MousePointer = 11
380        crCierroTrabajo jobnum
381        Screen.MousePointer = 0

' <VB WATCH>
382        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "InicializoReporteEImpresora"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Private Sub tabLista_Click()
' <VB WATCH>
383        On Error GoTo vbwErrHandler
' </VB WATCH>

384        Select Case tabLista.SelectedItem.Key
               Case "definidas"
385            picLista(1).ZOrder 0
386            Case "varias"
387            picLista(0).ZOrder 0
388            Case "etiquetas"
389            picLista(2).ZOrder 0
390        End Select

' <VB WATCH>
391        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tabLista_Click"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub tArticulo_Change()
' <VB WATCH>
392        On Error GoTo vbwErrHandler
' </VB WATCH>
393        tArticulo.Tag = ""
' <VB WATCH>
394        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tArticulo_Change"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub tArticulo_GotFocus()
' <VB WATCH>
395        On Error GoTo vbwErrHandler
' </VB WATCH>
396        With tArticulo
397            .SelStart = 0
398            .SelLength = Len(.Text)
399        End With
' <VB WATCH>
400        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tArticulo_GotFocus"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
' <VB WATCH>
401        On Error GoTo vbwErrHandler
' </VB WATCH>
402    On Error GoTo ErrAP
403    Dim lEsta As Long, lid As Long
404    Dim sNombre As String

405        If KeyAscii = vbKeyReturn Then
406            If Val(tArticulo.Tag) <> 0 Then
                   'cQueEtiqueta.ListIndex = 2
407                cQueEtiqueta.SetFocus
408                Exit Sub
409            End If
410            Screen.MousePointer = 11
411            If Trim(tArticulo.Text) <> "" Then
412                If IsNumeric(tArticulo.Text) Then
413                    BuscoArticuloPorCodigo tArticulo.Text, sNombre, lid
414                    If lid > 0 Then
415                        tArticulo.Text = sNombre
416                        tArticulo.Tag = lid
417                    ElseIf lid = -1 Then
                           'No tiene precio
418                        tArticulo.Tag = "0"
419                    Else
420                        tArticulo.Tag = "0"
421                        MsgBox "No existe un artículo que posea ese código.", vbExclamation, "ATENCIÓN"
422                    End If
423                Else
424                    lid = BuscoArticuloPorNombre(tArticulo.Text)
425                    BuscoArticuloPorCodigo lid, sNombre, lid
426                    If lid > 0 Then
427                        tArticulo.Text = sNombre
428                        tArticulo.Tag = lid
429                    Else
                           'No tiene precio
430                        tArticulo.Tag = "0"
431                    End If
432                End If
433                If Val(tArticulo.Tag) > 0 Then
434                    For lEsta = 1 To vsEtiquetaArt.Rows - 1
435                        If Val(vsEtiquetaArt.Cell(flexcpData, lEsta, 0)) = Val(tArticulo.Tag) Then
436                            MsgBox "El artículo ya esta ingresado, edite la columna de cantidades si desea modificarlas.", vbInformation, "ATENCIÓN"
437                            vsEtiquetaArt.Select lEsta, 0, lEsta, vsEtiquetaArt.Cols - 1
438                            vsEtiquetaArt.SetFocus
439                            tArticulo.Text = ""
440                            tArticulo.Tag = ""
441                            Exit Sub
442                        End If
443                    Next
                       'cQueEtiqueta.ListIndex = 2
444                    cQueEtiqueta.SetFocus
445                End If
446            End If
447            Screen.MousePointer = 0
448        End If
449        Exit Sub
450    ErrAP:
451        clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
452        Screen.MousePointer = 0
' <VB WATCH>
453        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tArticulo_KeyPress"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub tCantidad_GotFocus()
' <VB WATCH>
454        On Error GoTo vbwErrHandler
' </VB WATCH>
455        With tCantidad
456            .SelStart = 0
457            .SelLength = Len(.Text)
458        End With
459        If Val(tCantidad.Text) = 0 Then
460             tCantidad.Text = vscCantidad.Value
461        End If
' <VB WATCH>
462        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tCantidad_GotFocus"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
' <VB WATCH>
463        On Error GoTo vbwErrHandler
' </VB WATCH>
464        If KeyAscii = vbKeyReturn Then
465            If IsNumeric(tCantidad.Text) Then
466                If Val(tCantidad.Text) < 1 Then
467                     tCantidad.Text = vscCantidad.Value
468                End If
469                vscCantidad.Value = Val(tCantidad.Text)
470                bAgregar_Click
471            Else
472                MsgBox "Formato incorrecto.", vbExclamation, "ATENCIÓN"
473                tCantidad.Text = vscCantidad.Value
474            End If
475        End If
' <VB WATCH>
476        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tCantidad_KeyPress"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub tCantidad_LostFocus()
' <VB WATCH>
477        On Error GoTo vbwErrHandler
' </VB WATCH>
478        If Not IsNumeric(tCantidad.Text) Then
479            tCantidad.Text = vscCantidad.Value
480        Else
481            If Val(tCantidad.Text) < 1 Then
482                 tCantidad.Text = vscCantidad.Value
483            End If
484        End If
' <VB WATCH>
485        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tCantidad_LostFocus"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub tVigencia_Change()
' <VB WATCH>
486        On Error GoTo vbwErrHandler
' </VB WATCH>
487        lVigencia.Caption = ""
' <VB WATCH>
488        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tVigencia_Change"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub tVigencia_GotFocus()
' <VB WATCH>
489        On Error GoTo vbwErrHandler
' </VB WATCH>
490        tVigencia.Appearance = 1
491        tVigencia.BackColor = vbWindowBackground
492        tVigencia.SelStart = 0
493        tVigencia.SelLength = Len(tVigencia.Text)
' <VB WATCH>
494        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tVigencia_GotFocus"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub tVigencia_KeyPress(KeyAscii As Integer)
' <VB WATCH>
495        On Error GoTo vbwErrHandler
' </VB WATCH>
496        If KeyAscii = vbKeyReturn Then
497            If IsDate(tVigencia.Text) Then
498                tVigencia.Text = Format(tVigencia, "dd/mm/yyyy")
499                lVigencia.Caption = Format(tVigencia, "Ddd d/Mmm/yyyy")
500            Else
501                lVigencia.Caption = "#Error"
502            End If
503        End If
' <VB WATCH>
504        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tVigencia_KeyPress"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub tVigencia_LostFocus()
' <VB WATCH>
505        On Error GoTo vbwErrHandler
' </VB WATCH>
506        tVigencia.BackColor = vbButtonFace
' <VB WATCH>
507        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tVigencia_LostFocus"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub



Private Sub ImprimoContadoII()
' <VB WATCH>
508        On Error GoTo vbwErrHandler
' </VB WATCH>
509    On Error GoTo ErrCrystal
510    Dim I As Integer
511    Dim dRige As Date

512        Screen.MousePointer = 11

           'Saco la máxima fecha de vigencia para cargar valor Rige
513        dRige = Now

514        Cons = "Select Max(Precios.HPrVigencia) from HistoriaPrecio Precios, Articulo, ArticuloFacturacion " & _
                      " Where (Precios.HPrVigencia IN " & _
                           " (Select MAX(H.HPrVigencia)" & _
                                   " FROM HistoriaPrecio H" & _
                                   " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                                   " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                                   " And H.HPrMoneda = Precios.HPrMoneda " & _
                                   " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                                   " )) " & _
                       " And Precios.HPrArticulo = ArtID" & _
                       " And ArtId = AFaArticulo " & _
                       " And ArtEnUso = 1 And AFaInterior = 1 " & _
                       " And Precios.HPrMoneda = " & paMonedaPesos & _
                       " And Precios.HPrHabilitado = 1"
515        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
516        If Not RsAux.EOF Then
517             dRige = RsAux(0)
518        End If
519        RsAux.Close

           '--------------------------------------------------------------------------------------------------------------------------------------------

           'Obtengo la cantidad de formulas que tiene el reporte.----------------------
520        CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
521        If CantForm = -1 Then
522             GoTo ErrCrystal
523        End If

           'Cargo Propiedades para el reporte Contado --------------------------------
524        For I = 0 To CantForm - 1
525            NombreFormula = crObtengoNombreFormula(jobnum, I)

526            Select Case LCase(NombreFormula)
                   Case ""
527                GoTo ErrCrystal
528                Case "prmvigencia"
529                Result = crSeteoFormula(jobnum%, NombreFormula, "'Rige desde el: " & Format(dRige, "dd/Mmm/yyyy") & "'")

530                Case Else
531                Result = 1
532            End Select
533            If Result = 0 Then
534                 GoTo ErrCrystal
535            End If
536        Next
           '--------------------------------------------------------------------------------------------------------------------------------------------

           'Seteo la Query del reporte-----------------------------------------------------------------
537        Cons = "Select * " & _
                      " From " & _
                               paBD & ".dbo.HistoriaPrecio Precios, " & _
                               paBD & ".dbo.Articulo, " & paBD & ".dbo.Tipo, " & paBD & ".dbo.ArticuloFacturacion, " & paBD & ".dbo.Especie, " & _
                               paBD & ".dbo.ListasDePrecios, " & paBD & ".dbo.TipoCuota, " & paBD & ".dbo.TipoPlan" & _
                       " WHERE (HPrVigencia IN" & _
                                   " (Select MAX(H.HPrVigencia)" & _
                                   " FROM HistoriaPrecio H" & _
                                   " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                                   " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                                   " And H.HPrMoneda = Precios.HPrMoneda " & _
                                   " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                                   " )) " & _
                       " And Precios.HPrArticulo = ArtID" & _
                       " And ArtId = AFaArticulo " & _
                       " And ArtTipo = TipCodigo " & _
                       " And TipEspecie = EspCodigo" & _
                       " And ArtEnUso = 1 And AFaInterior = 1 " & _
                       " And Precios.HPrMoneda = " & paMonedaPesos & _
                       " And Precios.HPrHabilitado = 1" & _
                       " And ArticuloFacturacion.AFaLista = ListasDePrecios.LDPCodigo " & _
                       " And Precios.HPrTipoCuota = TipoCuota.TCuCodigo" & _
                       " And Precios.HPrPlan = TipoPlan.PlaCodigo"

538        Cons = Cons & _
                        " AND LDPNumero = 1 " & _
                        " And TipoCuota.TCuVencimientoE is null " & _
                        " And TCuEspecial = 0 " & _
                        " and TCuDeshabilitado is null"
                       '" Order By Especie.EspNombre Asc, Articulo.ArtNombre Asc"

539        Cons = Trim(Cons) & Chr$(0)

540        If crSeteoSqlQuery(jobnum%, Cons) = 0 Then
541             GoTo ErrCrystal
542        End If
           '-------------------------------------------------------------------------------------------------------------------------------------

543        If crMandoAPantalla(jobnum, "Lista de Precios") = 0 Then
544             GoTo ErrCrystal
545        End If
           'If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
546        If Not crInicioImpresion(jobnum, True, False) Then
547             GoTo ErrCrystal
548        End If

       '    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal

       '    crEsperoCierreReportePantalla
549        Screen.MousePointer = 0
550        Exit Sub

551    ErrCrystal:
552        Screen.MousePointer = 0
553        clsGeneral.OcurrioError crMsgErr
554        On Error Resume Next
555        Screen.MousePointer = 11
556        crCierroSubReporte JobSRep1
557        Screen.MousePointer = 0
558        Exit Sub
' <VB WATCH>
559        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ImprimoContadoII"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub


Private Sub InicializoGrillas()
' <VB WATCH>
560        On Error GoTo vbwErrHandler
' </VB WATCH>

561        On Error Resume Next
562        With vsLista
563            .Rows = 1
564            .Cols = 1
565            .Editable = False
566            .FormatString = ">Nº|<Listas"
567            .WordWrap = False
568            .ColWidth(0) = 500
569            .ColWidth(1) = 2100
570            .ExtendLastCol = True
571        End With

572        Dim aValor As Long
573        Cons = "Select * from ListasDePrecios order by LDPNumero"
574        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
575        Do While Not RsAux.EOF
576            With vsLista
577                .AddItem ""
578                .Cell(flexcpText, .Rows - 1, 0) = RsAux!LDPNumero
579                aValor = RsAux!LDPCodigo
580                .Cell(flexcpData, .Rows - 1, 0) = aValor

581                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!LDPNombre)
582            End With

583            RsAux.MoveNext
584        Loop
585        RsAux.Close

' <VB WATCH>
586        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "InicializoGrillas"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub vscCantidad_Change()
' <VB WATCH>
587        On Error GoTo vbwErrHandler
' </VB WATCH>
588        tCantidad.Text = vscCantidad.Value
' <VB WATCH>
589        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vscCantidad_Change"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub vsEtiquetaArt_AfterEdit(ByVal Row As Long, ByVal Col As Long)
' <VB WATCH>
590        On Error GoTo vbwErrHandler
' </VB WATCH>
591        AjustoTotalEtiqueta
' <VB WATCH>
592        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vsEtiquetaArt_AfterEdit"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub vsEtiquetaArt_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
' <VB WATCH>
593        On Error GoTo vbwErrHandler
' </VB WATCH>
594    On Error Resume Next

595        If vsEtiquetaArt.IsSubtotal(Row) Or Col = 0 Then
596             Cancel = True
597             Exit Sub
598        End If
599        If Col = 2 And vsEtiquetaArt.Cell(flexcpText, Row, 4) = "" Then
600             Cancel = True
601             Exit Sub
602        End If
603        If Col = 3 And vsEtiquetaArt.Cell(flexcpText, Row, 4) <> "" Then
604             Cancel = True
605             Exit Sub
606        End If

' <VB WATCH>
607        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vsEtiquetaArt_BeforeEdit"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub vsEtiquetaArt_KeyDown(KeyCode As Integer, Shift As Integer)
' <VB WATCH>
608        On Error GoTo vbwErrHandler
' </VB WATCH>
609        If vsEtiquetaArt.Rows = 1 Then
610             Exit Sub
611        End If
612        Select Case KeyCode
               Case vbKeyDelete
613                If Not vsEtiquetaArt.IsSubtotal(vsEtiquetaArt.Row) Then
614                    vsEtiquetaArt.RemoveItem vsEtiquetaArt.Row
615                    AjustoTotalEtiqueta
616                End If
617            Case vbKeyReturn
618            cEtiquetaAImprimir.SetFocus
619        End Select
' <VB WATCH>
620        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vsEtiquetaArt_KeyDown"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub vsEtiquetaArt_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
' <VB WATCH>
621        On Error GoTo vbwErrHandler
' </VB WATCH>

622        If vsEtiquetaArt.EditText = "" Then
623            vsEtiquetaArt.EditText = "0"
624        Else
625            If Not IsNumeric(vsEtiquetaArt.EditText) Then
626                Cancel = True
627                MsgBox "Formato inválido.", vbExclamation, "ATENCIÓN"
628            Else
629                If Val(vsEtiquetaArt.EditText) < 0 Then
630                    Cancel = True
631                    MsgBox "La cantidad tiene que ser mayor o igual a cero.", vbExclamation, "ATENCIÓN"
632                End If
633            End If
634        End If
' <VB WATCH>
635        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vsEtiquetaArt_ValidateEdit"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub vsLista_DblClick()
' <VB WATCH>
636        On Error GoTo vbwErrHandler
' </VB WATCH>
637        If vsLista.Rows = 1 Then
638             Exit Sub
639        End If
640        AccionImprimir TReporte.AlPublico
' <VB WATCH>
641        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vsLista_DblClick"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub rptListaAlPublico()
' <VB WATCH>
642        On Error GoTo vbwErrHandler
' </VB WATCH>
643    On Error GoTo errSel

644        If vsLista.Rows = 1 Then
645             Exit Sub
646        End If
647        Dim miF As New frmPreview

648        With miF
649            .prmHeaderReport = "Lista Nº " & vsLista.Cell(flexcpText, vsLista.Row, 0) & ": " & vsLista.Cell(flexcpText, vsLista.Row, 1)
650            .prmCaption = "Listas de Precios Público"
651            .prmIDLista = vsLista.Cell(flexcpData, vsLista.Row, 0)
652            .prmMonedaPesos = paMonedaPesos
653            .prmVigencia = prmVigencia
654            .Show
655        End With

656        Set miF = Nothing
657        Exit Sub

658    errSel:
659        clsGeneral.OcurrioError "Error al activar la lista de precios.", Err.Description
660        Screen.MousePointer = 0
' <VB WATCH>
661        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "rptListaAlPublico"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub BuscoArticuloPorCodigo(ByVal lCodArticulo As Long, ByRef sNombre As String, ByRef lIDArt As Long)
' <VB WATCH>
662        On Error GoTo vbwErrHandler
' </VB WATCH>
663    On Error GoTo errBA

664        Screen.MousePointer = 11
665        sNombre = ""
666        lIDArt = 0
667        Cons = "Select * From Articulo Where ArtCodigo = " & lCodArticulo
668        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
669        If RsAux.EOF Then
670            RsAux.Close
671        Else
672            sNombre = Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
673            lIDArt = RsAux!ArtID
674            RsAux.Close

675            If Not ArticuloTienePrecio(lIDArt) Then
676                MsgBox "El artículo no posee precios.", vbInformation, "ATENCIÓN"
677                lIDArt = -1
678                sNombre = ""
679            End If
680        End If
681        Screen.MousePointer = 0
682        Exit Sub
683    errBA:
684        clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo por código.", Err.Description, "Error (buscoarticuloporcodigo)"
685        Screen.MousePointer = 0
' <VB WATCH>
686        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "BuscoArticuloPorCodigo"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Function BuscoArticuloPorNombre(NomArticulo As String) As Long
' <VB WATCH>
687        On Error GoTo vbwErrHandler
' </VB WATCH>
688    On Error GoTo errBA
689    Dim lResultado As Long
690        Screen.MousePointer = 11
691        BuscoArticuloPorNombre = 0
692        Cons = "Select ArtCodigo, ArtCodigo as 'Código', ArtNombre as 'Nombre' From Articulo" _
               & " Where ArtNombre LIKE '" & Replace(NomArticulo, " ", "%") & "%'" _
               & " Order By ArtNombre"

693        Dim objAyuda As New clsListadeAyuda
694        If objAyuda.ActivarAyuda(cBase, Cons, 5000, 1, "Lista de Artículos") Then
695            lResultado = objAyuda.RetornoDatoSeleccionado(1)
696        Else
697            lResultado = 0
698        End If
699        Set objAyuda = Nothing       'Destruyo la clase.
700        Screen.MousePointer = 11
701        If lResultado > 0 Then
702             BuscoArticuloPorNombre = lResultado
703        End If
704        Screen.MousePointer = 0
705        Exit Function
706    errBA:
707        clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo por código.", Err.Description, "Error (buscoarticuloporcodigo)"
708        Screen.MousePointer = 0
' <VB WATCH>
709        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "BuscoArticuloPorNombre"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Private Sub InicializoObjetosEtiqueta()
' <VB WATCH>
710        On Error GoTo vbwErrHandler
' </VB WATCH>
711        LimpioIngresoArticuloEtiqueta
712        With vsEtiquetaArt
713            .Rows = 1
714            .Cols = 1
715            .FormatString = "<Artículo|>Q Normal|>Q c/Argum.|>Q s/Argum.|Arg.Largo|Medida"
716            .ColWidth(0) = 3000
717            .ColHidden(4) = True
718            .ColHidden(5) = True
719            .Editable = True
720        End With
           'Cargo combos
721        With cQueEtiqueta
722            .Clear
723            .AddItem "Ambas"
724            .AddItem "Normal (chica)"
725            .AddItem "Según tabla"
726            .AddItem "Vidriera (grande)"
727            .ListIndex = 2
728        End With
729        With cEtiquetaAImprimir
730            .Clear
731            .AddItem "Normales"
732            .AddItem "Vidriera c/Argumento"
733            .AddItem "Vidriera s/Argumento"
734        End With
' <VB WATCH>
735        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "InicializoObjetosEtiqueta"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub LimpioIngresoArticuloEtiqueta()
' <VB WATCH>
736        On Error GoTo vbwErrHandler
' </VB WATCH>

737        tArticulo.Text = ""
738        vscCantidad.Value = 1
739        tCantidad.Text = vscCantidad.Value

' <VB WATCH>
740        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "LimpioIngresoArticuloEtiqueta"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub etiqueta_AgregoArticuloALista(ByVal lIDArticulo As Long, ByVal sNombre As String, ByVal iCantidad As Integer, ByVal iQueEtiqueta As Integer)
' <VB WATCH>
741        On Error GoTo vbwErrHandler
' </VB WATCH>
742    On Error GoTo errAA
743    Dim sArgumLargo As String, sMedida As String
744    Dim lQEN As Long, lQEV As Long

           'Tengo que buscar en la tabla artículo facturación si posee argumento largo.
745        GetArticuloArgumLargoMedidas lIDArticulo, sArgumLargo, sMedida, lQEN, lQEV

746        If iQueEtiqueta <> 2 Then
747            lQEN = iCantidad
748            lQEV = iCantidad
749            If lQEN = 0 And lQEV = 0 Then
750                 Exit Sub
751            End If
752        End If



753        With vsEtiquetaArt
754            .AddItem sNombre
755            .Cell(flexcpData, .Rows - 1, 0) = lIDArticulo

756            Select Case iQueEtiqueta
                   Case 0, 2
757                    .Cell(flexcpText, .Rows - 1, 1) = lQEN
758                    If Trim(sArgumLargo) <> "" Then
759                        .Cell(flexcpText, .Rows - 1, 2) = lQEV
760                        .Cell(flexcpText, .Rows - 1, 3) = "0"
761                    Else
762                        .Cell(flexcpText, .Rows - 1, 2) = "0"
763                        .Cell(flexcpText, .Rows - 1, 3) = lQEV
764                    End If

765                Case 1
766                    .Cell(flexcpText, .Rows - 1, 1) = lQEN
767                    .Cell(flexcpText, .Rows - 1, 2) = "0"
768                    .Cell(flexcpText, .Rows - 1, 3) = "0"

769                Case 3
770                    .Cell(flexcpText, .Rows - 1, 1) = 0
771                    If Trim(sArgumLargo) <> "" Then
772                        .Cell(flexcpText, .Rows - 1, 2) = lQEV
773                        .Cell(flexcpText, .Rows - 1, 3) = "0"
774                    Else
775                        .Cell(flexcpText, .Rows - 1, 2) = "0"
776                        .Cell(flexcpText, .Rows - 1, 3) = lQEV
777                    End If
778            End Select
779            .Cell(flexcpText, .Rows - 1, 4) = Trim(sArgumLargo)
780            .Cell(flexcpText, .Rows - 1, 5) = sMedida
781        End With
782        AjustoTotalEtiqueta
783        Exit Sub
784    errAA:
785        clsGeneral.OcurrioError "Ocurrió un error al intentar agregar el artículo a la lista.", Err.Description, "Error (agregoarticuloalista)"
' <VB WATCH>
786        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "etiqueta_AgregoArticuloALista"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub GetArticuloArgumLargoMedidas(ByVal lIDArticulo As Long, ByRef sArgum As String, _
                                ByRef sMedida As String, ByRef lQEN As Long, ByRef lQEV As Long)
' <VB WATCH>
787        On Error GoTo vbwErrHandler
' </VB WATCH>

788    On Error GoTo errGA
789    Dim cAlto As Currency, cFrente As Currency, cProf As Currency

790        sArgum = ""
791        sMedida = ""
792        lQEN = 0
793        lQEV = 0

794        Cons = "Select * From ArticuloFacturacion Where AFaArticulo = " & lIDArticulo
795        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
796        If Not RsAux.EOF Then

797            If Not IsNull(RsAux!AFAEtNormales) Then
798                 lQEN = RsAux!AFAEtNormales
799            End If
800            If Not IsNull(RsAux!AFAEtVidriera) Then
801                 lQEV = RsAux!AFAEtVidriera
802            End If

803            If Not IsNull(RsAux!AFaArgumLargo) Then
804                sArgum = Trim(RsAux!AFaArgumLargo)
805                If Not sArgum Like "*[0-z]*" Then
806                     sArgum = ""
807                End If
808            End If

               '-----------------------------------------------------------------------
809            cAlto = 0
810            cFrente = 0
811            cProf = 0
812            If Not IsNull(RsAux!AFaAlto) Then
813                 cAlto = Trim(RsAux!AFaAlto)
814            End If
815            If Not IsNull(RsAux!AFaFrente) Then
816                 cFrente = Trim(RsAux!AFaFrente)
817            End If
818            If Not IsNull(RsAux!AFaProfundidad) Then
819                 cProf = RsAux!AFaProfundidad
820            End If
               '-----------------------------------------------------------------------

821            If cAlto <> 0 Then
822                 sMedida = "(" & cAlto
823            End If
824            If cFrente <> 0 Then
825                If sMedida <> "" Then
826                    sMedida = sMedida & "x" & cFrente
827                Else
828                    sMedida = "(" & cAlto
829                End If
830            End If
831            If cProf <> 0 Then
832                If sMedida <> "" Then
833                    sMedida = sMedida & "x" & cProf
834                Else
835                    sMedida = "(" & cProf
836                End If
837            End If
838            If cAlto <> 0 And cFrente <> 0 And cProf <> 0 Then
839                sMedida = sMedida & ")"
840            Else
841                If sMedida <> "" Then
842                     sMedida = sMedida & " cm)"
843                End If
844            End If
845        End If
846        RsAux.Close
847        Exit Sub
848    errGA:
849        clsGeneral.OcurrioError "Ocurrió un error al intentar obtener el argumento largo del artículo.", Err.Description, "Error (getargumento)"
' <VB WATCH>
850        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetArticuloArgumLargoMedidas"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub AjustoTotalEtiqueta()
' <VB WATCH>
851        On Error GoTo vbwErrHandler
' </VB WATCH>
852    On Error Resume Next
853        With vsEtiquetaArt
854            .Subtotal flexSTClear
855            .SubtotalPosition = flexSTBelow
856            .Subtotal flexSTSum, -1, 1, "#,###", &HC0FFFF, vbRed, True, "Total", , True
857            .Subtotal flexSTSum, -1, 2, "#,###"
858            .Subtotal flexSTSum, -1, 3, "#,###"
859        End With
' <VB WATCH>
860        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "AjustoTotalEtiqueta"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub etiqueta_MandoAImprimir()
' <VB WATCH>
861        On Error GoTo vbwErrHandler
' </VB WATCH>

862        If Not IsDate(tVigencia.Text) Then
863            MsgBox "La fecha de vigencia no es correcta.", vbExclamation, "Datos Incorrectos"
864            tVigencia.SetFocus
865            Exit Sub
866        End If

867        If etiqueta_BorroTablaAuxiliar Then
868            If etiqueta_InsertoTablaAuxiliar Then
869                Select Case cEtiquetaAImprimir.ListIndex
                       Case 0
870                    AccionImprimir TReporte.EtiquetaNormal
871                    Case 1
872                    AccionImprimir TReporte.EtiquetaConArgumento
873                    Case 2
874                    AccionImprimir TReporte.EtiquetaSinArgumento
875                End Select
876                etiqueta_BorroTablaAuxiliar
877            End If
878       End If

' <VB WATCH>
879        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "etiqueta_MandoAImprimir"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Function etiqueta_BorroTablaAuxiliar() As Boolean
' <VB WATCH>
880        On Error GoTo vbwErrHandler
' </VB WATCH>
881    On Error GoTo errBTA
882        etiqueta_BorroTablaAuxiliar = False
883        Cons = "Delete EtiquetaAImprimir"
884        cBase.Execute (Cons)
885        etiqueta_BorroTablaAuxiliar = True
886        Exit Function
887    errBTA:
888        clsGeneral.OcurrioError "Ocurrió un error al vaciar la tabla auxiliar de impresión.", Err.Description, "Error (borrotablaauxiliar)"
' <VB WATCH>
889        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "etiqueta_BorroTablaAuxiliar"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Private Function etiqueta_InsertoTablaAuxiliar() As Boolean
' <VB WATCH>
890        On Error GoTo vbwErrHandler
' </VB WATCH>
891    On Error GoTo errITA
892    Dim lCont As Long
893    Dim sCtdo As String, sCuota As String, sPlan As String, sTotFin As String

894        etiqueta_InsertoTablaAuxiliar = False
895        Select Case cEtiquetaAImprimir.ListIndex
               Case 0  'Normales
896                For lCont = 1 To vsEtiquetaArt.Rows - 1
897                    If Val(vsEtiquetaArt.Cell(flexcpText, lCont, 1)) > 0 And vsEtiquetaArt.IsSubtotal(lCont) = False Then
                           'Armo los precios para este artículo.
898                        etiqueta_CargoPrecios Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, ""
                           'Inserto la cantidad de copias que pidio.
899                        etiqueta_AddRowTablaAux Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, Trim(vsEtiquetaArt.Cell(flexcpText, lCont, 5)), sPlan, Val(vsEtiquetaArt.Cell(flexcpText, lCont, 1))
900                    End If
901                Next lCont

902            Case 1 'c/ y s/argumento
903                For lCont = 1 To vsEtiquetaArt.Rows - 1
904                    If Val(vsEtiquetaArt.Cell(flexcpText, lCont, 2)) > 0 And vsEtiquetaArt.IsSubtotal(lCont) = False Then
                           'Armo los precios para este artículo.
905                        etiqueta_CargoPrecios Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin
                           'Inserto la cantidad de copias que pidio.
906                        etiqueta_AddRowTablaAux Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, Trim(sTotFin), sPlan, Val(vsEtiquetaArt.Cell(flexcpText, lCont, 2))
907                    End If
908                Next lCont
909            Case 2
910                For lCont = 1 To vsEtiquetaArt.Rows - 1
911                    If Val(vsEtiquetaArt.Cell(flexcpText, lCont, 3)) > 0 And vsEtiquetaArt.IsSubtotal(lCont) = False Then
                           'Armo los precios para este artículo.
912                        etiqueta_CargoPrecios Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin
                           'Inserto la cantidad de copias que pidio.
913                        etiqueta_AddRowTablaAux Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, Trim(sTotFin), sPlan, Val(vsEtiquetaArt.Cell(flexcpText, lCont, 3))
914                    End If
915                Next lCont
916        End Select
917        etiqueta_InsertoTablaAuxiliar = True
918        Exit Function

919    errITA:
920        clsGeneral.OcurrioError "Ocurrió un error al insertar los artículos en la tabla auxiliar.", Err.Description, "Error (insertotablaauxiliar)"
921        etiqueta_BorroTablaAuxiliar
' <VB WATCH>
922        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "etiqueta_InsertoTablaAuxiliar"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Public Function etiqueta_AddRowTablaAux(ByVal lArt As Long, ByVal sCtdo As String, ByVal sCuota As String, ByVal sFinanMedida As String, ByVal sPlan As String, ByVal iCant As Integer)
' <VB WATCH>
923        On Error GoTo vbwErrHandler
' </VB WATCH>
924    Dim rsAdd As rdoResultset
925    Dim iCont As Integer
926        Cons = "Select * From EtiquetaAImprimir Where EImArticulo = " & lArt
927        Set rsAdd = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
928        For iCont = 1 To iCant
929            rsAdd.AddNew
930            rsAdd!EImArticulo = lArt
931            rsAdd!EImImporteCtdo = sCtdo
932            rsAdd!EImCuota = sCuota
933            rsAdd!EImTotalFinanciado = sFinanMedida
934            rsAdd!EImPlan = sPlan
935            rsAdd.Update
936        Next
937        rsAdd.Close
' <VB WATCH>
938        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "etiqueta_AddRowTablaAux"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Private Sub etiqueta_CargoPrecios(ByVal lArt As Long, ByRef sCtdo As String, ByRef sCuota As String, ByRef sPlan As String, ByRef sTotFin As String)
' <VB WATCH>
939        On Error GoTo vbwErrHandler
' </VB WATCH>
940    Dim cCuotaAnt As Currency
941    Dim iCountEnter As Integer
942        sCtdo = ""
943        sPlan = ""
944        sCuota = ""
945        sTotFin = ""

946        Cons = "Select * " & _
                      " From " & _
                               "HistoriaPrecio Precios, TipoCuota, TipoPlan" & _
                       " WHERE (HPrVigencia IN" & _
                                   " (Select MAX(H.HPrVigencia)" & _
                                   " FROM HistoriaPrecio H" & _
                                   " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                                   " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                                   " And H.HPrMoneda = Precios.HPrMoneda " & _
                                   " And H.HPrVigencia <= '" & Format(tVigencia.Text, "mm/dd/yyyy 23:59:59") & "'" & _
                                   " )) " & _
                       " And Precios.HPrArticulo = " & lArt & _
                       " And Precios.HPrMoneda = " & paMonedaPesos & _
                       " And Precios.HPrHabilitado = 1" & _
                       " And Precios.HPrTipoCuota = TipoCuota.TCuCodigo" & _
                       " And Precios.HPrPlan = TipoPlan.PlaCodigo" & _
                       " And TipoCuota.TCuVencimientoE is null " & _
                       " And TCuEspecial = 0 And TCuVencimientoC = 0 " & _
                       " And TCuDeshabilitado is Null" & _
                       " Order By TCuCantidad Asc"

947        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
948        cCuotaAnt = -1
949        Do While Not RsAux.EOF
950            If Not IsNull(RsAux!PlaNombre) Then
951                 sPlan = "(" & Trim(RsAux!PlaNombre) & ")"
952            End If
953            If paTipoCuotaContado = RsAux!TCuCodigo Then
954                Select Case cEtiquetaAImprimir.ListIndex
                       Case 0
955                    sCtdo = "Cdo: $ " & Format(RsAux!HPrPrecio, "#,###")
956                    Case 1, 2
957                    sCtdo = "Contado: $ " & Format(RsAux!HPrPrecio, "#,###")
958                End Select
959            Else
960                Select Case cEtiquetaAImprimir.ListIndex
                       Case 0
961                        If sCuota <> "" Then
962                             sCuota = sCuota & vbCrLf
963                        End If
964                        sCuota = sCuota & Trim(RsAux!TCuCantidad) & " x $ " & Format(RsAux!HPrPrecio / RsAux!TCuCantidad, "#,###")
965                    Case 1, 2
966                        If (cCuotaAnt > RsAux!HPrPrecio / RsAux!TCuCantidad _
                               Or cCuotaAnt = -1) And RsAux!HPrPrecio / RsAux!TCuCantidad > paCuotaMin Then

967                            cCuotaAnt = RsAux!HPrPrecio / RsAux!TCuCantidad
968                            sCuota = "...o en " & Trim(RsAux!TCuCantidad) & " Cuotas de $ " & Format(RsAux!HPrPrecio / RsAux!TCuCantidad, "#,###")
969                            sTotFin = "Total financiado = $ " & Format(RsAux!HPrPrecio, "#,###")
970                        End If
971                End Select
972            End If
973            RsAux.MoveNext
974        Loop
975        RsAux.Close

976        If cEtiquetaAImprimir.ListIndex = 0 Then
977            Dim arrEnter() As String, I As Integer
978            arrEnter = Split(sCuota, vbCrLf)
979            I = UBound(arrEnter)
980            For I = I To 4
981                sCuota = sCuota & vbCrLf & " "
982            Next
983            sCuota = Trim(sCuota) & "...................."
984        End If

' <VB WATCH>
985        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "etiqueta_CargoPrecios"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub rptImprimoEtiquetas(ByVal sTitulo As String)
' <VB WATCH>
986        On Error GoTo vbwErrHandler
' </VB WATCH>
987    On Error GoTo ErrCrystal
988        Screen.MousePointer = 11

989        If crMandoAPantalla(jobnum, sTitulo) = 0 Then
990             GoTo ErrCrystal
991        End If
992        If Not crInicioImpresion(jobnum, True, False) Then
993             GoTo ErrCrystal
994        End If

995        Screen.MousePointer = 0
996        Exit Sub

997    ErrCrystal:
998        clsGeneral.OcurrioError crMsgErr
999        Screen.MousePointer = 0
' <VB WATCH>
1000       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "rptImprimoEtiquetas"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Function ArticuloTienePrecio(ByVal lArt As Long) As Boolean
' <VB WATCH>
1001       On Error GoTo vbwErrHandler
' </VB WATCH>

1002       Cons = "Select * " & _
                   " From " & _
                           "HistoriaPrecio Precios, TipoCuota, TipoPlan" & _
                   " WHERE (HPrVigencia IN" & _
                               " (Select MAX(H.HPrVigencia)" & _
                               " FROM HistoriaPrecio H" & _
                               " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                               " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                               " And H.HPrMoneda = Precios.HPrMoneda " & _
                               " And H.HPrVigencia <= '" & Format(tVigencia.Text, "mm/dd/yyyy 23:59:59") & "'" & _
                               " )) " & _
                   " And Precios.HPrArticulo = " & lArt & _
                   " And Precios.HPrMoneda = " & paMonedaPesos & _
                   " And Precios.HPrHabilitado = 1" & _
                   " And Precios.HPrTipoCuota = TipoCuota.TCuCodigo" & _
                   " And Precios.HPrPlan = TipoPlan.PlaCodigo" & _
                   " And TipoCuota.TCuVencimientoE Is Null " & _
                   " And TCuEspecial = 0 And TCuVencimientoC = 0 " & _
                   " And TCuDeshabilitado Is Null"

1003       Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
1004       If Not RsAux.EOF Then
1005           ArticuloTienePrecio = True
1006       Else
1007           ArticuloTienePrecio = False
1008       End If
1009       RsAux.Close

' <VB WATCH>
1010       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ArticuloTienePrecio"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

