VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmPreview 
   Caption         =   "Preview"
   ClientHeight    =   3450
   ClientLeft      =   1230
   ClientTop       =   2010
   ClientWidth     =   7380
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   7380
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picCopia 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   5400
      ScaleHeight     =   315
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
      Begin VB.VScrollBar vsCopias 
         Height          =   285
         Left            =   960
         Max             =   -1
         Min             =   -1111
         TabIndex        =   6
         Top             =   0
         Value           =   -1
         Width           =   255
      End
      Begin VB.TextBox tCopias 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         MaxLength       =   5
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Copias:"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   40
         Width           =   615
      End
   End
   Begin VB.HScrollBar fsbZoom 
      Height          =   255
      Left            =   6480
      TabIndex        =   3
      Top             =   1620
      Width           =   1095
   End
   Begin VB.TextBox tPage 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6480
      MaxLength       =   5
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin MSComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Cerrar [Ctrl+X]"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "play"
            Object.ToolTipText     =   "Refrescar consulta. [Ctrl+E]"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stop"
            Object.ToolTipText     =   "Cancelar carga. [Ctrl+C]"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "separator1"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir. [Ctrl+P]"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "printconfig"
            Object.ToolTipText     =   "Configurar página."
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "printcopies"
            Style           =   4
            Object.Width           =   1330
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "firstpage"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previouspage"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pagenumber"
            Style           =   4
            Object.Width           =   815
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nextpage"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lastpage"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "separator4"
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "zoom"
            Style           =   4
            Object.Width           =   1500
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   4920
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":0442
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":075C
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":08B6
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":0D08
            Key             =   "print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":0E1A
            Key             =   "printcfg"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":0F2C
            Key             =   "first"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":137E
            Key             =   "previous"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":17D0
            Key             =   "next"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":1C22
            Key             =   "last"
         EndProperty
      EndProperty
   End
   Begin vsViewLib.vsPrinter vspReporte 
      Height          =   2055
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   4215
      _Version        =   196608
      _ExtentX        =   7435
      _ExtentY        =   3625
      _StockProps     =   229
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
      PhysicalPage    =   -1  'True
      Zoom            =   80
      ZoomMax         =   160
      ZoomStep        =   10
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Properties.
Private m_Lista As Long
Private m_Header As String
Private m_Vigencia As String
Private m_MonedaPesos As Long
Private m_Caption As String
'------------------------------------------------

Private Type tRepProperties
    Device As String
    MarginL As Long
    MarginR As Long
    MarginB As Long
    MarginT As Long
    Orientation As Integer
    PaperSize As Integer
End Type

Private bHeader As Boolean

Private dTitRige As Date
Private sTitTableTC As String
Private sTitTableTCD As String      'Diferidas.
Private lHeightCol As Long

'Guardo dimensión de tabla.
Private sDimTableTC As String, sDimTableTCD As String

Private Const sDimTableFecha = "|>770"
Private Const sDimTableArt = "<300|>780|<3150"
Private Const lDimTable = 5000
'------------------------------------------------

Private arrCuota() As Long
Private arrDiferidos() As Long
Private arrPrecios() As String, arrDif() As String

Private bLoad As Boolean
Private bCancelQuery As Boolean

' <VB WATCH>
Const VBWMODULE = "frmPreview"
' </VB WATCH>

Public Property Let prmVigencia(ByVal dRige As String)
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>
2          m_Vigencia = dRige
' <VB WATCH>
3          Exit Property
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "prmVigencia[Let]"

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
End Property

Public Property Let prmMonedaPesos(ByVal lMon As Long)
' <VB WATCH>
4          On Error GoTo vbwErrHandler
' </VB WATCH>
5          m_MonedaPesos = lMon
' <VB WATCH>
6          Exit Property
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "prmMonedaPesos[Let]"

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
End Property

Public Property Let prmIDLista(ByVal lid As Long)
' <VB WATCH>
7          On Error GoTo vbwErrHandler
' </VB WATCH>
8          m_Lista = lid
' <VB WATCH>
9          Exit Property
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "prmIDLista[Let]"

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
End Property

Public Property Let prmCaption(ByVal sCaption As String)
' <VB WATCH>
10         On Error GoTo vbwErrHandler
' </VB WATCH>
11         m_Caption = sCaption
' <VB WATCH>
12         Exit Property
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "prmCaption[Let]"

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
End Property

Public Property Let prmHeaderReport(ByVal sHeader As String)
' <VB WATCH>
13         On Error GoTo vbwErrHandler
' </VB WATCH>
14         m_Header = sHeader
' <VB WATCH>
15         Exit Property
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "prmHeaderReport[Let]"

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
End Property

Private Sub Form_Load()
' <VB WATCH>
16         On Error GoTo vbwErrHandler
' </VB WATCH>
17     On Error GoTo errLoad

18         bLoad = True
19         bHeader = False
20         If m_Vigencia = "0:00:00" Then
21              prmVigencia = Format(Now, "mm/dd/yyyy hh:nn:ss")
22         End If
23         With tooMenu
24             .ImageList = imgIcon
25             .Buttons("salir").Image = "salir"
26             .Buttons("play").Image = "refresh"
27             .Buttons("stop").Image = "stop"
28             .Buttons("print").Image = "print"
29             .Buttons("printconfig").Image = "printcfg"
30             .Buttons("firstpage").Image = "first"
31             .Buttons("previouspage").Image = "previous"
32             .Buttons("nextpage").Image = "next"
33             .Buttons("lastpage").Image = "last"
34         End With

35         With vspReporte
36             .Zoom = 100
37             fsbZoom.LargeChange = .ZoomStep
38             fsbZoom.SmallChange = .ZoomStep / 2
39             fsbZoom.Min = .ZoomMin
40             fsbZoom.Max = .ZoomMax
41             fsbZoom.Value = .Zoom

42             .PaperSize = 1
43             .Orientation = orPortrait
               'OJO EN WIN2000
               'si ponemos printer se cuelga el equipo al hacer render control
44             .PreviewMode = pmScreen
45             .MarginLeft = 576 '649
46             .MarginRight = 576 '609
47             .MarginTop = 1200
48             .PhysicalPage = True
49             .PageBorder = pbBottom

50         End With

51         StartReport

52         If m_Caption <> "" Then
53              Me.Caption = m_Caption
54         End If

55         With vspReporte
56             .Zoom = 100
57             fsbZoom.LargeChange = .ZoomStep
58             fsbZoom.SmallChange = .ZoomStep / 2
59             fsbZoom.Min = .ZoomMin
60             fsbZoom.Max = .ZoomMax
61             fsbZoom.Value = .Zoom
62         End With
63         Exit Sub

64     errLoad:
' <VB WATCH>
65         Exit Sub
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
66         On Error GoTo vbwErrHandler
' </VB WATCH>
67     On Error Resume Next

68         If Me.WindowState = vbMinimized Then
69              Exit Sub
70         End If
71         With vspReporte
72             .Top = tooMenu.Top + tooMenu.Height
73             .Left = ScaleLeft
74             .Width = ScaleWidth
75             .Height = ScaleHeight - .Top
76         End With
77         With fsbZoom
78             .Move tooMenu.Buttons("zoom").Left, tooMenu.Buttons("zoom").Top + ((tooMenu.Height - .Height) / 1.5), tooMenu.Buttons("zoom").Width
79         End With
80         With picCopia
81             .Move tooMenu.Buttons("printcopies").Left + 100, tooMenu.Buttons("printcopies").Top + 50  '((tooMenu.Height - .Height) / 1.5)   ', tooMenu.Buttons("printcopies").Width
82         End With

83         With tPage
84             .Move tooMenu.Buttons("pagenumber").Left + 150, tooMenu.Buttons("pagenumber").Top + ((tooMenu.Height - .Height) / 1.5)  ', tooMenu.Buttons("pagenumber").Width
85         End With
' <VB WATCH>
86         Exit Sub
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

Private Sub fsbZoom_Change()
' <VB WATCH>
87         On Error GoTo vbwErrHandler
' </VB WATCH>
88     On Error Resume Next
89         If vspReporte Is Nothing Then
90              Exit Sub
91         End If
92         vspReporte.Zoom = fsbZoom.Value
' <VB WATCH>
93         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fsbZoom_Change"

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

Private Sub tCopias_KeyPress(KeyAscii As Integer)
' <VB WATCH>
94         On Error GoTo vbwErrHandler
' </VB WATCH>
95     On Error Resume Next
96         If KeyAscii = vbKeyReturn Then
97             If IsNumeric(tCopias.Text) Then
98                 vsCopias.Value = Val(tCopias.Text) * -1
99             Else
100                MsgBox "Formato incorrecto.", vbExclamation, "ATENCIÓN"
101                tCopias.Text = vsCopias.Value * -1
102            End If
103        End If
' <VB WATCH>
104        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tCopias_KeyPress"

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

Private Sub tCopias_LostFocus()
' <VB WATCH>
105        On Error GoTo vbwErrHandler
' </VB WATCH>
106    On Error Resume Next
107        If IsNumeric(tCopias.Text) Then
108            vsCopias.Value = Val(tCopias.Text) * -1
109        Else
110            tCopias.Text = vsCopias.Value * -1
111        End If
' <VB WATCH>
112        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tCopias_LostFocus"

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

Private Sub tooMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
' <VB WATCH>
113        On Error GoTo vbwErrHandler
' </VB WATCH>
114        Select Case Button.Key
               Case "salir"
115            Unload Me
116            Case "play"
117            StartReport
118            Case "stop"
119            ActionStop
120            Case "print"
121            ActionPrint
122            Case "printconfig"
123            ActionConfigPage

               'Botones de reporte.
124            Case "firstpage"
125                vspReporte.PreviewPage = 1
126                SetButtonReport
127            Case "previouspage"
128                vspReporte.PreviewPage = vspReporte.PreviewPage - 1
129                SetButtonReport
130            Case "nextpage"
131                vspReporte.PreviewPage = vspReporte.PreviewPage + 1
132                SetButtonReport
133            Case "lastpage"
134                vspReporte.PreviewPage = vspReporte.PageCount
135                SetButtonReport
136        End Select
' <VB WATCH>
137        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tooMenu_ButtonClick"

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

Private Sub ActionStop()
' <VB WATCH>
138        On Error GoTo vbwErrHandler
' </VB WATCH>
139        bCancelQuery = True
' <VB WATCH>
140        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ActionStop"

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

Private Sub ActionConfigPage()
' <VB WATCH>
141        On Error GoTo vbwErrHandler
' </VB WATCH>
142    On Error GoTo errCancel
143    Dim vProperties As tRepProperties

144        With vProperties
145            .Device = vspReporte.Device
146            .MarginB = vspReporte.MarginBottom
147            .MarginL = vspReporte.MarginLeft
148            .MarginR = vspReporte.MarginRight
149            .MarginT = vspReporte.MarginTop
150            .Orientation = vspReporte.Orientation
151            .PaperSize = vspReporte.PaperSize
152        End With

153        If vspReporte.PrintDialog(pdPageSetup) Then

154            If vProperties.Device <> vspReporte.Device Or _
                   vProperties.MarginB <> vspReporte.MarginBottom Or _
                   vProperties.MarginL <> vspReporte.MarginLeft Or _
                   vProperties.MarginR <> vspReporte.MarginRight Or _
                   vProperties.MarginT <> vspReporte.MarginTop Or _
                   vProperties.Orientation <> vspReporte.Orientation Or _
                   vProperties.PaperSize <> vspReporte.PaperSize Then

155                StartReport

156            End If
157        End If
158        Exit Sub
159    errCancel:
160        Screen.MousePointer = 0
161        Exit Sub
' <VB WATCH>
162        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ActionConfigPage"

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
Private Sub ActionPrint()
' <VB WATCH>
163        On Error GoTo vbwErrHandler
' </VB WATCH>
164    Dim vProperties As tRepProperties
165    Dim lCopies As Long

166        With vProperties
167            .Device = vspReporte.Device
168            .MarginB = vspReporte.MarginBottom
169            .MarginL = vspReporte.MarginLeft
170            .MarginR = vspReporte.MarginRight
171            .MarginT = vspReporte.MarginTop
172            .Orientation = vspReporte.Orientation
173            .PaperSize = vspReporte.PaperSize
174        End With

175        If Val(tCopias.Text) < 1 Then
176             tCopias.Text = 1
177        End If
178        lCopies = Val(tCopias.Text)

179        If vspReporte.PrintDialog(pdPrinterSetup) Then

180            If vProperties.Device <> vspReporte.Device Or _
                   vProperties.MarginB <> vspReporte.MarginBottom Or _
                   vProperties.MarginL <> vspReporte.MarginLeft Or _
                   vProperties.MarginR <> vspReporte.MarginRight Or _
                   vProperties.MarginT <> vspReporte.MarginTop Or _
                   vProperties.Orientation <> vspReporte.Orientation Or _
                   vProperties.PaperSize <> vspReporte.PaperSize Then

181                StartReport

182            End If
183            vspReporte.AbortWindow = False
184            vspReporte.FileName = Me.Caption

185            tCopias.Text = lCopies
               'Como no se si la impresora acepta cant. de copias hago un loop.
186            For lCopies = 1 To Val(tCopias.Text)
187                vspReporte.PrintDoc False
188            Next

189        End If

' <VB WATCH>
190        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ActionPrint"

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

Private Sub ButtonMenu(ByVal bPlay As Boolean, ByVal bClean As Boolean, ByVal bCancel As Boolean)
' <VB WATCH>
191        On Error GoTo vbwErrHandler
' </VB WATCH>

192        With tooMenu
193            .Buttons("play").Enabled = bPlay
194            .Buttons("stop").Enabled = bCancel
195            .Buttons("print").Enabled = bClean
196        End With

' <VB WATCH>
197        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ButtonMenu"

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

Private Sub tPage_KeyPress(KeyAscii As Integer)
' <VB WATCH>
198        On Error GoTo vbwErrHandler
' </VB WATCH>
199    On Error Resume Next
200        If KeyAscii = vbKeyReturn Then
201            If IsNumeric(tPage.Text) Then
202                If CInt(tPage.Text) > 0 And CInt(tPage.Text) <= vspReporte.PageCount Then
203                    vspReporte.PreviewPage = CInt(tPage.Text)
204                Else
205                    If CInt(tPage.Text) > vspReporte.PageCount Then
206                        vspReporte.PreviewPage = vspReporte.PageCount
207                    Else
208                        vspReporte.PreviewPage = 1
209                    End If
210                End If
211                vspReporte.SetFocus
212            End If
213            SetButtonReport
214        End If
' <VB WATCH>
215        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tPage_KeyPress"

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

Private Sub vsCopias_Change()
' <VB WATCH>
216        On Error GoTo vbwErrHandler
' </VB WATCH>
217        tCopias.Text = vsCopias.Value * -1
' <VB WATCH>
218        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vsCopias_Change"

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

Private Sub vspReporte_EndDoc()
' <VB WATCH>
219        On Error GoTo vbwErrHandler
' </VB WATCH>
220        SetHeader
' <VB WATCH>
221        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vspReporte_EndDoc"

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

Private Sub SetHeader()
' <VB WATCH>
222        On Error GoTo vbwErrHandler
' </VB WATCH>
223    Dim lCantPage As Long
224    Dim lWidth As Long
225    Dim sAncho As String
226    Dim lHeader As Long

227        With vspReporte
228            bHeader = True
229            For lCantPage = 1 To .PageCount
230                .StartOverlay lCantPage

231                lWidth = (.PageWidth - .MarginLeft - .MarginRight) / 3
232                sAncho = Format(lWidth) * 2
                   'sAncho = 10 & "|^" & ((lWidth * 2) - 750)
233                sAncho = 10 & "|^" & sAncho - 10
234                lHeader = .MarginTop / 3

235                .CurrentY = lHeader * 1.8 ' + (lHeader / 2)
236                .TextAlign = taLeftBottom
237                .Font = "Tahoma"
238                .Font.Size = 14
239                .Font.Bold = True
240                .Font.Italic = True
241                .TableBorder = tbNone
242                .Table = sAncho + ";" + "|" + m_Header

243                .CurrentY = lHeader * 1.8 ' + (lHeader / 2)
244                .TextAlign = taRightBaseline ' taRightBottom
245                .Font = "tahoma"
246                .Font.Size = 9
247                .Font.Bold = True
248                .Font.Italic = False
249                .FontUnderline = True
250                .Table = ">" + CStr(lWidth - 50) + ";" + "Rige desde el: " & Format(dTitRige, "dd/mmm/yy")


251                .TextAlign = taLeftBottom
                   ' print second header line
252                .Font = "Tahoma"
253                .Font.Size = 8
254                .Font.Bold = True
255                .Font.Italic = False
256                .FontUnderline = False
257                .CurrentY = .MarginTop - .TextHeight("H")
258                .TableBorder = tbBottom
259                .Table = sDimTableArt + sDimTableTC + sDimTableTCD + sDimTableFecha + ";" + "|Código|  Artículo" & sTitTableTC & sTitTableTCD & "|Fecha"

260                .TableBorder = tbNone
261                .TextAlign = taLeftTop
262                .EndOverlay
263            Next
264            .TextAlign = 0 'taLeft
265        End With
266        bHeader = False
267        Exit Sub

' <VB WATCH>
268        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetHeader"

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

Private Sub vspReporte_EndPage()
       '    vspReporte.TableBorder = tbNone
' <VB WATCH>
269        On Error GoTo vbwErrHandler
' </VB WATCH>
' <VB WATCH>
270        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vspReporte_EndPage"

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

Private Sub vspReporte_MousePage(NewPage As Integer)
' <VB WATCH>
271        On Error GoTo vbwErrHandler
' </VB WATCH>
272        If vspReporte Is Nothing Then
273             Exit Sub
274        End If
275        SetButtonReport
' <VB WATCH>
276        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vspReporte_MousePage"

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

Private Sub vspReporte_MouseZoom(NewZoom As Integer)
' <VB WATCH>
277        On Error GoTo vbwErrHandler
' </VB WATCH>
278        If vspReporte Is Nothing Then
279             Exit Sub
280        End If
281        fsbZoom.Value = vspReporte.Zoom
' <VB WATCH>
282        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vspReporte_MouseZoom"

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

Private Sub vspReporte_NewPage()
' <VB WATCH>
283        On Error GoTo vbwErrHandler
' </VB WATCH>
284    On Error Resume Next
           'Cdo. carga el report voy enumerando en el textbox.
285        tPage.Text = Val(tPage.Text) + 1
286        tPage.Refresh
' <VB WATCH>
287        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vspReporte_NewPage"

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

Private Sub SetButtonReport()
' <VB WATCH>
288        On Error GoTo vbwErrHandler
' </VB WATCH>
289    On Error Resume Next
290    Dim iCantPag As Integer, iPrePag As Integer

291        With vspReporte
292            iCantPag = .PageCount
293            iPrePag = .PreviewPage
294        End With
295        With tooMenu
296            .Buttons("firstpage").Enabled = (iPrePag > 1)
297            .Buttons("previouspage").Enabled = (iPrePag > 1)
298            .Buttons("nextpage").Enabled = (iPrePag < iCantPag)
299            .Buttons("lastpage").Enabled = (iPrePag < iCantPag)
300        End With

301        picCopia.Enabled = (iCantPag > 0)
302        If iCantPag = 0 Then
303            tCopias.Text = ""
304        Else
305            tCopias.Text = 1
306        End If
307        tPage.Text = iPrePag
308        If iCantPag > 1 Then
309            tPage.Enabled = True
310        Else
311            tPage.Enabled = False
312        End If

' <VB WATCH>
313        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetButtonReport"

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

Private Sub StartReport()
' <VB WATCH>
314        On Error GoTo vbwErrHandler
' </VB WATCH>
315    On Error GoTo errSR
316        bCancelQuery = False
317        Screen.MousePointer = 11
318        ButtonMenu False, False, True
319        vspReporte.StartDoc
320        SetButtonReport
321        DoEvents
322        If vspReporte.Error <> 0 Then
323            MsgBox "Ocurrio un error al iniciar el reporte.", vbCritical, "ATENCIÓN"
324            vspReporte.EndDoc
325            GoTo evFin
326        End If
327        If SetVarGlobalReport Then
328            If LoadQuery Then
329                Screen.MousePointer = 0 'Saco el puntero para que tome el evento cancelar
330                DoReport
331                vspReporte.EndDoc
332            End If
333        End If

334    evFin:
335        cBase.QueryTimeout = 15
336        Screen.MousePointer = 0
337        ButtonMenu True, True, False
338        SetButtonReport
339        Exit Sub

340    errSR:
341        cBase.QueryTimeout = 15
342        MsgBox Err.Description
' <VB WATCH>
343        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "StartReport"

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

Private Function SetVarGlobalReport() As Boolean
' <VB WATCH>
344        On Error GoTo vbwErrHandler
' </VB WATCH>
345    On Error GoTo errSV
346    Dim lWidthR As Long, lCont As Long
347    Dim lWidthCtdo As Long
348        SetVarGlobalReport = False
349        dTitRige = GetDateRige      'Fecha rige
350        If LoadPlan Then
               'Dada la cantidad de cuotas que cargue doy un largo a la tabla.
351            lWidthR = vspReporte.PageWidth - vspReporte.MarginLeft - vspReporte.MarginRight - lDimTable

352            If UBound(arrCuota) >= 0 Then
353                 lCont = UBound(arrCuota) + 1
354            End If
355            If UBound(arrDiferidos) >= 0 Then
356                 lCont = lCont + UBound(arrDiferidos) + 1
357            End If
358            If lCont = 0 Then
359                 SetVarGlobalReport = True
360                 Exit Function
361            End If

362            lWidthCtdo = 920
363            sDimTableTC = "|>" & lWidthCtdo

364            lWidthR = lWidthR - lWidthCtdo
365            lCont = lCont - 1
366            If lCont = 0 Then
367                 SetVarGlobalReport = True
368                 Exit Function
369            End If
370            lWidthR = lWidthR / lCont

               'Siempre al ctdo le doy + que a los otros.
371            If (lWidthR * lCont) + lDimTable > vspReporte.PageWidth - vspReporte.MarginLeft - vspReporte.MarginRight Then
372                sDimTableTC = sDimTableTC & "|>" & lWidthR - (((lWidthR * lCont) + lDimTable) - (vspReporte.PageWidth - vspReporte.MarginLeft - vspReporte.MarginRight))
373            Else
374                sDimTableTC = sDimTableTC & "|>" & lWidthR
375            End If
376            For lCont = 2 To UBound(arrCuota)
377                sDimTableTC = sDimTableTC & "|>" & lWidthR
378            Next lCont
379            For lCont = 0 To UBound(arrDiferidos)
380                sDimTableTCD = sDimTableTCD & "|>" & lWidthR
381            Next
382            SetVarGlobalReport = True
383        End If
384        Exit Function
385    errSV:
386        clsGeneral.OcurrioError "Ocurrió un error al cargar los parámetros globales.", Err.Description, "ATENCIÓN"
' <VB WATCH>
387        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetVarGlobalReport"

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

Private Function LoadQuery() As Boolean
' <VB WATCH>
388        On Error GoTo vbwErrHandler
' </VB WATCH>
389    On Error GoTo errLQ
390    Dim sCharArt As String
391    Dim lEspecie As Long

392        LoadQuery = False
           '--------------------------------------------------------------------------------------------------------------------------------------------

           'Seteo la Query del reporte-----------------------------------------------------------------
393        Cons = "Select Articulo.*, Especie.*, Precios.HPrVigencia, Precios.HPrPrecio, TCuCodigo, TCuCantidad, TCuVencimientoC, PlaNombre " & _
                      " From " & _
                               "HistoriaPrecio Precios, " & _
                               "Articulo, Tipo, ArticuloFacturacion, Especie, " & _
                               "TipoCuota, TipoPlan" & _
                       " WHERE (HPrVigencia IN" & _
                                   " (Select MAX(H.HPrVigencia)" & _
                                   " FROM HistoriaPrecio H" & _
                                   " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                                   " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                                   " And H.HPrMoneda = Precios.HPrMoneda " & _
                                   " And H.HPrVigencia <= '" & m_Vigencia & "'" & _
                                   " )) " & _
                       " And Precios.HPrArticulo = ArtID" & _
                       " And ArtId = AFaArticulo And ArtTipo = TipCodigo " & _
                       " And TipEspecie = EspCodigo And ArtEnUso = 1 " & _
                       " And Precios.HPrMoneda = " & m_MonedaPesos & _
                       " And Precios.HPrHabilitado = 1" & _
                       " And ArticuloFacturacion.AFaLista = " & m_Lista & _
                       " And Precios.HPrTipoCuota = TipoCuota.TCuCodigo" & _
                       " And Precios.HPrPlan = TipoPlan.PlaCodigo"

394        Cons = Cons & _
                       " And TipoCuota.TCuVencimientoE is null " & _
                       " And TCuEspecial = 0 " & _
                       " and TCuDeshabilitado is null"

           'Hago Union con combos
395        Cons = Cons & _
                   " Union All " & _
                       " Select Articulo.*, Especie.*, '' as HPrVigencia, 0 as HPrPrecio, 0 as TCuCodigo, 0 as TCuCantidad, 0 as TCuVencimientoC, '' as PlaNombre " & _
                       " From Articulo, Tipo, ArticuloFacturacion, Especie, Presupuesto " & _
                       " Where ArtEsCombo = 1 And PreHabilitado = 1 And PreEsPresupuesto = 0 And PreArtCombo = ArtID " & _
                       " And ArticuloFacturacion.AFaLista = " & m_Lista & _
                       " And ArtId = AFaArticulo And ArtTipo = TipCodigo " & _
                       " And TipEspecie = EspCodigo And ArtEnUso = 1 "

396        Cons = Cons & _
                       " Order By Especie.EspNombre Asc, Articulo.ArtNombre Asc"

397        cBase.QueryTimeout = 50
398        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
399        LoadQuery = True
400        Exit Function

401    errLQ:
402        cBase.QueryTimeout = 15
403        clsGeneral.OcurrioError "Ocurrió un error al iniciar la consulta.", Err.Description, "Load Query"
' <VB WATCH>
404        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "LoadQuery"

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

Private Sub DoReport()
' <VB WATCH>
405        On Error GoTo vbwErrHandler
' </VB WATCH>
406    On Error GoTo errDR
407    Dim sCharArt As String
408    Dim lEspecie As Long, lArticulo As Long

409    Dim sPrecio As String, sDiferido As String
410    Dim lY As Long


411    Dim mEspecie As Long, mArticulo As Long
412    Dim mNameArticulo As String, mPlan As String
413    Dim mLetra As String

414        ReDim arrPrecios(UBound(arrCuota))
415        ReDim arrDif(UBound(arrDiferidos))

416        Do While Not RsAux.EOF


417            mEspecie = RsAux!EspCodigo
418            mArticulo = RsAux!ArtCodigo

419            If lArticulo = 0 Then
420                lArticulo = mArticulo
421                mNameArticulo = "|" & Format(RsAux!ArtCodigo, "#,000,000") & "|" & Trim(RsAux!ArtNombre)
422                mLetra = Mid(RsAux!ArtNombre, 1, 1)
423                mPlan = "|" & Format(RsAux!HPrVigencia, "dd/mm") & " " & Trim(RsAux!PlaNombre)
424                NuevaEspecie
425                lEspecie = mEspecie
426            End If

427            If lArticulo <> mArticulo Then

428                lArticulo = mArticulo
429                With vspReporte
430                    .Font = "Tahoma"
431                    .FontSize = 8
432                    .FontBold = False
433                    .FontItalic = False
434                    .FontUnderline = False

435                    lHeightCol = .TextHeight("HOLA")
436                    sPrecio = Join(arrPrecios, "|")
437                    sDiferido = Join(arrDif, "|")
438                    .TableBorder = tbBottom
439                    .SpaceAfter = 50

440                    If sCharArt <> mLetra Then
441                        sCharArt = mLetra
442                        .AddTable sDimTableArt + sDimTableTC + sDimTableTCD + sDimTableFecha, "", sCharArt & mNameArticulo & "|$ " & sPrecio & "|" & Join(arrDif, "|") & mPlan
443                    Else
444                        .AddTable sDimTableArt + sDimTableTC + sDimTableTCD + sDimTableFecha, "", mNameArticulo & "|$ " & sPrecio & "|" & Join(arrDif, "|") & mPlan
445                    End If
446                End With


447                mNameArticulo = "|" & Format(RsAux!ArtCodigo, "#,000,000") & "|" & Trim(RsAux!ArtNombre)
448                mLetra = Mid(RsAux!ArtNombre, 1, 1)
449                mPlan = "|" & Format(RsAux!HPrVigencia, "dd/mm") & " " & Trim(RsAux!PlaNombre)

450                ReDim arrPrecios(UBound(arrCuota))
451                ReDim arrDif(UBound(arrDiferidos))

452                If lEspecie <> mEspecie Then
453                    lEspecie = mEspecie
454                    NuevaEspecie
455                End If

456            End If
457            Dim iIndex As Integer
458            If RsAux!TCuVencimientoC = 0 Then
459                If RsAux!ArtEsCombo Then
                       'Cargo todos los precios para el combo
460                    s_LoadPrecioCombo
461                Else

462                    iIndex = GetColPlan(arrCuota, RsAux!TCuCodigo)
463                    If iIndex >= 0 Then
464                         arrPrecios(iIndex) = RsAux!HPrPrecio / RsAux!TCuCantidad
465                    End If
466                End If
467            Else
468                iIndex = GetColPlan(arrDiferidos, RsAux!TCuCodigo)
469                If iIndex >= 0 Then
470                     arrDif(iIndex) = RsAux!HPrPrecio / RsAux!TCuCantidad
471                End If
472            End If

473            RsAux.MoveNext

474            If RsAux.EOF Then
475                lArticulo = mArticulo
476                With vspReporte
477                    .Font = "Tahoma"
478                    .FontSize = 8
479                    .FontBold = False
480                    .FontItalic = False
481                    .FontUnderline = False
482                    lHeightCol = .TextHeight("HOLA")
483                    sPrecio = Join(arrPrecios, "|")
484                    sDiferido = Join(arrDif, "|")
485                    .TableBorder = tbBottom
486                    If sCharArt <> mLetra Then
487                        sCharArt = mLetra
488                        .AddTable sDimTableArt + sDimTableTC + sDimTableTCD + sDimTableFecha, "", sCharArt & mNameArticulo & "|$ " & sPrecio & "|" & Join(arrDif, "|") & mPlan
489                    Else
490                        .AddTable sDimTableArt + sDimTableTC + sDimTableTCD + sDimTableFecha, "", mNameArticulo & "|$ " & sPrecio & "|" & Join(arrDif, "|") & mPlan
491                    End If
492                End With
493            End If
494            DoEvents
495            If bCancelQuery Then
496                 Exit Do
497            End If

498        Loop
499        RsAux.Close

500        Erase arrPrecios
501        Erase arrDif
502        Exit Sub
503    errDR:
504        clsGeneral.OcurrioError "Ocurrió un error al cargar el reporte.", Err.Description, "Error (doreport)"
' <VB WATCH>
505        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DoReport"

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
Private Function GetColPlan(ByVal arrCol As Variant, ByVal lPlan As Long) As Integer
' <VB WATCH>
506        On Error GoTo vbwErrHandler
' </VB WATCH>
507    Dim iIndex As Integer
508        GetColPlan = -1
509        For iIndex = 0 To UBound(arrCol)
510            If arrCol(iIndex) = lPlan Then
511                 GetColPlan = iIndex
512                 Exit For
513            End If
514        Next
' <VB WATCH>
515        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetColPlan"

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


Private Sub NuevaEspecie()
' <VB WATCH>
516        On Error GoTo vbwErrHandler
' </VB WATCH>
517    Dim p As Long

518        With vspReporte
519            .TableBorder = tbNone
520            .FontBold = True
521            .Font = "Wingdings"
522            .FontSize = 9
523            .CurrentY = .CurrentY + 100
524            p = .CurrentY
525            .Text = "|"
526            .FontUnderline = True
527            .FontSize = 8
528            .Font = "tahoma"
529            .CurrentY = p
530            .TextAlign = taRightTop
531            .AddTable CStr(.PageWidth - .MarginLeft - .MarginRight - 400), "", Trim(RsAux!EspNombre)
532            .FontBold = False
533            .FontUnderline = False
534            .TextAlign = 0
535        End With


' <VB WATCH>
536        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "NuevaEspecie"

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

Private Function LoadPlan() As Boolean
' <VB WATCH>
537        On Error GoTo vbwErrHandler
' </VB WATCH>
538    On Error GoTo errLP
539    Dim iIndex As Integer, iIndexD As Integer

540        LoadPlan = False
541        ReDim arrCuota(0)
542        ReDim arrDiferidos(0)

543        iIndex = 0
544        iIndexD = 0
545        sTitTableTC = ""
546        sTitTableTCD = ""
547        sDimTableTC = ""
548        sDimTableTCD = ""

549        Cons = "Select * From TipoCuota" _
                & " Where TCuVencimientoE Is Null And TCuEspecial = 0 " _
                & " And TCuDeshabilitado Is Null Order By TCuOrden"

550        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
551        Do While Not RsAux.EOF
552            If ExistePlanEnPrecio(RsAux!TCuCodigo) Then
553                If RsAux!TCuVencimientoC = 0 Then
554                    ReDim Preserve arrCuota(iIndex)
555                    arrCuota(iIndex) = RsAux!TCuCodigo
556                    iIndex = iIndex + 1
557                    sTitTableTC = sTitTableTC & "|" & Trim(RsAux!TCuAbreviacion)
558                Else
559                    sTitTableTCD = sTitTableTCD & "|" & Trim(RsAux!TCuAbreviacion)
560                    ReDim Preserve arrDiferidos(iIndexD)
561                    arrDiferidos(iIndexD) = RsAux!TCuCodigo
562                    iIndexD = iIndexD + 1
563                End If
564            End If
565            RsAux.MoveNext
566        Loop
567        RsAux.Close
568        LoadPlan = True
569        Exit Function
570    errLP:
571        clsGeneral.OcurrioError "Ocurrió un error al cargar los tipos de cuotas.", Err.Description, "ATENCIÓN"
' <VB WATCH>
572        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "LoadPlan"

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

Private Function GetDateRige() As Date
' <VB WATCH>
573        On Error GoTo vbwErrHandler
' </VB WATCH>

574        GetDateRige = Now
575        Cons = "Select Max(Precios.HPrVigencia) from HistoriaPrecio Precios, Articulo, ArticuloFacturacion " & _
                      " Where (Precios.HPrVigencia IN " & _
                           " (Select MAX(H.HPrVigencia)" & _
                                   " FROM HistoriaPrecio H" & _
                                   " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                                   " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                                   " And H.HPrMoneda = Precios.HPrMoneda " & _
                                   " And H.HPrVigencia <= '" & m_Vigencia & "'" & _
                                   " )) " & _
                       " And Precios.HPrArticulo = ArtID" & _
                       " And ArtId = AFaArticulo " & _
                       " And ArtEnUso = 1 " & _
                       " And Precios.HPrMoneda = " & m_MonedaPesos & _
                       " And Precios.HPrHabilitado = 1" & _
                       " And AFaLista = " & m_Lista

576        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
577        If Not IsNull(RsAux(0)) Then
578             GetDateRige = RsAux(0)
579        End If
580        RsAux.Close

' <VB WATCH>
581        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetDateRige"

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

Private Sub vspReporte_NewTableCell(Row As Integer, Column As Integer, Cell As String)
' <VB WATCH>
582        On Error GoTo vbwErrHandler
' </VB WATCH>
583    On Error Resume Next
584    Dim lColor

585        If bHeader Then
586             Exit Sub
587        End If
588        vspReporte.FontItalic = False
589        If Column = 4 Or Column = 1 Then
590             vspReporte.FontBold = True
591        Else
592             vspReporte.FontBold = False
593        End If

594        If Column = UBound(arrCuota) + UBound(arrDiferidos) + 6 Or Column = UBound(arrCuota) + 5 Then
595            vspReporte.DrawLine vspReporte.MarginLeft, vspReporte.CurrentY, vspReporte.MarginLeft, vspReporte.CurrentY + lHeightCol + vspReporte.SpaceAfter
596        End If

597        If Column > 4 And Column < 10 Then
598            If IsNumeric(Cell) Then
599                If CCur(Cell) < paCuotaMin Then
600                    vspReporte.FontItalic = True
601                End If
602            End If
603        End If

' <VB WATCH>
604        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vspReporte_NewTableCell"

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

Private Function ExistePlanEnPrecio(ByVal idPlan As Long) As Boolean
' <VB WATCH>
605        On Error GoTo vbwErrHandler
' </VB WATCH>
606    Dim rsTC As rdoResultset
607        ExistePlanEnPrecio = False
608        Cons = "Select Top 1 * " & _
                      " From " & _
                               "HistoriaPrecio Precios, " & _
                               "Articulo, ArticuloFacturacion " & _
                       " WHERE (HPrVigencia IN" & _
                                   " (Select MAX(H.HPrVigencia)" & _
                                   " FROM HistoriaPrecio H" & _
                                   " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                                   " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                                   " And H.HPrMoneda = Precios.HPrMoneda " & _
                                   " And H.HPrVigencia <= '" & m_Vigencia & "'" & _
                                   " )) " & _
                       " And Precios.HPrArticulo = ArtID" & _
                       " And ArtId = AFaArticulo " & _
                       " And ArtEnUso = 1 " & _
                       " And Precios.HPrMoneda = " & m_MonedaPesos & _
                       " And Precios.HPrHabilitado = 1" & _
                       " And ArticuloFacturacion.AFaLista = " & m_Lista & _
                       " And Precios.HPrTipoCuota = " & idPlan
609        Set rsTC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
610        If Not rsTC.EOF Then
611            ExistePlanEnPrecio = True
612        End If
613        rsTC.Close
' <VB WATCH>
614        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ExistePlanEnPrecio"

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

Private Sub s_LoadPrecioCombo()
' <VB WATCH>
615        On Error GoTo vbwErrHandler
' </VB WATCH>
616    Dim rsPC As rdoResultset
617    Dim iIndex As Integer

618            Cons = "Select PViTipoCuota, MAx(PlaNombre) as PlaNombre, TCuCodigo, TCuCantidad, TCuAbreviacion, TCuVencimientoC, sum((PViPrecio * PArCantidad)) as Precio, Count(*) as Cant " & _
                       " From Presupuesto, PresupuestoArticulo, PrecioVigente, TipoCuota, TipoPlan " & _
                       " Where PreArtCombo = " & RsAux!ArtID & " And PreID = PArPresupuesto And PViArticulo = PArArticulo " & _
                       " And TCuVencimientoE Is Null And TCuEspecial = 0 And TCuDeshabilitado Is Null " & _
                       " And PViHabilitado <> 0  And PViMoneda = 1 And PViTipoCuota = TCuCodigo And PViPlan = PlaCodigo " & _
                       " Group By PViTipoCuota, TCuCodigo, TCuCAntidad, TcuAbreviacion, TCuVencimientoC " & _
                       " Order by PViTipoCuota"

619            Set rsPC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
620            Do While Not rsPC.EOF
621                If rsPC!TCuVencimientoC = 0 Then
622                    iIndex = GetColPlan(arrCuota, rsPC!TCuCodigo)
623                    If iIndex >= 0 Then
624                         arrPrecios(iIndex) = Format(rsPC("Precio") / rsPC("TCuCantidad"), "###0")
625                    End If
626                Else
627                    iIndex = GetColPlan(arrDiferidos, rsPC!TCuCodigo)
628                    If iIndex >= 0 Then
629                         arrDif(iIndex) = Format(rsPC("Precio") / rsPC("TCuCantidad"), "###0")
630                    End If
631                End If
632                rsPC.MoveNext
633            Loop
634            rsPC.Close

' <VB WATCH>
635        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "s_LoadPrecioCombo"

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

