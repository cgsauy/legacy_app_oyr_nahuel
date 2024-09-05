VERSION 5.00
Begin VB.Form frmFiltroEtiqueta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Etiquetas según filtro"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bCancel 
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton bAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox tCantEtiqueta 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Text            =   "1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.VScrollBar vscCantidad 
      Height          =   285
      Left            =   1920
      Min             =   1
      TabIndex        =   10
      Top             =   1680
      Value           =   1
      Width           =   255
   End
   Begin VB.ComboBox cQueEtiqueta 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox tLista 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox tMarca 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox tTipo 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox tFecha 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cantidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "¿&Qué etiqueta?:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Lista:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Marca:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo de Artículo:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Precio modificado después del:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmFiltroEtiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private colArticulo As New Collection
Private m_CantE As Long, m_QueEtiqueta As Integer

' <VB WATCH>
Const VBWMODULE = "frmFiltroEtiqueta"
' </VB WATCH>

Public Property Get prmHayDatos() As Boolean
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>
2          If colArticulo Is Nothing Then
3              prmHayDatos = False
4          Else
5              If colArticulo.Count = 0 Then
6                  prmHayDatos = False
7              Else
8                  prmHayDatos = True
9              End If
10         End If
' <VB WATCH>
11         Exit Property
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "prmHayDatos[Get]"

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
Public Property Get prmCantResultado() As Long
' <VB WATCH>
12         On Error GoTo vbwErrHandler
' </VB WATCH>
13         prmCantResultado = colArticulo.Count
' <VB WATCH>
14         Exit Property
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "prmCantResultado[Get]"

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

Public Property Get prmIDResultado(ByVal lPos As Long) As Long
' <VB WATCH>
15         On Error GoTo vbwErrHandler
' </VB WATCH>
16         If colArticulo Is Nothing Then
17              Exit Sub
18         End If
19         If lPos > colArticulo.Count Or lPos < 1 Then
20              Exit Sub
21         End If
22         prmIDResultado = colArticulo(lPos)
' <VB WATCH>
23         Exit Property
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "prmIDResultado[Get]"

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

Public Property Get prmCantidad() As Long
' <VB WATCH>
24         On Error GoTo vbwErrHandler
' </VB WATCH>
25         prmCantidad = m_CantE
' <VB WATCH>
26         Exit Property
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "prmCantidad[Get]"

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

Public Property Get prmQueEtiqueta() As Integer
' <VB WATCH>
27         On Error GoTo vbwErrHandler
' </VB WATCH>
28         prmQueEtiqueta = m_QueEtiqueta
' <VB WATCH>
29         Exit Property
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "prmQueEtiqueta[Get]"

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

Private Sub bAplicar_Click()
       'Consulto dados los filtros
' <VB WATCH>
30         On Error GoTo vbwErrHandler
' </VB WATCH>
31         m_QueEtiqueta = -1
32         m_CantE = 0
33         Set colArticulo = Nothing
34         If tFecha.Text <> "" Or Val(tTipo.Tag) > 0 Or Val(tMarca.Tag) > 0 Or Val(tLista.Tag) > 0 Then

35             Cons = "Select Distinct(ArtCodigo) From " & _
                               "HistoriaPrecio Precios, " & _
                               "Articulo, ArticuloFacturacion, TipoCuota " & _
                       " Where Precios.HPrArticulo = ArtID"

36             If IsDate(tFecha.Text) Then
37                 Cons = Cons & " And HPrVigencia >= '" & Format(tFecha.Text, "mm/dd/yyyy 00:00:00") & "'"
38             End If
39             If Val(tTipo.Tag) > 0 Then
40                  Cons = Cons & " And ArtTipo = " & Val(tTipo.Tag)
41             End If
42             If Val(tMarca.Tag) > 0 Then
43                  Cons = Cons & " And ArtMarca = " & Val(tMarca.Tag)
44             End If
45             If Val(tLista.Tag) > 0 Then
46                  Cons = Cons & " And AFaLista = " & Val(tLista.Tag)
47             End If


48             Cons = Cons & " And ArtId = AFaArticulo " & _
                       " And ArtEnUso = 1" & _
                       " And Precios.HPrMoneda = " & paMonedaPesos & _
                       " And Precios.HPrHabilitado = 1" & _
                       " And Precios.HPrTipoCuota = TipoCuota.TCuCodigo"
49             Cons = Cons & _
                       " And TipoCuota.TCuVencimientoE Is Null " & _
                       " And TCuEspecial = 0 And TCuDeshabilitado Is Null"

50             Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

51             If RsAux.EOF Then
52                 MsgBox "No hay datos para los filtros ingresados.", vbInformation, "ATENCIÓN"
53                 RsAux.Close
54             Else
55                 Do While Not RsAux.EOF
56                     colArticulo.Add CStr(RsAux(0))
57                     RsAux.MoveNext
58                 Loop
59                 RsAux.Close
60                 m_QueEtiqueta = cQueEtiqueta.ListIndex
61                 m_CantE = CLng(tCantEtiqueta.Text)
62                 Unload Me
63             End If
64         Else
65             MsgBox "No se ingresaron filtros.", vbExclamation, "ATENCIÓN"
66         End If

' <VB WATCH>
67         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "bAplicar_Click"

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

Private Sub bCancel_Click()
' <VB WATCH>
68         On Error GoTo vbwErrHandler
' </VB WATCH>
69         Unload Me
' <VB WATCH>
70         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "bCancel_Click"

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
71         On Error GoTo vbwErrHandler
' </VB WATCH>
72         If KeyAscii = vbKeyReturn Then
73              bAplicar.SetFocus
74         End If
' <VB WATCH>
75         Exit Sub
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

Private Sub Form_Load()
' <VB WATCH>
76         On Error GoTo vbwErrHandler
' </VB WATCH>
77         With cQueEtiqueta
78             .Clear
79             .AddItem "Ambas"
80             .AddItem "Normal (chica)"
81             .AddItem "Según tabla"
82             .AddItem "Vidriera (grande)"
83             .ListIndex = 2
84         End With
' <VB WATCH>
85         Exit Sub
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

Private Sub tCantEtiqueta_GotFocus()
' <VB WATCH>
86         On Error GoTo vbwErrHandler
' </VB WATCH>
87         With tCantEtiqueta
88             .SelStart = 0
89             .SelLength = Len(.Text)
90         End With
91         If Val(tCantEtiqueta.Text) = 0 Then
92              tCantEtiqueta.Text = vscCantidad.Value
93         End If
' <VB WATCH>
94         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tCantEtiqueta_GotFocus"

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

Private Sub tCantEtiqueta_KeyPress(KeyAscii As Integer)
' <VB WATCH>
95         On Error GoTo vbwErrHandler
' </VB WATCH>
96         If KeyAscii = vbKeyReturn Then
97             If IsNumeric(tCantEtiqueta.Text) Then
98                 If Val(tCantEtiqueta.Text) < 1 Then
99                      tCantEtiqueta.Text = vscCantidad.Value
100                End If
101                vscCantidad.Value = Val(tCantEtiqueta.Text)
102                cQueEtiqueta.SetFocus
103            Else
104                MsgBox "Formato incorrecto.", vbExclamation, "ATENCIÓN"
105                tCantEtiqueta.Text = vscCantidad.Value
106            End If
107        End If
' <VB WATCH>
108        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tCantEtiqueta_KeyPress"

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

Private Sub tCantEtiqueta_LostFocus()
' <VB WATCH>
109        On Error GoTo vbwErrHandler
' </VB WATCH>
110        If Not IsNumeric(tCantEtiqueta.Text) Then
111            tCantEtiqueta.Text = vscCantidad.Value
112        Else
113            If Val(tCantEtiqueta.Text) < 1 Then
114                 tCantEtiqueta.Text = vscCantidad.Value
115            End If
116        End If
' <VB WATCH>
117        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tCantEtiqueta_LostFocus"

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

Private Sub tFecha_GotFocus()
' <VB WATCH>
118        On Error GoTo vbwErrHandler
' </VB WATCH>
119        With tFecha
120            .SelStart = 0
121            .SelLength = Len(.Text)
122        End With
' <VB WATCH>
123        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tFecha_GotFocus"

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

Private Sub tFecha_KeyPress(KeyAscii As Integer)
' <VB WATCH>
124        On Error GoTo vbwErrHandler
' </VB WATCH>
125    On Error Resume Next
126        If KeyAscii = vbKeyReturn Then
127            If IsDate(tFecha.Text) Then
128                tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
129                tTipo.SetFocus
130            Else
131                If tFecha.Text <> "" Then
132                    MsgBox "Formato de fecha incorrecto.", vbExclamation, "ATENCIÓN"
133                    tFecha.Text = ""
134                Else
135                    tTipo.SetFocus
136                End If
137            End If
138        End If
' <VB WATCH>
139        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tFecha_KeyPress"

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

Private Sub tFecha_LostFocus()
' <VB WATCH>
140        On Error GoTo vbwErrHandler
' </VB WATCH>
141    On Error Resume Next
142        If IsDate(tFecha.Text) Then
143            tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
144        Else
145            tFecha.Text = ""
146        End If
' <VB WATCH>
147        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tFecha_LostFocus"

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

Private Sub tLista_Change()
' <VB WATCH>
148        On Error GoTo vbwErrHandler
' </VB WATCH>
149        If Val(tLista.Tag) > 0 Then
150             tLista.Tag = ""
151        End If
' <VB WATCH>
152        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tLista_Change"

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

Private Sub tLista_KeyPress(KeyAscii As Integer)
' <VB WATCH>
153        On Error GoTo vbwErrHandler
' </VB WATCH>
154    On Error Resume Next

155        If KeyAscii = vbKeyReturn Then
156            If Trim(tLista.Text) = "" Or Val(tLista.Tag) > 0 Then
157                tCantEtiqueta.SetFocus
158            Else
159                If Trim(tLista.Text) <> "" Then
160                    Cons = "Select LDPCodigo, LDPDescripcion as 'Lista de Precios' From ListasDePrecios " _
                           & " Where  LDPDescripcion Like '" & Replace(tLista.Text, " ", "%") & "%'" _
                           & "Order by LDPDescripcion"
161                    ListaAyuda Cons, tLista, "Listas de Precios"
162                    If Val(tLista.Tag) > 0 Then
163                         tCantEtiqueta.SetFocus
164                    End If
165                End If
166            End If
167        End If

' <VB WATCH>
168        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tLista_KeyPress"

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

Private Sub tMarca_Change()
' <VB WATCH>
169        On Error GoTo vbwErrHandler
' </VB WATCH>
170        If Val(tMarca.Text) > 0 Then
171             tMarca.Tag = ""
172        End If
' <VB WATCH>
173        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tMarca_Change"

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

Private Sub tMarca_GotFocus()
' <VB WATCH>
174        On Error GoTo vbwErrHandler
' </VB WATCH>
175        With tMarca
176            .SelStart = 0
177            .SelLength = Len(.Text)
178        End With
' <VB WATCH>
179        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tMarca_GotFocus"

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

Private Sub tMarca_KeyPress(KeyAscii As Integer)
' <VB WATCH>
180        On Error GoTo vbwErrHandler
' </VB WATCH>
181    On Error Resume Next

182        If KeyAscii = vbKeyReturn Then
183            If Trim(tMarca.Text) = "" Or Val(tMarca.Tag) > 0 Then
184                tLista.SetFocus
185            Else
186                If Trim(tMarca.Text) <> "" Then
187                    Cons = "Select MarCodigo, MarNombre as 'Marca' From Marca " _
                           & " Where MarNombre Like '" & Replace(tMarca.Text, " ", "%") & "%'" _
                           & "Order by MarNombre"
188                    ListaAyuda Cons, tMarca, "Lista de Marcas"
189                    If Val(tMarca.Tag) > 0 Then
190                         tLista.SetFocus
191                    End If
192                End If
193            End If
194        End If

' <VB WATCH>
195        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tMarca_KeyPress"

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

Private Sub tTipo_Change()
' <VB WATCH>
196        On Error GoTo vbwErrHandler
' </VB WATCH>
197        If Val(tTipo.Tag) > 0 Then
198             tTipo.Tag = ""
199        End If
' <VB WATCH>
200        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tTipo_Change"

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

Private Sub tTipo_GotFocus()
' <VB WATCH>
201        On Error GoTo vbwErrHandler
' </VB WATCH>
202        With tTipo
203            .SelStart = 0
204            .SelLength = Len(.Text)
205        End With
' <VB WATCH>
206        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tTipo_GotFocus"

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

Private Sub tTipo_KeyPress(KeyAscii As Integer)
' <VB WATCH>
207        On Error GoTo vbwErrHandler
' </VB WATCH>
208    On Error Resume Next

209        If KeyAscii = vbKeyReturn Then
210            If Trim(tTipo.Text) = "" Or Val(tTipo.Tag) > 0 Then
211                tMarca.SetFocus
212            Else
213                If Trim(tTipo.Text) <> "" Then
214                    Cons = "Select TipCodigo, TipNombre as 'Tipo' From Tipo " _
                           & " Where TipNombre Like '" & Replace(tTipo.Text, " ", "%") & "%'" _
                           & "Order by TipNombre"
215                    ListaAyuda Cons, tTipo, "Tipos de Artículos"
216                    If Val(tTipo.Tag) > 0 Then
217                         tMarca.SetFocus
218                    End If
219                End If
220            End If
221        End If

' <VB WATCH>
222        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tTipo_KeyPress"

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

Private Sub ListaAyuda(ByVal sCons As String, tControl As Control, ByVal sTitulo As String)
' <VB WATCH>
223        On Error GoTo vbwErrHandler
' </VB WATCH>
224    Dim objLista As New clsListadeAyuda
           '1ero hago cons. para ver cantidad.

225        Set RsAux = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
226        If RsAux.EOF Then
227            MsgBox "No se encontraron datos para el filtro ingresado.", vbExclamation, "ATENCIÓN"
228            RsAux.Close
229        Else
230            RsAux.MoveNext
231            If RsAux.EOF Then
232                RsAux.MoveFirst
233                tControl.Text = Trim(RsAux(1))
234                tControl.Tag = RsAux(0)
235                RsAux.Close
236            Else
237                RsAux.Close
238                If objLista.ActivarAyuda(cBase, sCons, 5000, 1, sTitulo) > 0 Then
239                    tControl.Text = objLista.RetornoDatoSeleccionado(1)
240                    tControl.Tag = objLista.RetornoDatoSeleccionado(0)
241                End If
242            End If
243        End If
244        Set objLista = Nothing

' <VB WATCH>
245        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ListaAyuda"

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
246        On Error GoTo vbwErrHandler
' </VB WATCH>
247        tCantEtiqueta.Text = vscCantidad.Value
' <VB WATCH>
248        Exit Sub
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

