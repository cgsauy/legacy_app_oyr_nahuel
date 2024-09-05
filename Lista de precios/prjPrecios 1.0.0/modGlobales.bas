Attribute VB_Name = "ModGlobales"
Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Colores.-----------------------------------------
Public Enum Colores
    Obligatorio = &HC0FFFF
    Inactivo = &HE0E0E0
    Blanco = &HFFFFFF
    Rojo = &H80&
    RojoClaro = &HC0&
    Gris = &HE0E0E0
    GrisOscuro = &H8000000F
    Azul = &H800000
    clNaranja = &HC0E0FF
    clVioleta = &HFFC0C0      'violeta
    clVerde = &HB5DDA7
    clCeleste = &HFFEEB0
    clRosado = &HFFC0EE
End Enum

'Constantes.------------------------------------
Public Const sqlFormatoFH = "mm/dd/yyyy hh:nn:ss"
Public Const sqlFormatoF = "mm/dd/yyyy"
Public Const FormatoFP = "dd/mm/yyyy"
Public Const FormatoFHP = "d-Mmm yyyy hh:mm:ss"
Public Const FormatoMonedaP = "#,##0.00"

Public I As Integer

Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, _
                                ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, _
                                ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
                                ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, _
                                ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, _
                                lpProcessInformation As PROCESS_INFORMATION) As Long

Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Byte
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Global Const NORMAL_PRIORITY_CLASS = &H20
Global Const INFINITE = &HFFFF


'------------------------------------------------------------------------------------------------
 '   Funciones:
'       TextoValido (S as String)
'       ValidoFormatoFolder(Folder As String)
'       Encripto (S As String)
'       DesEncripto (S As String)
'------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------
 '   Procedimientos:
'       Botones(Nu As Boolean, Mo As Boolean, El As Boolean, Gr As Boolean, Ca As Boolean, Toolbar1 As Control, nForm As Form)
'       BotonesRegistro(Pri As Boolean, Ant As Boolean, Sig As Boolean, Ult As Boolean, Toolbar1 As Control, nForm As Form)
'------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
'   Habilita y deshabilita los botones y menus del toolbar
'-----------------------------------------------------------------------------------
' <VB WATCH>
Const VBWMODULE = "ModGlobales"
' </VB WATCH>

Public Sub Botones(Nu As Boolean, Mo As Boolean, El As Boolean, Gr As Boolean, Ca As Boolean, Toolbar1 As Control, nForm As Form)
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>

           'Habilito y Desabilito Botones.
2          Toolbar1.Buttons("nuevo").Enabled = Nu
3          nForm.MnuNuevo.Enabled = Nu

4          Toolbar1.Buttons("modificar").Enabled = Mo
5          nForm.MnuModificar.Enabled = Mo

6          Toolbar1.Buttons("eliminar").Enabled = El
7          nForm.MnuEliminar.Enabled = El

8          Toolbar1.Buttons("grabar").Enabled = Gr
9          nForm.MnuGrabar.Enabled = Gr

10         Toolbar1.Buttons("cancelar").Enabled = Ca
11         nForm.MnuCancelar.Enabled = Ca

' <VB WATCH>
12         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Botones"

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

'-----------------------------------------------------------------------------------
'   Habilita y deshabilita los botones y menus de registros del toolbar.
'-----------------------------------------------------------------------------------
Public Sub BotonesRegistro(Pri As Boolean, Ant As Boolean, Sig As Boolean, Ult As Boolean, Toolbar1 As Control, nForm As Form)
' <VB WATCH>
13         On Error GoTo vbwErrHandler
' </VB WATCH>

           'Habilito y Desabilito Botones.
14         Toolbar1.Buttons("primero").Enabled = Pri
15         nForm.MnuPrimero.Enabled = Pri

16         Toolbar1.Buttons("anterior").Enabled = Ant
17         nForm.MnuAnterior.Enabled = Ant

18         Toolbar1.Buttons("siguiente").Enabled = Sig
19         nForm.MnuSiguiente.Enabled = Sig

20         Toolbar1.Buttons("ultimo").Enabled = Ult
21         nForm.MnuUltimo.Enabled = Ult

' <VB WATCH>
22         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "BotonesRegistro"

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

'--------------------------------------------------------------------------------------------------------
'   PROCEDMIENTO BuscoCodigoEnCombo: Busca un el codigo pasado como parámetro dentro del itemData del combo.
'
'   PARÁMETROS:
'       lngCodigo: Codigo a buscar.
'
'   RETORNA:
'       Si encuentra el dato, setea automáticamente el combo, sino lo marca en vacio.
'--------------------------------------------------------------------------------------------------------

Public Sub BuscoCodigoEnCombo(cCombo As Control, lngCodigo As Long)
' <VB WATCH>
23         On Error GoTo vbwErrHandler
' </VB WATCH>
24     Dim I As Integer

25         If cCombo.ListCount > 0 Then
26             For I = 0 To cCombo.ListCount - 1
27                 If cCombo.ItemData(I) = lngCodigo Then
28                     cCombo.ListIndex = I
29                     Exit Sub
30                 End If
31             Next I
32             cCombo.ListIndex = -1
33         Else
34             cCombo.ListIndex = -1
35         End If

' <VB WATCH>
36         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "BuscoCodigoEnCombo"

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

Public Sub Foco(C As Control)
' <VB WATCH>
37         On Error GoTo vbwErrHandler
' </VB WATCH>
38         On Error Resume Next
39         If C.Enabled Then
40             C.SelStart = 0
41             C.SelLength = Len(C.Text)
42             C.SetFocus
43         End If
' <VB WATCH>
44         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Foco"

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

'--------------------------------------------------------------------------------------------------------
'   PROCEDMIENTO CargoCombo: Carga el combo con los datos de la consulta pasada
'   como parámetro.
'
'   PARÁMETROS:
'       Cons: Cosulta seleccionando los datos a cargar - RS(0) = Codigo, RS(1) = Dato.
'       Combo: Combo a cargar.
'       Seleccionado: Dato a seleccionar por defecto (Texto).
'--------------------------------------------------------------------------------------------------------
Public Sub CargoCombo(Consulta As String, Combo As Control, Optional Seleccionado As String = "")
' <VB WATCH>
45         On Error GoTo vbwErrHandler
' </VB WATCH>

46     Dim RsAuxiliar As rdoResultset
47     Dim iSel As Integer 'Guardo el indice del seleccionado
48     iSel = -1

49     On Error GoTo ErrCC

50         Screen.MousePointer = 11
51         Combo.Clear
52         Set RsAuxiliar = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)

53         Do While Not RsAuxiliar.EOF
54             Combo.AddItem Trim(RsAuxiliar(1))
55             Combo.ItemData(Combo.NewIndex) = RsAuxiliar(0)

56             If Trim(RsAuxiliar(1)) = Trim(Seleccionado) Then
57                  iSel = Combo.ListCount
58             End If
59             RsAuxiliar.MoveNext
60         Loop
61         RsAuxiliar.Close

62         If iSel = -1 Then
63              Combo.ListIndex = iSel
64         Else
65              Combo.ListIndex = iSel - 1
66         End If
67         Screen.MousePointer = 0
68         Exit Sub

69     ErrCC:
70         Screen.MousePointer = 0
71         MsgBox "Ocurrió un error al cargar el combo: " & Trim(Combo.Name) & "." & vbCrLf & Err.Description, vbCritical, "ERROR"
' <VB WATCH>
72         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CargoCombo"

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

Public Function SacoEnter(Texto) As String
' <VB WATCH>
73         On Error GoTo vbwErrHandler
' </VB WATCH>
74     Dim aTexto As String

75         I = 1
76         aTexto = ""
77         On Error GoTo errEnter
78         If InStr(1, Texto, Chr(13), vbTextCompare) = 0 Then
79              SacoEnter = Texto
80              Exit Function
81         End If
82         Do While I <= Len(Texto)
83             If Asc(Mid(Texto, I, 1)) = vbKeyReturn Then
84                 If Asc(Mid(Texto, I + 1, 1)) = 10 Then
85                      I = I + 2
86                 Else
87                      I = I + 1
88                 End If
89                 aTexto = aTexto & " "
90             Else
91                 aTexto = aTexto & Mid(Texto, I, 1)
92                 I = I + 1
93             End If
94         Loop
95         SacoEnter = Trim(aTexto)
96         Exit Function

97     errEnter:
98         SacoEnter = Texto
' <VB WATCH>
99         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SacoEnter"

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

Public Sub EjecutarApp(Path As String, Optional Valor As String = "", Optional Modal As Boolean = False, Optional bAtrapoError As Boolean = True)
' <VB WATCH>
100        On Error GoTo vbwErrHandler
' </VB WATCH>

101        If bAtrapoError Then
102             On Error GoTo errApp
103        End If

104        Screen.MousePointer = 11
105        Dim plngRet As Long

106        If Not Modal Then   'APLICACION NO MODAL--------------------------------------------------------------------------------
107            If Valor = "" Then
108                Dim aPos As Integer
109                aPos = InStr(Path, ".exe")
110                If aPos <> 0 Then
111                    Valor = Mid(Path, aPos + Len(".exe") + 1)
112                    Path = Mid(Path, 1, aPos - 1)
113                End If
114            End If
115            plngRet = ShellExecute(0, "open", Path, Valor, 0, 1)
116            If plngRet = 0 And Err.Description <> "" Then
117                 GoTo errApp
118            End If

119        Else                       'APLICACION MODAL------------------------------------------------------------------------------------
120            Dim pobjProcess As PROCESS_INFORMATION, pobjStart As STARTUPINFO

121            If Trim(Valor) <> "" Then
122                 Path = Path & " " & Valor
123            End If

               'Inicializa la estructura STARTUPINFO
124            pobjStart.cb = Len(pobjStart)
               'pobjStart.dwFlags = STARTF_USESHOWWINDOW

               'pobjstart.wShowWindow=

               'Dispara la aplicación
125            plngRet = CreateProcessA(0, Path, 0, 0, 1, NORMAL_PRIORITY_CLASS, 0, 0, pobjStart, pobjProcess)

               ' Espera a que termine la aplicación
126            plngRet = WaitForSingleObject(pobjProcess.hProcess, INFINITE)
127            plngRet = CloseHandle(pobjProcess.hProcess)
               '----------------------------------------------------------------------------------------------------------
128        End If
129        Screen.MousePointer = 0
130        Exit Sub

131    errApp:
132        Screen.MousePointer = 0
133        MsgBox "Ocurrió un error al ejecutar la aplicación " & Path & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbCritical, "Error de Aplicación"
' <VB WATCH>
134        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EjecutarApp"

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

Public Sub snd_ActivarSonido(sndFile As String)
' <VB WATCH>
135        On Error GoTo vbwErrHandler
' </VB WATCH>

136        On Error Resume Next
137        Dim Result As Long

138        Result = sndPlaySound(sndFile, 1)

' <VB WATCH>
139        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "snd_ActivarSonido"

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

