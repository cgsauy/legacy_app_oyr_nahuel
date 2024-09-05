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
    osVerde = &H6000&
    osGris = &H808080
End Enum

'Constantes.------------------------------------
Public Const sqlFormatoFH = "mm/dd/yyyy hh:nn:ss"
Public Const sqlFormatoF = "mm/dd/yyyy"
Public Const FormatoFP = "dd/mm/yyyy"
Public Const FormatoFHP = "d-Mmm yyyy hh:mm:ss"
Public Const FormatoMonedaP = "#,##0.00"

Public i As Integer

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
Public Sub Botones(Nu As Boolean, Mo As Boolean, El As Boolean, Gr As Boolean, Ca As Boolean, Toolbar1 As Control, nForm As Form)

    'Habilito y Desabilito Botones.
    Toolbar1.Buttons("nuevo").Enabled = Nu
    nForm.MnuNuevo.Enabled = Nu
    
    Toolbar1.Buttons("modificar").Enabled = Mo
    nForm.MnuModificar.Enabled = Mo
    
    Toolbar1.Buttons("eliminar").Enabled = El
    nForm.MnuEliminar.Enabled = El
    
    Toolbar1.Buttons("grabar").Enabled = Gr
    nForm.MnuGrabar.Enabled = Gr
    
    Toolbar1.Buttons("cancelar").Enabled = Ca
    nForm.MnuCancelar.Enabled = Ca

End Sub

'-----------------------------------------------------------------------------------
'   Habilita y deshabilita los botones y menus de registros del toolbar.
'-----------------------------------------------------------------------------------
Public Sub BotonesRegistro(Pri As Boolean, Ant As Boolean, Sig As Boolean, Ult As Boolean, Toolbar1 As Control, nForm As Form)

    'Habilito y Desabilito Botones.
    Toolbar1.Buttons("primero").Enabled = Pri
    nForm.MnuPrimero.Enabled = Pri
    
    Toolbar1.Buttons("anterior").Enabled = Ant
    nForm.MnuAnterior.Enabled = Ant
    
    Toolbar1.Buttons("siguiente").Enabled = Sig
    nForm.MnuSiguiente.Enabled = Sig
    
    Toolbar1.Buttons("ultimo").Enabled = Ult
    nForm.MnuUltimo.Enabled = Ult
    
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
Dim i As Integer
    
    If cCombo.ListCount > 0 Then
        For i = 0 To cCombo.ListCount - 1
            If cCombo.ItemData(i) = lngCodigo Then
                cCombo.ListIndex = i
                Exit Sub
            End If
        Next i
        cCombo.ListIndex = -1
    Else
        cCombo.ListIndex = -1
    End If

End Sub

Public Sub Foco(c As Control)
    On Error Resume Next
    If c.Enabled Then
        c.SelStart = 0
        c.SelLength = Len(c.Text)
        c.SetFocus
    End If
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

Dim RsAuxiliar As rdoResultset
Dim iSel As Integer: iSel = -1     'Guardo el indice del seleccionado
    
On Error GoTo ErrCC
    
    Screen.MousePointer = 11
    Combo.Clear
    Set RsAuxiliar = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAuxiliar.EOF
        Combo.AddItem Trim(RsAuxiliar(1))
        Combo.ItemData(Combo.NewIndex) = RsAuxiliar(0)
        
        If Trim(RsAuxiliar(1)) = Trim(Seleccionado) Then iSel = Combo.ListCount
        RsAuxiliar.MoveNext
    Loop
    RsAuxiliar.Close
    
    If iSel = -1 Then Combo.ListIndex = iSel Else Combo.ListIndex = iSel - 1
    Screen.MousePointer = 0
    Exit Sub
    
ErrCC:
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al cargar el combo: " & Trim(Combo.Name) & "." & vbCrLf & Err.Description, vbCritical, "ERROR"
End Sub

Public Function SacoEnter(Texto) As String
Dim aTexto As String
    
    i = 1
    aTexto = ""
    On Error GoTo errEnter
    If InStr(1, Texto, Chr(13), vbTextCompare) = 0 Then SacoEnter = Texto: Exit Function
    Do While i <= Len(Texto)
        If Asc(Mid(Texto, i, 1)) = vbKeyReturn Then
            If Asc(Mid(Texto, i + 1, 1)) = 10 Then i = i + 2 Else i = i + 1
            aTexto = aTexto & " "
        Else
            aTexto = aTexto & Mid(Texto, i, 1)
            i = i + 1
        End If
    Loop
    SacoEnter = Trim(aTexto)
    Exit Function

errEnter:
    SacoEnter = Texto
End Function

Public Sub EjecutarApp(Path As String, Optional Valor As String = "", Optional Modal As Boolean = False, Optional bAtrapoError As Boolean = True)

    If bAtrapoError Then On Error GoTo errApp
    
    Screen.MousePointer = 11
    Dim plngRet As Long
    
    If Not Modal Then   'APLICACION NO MODAL--------------------------------------------------------------------------------
        If Valor = "" Then
            Dim aPos As Integer
            aPos = InStr(Path, ".exe")
            If aPos <> 0 Then
                Valor = Mid(Path, aPos + Len(".exe") + 1)
                Path = Mid(Path, 1, aPos - 1)
            End If
        End If
        plngRet = ShellExecute(0, "open", Path, Valor, 0, 1)
        If plngRet = 0 And Err.Description <> "" Then GoTo errApp
    
    Else                       'APLICACION MODAL------------------------------------------------------------------------------------
        Dim pobjProcess As PROCESS_INFORMATION, pobjStart As STARTUPINFO
        
        If Trim(Valor) <> "" Then Path = Path & " " & Valor
        
        'Inicializa la estructura STARTUPINFO
        pobjStart.cb = Len(pobjStart)
        'pobjStart.dwFlags = STARTF_USESHOWWINDOW
    
        'pobjstart.wShowWindow=
        
        'Dispara la aplicación
        plngRet = CreateProcessA(0, Path, 0, 0, 1, NORMAL_PRIORITY_CLASS, 0, 0, pobjStart, pobjProcess)
                
        ' Espera a que termine la aplicación
        plngRet = WaitForSingleObject(pobjProcess.hProcess, INFINITE)
        plngRet = CloseHandle(pobjProcess.hProcess)
        '----------------------------------------------------------------------------------------------------------
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errApp:
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al ejecutar la aplicación " & Path & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbCritical, "Error de Aplicación"
End Sub

Public Sub snd_ActivarSonido(sndFile As String)
    
    On Error Resume Next
    Dim Result As Long
    
    Result = sndPlaySound(sndFile, 1)
    
End Sub
