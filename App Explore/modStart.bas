Attribute VB_Name = "modStart"
Option Explicit

Global Const NORMAL_PRIORITY_CLASS = &H20
Global Const INFINITE = &HFFFF

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, _
                                ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, _
                                ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
                                ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, _
                                ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, _
                                lpProcessInformation As PROCESS_INFORMATION) As Long
                                
                                
Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

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

Sub Main()
    
    On Error GoTo errMain
    Dim aTexto As String
    
    aTexto = Trim(Command())
    
    If Trim(aTexto) <> "" Then
'        frmSumi.prmMensaje = aTexto
'        Load frmSumi
        Dim vPrm As String
        If InStr(1, aTexto, "CGSAQuery:", vbTextCompare) > 0 Then
            EjecutarApp "C:\AA Aplicaciones\System\cgsareporteador.exe", aTexto
        Else
            If InStr(1, aTexto, ":") > 0 Then vPrm = "CGSAPlantilla:" Else vPrm = "CGSAMensaje:"
            EjecutarApp "C:\AA Aplicaciones\System\cgsareporteador.exe", vPrm & aTexto & "[PRM/]"
        End If
        'EjecutarApp "C:\Desarrollo\Visual Basic\AA Aplicaciones\System\cgsareporteador.exe", vPrm & aTexto & "[PRM/]"
        'C:\AA Aplicaciones\System
        
        End
    End If
    Exit Sub
errMain:
    MsgBox "Error al activar explorador de mensajes." & vbCrLf & _
                Err.Number & "- " & Err.Description, vbCritical, "(Main) Error de Aplicación - Explorador MSG"
    End
End Sub

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
