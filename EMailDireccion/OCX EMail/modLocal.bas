Attribute VB_Name = "modLocal"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public idX As Integer
Public mSQL As String
Public rsAux As rdoResultset

Public clsGeneral As New clsorCGSA
'-----------------------------------------------------------------
Public Type typCorreo
    ClaveCompleta As String

    ServidorID As Long
    ServidorNombre As String
    
    Direccion As String
    DireccionNombre As String
    DireccionID As Long
End Type

Public arrCorreo()  As typCorreo
'-----------------------------------------------------------------
Private arrI As Integer

Public Const taba_EMailDireccion = 8
Dim idContador As Long


Public Function arrIndex(findClave As String) As Integer
' Se pasa la clave (direccion completa de correo) y retorna el indice del array en que esta
' Si no hay datos retorna -1

On Error GoTo errFnc
    arrIndex = -1

    For arrI = LBound(arrCorreo) To UBound(arrCorreo)
        If LCase(Trim(arrCorreo(arrI).ClaveCompleta)) = Trim(LCase(findClave)) Then
            arrIndex = arrI
            Exit For
        End If
    Next

errFnc:
End Function

Public Function arrDeleteIndex(iDel As Integer) As Boolean
On Error GoTo errDelete
arrDeleteIndex = False
'   Hace copia del array y elimina la posicion iDel
Dim arrTmp() As typCorreo
Dim newI As Integer

    ReDim arrTmp(0)
    newI = 0
    
    For arrI = LBound(arrCorreo) To UBound(arrCorreo)
        
        If arrI <> iDel Then
            If arrTmp(0).ClaveCompleta <> "" Then newI = newI + 1
            ReDim Preserve arrTmp(newI)
            arrTmp(newI) = arrCorreo(arrI)
        End If
    Next
    
    arrCorreo = arrTmp
    arrDeleteIndex = True
    Exit Function
errDelete:
End Function

Public Function arrNewItem() As Integer
    'Agrega un nuevo elemento al array y retorna el index
    If arrCorreo(0).ClaveCompleta = "" Then
        arrNewItem = 0
    Else
        arrNewItem = UBound(arrCorreo) + 1
        ReDim Preserve arrCorreo(arrNewItem)
    End If
    
End Function

Public Function Trabuque(Palabra As String, Optional Sep As String = ",") As String
'Devuelve un Array con todas las posibles palabras que contengan trabuques de una sola letra.
Dim I As Byte, arrTrab As String
arrTrab = Palabra
For I = 2 To Len(Palabra)
    arrTrab = arrTrab & Sep & Left$(Palabra, I - 2) & Mid$(Palabra, I, 1) & Mid$(Palabra, I - 1, 1) & Mid$(Palabra, I + 1)
Next
'Trabuque = Mid$(arrTrab, 2)
Trabuque = arrTrab
End Function


Public Function OrtComodines(TextoErr As String, Optional PorcentajeAlInicio As Boolean = True) As String
'Sustituye los caracteres donde hay posibles errores por comodines.
Dim I As Integer, Med As String, TxtoNew As String, Letra As String
    
    TextoErr = UCase(Trim(TextoErr))
    For I = 1 To Len(TextoErr)
        Med = Mid(TextoErr, I, 1): Letra = ""
        If Med Like "[BV]" Then
            Letra = "[BV]"
        ElseIf Med Like "[CSZ][EI]" Then
            Letra = "[CSZ]"
        ElseIf Med Like "[CK][AOU]" Then
            Letra = "[CK]"
        ElseIf Med Like "I" Then
            Letra = "[IY]"
        ElseIf Med Like "[GJ][EI]" Then
            Letra = "[GJ]"
        ElseIf Med Like "NI[AEOU]" Then
            Letra = "[NÑ]"
        ElseIf Med Like "ÑI" Then
            Letra = "[NÑ]"
        ElseIf Med Like "H" Then
            If I > 1 Then
                If Mid(TextoErr, I - 1, 1) <> "C" Then Letra = ""
            Else
                Letra = ""
            End If
        ElseIf Med Like "[OU]A" Then
            If I > 1 Then
                If Mid(TextoErr, I - 1) Like "J" Then Letra = "[OU]"
            End If
        ElseIf Med Like "Q[EI]" Then
            Letra = "Q"
        ElseIf Med Like "[SZ]" Then
            Letra = "[SZ]"
        ElseIf Med Like "Y" Then
            If I > 1 Then
                If Mid(TextoErr, I - 1) Like "[A-Z]" Then Letra = "[IY]"
            End If
        End If
        If Letra = "" Then Letra = Left(Med, 1)
        TxtoNew = TxtoNew & Letra
    Next
    
    If PorcentajeAlInicio Then OrtComodines = "%" & TxtoNew Else OrtComodines = TxtoNew
    
End Function


Public Function Autonumerico(Tabla As Integer, ByVal rdoCB1 As rdoConnection) As Long
   
    Dim Ok As Boolean, Intentos As Integer
    Intentos = 0: Ok = False
    
    Do While Not Ok And Intentos < 10
        Intentos = Intentos + 1
        Ok = PidoAutonumerico(Tabla, rdoCB1)
    Loop
    
    Autonumerico = idContador
    
End Function

Private Function PidoAutonumerico(idTabla As Integer, ByVal rdoCB1 As rdoConnection) As Boolean

    On Error GoTo errConcurr
    PidoAutonumerico = False
    Dim RsDoc As rdoResultset
    
    mSQL = "Select * from Autonumerico Where Tabla = " & idTabla
    Set RsDoc = rdoCB1.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    
    If RsDoc.EOF Then
        RsDoc.AddNew
        RsDoc!Tabla = idTabla
        RsDoc!Contador = 1
        idContador = 1
        RsDoc.Update
    Else
        idContador = RsDoc!Contador + 1
        RsDoc.Edit
        RsDoc!Contador = idContador
        RsDoc.Update
    End If
    RsDoc.Close
    
    PidoAutonumerico = True
    
    Exit Function
    
errConcurr:
    RsDoc.Close
End Function

Public Sub EjecutarApp(Path As String, Optional Valor As String = "")
On Error GoTo errApp

    Screen.MousePointer = 11
    Dim plngRet As Long
    
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
    
    Screen.MousePointer = 0
    Exit Sub
errApp:
    clsGeneral.OcurrioError "Error al ejecutar la aplicación " & Path, Err.Number & "- " & Err.Description
    Screen.MousePointer = 0
End Sub

