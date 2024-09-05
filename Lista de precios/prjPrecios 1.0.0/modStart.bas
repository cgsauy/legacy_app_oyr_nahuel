Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmPathListados As String
Public paBD As String

Public paTipoCuotaContado As Long
Public paMonedaPesos As Long
Public paCuotaMin As Currency

' <VB WATCH>
Const VBWMODULE = "modStart"
' </VB WATCH>

Public Sub Main()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          vbwInitializeProtector ' Initialize VB Watch
' </VB WATCH>

3          On Error GoTo errMain
4          Screen.MousePointer = 11
5          Dim aTexto As String

6          If Not miConexion.AccesoAlMenu("Listas de Precios") Then
7              MsgBox "Acceso denegado. " & vbCrLf & "Consulte a su administrador de Sistemas", vbExclamation, "Acceso Denegado"
8              End
9          End If

10         paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)

11         If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then
12              End
13         End If

14         paBD = miConexion.RetornoPropiedad(bDB:=True)

15         CargoParametrosLocal
           'prmPathListados = "C:\Proyectos\Precios\Reportes\"
16         frmListas.Show vbModeless

17         Exit Sub

18     errMain:
19         On Error Resume Next
20         Screen.MousePointer = 0
21         MsgBox "Error al inicializar la aplicación " & App.title & Chr(13) & "Error: " & Trim(Err.Description)
22         End
' <VB WATCH>
23         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Main"

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

Private Sub CargoParametrosLocal()
' <VB WATCH>
24         On Error GoTo vbwErrHandler
' </VB WATCH>
25     On Error Resume Next
26         prmPathListados = ""

27         Cons = "Select * from Parametro Where ParNombre In ('pathapp', 'TipoCuotaContado', 'MonedaPesos', 'webminimportecuota')"
28         Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
29         Do While Not RsAux.EOF
30             Select Case LCase(Trim(RsAux!ParNombre))
                   Case "webminimportecuota"
31                 paCuotaMin = RsAux("ParValor")
32                 Case "pathapp"
33                 prmPathListados = Trim(RsAux!ParTexto)
34                 Case "tipocuotacontado"
35                 paTipoCuotaContado = RsAux!ParValor
36                 Case "monedapesos"
37                 paMonedaPesos = RsAux!ParValor
38             End Select

39             RsAux.MoveNext
40         Loop
41         RsAux.Close

42         Cons = ""
43         Dim aPos As Integer, aT2 As String
44         aT2 = prmPathListados
45         Do While InStr(aT2, "\") <> 0
46             aPos = InStr(aT2, "\")
47             Cons = Cons & Mid(aT2, 1, aPos)
48             aT2 = Mid(aT2, aPos + 1)
49         Loop
50         prmPathListados = Cons & "Reportes\"


           'paCodigoDeSucursal
       '    cons = miConexion.NombreTerminal
       '    cons = "Select * from Terminal Where TerNombre = '" & cons & "'"
       '    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
       '    If Not rsAux.EOF Then If Not IsNull(rsAux!TerSucursal) Then paCodigoDeSucursal = rsAux!TerSucursal
       '    rsAux.Close

' <VB WATCH>
51         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CargoParametrosLocal"

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

