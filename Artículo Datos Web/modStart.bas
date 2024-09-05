Attribute VB_Name = "modStart"
Option Explicit

Public gFechaServidor As String

Public idArticulo As Long
Public pathFotos As String, pathWeb As String, pathIntra As String
Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Sub Main()
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If miConexion.AccesoAlMenu(App.Title) Then
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            Screen.MousePointer = 0
            End: Exit Sub
        End If
        
        Cons = "Select * From Parametro Where ParNombre Like 'ArticuloPath%'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            Select Case LCase(Trim(RsAux!ParNombre))
                Case "articulopathfotos": pathFotos = Trim(RsAux!ParTexto)
                Case "articulopathpageweb": pathWeb = Trim(RsAux!ParTexto)
                Case "articulopathpageintra": pathIntra = Trim(RsAux!ParTexto)
            End Select
            RsAux.MoveNext
        Loop
        RsAux.Close
        If Trim(pathFotos) <> "" Then
            If Right(pathFotos, 1) <> "\" Then pathFotos = pathFotos & "\"
        Else
            pathFotos = App.Path & "\"
        End If
        If Trim(Command()) <> "" Then idArticulo = Val(Command())
        frmDatosWeb.Show
    Else
        MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurri� un error al inicializar la aplicaci�n " & App.Title & Chr(13) & "Error: " & Trim(Err.Description), vbCritical, "ATENCI�N"
    End
End Sub
Public Sub FechaDelServidor()
Dim RsF As rdoResultset
    
    On Error GoTo errFecha
    Cons = "Select GetDate()"
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    gFechaServidor = RsF(0)
    RsF.Close
    
    Time = gFechaServidor
    Date = gFechaServidor
    Exit Sub

errFecha:
    gFechaServidor = Date
End Sub

