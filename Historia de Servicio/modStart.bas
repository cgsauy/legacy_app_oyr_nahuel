Attribute VB_Name = "modStart"
Option Explicit

Public Enum TipoCliente
    Cliente = 1
    Empresa = 2
End Enum

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public txtConexion As String
'---------------------------------------------------------------------------------------

Public Sub Main()
Dim aSucursal As String
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion("comercio")
        InicioConexionBD txtConexion
        CargoParametrosSucursal
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        frmHisServicio.Show vbModeless
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    End
End Sub

Public Function CargoParametrosSucursal() As String

Dim aNombreTerminal As String

    CargoParametrosSucursal = ""
    aNombreTerminal = miConexion.NombreTerminal
    
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & aNombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
'        If Not IsNull(RsAux!SucDisponibilidad) Then paDisponibilidad = RsAux!SucDisponibilidad Else paDisponibilidad = 0
        CargoParametrosSucursal = Trim(RsAux!SucAbreviacion)
    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(aNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        End
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    
End Function

Public Function f_GetEventos(ByVal sAux As String) As String
On Error Resume Next
    f_GetEventos = ""
    If InStr(1, sAux, "[", vbTextCompare) = 1 And InStr(1, sAux, "/", vbTextCompare) > 1 And InStr(1, sAux, ":", vbTextCompare) > 2 And InStr(1, sAux, "]", vbTextCompare) > 1 Then
        f_GetEventos = Mid(sAux, InStr(1, sAux, "[", vbTextCompare), InStr(InStr(1, sAux, "[", vbTextCompare) + 1, sAux, "]"))
    End If
End Function

Public Function f_QuitarClavesDelComentario(ByVal sComentario As String) As String
Dim sAux As String
    sAux = f_GetEventos(sComentario)
    If sAux <> "" Then
        f_QuitarClavesDelComentario = Replace(sComentario, sAux, "")
    Else
        f_QuitarClavesDelComentario = sComentario
    End If
End Function


