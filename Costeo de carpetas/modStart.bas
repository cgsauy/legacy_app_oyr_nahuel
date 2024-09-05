Attribute VB_Name = "modStart"
Option Explicit

'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno
Public RsAux As rdoResultset         'Resultset Auxiliar

'String.
Public Cons As String
Public paCodigoDeUsuario As Long

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
'Public txtConeccion As String

Public Sub Main()
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        'txtConeccion = miConexion.TextoConexion(logImportaciones)
        'InicioConexionBD txtConeccion
        
        If Not ObtenerConexionBD(cBase, logImportaciones) Then Screen.MousePointer = 0: End
        CargoParametrosImportaciones
        frmCosteo.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmCosteo.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmCosteo.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        CargoParametrosParaZureo
        
        If Not fnc_ConnectZureo Then
            MsgBox "Sin acceso a Zureo, no podrá continuar.", vbCritical, "ATENCIÓN"
            End
        End If
        frmCosteo.Show
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub

errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & _
                Err.Number & " - " & Err.Description
    End
End Sub

Public Function PropiedadesConnect(Conexion As String, _
                                                    Optional Database As Boolean = True, Optional DSN As Boolean = False, _
                                                    Optional Server As Boolean = True) As String
Dim aRetorno As String

    On Error GoTo errConnect
    PropiedadesConnect = ""
    Conexion = UCase(Conexion)
    If DSN Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "DSN=") + 4, Len(Conexion)))
    If Server Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "SERVER=") + 7, Len(Conexion)))
    If Database Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "DATABASE=") + 9, Len(Conexion)))
    
    aRetorno = Trim(Mid(aRetorno, 1, InStr(aRetorno, ";") - 1))
    
    PropiedadesConnect = aRetorno
    Exit Function
    
errConnect:
End Function

