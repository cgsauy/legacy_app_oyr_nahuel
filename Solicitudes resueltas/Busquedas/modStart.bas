Attribute VB_Name = "modStart"
Option Explicit

Public Const FormatoCedula = "_.___.___-_"

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmPathApp As String

Public paTipoCuotaContado As Long

'Definicion de Tipos de Clientes------------------------------------------------------------------------------------
Public Enum TipoCliente
    Cliente = 1
    Empresa = 2
End Enum
'-----------------------------------------------------------------------------------------------------------------------

Public Enum EstadoSolicitud
    Pendiente = 0
    Aprovada = 1
    Rechazada = 2
    Condicional = 3
    ParaRetomar = 4
End Enum

'Definicion de Tipos Resolucion de Solicitud
Public Enum TipoResolucionSolicitud
    Automatica = 1
    Manual = 2
    Facturada = 3
    Facturando = 4
    LlamarA = 5
End Enum

Public Sub Main()

    On Error GoTo errMain
    
    If Not miConexion.AccesoAlMenu("Buscar Solicitudes") Then
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then
            MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "Acceso Denegado"
        End If
        End
    End If
    
    Screen.MousePointer = 11

    Dim txtConexion As String
    txtConexion = miConexion.TextoConexion("comercio")
    InicioConexionBD txtConexion
        
    CargoParametrosLocal
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
    frmBuscar.Show
        
    Screen.MousePointer = 0
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & vbCrLf & "Error: " & Trim(Err.Description)
    End
End Sub

Public Function EndMain()

    On Error Resume Next
    
    CierroConexion
    
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Function

Private Sub CargoParametrosLocal()

    prmPathApp = App.Path
    paTipoCuotaContado = 0
    
    cons = "Select * from Parametro Where ParNombre IN ( 'tipocuotacontado')"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            
            Case "tipocuotacontado": paTipoCuotaContado = rsAux!ParValor
            
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
End Sub

