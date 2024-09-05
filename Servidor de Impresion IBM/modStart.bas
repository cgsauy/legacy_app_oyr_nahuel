Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public prmNombreSucursal As String
Public prmTickeadorasAsignadas As String
Public prmTickeadoraCuotasGiros As Integer


Public Sub Main()
     
    If App.PrevInstance Then End: Exit Sub
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    Dim aTexto As String
    
    If Not miConexion.AccesoAlMenu("Reimprimir Documentos") Then
        MsgBox "Acceso denegado. " & vbCrLf & "Consulte a su administrador de Sistemas", vbExclamation, "Acceso Denegado"
        End
    End If
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then End
    
    CargoParametrosLocal
    CargarImpresorasAsignadas
   
    Load frmInicio
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Public Sub CargarImpresorasAsignadas()
    
    Cons = "SELECT ParNombre, ParValor, ParTexto from Parametro Where ParNombre In ('TickeadoraAsignadaTerminal', 'TickeadoraGirosCuotas')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux("ParNombre")))
            Case LCase("TickeadoraAsignadaTerminal")
                If Not IsNull(RsAux("ParTexto")) Then prmTickeadorasAsignadas = Trim(RsAux("ParTexto"))
            Case LCase("TickeadoraGirosCuotas")
                If Not IsNull(RsAux("ParValor")) Then prmTickeadoraCuotasGiros = RsAux("ParValor")
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

Private Sub CargoParametrosLocal()
On Error Resume Next
    
    Cons = miConexion.NombreTerminal
    Cons = "Select * from Terminal Left Outer Join Sucursal On TerSucursal = SucCodigo Where TerNombre = '" & Cons & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
    
        If Not IsNull(RsAux!TerSucursal) Then
            
            paCodigoDeSucursal = RsAux!TerSucursal
            prmNombreSucursal = Trim(RsAux!SucAbreviacion)
            paCodigoDeTerminal = RsAux!TerCodigo
            
        End If
        
    End If
    RsAux.Close
    
End Sub

