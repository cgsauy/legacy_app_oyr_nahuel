Attribute VB_Name = "modStart"
Option Explicit
Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral

Public paEstadoArticuloARecuperar As Long
Public paEstadoArticuloRoto As Long
Public paGrupoRepuesto As Long
Public paTipoArticuloServicio As Long

Public paPlBalance As String

Public paLocalCompania As Long
Public paLocalEduardo As Long

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    If Not miConexion.AccesoAlMenu(App.Title) Then End
    
    If Not InicioConexionBD(miConexion.TextoConexion("balance")) Then End
    
    CargoParametrosLocal
    
    frmListado.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    frmListado.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    frmListado.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
            
    frmListado.Show vbModeless
    
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocal()

    cons = "Select * from cgsa.dbo.Parametro"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            Case "estadoarecuperar": paEstadoArticuloARecuperar = rsAux!ParValor
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = rsAux!ParValor
            Case "estadoroto": paEstadoArticuloRoto = rsAux!ParValor
            
            Case "repuesto": paGrupoRepuesto = rsAux!ParValor
            
            Case "tipoarticuloservicio": paTipoArticuloServicio = rsAux!ParValor
            
            Case "plbalance": paPlBalance = Trim(rsAux!ParTexto)
            
            Case "localcompañia": paLocalCompania = rsAux!ParValor
        End Select
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    cons = "Select * from cgsa.dbo.Local Where LocNombre = 'Eduardo'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then paLocalEduardo = rsAux!LocCodigo
    rsAux.Close
    
End Sub
