Attribute VB_Name = "modStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion

Public paLocalPuerto As Long, paLocalZF As Long

Public Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miConexion.AccesoAlMenu(App.Title) Then
            InicioConexionBD miConexion.TextoConexion("comercio")
            paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
            
            CargoParametrosSucursal
            'CargoParametrosImportaciones
            CargoParametrosCaja
            CargoParametrosComercio
            
            CargoBasesDeDatos
            
            frmListado.Show vbModeless
            
        Else
            If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contrase�a.
        miConexion.AccesoAlMenu (App.Title)
        InicioConexionBD miConexion.TextoConexion("comercio")
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub


Public Function AccionCambiarBase(mKey As String, mNombre As String) As Boolean
    
    On Error GoTo errCh
    AccionCambiarBase = False
    Dim miConnect As String: miConnect = ""
    miConnect = miConexion.TextoConexion(mKey)
    
    If Trim(miConnect) = "" Then
        MsgBox "No hay una conexi�n disponible para �sta base de datos." & vbCrLf & _
                    "Consulte con el administrador de bases de datos.", vbExclamation, "Falta Conexi�n a " & mNombre
        Screen.MousePointer = 0: Exit Function
    End If
    
    If MsgBox("Cambiar de base a " & mNombre & vbCrLf & _
                   "Confirma cambiar la base de datos. ?", vbQuestion + vbYesNo + vbDefaultButton2, "Realmente desea cambiar la base") = vbNo Then Exit Function
    
    Screen.MousePointer = 11
    
    On Error Resume Next
    cBase.Close
    On Error GoTo errCh
    Set cBase = Nothing
    Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , miConnect)
    
    MsgBox "Ahora est� trabajando en la base de datos " & mNombre & vbCrLf & _
                "Presione aceptar para actualizar la informaci�n.", vbExclamation, "Base Cambiada OK"
    
    Screen.MousePointer = 0
    AccionCambiarBase = True
    Exit Function
    
errCh:
    clsGeneral.OcurrioError "Error de conexi�n al cambiar la base de datos.", Err.Description
    Screen.MousePointer = 0
End Function


Private Sub CargoBasesDeDatos()
On Error GoTo errCBase

    frmListado.MnuBases.Enabled = False
    If Not miConexion.AccesoAlMenu("Cambiar_Conexion") Then Exit Sub
    
    Dim aItem As Integer, I As Integer
    I = 0
    cons = "Select * from Bases Order by BasNombre"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        Do While Not rsAux.EOF

            If Not IsNull(rsAux!BasBColor) Then
                frmListado.MnuBases.Tag = frmListado.MnuBases.Tag & rsAux!BasBColor & "|"
            Else
                frmListado.MnuBases.Tag = frmListado.MnuBases.Tag & "|"
            End If

            If I > 0 Then Load frmListado.MnuBx(I)
            
            With frmListado.MnuBx(I)
                .Tag = Trim(rsAux!BasConexion)
                .Caption = Trim(rsAux!BasNombre)
                .Visible = True
            End With
            
            I = I + 1
            rsAux.MoveNext
            
        Loop
        frmListado.MnuBases.Enabled = True
    End If
    rsAux.Close
    Exit Sub

errCBase:
End Sub


