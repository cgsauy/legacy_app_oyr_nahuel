Attribute VB_Name = "modStart"

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Sub Main()
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion(logImportaciones)
        
        CargoBasesDeDatos
        
        CargoParametrosImportaciones
                
        frmMaDisponibilidad.Show
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub

errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & _
                Err.Number & " - " & Err.Description
    End
End Sub

Public Function AccionCambiarBase(mKey As String, mNombre As String) As Boolean
    
    On Error GoTo errCh
    AccionCambiarBase = False
    Dim miConnect As String: miConnect = ""
    miConnect = miConexion.TextoConexion(mKey)
    
    If Trim(miConnect) = "" Then
        MsgBox "No hay una conexión disponible para ésta base de datos." & vbCrLf & _
                    "Consulte con el administrador de bases de datos.", vbExclamation, "Falta Conexión a " & mNombre
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
    
    MsgBox "Ahora está trabajando en la base de datos " & mNombre & vbCrLf & _
                "Presione aceptar para actualizar la información.", vbExclamation, "Base Cambiada OK"
    
    Screen.MousePointer = 0
    AccionCambiarBase = True
    Exit Function
    
errCh:
    clsGeneral.OcurrioError "Error de conexión al cambiar la base de datos.", Err.Description
    Screen.MousePointer = 0
End Function


Private Sub CargoBasesDeDatos()
On Error GoTo errCBase

    frmMaDisponibilidad.MnuBases.Enabled = False
    If Not miConexion.AccesoAlMenu("Cambiar_Conexion") Then Exit Sub
    
    Dim aItem As Integer, I As Integer
    I = 0
    Cons = "Select * from Bases Order by BasNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF

            If Not IsNull(RsAux!BasBColor) Then
                frmMaDisponibilidad.MnuBases.Tag = frmMaDisponibilidad.MnuBases.Tag & RsAux!BasBColor & "|"
            Else
                frmMaDisponibilidad.MnuBases.Tag = frmMaDisponibilidad.MnuBases.Tag & "|"
            End If

            If I > 0 Then Load frmMaDisponibilidad.MnuBx(I)
            
            With frmMaDisponibilidad.MnuBx(I)
                .Tag = Trim(RsAux!BasConexion)
                .Caption = Trim(RsAux!BasNombre)
                .Visible = True
            End With
            
            I = I + 1
            RsAux.MoveNext
            
        Loop
        frmMaDisponibilidad.MnuBases.Enabled = True
    End If
    RsAux.Close
    Exit Sub

errCBase:
End Sub


