Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Sub Main()

    On Error GoTo errMain
    Dim sParams() As String, sConexion As String
    
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
    
        If Trim(Command()) <> "" Then
            sParams = Split(Trim(Command()), "|")
            '0- id Gasto    1-conexion a bd
            frmPagos.prmIDGasto = sParams(0)
            If UBound(sParams) >= 1 Then sConexion = Trim(sParams(1))
        End If
        
        If Trim(sConexion) <> "" Then
            InicioConexionBD miConexion.TextoConexion(sConexion)
        Else
            InicioConexionBD miConexion.TextoConexion("comercio")
        End If
        
        CargoBasesDeDatos
        If Trim(sConexion) <> "" Then frmPagos.prmColorBase (sConexion)
        frmPagos.Show
        
    
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

    frmPagos.MnuBases.Enabled = False
    If Not miConexion.AccesoAlMenu("Cambiar_Conexion") Then Exit Sub
    
    Dim aItem As Integer, I As Integer
    I = 0
    cons = "Select * from Bases Order by BasNombre"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        Do While Not rsAux.EOF

            If Not IsNull(rsAux!BasBColor) Then
                frmPagos.MnuBases.Tag = frmPagos.MnuBases.Tag & rsAux!BasBColor & "|"
            Else
                frmPagos.MnuBases.Tag = frmPagos.MnuBases.Tag & "|"
            End If

            If I > 0 Then Load frmPagos.MnuBx(I)
            
            With frmPagos.MnuBx(I)
                .Tag = Trim(rsAux!BasConexion)
                .Caption = Trim(rsAux!BasNombre)
                .Visible = True
            End With
            
            I = I + 1
            rsAux.MoveNext
            
        Loop
        frmPagos.MnuBases.Enabled = True
    End If
    rsAux.Close
    Exit Sub

errCBase:
End Sub


