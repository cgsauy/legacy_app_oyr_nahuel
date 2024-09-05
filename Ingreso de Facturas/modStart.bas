Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public txtConexion As String

Public prmKeyConnect As String
Public Const prmKeyApp = "Ingreso de Facturas"
Public Const prmKeyAppADM = "GastosADM"

Public Type typSuceso
    Tipo As Integer
    Titulo As String
    Defensa As String
    Usuario As Long
    Autoriza As Long
    Valor As Currency
    Cliente As Long
End Type

Public dSuceso As typSuceso
Public prmSucesoGastos As Integer
Public prmSucesoModGastos As Integer

Public prmPathApp As String

Private Type typRub        'Definicion para Rubros y Subrubros
    IdRubro As Long
    TextoRubro As String
    IdSRubro As Long
    TextoSRubro As String
    Importe As Currency
End Type
Public arrRubros() As typRub


Public Sub Main()
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If miConexion.AccesoAlMenu(prmKeyApp) Then
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        prmKeyConnect = "comercio"
        txtConexion = miConexion.TextoConexion(prmKeyConnect)
        InicioConexionBD txtConexion
        
        prmSucesoGastos = 19
        prmSucesoModGastos = 20
        
        CargoParametrosImportaciones
        CargoParametrosComercio
        CargoParametrosSucursal
        
        CargoParametrosLocales
        CargoBasesDeDatos
        
        frmFacturas.Show
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub

errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & prmKeyApp & Chr(vbKeyReturn) & _
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
    
    prmKeyConnect = mKey
    
    MsgBox "Ahora está trabajando en la base de datos " & mNombre & vbCrLf & _
                "Presione aceptar para actualizar la información.", vbExclamation, "Base Cambiada OK"
    
    txtConexion = miConnect
    Screen.MousePointer = 0
    AccionCambiarBase = True
    Exit Function
    
errCh:
    clsGeneral.OcurrioError "Error de conexión al cambiar la base de datos.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub CargoBasesDeDatos()
On Error GoTo errCBase

    frmFacturas.MnuBases.Enabled = False
    If Not miConexion.AccesoAlMenu("Cambiar_Conexion") Then Exit Sub
    
    Dim aItem As Integer, I As Integer
    I = 0
    cons = "Select * from Bases Order by BasNombre"
    Set RsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF

            If Not IsNull(RsAux!BasBColor) Then
                frmFacturas.MnuBases.Tag = frmFacturas.MnuBases.Tag & RsAux!BasBColor & "|"
            Else
                frmFacturas.MnuBases.Tag = frmFacturas.MnuBases.Tag & "|"
            End If

            If I > 0 Then Load frmFacturas.MnuBx(I)
            
            With frmFacturas.MnuBx(I)
                .Tag = Trim(RsAux!BasConexion)
                .Caption = Trim(RsAux!BasNombre)
                .Visible = True
            End With
            
            I = I + 1
            RsAux.MoveNext
            
        Loop
        frmFacturas.MnuBases.Enabled = True
    End If
    RsAux.Close
    Exit Sub

errCBase:
End Sub

Private Sub CargoParametrosLocales()
On Error GoTo errCP
    
    prmFCierreIVA = CDate("1/1/2002")
    
    cons = "Select * from Parametro " & _
            " Where ParNombre IN ( 'pathApp', 'FechaCierreIVA')"
            
    Set RsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case (Trim(LCase(RsAux!ParNombre)))
            
            Case "pathapp": prmPathApp = Trim(RsAux!ParTexto) & "\"
            
            Case "fechacierreiva": prmFCierreIVA = CDate(RsAux!ParTexto)
            
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Sub
errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros (locales).", Err.Description
End Sub


Public Function ing_BuscoSubrubro(mControlR As TextBox, mControlSR As TextBox) As Boolean
On Error GoTo errBS

    ing_BuscoSubrubro = False
    Dim aQ As Integer, aID As Long, aTexto As String
    aQ = 0: aID = 0
    
    mControlSR.Text = Replace(RTrim(mControlSR.Text), " ", "%")
    
    cons = "Select SRuID, SRuNombre as 'SubRubro', SRuCodigo as 'Cód. SR', RubNombre as 'Rubro', RubCodigo as 'Cód. Rubro'" _
            & " from SubRubro, Rubro " _
            & " Where SRuNombre like '" & Trim(mControlSR.Text) & "%'" _
            & " And SRuCodigo Not like '" & paRubroDisponibilidad & "%'" _
            & " And SRuRubro = RubID "
    If Val(mControlR.Tag) <> 0 Then cons = cons & " And RubID= " & Val(mControlR.Tag)
    cons = cons & " Order by SRuNombre"
                
    Set RsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        aQ = 1: aID = RsAux!SRuID: aTexto = Trim(RsAux(1))
        RsAux.MoveNext
        If Not RsAux.EOF Then
            aQ = 2: aID = 0
        End If
    End If
    RsAux.Close
    
    Select Case aQ
        Case 0: MsgBox "No existen Subrubros para el texto ingresado.", vbExclamation, "No hay datos"
        
        Case 1:
                mControlSR.Text = aTexto: mControlSR.Tag = aID
        
        Case 2:
                Dim aLista As New clsListadeAyuda
                aID = aLista.ActivarAyuda(cBase, cons, 5500, 1, "Sub Rubros")
                
                If aID <> 0 Then
                    aTexto = Trim(aLista.RetornoDatoSeleccionado(1))
                    aID = aLista.RetornoDatoSeleccionado(0)
                End If
                
                Set aLista = Nothing
    End Select
    
    If aID <> 0 Then
        cons = "Select * from Subrubro, Rubro " & _
                   " Where SRuID = " & aID & _
                   " And SRuRubro = RubID"
        
        Set RsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            mControlR.Text = Trim(RsAux!RubNombre)
            mControlR.Tag = RsAux!SRuRubro
            
            mControlSR.Text = Trim(RsAux!SRuNombre)
            mControlSR.Tag = RsAux!SRuID
            ing_BuscoSubrubro = True
        End If
        RsAux.Close
        
    End If
    
    Screen.MousePointer = 0
    Exit Function

errBS:
    clsGeneral.OcurrioError "Error al buscar el subrubro.", Err.Description
    Screen.MousePointer = 0
End Function


Public Function ing_BuscoRubro(mControlR As TextBox) As Boolean
On Error GoTo errBS

    ing_BuscoRubro = False
    Dim aQ As Integer, aID As Long, aTexto As String
    aQ = 0: aID = 0
    
    mControlR.Text = Replace(RTrim(mControlR.Text), " ", "%")
    
    cons = "Select RubID, RubNombre as 'Rubro', RubCodigo as 'Cód. Rubro'" _
            & " from Rubro " _
            & " Where RubNombre like '" & Trim(mControlR.Text) & "%'" _
            & " Order by RubNombre"
                
    Set RsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        aQ = 1: aID = RsAux!RubID: aTexto = Trim(RsAux(1))
        RsAux.MoveNext
        If Not RsAux.EOF Then
            aQ = 2: aID = 0
        End If
    End If
    RsAux.Close
    
    Select Case aQ
        Case 0: MsgBox "No existen rubros para el texto ingresado.", vbExclamation, "No hay datos"
        
        Case 1:
                mControlR.Text = aTexto: mControlR.Tag = aID
        
        Case 2:
                Dim aLista As New clsListadeAyuda
                aID = aLista.ActivarAyuda(cBase, cons, 4500, 1, "Rubros")
                If aID <> 0 Then
                    aTexto = Trim(aLista.RetornoDatoSeleccionado(1))
                    aID = aLista.RetornoDatoSeleccionado(0)
                End If
                Set aLista = Nothing
    End Select
    
    If aID <> 0 Then
        mControlR.Text = aTexto
        mControlR.Tag = aID
            
        cons = "Select Top 2 * from Subrubro" & _
                   " Where SRuRubro = " & aID & _
                   " And SRuCodigo Not like '" & paRubroDisponibilidad & "%'"
        Set RsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aTexto = Trim(RsAux!SRuNombre)
            aID = RsAux!SRuID
            RsAux.MoveNext
            If RsAux.EOF Then
                mControlR.Text = aTexto
                mControlR.Tag = aID
            End If
        End If
        RsAux.Close
        
    End If
    
    Screen.MousePointer = 0
    Exit Function

errBS:
    clsGeneral.OcurrioError "Error al buscar el rubro.", Err.Description
    Screen.MousePointer = 0
End Function


