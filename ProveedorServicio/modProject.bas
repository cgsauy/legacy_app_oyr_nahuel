Attribute VB_Name = "modProject"
Option Explicit
Public objGral As New clsorCGSA

Public prmPlantillaProcesa As Long
Public prmSitioHome As String

Sub Main()
On Error GoTo errMain
Dim objC As New clsConexion

    If Not objC.AccesoAlMenu("ProveedorService") Then
        MsgBox "No tiene permiso de acceso a esta aplicación.", vbCritical, "Atención"
        GoTo evFin
    End If
    If Not InicioConexionBD(objC.TextoConexion("Comercio")) Then GoTo evFin
    If Val(Command()) > 0 Then frmProveedorService.prm_Proveedor = Val(Command())
    frmProveedorService.Show
    Set objC = Nothing
    Exit Sub
    
errMain:
    objGral.OcurrioError "Error al iniciar la aplicación.", Err.Description
    
evFin:
    Screen.MousePointer = 0
    End
End Sub

Private Function f_LoadPrm() As Boolean
On Error GoTo errLP
    f_LoadPrm = False
    Cons = "Select * From Parametro Where ParNombre Like 'clearing%'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "clearingidplaprocesa": prmPlantillaProcesa = RsAux!ParValor
            Case "clearingsitiobusquedamanual": prmSitioHome = Trim(RsAux!ParTexto)
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    If prmPlantillaProcesa = 0 Or prmSitioHome = "" Then
        MsgBox "Algunos de los parámetros esenciales para el funcionamiento del programa no existe , comuníueselo al administrador del sistema.", vbCritical, "Lectura de parámetros."
        Exit Function
    End If
    f_LoadPrm = True
    Exit Function
errLP:
    objGral.OcurrioError "Error al leer los parámetros.", Err.Description
End Function

Public Sub prj_SetFocus(ctrl As Control)
On Error Resume Next
    With ctrl
        If .Enabled Then
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End If
    End With
End Sub
Public Sub ObtengoSeteoForm(Formulario As Form, Optional LeftIni As Currency = 0, Optional TopIni As Currency = 0, _
                                            Optional WidthIni As Currency = 0, Optional HeightIni As Currency = 0)
    If LeftIni = 0 Then LeftIni = Formulario.Left
    If TopIni = 0 Then TopIni = Formulario.Top
    If WidthIni = 0 Then WidthIni = Formulario.Width
    If HeightIni = 0 Then HeightIni = Formulario.Height
    
    'Busco si tengo seteada la última posición y tamaño del formulario
    'Sino le marco yo los tamaños iniciales. ------------------------------------------
    Formulario.Left = GetSetting(App.Title, "Settings", "AA" & Formulario.Name & "Left", LeftIni)
    Formulario.Top = GetSetting(App.Title, "Settings", "AA" & Formulario.Name & "Top", TopIni)
    Formulario.Width = GetSetting(App.Title, "Settings", "AA" & Formulario.Name & "Width", WidthIni)
    Formulario.Height = GetSetting(App.Title, "Settings", "AA" & Formulario.Name & "Height", HeightIni)
    
End Sub

Public Sub GuardoSeteoForm(Formulario As Form)
    'Guarda la posicion y tamaño del formulario, si su estado es normal.
    If Formulario.WindowState <> vbMinimized And Formulario.WindowState <> vbMaximized Then
        SaveSetting App.Title, "Settings", "AA" & Formulario.Name & "Left", Formulario.Left
        SaveSetting App.Title, "Settings", "AA" & Formulario.Name & "Top", Formulario.Top
        SaveSetting App.Title, "Settings", "AA" & Formulario.Name & "Width", Formulario.Width
        SaveSetting App.Title, "Settings", "AA" & Formulario.Name & "Height", Formulario.Height
    End If
End Sub


