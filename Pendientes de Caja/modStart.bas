Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmPathApp As String
Public paDisponibilidad As Long
Public prmPlantillas As String

'Parámetros de Entrada
Dim prmE_ID As Long, prmE_IDDocumento As Long
    
Public Sub Main()
     
    On Error GoTo errMain
    
    Screen.MousePointer = 11
    
    Dim bAccesoOK As Boolean, mAppVer As String
    
    bAccesoOK = False: mAppVer = "Pendientes de Caja"
    bAccesoOK = miConexion.AccesoAlMenu(mAppVer)
    
    If bAccesoOK Then
        
        If mAppVer <> "" And mAppVer <> App.Major & "." & App.Minor & "." & App.Revision Then
            MsgBox "La versión del programa no es la última disponible." & vbCr & vbCr & _
                        "Su versión es la " & App.Major & "." & App.Minor & "." & App.Revision & vbCr & _
                        "Ud. debe actualizar el software a la versión " & mAppVer, vbExclamation, "Actualizar a Versión " & mAppVer
            EndMain
        End If
        
        If Not InicioConexionBD(miConexion.TextoConexion("comercio"), 45) Then End
    
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        CargoParametrosLocales
        CargoParametrosEntrada Trim(Command())
        
        frmMain.Show vbModeless
                
        frmMain.gbl_CargaConParametros prmE_ID, prmE_IDDocumento
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then
            MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "Usuario sin Autorización"
        End If
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocales()
On Error GoTo errCP

    cons = "Select * from Parametro " & _
            " Where ParNombre IN ( 'PathApp', 'PlantillasPendienteCaja')"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto)
            
            Case "plantillaspendientecaja": prmPlantillas = Trim(rsAux!ParTexto)
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & miConexion.NombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        paCodigoDeSucursal = rsAux!TerSucursal
        paCodigoDeTerminal = rsAux!TerCodigo
        If Not IsNull(rsAux!SucDisponibilidad) Then paDisponibilidad = rsAux!SucDisponibilidad Else paDisponibilidad = 0
    End If
    rsAux.Close
    Exit Sub

errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub


Public Function EndMain()
    On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Function

Private Function CargoParametrosEntrada(mPrms As String)
    On Error GoTo errCPE
    'ID IDPte |D IDDocumento
    prmE_IDDocumento = 0: prmE_ID = 0
    
    If Trim(mPrms) = "" Then Exit Function
    
    Dim I As Integer
    Dim arrPrms() As String, arrValues() As String
    arrPrms = Split(Trim(mPrms), "|")
    
    For I = LBound(arrPrms) To UBound(arrPrms)
        arrValues = Split(arrPrms(I), " ")
        Select Case UCase(arrValues(0))
            
            Case "ID"
                prmE_ID = Val(arrValues(1))
                        
            Case "D": prmE_IDDocumento = Val(arrValues(1))
            
        End Select
        
    Next
    
    Exit Function
    
errCPE:
    clsGeneral.OcurrioError "Error al cargar los parámetros de entrada: " & mPrms, Err.Description
End Function

