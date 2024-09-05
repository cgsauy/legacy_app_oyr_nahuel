Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmIDMail As Long
Public prmIDCliente As Long
Public prmIDArticulo As Long

Public mSQL As String

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    Dim bAccesoOK As Boolean
    bAccesoOK = False
    
    'If Not bAccesoOK Then bAccesoOK = miConexion.AccesoAlMenu("ContadorComun")
    bAccesoOK = True
    If bAccesoOK Then
    
        If Not InicioConexionBD(miConexion.TextoConexion("comercio"), 45) Then End
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        CargoParametrosEntrada Trim(Command())
        If prmIDMail > 0 Then
            frmAviso.Show vbModeless
        Else
            EndMain
        End If
        
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


Public Function EndMain()
    On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Function

Private Function CargoParametrosEntrada(mPrms As String)
    On Error GoTo errCPE
    
    prmIDMail = 0: prmIDCliente = 0
    If Trim(mPrms) = "" Then Exit Function
    
    Dim I As Integer
    Dim arrPrms() As String, arrVals() As String
    arrPrms = Split(mPrms, "|")
    'M=XX    id de mail
    'C=XX     id de Cliente

    For I = LBound(arrPrms) To UBound(arrPrms)
        arrVals = Split(arrPrms(I), "=")
        Select Case UCase(arrVals(0))
            Case "M": prmIDMail = arrVals(1)
            Case "C": prmIDCliente = arrVals(1)
        End Select
    Next
    
    If prmIDMail = 0 And prmIDCliente <> 0 Then
                        
        mSQL = "Select EMDCodigo, EMDDireccion as Nombre, (RTrim(EMDDireccion) + '@' + RTrim(EMSDireccion))  as Direccion " & _
                    " From CGSA.dbo.EMailDireccion, CGSA.dbo.EMailServer" & _
                    " Where EMDIDCliente = " & prmIDCliente & " And EMDServidor = EMSCodigo"

        Dim objLista As New clsListadeAyuda

        If objLista.ActivarAyuda(cBase, mSQL, , 1, "Direcciones de correo") > 0 Then
            prmIDMail = objLista.RetornoDatoSeleccionado(0)
        End If
        Set objLista = Nothing
        
    End If
    
    Exit Function
    
errCPE:
    clsGeneral.OcurrioError "Error al cargar los parámetros de entrada: " & mPrms, Err.Description
End Function


Public Function ValidoAccesoMnu(Menu As String, idUsr As Long) As Boolean

    On Error Resume Next
    ValidoAccesoMnu = False
    
    If idUsr = 0 Then Exit Function
    If Menu = "" Then Exit Function
    
    cons = " Select * from logdb.dbo.NivelPermiso, logdb.dbo.Aplicacion " _
            & " Where NPeNivel IN (Select UNiNivel from UsuarioNivel Where UNiUsuario = " & idUsr & ")" _
            & " And NPeAplicacion = AplCodigo" _
            & " And AplNombre = '" & Trim(Menu) & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then ValidoAccesoMnu = True
    rsAux.Close

End Function


