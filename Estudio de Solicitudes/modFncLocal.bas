Attribute VB_Name = "modFncLocal"
Option Explicit

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Variables para los sonidos---------------------------------------------------
Dim wavResolucionN1 As String, wavNivel1C As Integer
Dim wavResolucionN2 As String, wavNivel2C As Integer
'----------------------------------------------------------------------------------

Dim colUsuarios As New Collection   'Usuarios
Public colImportes As New Collection    'Importes de solicitudes

Public Sub fnc_CargoParametrosSonido()

Dim intFile As Integer, varTmp As String, strFile As String
Dim aAux As String

    On Error GoTo PROC_ERR
    intFile = FreeFile
    strFile = "C:\AA Aplicaciones\System\Parametros.ini"
    Open strFile For Input As intFile
    
    Do Until EOF(intFile)
        Input #intFile, varTmp
        If InStr(varTmp, "=") <> 0 Then
            aAux = LCase(Mid(varTmp, 1, InStr(varTmp, "=") - 1))
            Select Case aAux
                Case "wavresolucionn1"
                      wavResolucionN1 = Mid(varTmp, InStr(varTmp, "=") + 1, InStr(varTmp, "|") - InStr(varTmp, "=") - 1)
                      wavNivel1C = Trim(Mid(varTmp, InStr(varTmp, "|") + 1, Len(varTmp)))
              
                  Case "wavresolucionn2"
                      wavResolucionN2 = Mid(varTmp, InStr(varTmp, "=") + 1, InStr(varTmp, "|") - InStr(varTmp, "=") - 1)
                      wavNivel2C = Trim(Mid(varTmp, InStr(varTmp, "|") + 1, Len(varTmp)))
            End Select
        End If
    Loop
    
    Close #intFile
    Exit Sub
  
PROC_ERR:
  MsgBox "Error al cargar los parámetros de sonido." & vbCrLf & _
               Err.Number & "- " & Err.Description, vbCritical, "Archivo: " & strFile
  Close #intFile
End Sub

Public Sub fnc_ActivoSonido(Cantidad As Integer)
    
    Dim Result As Long, aFile As String
    On Error Resume Next
    If Trim(wavResolucionN1) = "" And Trim(wavResolucionN2) = "" Then Exit Sub
    
    Select Case Cantidad
        Case Is >= wavNivel2C: If Trim(wavResolucionN2) <> "" Then aFile = wavResolucionN2 Else aFile = wavResolucionN1
        Case Is >= wavNivel1C: If Trim(wavResolucionN1) <> "" Then aFile = wavResolucionN1 Else aFile = wavResolucionN2
    End Select
    Result = sndPlaySound(aFile, 1)
    
End Sub

Public Function fnc_IconName(QSolicitud As Integer, QServicio As Integer) As String

    On Error GoTo errIcono
    
    fnc_IconName = "s0"
    If QSolicitud = 0 And QServicio = 0 Then Exit Function
    
    Dim aIcon As String
    If QServicio > 0 Then aIcon = "e" Else aIcon = "s"
    
    Select Case QSolicitud
        Case Is > 9: aIcon = aIcon & 9
        Case Else: aIcon = aIcon & QSolicitud
     End Select
     
     fnc_IconName = aIcon
     
errIcono:
End Function


'---------------------------------------------------------------------------------------------------------------
'   Valores que Retorna:    -1: Error o No Existe
'                                       0: Hay Otro Usuario
'                                       1: Bloqueada
Public Function fnc_BlockearSolicitud(Codigo As Long, Optional retUsuarioR As Long = 0) As Integer

    fnc_BlockearSolicitud = 0
    Screen.MousePointer = 11
    On Error GoTo errorBT
    
    'Bloqueo la solicitud y Actulizo el SolUsuarioR (Analizando)
    Cons = "Select * from Solicitud Where SolCodigo = " & Codigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        If IsNull(RsAux!SolUsuarioR) Or (Not IsNull(RsAux!SolUsuarioR) And RsAux!SolEstado = EstadoSolicitud.ParaRetomar) Then

            cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
            On Error GoTo errorET
            
            RsAux.Requery
            
            If Not IsNull(RsAux!SolUsuarioR) And RsAux!SolEstado <> EstadoSolicitud.ParaRetomar Then
                cBase.RollbackTrans
                Exit Function
            End If
            
            RsAux.Edit
            RsAux!SolUsuarioR = paCodigoDeUsuario
            RsAux!SolEstado = EstadoSolicitud.Pendiente
            If Not IsNull(RsAux!SolUsuarioR) Then retUsuarioR = RsAux!SolUsuarioR
            RsAux.Update
            
            cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
            RsAux.Requery
            fnc_BlockearSolicitud = 1
            
        Else
            If Not IsNull(RsAux!SolFResolucion) Then fnc_BlockearSolicitud = -1
            retUsuarioR = RsAux!SolUsuarioR
        End If
    
    Else
        fnc_BlockearSolicitud = -1
    End If
    
    RsAux.Close
    Screen.MousePointer = 0
    Exit Function

errorBT:
    fnc_BlockearSolicitud = -1
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Function
errorET:
    Resume ErrorRoll
ErrorRoll:
    fnc_BlockearSolicitud = -1
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
End Function

Public Function fnc_AutorizarCredito(mIDSolicitud As Long)
On Error GoTo errorBT

    Dim rsRes As rdoResultset
    Dim mMsgError As String: mMsgError = ""
    
    Screen.MousePointer = 11
    FechaDelServidor
    
    'Selecciono la solicitud para ver si aún no ha sido resuelta
    Cons = "Select * from Solicitud Where SolCodigo = " & mIDSolicitud
    Set rsRes = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsRes.EOF Then
    
        If rsRes!SolEstado <> EstadoSolicitud.Pendiente Then    'Solicitud ya resuelta
            
            If Not IsNull(rsRes!SolUsuarioR) Then
                MsgBox "La solicitud ha sido resuelta por otro usuario (" & z_BuscoUsuario(rsRes!SolUsuarioR, Identificacion:=True) & ").", vbExclamation, "Solicitud Resuelta"
            Else
                MsgBox "La solicitud ha sido resuelta por otro usuario.", vbCritical, "Posible Error"
            End If
        
        Else        'Solicitud SIGUE PENDIENTE
            
            cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
            On Error GoTo errorET
            rsRes.Requery
            
            If rsRes!SolEstado <> EstadoSolicitud.Pendiente Then        'Valido con RS Bloqueado
                mMsgError = "La solicitud ha sido resuelta por otro usuario."
                rsRes.Close
                GoTo errorET: Exit Function
            End If  '----------------------------------------------------------------------------------------
            
            '1) Cambio el Estado de la Solicitud                    ------------------------------------------------
            rsRes.Edit
            rsRes!SolEstado = EstadoSolicitud.Aprovada
            rsRes!SolFResolucion = Format(gFechaServidor, sqlFormatoFH)
            rsRes!SolUsuarioR = paCodigoDeUsuario
            rsRes!SolCondicionR = paResolucionEstandar
            rsRes.Update
                        
            '2) Inserto los renglones de Solicitud Resolucion   ------------------------------------------------
            Dim m_Numero As Byte, rs1 As rdoResultset
            
            Cons = "Select Top 1 * from SolicitudResolucion Where ResSolicitud = " & mIDSolicitud & _
                       " Order by ResNumero DESC"
            
            Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rs1.EOF Then m_Numero = rs1!ResNumero Else m_Numero = 0
            m_Numero = m_Numero + 1
            
            rs1.AddNew
            rs1!ResSolicitud = mIDSolicitud
            rs1!ResNumero = m_Numero
            rs1!ResCondicion = paResolucionEstandar
            rs1!ResTexto = Null
            rs1!ResComentario = "Si (automática)"
            rs1!ResFecha = Format(gFechaServidor, sqlFormatoFH)
            rs1!ResUsuario = paCodigoDeUsuario
            rs1.Update
            
            rs1.Close
            
            cBase.CommitTrans   'FINALIZO TRANSACCION   --------------------------------------------------
            rsRes.Requery
        End If
        
    Else
        MsgBox "La solicitud ha sido eliminada. " & vbCrLf & _
                    "Verifique la lista de solicitudes pendientes.", vbCritical, "Solicitud Eliminada"
    End If
    
    rsRes.Close
    
    Screen.MousePointer = 0
    Exit Function

errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0
    Exit Function
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    If Trim(mMsgError) = "" Then mMsgError = "No se ha podido realizar la transacción. Reintente la operación."
    clsGeneral.OcurrioError mMsgError, Err.Description
    Screen.MousePointer = 0
    Exit Function
End Function


Public Function fnc_ItemUsuarios(idUsr As Long) As String
    On Error GoTo usrAgregar
    
    Dim aItem As String
    aItem = "I" & CStr(idUsr)
    
    fnc_ItemUsuarios = colUsuarios.Item(aItem)
    Exit Function
    
usrAgregar:
    On Error GoTo usrErrAdd
    fnc_ItemUsuarios = z_BuscoUsuario(idUsr, Identificacion:=True)
    colUsuarios.Add fnc_ItemUsuarios, CStr(aItem)
    
usrErrAdd:
End Function


Public Function fnc_ItemImportes(idSol As Long)
    On Error GoTo colAgregar
    
    Dim aItem As String
    aItem = "I" & CStr(idSol)
    
    fnc_ItemImportes = colImportes.Item(aItem)
    fnc_ItemImportes = Trim(Mid(fnc_ItemImportes, InStr(fnc_ItemImportes, "|") + 1))
    Exit Function
    
colAgregar:
    On Error GoTo colErrAdd
    Dim rsMon As rdoResultset
    Dim aMontoSol As Currency
    
    Cons = "Select * from RenglonSolicitud,  TipoCuota" & _
                " Where RSoSolicitud = " & idSol & _
                " And RSoTipoCuota = TCuCodigo"
    Set rsMon = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsMon.EOF
        If Not IsNull(rsMon!RSoValorEntrega) Then aMontoSol = aMontoSol + rsMon!RSoValorEntrega
        If Not IsNull(rsMon!RSoValorCuota) And Not IsNull(rsMon!TCuCantidad) Then aMontoSol = aMontoSol + (rsMon!RSoValorCuota * rsMon!TCuCantidad)
        rsMon.MoveNext
    Loop
    rsMon.Close
    
    fnc_ItemImportes = Format(aMontoSol, "#,##0.00")
    colImportes.Add CStr(idSol) & "|" & fnc_ItemImportes, CStr(aItem)
    
    Exit Function
colErrAdd:
End Function


