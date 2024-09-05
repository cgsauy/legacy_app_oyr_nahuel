Attribute VB_Name = "modConeccion"
'MODULO Conección
'Contiene rutinas y variables del entorno RDO.
Option Explicit

'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno
Public RsAux As rdoResultset         'Resultset Auxiliar

'String.
Public Cons As String
Public paCodigoDeUsuario As Long
Public paCodigoDeSucursal As Long
Public paCodigoDGI As Long
Public paCodigoDeTerminal As Long

Public Function InicioConexionBD(strConexion, Optional sqlTimeOut As Integer = 15) As Boolean
On Error GoTo ErrICBD
    InicioConexionBD = False
    Set eBase = rdoCreateEnvironment("", "", "")
    eBase.CursorDriver = rdUseServer
    'Conexion a la base de datos----------------------------------------
    Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , strConexion)
    cBase.QueryTimeout = sqlTimeOut
    InicioConexionBD = True
    Exit Function
ErrICBD:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al intentar comunicarse con la Base de Datos, se cancelará la ejecución.", vbExclamation, "ATENCIÓN"
End Function

Public Sub CierroConexion()
    On Error GoTo ErrCC
    cBase.Close
    eBase.Close
    Exit Sub
ErrCC:
    On Error Resume Next
End Sub

Public Function PropiedadesConnect(Conexion As String, _
                                                    Optional Database As Boolean = True, Optional DSN As Boolean = False, _
                                                    Optional Server As Boolean = True) As String
Dim aRetorno As String

    On Error GoTo errConnect
    PropiedadesConnect = ""
    Conexion = UCase(Conexion)
    If DSN Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "DSN=") + 4, Len(Conexion)))
    If Server Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "SERVER=") + 7, Len(Conexion)))
    If Database Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "DATABASE=") + 9, Len(Conexion)))
    
    aRetorno = Trim(Mid(aRetorno, 1, InStr(aRetorno, ";") - 1))
    
    PropiedadesConnect = aRetorno
    Exit Function
    
errConnect:
End Function


Public Function RetornoTipoDeUnCampo(rdoTipo As Integer)
    Select Case rdoTipo
        Case rdTypeCHAR             '1
            RetornoTipoDeUnCampo = "char"
        Case rdTypeNUMERIC      '2
            RetornoTipoDeUnCampo = "numeric"
        Case rdTypeDECIMAL      '3
            RetornoTipoDeUnCampo = "decimal"
        Case rdTypeINTEGER         '4
            RetornoTipoDeUnCampo = "int"
        Case rdTypeSMALLINT     ' 5
            RetornoTipoDeUnCampo = "smallint"
        Case rdTypeFLOAT            '6
            RetornoTipoDeUnCampo = "float"
        Case rdTypeREAL             '7
            RetornoTipoDeUnCampo = "real"
        Case rdTypeDOUBLE       '8
            RetornoTipoDeUnCampo = "double"
        Case rdTypeDATE             '9
            RetornoTipoDeUnCampo = "date"
        Case rdTypeTIME                 '10
            RetornoTipoDeUnCampo = "time"
        Case rdTypeTIMESTAMP    '11
            RetornoTipoDeUnCampo = "timestamp"
        Case rdTypeVARCHAR      '12
            RetornoTipoDeUnCampo = "varchar"
        Case rdTypeLONGVARCHAR   '-1
            RetornoTipoDeUnCampo = "longvarchar"
        Case rdTypeBINARY               '-2
            RetornoTipoDeUnCampo = "binary"
        Case rdTypeVARBINARY        '-3
           RetornoTipoDeUnCampo = "varbinary"
        Case rdTypeLONGVARBINARY '-4
            RetornoTipoDeUnCampo = "longvarbinary"
        Case rdTypeBIGINT                   '-5
            RetornoTipoDeUnCampo = "bigint"
        Case rdTypeTINYINT                  '-6
            RetornoTipoDeUnCampo = "tinyint"
        Case rdTypeBIT                          '-7
            RetornoTipoDeUnCampo = "bit"
    End Select

End Function

Public Function RetornoFormatoSegunTipo(rdoTipo As Integer)

    Select Case rdoTipo
        Case rdTypeCHAR, rdTypeVARCHAR, rdTypeLONGVARCHAR              '1, 12, -1
            RetornoFormatoSegunTipo = "#"
        
        Case rdTypeNUMERIC, rdTypeDECIMAL, rdTypeFLOAT, rdTypeREAL, rdTypeDOUBLE       '2, 3, , 7, 8
            RetornoFormatoSegunTipo = "#,##0.00"
        
        Case rdTypeINTEGER, rdTypeSMALLINT, rdTypeBIGINT, rdTypeTINYINT, rdTypeBINARY, rdTypeVARBINARY, rdTypeLONGVARBINARY           '4, 5, -5, -6, -2,-3,-4
            RetornoFormatoSegunTipo = "#,##0"
        
        Case rdTypeDATE             '9
            RetornoFormatoSegunTipo = "d/Mmm/yyyy"
            
        Case rdTypeTIME                 '10
            RetornoFormatoSegunTipo = "hh:mm:ss"
            
        Case rdTypeTIMESTAMP    '11
            RetornoFormatoSegunTipo = "d/mm/yyyy hh:mm:ss"
        
        Case rdTypeBIT                   '-7
            'la barra imprime, Formato (valores +; valores -; valor = 0)
            RetornoFormatoSegunTipo = "\S\i;\S\i;\N\o"
            
    End Select

End Function
Public Function ValidoCampoSegunFormato(rdoTipo As Integer, Campo As String) As Boolean

    ValidoCampoSegunFormato = False
    
    Select Case rdoTipo
        Case rdTypeCHAR, rdTypeVARCHAR, rdTypeLONGVARCHAR
            ValidoCampoSegunFormato = True
        
        Case rdTypeNUMERIC, rdTypeDECIMAL, rdTypeFLOAT, rdTypeREAL, rdTypeDOUBLE
            If IsNumeric(Campo) Then ValidoCampoSegunFormato = True
        
        Case rdTypeINTEGER, rdTypeSMALLINT, rdTypeBIGINT, rdTypeTINYINT, rdTypeBINARY, rdTypeVARBINARY, rdTypeLONGVARBINARY           '4, 5, -5, -6, -2,-3,-4
            If IsNumeric(Campo) Then ValidoCampoSegunFormato = True
        
        Case rdTypeDATE, rdTypeTIMESTAMP
            If IsDate(Campo) Then ValidoCampoSegunFormato = True
        
        Case rdTypeBIT
            If IsNumeric(Campo) Then ValidoCampoSegunFormato = True
    End Select

End Function

Public Function z_BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
    z_BuscoUsuario = BuscoUsuario(Codigo, Identificacion, Digito, Iniciales)
End Function

Public Function BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
Dim RsUsr As rdoResultset
Dim aRetorno As String: aRetorno = ""
    
    On Error Resume Next
    
    Cons = "Select * from Usuario Where UsuCodigo = " & Codigo
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Digito Then aRetorno = Trim(RsUsr!UsuDigito)
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    
    BuscoUsuario = aRetorno
    
End Function
Public Function z_BuscoUsuarioDigito(Digito As Long, Optional Codigo As Boolean = False, Optional Identificacion As Boolean = False, Optional Iniciales As Boolean = False) As Variant
    z_BuscoUsuarioDigito = BuscoUsuarioDigito(Digito, Codigo, Identificacion, Iniciales)
End Function

Public Function BuscoUsuarioDigito(Digito As Long, Optional Codigo As Boolean = False, Optional Identificacion As Boolean = False, Optional Iniciales As Boolean = False) As Variant
Dim RsUsr As rdoResultset
Dim aRetorno As Variant
On Error GoTo ErrBUD

    Cons = "Select * from Usuario Where UsuDigito = " & Digito
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Codigo Then aRetorno = RsUsr!UsuCodigo
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    BuscoUsuarioDigito = aRetorno
    Exit Function
    
ErrBUD:
    MsgBox "Error inesperado al buscar el usuario.", vbCritical, "ATENCIÓN"
End Function

Public Function RetornoDataTypeGrilla(intRdoType As Integer) As Integer
    
    Select Case intRdoType
        Case rdTypeCHAR, rdTypeVARCHAR, rdTypeLONGVARCHAR              '1, 12, -1
            RetornoDataTypeGrilla = 8
        
        Case rdTypeDOUBLE: RetornoDataTypeGrilla = 5
        
        Case rdTypeNUMERIC, rdTypeDECIMAL, rdTypeFLOAT, rdTypeREAL        '2, 3, , 7, 8
            RetornoDataTypeGrilla = 6
        
        Case rdTypeINTEGER, rdTypeSMALLINT, rdTypeBIGINT, rdTypeTINYINT, rdTypeBINARY, rdTypeVARBINARY, rdTypeLONGVARBINARY           '4, 5, -5, -6, -2,-3,-4
            RetornoDataTypeGrilla = 20
        
        Case rdTypeDATE, rdTypeTIME, rdTypeTIMESTAMP
            RetornoDataTypeGrilla = 7
        
        Case rdTypeBIT                   '-7
            RetornoDataTypeGrilla = 11
            
    End Select
End Function

Public Function TesteoConexion(ByVal strDSN As String, strConexion As String) As Boolean
Dim rCon As rdoConnection
'...........................................................................................................................................................
'Retorna si la conexión del odbc se realizo con exito y  además toda la cadena de conexión.
'Parametros: Nombre del odbc y cadena de conexión (opcional).
'
'OIR 18-9-2000
'...........................................................................................................................................................

    Screen.MousePointer = 11
    TesteoConexion = False
    'Mapeo el error.
    On Error GoTo ErrCC
    
    'Si requiere pwd invoca automáticamente para el logueo del mismo.
    Set rCon = eBase.OpenConnection(strDSN, rdDriverCompleteRequired, , strConexion)
        
    strConexion = rCon.Connect      'Cargo el string de conexión.
    
    rCon.Close  'Cierro conexión.
    
    TesteoConexion = True
    Screen.MousePointer = 0
    Exit Function
    
ErrCC:
    Screen.MousePointer = 0
    Exit Function
    
End Function

Public Function PropiedadesConnectPorClave(ByVal strConexion As String, _
                                                    ByVal strClave As String) As String
Dim strRetorno As String
Dim intPos As Integer

'...........................................................................................................................................................
'Dada un clave en una cadena de conexión retorna el valor de dicha clave.
'
'OIR 18-9-2000
'...........................................................................................................................................................
    On Error GoTo errConnect
    
    PropiedadesConnectPorClave = ""
    strRetorno = strConexion    'Hago copia para poder retornar el verdadero formato de la clave.
    strConexion = UCase(strConexion)
    strClave = UCase(strClave) & "="
    intPos = InStr(1, strConexion, strClave)
    If intPos > 0 Then
        strRetorno = Trim(Mid(strRetorno, intPos + Len(strClave), Len(strConexion)))
        PropiedadesConnectPorClave = Trim(Mid(strRetorno, 1, InStr(strRetorno, ";") - 1))
    End If
    Exit Function
    
errConnect:
End Function


