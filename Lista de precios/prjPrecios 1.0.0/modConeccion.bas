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
Public paCodigoDeTerminal As Long

' <VB WATCH>
Const VBWMODULE = "modConeccion"
' </VB WATCH>

Public Function InicioConexionBD(strConexion, Optional sqlTimeOut As Integer = 15) As Boolean
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>
2      On Error GoTo ErrICBD
3          InicioConexionBD = False
4          Set eBase = rdoCreateEnvironment("", "", "")
5          eBase.CursorDriver = rdUseServer
           'Conexion a la base de datos----------------------------------------
6          Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , strConexion)
7          cBase.QueryTimeout = sqlTimeOut
8          InicioConexionBD = True
9          Exit Function
10     ErrICBD:
11         On Error Resume Next
12         Screen.MousePointer = 0
13         MsgBox "Ocurrió un error al intentar comunicarse con la Base de Datos, se cancelará la ejecución.", vbExclamation, "ATENCIÓN"
' <VB WATCH>
14         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "InicioConexionBD"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Public Sub CierroConexion()
' <VB WATCH>
15         On Error GoTo vbwErrHandler
' </VB WATCH>
16         On Error GoTo ErrCC
17         cBase.Close
18         eBase.Close
19         Exit Sub
20     ErrCC:
21         On Error Resume Next
' <VB WATCH>
22         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CierroConexion"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Public Function PropiedadesConnect(Conexion As String, _
                                                    Optional Database As Boolean = True, Optional DSN As Boolean = False, _
                                                    Optional Server As Boolean = True) As String
' <VB WATCH>
23         On Error GoTo vbwErrHandler
' </VB WATCH>
24     Dim aRetorno As String

25         On Error GoTo errConnect
26         PropiedadesConnect = ""
27         Conexion = UCase(Conexion)
28         If DSN Then
29              aRetorno = Trim(Mid(Conexion, InStr(Conexion, "DSN=") + 4, Len(Conexion)))
30         End If
31         If Server Then
32              aRetorno = Trim(Mid(Conexion, InStr(Conexion, "SERVER=") + 7, Len(Conexion)))
33         End If
34         If Database Then
35              aRetorno = Trim(Mid(Conexion, InStr(Conexion, "DATABASE=") + 9, Len(Conexion)))
36         End If

37         aRetorno = Trim(Mid(aRetorno, 1, InStr(aRetorno, ";") - 1))

38         PropiedadesConnect = aRetorno
39         Exit Function

40     errConnect:
' <VB WATCH>
41         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PropiedadesConnect"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function


Public Function RetornoTipoDeUnCampo(rdoTipo As Integer)
' <VB WATCH>
42         On Error GoTo vbwErrHandler
' </VB WATCH>
43         Select Case rdoTipo
               Case rdTypeCHAR             '1
44                 RetornoTipoDeUnCampo = "char"
45             Case rdTypeNUMERIC      '2
46                 RetornoTipoDeUnCampo = "numeric"
47             Case rdTypeDECIMAL      '3
48                 RetornoTipoDeUnCampo = "decimal"
49             Case rdTypeINTEGER         '4
50                 RetornoTipoDeUnCampo = "int"
51             Case rdTypeSMALLINT     ' 5
52                 RetornoTipoDeUnCampo = "smallint"
53             Case rdTypeFLOAT            '6
54                 RetornoTipoDeUnCampo = "float"
55             Case rdTypeREAL             '7
56                 RetornoTipoDeUnCampo = "real"
57             Case rdTypeDOUBLE       '8
58                 RetornoTipoDeUnCampo = "double"
59             Case rdTypeDATE             '9
60                 RetornoTipoDeUnCampo = "date"
61             Case rdTypeTIME                 '10
62                 RetornoTipoDeUnCampo = "time"
63             Case rdTypeTIMESTAMP    '11
64                 RetornoTipoDeUnCampo = "timestamp"
65             Case rdTypeVARCHAR      '12
66                 RetornoTipoDeUnCampo = "varchar"
67             Case rdTypeLONGVARCHAR   '-1
68                 RetornoTipoDeUnCampo = "longvarchar"
69             Case rdTypeBINARY               '-2
70                 RetornoTipoDeUnCampo = "binary"
71             Case rdTypeVARBINARY        '-3
72                RetornoTipoDeUnCampo = "varbinary"
73             Case rdTypeLONGVARBINARY '-4
74                 RetornoTipoDeUnCampo = "longvarbinary"
75             Case rdTypeBIGINT                   '-5
76                 RetornoTipoDeUnCampo = "bigint"
77             Case rdTypeTINYINT                  '-6
78                 RetornoTipoDeUnCampo = "tinyint"
79             Case rdTypeBIT                          '-7
80                 RetornoTipoDeUnCampo = "bit"
81         End Select

' <VB WATCH>
82         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "RetornoTipoDeUnCampo"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Public Function RetornoFormatoSegunTipo(rdoTipo As Integer)
' <VB WATCH>
83         On Error GoTo vbwErrHandler
' </VB WATCH>

84         Select Case rdoTipo
               Case rdTypeCHAR, rdTypeVARCHAR, rdTypeLONGVARCHAR              '1, 12, -1
85                 RetornoFormatoSegunTipo = "#"

86             Case rdTypeNUMERIC, rdTypeDECIMAL, rdTypeFLOAT, rdTypeREAL, rdTypeDOUBLE       '2, 3, , 7, 8
87                 RetornoFormatoSegunTipo = "#,##0.00"

88             Case rdTypeINTEGER, rdTypeSMALLINT, rdTypeBIGINT, rdTypeTINYINT, rdTypeBINARY, rdTypeVARBINARY, rdTypeLONGVARBINARY           '4, 5, -5, -6, -2,-3,-4
89                 RetornoFormatoSegunTipo = "#,##0"

90             Case rdTypeDATE             '9
91                 RetornoFormatoSegunTipo = "d/Mmm/yyyy"

92             Case rdTypeTIME                 '10
93                 RetornoFormatoSegunTipo = "hh:mm:ss"

94             Case rdTypeTIMESTAMP    '11
95                 RetornoFormatoSegunTipo = "d/mm/yyyy hh:mm:ss"

96             Case rdTypeBIT                   '-7
                   'la barra imprime, Formato (valores +; valores -; valor = 0)
97                 RetornoFormatoSegunTipo = "\S\i;\S\i;\N\o"

98         End Select

' <VB WATCH>
99         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "RetornoFormatoSegunTipo"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function
Public Function ValidoCampoSegunFormato(rdoTipo As Integer, Campo As String) As Boolean
' <VB WATCH>
100        On Error GoTo vbwErrHandler
' </VB WATCH>

101        ValidoCampoSegunFormato = False

102        Select Case rdoTipo
               Case rdTypeCHAR, rdTypeVARCHAR, rdTypeLONGVARCHAR
103                ValidoCampoSegunFormato = True

104            Case rdTypeNUMERIC, rdTypeDECIMAL, rdTypeFLOAT, rdTypeREAL, rdTypeDOUBLE
105                If IsNumeric(Campo) Then
106                     ValidoCampoSegunFormato = True
107                End If

108            Case rdTypeINTEGER, rdTypeSMALLINT, rdTypeBIGINT, rdTypeTINYINT, rdTypeBINARY, rdTypeVARBINARY, rdTypeLONGVARBINARY           '4, 5, -5, -6, -2,-3,-4
109                If IsNumeric(Campo) Then
110                     ValidoCampoSegunFormato = True
111                End If

112            Case rdTypeDATE, rdTypeTIMESTAMP
113                If IsDate(Campo) Then
114                     ValidoCampoSegunFormato = True
115                End If

116            Case rdTypeBIT
117                If IsNumeric(Campo) Then
118                     ValidoCampoSegunFormato = True
119                End If
120        End Select

' <VB WATCH>
121        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ValidoCampoSegunFormato"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Public Function BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
' <VB WATCH>
122        On Error GoTo vbwErrHandler
' </VB WATCH>
123    Dim RsUsr As rdoResultset
124    Dim aRetorno As String
125    aRetorno = ""

126        On Error Resume Next

127        Cons = "Select * from Usuario Where UsuCodigo = " & Codigo
128        Set RsUsr = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
129        If Not RsUsr.EOF Then
130            If Identificacion Then
131                 aRetorno = Trim(RsUsr!UsuIdentificacion)
132            End If
133            If Digito Then
134                 aRetorno = Trim(RsUsr!UsuDigito)
135            End If
136            If Iniciales Then
137                 aRetorno = Trim(RsUsr!UsuInicial)
138            End If
139        End If
140        RsUsr.Close

141        BuscoUsuario = aRetorno

' <VB WATCH>
142        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "BuscoUsuario"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Public Function BuscoUsuarioDigito(Digito As Long, Optional Codigo As Boolean = False, Optional Identificacion As Boolean = False, Optional Iniciales As Boolean = False) As Variant
' <VB WATCH>
143        On Error GoTo vbwErrHandler
' </VB WATCH>
144    Dim RsUsr As rdoResultset
145    Dim aRetorno As Variant
146    On Error GoTo ErrBUD

147        Cons = "Select * from Usuario Where UsuDigito = " & Digito
148        Set RsUsr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
149        If Not RsUsr.EOF Then
150            If Identificacion Then
151                 aRetorno = Trim(RsUsr!UsuIdentificacion)
152            End If
153            If Codigo Then
154                 aRetorno = RsUsr!UsuCodigo
155            End If
156            If Iniciales Then
157                 aRetorno = Trim(RsUsr!UsuInicial)
158            End If
159        End If
160        RsUsr.Close
161        BuscoUsuarioDigito = aRetorno
162        Exit Function

163    ErrBUD:
164        MsgBox "Error inesperado al buscar el usuario.", vbCritical, "ATENCIÓN"
' <VB WATCH>
165        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "BuscoUsuarioDigito"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Public Function RetornoDataTypeGrilla(intRdoType As Integer) As Integer
' <VB WATCH>
166        On Error GoTo vbwErrHandler
' </VB WATCH>

167        Select Case intRdoType
               Case rdTypeCHAR, rdTypeVARCHAR, rdTypeLONGVARCHAR              '1, 12, -1
168                RetornoDataTypeGrilla = 8

169            Case rdTypeDOUBLE
170            RetornoDataTypeGrilla = 5

171            Case rdTypeNUMERIC, rdTypeDECIMAL, rdTypeFLOAT, rdTypeREAL        '2, 3, , 7, 8
172                RetornoDataTypeGrilla = 6

173            Case rdTypeINTEGER, rdTypeSMALLINT, rdTypeBIGINT, rdTypeTINYINT, rdTypeBINARY, rdTypeVARBINARY, rdTypeLONGVARBINARY           '4, 5, -5, -6, -2,-3,-4
174                RetornoDataTypeGrilla = 20

175            Case rdTypeDATE, rdTypeTIME, rdTypeTIMESTAMP
176                RetornoDataTypeGrilla = 7

177            Case rdTypeBIT                   '-7
178                RetornoDataTypeGrilla = 11

179        End Select
' <VB WATCH>
180        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "RetornoDataTypeGrilla"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Public Function TesteoConexion(ByVal strDSN As String, strConexion As String) As Boolean
' <VB WATCH>
181        On Error GoTo vbwErrHandler
' </VB WATCH>
182    Dim rCon As rdoConnection
       '...........................................................................................................................................................
       'Retorna si la conexión del odbc se realizo con exito y  además toda la cadena de conexión.
       'Parametros: Nombre del odbc y cadena de conexión (opcional).
       '
       'OIR 18-9-2000
       '...........................................................................................................................................................

183        Screen.MousePointer = 11
184        TesteoConexion = False
           'Mapeo el error.
185        On Error GoTo ErrCC

           'Si requiere pwd invoca automáticamente para el logueo del mismo.
186        Set rCon = eBase.OpenConnection(strDSN, rdDriverCompleteRequired, , strConexion)

187        strConexion = rCon.Connect      'Cargo el string de conexión.

188        rCon.Close  'Cierro conexión.

189        TesteoConexion = True
190        Screen.MousePointer = 0
191        Exit Function

192    ErrCC:
193        Screen.MousePointer = 0
194        Exit Function

' <VB WATCH>
195        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "TesteoConexion"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

Public Function PropiedadesConnectPorClave(ByVal strConexion As String, _
                                                    ByVal strClave As String) As String
' <VB WATCH>
196        On Error GoTo vbwErrHandler
' </VB WATCH>
197    Dim strRetorno As String
198    Dim intPos As Integer

       '...........................................................................................................................................................
       'Dada un clave en una cadena de conexión retorna el valor de dicha clave.
       '
       'OIR 18-9-2000
       '...........................................................................................................................................................
199        On Error GoTo errConnect

200        PropiedadesConnectPorClave = ""
201        strRetorno = strConexion    'Hago copia para poder retornar el verdadero formato de la clave.
202        strConexion = UCase(strConexion)
203        strClave = UCase(strClave) & "="
204        intPos = InStr(1, strConexion, strClave)
205        If intPos > 0 Then
206            strRetorno = Trim(Mid(strRetorno, intPos + Len(strClave), Len(strConexion)))
207            PropiedadesConnectPorClave = Trim(Mid(strRetorno, 1, InStr(strRetorno, ";") - 1))
208        End If
209        Exit Function

210    errConnect:
' <VB WATCH>
211        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PropiedadesConnectPorClave"

    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function



