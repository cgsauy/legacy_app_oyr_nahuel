Attribute VB_Name = "ModCrystal"
'------------------------------------------------------------------------------------------------
'Modulo Crystal Engine
'Contiene rutinas de impresión utilizando el Engine del Crystal Report.
'Autor    = A&A analistas
'Fecha  = Junio-1999
'------------------------------------------------------------------------------------------------
Option Explicit

'Constantes locales de seteo de impresion
Dim cnUsuario As String
Dim cnPass As String
Dim cnDSN As String
Dim cnBD As String

' Logging on is performed when printing the report, but the correct
' log on information must first be set using PESetNthTableLogOnInfo.
' Only the password is required, but the server, database, and
' user names may optionally be overriden as well.
'
' If the parameter propagateAcrossTables is TRUE, the new log on info
' is also applied to any other tables in this report that had the
' same original server and database names as this table.  If FALSE
' only this table is updated.
'
' Logging off is performed automatically when the print job is closed.

Global Const PE_SERVERNAME_LEN = 128
Global Const PE_DATABASENAME_LEN = 128
Global Const PE_USERID_LEN = 128
Global Const PE_PASSWORD_LEN = 128
Global Const PE_SIZEOF_LOGON_INFO = 514  ' # bytes in PELogOnInfo

'Constants using to calculate structure size constants
Global Const PE_BYTE_LEN = 1
Global Const PE_WORD_LEN = 2
Global Const PE_LONG_LEN = 4
Global Const PE_DOUBLE_LEN = 8

Type PELogOnInfo
    ' initialize to # bytes in PELogOnInfo
    StructSize As Integer

    ' For any of the following values an empty string ("") means to use
    ' the value already set in the report.  To override a value in the
    ' report use a non-empty string (e.g. "Server A").
    '
    ' For Netware SQL, pass the dictionary path name in ServerName and
    ' data path name in DatabaseName.

    ServerName As String * PE_SERVERNAME_LEN
    DatabaseName  As String * PE_DATABASENAME_LEN
    UserID As String * PE_USERID_LEN

    ' Password is undefined when getting information from report.
    Password  As String * PE_PASSWORD_LEN
End Type

Global Const PE_DLL_NAME_LEN = 64
Global Const PE_FULL_NAME_LEN = 256
Global Const PE_SIZEOF_TABLE_TYPE = 324 ' # bytes in PETableType

Global Const PE_DT_STANDARD = 1  ' values for DBType
Global Const PE_DT_SQL = 2

Type PETableType
    StructSize As Integer   ' initialize to # bytes in PETableType

    DLLName As String * PE_DLL_NAME_LEN
    DescriptiveName  As String * PE_FULL_NAME_LEN

    DBType As Integer
End Type

Declare Function PEGetNthTableType Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, TableType As PETableType) As Integer
Declare Function PEGetNthTableLogOnInfo Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, LogOnInfo As PELogOnInfo) As Integer
Declare Function crPESetNthTableLogOnInfo Lib "crwrap32.dll" (ByVal printJob As Integer, ByVal TableN As Integer, ByVal ServerName As String, ByVal dbName As String, ByVal UserID As String, ByVal Password As String, ByVal PropagateAcrossTables As Long) As Integer


'Declaración de API.
Private Declare Function GetActiveWindow& Lib "user32" ()        'Obtengo la ventana abierta.
Private Declare Function IsWindow& Lib "user32" (ByVal hwnd As Long)    'Verifico si esta abierta la ventana.

Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
Private Declare Function PEGetWindowHandle Lib "crpe32.dll" (ByVal printJob As Integer) As Integer

Private Declare Function crPEGetSelectedPrinter Lib "crwrap32.dll" Alias "crvbPEGetSelectedPrinter" (ByVal printJob As Integer, ByRef driverName As String, ByRef printerName As String, ByRef portName As String, crmode As crDEVMODE) As Integer
Private Declare Function crPESelectPrinter Lib "crwrap32.dll" (ByVal printJob As Integer, ByVal driverName As String, ByVal printerName As String, ByVal portName As String, crmode As crDEVMODE) As Integer

Declare Function crPELogOnServer Lib "crwrap32.dll" (ByVal DLLName As String, ByVal ServerName As String, ByVal dbName As String, ByVal UserID As String, ByVal Password As String) As Integer
Declare Function crPELogOffServer Lib "crwrap32.dll" (ByVal DLLName As String, ByVal ServerName As String, ByVal dbName As String, ByVal UserID As String, ByVal Password As String) As Integer

Private Declare Function PEOpenEngine Lib "crpe32.dll" () As Boolean
Private Declare Sub PECloseEngine Lib "crpe32.dll" ()
Private Declare Function PEOpenPrintJob Lib "crpe32.dll" (ByVal RptName As String) As Integer
Private Declare Sub PEClosePrintJob Lib "crpe32.dll" (ByVal printJob As Integer)
Private Declare Function PEOutputToPrinter Lib "crpe32.dll" (ByVal printJob As Integer, ByVal nCopies As Integer) As Integer
Private Declare Function PEOutputToWindow Lib "crpe32.dll" (ByVal printJob As Integer, ByVal title As String, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal style As Long, ByVal PWindow As Long) As Integer
Private Declare Function PEStartPrintJob Lib "crpe32.dll" (ByVal printJob As Integer, ByVal WaitOrNot As Integer) As Integer
Private Declare Function PEEnableProgressDialog Lib "crpe32.dll" (ByVal printJob%, ByVal enable%) As Integer
Private Declare Function PEDiscardSavedData Lib "crpe32.dll" (ByVal printJob As Integer) As Integer

'SUBREPORTES----------------------------
Private Declare Function PEOpenSubreport Lib "crpe32.dll" (ByVal parentJob As Integer, ByVal subreportName As String) As Integer
Private Declare Function PECloseSubreport Lib "crpe32.dll" (ByVal printJob As Integer) As Integer
'--------------------------------------------------------

Private Declare Function PESetSQLQuery Lib "crpe32.dll" (ByVal printJob As Integer, ByVal QueryString As String) As Integer

'FORMULAS------------------------------------------------------------------------------------
Private Declare Function PEGetFormula Lib "crpe32.dll" (ByVal printJob As Integer, ByVal FormulaName As String, textHandle As Long, TextLength As Integer) As Integer
Private Declare Function PESetFormula Lib "crpe32.dll" (ByVal printJob As Integer, ByVal FormulaName As String, ByVal FormulaString As String) As Integer
Private Declare Function PEGetNFormulas Lib "crpe32.dll" (ByVal printJob As Integer) As Integer
Private Declare Function PEGetNthFormula Lib "crpe32.dll" (ByVal printJob As Integer, ByVal FormulaN As Integer, NameHandle As Long, NameLength As Integer, textHandle As Long, TextLength As Integer) As Integer
Private Declare Function PECheckFormula Lib "crpe32.dll" (ByVal printJob As Integer, ByVal FormulaName As String) As Integer

'Obtiene el string al cual apunta el HANDLE
Private Declare Function crvbHandleToBStr Lib "crwrap32.dll" (ByRef BString As String, ByVal strHandle As Long, ByVal strLength As Integer) As Integer
'----------------------------------------------------------------------------------------------------------------

'CONSTANTES--------------------------------------------------------------------------------------------
Const CW_USEDEFAULT = &H80000000

Const WS_MINIMIZE = 536870912
Const WS_VISIBLE = 268435456
Const WS_DISABLED = 134217728  'Make a window that is disabled when it first appears.
Const WS_CLIPSIBLINGS = 67108864   'Clip child windows with respect to one another.
Const WS_CLIPCHILDREN = 33554432  ' Exclude the area occupied by child windows when drawing inside the parent window.
Const WS_MAXIMIZE = 16777216   'Make a window of maximum size.
Const WS_CAPTION = 12582912    'Make a window that includes a title bar.
Const WS_BORDER = 8388608  'Make a window that includes a border.
Const WS_DLGFRAME = 4194304 'Make a window that has a double border but no title.
Const WS_VSCROLL = 2097152 'Make a window that includes a vertical scroll bar.
Const WS_HSCROLL = 1048576 'Make a window that includes a horizontal scroll bar.
Const WS_SYSMENU = 524288  'Include the system menu box.
Const WS_THICKFRAME = 262144   'Include the thick frame that can be used to size the window.
Const WS_MINIMIZEBOX = 131072   'Include the minimize box.
Const WS_MAXIMIZEBOX = 65536    'Include the maximize box.
Const CW_USEDFAULT = -32768     'Assign the child window the default horizontal and vertical position, and the default height and width.

'================================================================================
'SETEOS PARA EL CAMPO DMFIELD DEL DEVMODE
'/* field selection bits */
 Const DM_ORIENTATION = &H1
 Const DM_PAPERSIZE = &H2
 Const DM_PAPERLENGTH = &H4
 Const DM_PAPERWIDTH = &H8
 Const DM_SCALE = &H10
 Const DM_COPIES = &H100
 Const DM_DEFAULTSOURCE = &H200
 Const DM_PRINTQUALITY = &H400
 Const DM_COLOR = &H800
 Const DM_DUPLEX = &H1000
 Const DM_YRESOLUTION = &H2000
 Const DM_TTOPTION = &H4000

'/* orientation selections */
 Const DMORIENT_PORTRAIT = 1
 Const DMORIENT_LANDSCAPE = 2

'/* paper selections */
 '/*  Warning: The PostScript driver mistakingly uses DMPAPER_ values between
' *  50 and 56.  Don't use this range when defining new paper sizes.
' */
 Const DMPAPER_FIRST = 1
 Const DMPAPER_LETTER = 1               '/* Letter 8 1/2 x 11 in               */
 Const DMPAPER_LETTERSMALL = 2          '/* Letter Small 8 1/2 x 11 in         */
 Const DMPAPER_TABLOID = 3              '/* Tabloid 11 x 17 in                 */
 Const DMPAPER_LEDGER = 4               '/* Ledger 17 x 11 in                  */
 Const DMPAPER_LEGAL = 5                '/* Legal 8 1/2 x 14 in                */
 Const DMPAPER_STATEMENT = 6            '/* Statement 5 1/2 x 8 1/2 in         */
 Const DMPAPER_EXECUTIVE = 7            '/* Executive 7 1/4 x 10 1/2 in        */
 Const DMPAPER_A3 = 8                   '/* A3 297 x 420 mm                    */
 Const DMPAPER_A4 = 9                   '/* A4 210 x 297 mm                    */
 Const DMPAPER_A4SMALL = 10             '/* A4 Small 210 x 297 mm              */
 Const DMPAPER_A5 = 11                  '/* A5 148 x 210 mm                    */
 Const DMPAPER_B4 = 12                  '/* B4 250 x 354                       */
 Const DMPAPER_B5 = 13                  '/* B5 182 x 257 mm                    */
 Const DMPAPER_FOLIO = 14               '/* Folio 8 1/2 x 13 in                */
 Const DMPAPER_QUARTO = 15              '/* Quarto 215 x 275 mm                */
 Const DMPAPER_10X14 = 16               '/* 10x14 in                           */
 Const DMPAPER_11X17 = 17               '/* 11x17 in                           */
 Const DMPAPER_NOTE = 18                '/* Note 8 1/2 x 11 in                 */
 Const DMPAPER_ENV_9 = 19               '/* Envelope #9 3 7/8 x 8 7/8          */
 Const DMPAPER_ENV_10 = 20              '/* Envelope #10 4 1/8 x 9 1/2         */
 Const DMPAPER_ENV_11 = 21              '/* Envelope #11 4 1/2 x 10 3/8        */
 Const DMPAPER_ENV_12 = 22              '/* Envelope #12 4 \276 x 11           */
 Const DMPAPER_ENV_14 = 23              '/* Envelope #14 5 x 11 1/2            */
 Const DMPAPER_CSHEET = 24              '/* C size sheet                       */
 Const DMPAPER_DSHEET = 25              '/* D size sheet                       */
 Const DMPAPER_ESHEET = 26              '/* E size sheet                       */
 Const DMPAPER_ENV_DL = 27              '/* Envelope DL 110 x 220mm            */
 Const DMPAPER_ENV_C5 = 28              '/* Envelope C5 162 x 229 mm           */
 Const DMPAPER_ENV_C3 = 29              '/* Envelope C3  324 x 458 mm          */
 Const DMPAPER_ENV_C4 = 30              '/* Envelope C4  229 x 324 mm          */
 Const DMPAPER_ENV_C6 = 31              '/* Envelope C6  114 x 162 mm          */
 Const DMPAPER_ENV_C65 = 32             '/* Envelope C65 114 x 229 mm          */
 Const DMPAPER_ENV_B4 = 33              '/* Envelope B4  250 x 353 mm          */
 Const DMPAPER_ENV_B5 = 34              '/* Envelope B5  176 x 250 mm          */
 Const DMPAPER_ENV_B6 = 35              '/* Envelope B6  176 x 125 mm          */
 Const DMPAPER_ENV_ITALY = 36           '/* Envelope 110 x 230 mm              */
 Const DMPAPER_ENV_MONARCH = 37         '/* Envelope Monarch 3.875 x 7.5 in    */
 Const DMPAPER_ENV_PERSONAL = 38        '/* 6 3/4 Envelope 3 5/8 x 6 1/2 in    */
 Const DMPAPER_FANFOLD_US = 39          '/* US Std Fanfold 14 7/8 x 11 in      */
 Const DMPAPER_FANFOLD_STD_GERMAN = 40  '/* German Std Fanfold 8 1/2 x 12 in   */
 Const DMPAPER_FANFOLD_LGL_GERMAN = 41  '/* German Legal Fanfold 8 1/2 x 13 in */

 Const DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN

 Const DMPAPER_USER = 256

'/* bin selections */
 Const DMBIN_FIRST = 1
 Const DMBIN_UPPER = 1
 Const DMBIN_ONLYONE = 1
 Const DMBIN_LOWER = 2
 Const DMBIN_MIDDLE = 3
 Const DMBIN_MANUAL = 4
 Const DMBIN_ENVELOPE = 5
 Const DMBIN_ENVMANUAL = 6
 Const DMBIN_AUTO = 7
 Const DMBIN_TRACTOR = 8
 Const DMBIN_SMALLFMT = 9
 Const DMBIN_LARGEFMT = 10
 Const DMBIN_LARGECAPACITY = 11
 Const DMBIN_CASSETTE = 14
 Const DMBIN_LAST = DMBIN_CASSETTE

 Const DMBIN_USER = 256             '/* device specific bins start here */

'/* print qualities */
 Const DMRES_DRAFT = -1
 Const DMRES_LOW = -2
 Const DMRES_MEDIUM = -3
 Const DMRES_HIGH = -4

'/* color enable/disable for color printers */
 Const DMCOLOR_MONOCHROME = 1
 Const DMCOLOR_COLOR = 2

'/* duplex enable */
 Const DMDUP_SIMPLEX = 1
 Const DMDUP_VERTICAL = 2
 Const DMDUP_HORIZONTAL = 3

'/* TrueType options */
 Const DMTT_BITMAP = 1          '/* print TT fonts as graphics */
 Const DMTT_DOWNLOAD = 2        '/* download TT fonts as soft fonts */
 Const DMTT_SUBDEV = 3          '/* substitute device fonts for TT fonts */

'================================================================================

'Tipos.----------------------------------------------------------------

Public Type crDEVMODE
    dmDriverVersion As Integer
    ' printer driver version number (usually not required)
#If Win16 Then
    ' add padding so it aligns the same way under both 16-and 32-bit environments
    pad1 As Integer
#End If
    dmFields As Long
   'flags indicating fields to modify (required)
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
End Type


Type PEWindowOptions
    StructSize As Integer
    hasGroupTree As Integer
    canDrillDown As Integer
    hasNavigationControls As Integer
    hasCancelButton As Integer
    hasPrintButton As Integer
    hasExportButton As Integer
    hasZoomControl As Integer
    hasCloseButton As Integer
    hasProgressControls As Integer
    hasSearchButton As Integer
    hasPrintSetupButton As Integer
    hasRefreshButton As Integer
End Type

Global Const PE_SIZEOF_WINDOW_OPTIONS = 13 * PE_WORD_LEN

Declare Function PEGetWindowOptions Lib "crpe32.dll" (ByVal printJob%, Options As PEWindowOptions) As Integer
Declare Function PESetWindowOptions Lib "crpe32.dll" (ByVal printJob%, Options As PEWindowOptions) As Integer

Dim crResult As Long                  'Resultado de cada operacion
Public crMsgErr As String         'Variable global para cargar mensajes de error

' <VB WATCH>
Const VBWMODULE = "ModCrystal"
' </VB WATCH>

Public Function crDescartoDatos(jobnum As Integer) As Integer
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>
2           crMsgErr = ""
3           crDescartoDatos = PEDiscardSavedData(jobnum%)
4           If crDescartoDatos = 0 Then
5                crMsgErr = "Ocurrió un error al intentar descartar datos almacenados."
6           End If
' <VB WATCH>
7          Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crDescartoDatos"

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

Public Function crSeteoFormula(jobnum As Integer, NombreFormula As String, TextoFormula As String) As Integer
' <VB WATCH>
8          On Error GoTo vbwErrHandler
' </VB WATCH>

9          crMsgErr = ""
           'Si retorna Cero dio error.
10         crResult = PESetFormula(jobnum%, NombreFormula, TextoFormula)
11         crSeteoFormula = crResult

12         If crResult = 0 Then
13              crMsgErr = "Ocurrió un error al setear la formula " & Trim(NombreFormula) & " := " & Trim(TextoFormula) & "."
14              Exit Function
15         End If

           'Si retorna Cero dio error.
16         crResult = PECheckFormula(jobnum%, NombreFormula)
17         crSeteoFormula = crResult
18         If crResult = 0 Then
19              crMsgErr = "Ocurrió un error al setear la formula " & Trim(NombreFormula) & " := " & Trim(TextoFormula) & "."
20         End If

' <VB WATCH>
21         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crSeteoFormula"

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

Public Function crAbroEngine() As Integer
' <VB WATCH>
22         On Error GoTo vbwErrHandler
' </VB WATCH>

23         crMsgErr = ""
24         crResult = PEOpenEngine()   'Si retorna Cero dio error.
25         If crResult = 0 Then
26              crMsgErr = "No se pudo abrir el motor de impresión."
27         End If

28         cnUsuario = miConexion.RetornoPropiedad(bUID:=True)
29         cnPass = miConexion.RetornoPropiedad(bPWD:=True)
30         cnDSN = miConexion.RetornoPropiedad(bDSN:=True)
31         cnBD = miConexion.RetornoPropiedad(bDB:=True)

32         crAbroEngine = crResult

' <VB WATCH>
33         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crAbroEngine"

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

Public Function crAbroReporte(Camino As String) As Integer
' <VB WATCH>
34         On Error GoTo vbwErrHandler
' </VB WATCH>

35     Dim Result%, jobnum%, mainjob%, textHandle&, TextLength%
36     Dim TableType As PETableType, LogOnInfo As PELogOnInfo

           ' Define the size of the structures
37         TableType.StructSize = PE_SIZEOF_TABLE_TYPE
38         LogOnInfo.StructSize = PE_SIZEOF_LOGON_INFO

39         crMsgErr = ""
40         Result% = PEOpenPrintJob(Camino)  'Si retorna Cero dio error.
41         jobnum% = Result%
42         If Result% = 0 Then
43              crMsgErr = "Ocurrió un error al iniciar el reporte."
44              Exit Function
45         End If

46         Result% = PEGetNthTableType(jobnum%, 0, TableType)
47         Result% = PEGetNthTableLogOnInfo(jobnum%, 0, LogOnInfo)

           'TableType.DBType = 2
           'TableType.DescriptiveName = "ODBC - SSFF" & Chr(0)
           'TableType.DLLName = "PdSODBC.DLL" & Chr(0)
           'TableType.StructSize = 324 & Chr(0)
48         TableType.DescriptiveName = "ODBC - " & cnDSN & Chr(0)

           ' Get the fields needed for the LogOn Server call from the user, defaulting with the data
49         LogOnInfo.ServerName = cnDSN & Chr$(0)
50         LogOnInfo.DatabaseName = cnBD & Chr$(0)
51         LogOnInfo.UserID = cnUsuario & Chr$(0)
52         LogOnInfo.Password = cnPass & Chr$(0)

53         Result% = crPESetNthTableLogOnInfo(jobnum%, 0, LogOnInfo.ServerName, LogOnInfo.DatabaseName, LogOnInfo.UserID, LogOnInfo.Password, True)
54         If Result% = 0 Then
55              crMsgErr = "Ocurrió un error al setear los valores de conexión."
56              Exit Function
57         End If

           ' Attempt to log on server using parameters
58         Result% = crPELogOnServer(TableType.DLLName, LogOnInfo.ServerName, LogOnInfo.DatabaseName, LogOnInfo.UserID, LogOnInfo.Password)
59         If Result% = 0 Then
60              crMsgErr = "Ocurrió un error al iniciar la conexión para  el reporte."
61              Exit Function
62         End If
63         crAbroReporte = jobnum%

' <VB WATCH>
64         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crAbroReporte"

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

Public Function crObtengoNombreFormula(NroTrabajo As Integer, Posicion As Integer) As String
' <VB WATCH>
65         On Error GoTo vbwErrHandler
' </VB WATCH>
66     Dim NameHandle As Long, NameLength As Integer, textHandle As Long, TextLength As Integer

67         crMsgErr = ""
68         crObtengoNombreFormula = ""     'Si retorna "" dio error.
69         crResult = PEGetNthFormula(NroTrabajo, Posicion, NameHandle, NameLength, textHandle, TextLength)
70         If crResult = 0 Then
71              crMsgErr = "Ocurrió un error al buscar el nombre de las formulas del reporte."
72              Exit Function
73         End If

74         crObtengoNombreFormula = String$(NameLength, 0)

           'Paso el puntero del string y cargo el nombre.
75         crResult = crvbHandleToBStr(crObtengoNombreFormula, NameHandle, NameLength)
76         If crResult = 0 Then
77             crObtengoNombreFormula = ""
78             crMsgErr = "Ocurrió un error al buscar el nombre de las formulas del reporte."
79         End If

' <VB WATCH>
80         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crObtengoNombreFormula"

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

Public Function crSeteoSqlQuery(NroTrabajo As Integer, SQLText As String) As Integer
' <VB WATCH>
81         On Error GoTo vbwErrHandler
' </VB WATCH>
82         crMsgErr = ""
83         crSeteoSqlQuery = PESetSQLQuery(NroTrabajo, SQLText)
84         If crSeteoSqlQuery = 0 Then
85              crMsgErr = "Ocurrió un error al setear la query del reporte."
86         End If
' <VB WATCH>
87         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crSeteoSqlQuery"

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

Public Function crAbroSubreporte(NroTrabajo As Integer, NombreSubReporte As String)
' <VB WATCH>
88         On Error GoTo vbwErrHandler
' </VB WATCH>

89     Dim Result%, jobnum%
90     Dim TableType As PETableType, LogOnInfo As PELogOnInfo

91         crMsgErr = ""
92         crAbroSubreporte = 0

           ' Define the size of the structures
93         TableType.StructSize = PE_SIZEOF_TABLE_TYPE
94         LogOnInfo.StructSize = PE_SIZEOF_LOGON_INFO

           'Retorna el nro. de trabajo del subreporte. Ojo es distinto al del reporte principal.
95         Result% = PEOpenSubreport(NroTrabajo, NombreSubReporte)
96         jobnum% = Result%
97         If Result% = 0 Then
98              crMsgErr = "Ocurrió un error al abrir el subreporte: " & Trim(NombreSubReporte)
99              Exit Function
100        End If

101        Result% = PEGetNthTableType(jobnum%, 0, TableType)
102        Result% = PEGetNthTableLogOnInfo(jobnum%, 0, LogOnInfo)

103        TableType.DescriptiveName = "ODBC - " & cnDSN & Chr(0)

           ' Get the fields needed for the LogOn Server call from the user, defaulting with the data
104        LogOnInfo.ServerName = cnDSN & Chr$(0)
105        LogOnInfo.DatabaseName = cnBD & Chr$(0)
106        LogOnInfo.UserID = cnUsuario & Chr$(0)
107        LogOnInfo.Password = cnPass & Chr$(0)

108        Result% = crPESetNthTableLogOnInfo(jobnum%, 0, LogOnInfo.ServerName, LogOnInfo.DatabaseName, LogOnInfo.UserID, LogOnInfo.Password, True)
109        If Result% = 0 Then
110             crMsgErr = "Ocurrió un error al setear los valores de conexión."
111             Exit Function
112        End If

113        crAbroSubreporte = jobnum%

' <VB WATCH>
114        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crAbroSubreporte"

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

Public Function crMandoAPantalla(NroTrabajo As Integer, Titulo As String) As Integer
' <VB WATCH>
115        On Error GoTo vbwErrHandler
' </VB WATCH>
116        crMsgErr = ""
           'crMandoAPantalla = PEOutputToWindow(NroTrabajo, Titulo, 0, 0, 0, 0, _
                WS_MAXIMIZE + WS_SYSMENU + WS_MINIMIZEBOX + WS_MAXIMIZEBOX + WS_THICKFRAME, 0)

117        crMandoAPantalla = PEOutputToWindow(NroTrabajo, Titulo & Chr$(0), 0, 0, 0, 0, _
           WS_VISIBLE + WS_CAPTION + WS_BORDER + WS_SYSMENU + WS_THICKFRAME + WS_MINIMIZEBOX + WS_MAXIMIZEBOX, 0)

118        If crMandoAPantalla = 0 Then
119             crMsgErr = "Error al redireccionar el reporte a pantalla."
120        End If

121        Dim winBut As PEWindowOptions
122        winBut.StructSize = PE_SIZEOF_WINDOW_OPTIONS
123        crMandoAPantalla = PEGetWindowOptions(NroTrabajo, winBut)
124        If crMandoAPantalla = 0 Then
125             crMsgErr = "Error al obtener la configuración de los controles."
126        End If

127        With winBut
128            .hasPrintSetupButton = 1
129            .hasCloseButton = 1
130            .hasSearchButton = 1
               '.hasGroupTree = 1
               '.canDrillDown = 1
               '.hasRefreshButton = 1
131        End With
132        crMandoAPantalla = PESetWindowOptions(NroTrabajo, winBut)
133        If crMandoAPantalla = 0 Then
134             crMsgErr = "Error al setear la configuración de los controles de impresión."
135        End If

' <VB WATCH>
136        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crMandoAPantalla"

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
Public Function crMandoAImpresora(NroTrabajo As Integer, CantCopias As Integer)
' <VB WATCH>
137        On Error GoTo vbwErrHandler
' </VB WATCH>
138        crMsgErr = ""
139        crMandoAImpresora = PEOutputToPrinter(NroTrabajo, CantCopias)
140        If crMandoAImpresora = 0 Then
141             crMsgErr = "Ocurrió un error al imprimir el reporte. " & vbCr & Err.Description
142        End If
' <VB WATCH>
143        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crMandoAImpresora"

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

Public Function crInicioImpresion(Trabajo As Integer, Espera As Boolean, ProgressDialog As Boolean) As Boolean
' <VB WATCH>
144        On Error GoTo vbwErrHandler
' </VB WATCH>
145        crMsgErr = ""
146        crInicioImpresion = True

147        crEstadoProgressDialog Trabajo, ProgressDialog
148        crResult = PEStartPrintJob(Trabajo, Espera)
149        If crResult = 0 Then
150            crInicioImpresion = False
151            crMsgErr = "Ocurrió un error al iniciar la impresión del reporte."
152        End If

' <VB WATCH>
153        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crInicioImpresion"

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

Public Function crCierroSubReporte(NroSubReporte As Integer) As Boolean
' <VB WATCH>
154        On Error GoTo vbwErrHandler
' </VB WATCH>

155        On Error GoTo errCerrar
156        crMsgErr = ""
157        crCierroSubReporte = True
158        crResult = PECloseSubreport(NroSubReporte)
159        If crResult = 0 Then
160            crCierroSubReporte = False
161            crMsgErr = "Ocurrió un error al cerrar el subreporte."
162        End If

163        Exit Function

164    errCerrar:
165        crCierroSubReporte = False
166        crMsgErr = "Ocurrió un error al cerrar el subreporte."
' <VB WATCH>
167        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crCierroSubReporte"

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
Public Function crEsperoCierreReportePantalla()
' <VB WATCH>
168        On Error GoTo vbwErrHandler
' </VB WATCH>
169    Dim hwndVentana
170        On Error GoTo errEspero
171        hwndVentana = GetActiveWindow()
172        Do While IsWindow(hwndVentana)
173            DoEvents
174        Loop
175    errEspero:
' <VB WATCH>
176        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crEsperoCierreReportePantalla"

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
Public Function crCierroTrabajo(NroTrabajo As Integer) As Boolean
' <VB WATCH>
177        On Error GoTo vbwErrHandler
' </VB WATCH>
178        On Error GoTo errCerrar
179        crMsgErr = ""
180        PEClosePrintJob NroTrabajo
181        crCierroTrabajo = True
182        Exit Function
183    errCerrar:
184        crCierroTrabajo = False
185        crMsgErr = "Ocurrió un error al cerrar el reporte."
' <VB WATCH>
186        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crCierroTrabajo"

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

Public Function crCierroEngine() As Boolean
' <VB WATCH>
187        On Error GoTo vbwErrHandler
' </VB WATCH>
188        On Error GoTo errCerrar
189        crMsgErr = ""
190        PECloseEngine
191        crCierroEngine = True
192        Exit Function
193    errCerrar:
194        crCierroEngine = False
195        crMsgErr = "Ocurrió un error al cerrar el reporte."
' <VB WATCH>
196        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crCierroEngine"

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

Public Function crSeteoImpresora(NroTrabajo As Integer, Impresora As Printer, _
            Optional NroBandeja As Integer = -1, Optional Orientacion As Integer = 1) As Boolean
' <VB WATCH>
197        On Error GoTo vbwErrHandler
' </VB WATCH>
198    On Error GoTo errSetear

199    Dim aModo As crDEVMODE
200    Dim aDriver As String, aDevice As String, aPuerto As String

201        crMsgErr = ""

202        crPEGetSelectedPrinter NroTrabajo, aDriver, aDevice, aPuerto, aModo  'Tomo la impresora del reporte.

203        If NroBandeja <> -1 Then 'Indico la bandeja.
204             aModo.dmDefaultSource = NroBandeja
205        End If

           'DRIVER = HP LaserJet 4050 TN PCL 6 aModo.dmPaperSize = 258
206        aModo.dmOrientation = Orientacion
207        aModo.dmPaperSize = 1
           'Selecciono esta impresora en el reporte.
208        crPESelectPrinter NroTrabajo, Impresora.driverName, Impresora.DeviceName, Impresora.Port, aModo

209        crSeteoImpresora = True
210        Exit Function

211    errSetear:
212        crSeteoImpresora = False
213        crMsgErr = "Ocurrió un error al setear la impresora. " & Err.Description
' <VB WATCH>
214        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crSeteoImpresora"

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

Public Function crObtengoCantidadFormulasEnReporte(NroTrabajo As Integer) As Integer
' <VB WATCH>
215        On Error GoTo vbwErrHandler
' </VB WATCH>
216        crMsgErr = ""
           ' Obtengo la cantidad de formulas que hay en el reporte.
           'Resultados posibles:  -1 Error ó 0...n
217        crObtengoCantidadFormulasEnReporte = PEGetNFormulas(NroTrabajo)
218        If crObtengoCantidadFormulasEnReporte = -1 Then
219             crMsgErr = "Ocurrió un error al obtener la cantidad de formulas."
220        End If

' <VB WATCH>
221        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crObtengoCantidadFormulasEnReporte"

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

Public Function crEstadoProgressDialog(Trabajo As Integer, Estado As Boolean)
' <VB WATCH>
222        On Error GoTo vbwErrHandler
' </VB WATCH>
223        crResult = PEEnableProgressDialog(Trabajo, Estado)
' <VB WATCH>
224        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "crEstadoProgressDialog"

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




