Attribute VB_Name = "vbwProtector"
 ' vbwProtector.bas file - Location: \VB Watch 2\Templates\VB6\Protector\ '
'                                                                        '
' This module contains all procedures common to the VB Watch tools.      '
' It will be added to every project instrumented with VB Watch.          '
'                                                                        '
' ************************* WARNING *******************************      '
' You should not modify it unles you know what you are doing.            '
' To modify it, remove the read-only attribute of vbwProtector.bas.      '
 ' WARNING: modifications of this file will apply to all error handling   '
'          plans !!!                                                     '

Option Explicit

' Options '
Public vbwCatchException As Boolean
Public vbwTraceProc As Boolean
Public vbwTraceParameters As Boolean
Public vbwTraceLine As Boolean
Public vbwCallStack As Boolean
Public vbwEmailRecipientAdress As String
Public vbwDumpStringMaxLength As Long
Public vbwSystemInfo As Boolean
Public vbwScreenshot As Boolean

' Variables for use with vbwFunctions.dll '
Public vbwAdvancedFunctions As Object          ' this will be used only if vbwFunctions.dll is installed on the enduser machine '
Public fIsVbwFunctionsInitialized As Boolean   ' true if vbwFunctions.dll is installed and instanciated                         '

' Call Stack '
Public vbwStackCalls() As String     ' array containing each call of the stack '
Public vbwStackCallsNumber As Long   ' number of calls = Ubound(vbwStackCalls) '

' Trace '
Public vbwTraceCallsNumber As Long   ' number of calls '

' Log File
Dim fIsLogInitialize As Boolean
Public vbwLogFile As String
Public vbwLogTraceToFile As Boolean
Dim fLogFileOpen As Boolean
Dim lLogFileNumber As Long
Dim lLogFileOffset As Long
' file I/O
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const OPEN_ALWAYS = 4
Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        OffSet As Long
        OffsetHigh As Long
        hEvent As Long
End Type
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFileEx Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long

' Var Dump
Const VBW_STRING = "**************************"
Global Const VBW_LOCAL_STRING = vbCrLf & VBW_STRING & vbCrLf & "* LOCAL LEVEL VARIABLES  *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_MODULE_STRING = vbCrLf & VBW_STRING & vbCrLf & "* MODULE LEVEL VARIABLES *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_GLOBAL_STRING = vbCrLf & VBW_STRING & vbCrLf & "* GLOBAL LEVEL VARIABLES *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_TYPE_STRING = " (User Defined Type Array)"
Global Const VBW_UNKNOWN_STRING = " = {Unknown Type}"
Global Const VBW_LOCAL_NOT_REPORTED = "Local Variables: not reported"
Global Const VBW_MODULE_NOT_REPORTED = "Module Variables: not reported"
Global Const VBW_GLOBAL_NOT_REPORTED = "Global Variables: not reported"
Global Const VBW_NO_LOCAL_VARIABLES = "No Local Variables"
Global vbwDumpFile As String
Global vbwDumpFileNum As Long

' Thread & processes
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

' Exception handling declarations
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const EXCEPTION_CONTINUE_EXECUTION = -1
Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15
Private Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long    ' Pointer to an EXCEPTION_RECORD structure
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type
Private Type EXCEPTION_DEBUG_INFO
        pExceptionRecord As EXCEPTION_RECORD
        dwFirstChance As Long
End Type
Private Type CONTEXT
    dblVar(66) As Double ' The real structure is more complex
    lngVar(6) As Long    ' but we don't need those details
End Type
Private Type EXCEPTION_POINTERS
    pExceptionRecord As EXCEPTION_RECORD
    ContextRecord As CONTEXT
End Type
Private Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
Private Const EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
Private Const EXCEPTION_BREAKPOINT = &H80000003
Private Const EXCEPTION_SINGLE_STEP = &H80000004
Private Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Private Const EXCEPTION_FLT_DENORMAL_OPERAND = &HC000008D
Private Const EXCEPTION_FLT_DIVIDE_BY_ZERO = &HC000008E
Private Const EXCEPTION_FLT_INEXACT_RESULT = &HC000008F
Private Const EXCEPTION_FLT_INVALID_OPERATION = &HC0000090
Private Const EXCEPTION_FLT_OVERFLOW = &HC0000091
Private Const EXCEPTION_FLT_STACK_CHECK = &HC0000092
Private Const EXCEPTION_FLT_UNDERFLOW = &HC0000093
Private Const EXCEPTION_INT_DIVIDE_BY_ZERO = &HC0000094
Private Const EXCEPTION_INT_OVERFLOW = &HC0000095
Private Const EXCEPTION_PRIV_INSTRUCTION = &HC0000096
Private Const EXCEPTION_IN_PAGE_ERROR = &HC0000006
Private Const EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
Private Const EXCEPTION_NONCONTINUABLE_EXCEPTION = &HC0000025
Private Const EXCEPTION_STACK_OVERFLOW = &HC00000FD
Private Const EXCEPTION_INVALID_DISPOSITION = &HC0000026
Private Const EXCEPTION_GUARD_PAGE = &H80000001
Private Const EXCEPTION_INVALID_HANDLE = &HC0000008
Private Const CONTROL_C_EXIT = &HC000013A

' Variable to Save the Err object
Dim ErrObjectDescription As String
Dim ErrObjectHelpContext As Long
Dim ErrObjectHelpFile As String
Dim ErrObjectLastDllError As Long
Dim ErrObjectNumber As Long
Dim ErrObjectSource As String
Dim ErrLine As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public VBWPROTECTOR_EMPTY As Variant ' for use with vbwExecuteLine() in IIf structures

'Const VBW_EXE_EXTENSION = ".exe" ' this line will be rewritten by VB Watch with the right extension

' vbwNoTraceProc vbwNoTraceLine ' don't remove this !

#Const PROJECT = "prjPrecios.vbp"
' <VB WATCH>
Const VBWMODULE = "vbwProtector"
Global Const VBWPROJECT = "prjPrecios"
Global Const VBW_EXE_EXTENSION = ".exe"
' </VB WATCH>

Sub vbwInitializeProtector()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>

2          Static vbwIsInitialized As Boolean

3          If vbwIsInitialized Then
4              Exit Sub
5          End If

       ' Don't remove the following comments !                                                         '
       ' VB Watch will replace next line with the initialization code as set in the plan being applied '
6      vbwCatchException = True


7          vbwLogTraceToFile = vbwTraceProc Or vbwTraceLine
8          If vbwCallStack Then ' needed to track call stack
9               vbwTraceProc = True
10         End If

11         vbwLogFile = App.Path & "\vbw" & App.EXEName & VBW_EXE_EXTENSION & ".log"
12         vbwDumpFile = App.Path & "\vbw" & App.EXEName & VBW_EXE_EXTENSION & ".dmp"

13         If vbwCatchException Then
14             vbwHandleException
15         End If

16         vbwDumpStringMaxLength = 128 ' change this value to suit your need - make it 0 to remove the size check (to use with caution)

17         vbwIsInitialized = True

' <VB WATCH>
18         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwInitializeProtector"

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

Sub vbwReportVariable(ByVal lName As String, ByVal lValue As Variant, Optional ByVal lTab As Long)
       ' vbwNoErrorHandler ' don't remove this !
19         Dim i As Long, j As Long, k As Long, L As Long
20         Dim tDim As Long

21         On Error GoTo ErrDump

22         If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
23             tDim = GetArrayDimension(lValue)
24             Select Case tDim
                   Case 1
25                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & ") As " & TypeName(lValue))
26                     For i = LBound(lValue) To UBound(lValue)
27                         vbwReportVariable lName & "(" & i & ")", lValue(i), lTab + 1
28                     Next i
29                 Case 2
30                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & ") As " & TypeName(lValue))
31                     For j = LBound(lValue, 2) To UBound(lValue, 2)
32                         For i = LBound(lValue, 1) To UBound(lValue, 1)
33                             vbwReportVariable lName & "(" & i & "," & j & ")", lValue(i, j), lTab + 1
34                         Next i
35                     Next j
36                 Case 3
37                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & "," & LBound(lValue, 3) & " To " & UBound(lValue, 3) & ") As " & TypeName(lValue))
38                     For k = LBound(lValue, 3) To UBound(lValue, 3)
39                         For j = LBound(lValue, 2) To UBound(lValue, 2)
40                             For i = LBound(lValue, 1) To UBound(lValue, 1)
41                                 vbwReportVariable lName & "(" & i & "," & j & "," & k & ")", lValue(i, j, k), lTab + 1
42                             Next i
43                         Next j
44                     Next k
45                 Case 4
46                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & "," & LBound(lValue, 3) & " To " & UBound(lValue, 3) & "," & LBound(lValue, 4) & " To " & UBound(lValue, 4) & ") As " & TypeName(lValue))
47                     For L = LBound(lValue, 4) To UBound(lValue, 4)
48                         For k = LBound(lValue, 3) To UBound(lValue, 3)
49                             For j = LBound(lValue, 2) To UBound(lValue, 2)
50                                 For i = LBound(lValue, 1) To UBound(lValue, 1)
51                                     vbwReportVariable lName & "(" & i & "," & j & "," & k & "," & L & ")", lValue(i, j, k, L), lTab + 1
52                                 Next i
53                             Next j
54                         Next k
55                     Next L
56                 Case Else
57                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "() not processed: " & tDim & " dimensions")
58             End Select
59         Else
               ' non-array '
60             If IsObject(lValue) Then
61                 vbwReportObject lName, lValue, lTab
62             Else
63                 If VarType(lValue) = vbString Then
64                     lValue = FormatString(lValue)
65                 End If
66                 vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & " = " & lValue & " (" & TypeName(lValue) & ")")
67             End If
68         End If
69         Exit Sub

70     ErrDump:
71         Err.Clear
72         vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & " = {Variable Dumping Error}")
End Sub

Public Sub vbwReportObject(lName As String, ByVal lObject As Object, Optional ByVal lTab As Long)
       ' vbwNoErrorHandler ' don't remove this !

73         On Error GoTo ErrDump

74         If TypeName(lObject) <> "ErrObject" Then
75             If fIsVbwFunctionsInitialized Then
                   ' this should be executed only if you are using a global error handler '
                   ' that prepares properly the vbwAdvancedFunctions for object dumping   '
76                 vbwCloseDumpFile       ' close it because vbwAdvancedFunctions uses its own file writing routines '
77                 vbwAdvancedFunctions.ReportObject lName, lObject, lTab, TypeOf lObject Is Form, TypeOf lObject Is MDIForm
78                 vbwOpenDumpFile
79             Else
                   ' no vbwFunctions.dll available                         '
                   ' only report the default value of objects and controls '
80                 If TypeOf lObject Is Form Or TypeOf lObject Is MDIForm Then
81                    On Error Resume Next
82                    vbwReportToFile vbwEncryptString("Form " & lName)
83                    Dim c As Control
84                    For Each c In lObject.Controls
85                        vbwReportObject c.Name & vbwGetIndex(c), c, 1
86                    Next c
87                 Else
88                     If IsNumeric(lObject) Then
89                         vbwReportVariable lName, CDbl(lObject), lTab
90                     Else
91                         vbwReportVariable lName, CStr(lObject), lTab
92                     End If
93                 End If
94             End If
95         Else
96             vbwReportToFile vbCrLf & vbwEncryptString("**** ErrObject Err ****")
97             vbwReportVariable "Err.Number", ErrObjectNumber
98             vbwReportVariable "Err.Source", ErrObjectSource
99             vbwReportVariable "Err.Description", ErrObjectDescription
100            vbwReportVariable "Err.HelpContext", ErrObjectHelpContext
101            vbwReportVariable "Err.HelpFile", ErrObjectHelpFile
102            If ErrObjectLastDllError = 0 Then
103                vbwReportVariable "Err.LastDllError", ErrObjectLastDllError
104            Else
                   ' get the API error description from the system
105                Dim sBuffer As String * 512
106                Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
107                FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, Null, ErrObjectLastDllError, 0, sBuffer, 512, 0
108                If InStr(sBuffer, Chr(0)) Then
109                    vbwReportVariable "Err.LastDllError", ErrObjectLastDllError & " (" & Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1) & ")"
110                Else
111                    vbwReportVariable "Err.LastDllError", ErrObjectLastDllError
112                End If
113            End If
114        End If

115        Exit Sub

116    ErrDump:
117        Err.Clear
118        vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & ".Value = {No Value Property}")
End Sub

Public Function vbwEncryptString(ByRef sString As String, Optional sKey) As String
       ' vbwNoErrorHandler ' don't remove this ! '

119        On Error Resume Next
120        If fIsVbwFunctionsInitialized = False Then
               ' no encryption without vbwFunctions.dll            '
               ' you may want to write your own encryption routine '
121            vbwEncryptString = sString
122        Else
123            If IsMissing(sKey) Then
                   ' If you filled the vardump encryption key in the VB Watch Options, your key will   '
                   ' be already embeded in the vbwAdvancedFunctions.ObjectInfo property, so you do not '
                   ' have to care about provideing a key                                               '
124                vbwEncryptString = vbwAdvancedFunctions.EncryptString(sString)
125            Else
                   ' Yet if you wish to overide  your default encryption key, simply pass it '
                   ' in the sKey parameter                                                   '
126                vbwEncryptString = vbwAdvancedFunctions.EncryptString(sString, sKey)
127            End If
128        End If
End Function

Function vbwReportParameter(ByVal lName As String, ByRef lValue As Variant) As String
       ' vbwNoErrorHandler ' don't remove this !
129        Dim i As Long, j As Long, k As Long
130        Dim tDim As Long
131        Dim retString As String

132        On Error GoTo ErrDump

133        If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
134            tDim = GetArrayDimension(lValue)
135            If tDim Then
136                retString = lName & "("
137                For i = 1 To tDim
138                    retString = retString & LBound(lValue, i) & " To " & UBound(lValue, i) & ","
139                Next i
140                Mid$(retString, Len(retString)) = ")"   ' Close the brackets by overwriting the last comma '
141            Else
142                retString = lName & "(Undimensioned Array)"
143            End If
144        Else
               ' non-array '
145            If IsObject(lValue) Then
                   ' object
146                On Error Resume Next
147                retString = TypeName(lValue) & " " & lName & " = " & CStr(lValue)
148                If Err.Number Then
149                    On Error GoTo ErrDump
150                    retString = TypeName(lValue) & " " & lName & " = " & lValue.Name & vbwGetIndex(lValue)
151                End If
152            Else
                   ' non-object
153                If VarType(lValue) = vbString Then
154                   retString = lName & " = " & FormatString(lValue)
155                Else
156                   retString = lName & " = " & lValue
157                End If
158            End If
159        End If

160        vbwReportParameter = retString
161        Exit Function

162    ErrDump:
163        Err.Clear
164        vbwReportParameter = lName & " = {" & TypeName(lValue) & ": Parameter Dumping Error}"
End Function

Function vbwReportParameterByVal(ByVal lName As String, ByVal lValue As Variant) As String
       ' vbwNoErrorHandler ' don't remove this !
165        Dim i As Long, j As Long, k As Long
166        Dim tDim As Long
167        Dim retString As String

168        On Error GoTo ErrDump

169        If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
170            tDim = GetArrayDimension(lValue)
171            If tDim Then
172                retString = lName & "("
173                For i = 1 To tDim
174                    retString = retString & LBound(lValue, i) & " To " & UBound(lValue, i) & ","
175                Next i
176                Mid$(retString, Len(retString)) = ")"   ' Close the brackets by overwriting the last comma '
177            Else
178                retString = lName & "(Undimensioned Array)"
179            End If
180        Else
               ' non-array '
181            If IsObject(lValue) Then
                   ' object
182                On Error Resume Next
183                retString = TypeName(lValue) & " " & lName & " = " & CStr(lValue)
184                If Err.Number Then
185                    On Error GoTo ErrDump
186                    retString = TypeName(lValue) & " " & lName & " = " & lValue.Name & vbwGetIndex(lValue)
187                End If
188            Else
                   ' non-object
189                If VarType(lValue) = vbString Then
190                   retString = lName & " = " & FormatString(lValue)
191                Else
192                   retString = lName & " = " & lValue
193                End If
194            End If
195        End If

196        vbwReportParameterByVal = retString
197        Exit Function

198    ErrDump:
199        Err.Clear
200        vbwReportParameterByVal = lName & " = {" & TypeName(lValue) & ": Parameter Dumping Error}"
End Function

Sub vbwReportToFile(ByRef lString As String)
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
201        On Error GoTo vbwErrHandler
' </VB WATCH>
202         If vbwDumpFileNum = 0 Then
203              vbwOpenDumpFile
204         End If
205         On Error Resume Next
206         Print #vbwDumpFileNum, lString
207         If Err = 52 Then
208            vbwCloseDumpFile
209            vbwOpenDumpFile
210            Print #vbwDumpFileNum, lString
211         End If
' <VB WATCH>
212        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwReportToFile"

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

Sub vbwOpenDumpFile()
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
213        On Error GoTo vbwErrHandler
' </VB WATCH>
214       vbwDumpFileNum = FreeFile
215       Open vbwDumpFile For Append As #vbwDumpFileNum
' <VB WATCH>
216        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwOpenDumpFile"

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

Sub vbwCloseDumpFile()
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
217        On Error GoTo vbwErrHandler
' </VB WATCH>
218       Close #vbwDumpFileNum
' <VB WATCH>
219        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwCloseDumpFile"

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

Private Function GetArrayDimension(ByRef arg As Variant) As Long
       ' vbwNoErrorHandler ' don't remove this !
220        Dim i As Long, j As Long
221        On Error Resume Next
222        i = 0
223        Do
224            i = i + 1
225            j = LBound(arg, i)
226        Loop Until Err.Number
227        GetArrayDimension = i - 1
End Function

Function vbwGetIndex(tObject As Variant) As String
       ' vbwNoErrorHandler ' don't remove this !
228        On Error Resume Next
229        vbwGetIndex = "(" & tObject.Index & ")"
End Function

Private Function FormatString(ByVal arg As String) As String
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
230        On Error GoTo vbwErrHandler
' </VB WATCH>

231        If Right$(arg, 1) = "}" Then ' probably a VB Watch built-in message
232             FormatString = arg
233             Exit Function
234        End If

           ' 1. truncate according to the vbwDumpStringMaxLength value
235        If vbwDumpStringMaxLength Then
236            If Len(arg) > vbwDumpStringMaxLength Then
237                arg = Left$(arg, vbwDumpStringMaxLength + 1)   ' +1: avoids to cut inside a vbCrLf '
238                If Right$(arg, 2) = vbCrLf Then
                       ' don't cut inside a vbCrLf
239                Else
240                    arg = Left$(arg, vbwDumpStringMaxLength)
241                End If
242                arg = arg & "{...}" ' truncated
243            End If
244        End If

           ' 2. make sure string isn't multiline
245        arg = Replace(arg, vbCrLf, "<CrLf>", , , vbBinaryCompare)
246        arg = Replace(arg, Chr(13), "<Cr>", , , vbBinaryCompare)
247        arg = Replace(arg, Chr(10), "<Lf>", , , vbBinaryCompare)

           ' 3. add quotes
248        FormatString = Chr(34) & arg & Chr(34)
' <VB WATCH>
249        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FormatString"

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

Sub vbwProcIn(ByRef lProc As String, Optional ByRef lParameters As String)
' <VB WATCH>
250        On Error GoTo vbwErrHandler
' </VB WATCH>

251        vbwTraceCallsNumber = vbwTraceCallsNumber + 1

252        vbwStackCallsNumber = vbwStackCallsNumber + 1
253        ReDim Preserve vbwStackCalls(1 To vbwStackCallsNumber)
254        vbwStackCalls(vbwStackCallsNumber) = lProc

255        Dim lString As String
256        lString = String$(vbwTraceCallsNumber - 1, vbTab) & lProc

257        If vbwLogTraceToFile Then
258             vbwSendLog lString & lParameters
259        End If

' <VB WATCH>
260        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwProcIn"

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

Sub vbwProcOut(ByRef lProc As String)
' <VB WATCH>
261        On Error GoTo vbwErrHandler
' </VB WATCH>

262        If vbwTraceCallsNumber > 0 Then ' should always be true
263           vbwTraceCallsNumber = vbwTraceCallsNumber - 1
264        End If

265        If vbwStackCallsNumber > 0 Then ' should always be true
266           vbwStackCallsNumber = vbwStackCallsNumber - 1
267        End If

' <VB WATCH>
268        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwProcOut"

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


Function vbwExecuteLine(ByRef fEncrypted As String, ByRef lLine As String) As Boolean
' <VB WATCH>
269        On Error GoTo vbwErrHandler
' </VB WATCH>

270        If vbwTraceLine Then

271            If fEncrypted Then
272                lLine = "<CRY>" & lLine & "</CRY>"
273            End If

274            If vbwLogTraceToFile Then
275                If vbwTraceCallsNumber > 0 Then
276                    vbwSendLog String$(vbwTraceCallsNumber - 1, vbTab) & " -> " & lLine
277                Else
278                    vbwSendLog " -> " & lLine
279                End If
280            End If

281        End If

           ' This function always returns false
' <VB WATCH>
282        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwExecuteLine"

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

Function vbwGetStack() As String
' <VB WATCH>
283        On Error GoTo vbwErrHandler
' </VB WATCH>

284        If vbwTraceProc = False Then
285            vbwGetStack = "{Unavailable}"
286            Exit Function
287        End If

288        Dim vbwStackString As String
289        Dim i As Long

290        For i = vbwStackCallsNumber To 1 Step -1
291            vbwStackString = vbwStackString & String$(i - 1, vbTab) & vbwStackCalls(i) & vbCrLf
292        Next i
293        vbwGetStack = IIf(vbwStackString <> "", vbwStackString, "{Empty}")
' <VB WATCH>
294        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwGetStack"

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

Sub vbwSendLog(ByRef tMsg As String)
' <VB WATCH>
295        On Error GoTo vbwErrHandler
' </VB WATCH>
296        If Err.Number Then
               ' Save Err object before being cleared by "On Error Resume Next"
297            Dim ErrDescription As String, ErrHelpFile As String, ErrSource As String
298            Dim ErrHelpContext As Long, ErrNumber As Long
299            ErrDescription = Err.Description
300            ErrHelpContext = Err.HelpContext
301            ErrHelpFile = Err.HelpFile
302            ErrNumber = Err.Number
303            ErrSource = Err.Source
304        End If

305        On Error Resume Next

306        If Not fLogFileOpen Then
307            fLogFileOpen = True
308            Dim suffix As Long
309            Do
310                Kill vbwLogFile
311                lLogFileNumber = CreateFile(vbwLogFile, GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_ALWAYS, FILE_FLAG_OVERLAPPED, 0)
312                If lLogFileNumber < 0 Then
                       ' under some circumstances (retained in memory applications or while in the IDE)
                       ' the previous log file might not have been freed yet, so we must use another one
313                    suffix = suffix + 1
314                    vbwLogFile = App.Path & "\vbw" & App.EXEName & VBW_EXE_EXTENSION & suffix & ".log"
315                End If
316            Loop Until lLogFileNumber >= 0 Or suffix > 1000
317        End If

318        If Not fIsLogInitialize Then
               ' init file '
319            fIsLogInitialize = True
320            WriteToLogFile "Tracing " & App.Title
321            WriteToLogFile "Session started " & Now
322            WriteToLogFile ""
323        End If

           ' log to file
324        WriteToLogFile tMsg

325       If ErrNumber Then
               ' Restore Err object if cleared by "On Error Resume Next"
326            Err.Description = ErrDescription
327            Err.HelpContext = ErrHelpContext
328            Err.HelpFile = ErrHelpFile
329            Err.Number = ErrNumber
330            Err.Source = ErrSource
331        End If

' <VB WATCH>
332        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwSendLog"

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

' Writes Str as a new line in the log file (adding a vbCrLf to the end)
Private Function WriteToLogFile(Str As String) As Long
' <VB WATCH>
333        On Error GoTo vbwErrHandler
' </VB WATCH>
334        Dim ol As OVERLAPPED
335        Dim bBytes() As Byte, StrLength As Long
336        StrLength = Len(Str) + 2
337        ReDim bBytes(0 To StrLength - 1)
338        CopyMemory bBytes(0), ByVal Str & vbCrLf, StrLength
339        ol.OffSet = lLogFileOffset
340        WriteToLogFile = WriteFileEx(lLogFileNumber, bBytes(0), StrLength, ol, ByVal 0&)
341        lLogFileOffset = lLogFileOffset + StrLength
' <VB WATCH>
342        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WriteToLogFile"

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


' Ends a component's thread. If this was the last active thread, ends the component's process.
Public Sub vbwExitThread()
' <VB WATCH>
343        On Error GoTo vbwErrHandler
' </VB WATCH>
344        If vbwIsInIDE Then
               ' Executing ExitThread within the IDE will terminate VB without ceremony !
345            Stop ' Press the End button now
346        Else
347            Dim lpExitCode As Long
348            If GetExitCodeThread(GetCurrentThread(), lpExitCode) Then
349                ExitThread lpExitCode
350            End If
351        End If
' <VB WATCH>
352        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwExitThread"

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

' Ends a component's process. Equivalent to the End statement.
Public Sub vbwExitProcess()
' <VB WATCH>
353        On Error GoTo vbwErrHandler
' </VB WATCH>
354        If vbwIsInIDE Then
               ' Executing ExitProcess within the IDE will terminate VB without ceremony !
355            Stop ' Press the End button now
356        Else
357            Dim lpExitCode As Long
358            If GetExitCodeProcess(GetCurrentProcess(), lpExitCode) Then
359                ExitProcess lpExitCode
360            End If
361        End If
' <VB WATCH>
362        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwExitProcess"

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

' determines if the program is running in the IDE or an EXE File
Private Function vbwIsInIDE() As Boolean
' <VB WATCH>
363        On Error GoTo vbwErrHandler
' </VB WATCH>

364        Dim strFileName As String
365        Dim lngCount As Long

366        strFileName = String(255, 0)
367        lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
368        strFileName = Left(strFileName, lngCount)

369        vbwIsInIDE = UCase$(Right$(strFileName, 8)) Like "\VB#.EXE"

' <VB WATCH>
370        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwIsInIDE"

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

' Exception handling stuff
Public Sub vbwHandleException()
           ' Exceptions will be caught and redirected to the failing procedure
' <VB WATCH>
371        On Error GoTo vbwErrHandler
' </VB WATCH>
372        SetUnhandledExceptionFilter AddressOf vbwExceptionFilter
' <VB WATCH>
373        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwHandleException"

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

' Exception handling stuff
Public Sub vbwUnHandleException()
           ' Exceptions are no longer caught and will cause Exceptions
           ' Whenever possible, call this procedure before returning to the VB's IDE
' <VB WATCH>
374        On Error GoTo vbwErrHandler
' </VB WATCH>
375        SetUnhandledExceptionFilter 0
' <VB WATCH>
376        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwUnHandleException"

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

' Exception handling stuff
Public Function vbwExceptionFilter(ByRef pExceptionInfo As EXCEPTION_POINTERS) As Long
       'vbwNoErrorHandler ' DO NOT remove this !!!

377        Dim ExceptionRecord As EXCEPTION_RECORD
378        ExceptionRecord = pExceptionInfo.pExceptionRecord

379        Do While ExceptionRecord.pExceptionRecord ' Empties the exceptions stack
380            CopyMemory ExceptionRecord, ByVal ExceptionRecord.pExceptionRecord, Len(ExceptionRecord)
381        Loop

382        vbwExceptionFilter = EXCEPTION_CONTINUE_EXECUTION

       'vbwExitProc ' because the next instruction causes to exit the function ' ' DO NOT remove this !!!

           ' Convert the exception to a normal VB error and go back to the failing procedure '
383        Err.Raise 65535, , ExceptionDescription(ExceptionRecord.ExceptionCode)

End Function

' Exception handling stuff
Private Function ExceptionDescription(ByVal ExceptionCode As Long) As String
       ' vbwNoErrorHandler ' don't remove this !
384        Select Case ExceptionCode
               Case EXCEPTION_ACCESS_VIOLATION
385                ExceptionDescription = "Exception: Access Violation"
386            Case EXCEPTION_DATATYPE_MISALIGNMENT
387                ExceptionDescription = "Exception: Datatype Misalignment"
388            Case EXCEPTION_BREAKPOINT
389                ExceptionDescription = "Exception: Breakpoint"
390            Case EXCEPTION_SINGLE_STEP
391                ExceptionDescription = "Exception: Single Step"
392            Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
393                ExceptionDescription = "Exception: Array Bounds Exceeded"
394            Case EXCEPTION_FLT_DENORMAL_OPERAND
395                ExceptionDescription = "Exception: Float Denormal Operand"
396            Case EXCEPTION_FLT_DIVIDE_BY_ZERO
397                ExceptionDescription = "Exception: Float Divide By Zero"
398            Case EXCEPTION_FLT_INEXACT_RESULT
399                ExceptionDescription = "Exception: Float Inexact Result"
400            Case EXCEPTION_FLT_INVALID_OPERATION
401                ExceptionDescription = "Exception: Float Invalid Operation"
402            Case EXCEPTION_FLT_OVERFLOW
403                ExceptionDescription = "Exception: Float Overflow"
404            Case EXCEPTION_FLT_STACK_CHECK
405                ExceptionDescription = "Exception: Float Stack Check"
406            Case EXCEPTION_FLT_UNDERFLOW
407                ExceptionDescription = "Exception: Float Underflow"
408            Case EXCEPTION_INT_DIVIDE_BY_ZERO
409                ExceptionDescription = "Exception: Integer Divide By Zero"
410            Case EXCEPTION_INT_OVERFLOW
411                ExceptionDescription = "Exception: Integer Overflow"
412            Case EXCEPTION_PRIV_INSTRUCTION
413                ExceptionDescription = "Exception: Priv Instruction"
414            Case EXCEPTION_IN_PAGE_ERROR
415                ExceptionDescription = "Exception: In Page Error"
416            Case EXCEPTION_ILLEGAL_INSTRUCTION
417                ExceptionDescription = "Exception: Illegal Instruction"
418            Case EXCEPTION_NONCONTINUABLE_EXCEPTION
419                ExceptionDescription = "Exception: Non Continuable Exception"
420            Case EXCEPTION_STACK_OVERFLOW
421                ExceptionDescription = "Exception: Stack Overflow"
422            Case EXCEPTION_INVALID_DISPOSITION
423                ExceptionDescription = "Exception: Invalid Disposition"
424            Case EXCEPTION_GUARD_PAGE
425                ExceptionDescription = "Exception: Guard Page"
426            Case EXCEPTION_INVALID_HANDLE
427                ExceptionDescription = "Exception: Invalid Handle"
428            Case CONTROL_C_EXIT
429                ExceptionDescription = "Exception: Control C Exit"
430            Case Else
431                ExceptionDescription = "Unknown Exception"
432        End Select

End Function

Public Sub vbwSaveErrObject()
       ' vbwNoErrorHandler ' don't remove this !
433        ErrObjectDescription = Err.Description
434        ErrObjectHelpContext = Err.HelpContext
435        ErrObjectHelpFile = Err.HelpFile
436        ErrObjectLastDllError = Err.LastDllError
437        ErrObjectNumber = Err.Number
438        ErrObjectSource = Err.Source
End Sub

Public Sub vbwRestoreErrObject()
       ' vbwNoErrorHandler ' don't remove this !
439       Err.Description = ErrObjectDescription
440       Err.HelpContext = ErrObjectHelpContext
441       Err.HelpFile = ErrObjectHelpFile
442       Err.Number = ErrObjectNumber
443       Err.Source = ErrObjectSource
End Sub


