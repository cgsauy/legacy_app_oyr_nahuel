VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVBLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public FileLog As String

Public Enum TipoLog
    TL_Info = 1
    TL_Debug = 2
    TL_Error = 9
End Enum

Private Function TipoLog_ToString(ByVal tipo As TipoLog) As String
    Select Case tipo
        Case TL_Debug
            TipoLog_ToString = "Debug"
        Case TL_Info
            TipoLog_ToString = "Info"
        Case TL_Error
            TipoLog_ToString = "Error"
    End Select
End Function

Public Sub InsertoLog(ByVal tipo As TipoLog, ByVal sLog As String)
On Error Resume Next
    If FileLen(FileLog) > 80000 Then
        Open FileLog For Output As #1
    Else
        Open FileLog For Append As #1
    End If
    Print #1, Format(Now, "dd/mm hh:nn") & vbTab & TipoLog_ToString(tipo) & vbTab & Trim(sLog)
    'Print #1, String(30, "-")
    Close #1
End Sub

