VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTablaRenglon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IndexOrig As Integer
Public PosInit As clsPosXY
'Public PosWidth As String
'Public DataField As String
Private m_Renglones As Collection

Public Sub AddRenglon(ByVal sValue As String)
    m_Renglones.Add sValue
End Sub

Public Function GetRenglon(ByVal iIdx As Integer) As String
    If iIdx > 0 And iIdx <= m_Renglones.Count Then GetRenglon = m_Renglones(iIdx)
End Function

Public Function Count() As Integer
    Count = m_Renglones.Count
End Function

Private Sub Class_Initialize()
    Set m_Renglones = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Renglones = Nothing
End Sub
