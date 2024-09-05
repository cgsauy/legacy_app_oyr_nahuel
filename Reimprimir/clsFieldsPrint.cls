VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFieldsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Field As Collection

Public Function Count() As Integer
    Count = m_Field.Count
End Function

Public Sub AddField(ByVal oField As clsRegFieldPrint)
    m_Field.Add oField
End Sub

Public Function GetField(ByVal iIdx As Integer) As clsRegFieldPrint
    Set GetField = m_Field(iIdx)
End Function

Public Function FindField(ByVal iTipo As Integer, sNombre As String) As clsRegFieldPrint
Dim iQ As Integer
Dim oF As clsRegFieldPrint
    For iQ = 1 To m_Field.Count
        Set oF = m_Field(iQ)
        If oF.Tipo = iTipo And LCase(oF.Nombre) = LCase(sNombre) Then
            Set FindField = oF
            Exit Function
        End If
    Next
End Function

Private Sub Class_Initialize()
    Set m_Field = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Field = Nothing
End Sub
