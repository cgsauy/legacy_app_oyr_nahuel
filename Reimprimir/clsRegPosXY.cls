VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegPosXY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PosXY As Collection

Public Sub AddPosicion(ByVal oPXY As clsPosXY)
    m_PosXY.Add oPXY
End Sub

Public Function GetPosicion(ByVal iIndex As Byte) As clsPosXY
    Set GetPosicion = m_PosXY(iIndex)
End Function

Public Function Count() As Byte
    Count = m_PosXY.Count
End Function

Private Sub Class_Initialize()
    Set m_PosXY = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_PosXY = Nothing
End Sub
