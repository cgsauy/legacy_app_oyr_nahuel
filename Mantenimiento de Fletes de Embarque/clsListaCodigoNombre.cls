VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListaCodigoNombre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Lista As Collection

Private Sub Class_Initialize()
    Set Lista = New Collection
End Sub

Private Sub Class_Terminate()
    Set Lista = Nothing
End Sub

Public Function BuscarElemento(ByVal codigo As Long) As clsCodigoNombre
Dim oElem As clsCodigoNombre
    Set BuscarElemento = Nothing
    For Each oElem In Lista
        If oElem.codigo = codigo Then
            Set BuscarElemento = oElem
            Exit Function
        End If
    Next
End Function

Public Function BuscarTexto(ByVal texto As String) As Integer
Dim oElem As clsCodigoNombre
    BuscarTexto = 0
    'For Each oElem In Lista
    Dim iE As Integer
    For iE = 1 To Lista.Count
        Set oElem = Lista.Item(iE)
        If oElem.Nombre = Trim(texto) Then
            BuscarTexto = iE
            Exit Function
        ElseIf oElem.Nombre & "," = Trim(texto) Then
            BuscarTexto = iE
            Exit Function
        End If
    Next
End Function

Public Function ListaElementosAsignado() As String
Dim oElem As clsCodigoNombre
    For Each oElem In Lista
        ListaElementosAsignado = ListaElementosAsignado & IIf(ListaElementosAsignado <> "", ", ", "") & Trim(oElem.Nombre)
    Next
End Function

Public Function ListaCodigos() As String
Dim oElem As clsCodigoNombre
    
    For Each oElem In Lista
        ListaCodigos = ListaCodigos & IIf(ListaCodigos <> "", ", ", "") & Trim(oElem.codigo)
    Next

End Function
