VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatosPrecioFlete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Origen As clsCodigoNombre
Public Destino As clsCodigoNombre
Public FechaEmbarque As Date
Public SiEmbarco As Boolean
Public Agencia As clsCodigoNombre
Public Linea As clsCodigoNombre
Public Contenedores As Collection

Public Function ContenedoresALineaTexto() As String
    Dim oContenedor As clsContenedoresEmbarque
    For Each oContenedor In Contenedores
        ContenedoresALineaTexto = ContenedoresALineaTexto & IIf(ContenedoresALineaTexto = "", "", ", ") & Trim(oContenedor.Contenedor.Nombre)
    Next
End Function

Public Function ContenedoresIDs()
    Dim oContenedor As clsContenedoresEmbarque
    For Each oContenedor In Contenedores
        ContenedoresIDs = ContenedoresIDs & IIf(ContenedoresIDs = "", "", ",") & Trim(oContenedor.Contenedor.Codigo)
    Next
End Function

Public Function Contenedor(ByVal idCon As Long) As clsContenedoresEmbarque
    Dim oCEm As clsContenedoresEmbarque
    For Each oCEm In Contenedores
        If oCEm.Contenedor.Codigo = idCon Then
            Set Contenedor = oCEm
            Exit Function
        End If
    Next
End Function

Public Sub AgregarContenedor(ByVal Contenedor As clsContenedoresEmbarque)
    Contenedores.Add Contenedor
End Sub

Private Sub Class_Initialize()
    Set Contenedores = New Collection
End Sub

Private Sub Class_Terminate()
    Set Contenedores = Nothing
End Sub


