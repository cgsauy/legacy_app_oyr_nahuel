VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRenglonIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public Tipo As Byte                    '  1=entrada, 2 = salida
Public Articulo As Long                 '  ID del artículo
Public ArticuloNombre As String
Public TipoNroSerie As Byte
Public NroSerie As String
Public NroSerieArtCambio As String      'Es el nro. de serie del artículo que se lleva en el cambio.
Public EsRoto As Boolean
Public Estados As String
Public Servicio As clsServicio
Public FichaDevolucion As Long

Private Sub Class_Initialize()
    Set Servicio = New clsServicio
End Sub

Private Sub Class_Terminate()
    Set Servicio = Nothing
End Sub
