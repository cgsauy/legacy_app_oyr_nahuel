VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTiposDocIdent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TiposDocIdent As Collection

Private Sub Class_Initialize()
    Set TiposDocIdent = New Collection
End Sub

Private Sub Class_Terminate()
    Set TiposDocIdent = Nothing
End Sub

Public Sub CargoTiposActivos(ByVal m_Connect As rdoConnection)
On Error GoTo errCTA
    Dim oTD As clsTipoDocumento
    Dim sQy As String
    sQy = "SELECT PDDId, PDDNombre, PDDAbreviacion, PDDTipoDocIdentidad, PDDTipoCliente, PaiCodigo, PaiNombre FROM PaisDelDocumento inner join Pais ON PDDPais = PaiCodigo ORDER BY PaiNombre, PDDNombre"
    Dim rsT As rdoResultset
    Set rsT = m_Connect.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    Do While Not rsT.EOF
        Set oTD = New clsTipoDocumento
        With oTD
            .ID = rsT("PDDId")
            .Nombre = Trim(rsT("PDDNombre"))
            .Abreviacion = Trim(rsT("PDDAbreviacion"))
            .TipoDeCliente = rsT("PDDTipoCliente")
            .TipoDocIdDGI = rsT("PDDTipoDocIdentidad")
            .Pais.ID = rsT("PaiCodigo")
            .Pais.Nombre = Trim(rsT("PaiNombre"))
        End With
        TiposDocIdent.Add oTD
        rsT.MoveNext
    Loop
    rsT.Close
    Exit Sub
errCTA:
    MsgBox "Error al cargar los tipos de documentos: " & Err.Description, vbCritical, "Tipos de documentos"
End Sub

Public Function ObtenerTipoDocumento(ByVal ID As Integer) As clsTipoDocumento
    Dim tipo As clsTipoDocumento
    For Each tipo In TiposDocIdent
        If tipo.ID = ID Then
            Set ObtenerTipoDocumento = tipo
            Exit Function
        End If
    Next
End Function

