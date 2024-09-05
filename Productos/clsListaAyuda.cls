VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsListaAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------------------------
'Metodo: ActivarAyuda
'       Consulta:  Consulta para realizar la seleccion. Se deben enviar los datos ya filtrados.
'                        Ejemplo: Select CliCodigo, CliCodigo, Nombre from .....
'                                      - El primero es utilizado como la KEY y  desde el segundo en adelante
'                                        son los resultados a cargar en la lista.
'----------------------------------------------------------------------------------------------------------------------
Public Sub ActivarAyuda(ByVal cConexion As RDO.rdoConnection, Consulta As String, Optional AnchoForm As Currency = 6000, Optional OcultoCol1 As Boolean = True, Optional Titulo As String = "")
On Error GoTo ErrActivo
    Screen.MousePointer = 11
    
    Set cBase = cConexion   'Tomo la conexión y la asigno a mi var global.
    
    Set RsAux = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No se encontraron datos para los filtros seleccionados.", vbInformation, "Lista de Ayuda"
    Else
        frmAyuda.bolHidePrimera = OcultoCol1
        If Titulo <> "" Then frmAyuda.Caption = Titulo
        frmAyuda.Width = AnchoForm
        frmAyuda.Show vbModal
    End If
    Set cBase = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrActivo:
    miRutinas.OcurrioError rdoErrors, "Clase Lista de Ayuda", "Ocurrio un error al iniciar el metodo ActivarAyuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Public Function RetornoFilaSeleccionada() As Integer
    RetornoFilaSeleccionada = frmAyuda.prtFilaSeleccionada
End Function

Public Function RetornoDatoSeleccionado(intCol As Integer) As Variant
    'Retorna el dato que contiene la columna seleccionada.
    RetornoDatoSeleccionado = frmAyuda.vsAyuda.Cell(flexcpText, frmAyuda.prtFilaSeleccionada, intCol)
End Function

Private Sub Class_Terminate()
On Error Resume Next
    If FormActivo("frmAyuda") Then Unload frmAyuda
End Sub
