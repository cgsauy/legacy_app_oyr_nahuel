VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub ActivoEmbarque(Modal As Boolean, Optional Codigo As Long = 0)
On Error GoTo ErrAE
    RelojA
    If miconexion.AccesoAlMenu("MaEmbarque") Then
        InicioConexionBD miconexion.TextoConexion(logImportaciones)
        MaEmbarque.Status.Panels("usuario") = "Usuario: " & miconexion.UsuarioLogueado(Nombre:=True)
        MaEmbarque.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        MaEmbarque.pModal = Modal
        If Modal Then MaEmbarque.Show vbModal Else MaEmbarque.Show vbModeless
    Else
        If miconexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
    End If
    RelojD
    Exit Sub
ErrAE:
    clsGeneral.OcurrioError "Ocurrio un error al instanciar la clase Embarque.", Trim(Err.Description)
    RelojD
End Sub

Public Function RetornoCodigoEmbarque() As Long
On Error GoTo ErrRCE
    RetornoCodigoEmbarque = MaEmbarque.pSeleccionado
    Exit Function
ErrRCE:
    clsGeneral.OcurrioError "Ocurrio un error al retornar el código de embarque.", Trim(Err.Description)
End Function
