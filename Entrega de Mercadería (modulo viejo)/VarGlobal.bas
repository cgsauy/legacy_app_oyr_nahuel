Attribute VB_Name = "VarGlobal"
'------------------------------------------------------------------------------------------------
'   Funciones:
'       TextoValido (S as String)
'       ValidoFormatoFolder(Folder As String)
'       Encripto (S As String)
'       DesEncripto (S As String)
'------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------
'   Procedimientos:
'       Botones(Nu As Boolean, Mo As Boolean, El As Boolean, Gr As Boolean, Ca As Boolean, Toolbar1 As Control, nForm As Form)
'       BotonesRegistro(Pri As Boolean, Ant As Boolean, Sig As Boolean, Ult As Boolean, Toolbar1 As Control, nForm As Form)
'------------------------------------------------------------------------------------------------

Option Explicit
Public pathApp As String
Global Const paEmpresa = "Carlos Gutiérrez s.a."

Public clsGeneral  As New clsorCGSA
Public miConexion As New clsConexion

'Datos de los Usuarios-------------------------------
Public paInicialDeUsuario As String
Public paCodigoDeUsuario As Long
Public paCodigoDeSucursal As Long
Public paCodigoDeTerminal As Long
Public paBD As String

Public LoginOK As Boolean

'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno
Public RsAux As rdoResultset
Public QAux As rdoQuery

'BOOLEANAS-------------------------------
Public sCargoParametros As Boolean    'Parametrizar
Public sAyuda As Boolean

'STRING----------------------------------
Public Cons As String
Public aTexto As String

'ENTEROS---------------------------------
Public Col As Integer
Public I As Integer

'LisView---------------------------------
Public itmX As ListItem

Public gPathListados As String

'-----------------------------------------------------------------------------------
'   Habilita y deshabilita los botones y menus del toolbar
'-----------------------------------------------------------------------------------
Public Sub Botones(Nu As Boolean, Mo As Boolean, El As Boolean, Gr As Boolean, Ca As Boolean, Toolbar1 As Control, nForm As Form)

    'Habilito y Desabilito Botones.
    Toolbar1.Buttons("nuevo").Enabled = Nu
    nForm.MnuNuevo.Enabled = Nu
    
    Toolbar1.Buttons("modificar").Enabled = Mo
    nForm.MnuModificar.Enabled = Mo
    
    Toolbar1.Buttons("eliminar").Enabled = El
    nForm.MnuEliminar.Enabled = El
    
    Toolbar1.Buttons("grabar").Enabled = Gr
    nForm.MnuGrabar.Enabled = Gr
    
    Toolbar1.Buttons("cancelar").Enabled = Ca
    nForm.MnuCancelar.Enabled = Ca

End Sub

'---------------------------------------------------
'   Verifica si el texto ingresado es valido.
'       Controla si hay comillas simples.
'---------------------------------------------------
Public Function TextoValido(S As String)

    If InStr(S, "'") > 0 Then
        TextoValido = False
    Else
        TextoValido = True
    End If
    
End Function

'-----------------------------------------------------------------------------------
'   Habilita y deshabilita los botones y menus de registros del toolbar.
'-----------------------------------------------------------------------------------
Public Sub BotonesRegistro(Pri As Boolean, Ant As Boolean, Sig As Boolean, Ult As Boolean, Toolbar1 As Control, nForm As Form)

    'Habilito y Desabilito Botones.
    Toolbar1.Buttons("primero").Enabled = Pri
    nForm.MnuPrimero.Enabled = Pri
    
    Toolbar1.Buttons("anterior").Enabled = Ant
    nForm.MnuAnterior.Enabled = Ant
    
    Toolbar1.Buttons("siguiente").Enabled = Sig
    nForm.MnuSiguiente.Enabled = Sig
    
    Toolbar1.Buttons("ultimo").Enabled = Ult
    nForm.MnuUltimo.Enabled = Ult
    
End Sub

