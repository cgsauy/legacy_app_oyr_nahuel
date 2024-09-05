Attribute VB_Name = "modStart"
Option Explicit

Public Const sPathAppIni = "\app\aplicaciones.ini"

Public sServer As String
Public objRutinas As New clsRutinas
Public sTextConexion As String

Public Sub Main()

    'Pido Acceso.
    sTextConexion = objPermisos.TextoConexion("Login")
    If objConexion.InicioConexionBD(cBase, sTextConexion) Then
        frmConsola.Show
    Else
        Set objConexion = Nothing
        Set objPermisos = Nothing
        MsgBox "No se logró la conexión a la base de datos.", vbCritical, "ATENCIÓN"
        End
    End If
    
    
End Sub

Public Function CreateTextFromFile(ByVal sFileName As String) As String
On Error GoTo ErrorHandler
Dim oFileSys As Object, oFileObj As Object
Dim sData As String
    CreateTextFromFile = vbNullString
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    Set oFileObj = oFileSys.opentextfile(sFileName, 1, False, 0)
    sData = oFileObj.Readall()
    oFileObj.Close
    CreateTextFromFile = sData
ErrorHandler:
    Set oFileObj = Nothing
    Set oFileSys = Nothing
    If Err.Number > 0 Then CreateTextFromFile = vbNullString
End Function

