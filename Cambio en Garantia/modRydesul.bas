Attribute VB_Name = "modRydesul"
Option Explicit

Public cBaseRD As rdoConnection       'Conexion a la Base de Datos
Public prmBDRydesul As Boolean

Public Function InicioConexionBDRydesul(strConexion, Optional sqlTimeOut As Integer = 15) As Boolean
    
    On Error GoTo ErrICBD
    prmBDRydesul = False
    InicioConexionBDRydesul = False
    
    'Conexion a la base de datos----------------------------------------
    Set cBaseRD = eBase.OpenConnection("", rdDriverNoPrompt, , strConexion)
    cBaseRD.QueryTimeout = sqlTimeOut
    
    InicioConexionBDRydesul = True
    
    prmBDRydesul = True
    Exit Function
    
ErrICBD:
    On Error Resume Next
    Screen.MousePointer = 0
'    MsgBox "Error al intentar comunicarse con la Base de Datos de RYDESUL." & vbCrLf & _
                "Error: " & Err.Description, vbExclamation, "Error de Conexión a Rydesul"
End Function

Public Function CierroConexionBDRydesul()
    On Error Resume Next
    cBaseRD.Close
End Function



