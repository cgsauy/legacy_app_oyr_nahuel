Attribute VB_Name = "modTest"
Option Explicit

'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno


Sub Main()
    
    InicioConexionBD
    
    
    frmTest.Show vbModal
    
    CierroConexion
    
End Sub

Public Function InicioConexionBD() As Boolean
    
    On Error GoTo ErrICBD
    InicioConexionBD = False
    
    Dim mConexion As String
    Dim mTimeOut As Integer
    
    mConexion = "dsn=oyr;uid=SA;pwd=BARTOL;server=polenta;database=org;"
    mConexion = "dsn=org;uid=;pwd=;server=;dbq=C:\Punto ORG\Bases\ORG.mdb;"
    mTimeOut = 15
    
    Set eBase = rdoCreateEnvironment("", "", "")
    eBase.CursorDriver = rdUseServer
    
    'Conexion a la base de datos----------------------------------------
    Set cBase = eBase.OpenConnection("", , , mConexion)
    cBase.QueryTimeout = mTimeOut
    
    InicioConexionBD = True
    Exit Function
    
ErrICBD:
End Function

Public Function CierroConexion() As Boolean
    
    On Error GoTo ErrCC
    CierroConexion = False
    
    cBase.Close
    eBase.Close
    
    CierroConexion = True

ErrCC:
End Function


