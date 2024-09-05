Attribute VB_Name = "ModStart"
Option Explicit

Public eBase As rdoEnvironment
Public cBase As rdoConnection

Public clsGeneral As New clsorCGSA 'clsLibGeneral
Public miConexion As New clsConexion
Public txtConexion As String

Public Sub Main()

On Error GoTo ErrMain
    
    InicioConexionBD
    
    Form1.Show
    
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Ocurrio un error al activar el ejecutable."
    Screen.MousePointer = 0
    End
End Sub

Public Function InicioConexionBD() As Boolean
    
    On Error GoTo ErrICBD
    InicioConexionBD = False
    
    Dim mConexion As String
    Dim mTimeOut As Integer
    
    'mConexion = "dsn=oyr;uid=SA;pwd=BARTOL;server=polenta;database=org;"
    'mConexion = "dsn=org;uid=;pwd=;server=;dbq=C:\Punto ORG\Bases\ORG.mdb;"
    
    mConexion = miConexion.TextoConexion("comercio")
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
