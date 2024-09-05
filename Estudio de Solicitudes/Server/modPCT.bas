Attribute VB_Name = "modPCT"
Option Explicit

Public Enum enuEventoPCT
    Login = 1             'Pedido de conexión de una terminal.
    LogOff = 2          'Aviso de una terminal que cerro la conexión.
    Mensaje = 3        'Manda el servidor mensaje con información.
    Error = 4
    ListaError = 5
    VaciarError = 6
End Enum

Public Function ws_InstanciaSocket(Sockets As Variant) As Long

Dim Indice As Long
    
    On Error GoTo InicioControl
    For Indice = 1 To 10000
        If Sockets(Indice).Name = "" Then       'Si tengo uno vacío da error
        End If
    Next Indice
    
    'Si no dio error tengo todos asignados entonces le sumo uno al array.
    Indice = Indice + 1
    
InicioControl:
    On Error GoTo errIS
    Load Sockets(Indice)
    ws_InstanciaSocket = Indice
    Exit Function
    
errIS:
    Indice = -1
    Resume Next
End Function

Public Sub ws_Desconecto(Socket As Winsock)
    On Error Resume Next
    If Socket.State <> sckClosed Then Socket.Close
End Sub
