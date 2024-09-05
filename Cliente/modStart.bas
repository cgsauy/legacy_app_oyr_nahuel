Attribute VB_Name = "modStart"
Option Explicit

Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public prmIPServer As String
Public prmPortServer As Long

Public Const sc_FIN = vbCrLf


Sub Main()

    prmIPServer = "192.168.1.79"
    prmPortServer = 1251
    
    frmCliente.Show
    
End Sub
