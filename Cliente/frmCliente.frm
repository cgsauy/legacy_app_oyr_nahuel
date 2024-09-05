VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Líneas Telefónicas"
   ClientHeight    =   4605
   ClientLeft      =   2580
   ClientTop       =   5055
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7980
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   1980
      TabIndex        =   0
      Top             =   300
      Width           =   975
   End
   Begin MSWinsockLib.Winsock wsSocket 
      Left            =   420
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lStatus 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3420
      TabIndex        =   2
      Top             =   240
      Width           =   1635
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim globalData As String

Private Sub Form_Load()
    
    On Error Resume Next
    
    Me.Show
    ws_IniciarConexion
    
    Screen.MousePointer = 0
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wsSocket.Close
    End
End Sub

Private Sub wsSocket_Close()
    On Error Resume Next
    lStatus.Caption = "Desconectado": lStatus.Refresh
End Sub

Private Sub wsSocket_DataArrival(ByVal bytesTotal As Long)
    
    Dim strDato As String
    Dim xPos As Long
    
    If wsSocket.BytesReceived > 0 Then
        
        wsSocket.GetData strDato
        globalData = globalData & strDato
    
        Do While InStr(globalData, sc_FIN) <> 0
            xPos = InStr(globalData, sc_FIN)
            
            strDato = Mid(globalData, 1, xPos)
            globalData = Mid(globalData, xPos + Len(sc_FIN))
        
            Text1.Text = Text1.Text & strDato & vbCrLf
        Loop
        
    End If
    
End Sub

Private Sub ws_IniciarConexion()

    lStatus.Caption = "Conectando ...": lStatus.Refresh
    
    wsSocket.Connect prmIPServer, prmPortServer
    
    Dim aQIntentos As Integer
    aQIntentos = 1
    
    Do While aQIntentos <= 4
        DoEvents
        If wsSocket.State = 7 Then Exit Do
        Sleep 1000
        aQIntentos = aQIntentos + 1
        
        lStatus.Caption = "Conectando ...(" & aQIntentos & ")"
        lStatus.Refresh
    Loop
    
    If wsSocket.State <> 7 Then
        lStatus.Caption = "Sin Conexión."
    Else
        lStatus.Caption = "Conectado."
    End If
    lStatus.Refresh
    
End Sub

Public Sub ws_Reconectar()
    
    On Error Resume Next
    wsSocket.Close
    
    ws_IniciarConexion
End Sub

Public Sub ws_SendData(Trama As String)

    On Error Resume Next
    If wsSocket.State = 7 Then
        wsSocket.SendData Trama
        DoEvents
    End If
    

End Sub
