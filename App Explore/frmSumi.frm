VERSION 5.00
Begin VB.Form frmSumi 
   Caption         =   "frmSumi"
   ClientHeight    =   3585
   ClientLeft      =   3180
   ClientTop       =   3000
   ClientWidth     =   6585
   Icon            =   "frmSumi.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   6585
End
Attribute VB_Name = "frmSumi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmMensaje As String

Private Sub Form_Load()
    
    On Error GoTo errMain
    
    Dim objMsg As New clsExplore
    objMsg.VerMensaje prmMensaje
    Set objMsg = Nothing
    End
    Exit Sub
    
errMain:
    MsgBox "Error al activar explorador de mensajes." & vbCrLf & _
                Err.Number & "- " & Err.Description, vbCritical, "(Load) Error de Aplicación - Explorador MSG"
    End
End Sub
