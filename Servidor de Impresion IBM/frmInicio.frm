VERSION 5.00
Begin VB.Form frmInicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgOK 
      Height          =   480
      Left            =   840
      Picture         =   "frmInicio.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgError 
      Height          =   480
      Left            =   240
      Picture         =   "frmInicio.frx":1708A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum enuEstado
    Normal = 1
    Error = 2
End Enum

Public Sub CambiarIcono(ByVal mensaje As String, icono As enuEstado)
    
    Select Case icono
        Case enuEstado.Normal
            Me.Icon = imgOK.Picture
        Case enuEstado.Error
            Me.Icon = imgError.Picture
    End Select
    TrayModify Me.hwnd, Me, mensaje
    
End Sub

Private Sub Form_Load()
    Me.Visible = False
    Me.Icon = imgError.Picture
    TrayNotify Me.hwnd, Me, "Conectando ..."
    Me.Refresh
    'Abro el form para que lo activen.
    frmServerPosPrinter.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ret As Integer
    ret = TrayClick(Button, Shift, X, Y)
    Select Case ret
        Case 1: frmServerPosPrinter.Show
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    TrayRemove Me.hwnd
End Sub
