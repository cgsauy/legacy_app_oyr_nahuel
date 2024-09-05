VERSION 5.00
Begin VB.Form frmTOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reintentar Consulta"
   ClientHeight    =   1875
   ClientLeft      =   4710
   ClientTop       =   4620
   ClientWidth     =   4275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   1875
   ScaleWidth      =   4275
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tm1 
      Interval        =   1000
      Left            =   3240
      Top             =   120
   End
   Begin VB.CommandButton bCancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   1500
      Width           =   975
   End
   Begin VB.Label lTime 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1500
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Para cancelar presione el botón ""cancelar""."
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   4275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "El sistema volverá a ejecutar la consulta en 5 segundos:"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   780
      Width           =   4275
   End
   Begin VB.Label lError 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4035
   End
End
Attribute VB_Name = "frmTOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmOK As Boolean

Private Sub bCancel_Click()
    prmOK = False
    Unload Me
End Sub

Private Sub Form_Load()
    prmOK = False
    
    lError.Caption = "ERROR: " & Err.Description
    tm1.Tag = 5
    tm1.Enabled = True
    
End Sub

Private Sub tm1_Timer()
    
    tm1.Tag = Val(tm1.Tag) - 1
    If Val(tm1.Tag) = 0 Then
        prmOK = True
        Unload Me
        
    Else
        lTime.Caption = tm1.Tag: lTime.Refresh
        tm1.Enabled = True
    End If
        
End Sub
