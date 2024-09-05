VERSION 5.00
Begin VB.Form frmNumeroDeSerie 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7560
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
   ScaleHeight     =   2340
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtNroSerie 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1200
      Width           =   7335
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   7530
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lavarropa nuevo y sin uso."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   7335
      End
      Begin VB.Label lblTitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Escanear número de serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmNumeroDeSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Result As VbMsgBoxResult
Public Articulo As clsArticuloEntrega

Private Sub Form_Load()
    Result = vbNo
End Sub

Private Sub txtNroSerie_GotFocus()
On Error Resume Next
    Foco txtNroSerie
End Sub

Private Sub txtNroSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And txtNroSerie.Text <> "" Then
        Articulo.NroSerieEscaneado = txtNroSerie.Text
        Result = vbOK
        Unload Me
    End If
End Sub
