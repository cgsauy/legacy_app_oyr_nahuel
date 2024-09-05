VERSION 5.00
Begin VB.Form frmSaldoAFavor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Saldo a favor"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7215
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
   ScaleHeight     =   2895
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton butAporte 
      Caption         =   "Aporte a cuenta"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton butPendiente 
      Caption         =   "Pendiente"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.PictureBox picTitulo 
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
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7185
      TabIndex        =   2
      Top             =   0
      Width           =   7215
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo a favor de cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Indique que opción es conveniente utilizar para asignar el saldo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hago un aporte a la cuenta personal del cliente, no se en que momento se utilizará este saldo."
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voy a utilizar este aporte a la brevedad, se hace un pendiente de caja negativo para el cliente"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   4815
   End
End
Attribute VB_Name = "frmSaldoAFavor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TipoAporte As Byte

Private Sub butAporte_Click()
    TipoAporte = 2
    Unload Me
End Sub

Private Sub butPendiente_Click()
    TipoAporte = 1
    Unload Me
End Sub

Private Sub Form_Load()
    TipoAporte = 0
End Sub
