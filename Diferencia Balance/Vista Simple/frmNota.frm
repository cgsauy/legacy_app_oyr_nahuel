VERSION 5.00
Begin VB.Form frmNota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresar Notas"
   ClientHeight    =   2370
   ClientLeft      =   3615
   ClientTop       =   4230
   ClientWidth     =   5340
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
   ScaleHeight     =   2370
   ScaleWidth      =   5340
   Begin VB.CommandButton Command1 
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
      Left            =   4200
      TabIndex        =   2
      Top             =   1920
      Width           =   1035
   End
   Begin VB.CommandButton bOK 
      Caption         =   "&Aceptar"
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
      Left            =   3060
      TabIndex        =   1
      Top             =   1920
      Width           =   1035
   End
   Begin VB.TextBox tNotas 
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
      Height          =   1455
      Left            =   60
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   300
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   5115
   End
End
Attribute VB_Name = "frmNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
