VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Tester de impresora"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5460
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
   ScaleHeight     =   1545
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butTest 
      Caption         =   "Limpiar Mem."
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton butCargarLogo 
      Caption         =   "Cargar logo"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton butTest 
      Caption         =   "Test"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cboBandeja 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Imprimir en:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butTest_Click()
    frmServerPosPrinter.EnviarAImpresora Chr$(&H1B) & Chr$(&H78)
End Sub

Private Sub Form_Load()
    
    cboBandeja.Clear
    cboBandeja.AddItem "Ambas"
    cboBandeja.AddItem "Cliente"
    cboBandeja.AddItem "Diaria"
    cboBandeja.ListIndex = 1
    
End Sub
