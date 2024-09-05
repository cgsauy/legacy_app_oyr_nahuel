VERSION 5.00
Begin VB.Form frmEnQueVino 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "En que se va el cliente?"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6240
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
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboOpciones 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox cboDondeEstaciona 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   6240
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "¿En que se va el cliente?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Estamos haciendo una encuesta para evaluar el impacto de nuestro depósito en el tránsito."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "¿En que se va?"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "¿Dónde estacionó?"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "frmEnQueVino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Respuesta As Byte
Public SubRespuesta As Integer

Private Sub cboOpciones_Click()
    cboDondeEstaciona.ListIndex = 0
    cboDondeEstaciona.Enabled = (cboOpciones.ListIndex = 1)
End Sub

Private Sub cboOpciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cboDondeEstaciona.Enabled Then cboDondeEstaciona.SetFocus Else cmdOK.SetFocus
    End If
End Sub

Private Sub cmdOK_Click()
    
    If (cboOpciones.ListIndex = 1 And cboDondeEstaciona.ListIndex = 0) Then
        MsgBox "Debe indicar dónde estacionó?", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("¿Confirma la respuesta " & cboOpciones.Text & "?", vbQuestion + vbYesNo, "En que se va?") = vbYes Then
        Respuesta = cboOpciones.ListIndex
        If (cboDondeEstaciona.Enabled) Then SubRespuesta = cboDondeEstaciona.ListIndex Else SubRespuesta = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    cboOpciones.AddItem "No contesta"
    cboOpciones.AddItem "Auto"
    cboOpciones.AddItem "Bicicleta"
    cboOpciones.AddItem "Caminando"
    cboOpciones.AddItem "Moto"
    cboOpciones.AddItem "Ómnibus"
    cboOpciones.AddItem "Taxi"
    
    
    cboDondeEstaciona.AddItem ""
    cboDondeEstaciona.AddItem "Torre"
    cboDondeEstaciona.AddItem "Otro lugar"
    cboDondeEstaciona.AddItem "Da vueltas esperando"
    
    cboOpciones.ListIndex = 0
End Sub


