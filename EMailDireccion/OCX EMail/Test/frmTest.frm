VERSION 5.00
Object = "{1292AE18-2B08-4CE3-9F79-9CB06F26AB54}#1.6#0"; "orEMails.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   3330
   ClientTop       =   2625
   ClientWidth     =   6195
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
   ScaleHeight     =   4365
   ScaleWidth      =   6195
   Begin orEMails.ctrEMails ctrEMails2 
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Top             =   2160
      Width           =   3555
      _ExtentX        =   5980
      _ExtentY        =   556
      ForeColor       =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   315
      Left            =   540
      TabIndex        =   8
      Top             =   2640
      Width           =   555
   End
   Begin orEMails.ctrEMails ctrEMails1 
      Height          =   315
      Left            =   1500
      TabIndex        =   7
      Top             =   2640
      Width           =   2760
      _ExtentX        =   5980
      _ExtentY        =   556
      BackColor       =   12640511
      ForeColor       =   16711680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enabled"
      Height          =   315
      Left            =   4260
      TabIndex        =   6
      Top             =   420
      Width           =   975
   End
   Begin VB.TextBox tNombres 
      Height          =   285
      Left            =   900
      TabIndex        =   5
      Text            =   "Nombre Apellido"
      Top             =   840
      Width           =   2895
   End
   Begin orEMails.ctrEMails mControl 
      Height          =   315
      Left            =   1020
      TabIndex        =   4
      Top             =   1560
      Width           =   2880
      _ExtentX        =   5980
      _ExtentY        =   556
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Asignar"
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   420
      Width           =   1155
   End
   Begin VB.TextBox tIDCliente 
      Height          =   285
      Left            =   900
      TabIndex        =   0
      Text            =   "179446"
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "eMails:"
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   1620
      Width           =   675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   675
   End
   Begin VB.Menu main 
      Caption         =   "main"
      Begin VB.Menu mnu 
         Caption         =   "mnu"
         Index           =   0
      End
      Begin VB.Menu mnu 
         Caption         =   "sbmenu1"
         Index           =   1
      End
      Begin VB.Menu mnu 
         Caption         =   "sbmenu2"
         Index           =   2
      End
   End
   Begin VB.Menu m2 
      Caption         =   "m2"
      Begin VB.Menu o 
         Caption         =   "oooooooo"
      End
      Begin VB.Menu oo 
         Caption         =   "ooooooo"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    mControl.CargarDatos Val(tIDCliente.Text)
    mControl.IdsPorDefecto = Trim(tNombres.Text)
    mControl.IDUsuario = miConexion.UsuarioLogueado(True, False)
End Sub

Private Sub Command2_Click()
    mControl.Enabled = Not mControl.Enabled
End Sub

Private Sub Command3_Click()
Dim o1 As New clsCliente
o1.Personas 53405
Set o1 = Nothing

End Sub

Private Sub ctrEMails1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MsgBox "REtirn"
End Sub

Private Sub Form_Click()
    
    'Load mnu(1)
    'Load sbmenu1(1)
    'Load sbmenu2(1)
    
    'sbmenu2(1).Parent = 1
    
    PopupMenu main
    
End Sub

Private Sub Form_Load()
    fnc_Inicializo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierroConexion
End Sub

Private Function fnc_Inicializo()
    mControl.OpenControl cBase
End Function

Private Sub mnu_Click(Index As Integer)
    DoEvents
    
End Sub
