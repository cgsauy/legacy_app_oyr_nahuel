VERSION 5.00
Begin VB.Form frmAyuda 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda"
   ClientHeight    =   3270
   ClientLeft      =   3375
   ClientTop       =   3420
   ClientWidth     =   5820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAyuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5820
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "El costo de los artículos es aproximado, se toma de la última compra del lifo (no se van ""costeando"" para llegar a valorarlas)."
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2700
      Width           =   5655
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1860
      TabIndex        =   9
      Top             =   60
      Width           =   3855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Marca Utilizada"
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
      Left            =   0
      TabIndex        =   8
      Top             =   60
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Contados Balance: la cantidad de artículos sanos contados es menor o igual al 10% del total."
      Height          =   495
      Left            =   1860
      TabIndex        =   7
      Top             =   1920
      Width           =   3915
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "41"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "La última compra fue realizada hace más de 18 meses, contando desde la fecha de balance."
      Height          =   495
      Left            =   1860
      TabIndex        =   5
      Top             =   1320
      Width           =   3915
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "(100,000) Artículo X"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "La última compra fue realizada entre 6 y 18 meses, contando desde la fecha de balance."
      Height          =   495
      Left            =   1860
      TabIndex        =   3
      Top             =   780
      Width           =   3915
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "(100,000) Artículo X"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo del artículo en la última compra (lifo) es 0."
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(100,000) Artículo X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
