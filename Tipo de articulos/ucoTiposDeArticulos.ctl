VERSION 5.00
Begin VB.UserControl ucoTiposDeArticulos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   630
   ScaleWidth      =   5430
   Begin VB.TextBox txtTipo 
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
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "ucoTiposDeArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDTipo As Long
Public IDPadre As Long

Private Sub UserControl_Initialize()
    IDTipo = 0
    IDPadre = 0
End Sub

Private Sub UserControl_Resize()

    txtTipo.Move 0, 0, ScaleWidth, ScaleHeight

End Sub
