VERSION 5.00
Begin VB.Form frmOpcionSistarbanc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Devolucón Sistarbanc"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton butOk 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   5655
   End
   Begin VB.CheckBox chkDevuelveFlete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "¿Le devolvemos el importe del flete?"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   2280
      Width           =   5895
   End
   Begin VB.OptionButton opbDevolvemosDinero 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "¿Le devolvemos el dinero?"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   6615
   End
   Begin VB.OptionButton opbAporteACuenta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "¿El dinero lo deja en un aporte a su cuenta?"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1440
      Width           =   6615
   End
   Begin VB.CheckBox chkLlevaOtro 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "¿El cliente lleva otro producto?"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   6975
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
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
      ScaleWidth      =   6120
      TabIndex        =   0
      Top             =   0
      Width           =   6150
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Devolución pago por banco"
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
         Height          =   405
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3285
      End
   End
End
Attribute VB_Name = "frmOpcionSistarbanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Confirmo As Boolean
Public Respuesta As clsOpcionSistarbanc

Private Sub butOk_Click()
    If MsgBox("¿Confirma las opciones elegidas?", vbQuestion + vbYesNo + vbDefaultButton2, "Devolución pago por banco") = vbYes Then
        Set Respuesta = New clsOpcionSistarbanc
        If (chkLlevaOtro.value = vbUnchecked) Then
            Respuesta.NoLlevaProducto = True
            If (opbDevolvemosDinero.value) Then
                Respuesta.LeDevolvemosPlata = True
                Respuesta.DescontarFlete = (chkDevuelveFlete.value = vbChecked)
            End If
        End If
        Confirmo = True
        Unload Me
    End If
End Sub

Private Sub chkLlevaOtro_Click()
    opbAporteACuenta.Enabled = (chkLlevaOtro.value = vbUnchecked)
    opbDevolvemosDinero.Enabled = (chkLlevaOtro.value = vbUnchecked)
    chkDevuelveFlete.Enabled = (chkLlevaOtro.value = vbUnchecked And opbDevolvemosDinero.value)
    If (chkLlevaOtro.value = vbUnchecked And opbDevolvemosDinero.value = opbAporteACuenta.value) Then
        opbAporteACuenta.value = True
    End If
End Sub

Private Sub Form_Load()
    Confirmo = False
    chkLlevaOtro.value = vbChecked
End Sub

Private Sub opbAporteACuenta_Click()
    chkDevuelveFlete.Enabled = opbDevolvemosDinero.value And opbDevolvemosDinero.Enabled
    If (opbAporteACuenta.value) Then chkDevuelveFlete.value = vbUnchecked
End Sub

Private Sub opbDevolvemosDinero_Click()
    chkDevuelveFlete.Enabled = opbDevolvemosDinero.value And opbDevolvemosDinero.Enabled
End Sub
