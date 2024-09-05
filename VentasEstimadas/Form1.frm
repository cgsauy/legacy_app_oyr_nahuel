VERSION 5.00
Begin VB.Form frmMemoVtaEst 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comentario"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmMemoVtaEstimada"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton butCancel 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton butOK 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtComentario 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Form1.frx":0000
      Top             =   1680
      Width           =   5055
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
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   5280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5310
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario venta estimada"
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
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblMes 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mes:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblArticulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Articulo:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmMemoVtaEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ArticuloNombre As String
Public ArticuloID As Long
Public Fecha As Date
Public Comentario As String
Public Resultado As VbMsgBoxResult

Private Sub butCancel_Click()
    Resultado = vbCancel
    Unload Me
End Sub

Private Sub butOK_Click()
On Error GoTo errBO
    If Trim(txtComentario.Text) = "" And Comentario <> "" Then
        If MsgBox("¿Confirma ELIMINAR el comentario?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    Else
        If MsgBox("¿Confirma almacenar el comentario?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    End If
    Cons = "SELECT * FROM VentasEstimadas WHERE VEsArticulo = " & ArticuloID & " AND VEsMesAño = '" & Format(Fecha, "yyyymmdd") & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        If (txtComentario.Text) = "" Then
            RsAux("VEsComentario") = Null
        Else
            RsAux("VEsComentario") = txtComentario.Text
        End If
        RsAux.Update
    End If
    RsAux.Close
    Comentario = Trim(txtComentario.Text)
    Resultado = vbOK
    Unload Me
    Exit Sub
errBO:
    clsGeneral.OcurrioError "Error al grabar", Err.Description, "Comentario de venta estimada"
End Sub

Private Sub Form_Load()
    Resultado = vbIgnore
    Me.lblArticulo.Caption = ArticuloNombre
    Me.lblMes.Caption = Format(Fecha, "MMM yyyy")
    Me.txtComentario.Text = Comentario
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call butOK_Click
    End If
End Sub
