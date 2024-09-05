VERSION 5.00
Begin VB.Form frmDisposicion 
   Caption         =   "Disposición"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton butCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton butAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   720
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "El cliente lo retira a partir del:"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Disponer del producto"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmDisposicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gServicio As Long
Public boolGrabo As Boolean

Private Sub butAceptar_Click()
    
    If (MsgBox("¿Confirma grabar la información?", vbQuestion + vbYesNo, "Grabar") = vbYes) Then
        If Option1(0).Value = True Then
            Cons = "Update taller set TalDisposicionProducto = GETDATE(), TalNoDispuso = null WHERE TalServicio = " & gServicio
        Else
            If Not IsDate(txtFecha.Text) Then
                MsgBox "Fecha incorrecta.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        
            Cons = "Update taller set TalDisposicionProducto = '" & Format(Now, "MM/dd/yyyy") & "' , TalNoDispuso = 1 WHERE TalServicio = " & gServicio
        End If
        cBase.Execute Cons
        boolGrabo = True
        Unload Me
    End If
    
End Sub

Private Sub butCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    boolGrabo = False
    txtFecha.Text = Format(Date, "dd/MM/yyyy")
End Sub

Private Sub Option1_Click(Index As Integer)
    txtFecha.Enabled = (Index = 1)
End Sub
