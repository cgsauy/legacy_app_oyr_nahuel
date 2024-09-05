VERSION 5.00
Begin VB.Form frmPregunta 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pregunta"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   375
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
   Icon            =   "frmPregunta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmWait 
      Left            =   720
      Top             =   1920
   End
   Begin VB.CommandButton butCancelar 
      Caption         =   "&No"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton butOk 
      Caption         =   "&Si"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox cboResolucion 
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CheckBox chkHabilito 
      Appearance      =   0  'Flat
      Caption         =   "&Resolución accesible"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblPregunta 
      BackStyle       =   0  'Transparent
      Caption         =   "¿Está seguro?"
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   240
      Picture         =   "frmPregunta.frx":4888A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   660
   End
End
Attribute VB_Name = "frmPregunta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DialogResult As VbMsgBoxResult
'Public CodigoResolucion As Long
'Public TextoResolucion As String
'Public ResolucionAccesible As Boolean
Public ResolucionAccesible As clsResolucionAccesible

Public Property Let Titulo(ByVal value As String)
    Me.Caption = value
End Property

Public Property Let Pregunta(ByVal value As String)
    Me.lblPregunta.Caption = value
End Property

Public Sub ShowDialog()
'    butOk.SetFocus
    Me.Show vbModal
End Sub

Private Sub butCancelar_Click()
    Set Me.ResolucionAccesible = Nothing
    Unload Me
End Sub

Private Sub butOk_Click()
    'Si tengo seleccionado el combo grabo el dato.
    If (chkHabilito.value = 1) Then
        If cboResolucion.Text = "" Then
             If MsgBox("Marco la opción resolución accesible sin ingresar datos." & vbCrLf & "¿Desea continuar sin este dato?", vbQuestion + vbYesNo, "Posible error") = vbNo Then Exit Sub
        End If
        Set Me.ResolucionAccesible = New clsResolucionAccesible
        If cboResolucion.ListIndex > -1 Then Me.ResolucionAccesible.ID = cboResolucion.ItemData(cboResolucion.ListIndex)
        Me.ResolucionAccesible.Texto = cboResolucion.Text
    Else
        Set Me.ResolucionAccesible = Nothing
    End If
    DialogResult = vbYes
    Unload Me
    Exit Sub
End Sub

Private Sub cboResolucion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        butOk.SetFocus
    Else
        'si el usuario presiona un texto entonces prendo el timer de búsqueda
        tmWait.Enabled = True
        tmWait.Tag = 0
    End If
End Sub

Private Sub chkHabilito_Click()
    cboResolucion.Enabled = (chkHabilito.value = 1)
End Sub

Private Sub chkHabilito_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cboResolucion.Enabled Then cboResolucion.SetFocus Else
    End If
End Sub

Private Sub Form_Load()
    
    DialogResult = vbCancel
    With cboResolucion
        .Clear
        .Enabled = False
    End With
    
    With tmWait
        .Interval = 300
        .Enabled = False
    End With
    Dim sqlQry As String
    sqlQry = "SELECT CodId, CodTexto FROM Codigos WHERE CodCual = 166 ORDER BY CodValor1 desc, CodTexto"
    CargoCombo sqlQry, cboResolucion, ""
    
End Sub

Private Sub tmWait_Timer()

    tmWait.Enabled = False
    If Val(tmWait.Tag) > 2 Then
        'Consulto a base de datos.
        Dim sqlQry As String, txtIngresado As String
        txtIngresado = cboResolucion.Text
        sqlQry = "SELECT CodId, CodTexto FROM Codigos WHERE CodTexto Like '%" & Replace(cboResolucion.Text, " ", "%") & "%' AND CodCual = 166 ORDER BY CodValor1 desc, CodTexto"
        CargoCombo sqlQry, cboResolucion, ""
        If cboResolucion.ListCount > 0 Then
            cboResolucion.ListIndex = 0
        Else
            cboResolucion.Text = txtIngresado
            cboResolucion.SelStart = Len(txtIngresado)
        End If
    Else
        tmWait.Tag = Val(tmWait.Tag) + 1
        tmWait.Enabled = True
    End If
    
End Sub
