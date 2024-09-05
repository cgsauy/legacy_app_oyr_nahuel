VERSION 5.00
Begin VB.Form frmAusencias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ausencias y otras"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cbOK 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cbCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtMemo 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmAusencias.frx":0000
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtFecha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ComboBox cboQue 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   960
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5775
      TabIndex        =   8
      Top             =   0
      Width           =   5775
      Begin VB.Label lbInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ausencias y otras causas"
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
         Height          =   210
         Left            =   750
         TabIndex        =   9
         Top             =   240
         Width           =   1995
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmAusencias.frx":0006
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Memo:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Día:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblQue 
      BackStyle       =   0  'Transparent
      Caption         =   "Qué?:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "frmAusencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Que As HorariosFijos
Public Dia As Date
Public Usuario As Integer

Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub cbOK_Click()
On Error GoTo errGrabo

    If cboQue.ListIndex = -1 Then
        MsgBox "Indiqué el QUE", vbExclamation, "ATENCIÓN"
        cboQue.SetFocus
        Exit Sub
    End If
    If Not IsDate(txtFecha.Text) Then
        MsgBox "Ingrese la fecha a computar.", vbExclamation, "ATENCIÓN"
        txtFecha.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
        Dim query As String
        Dim rsQ As rdoResultset
        If Que = No Then
            query = "SELECT * FROM HorarioPersonal WHERE HPeUsuario = " & Me.Usuario & " AND HPeQue = " & cboQue.ItemData(cboQue.ListIndex) _
                & " AND HPeFechaHora = '" & Format(txtFecha.Text, "yyyymmdd 00:00:00") & "'"
        ElseIf Que <> TransitoriaR And Que <> TransitoriaS Then
            query = "SELECT * FROM HorarioPersonal WHERE HPeUsuario = " & Me.Usuario & " AND HPeQue = " & cboQue.ItemData(cboQue.ListIndex) _
                & " AND HPeFechaHora between '" & Format(txtFecha.Text, "yyyymmdd 00:00:00") & "' AND '" & Format(txtFecha.Text, "yyyymmdd 23:59:59") & "'"
        Else
            query = "SELECT * FROM HorarioPersonal WHERE HPeUsuario = " & Me.Usuario & " AND HPeQue = " & cboQue.ItemData(cboQue.ListIndex) _
                & " AND HPeFechaHora = '" & Format(txtFecha.Text, "yyyymmdd HH:nn") & "'"
        End If
        Set rsQ = cBase.OpenResultset(query, rdOpenDynamic, rdConcurValues)
        If Not rsQ.EOF Then
            rsQ.Edit
            rsQ("HPeComentario") = Trim(txtMemo.Text)
        Else
            rsQ.AddNew
            rsQ("HPeUsuario") = Me.Usuario
            If Que = No Then
                rsQ("HPeFechaHora") = Format(CDate(txtFecha.Text), "mm/dd/yyyy 00:00:00")
            Else
                rsQ("HPeFechaHora") = Format(CDate(txtFecha.Text), "mm/dd/yyyy HH:nn")
            End If
            rsQ("HPeQue") = cboQue.ItemData(cboQue.ListIndex)
            rsQ("HPeComentario") = txtMemo.Text
            rsQ("HPeEnUso") = 1
        End If
        rsQ.Update
        rsQ.Close
        Screen.MousePointer = 0
        Unload Me
    End If
    Exit Sub
errGrabo:
    clsGeneral.OcurrioError "Error al intengar grabar la información.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cboQue_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtFecha.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    
    If Que = HorariosFijos.No Then
        Dim ofncs As New clsFunciones
        ' (31,32,33, 34)
        ofncs.CargoCombo "SELECT QHPId, QHPNombre FROM QueHorarioPersonal WHERE QHPId between 31 and 37", cboQue
        Set ofncs = Nothing
        txtFecha.Text = Date
    Else
        cboQue.Clear
        Select Case Que
            Case HorariosFijos.AlmuerzoR: cboQue.AddItem "Retornó del almuerzo"
            Case HorariosFijos.AlmuerzoS: cboQue.AddItem "Salida almuerzo"
            Case HorariosFijos.Ingreso: cboQue.AddItem "Ingreso"
            Case HorariosFijos.Salida: cboQue.AddItem "Salida"
            Case HorariosFijos.TransitoriaR: cboQue.AddItem "Retorno salió 1 Min."
            Case HorariosFijos.TransitoriaS: cboQue.AddItem "Salió 1 Min."
        End Select
        cboQue.ItemData(cboQue.NewIndex) = Que
        cboQue.ListIndex = 0
        txtFecha.Text = Format(Dia, "dd/MM/yyyy") & " " & Format(Now, "HH:nn")
        lbInfo.Caption = cboQue.Text
    End If
    txtMemo.Text = ""
    Screen.MousePointer = 0
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsDate(txtFecha.Text) Then
        If Que = No Then
            txtFecha.Text = Format(txtFecha.Text, "dd/MM/yyyy")
        Else
            txtFecha.Text = Format(txtFecha.Text, "dd/MM/yyyy HH:nn")
        End If
        txtMemo.SetFocus
    End If
End Sub

Private Sub txtFecha_LostFocus()
    If IsDate(txtFecha.Text) Then
        If Que = No Then
            txtFecha.Text = Format(txtFecha.Text, "dd/MM/yyyy")
        Else
            txtFecha.Text = Format(txtFecha.Text, "dd/MM/yyyy HH:nn")
        End If
    End If
End Sub

Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cbOK.SetFocus
End Sub

Private Function StringAHora(ByVal texto As String) As String
    
    texto = Replace(texto, ".", ":")
    Select Case Len(texto)
        Case 1, 2
            If IsNumeric(texto) Then
                texto = Format(texto, "00") & ":00"
            Else
                MsgBox "No se interpreto la hora", vbExclamation, "Atención"
                Exit Function
            End If
        Case Is > 2
            If InStr(1, texto, ":") > 0 Then
                If InStr(1, texto, ":") = 1 Then
                    texto = "00:" & Format(Val(Mid(texto, 2)), "00")
                Else
                    If InStr(1, texto, ":") = 2 And Len(texto) = 3 Then
                        texto = "0" & texto & "0"
                    ElseIf InStr(1, texto, ":") = 3 And Len(texto) = 4 Then
                        texto = Format(Val(Mid(texto, 1, InStr(1, texto, ":") - 1)), "00") & ":" & Mid(texto, InStr(1, texto, ":") + 1) & "0"
                    Else
                        texto = Format(Val(Mid(texto, 1, InStr(1, texto, ":") - 1)), "00") & ":" & Format(Val(Mid(texto, InStr(1, texto, ":") + 1)), "00")
                    End If
                End If
            Else
                texto = Format(texto, "0000")
                texto = Mid(texto, 1, 2) & ":" & Mid(texto, 3)
            End If
    End Select
    StringAHora = texto
End Function
