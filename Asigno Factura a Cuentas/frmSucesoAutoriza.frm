VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSucesoAutoriza 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Autorizar aportes"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   390
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ilstIconos 
      Left            =   960
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSucesoAutoriza.frx":0000
            Key             =   "i2"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSucesoAutoriza.frx":05D3
            Key             =   "i3"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSucesoAutoriza.frx":0BA2
            Key             =   "i4"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSucesoAutoriza.frx":1175
            Key             =   "i1"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picEspera 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   6375
      TabIndex        =   7
      Top             =   720
      Width           =   6375
      Begin VB.Image imgCursor 
         Height          =   480
         Left            =   720
         Picture         =   "frmSucesoAutoriza.frx":1746
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Solicitando autorización, por favor espere ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   840
         Width           =   4215
      End
   End
   Begin VB.Timer tmEsperaAutorizacion 
      Left            =   240
      Top             =   3120
   End
   Begin VB.CommandButton butOk 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton butCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtComentario 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   5775
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
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6165
      TabIndex        =   4
      Top             =   0
      Width           =   6195
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmSucesoAutoriza.frx":1D07
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizar aportes vencidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   495
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Label lblImporteAutorizar 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitar autorización para utilizar aportes vencidos por :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00446537&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "frmSucesoAutoriza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ImporteAAutorizar As Currency
Public ImporteTotalAAsignar As Currency
Public Moneda As String
Public Cliente As Long
Public CodigoSuceso As Long

Private Sub butCancelar_Click()
    If Val(butOk.Tag) > 0 Then
        If MsgBox("Aún no se obtuvo autorización para el suceso." & vbCrLf & vbCrLf & "¿Confirma cancelar la espera?", vbYesNo + vbQuestion + vbDefaultButton2, "Autorización") = vbNo Then Exit Sub
        Unload Me
    End If
End Sub

Private Sub butOk_Click()
On Error GoTo errGS

    Screen.MousePointer = 11

    'clsGeneral.RegistroSucesoAutorizado cBase, Now, TipoSuceso.AportesACuenta, paCodigoDeTerminal, paCodigoDeUsuario, 0, 0, "Autorizar uso de aportes vencidos por " & Moneda & " " & ImporteAAutorizar & ", el monto total asignado es de " & cMoneda.Text & " " & ImporteTotalAAsignar, txtComentario.Text, ImporteAAutorizar, Cliente, 0
     'No utilizo la dll ya que controla el ID de usuario autoriza > 0
    
     'Para consultar en AP busco los Autoriza = 0 y verificado = 0.
     Cons = "Insert into Suceso" _
           & " (SucFecha, SucTipo, SucTerminal, SucUsuario, SucDescripcion, SucDefensa, SucValor, SucCliente, SucAutoriza, SucVerificado, SucDocumento)" _
           & " Values (GetDate(), " & TipoSuceso.AportesACuenta & ", " & paCodigoDeTerminal & ", " & paCodigoDeUsuario _
           & ", 'Aportes vencidos por " & Moneda & " " & ImporteAAutorizar & "', '" & txtComentario.Text & "' " _
           & ", " & ImporteAAutorizar & ", " & Cliente & ", 0, 0, 0)"
    cBase.Execute Cons
    
    'Me quedo con el ID del suceso.
    Cons = "SELECT Max(SucCodigo) FROM Suceso WHERE SucCliente = " & Cliente & " AND SucUsuario = " & paCodigoDeUsuario & " AND SucTerminal = " & paCodigoDeTerminal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    butOk.Tag = RsAux(0)
    RsAux.Close
    
    Screen.MousePointer = 0
    
    picEspera.Top = Picture1.Top + Picture1.Height
    picEspera.Width = Me.ScaleWidth
    picEspera.Visible = True
    
    butOk.Visible = False
    
    'Prendo timer en espera de autorización.
    tmEsperaAutorizacion.Tag = Now
    ilstIconos.Tag = 1
    tmEsperaAutorizacion.Enabled = True
    tmEsperaAutorizacion.Interval = 600
    Exit Sub
    
errGS:
    clsGeneral.OcurrioError "Error al intentar grabar el suceso.", Err.Description, "Suceso autoriza"
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    picEspera.Visible = False
    txtComentario.MaxLength = 100
    txtComentario.Text = "Utilizar aportes vencidos por " & Moneda & " " & ImporteAAutorizar & ", el monto total asignado es de " & Moneda & " " & ImporteTotalAAsignar
    lblImporteAutorizar.Caption = lblImporteAutorizar.Caption & " " & Moneda & " " & ImporteAAutorizar
End Sub

Private Sub tmEsperaAutorizacion_Timer()
On Error GoTo errEA
    tmEsperaAutorizacion.Enabled = False
    Screen.MousePointer = 11
    
    imgCursor.Tag = Val(imgCursor.Tag) + 1
    If imgCursor.Tag = 5 Then imgCursor.Tag = 1
    imgCursor.Picture = ilstIconos.ListImages("i" & imgCursor.Tag).Picture
    imgCursor.Refresh
    
    If DateDiff("s", tmEsperaAutorizacion.Tag, Now) > 5 Then
    
        Cons = "SELECT IsNull(SucAutoriza, 0) SucAutoriza, SucVerificado  FROM Suceso WHERE SucCodigo = " & Val(butOk.Tag) & " AND SucVerificado IS NOT NULL"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux("SucAutoriza") > 0 Then
            If RsAux("SucVerificado") = 9 Then
                MsgBox "NO PUEDE UTILIZAR LOS APORTES VENCIDOS ya que su solicitud NO fue autorizada", vbExclamation, "ATENCIÓN"
            Else
                CodigoSuceso = Val(butOk.Tag)
            End If
            RsAux.Close
            Screen.MousePointer = 0
            Unload Me
            Exit Sub
        End If
        RsAux.Close
        
        tmEsperaAutorizacion.Tag = Now
    End If
    Screen.MousePointer = 0
    tmEsperaAutorizacion.Enabled = True
    Exit Sub
errEA:
    clsGeneral.OcurrioError "Error al verificar si el suceso fue autorizado.", Err.Description, "Sucesos autorizados"
    tmEsperaAutorizacion.Enabled = True
End Sub
