VERSION 5.00
Begin VB.Form frmDondeImprimo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Impresión de salidas de caja"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6405
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
   ScaleHeight     =   2595
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6375
      TabIndex        =   5
      Top             =   0
      Width           =   6405
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione donde  imprimirá las salidas de caja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   5535
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmDondeImprimo.frx":0000
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.CommandButton butCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton butOk 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox cboPos 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1440
      Width           =   3495
   End
   Begin VB.OptionButton optImpresora 
      Caption         =   "Tickets en:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton optImpresora 
      Caption         =   "Papel A5 en impresora "
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   5055
   End
End
Attribute VB_Name = "frmDondeImprimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub butCancelar_Click()
    Unload Me
End Sub

Private Sub butOk_Click()
    
    If optImpresora(1).Value And cboPos.ListIndex < 0 Then
        MsgBox "Debe indicar la impresora donde se emiten los tickets.", vbExclamation, "ATENCIÓN"
        cboPos.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Confirma almacenar la configuración?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        Dim oCnfg As New clsCnfgImpresora
        If optImpresora(0).Value Then
            oCnfg.Opcion = 0
            oCnfg.ImpresoraTickets = ""
        Else
            oCnfg.Opcion = 1
            oCnfg.ImpresoraTickets = cboPos.Text
        End If
        oCnfg.GuardarConfiguracion cnfgAppNombreMovimientoCaja, cnfgKeyTicketMovimientoCaja
        End If
    Unload Me
    
End Sub

Private Sub Form_Load()
On Error GoTo errL
    
    cboPos.Clear
    Dim X As Printer
    
    For Each X In Printers
        cboPos.AddItem Trim(X.DeviceName)
    Next
    
    oCnfgPrint.CargarConfiguracion cnfgAppNombreMovimientoCaja, cnfgKeyTicketMovimientoCaja
    
    optImpresora(oCnfgPrint.Opcion).Value = True
    optImpresora(0).Caption = optImpresora(0).Caption & " " & paPrintConfD
    If (oCnfgPrint.Opcion > 0) Then
        cboPos.Text = Trim(oCnfgPrint.ImpresoraTickets)
        Dim I As Integer
        For I = 0 To cboPos.ListCount - 1
            If cboPos.List(I) = Trim(oCnfgPrint.ImpresoraTickets) Then
                cboPos.ListIndex = I
                Exit For
            End If
        Next
    End If
    Exit Sub
    
errL:
    clsGeneral.OcurrioError "Error al cargar la configuración.", Err.Description, "ATENCIÓN"
End Sub

Private Sub optImpresora_Click(Index As Integer)
    cboPos.Enabled = (Index = 1)
End Sub
