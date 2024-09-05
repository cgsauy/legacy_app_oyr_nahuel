VERSION 5.00
Begin VB.Form frmDondeImprimo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Impresión de conformes"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5940
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
   ScaleHeight     =   1800
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton butOk 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox cboPos 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   600
      Width           =   3495
   End
   Begin VB.OptionButton optImpresora 
      Caption         =   "Tickets en:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton optImpresora 
      Caption         =   "Papel comodin en impresora "
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5175
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
        oCnfg.GuardarConfiguracion cnfgAppNombreCopia, cnfgKeyTicketCopiaFactura
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
    
    oCnfgPrint.CargarConfiguracion cnfgAppNombreCopia, cnfgKeyTicketCopiaFactura
    
    optImpresora(oCnfgPrint.Opcion).Value = True
    optImpresora(0).Caption = optImpresora(0).Caption & " " & paIConformeN
    If (oCnfgPrint.Opcion > 0) Then
        cboPos.Text = Trim(oCnfgPrint.ImpresoraTickets)
        Dim i As Integer
        For i = 0 To cboPos.ListCount - 1
            If cboPos.List(i) = Trim(oCnfgPrint.ImpresoraTickets) Then
                cboPos.ListIndex = i
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
