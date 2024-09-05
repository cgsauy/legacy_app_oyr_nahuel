VERSION 5.00
Begin VB.Form frmDondeImprimo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de documentos"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5445
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
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton butCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton butOk 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox cboPos 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.OptionButton optImpresora 
      Caption         =   "Tickets"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OptionButton optImpresora 
      Caption         =   "Tickets y papel carta"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OptionButton optImpresora 
      Caption         =   "Papel carta"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Impresora Tickets"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "frmDondeImprimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCnfg As New clsImpresoraTicketsCnfg
Public prmKeyApp As String
Public prmKeyTicket As String

Private Sub butCancelar_Click()
    Unload Me
End Sub

Private Sub butOk_Click()
    
    If Not optImpresora(0).Value And cboPos.ListIndex < 0 Then
        MsgBox "Debe indicar la impresora donde se emiten los tickets.", vbExclamation, "ATENCIÓN"
        cboPos.SetFocus
        Exit Sub
    End If
    
    If (prmKeyTicket = "") Then
        MsgBox "NO HAY CLAVE", vbCritical, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("¿Confirma almacenar la configuración?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        If optImpresora(0).Value Then
            oCnfg.Opcion = 0
            oCnfg.ImpresoraTickets = 0
        Else
            If optImpresora(1).Value Then oCnfg.Opcion = 1 Else oCnfg.Opcion = 2
            oCnfg.ImpresoraTickets = cboPos.ItemData(cboPos.ListIndex)
        End If
        oCnfg.GuardarConfiguracion prmKeyApp, prmKeyTicket '"CuotasImpresora"
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    cboPos.Clear
    CargoCombo "SELECT TicID, TicNombre From Tickeadora WHERE TicSucursal = " & paCodigoDeSucursal & " AND TicDocumentos like '%,1,%'", cboPos, ""
    
    oCnfg.CargarConfiguracion prmKeyApp, "CuotasImpresora"
    optImpresora(2).Value = True
'    optImpresora(oCnfg.Opcion).Value = True
    If (oCnfg.Opcion > 0) Then
        BuscoCodigoEnCombo cboPos, oCnfg.ImpresoraTickets
    End If
End Sub

Private Sub optImpresora_Click(Index As Integer)
    cboPos.Enabled = (Index <> 0)
    
End Sub
