VERSION 5.00
Begin VB.Form frmDondeTicket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de documentos"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5610
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
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton butOk 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox cboPos 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Impresora Tickets"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmDondeTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oCnfg As New clsImpresoraTicketsCnfg
Public Titulo As String
Public prmKeyApp As String

Private Sub butCancelar_Click()
    Unload Me
End Sub

Private Sub butOk_Click()
    If cboPos.ListIndex < 0 Then
        MsgBox "Debe indicar la impresora donde se emiten los tickets.", vbExclamation, "ATENCIÓN"
        cboPos.SetFocus
        Exit Sub
    End If
    If MsgBox("¿Confirma almacenar la configuración?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        oCnfg.Opcion = 1
        oCnfg.ImpresoraTickets = cboPos.ItemData(cboPos.ListIndex)
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    lblTitulo.Caption = Titulo
    cboPos.Clear
    CargoCombo "SELECT TicID, TicNombre From Tickeadora WHERE TicSucursal = " & paCodigoDeSucursal, cboPos, ""
    If oCnfg.ImpresoraTickets <> "" Then BuscoCodigoEnCombo cboPos, oCnfg.ImpresoraTickets

End Sub

Private Sub optImpresora_Click(Index As Integer)
    cboPos.Enabled = (Index <> 0)
End Sub
