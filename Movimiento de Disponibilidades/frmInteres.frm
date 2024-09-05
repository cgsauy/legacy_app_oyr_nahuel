VERSION 5.00
Begin VB.Form frmInteres 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cálculo de Intereses Bancarios"
   ClientHeight    =   1485
   ClientLeft      =   3615
   ClientTop       =   5235
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInteres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tCBProm 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5100
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   855
      Width           =   1275
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3180
      TabIndex        =   13
      Top             =   720
      Width           =   15
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label6"
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   60
      TabIndex        =   12
      Top             =   720
      Width           =   6315
   End
   Begin VB.Label Label7 
      Caption         =   "Tasa Promedio (%):"
      Height          =   195
      Left            =   3300
      TabIndex        =   11
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label lBPTasa 
      Caption         =   "0.000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5160
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Comisión por Bajo Prom:"
      Height          =   195
      Left            =   3300
      TabIndex        =   8
      Top             =   900
      Width           =   1875
   End
   Begin VB.Label lIGTasa 
      Caption         =   "Q de Saldos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label Label5 
      Caption         =   "Tasa Promedio (%):"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label lInteresesG 
      Caption         =   "Q de Saldos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   900
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "Intereses Ganados:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   1515
   End
   Begin VB.Label lSaldoP 
      Alignment       =   1  'Right Justify
      Caption         =   "Q de Saldos:"
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   420
      Width           =   1395
   End
   Begin VB.Label lQSaldo 
      Alignment       =   1  'Right Justify
      Caption         =   "Q de Saldos:"
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "Promedio de Saldos:"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Q de Saldos:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "frmInteres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmQSaldos As Long
Public prmSumaSaldos As Currency

Public prmIGanados As Currency


Private Sub Form_Load()
    
    CentroForm Me
    
    lQSaldo.Caption = prmQSaldos
    lSaldoP.Caption = Format(prmSumaSaldos / prmQSaldos, "#,##0.00")
    
    lInteresesG.Caption = Format(prmIGanados, "#,##0.00")
    
    lIGTasa.Caption = Format((CCur(lInteresesG.Caption) * 100) / CCur(lSaldoP.Caption), "0.000")
    
End Sub

Private Sub tCBProm_Change()
    lBPTasa.Caption = "0.000"
End Sub

Private Sub tCBProm_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        On Error Resume Next
        
        If Not IsNumeric(tCBProm.Text) Then tCBProm.Text = "0"
        tCBProm.Text = Format(tCBProm.Text, "#,##0.00")
        If Val(tCBProm.Text) = 0 Then Exit Sub
        
        lBPTasa.Caption = Format((CCur(tCBProm.Text) * 100) / CCur(lSaldoP.Caption), "0.000")
        
    End If
End Sub
