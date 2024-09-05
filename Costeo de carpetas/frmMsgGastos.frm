VERSION 5.00
Begin VB.Form frmMsgGastos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faltan Gastos Obligatorios"
   ClientHeight    =   2745
   ClientLeft      =   2955
   ClientTop       =   4650
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMsgGastos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bNo 
      Caption         =   "&No"
      Default         =   -1  'True
      Height          =   315
      Left            =   4920
      TabIndex        =   0
      Top             =   2340
      Width           =   975
   End
   Begin VB.CommandButton bSi 
      Caption         =   "&Si"
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   2340
      Width           =   975
   End
   Begin VB.TextBox tGastos 
      Appearance      =   0  'Flat
      Height          =   1875
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "frmMsgGastos.frx":164A
      Top             =   420
      Width           =   5835
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Carpeta:"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lCarpeta 
      BackStyle       =   0  'Transparent
      Caption         =   "Incoterm"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label lPago 
      BackStyle       =   0  'Transparent
      Caption         =   "Incoterm"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4740
      TabIndex        =   7
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pago:"
      Height          =   255
      Left            =   4260
      TabIndex        =   6
      Top             =   120
      Width           =   555
   End
   Begin VB.Label lIncoterm 
      BackStyle       =   0  'Transparent
      Caption         =   "Incoterm"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Incoterm:"
      Height          =   255
      Left            =   2340
      TabIndex        =   4
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Desea continuar con el Costeo de la Carpeta ?"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   4155
   End
End
Attribute VB_Name = "frmMsgGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmIdCarpeta As Long
Public prmTexto As String
Public prmOK As Boolean

Private Sub bNo_Click()
    On Error Resume Next
    prmOK = False
    Unload Me
End Sub

Private Sub bSi_Click()

    On Error Resume Next
    prmOK = True
    Unload Me
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    lIncoterm.Caption = "N/D"
    lCarpeta.Caption = "N/D"
    lPago.Caption = "N/D"
    
    prmOK = False
    tGastos.Text = Trim(prmTexto)
    
    Cons = "Select * from Carpeta, Incoterm" & _
               " Where CarID = " & prmIdCarpeta & _
               " And CarIncoterm *= IncCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        lCarpeta.Caption = RsAux!CarCodigo
        If Not IsNull(RsAux!IncNombre) Then lIncoterm.Caption = Trim(RsAux!IncNombre)
        If Not IsNull(RsAux!CarFormaPago) Then
            Select Case RsAux!CarFormaPago
                Case FormaPago.cFPAnticipado: lPago.Caption = "Anticipado"
                Case FormaPago.cFPCobranza: lPago.Caption = "Cobranza"
                Case FormaPago.cFPPlazoBL: lPago.Caption = "Plazo BL"
                Case FormaPago.cFPVista: lPago.Caption = "Vista"
            End Select
        End If
    End If
    RsAux.Close
    
    Screen.MousePointer = 0
    
End Sub
