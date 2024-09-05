VERSION 5.00
Begin VB.Form frmEdit 
   Caption         =   "Editar Valores"
   ClientHeight    =   1980
   ClientLeft      =   4470
   ClientTop       =   4215
   ClientWidth     =   2415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   2415
   Begin VB.ComboBox cEstado 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1200
      Width           =   1155
   End
   Begin VB.CommandButton bCancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   1260
      TabIndex        =   9
      Top             =   1620
      Width           =   1095
   End
   Begin VB.CommandButton bGrabar 
      Caption         =   "Aplicar"
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   1620
      Width           =   1095
   End
   Begin VB.TextBox tNumero 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   840
      MaxLength       =   10
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cTipoDocum 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1155
   End
   Begin VB.TextBox tRojo 
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   840
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   1260
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Número:"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   900
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Docum.:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   915
   End
   Begin VB.Label lArranca 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Rojo:"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public prm_Ed_Caso As Byte  'Para mostrar/ocultar el combo

Public prm_Ed_ROJO As Long
Public prm_Ed_TipoD As Integer
Public prm_Ed_Numero As String
Public prm_Ed_Estado As Integer
Public prm_Ed_SerieRoja As String

Public prm_Tipos  As String

Public prm_Grabar As Boolean

Private Sub bCancel_Click()
    Unload Me
End Sub

Private Sub bGrabar_Click()

    If Val(tRojo.Text) = 0 Then tRojo.SetFocus: Exit Sub
    If cTipoDocum.ListIndex = -1 Then cTipoDocum.SetFocus: Exit Sub
    If cTipoDocum.ListIndex = 0 Then
        If cEstado.ListIndex <= 0 Then cEstado.SetFocus
    Else
        If Trim(tNumero.Text) = "" Then tNumero.SetFocus: Exit Sub
    End If
    prm_Ed_ROJO = Val(tRojo.Text)
    If cTipoDocum.ListIndex <= 0 Then
        prm_Ed_TipoD = 0
    Else
        prm_Ed_TipoD = cTipoDocum.ItemData(cTipoDocum.ListIndex)
    End If
    prm_Ed_Numero = Trim(tNumero.Text)
    
    prm_Ed_Estado = cEstado.ItemData(cEstado.ListIndex)
    
    prm_Grabar = True
    Unload Me
End Sub

Private Sub cTipoDocum_Change()
    If cTipoDocum.Text = "" Then tNumero.Text = ""
    tNumero.Enabled = (cTipoDocum.ListIndex)
End Sub

Private Sub cTipoDocum_Click()
    If cTipoDocum.Text = "" Then tNumero.Text = ""
    tNumero.Enabled = (cTipoDocum.ListIndex > 0)
    tNumero.BackColor = IIf(tNumero.Enabled, vbWindowBackground, vbButtonFace)
End Sub

Private Sub Form_Load()

    prm_Grabar = False
    fnc_Inicializar
    
End Sub

Private Function fnc_Inicializar()

    Dim arrX() As String, idX As Integer
    Dim indexSEL As Integer
    arrX = Split(prm_Tipos, ",")
    cTipoDocum.AddItem ""
    For idX = LBound(arrX) To UBound(arrX)
        cTipoDocum.AddItem RetornoNombreDocumento(CInt(arrX(idX)), True)
        cTipoDocum.ItemData(cTipoDocum.NewIndex) = arrX(idX)
        
        If Val(arrX(idX)) = prm_Ed_TipoD Then indexSEL = (cTipoDocum.ListCount - 1)
    Next
    
    tRojo.Text = prm_Ed_ROJO
    cTipoDocum.ListIndex = indexSEL
    tNumero.Text = prm_Ed_Numero
    
    With cEstado
    .AddItem " ": .ItemData(.NewIndex) = 0: If 0 = prm_Ed_Estado Then indexSEL = (.ListCount - 1)
    .AddItem "PA": .ItemData(.NewIndex) = 1: If 1 = prm_Ed_Estado Then indexSEL = (.ListCount - 1)
    .AddItem "EXT": .ItemData(.NewIndex) = 10: If 10 = prm_Ed_Estado Then indexSEL = (.ListCount - 1)
    .AddItem "ExA": .ItemData(.NewIndex) = 11: If 11 = prm_Ed_Estado Then indexSEL = (.ListCount - 1)
    
    .ListIndex = indexSEL
    
    cEstado.Enabled = (prm_Ed_Caso = 1)
    tRojo.Enabled = (prm_Ed_Caso = 0 Or prm_Ed_Caso = 1)
    End With
        
End Function

