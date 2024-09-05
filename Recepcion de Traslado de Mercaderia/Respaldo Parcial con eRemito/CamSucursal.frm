VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACOMBO.OCX"
Begin VB.Form CamSucursal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Sucursal"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CamSucursal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Begin AACombo99.AACombo cLocal 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      ListIndex       =   -1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
   End
   Begin VB.CommandButton bProceder 
      Caption         =   "&Proceder"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Sucursal:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "CamSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub bProceder_Click()
On Error GoTo ErrProceder
    
    If cLocal.ListIndex = -1 Then MsgBox "Se debe seleccionar un local.", vbInformation, "ATENCIÓN": Exit Sub
    If CLng(cLocal.ItemData(cLocal.ListIndex)) = paCodigoDeSucursal Then MsgBox "Ya pertenece a esa sucursal.", vbExclamation, "ATENCIÓN": Exit Sub
    paCodigoDeSucursal = cLocal.ItemData(cLocal.ListIndex)
    Cons = "Select * From Sucursal Where SucCodigo = " & paCodigoDeSucursal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        RecTransferencia.Caption = "Recepción de Traslado (Sucursal: " & Trim(RsAux!SucAbreviacion) & ") "
    End If
    RsAux.Close
    Exit Sub
ErrProceder:
    clsGeneral.OcurrioError "Ocurrio un error inesperado. ", Trim(Err.Description)
End Sub
Private Sub cLocal_GotFocus()
    cLocal.SelStart = 0
    cLocal.SelLength = Len(cLocal.Text)
End Sub
Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bProceder.SetFocus
End Sub
Private Sub cLocal_LostFocus()
    cLocal.SelLength = 0
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 0: Me.Refresh
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    'Cargo los LOCALES
    Screen.MousePointer = 11
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cLocal
    BuscoCodigoEnCombo cLocal, paCodigoDeSucursal
    Screen.MousePointer = 0
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar el formulario. ", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Forms(Forms.Count - 2).SetFocus
End Sub
