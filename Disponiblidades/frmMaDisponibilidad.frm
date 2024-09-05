VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmMaDisponibilidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disponibilidades"
   ClientHeight    =   4875
   ClientLeft      =   4095
   ClientTop       =   3405
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaDisponibilidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6720
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      ForeColor       =   &H00000080&
      Height          =   1755
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   6495
      Begin VB.TextBox tAplicacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   3960
         MaxLength       =   40
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin AACombo99.AACombo cSubRubro 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   3855
         _ExtentX        =   6800
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
      End
      Begin VB.TextBox tNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   5
         Top             =   960
         Width           =   5175
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
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
      Begin AACombo99.AACombo cRubro 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
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
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Rubro:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "N�mero de Aplicaci�n:"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "&SubRubro:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame fValor 
      Caption         =   "       Disponibilidad Bancaria"
      ForeColor       =   &H00000080&
      Height          =   2205
      Left            =   120
      TabIndex        =   21
      Top             =   2340
      Width           =   6495
      Begin VB.TextBox tMinimo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         MaxLength       =   16
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox tCuenta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox cBancaria 
         Height          =   255
         Left            =   160
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin AACombo99.AACombo cSucursal 
         Height          =   315
         Left            =   840
         TabIndex        =   14
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
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
      Begin AACombo99.AACombo cBanco 
         Height          =   315
         Left            =   840
         TabIndex        =   12
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
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
      Begin AACombo99.AACombo cBSubRubro 
         Height          =   315
         Left            =   1140
         TabIndex        =   23
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
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
      End
      Begin AACombo99.AACombo cBRubro 
         Height          =   315
         Left            =   1140
         TabIndex        =   24
         Top             =   1440
         Width           =   3495
         _ExtentX        =   6165
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
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "SubRubro para Cheques Diferidos a Pagar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1140
         Width           =   6255
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "SubRubro:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Rubro:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "M�nim&o:"
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "N� &Cta.:"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sucu&rsal:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   4620
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "bd"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3731
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaDisponibilidad.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaDisponibilidad.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaDisponibilidad.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaDisponibilidad.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaDisponibilidad.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaDisponibilidad.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaDisponibilidad.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaDisponibilidad.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaDisponibilidad.frx":10E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaDisponibilidad.frx":11F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuBases 
      Caption         =   "&Bases"
      Begin VB.Menu MnuBx 
         Caption         =   "MnuBx"
         Index           =   0
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMaDisponibilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean
Dim RsDisp As rdoResultset
Dim iSeleccionado As Long, aTexto As String

Dim gIDDisponibilidad As Long

Private Sub cBancaria_Click()
    If cBancaria.Enabled Then
        If cBancaria.Value = vbChecked Then HabilitoIngreso Bancaria:=True Else DeshabilitoIngreso Bancaria:=True
    End If
End Sub

Private Sub cBancaria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And cBancaria.Value = vbUnchecked Then AccionGrabar
    If KeyCode = vbKeyReturn And cBancaria.Value = vbChecked Then Foco cBanco
End Sub

Private Sub cBanco_Change()
    cSucursal.Clear: cSucursal.Text = ""
End Sub

Private Sub cBanco_Click()
    cSucursal.Clear: cSucursal.Text = ""
End Sub

Private Sub cBanco_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If cBanco.ListIndex <> -1 And cSucursal.ListCount = 0 Then
            Cons = "Select SBaCodigo, SBaNombre from SucursalDeBanco Where SBaBanco = " & cBanco.ItemData(cBanco.ListIndex)
            CargoCombo Cons, cSucursal
        End If
        Foco cSucursal
    End If
    
End Sub

Private Sub cBRubro_Change()
    cBSubRubro.Clear
End Sub

Private Sub cBRubro_Click()
    cBSubRubro.Clear
End Sub

Private Sub cBRubro_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cRubro.ListIndex = -1 Then
            MsgBox "Seleccione el rubro para cargar los subrubros.", vbExclamation, "ATENCI�N"
            Exit Sub
        End If
        If cBSubRubro.ListCount > 0 Then Foco cBSubRubro: Exit Sub
        
        On Error GoTo errCargar
        Screen.MousePointer = 11
        Cons = "Select SRuID, SRunombre From SubRubro Where SRuRubro = " & cBRubro.ItemData(cBRubro.ListIndex) _
                & " Order by SRuNombre"
        CargoCombo Cons, cBSubRubro
        Screen.MousePointer = 0
        Foco cBSubRubro
    End If
    Exit Sub

errCargar:
    clsGeneral.OcurrioError "Ocurri� un error al cargar los subrubros.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub cBSubRubro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tAplicacion
End Sub

Private Sub cRubro_Change()
    cSubRubro.Clear
End Sub

Private Sub cRubro_Click()
    cSubRubro.Clear
End Sub

Private Sub cRubro_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If cRubro.ListIndex = -1 Then
            MsgBox "Seleccione el rubro para cargar los subrubros.", vbExclamation, "ATENCI�N"
            Exit Sub
        End If
        If cSubRubro.ListCount > 0 Then Foco cSubRubro: Exit Sub
        
        On Error GoTo errCargar
        Screen.MousePointer = 11
        Cons = "Select SRuID, SRunombre From SubRubro Where SRuRubro = " & cRubro.ItemData(cRubro.ListIndex) _
                & " Order by SRuNombre"
        CargoCombo Cons, cSubRubro
        Screen.MousePointer = 0
        Foco cSubRubro
    End If
    Exit Sub

errCargar:
    clsGeneral.OcurrioError "Ocurri� un error al cargar los subrubros.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cSubRubro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tNombre
End Sub

Private Sub cSucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tCuenta
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    LoadME

End Sub

Private Sub LoadME()
    On Error Resume Next
    
    Botones True, False, False, False, False, Toolbar1, Me
    sNuevo = False: sModificar = False
    LimpioFicha
    
    Cons = "Select MonCodigo, MonSigno from Moneda": CargoCombo Cons, cMoneda
    Cons = "Select BLoCodigo, BLoNombre from BancoLocal Order by BLoNombre": CargoCombo Cons, cBanco
    
    Cons = "Select RubID, RubNombre From Rubro Order by RubNombre"
    CargoCombo Cons, cRubro
    CargoCombo Cons, cBRubro
    
    DeshabilitoIngreso True, True
    cBancaria.Enabled = False
    
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
            
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco tNombre
End Sub

Private Sub label2_Click()
    Foco cSucursal
End Sub

Private Sub Label3_Click()
    Foco tMinimo
End Sub

Private Sub Label4_Click()
    Foco tAplicacion
End Sub

Private Sub Label5_Click()
    Foco cMoneda
End Sub

Private Sub Label7_Click()
    Foco cBanco
End Sub

Private Sub Label9_Click()
    Foco cSubRubro
End Sub

Private Sub MnuBx_Click(Index As Integer)

On Error Resume Next

    If Not AccionCambiarBase(MnuBx(Index).Tag, MnuBx(Index).Caption) Then Exit Sub
    Screen.MousePointer = 11
    
    CargoParametrosImportaciones
    LoadME
   
    'Cambio el Color del fondo de controles ----------------------------------------------------------------------------------------
    Dim arrC() As String
    arrC = Split(MnuBases.Tag, "|")
    If arrC(Index) <> "" Then Me.BackColor = arrC(Index) Else Me.BackColor = vbButtonFace
    
    Frame1.BackColor = Me.BackColor
    fValor.BackColor = Me.BackColor
    cBancaria.BackColor = Me.BackColor
    '-------------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0

End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub AccionNuevo()
   
    sNuevo = True
    gIDDisponibilidad = 0
    Botones False, False, False, True, True, Toolbar1, Me
    LimpioFicha
    HabilitoIngreso True
    Foco cRubro
    cBancaria.Enabled = True
  
End Sub

Private Sub AccionModificar()

    sModificar = True
    
    Botones False, False, False, True, True, Toolbar1, Me
    cBancaria.Enabled = True
    If cBancaria.Value = vbChecked Then HabilitoIngreso True, True Else HabilitoIngreso True
            
End Sub

Private Sub AccionGrabar()

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar la informaci�n ingresada", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo errGrabar
    
    If sNuevo Then
        Cons = "Select * from Disponibilidad Where DisID = 0"
        Set RsDisp = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsDisp.AddNew
        CargoCamposBD
        RsDisp.Update: RsDisp.Close
        
        Cons = "Select Max(DisID) from Disponibilidad"
        Set RsDisp = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsDisp.EOF Then gIDDisponibilidad = RsDisp(0)
        RsDisp.Close
        
    Else                                    'Modificar----
    
        Cons = "Select * from Disponibilidad Where DisID =" & gIDDisponibilidad
        Set RsDisp = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsDisp.Edit
        CargoCamposBD
        RsDisp.Update: RsDisp.Close
        
    End If
    
    sNuevo = False: sModificar = False
    DeshabilitoIngreso True, True
    Botones True, True, True, False, False, Toolbar1, Me
    
    cBancaria.Enabled = False
    If cBancaria.Value = vbUnchecked Then cBanco.Text = "": tCuenta.Text = "": tMinimo.Text = ""
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operaci�n.", Err.Description
End Sub

Private Sub AccionEliminar()
    
    On Error GoTo Error
    Screen.MousePointer = 11
    
    Cons = "Select * from MovimientoDisponibilidadRenglon Where MDRIDDisponibilidad  = " & gIDDisponibilidad
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        MsgBox "Hay movimientos que est�n asociados a la disponibilidad seleccionada." & Chr(vbKeyReturn) & "No podr� eliminarla.", vbExclamation, "ATENCI�N"
        RsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If
    RsAux.Close
    
    If MsgBox("Confirma eliminar la disponibilidad '" & Trim(tNombre.Text) & "'", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Screen.MousePointer = 0: Exit Sub
    
    Cons = "Select * from Disponibilidad Where DisID = " & gIDDisponibilidad
    Set RsDisp = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsDisp.Delete
    RsDisp.Close
    
    LimpioFicha
    DeshabilitoIngreso True, True
    Botones True, False, False, False, False, Toolbar1, Me
    Screen.MousePointer = 0
    Exit Sub
    
Error:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operaci�n.", Err.Description
End Sub

Private Sub AccionCancelar()

    On Error Resume Next
    LimpioFicha
    If sModificar Then
         Botones True, True, True, False, False, Toolbar1, Me
        CargoCamposDesdeBD gIDDisponibilidad
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    
    DeshabilitoIngreso Cabezal:=True, Bancaria:=True
    sNuevo = False: sModificar = False
    cBancaria.Enabled = False
    
End Sub

Private Sub tAplicacion_GotFocus()
    With tAplicacion
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tAplicacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cBancaria
End Sub

Private Sub tAplicacion_LostFocus()
    tAplicacion.SelStart = 0
End Sub

Private Sub tCuenta_GotFocus()
    tCuenta.SelStart = 0: tCuenta.SelLength = Len(tCuenta.Text)
End Sub

Private Sub tCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tMinimo
End Sub

Private Sub tMinimo_GotFocus()
    tMinimo.SelStart = 0: tMinimo.SelLength = Len(tMinimo.Text)
End Sub

Private Sub tMinimo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cBRubro
End Sub


Private Sub tMinimo_LostFocus()
    If IsNumeric(tMinimo.Text) Then tMinimo.Text = Format(tMinimo.Text, FormatoMonedaP)
End Sub

Private Sub tNombre_Change()
    If Not sNuevo And Not sModificar Then Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub tNombre_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 And Not sNuevo And Not sModificar Then
        Cons = "Select DisID 'ID', Disponibilidad = DisNombre, DisCuentaBanco as  'N� Cuenta', BLoNombre 'Banco', SBaNombre 'Sucursal'" _
               & " From Disponibilidad Left Outer Join SucursalDeBanco On DisSucursal = SBaCodigo " _
                                            & " Left Outer Join BancoLocal on SBaBanco = BLoCodigo" _
               & " Where DisNombre like '" & Trim(tNombre.Text) & "%'" _
               & " Order by DisNombre"
        AccionAyuda Cons
    End If
    
End Sub

Private Sub AccionAyuda(Consulta As String)

    Screen.MousePointer = 11
    Dim aLista As New clsListadeAyuda
    iSeleccionado = aLista.ActivarAyuda(cBase, Consulta, 6500, 1)
    Me.Refresh
    Screen.MousePointer = 0
    If iSeleccionado <> 0 Then iSeleccionado = aLista.RetornoDatoSeleccionado(0)
    Set aLista = Nothing
    
    If iSeleccionado <> 0 Then CargoCamposDesdeBD iSeleccionado
    If gIDDisponibilidad <> 0 Then Botones True, True, True, False, False, Toolbar1, Me Else Botones True, False, False, False, False, Toolbar1, Me
        
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If (sNuevo Or sModificar) Then Foco cMoneda Else Foco cSubRubro
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select

End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    If cSubRubro.ListIndex = -1 Then
        MsgBox "Debe ingresar el subrubro al que est� asociada la disponibilidad.", vbExclamation, "ATENCI�N"
        Foco cSubRubro: Exit Function
    End If
    
    If Trim(tNombre.Text) = "" Then
        MsgBox "Debe ingresar un nombre para la disponibilidad.", vbExclamation, "ATENCI�N"
        Foco tNombre: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda de la disponibilidad.", vbExclamation, "ATENCI�N"
        Foco cMoneda: Exit Function
    End If
    
    If Trim(tAplicacion.Text) <> "" Then
        If Not IsNumeric(tAplicacion.Text) Then
            MsgBox "El dato ingresado no es num�rico.", vbExclamation, "ATENCI�N"
            Foco tAplicacion: Exit Function
        End If
    End If
    
    If cBancaria.Value = vbChecked Then
        If cBanco.ListIndex = -1 Then
            MsgBox "Debe seleccionar el banco para la disponibilidad bancaria.", vbExclamation, "ATENCI�N"
            Foco cBanco: Exit Function
        End If
    
        If cSucursal.ListIndex = -1 Then
            MsgBox "Debe seleccionar la sucursal para la disponibilidad bancaria.", vbExclamation, "ATENCI�N"
            Foco cSucursal: Exit Function
        End If
    
        If Trim(tMinimo.Text) = "" Or Not IsNumeric(tMinimo.Text) Then
            MsgBox "El importe m�nimo de la cuenta no es correcto.", vbExclamation, "ATENCI�N"
            Foco tMinimo: Exit Function
        End If
        
        If Trim(tCuenta.Text) = "" Then
            MsgBox "Debe ingresar el n�mero de la cuenta del banco.", vbExclamation, "ATENCI�N"
            Foco tCuenta: Exit Function
        End If
        
        If cBSubRubro.ListIndex = -1 Then
            MsgBox "Debe seleccionar el subrubro relacionado a los cheques diferidos a pagar.", vbExclamation, "ATENCI�N"
            Foco cBSubRubro: Exit Function
        End If
    End If
    
    ValidoCampos = True
    
End Function

Private Sub DeshabilitoIngreso(Optional Cabezal As Boolean = False, Optional Bancaria As Boolean = False)
       
    If Cabezal Then
        tNombre.BackColor = Blanco
        cRubro.Enabled = False: cRubro.BackColor = Inactivo
        cSubRubro.Enabled = False: cSubRubro.BackColor = Inactivo
        cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
        tAplicacion.Enabled = False: tAplicacion.BackColor = Inactivo
    End If
    
    If Bancaria Then
        cSucursal.Enabled = False: cSucursal.BackColor = Inactivo
        cBanco.Enabled = False: cBanco.BackColor = Inactivo
        tCuenta.Enabled = False: tCuenta.BackColor = Inactivo
        tMinimo.Enabled = False: tMinimo.BackColor = Inactivo
        
        cBRubro.Enabled = False: cBRubro.BackColor = Inactivo
        cBSubRubro.Enabled = False: cBSubRubro.BackColor = Inactivo
    End If
        
End Sub

Private Sub HabilitoIngreso(Optional Cabezal As Boolean = False, Optional Bancaria As Boolean = False)

    If Cabezal Then
        cRubro.Enabled = True: cRubro.BackColor = Obligatorio
        cSubRubro.Enabled = True: cSubRubro.BackColor = Obligatorio
        tNombre.BackColor = Obligatorio
        cMoneda.Enabled = True: cMoneda.BackColor = Obligatorio
        tAplicacion.Enabled = True: tAplicacion.BackColor = vbWhite
    End If
    
    If Bancaria Then
        cBanco.Enabled = True: cBanco.BackColor = Obligatorio
        cSucursal.Enabled = True: cSucursal.BackColor = Obligatorio
        tCuenta.Enabled = True: tCuenta.BackColor = Obligatorio
        tMinimo.Enabled = True: tMinimo.BackColor = Obligatorio
        
        cBRubro.Enabled = True: cBRubro.BackColor = Obligatorio
        cBSubRubro.Enabled = True: cBSubRubro.BackColor = Obligatorio
    End If
    
End Sub

Private Sub CargoCamposBD()
        
    RsDisp!DisIDSubRubro = cSubRubro.ItemData(cSubRubro.ListIndex)
    RsDisp!DisNombre = Trim(tNombre.Text)
    RsDisp!DisMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    If IsNumeric(tAplicacion.Text) Then
        RsDisp!DisAplicacion = tAplicacion.Text
    Else
        RsDisp!DisAplicacion = Null
    End If
    If cBancaria.Value = vbChecked Then
        RsDisp!DisSucursal = cSucursal.ItemData(cSucursal.ListIndex)
        RsDisp!DisCuentaBanco = Trim(tCuenta.Text)
        RsDisp!DisMinimo = CCur(tMinimo.Text)
        RsDisp!DisIDSRCheque = cBSubRubro.ItemData(cBSubRubro.ListIndex)
    Else
        RsDisp!DisSucursal = Null
        RsDisp!DisCuentaBanco = Null
        RsDisp!DisMinimo = Null
        RsDisp!DisIDSRCheque = Null
    End If
    
End Sub

Private Sub CargoCamposDesdeBD(aDisp As Long)
    
    On Error GoTo errCargo
    Screen.MousePointer = 11
    LimpioFicha
    Cons = "Select * from Disponibilidad, SubRubro " _
           & " Where DisIDSubRubro = SRuID " _
           & " And DisID = " & aDisp
    Set RsDisp = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsDisp.EOF Then
        gIDDisponibilidad = RsDisp!DisID
        
        BuscoCodigoEnCombo cRubro, RsDisp!SRuRubro
        If cRubro.ListIndex <> -1 Then
            Cons = "Select SRuID, SRuNombre From SubRubro Where SRuRubro = " & cRubro.ItemData(cRubro.ListIndex) _
                & " Order by SRuNombre"
            CargoCombo Cons, cSubRubro
        End If
        
        BuscoCodigoEnCombo cSubRubro, RsDisp!SRuID
        tNombre.Text = Trim(RsDisp!DisNombre)
        
        BuscoCodigoEnCombo cMoneda, RsDisp!DisMoneda
        
        If Not IsNull(RsDisp!DisAplicacion) Then tAplicacion.Text = Trim(RsDisp!DisAplicacion)
        
        If Not IsNull(RsDisp!DisSucursal) Then
            Cons = "Select * from SucursalDeBanco, BancoLocal Where SBaCodigo = " & RsDisp!DisSucursal
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            BuscoCodigoEnCombo cBanco, RsAux!SBaBanco
            RsAux.Close
            
            Cons = "Select SBaCodigo, SBaNombre from SucursalDeBanco Where SBaBanco = " & cBanco.ItemData(cBanco.ListIndex)
            CargoCombo Cons, cSucursal
        
            BuscoCodigoEnCombo cSucursal, RsDisp!DisSucursal
        
            If Not IsNull(RsDisp!DisCuentaBanco) Then tCuenta.Text = Trim(RsDisp!DisCuentaBanco)
            If Not IsNull(RsDisp!DisMinimo) Then tMinimo.Text = Format(RsDisp!DisMinimo, FormatoMonedaP)
            cBancaria.Value = vbChecked
            
            'Saco SR de cheque-------------------------------------------------------------------------------------------
            If Not IsNull(RsDisp!DisIDSRCheque) Then
                Cons = "Select * from SubRubro Where SRuID = " & RsDisp!DisIDSRCheque
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux.EOF Then
                    BuscoCodigoEnCombo cBRubro, RsAux!SRuRubro
                    If cBRubro.ListIndex <> -1 Then
                        Cons = "Select SRuID, SRuNombre From SubRubro Where SRuRubro = " & cBRubro.ItemData(cBRubro.ListIndex) _
                            & " Order by SRuNombre"
                        CargoCombo Cons, cBSubRubro
                    End If
                    BuscoCodigoEnCombo cBSubRubro, RsDisp!DisIDSRCheque
                End If
                RsAux.Close
            End If
            '------------------------------------------------------------------------------------------------------------------
        End If
    End If
    RsDisp.Close
    Screen.MousePointer = 0
    Exit Sub

errCargo:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurri� un error al cargar los datos de la disponibilidad.", Err.Description
End Sub

Private Sub LimpioFicha()
    
    cRubro.Text = ""
    cSubRubro.Text = ""
    tNombre.Text = ""
    cMoneda.Text = ""
    
    cBancaria.Value = vbUnchecked
    cSucursal.Text = ""
    cBanco.Text = ""
    tMinimo.Text = ""
    tCuenta.Text = ""
    tAplicacion.Text = ""
  
    cBRubro.Text = ""
    cBSubRubro.Text = ""
  
End Sub

