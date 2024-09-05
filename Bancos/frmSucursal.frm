VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSucursal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sucursales de Bancos"
   ClientHeight    =   3690
   ClientLeft      =   3495
   ClientTop       =   4260
   ClientWidth     =   4830
   Icon            =   "frmSucursal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4830
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   2100
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Depósitos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1035
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   4575
      Begin VB.TextBox tDepositoS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox tDepositoB 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "B&anco:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Sucursal:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Banco"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   4575
      Begin VB.TextBox tBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox tCuenta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   7
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox tNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   5
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox tCodigoS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Cuen&ta:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Banco:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lCodigo 
         BackStyle       =   0  'Transparent
         Caption         =   "&Código:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   11
      Top             =   3420
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   476
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSucursal.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSucursal.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSucursal.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSucursal.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSucursal.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSucursal.frx":0BA4
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
   Begin VB.Menu MnuBancos 
      Caption         =   "&Bancos"
      Begin VB.Menu MnuBaCodigo 
         Caption         =   "Ingresar Código de Banco"
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario   Alt+F4"
      End
   End
End
Attribute VB_Name = "frmSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private prmIDSucursal As Long
Private sNuevo, sModificar As Boolean
Private rsSuc As rdoResultset

Private Sub Form_Activate()

    Screen.MousePointer = 0
    Me.Refresh
    Foco tBanco
    
End Sub

Private Sub Form_Load()

    prmIDSucursal = 0
    EstadoIngreso False
    sNuevo = False: sModificar = False
    LimpioCampos
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.SimpleText = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If sNuevo Or sModificar Then
        If MsgBox("Antes de salir desea grabar los datos ingresados?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
            AccionGrabar
            If sNuevo Or sModificar Then Cancel = True
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    rsSuc.Close
    Forms(Forms.Count - 2).SetFocus
    
End Sub

Private Sub Label1_Click()
    Foco tBanco
End Sub

Private Sub Label2_Click()
    Foco tNombre
End Sub


Private Sub Label3_Click()
    Foco tDepositoS
End Sub

Private Sub Label5_Click()
    Foco tCuenta
End Sub

Private Sub Label6_Click()
    Foco tDepositoB
End Sub

Private Sub lCodigo_Click()
    Foco tCodigoS
End Sub

Private Sub MnuBaCodigo_Click()
    AccionCodigoBanco
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
    
    Screen.MousePointer = 11
    sNuevo = True
    
    Botones False, False, False, True, True, Toolbar1, Me
    LimpioCampos
    EstadoIngreso True
    prmIDSucursal = 0
    Screen.MousePointer = 0

End Sub

Private Sub LimpioCampos()

    tNombre.Text = ""
    tCodigoS.Text = ""
    tCuenta.Text = ""
    tDepositoB.Text = ""
    tDepositoS.Text = ""
    
End Sub

Private Sub AccionModificar()

    sModificar = True
    
    On Error GoTo Error
    Screen.MousePointer = 11
    CargoDatosSucursal prmIDSucursal
    If prmIDSucursal > 0 Then
        Botones False, False, False, True, True, Toolbar1, Me
        EstadoIngreso True
    End If
    Screen.MousePointer = 0
    Foco tBanco
    Exit Sub
    
Error:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación."
    sModificar = False
    EstadoIngreso False
End Sub

Sub AccionGrabar()

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    tBanco.SetFocus
    
    If sNuevo Then prmIDSucursal = 0 'x las dudas
    
    cons = "Select * from SucursalDeBanco Where SBaCodigo = " & prmIDSucursal
    Set rsSuc = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        
    If rsSuc.EOF Then rsSuc.AddNew Else rsSuc.Edit
        
    rsSuc!SBaBanco = Val(tBanco.Tag)
    rsSuc!SBaCodigoS = tCodigoS.Text
    rsSuc!SBaNombre = Trim(tNombre.Text)

    rsSuc!SBaDeposito = Val(tDepositoS.Tag)

    If Trim(tCuenta.Text) <> "" Then rsSuc!SBaCuentaCGSA = tCuenta.Text Else rsSuc!SBaCuentaCGSA = Null

    rsSuc.Update
    rsSuc.Close
        
    LimpioCampos
    Botones True, False, False, False, False, Toolbar1, Me
    EstadoIngreso False
    
    sNuevo = False: sModificar = False
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    On Error Resume Next
    rsSuc.Close
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Sub AccionEliminar()

    On Error GoTo errTest
    'Valido que la sucursal no este asignada a los cheques.--------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    cons = "Select * from SucursalDeBanco Where SBaDeposito = " & prmIDSucursal
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        Screen.MousePointer = 0
        MsgBox "La sucursal no se puede eliminar debido a que está asignada como sucursal de depósito para algunos de los bancos.", vbInformation, "ATENCIÓN"
        rsAux.Close: Exit Sub
    End If
    rsAux.Close
        
    cons = "Select * from ChequeDiferido Where CDiSucursal = " & prmIDSucursal
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        Screen.MousePointer = 0
        MsgBox "No es posible eliminarla debido a que hay cheques ingresados que pertenecen a la sucursal.", vbInformation, "ATENCIÓN"
        rsAux.Close: Exit Sub
    End If
    rsAux.Close
    '--------------------------------------------------------------------------------------------------.--------------------------------------------------------------
    
    Screen.MousePointer = 0
    If MsgBox("Confirma eliminar la sucursal " & Trim(tNombre.Text), vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Exit Sub
    
    On Error GoTo Error
    Screen.MousePointer = 11
    'Verifico que no existan artículos con ese tipo.
    cons = "Select * from SucursalDeBanco Where SBaCodigo = " & prmIDSucursal
    Set rsSuc = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsSuc.EOF Then rsSuc.Delete
    rsSuc.Close
    
    LimpioCampos
    Botones True, False, False, False, False, Toolbar1, Me
    Screen.MousePointer = 0
    Exit Sub
    
errTest:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar el control de foráneas."
    Exit Sub
Error:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación."
End Sub

Sub AccionCancelar()

    LimpioCampos
    
    If sNuevo Then
        Botones True, False, False, False, False, Toolbar1, Me
    Else
        If prmIDSucursal <> 0 Then
            CargoDatosSucursal prmIDSucursal
            Botones True, True, True, False, False, Toolbar1, Me
        End If
    End If
    
    EstadoIngreso False
    sNuevo = False: sModificar = False
    
End Sub

Private Sub CargoDatosSucursal(Codigo As Long)

    On Error GoTo errCargar
    Screen.MousePointer = 11
    cons = "Select * from SucursalDeBanco Where SBaCodigo = " & Codigo
    Set rsSuc = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsSuc.EOF Then
        
        tNombre.Text = Trim(rsSuc!SBaNombre)
        tCodigoS.Text = rsSuc!SBaCodigoS
        
        tDepositoB.Text = ""
        If Not IsNull(rsSuc!SBaDeposito) Then
             
             cons = "Select * From SucursalDeBanco, BancoSSFF " & _
                        " Where SBaBanco = BanCodigo " & _
                        " And SBaCodigo = " & rsSuc!SBaDeposito
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then
                tDepositoB.Text = Trim(rsAux!BanNombre)
                tDepositoB.Tag = rsAux!BanCodigo
            
                tDepositoS.Text = Trim(rsAux!SBaNombre)
                tDepositoS.Tag = rsAux!SBaCodigo
            
            End If
            rsAux.Close
        End If
        
        If Not IsNull(rsSuc!SBaCuentaCGSA) Then tCuenta.Text = Trim(rsSuc!SBaCuentaCGSA) Else tCuenta.Text = ""
        
    Else
        MsgBox "No hay datos para la sucursal seleccionada.", vbExclamation, "ATENCIÓN"
        prmIDSucursal = 0
    End If
    rsSuc.Close

    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    On Error Resume Next
    rsSuc.Close
    clsGeneral.OcurrioError "Error al cargar la sucursal.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tBanco_KeyDown(KeyCode As Integer, Shift As Integer)

    'Lista de Ayuda de Sucursales       Para cargar Datos p/modificar
    If Val(tBanco.Tag) = 0 Then Exit Sub
    If sNuevo Or sModificar Then Exit Sub
    
    If KeyCode = vbKeyF1 Then
        On Error GoTo errBuscar
        cons = "Select SBaCodigo, SBaCodigoS, SBaNombre, SBaCuentaCGSA " _
               & " From SucursalDeBanco Where SBaBanco = " & Val(tBanco.Tag) _
               & " Order by SBaNombre"
        
        Dim aObj As New clsListadeAyuda
        prmIDSucursal = aObj.ActivarAyuda(cBase, cons, 6200, 1, "Lista de Sucursales")
        Me.Refresh
        If prmIDSucursal <> 0 Then prmIDSucursal = aObj.RetornoDatoSeleccionado(0)
        Set aObj = Nothing
 
        If prmIDSucursal > 0 Then
            CargoDatosSucursal prmIDSucursal
            Botones True, True, True, False, False, Toolbar1, Me
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Error al acceder a la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCodigoS_GotFocus()
    tCodigoS.SelStart = 0
    tCodigoS.SelLength = Len(tCodigoS.Text)
    Status.SimpleText = "Ingrese el código de la sucursal."
End Sub

Private Sub tCodigoS_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Trim(tCodigoS.Text) <> "" Then Foco tNombre
    
End Sub

Private Sub tCuenta_GotFocus()
    tCuenta.SelStart = 0
    tCuenta.SelLength = Len(tCuenta.Text)
    Status.SimpleText = "Ingrese el número de cuenta para depósitos."
End Sub

Private Sub tCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDepositoB
End Sub

Private Sub tDepositoB_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tDepositoB.Text) = "" Then Exit Sub
        If Val(tDepositoB.Tag) <> 0 Then
            Foco tDepositoS: Exit Sub
        End If
        
        BuscoBanco tDepositoB
    End If
    
End Sub

Private Sub tDepositoS_Change()
    If Val(tDepositoS.Tag) <> 0 Then tDepositoS.Tag = 0
End Sub

Private Sub tNombre_GotFocus()

    tNombre.SelStart = 0
    tNombre.SelLength = Len(tNombre.Text)
    Status.SimpleText = "Ingrese el nombre de la sucursal."
    
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tNombre.Text) <> "" Then Foco tCuenta
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select

End Sub

Private Sub EstadoIngreso(Optional bEstado As Boolean = True)

On Error Resume Next
Dim mBkColor As Long

    If bEstado Then mBkColor = Colores.Blanco Else mBkColor = Colores.Inactivo
    
    tNombre.Enabled = bEstado: tNombre.BackColor = mBkColor
    tCodigoS.Enabled = bEstado: tCodigoS.BackColor = mBkColor
    tCuenta.Enabled = bEstado: tCuenta.BackColor = mBkColor
    tDepositoB.Enabled = bEstado: tDepositoB.BackColor = mBkColor
    tDepositoS.Enabled = bEstado: tDepositoS.BackColor = mBkColor
    
End Sub

Private Function ValidoCampos()

    On Error GoTo errValido
    ValidoCampos = False
    
    If Val(tBanco.Tag) = 0 Then
        MsgBox "Debe seleccionar el banco para ingresar la sucursal.", vbExclamation, "ATENCIÓN"
        Foco tBanco: Exit Function
    End If
    
    If Trim(tNombre.Text) = "" Then
        MsgBox "Debe ingresar el nombre de la sucursal.", vbExclamation, "ATENCIÓN"
        Foco tNombre: Exit Function
    End If
   
    If Not IsNumeric(tCodigoS.Text) Then
        MsgBox "Debe ingresar el código de la sucursal.", vbExclamation, "ATENCIÓN"
        Foco tCodigoS: Exit Function
    End If
   
    If Val(tDepositoB.Tag) = 0 Then
        MsgBox "Debe seleccionar el banco para realizar los depósitos de los cheques.", vbExclamation, "ATENCIÓN"
        Foco tDepositoB: Exit Function
    End If
    
    If Val(tDepositoS.Tag) = 0 Then
        MsgBox "Debe seleccionar la sucursal para realizar los depósitos de los cheques.", vbExclamation, "ATENCIÓN"
        Foco tDepositoS: Exit Function
    End If
    
    If Not clsGeneral.TextoValido(tNombre.Text) Then
        MsgBox "Se ha ingresado un caracter no válido en el nombre de la sucursal.", vbExclamation, "ATENCION"
        Foco tNombre: Exit Function
    End If
    
    If Not clsGeneral.TextoValido(tCuenta.Text) Then
        MsgBox "Se ha ingresado un caracter no válido en el número de cuenta.", vbExclamation, "ATENCION"
        Foco tCuenta: Exit Function
    End If
    
    'Valido el codigo ingresado de la sucursal
    cons = "Select * from SucursalDeBanco " _
           & " Where SBaCodigoS = " & tCodigoS.Text _
           & " And SBaBanco  = " & Val(tBanco.Tag) & " And SBaCodigo <> " & prmIDSucursal
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        MsgBox "La sucursal " & Trim(rsAux!SBaNombre) & " ya está ingresada con el código " & Format(rsAux!SBaCodigoS, "000") & Chr(vbKeyReturn) _
                & "Verifique los datos ingresados.", vbExclamation, "ATENCIÓN"
        rsAux.Close: Foco tCodigoS: Exit Function
    End If
    rsAux.Close
    
    ValidoCampos = True
    Exit Function

errValido:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al validar los datos ingresados."
End Function

Private Sub tDepositoB_Change()
    If Val(tDepositoB.Tag) <> 0 Then
        tDepositoB.Tag = 0
        tDepositoS.Text = "": tDepositoS.Tag = ""
    End If
End Sub

Private Sub tDepositoB_GotFocus()
    tDepositoB.SelStart = 0: tDepositoB.SelLength = Len(tDepositoB.Text)
    Status.SimpleText = "Seleccione el banco para realizar los depósitos."
End Sub

Private Sub tDepositoS_GotFocus()
    tDepositoS.SelStart = 0: tDepositoS.SelLength = Len(tDepositoS.Text)
    Status.SimpleText = "Seleccione la sucursal para realizar los depósitos."
End Sub

Private Sub tDepositoS_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Val(tDepositoS.Tag) <> 0 Then AccionGrabar: Exit Sub
    If Trim(tDepositoS.Text) = "" Then Exit Sub
    If KeyAscii = vbKeyReturn And Val(tDepositoB.Tag) <> 0 Then

        On Error GoTo errBuscaG
        Screen.MousePointer = 11
                   
        cons = "Select SBaCodigo, SBaCodigoS as Codigo, SBaNombre as Sucursal From SucursalDeBanco " & _
                  " Where SBaBanco = " & Val(tDepositoB.Tag)
        If IsNumeric(tDepositoS.Text) Then
            cons = cons & " And SBaCodigoS = " & Val(tDepositoS.Text)
        Else
            cons = cons & " And SBaNombre like '" & Replace(Trim(tDepositoS.Text), " ", "%") & "%'"
        End If
        cons = cons & " Order by SBaNombre"
        
        Dim aQ As Integer, aIdSelect As Long, aTexto As String
        aQ = 0: aIdSelect = 0
        
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            aQ = 1
            aIdSelect = rsAux!SBaCodigo: aTexto = Trim(rsAux!Sucursal)
            rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
        End If
        rsAux.Close
        
        Select Case aQ
            Case 0: MsgBox "No hay datos que coincidan con el texto ingersado.", vbExclamation, "No hay datos"
            
            Case 2:
                        Dim miLista As New clsListadeAyuda
                        aIdSelect = miLista.ActivarAyuda(cBase, cons, 4000, 1, "Lista de Sucursales")
                        Me.Refresh
                        If aIdSelect > 0 Then
                            aIdSelect = miLista.RetornoDatoSeleccionado(0)
                            aTexto = miLista.RetornoDatoSeleccionado(2)
                        End If
                        Set miLista = Nothing
        End Select
            
        If aIdSelect > 0 Then
            tDepositoS.Text = aTexto
            tDepositoS.Tag = aIdSelect
        End If
    
        Screen.MousePointer = 0
   End If
    Exit Sub
errBuscaG:
    clsGeneral.OcurrioError "Error al buscar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function AccionCodigoBanco()

    On Error GoTo errUpdate
    
    If Val(tBanco.Tag) = 0 Then
        MsgBox "Seleccione un banco para ingresar el código.", vbExclamation, "Falta Banco"
        Exit Function
    End If
    
    Dim mIdBanco As Long
    mIdBanco = Val(tBanco.Tag)
    
    Dim mRet As String
    
    mRet = InputBox("Ingrese el código de banco (de 2 dígitos) para el banco " & Trim(tBanco.Text), "Código de Banco", "00")
    If mRet = "" Then Exit Function
    
    If MsgBox("Confirma asignar el Código " & mRet & " al " & Trim(tBanco.Text), vbQuestion + vbYesNo, "Grabar") = vbNo Then Exit Function
    
    cons = "Select * from BancoSSFF " & _
                " Where BanCodigo = " & mIdBanco
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsAux.Edit
    rsAux!BanCodigoB = mRet
    rsAux.Update
    rsAux.Close
    
    Exit Function
    
errUpdate:
    clsGeneral.OcurrioError "Error al grabar el código de banco.", Err.Description
    Screen.MousePointer = 0
End Function


Private Sub tBanco_Change()

    If Val(tBanco.Tag) <> 0 Then tBanco.Tag = 0
        
    If Not sNuevo And Not sModificar Then
        Botones True, False, False, False, False, Toolbar1, Me
        prmIDSucursal = 0
    End If

End Sub

Private Sub tBanco_GotFocus()
    tBanco.SelStart = 0: tBanco.SelLength = Len(tBanco.Text)
    Status.SimpleText = "Seleccione el banco para ingresar la sucursal      [F1] - Ayuda."
End Sub

Private Sub tBanco_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tBanco.Text) = "" Then Exit Sub
        If Val(tBanco.Tag) <> 0 Then
            Foco tCodigoS: Exit Sub
        End If
        
        BuscoBanco tBanco
    End If
        
End Sub

Private Sub tBanco_LostFocus()
    Status.SimpleText = ""
End Sub

Private Function BuscoBanco(mControl As TextBox)

    On Error GoTo errBuscaG
    Screen.MousePointer = 11
               
    cons = "Select BanCodigo, BanNombre as Banco From BancoSSFF "
    If IsNumeric(mControl.Text) Then
        cons = cons & " Where BanCodigoB = " & Val(mControl.Text)
    Else
        cons = cons & "Where BanNombre like '" & Replace(Trim(mControl.Text), " ", "%") & "%'"
    End If
    cons = cons & " Order by BanNombre"
    
    Dim aQ As Integer, aIdSelect As Long, aTexto As String
    aQ = 0: aIdSelect = 0
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        aQ = 1
        aIdSelect = rsAux!BanCodigo: aTexto = Trim(rsAux!Banco)
        rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
    End If
    rsAux.Close
        
    Select Case aQ
        Case 0: MsgBox "No hay datos que coincidan con el texto ingersado.", vbExclamation, "No hay datos"
        
        Case 2:
                    Dim miLista As New clsListadeAyuda
                    aIdSelect = miLista.ActivarAyuda(cBase, cons, 4000, 1, "Lista de Bancos")
                    Me.Refresh
                    If aIdSelect > 0 Then
                        aIdSelect = miLista.RetornoDatoSeleccionado(0)
                        aTexto = miLista.RetornoDatoSeleccionado(1)
                    End If
                    Set miLista = Nothing
    End Select
        
    If aIdSelect > 0 Then
        mControl.Text = aTexto
        mControl.Tag = aIdSelect
        Foco mControl
    End If
    
    Screen.MousePointer = 0
   
    Exit Function
errBuscaG:
    clsGeneral.OcurrioError "Error al buscar los datos.", Err.Description
    Screen.MousePointer = 0
End Function

