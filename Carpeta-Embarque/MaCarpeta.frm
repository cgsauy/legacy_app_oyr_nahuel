VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form MaCarpeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carpeta"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MaCarpeta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5070
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "boquilla"
            Object.ToolTipText     =   "Pedidos de Boquilla"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2100
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CheckBox chAnular 
      Alignment       =   1  'Right Justify
      Caption         =   "Carpeta A&nulada:"
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   3900
      Width           =   1815
   End
   Begin AACombo99.AACombo cBcoEmisor 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   4185
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cIncoterm 
      Height          =   315
      Left            =   1680
      TabIndex        =   19
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
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
   Begin AACombo99.AACombo cFormaPago 
      Height          =   315
      Left            =   1680
      TabIndex        =   15
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
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
   Begin AACombo99.AACombo cBcoCorresponsal 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1680
      Width           =   3135
      _ExtentX        =   5530
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
   Begin AACombo99.AACombo cProveedor 
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   3135
      _ExtentX        =   5530
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
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   1680
      MaxLength       =   230
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox tPlazo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   17
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox tLC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   13
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox tFactura 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   11
      Text            =   ".0"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox tFApertura 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      MaxLength       =   12
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox tCodigo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
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
            Picture         =   "MaCarpeta.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaCarpeta.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaCarpeta.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaCarpeta.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaCarpeta.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaCarpeta.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaCarpeta.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaCarpeta.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaCarpeta.frx":10E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaCarpeta.frx":13FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lAnulada 
      BackStyle       =   0  'Transparent
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
      Left            =   2280
      TabIndex        =   26
      Top             =   3900
      Width           =   1755
   End
   Begin VB.Label labCosteada 
      BackStyle       =   0  'Transparent
      Caption         =   "Costeada:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Comen&tario:"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "&Incoterm:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Pla&zo:"
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Forma de Pago:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&L/C:"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Factura:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Bco. &Emisor:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Bco. Corresponsal:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Proveedor:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Apertura:"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Código:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   855
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
      Begin VB.Menu MnuLineaBoquilla 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBoquilla 
         Caption         =   "Pedidos de &Boquilla"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "MaCarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modificaciones
'   13-1-2004 Valido que ingresen la forma de pago, antes estaba obligatorio pero podían ingresar nulo.
'   9-5-2006    no dejo ingresar pago anticipado cuando es con lc.
Option Explicit

Private RsCarpeta As rdoResultset

'Booleanas.--------------------------------------------------
Private sNuevo As Boolean, sModificar As Boolean

Private sPedido As String

'Property.------------------------------------------------------
Private iSeleccionado As Long           'Indica el tipo de Llamado, 0 = Normal, -1 = Nuevo, n = Código de Carpeta.
Private iBoquilla As Long                 'Indica el código del pedido de boquilla de una carpeta nueva.
Public Property Get pCodigosBoquilla() As String
    pCodigosBoquilla = sPedido
End Property
Public Property Get pSeleccionado() As Long
    pSeleccionado = iSeleccionado
End Property
Public Property Let pSeleccionado(Codigo As Long)
    iSeleccionado = Codigo
End Property
Public Property Get pCodBoquilla() As Long
    pCodBoquilla = iBoquilla
End Property
Public Property Let pCodBoquilla(Codigo As Long)
    iBoquilla = Codigo
End Property
Private Sub cBcoCorresponsal_GotFocus()
    With cBcoCorresponsal
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Seleccione el banco corresponsal."
End Sub
Private Sub cBcoCorresponsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tFactura.SetFocus
End Sub
Private Sub cBcoCorresponsal_LostFocus()
    cBcoCorresponsal.SelStart = 0
    Status.SimpleText = ""
End Sub
Private Sub cBcoEmisor_GotFocus()
    With cBcoEmisor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Seleccione el banco emisor."
End Sub
Private Sub cBcoEmisor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cBcoCorresponsal.SetFocus
End Sub
Private Sub cBcoEmisor_LostFocus()
    cBcoEmisor.SelStart = 0
    Status.SimpleText = ""
End Sub
Private Sub cFormaPago_Change()
    If cFormaPago.ListIndex > -1 Then
        If cFormaPago.ItemData(cFormaPago.ListIndex) = cFPPlazoBL Then
            tPlazo.Enabled = True: tPlazo.BackColor = vbWhite
        Else
            tPlazo.Enabled = False: tPlazo.BackColor = Inactivo: tPlazo.Text = ""
        End If
    Else
        tPlazo.Enabled = False: tPlazo.BackColor = Inactivo: tPlazo.Text = ""
    End If
End Sub
Private Sub cFormaPago_GotFocus()
    cFormaPago.SelStart = 0
    cFormaPago.SelLength = Len(cFormaPago.Text)
    Status.SimpleText = " Seleccione una forma de pago."
End Sub
Private Sub cFormaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tPlazo.Enabled Then
            Foco tPlazo
        Else
            Foco cIncoterm
        End If
    End If
End Sub
Private Sub cFormaPago_LostFocus()
    cFormaPago.SelLength = 0
    Status.SimpleText = ""
    If cFormaPago.ListIndex > -1 Then
        If cFormaPago.ItemData(cFormaPago.ListIndex) = cFPPlazoBL Then
            tPlazo.Enabled = True: tPlazo.BackColor = vbWhite: Foco tPlazo
        Else
            tPlazo.Enabled = False: tPlazo.BackColor = Inactivo: tPlazo.Text = ""
        End If
    Else
        tPlazo.Enabled = False: tPlazo.BackColor = Inactivo: tPlazo.Text = ""
    End If
End Sub
Private Sub cIncoterm_GotFocus()
    With cIncoterm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Seleccione un incoterm."
End Sub
Private Sub cIncoterm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentario
End Sub
Private Sub cIncoterm_LostFocus()
    cIncoterm.SelStart = 0
    Status.SimpleText = ""
End Sub
Private Sub cProveedor_GotFocus()
    With cProveedor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Seleccione un proveedor de la carpeta o para buscar pedidos."
End Sub
Private Sub cProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cBcoEmisor
End Sub
Private Sub cProveedor_LostFocus()
    cProveedor.SelStart = 0
    Status.SimpleText = ""
    'Veo si el proveedor tiene pedidos pendientes y cargo los últimos datos
    ' de sus carpetas.
    If sNuevo And cProveedor.ListIndex > -1 Then VefificoPedidos
End Sub

Private Sub chAnular_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub Form_Activate()
    If iSeleccionado = -1 And Not sNuevo Then AccionNuevo
    Me.Refresh
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    'Inicializo-----------------------------------
    sNuevo = False: sModificar = False
    
    'Doy formato a los campos.-----------
    InicializoCampos
    InhabilitoCampos
    
    'Cargo combos.-------------------------
    Cons = "Select IncCodigo, IncNombre From IncoTerm Order by IncNombre"
    CargoCombo Cons, cIncoterm, ""
    
    cFormaPago.AddItem "Anticipado"
    cFormaPago.ItemData(cFormaPago.NewIndex) = cFPAnticipado
    cFormaPago.AddItem "Cobranza"
    cFormaPago.ItemData(cFormaPago.NewIndex) = cFPCobranza
    cFormaPago.AddItem "PlazoBL"
    cFormaPago.ItemData(cFormaPago.NewIndex) = cFPPlazoBL
    cFormaPago.AddItem "Vista"
    cFormaPago.ItemData(cFormaPago.NewIndex) = cFPVista
    
    'Proveedores.--------------------------
    Cons = "Select PExCodigo, PExNombre From ProveedorExterior Order by PExNombre"
    CargoCombo Cons, cProveedor, ""
    '----------------------------------------------
    
    'Bancos Emisores.----------------------
    Cons = "Select BLoCodigo, BLoNombre From BancoLocal Order by BLoNombre"
    CargoCombo Cons, cBcoEmisor, ""
    '----------------------------------------------
    
    'Bancos Corresponsales.--------------
    Cons = "Select BExCodigo, BExNombre From BancoExterior Order by BExNombre"
    CargoCombo Cons, cBcoCorresponsal, ""
    '----------------------------------------------
    'Inicializo Resultset.---------------------
    PidoConsulta iSeleccionado
    If iSeleccionado > 0 Then
        If Not RsCarpeta.EOF Then
            Botones True, True, True, False, False, Toolbar1, Me
            CargoCarpeta
        End If
    End If

    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrio un error iniciar el formulario.", Trim(Err.Description)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    RsCarpeta.Close
    Forms(Forms.Count - 2).SetFocus
End Sub
Private Sub Label1_Click()
    Foco tCodigo
End Sub
Private Sub Label10_Click()
    Foco cIncoterm
End Sub
Private Sub Label11_Click()
    Foco tComentario
End Sub

Private Sub Label2_Click()
    Foco tFApertura
End Sub
Private Sub Label3_Click()
    Foco cProveedor
End Sub
Private Sub Label4_Click()
    Foco cBcoCorresponsal
End Sub
Private Sub Label5_Click()
    Foco cBcoEmisor
End Sub
Private Sub Label6_Click()
    Foco tFactura
End Sub
Private Sub Label7_Click()
    Foco tLC
End Sub
Private Sub Label8_Click()
    Foco cFormaPago
End Sub
Private Sub Label9_Click()
    Foco tPlazo
End Sub

Private Sub MnuBoquilla_Click()
    AccionPedidoBoquilla
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
On Error GoTo ErrAN
    Screen.MousePointer = 0
    sNuevo = True
    sPedido = ""
    'Limpio y habilito campos.-------------------
    LimpioCampos
    HabilitoCampos
    'Prendo Señal que es uno nuevo.--------
    iBoquilla = 0
    'Habilito y Desabilito Botones.-------------
    Botones False, False, False, True, True, Toolbar1, Me
'    MnuBoquilla.Enabled = False: Toolbar1.Buttons("boquilla").Enabled = False
    'Limpio el seleccionado.---------------------
    If iSeleccionado > -1 Then iSeleccionado = 0
    'Sugiero el código de carpeta.--------------
    Cons = "Select MAX(CarCodigo) From Carpeta Where CarCodigo < 9000"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux(0)) Then tCodigo.Text = RsAux(0) + 1 Else tCodigo.Text = 1
    Else
        tCodigo.Text = 1
    End If
    RsAux.Close
    tPlazo.Enabled = False: tPlazo.BackColor = Inactivo: tPlazo.Text = ""
    tFApertura.Text = Format(Date, FormatoFP)
    tCodigo.SetFocus
    Screen.MousePointer = 0
    Exit Sub
ErrAN:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error en acción nuevo."
End Sub
Private Sub AccionModificar()

    'Me quedo con el ID.----------------------
    iSeleccionado = RsCarpeta!CarID
    BuscoCarpeta iSeleccionado
    
    If Not RsCarpeta.EOF Then
        'Prendo señal.-----------------------------
        sModificar = False
        'Habilito y Desabilito Botones.----------
        Botones False, False, False, True, True, Toolbar1, Me
        Toolbar1.Buttons("boquilla").Enabled = False: MnuBoquilla.Enabled = False
        HabilitoCampos
    End If
    If cFormaPago.ListIndex > -1 Then
        If cFormaPago.ItemData(cFormaPago.ListIndex) = cFPPlazoBL Then
            tPlazo.Enabled = True: tPlazo.BackColor = vbWhite
        Else
            tPlazo.Enabled = False: tPlazo.BackColor = Inactivo: tPlazo.Text = ""
        End If
    Else
        tPlazo.Enabled = False: tPlazo.BackColor = Inactivo: tPlazo.Text = ""
    End If
    
End Sub

Private Sub AccionGrabar()

    If MsgBox("Confirma almacenar los datos ingresados?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
        If Not ValidoCampos Then MsgBox "Los datos ingresados no son correctos.", vbExclamation, "ATENCIÓN": Exit Sub
        If sNuevo Then NuevaCarpeta Else ModificoCarpeta
    End If
    
End Sub

Private Sub AccionEliminar()
On Error GoTo ErrAE
    If MsgBox("¿Confirma eliminar la carpeta seleccionada?", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        Screen.MousePointer = 11
        Cons = "Select * From Embarque Where EmbCarpeta = " & RsCarpeta!CarID
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            RsAux.Close
            'No tiene embarques.--------------
            Cons = "Select * From Pedido Where PedCarpeta = " & RsCarpeta!CarID
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                MsgBox "Se eliminará la asociación del pedido de boquilla a esta carpeta.", vbInformation, "ATENCIÓN"
                RsAux.Edit
                RsAux!PedCarpeta = Null
                RsAux.Update
            End If
            RsAux.Close
            RsCarpeta.Delete
            iSeleccionado = 0
            BuscoCarpeta 0
        Else
            RsAux.Close
            MsgBox "La carpeta esta asociada  a embarques, elimine los embarques.", vbCritical, "ATENCIÓN"
            Screen.MousePointer = 0
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrAE:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al intentar eliminar la carpeta."
End Sub

Private Sub AccionCancelar()
    'Apago señal.----------------
    sNuevo = False: sModificar = False: sPedido = ""
    'Actualizo el formulario.---
    InhabilitoCampos
    BuscoCarpeta iSeleccionado
End Sub
Private Sub AccionPedidoBoquilla()
On Error GoTo ErrPB
Dim rsAyuda As rdoResultset
    
    Cons = "Select PedCodigo, Código = PedCodigo, Fecha = PedFPedido, Embarca=PedFEmbarque, Proveedor = EExNombre, Articulo = ArtNombre, IsNull(PedComentario, '') as Memo " _
        & " From Pedido, EmpresaExterior, Articulo, ArticuloFolder " _
        & " Where PedProveedor = EExCodigo" _
        & " And AFoTipo = " & cFPedido _
        & " And PedCodigo = AFoCodigo And PedCarpeta Is NULL" _
        & " And AFoArticulo = ArtID"
        
    If sPedido <> "" Then Cons = Cons & " And PedCodigo Not In (" & sPedido & ")"
        
    If cProveedor.ListIndex <> -1 Then Cons = Cons & " And PedProveedor = " & cProveedor.ItemData(cProveedor.ListIndex)
    
    Set rsAyuda = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If rsAyuda.EOF Then
        rsAyuda.Close
        If sPedido <> "" Then
            MsgBox "No hay más pedidos pendientes para el proveedor.", vbInformation, "ATENCIÓN"
        Else
            If Not sNuevo Then
                LimpioCampos
                MsgBox "No hay pedidos pendientes.", vbInformation, "ATENCIÓN"
            End If
        End If
    Else
        rsAyuda.Close
        Dim objAyuda As New clsListadeAyuda
        If objAyuda.ActivarAyuda(cBase, Cons, 8000, 1, "Lista de Pedidos") Then
'        objAyuda.ActivoListaAyuda Cons, False, miconexion.TextoConexion(logImportaciones), 7500
'        If objAyuda.ValorSeleccionado > 0 Then
                If Not sNuevo Then
                    AccionNuevo
                    If iBoquilla = 0 Then CargoInformacionBoquilla objAyuda.RetornoDatoSeleccionado(0)
                Else
                    If sPedido = "" Then CargoInformacionBoquilla objAyuda.RetornoDatoSeleccionado(0)
                End If
                If sPedido <> "" Then
                    sPedido = sPedido & ", " & objAyuda.RetornoDatoSeleccionado(0)
                Else
                    iBoquilla = objAyuda.RetornoDatoSeleccionado(0)
                    sPedido = objAyuda.RetornoDatoSeleccionado(0)
                End If
        End If
        Set objAyuda = Nothing
    End If
    
    'Instancio al formulario de ayuda.---------------------------------------------
'    Dim frmAyuda As New HeCarpeta
    
'    Select Case frmAyuda.PidoConsulta(Cons)
'        Case 1
'            frmAyuda.CierroCursor
'            If Not sNuevo Then
'                LimpioCampos
'                MsgBox "No hay pedidos pendientes.", vbInformation, "ATENCIÓN"
'            End If
            
'        Case 2
'            frmAyuda.Show vbModal
'            Screen.MousePointer = 11
'            Me.Refresh
''            If frmAyuda.pSeleccionado > 0 Then
 '               If Not sNuevo Then
 '                   AccionNuevo
 '                   CargoInformacionBoquilla frmAyuda.pSeleccionado
 '               End If
 ''               iBoquilla = frmAyuda.pSeleccionado
  '          End If
  '  End Select
  '  Set frmAyuda = Nothing
    '.-------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    Exit Sub
ErrPB:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese el código de Carpeta."
End Sub
Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If sNuevo Or sModificar Then tFApertura.SetFocus Else BuscoCarpetaPorCodigo CLng(tCodigo.Text)
    End If
End Sub
Private Sub tCodigo_LostFocus()
    Status.SimpleText = ""
End Sub
Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese un comentario para la carpeta."
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tComentario_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub tFactura_GotFocus()
    With tFactura
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese el número de factura."
End Sub
Private Sub tFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And tLC.Enabled Then tLC.SetFocus
End Sub
Private Sub tFactura_LostFocus()
    Status.SimpleText = ""
End Sub
Private Sub tFApertura_GotFocus()
    With tFApertura
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese la fecha de apertura."
End Sub
Private Sub tFApertura_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cProveedor.Enabled Then cProveedor.SetFocus
End Sub
Private Sub tFApertura_LostFocus()
    If IsDate(tFApertura.Text) Then tFApertura.Text = Format(tFApertura.Text, "d-Mmm-yyyy")
    Status.SimpleText = ""
End Sub

Private Sub tLC_GotFocus()
    With tLC
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese la carta de crédito."
End Sub
Private Sub tLC_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cFormaPago.Enabled Then cFormaPago.SetFocus
End Sub
Private Sub tLC_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)

    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "boquilla": AccionPedidoBoquilla
        Case "salir": Unload Me
    End Select

End Sub
Private Sub InicializoCampos()
    tCodigo.BackColor = Colores.Obligatorio
    tFApertura.BackColor = Colores.Obligatorio
    cProveedor.BackColor = Colores.Obligatorio
    cFormaPago.BackColor = Colores.Obligatorio
End Sub
Private Sub InhabilitoCampos()
    tFApertura.Enabled = False: tFApertura.BackColor = Colores.Inactivo
    cProveedor.Enabled = False: cProveedor.BackColor = Colores.Inactivo
    cBcoEmisor.BackColor = Colores.Inactivo
    cBcoEmisor.Enabled = False
    cBcoCorresponsal.Enabled = False
    cBcoCorresponsal.BackColor = Colores.Inactivo
    tFactura.Enabled = False: tFactura.BackColor = Colores.Inactivo
    tLC.Enabled = False: tLC.BackColor = Colores.Inactivo
    cFormaPago.Enabled = False
    cFormaPago.BackColor = Colores.Inactivo
    tPlazo.Enabled = False: tPlazo.BackColor = Colores.Inactivo
    cIncoterm.Enabled = False
    cIncoterm.BackColor = Colores.Inactivo
    tComentario.Enabled = False: tComentario.BackColor = Colores.Inactivo
    chAnular.Enabled = False
End Sub
Private Sub HabilitoCampos()
On Error GoTo ErrHC
    tFApertura.Enabled = True: tFApertura.BackColor = Colores.Obligatorio
    cProveedor.Enabled = True: cProveedor.BackColor = Colores.Obligatorio
    cBcoEmisor.BackColor = vbWhite
    cBcoEmisor.Enabled = True
    cBcoCorresponsal.Enabled = True
    cBcoCorresponsal.BackColor = vbWhite
    tFactura.Enabled = True: tFactura.BackColor = vbWhite
    tLC.Enabled = True: tLC.BackColor = vbWhite
    cFormaPago.Enabled = True
    cFormaPago.BackColor = Colores.Obligatorio
    tPlazo.Enabled = True: tPlazo.BackColor = vbWhite
    cIncoterm.Enabled = True
    cIncoterm.BackColor = vbWhite
    tComentario.Enabled = True: tComentario.BackColor = vbWhite
    chAnular.Enabled = True
    Exit Sub
ErrHC:
    clsGeneral.OcurrioError "Ocurrio un error inesperado."
End Sub
Private Sub LimpioCampos()
    tCodigo.Text = ""
    tFApertura.Text = ""
    cProveedor.Text = ""
    cBcoEmisor.Text = ""
    cBcoCorresponsal.Text = ""
    tFactura.Text = ""
    tLC.Text = ""
    cFormaPago.Text = ""
    tPlazo.Text = ""
    cIncoterm.Text = ""
    labCosteada.Caption = "Costeada:"
    tComentario.Text = ""
    chAnular.Value = 0: lAnulada.Caption = ""
End Sub
Private Sub tPlazo_GotFocus()
    With tPlazo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese un plazo."
End Sub
Private Sub tPlazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cIncoterm.Enabled Then cIncoterm.SetFocus
End Sub
Private Sub tPlazo_LostFocus()
    Status.SimpleText = ""
End Sub
Private Sub PidoConsulta(Codigo As Long)
    
    On Error Resume Next
    RsCarpeta.Close
    On Error GoTo ErrPC
    Cons = "Select * From Carpeta Where CarID = " & Codigo
    Set RsCarpeta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Exit Sub
    
ErrPC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al consultar."
End Sub
Private Sub BuscoCarpetaPorCodigo(ByVal Codigo As Long)
    On Error GoTo ErrBC
    Screen.MousePointer = 11
    Botones True, False, False, False, False, Toolbar1, Me
    Toolbar1.Buttons("boquilla").Enabled = True: MnuBoquilla.Enabled = True
    LimpioCampos
    RsCarpeta.Close
    Cons = "Select * From Carpeta Where CarCodigo = " & Codigo
    Set RsCarpeta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    FinalBusqueda Codigo
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al buscar la carpeta."
End Sub

Private Sub BuscoCarpeta(ByVal Codigo As Long)
    
    On Error GoTo ErrBC
    Screen.MousePointer = 11
    Botones True, False, False, False, False, Toolbar1, Me
    Toolbar1.Buttons("boquilla").Enabled = True: MnuBoquilla.Enabled = True
    LimpioCampos
    RsCarpeta.Close
    
    Cons = "Select * From Carpeta Where CarID = " & Codigo
    Set RsCarpeta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    FinalBusqueda Codigo
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al buscar la carpeta."
End Sub
Private Sub FinalBusqueda(Codigo As Long)
On Error GoTo ErrFB
    If Not RsCarpeta.EOF Then
        Botones True, True, True, False, False, Toolbar1, Me
        CargoCarpeta
        iSeleccionado = RsCarpeta!CarID
    ElseIf Codigo > 0 Then
        'Muestro mensaje si busco una carpeta.-----------------
        Screen.MousePointer = 0
        MsgBox "No existe una carpeta con ese código.", vbInformation, "ATENCIÓN"
        iSeleccionado = 0
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrFB:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al buscar la carpeta."
End Sub

Private Sub CargoCarpeta()

    tCodigo.Text = RsCarpeta!CarCodigo
    tFApertura.Text = Format(RsCarpeta!CarFApertura, "d-Mmm-yyyy")
    If Not IsNull(RsCarpeta!CarProveedor) Then BuscoCodigoEnCombo cProveedor, RsCarpeta!CarProveedor
    If Not IsNull(RsCarpeta!CarBcoEmisor) Then BuscoCodigoEnCombo cBcoEmisor, RsCarpeta!CarBcoEmisor
    If Not IsNull(RsCarpeta!CarBcoCorresponsal) Then BuscoCodigoEnCombo cBcoCorresponsal, RsCarpeta!CarBcoCorresponsal
    If Not IsNull(RsCarpeta!CarFactura) Then tFactura.Text = Trim(RsCarpeta!CarFactura)
    If Not IsNull(RsCarpeta!CarCartaCredito) Then tLC.Text = Trim(RsCarpeta!CarCartaCredito)
    If Not IsNull(RsCarpeta!CarFormaPago) Then BuscoCodigoEnCombo cFormaPago, RsCarpeta!CarFormaPago
    If Not IsNull(RsCarpeta!CarPlazo) Then tPlazo.Text = RsCarpeta!CarPlazo
    If Not IsNull(RsCarpeta!CarIncoterm) Then BuscoCodigoEnCombo cIncoterm, RsCarpeta!CarIncoterm
    If RsCarpeta!CarCosteada Then labCosteada.Caption = "Costeada: SI" Else labCosteada.Caption = "Costeada: NO"
    If Not IsNull(RsCarpeta!CarComentario) Then tComentario.Text = Trim(RsCarpeta!CarComentario)
    If Not IsNull(RsCarpeta!CarFAnulada) Then
        chAnular.Value = 1
        lAnulada.Caption = Format(RsCarpeta!CarFAnulada, "dd/mm/yy hh:mm")
    Else
        lAnulada.Caption = "": chAnular.Value = 0
    End If
    
End Sub
Private Sub NuevaCarpeta()
On Error GoTo ErrNC
Dim PidieronNueva As Boolean
    
    Screen.MousePointer = 11
    If iSeleccionado = -1 Then PidieronNueva = True Else PidieronNueva = False
    'Bloqueo para obtener el código.
    cBase.BeginTrans
    On Error GoTo ErrResumo
    RsCarpeta.AddNew
    CargoCamposBD
    RsCarpeta.Update
    Cons = "Select Max(CarID) From Carpeta"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    iSeleccionado = RsAux(0)
    RsAux.Close
    'Relaciono el pedido con la carpeta.
    If iBoquilla > 0 Then
        Cons = "Update Pedido Set PedCarpeta = " & iSeleccionado & " Where PedCodigo IN(" & sPedido & ")"
        cBase.Execute (Cons)
    End If
    cBase.CommitTrans
    'Si pidieron nueva cierro el formulario.------
    If PidieronNueva Then Unload Me: Exit Sub
    On Error Resume Next
    sNuevo = False
    InhabilitoCampos
    BuscoCarpeta iSeleccionado
    Screen.MousePointer = 0
    
    Exit Sub
ErrNC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información."
    Exit Sub
ErrResumo:
    Resume ErrTransaccion
ErrTransaccion:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información."
End Sub
Private Sub ModificoCarpeta()
On Error GoTo ErrMC
Dim FModificacion As Date
Dim sIngreso As Boolean

    Screen.MousePointer = 11
    FechaDelServidor
    FModificacion = RsCarpeta!CarFModificacion
    
    sIngreso = False
    If Not ExisteGastoDivisa Then
        'Si no tenía gasto ingresado y me ingresó el bco. emisor y la lc para pagos plazobl y vista ingreso el gasto.
        If RsCarpeta!CarFormaPago = FormaPago.cFPPlazoBL Or RsCarpeta!CarFormaPago = FormaPago.cFPVista Then
            If cBcoEmisor.ListIndex > -1 And Trim(tLC.Text) <> "" And (IsNull(RsCarpeta!CarBcoEmisor) Or IsNull(RsCarpeta!CarCartaCredito)) Then
                'Tengo que ingresar gasto.
                If MsgBox("Se ingresaran gastos de divisa para los embarques que pertenezcan a la carpeta, con serie y numero de la LC y como proveedor el Banco Emisor." & Chr(13) & "¿Confirma Grabar?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
                sIngreso = True
            Else
                If cFormaPago.ItemData(cFormaPago.ListIndex) = FormaPago.cFPAnticipado Or cFormaPago.ItemData(cFormaPago.ListIndex) = FormaPago.cFPCobranza And (IsNull(RsCarpeta!CarBcoEmisor) Or IsNull(RsCarpeta!CarCartaCredito)) Then
                    If MsgBox("Se ingresaran gastos de divisa para los embarques que pertenezcan a la carpeta, con serie C más letra de embarque y numero de la Carpeta  sin proveedor." & Chr(13) & "¿Confirma seguir con la modificación?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
                    sIngreso = True
                End If
            End If
        End If
    End If
    
'2/11/2007 saque transacción ya que separe con gasto zureo.
    On Error GoTo ErrBT
'    cBase.BeginTrans
'    On Error GoTo ErrResumir
    'Refresco el resultset para ver si fue modificado.
    RsCarpeta.Requery
    If Not RsCarpeta.EOF Then
        If FModificacion = RsCarpeta!CarFModificacion Then
            RsCarpeta.Edit
            CargoCamposBD
            RsCarpeta.Update
            If sIngreso Then IngresoGastoDivisa
        Else
            MsgBox "La carpeta seleccionada fue modificada por otra terminal, verifique.", vbInformation, "ATENCIÓN"
        End If
    Else
        MsgBox "La carpeta fue eliminada, verifique.", vbInformation, "ATENCIÓN"
        iSeleccionado = 0
    End If
    sModificar = False
    InhabilitoCampos
    BuscoCarpeta iSeleccionado
    Screen.MousePointer = 0
    Exit Sub
ErrMC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información."
    Exit Sub
ErrBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar la información.", Err.Description
    Exit Sub

End Sub
Private Function ExisteGastoDivisa() As Boolean
Dim rsEmb As rdoResultset
    
    ExisteGastoDivisa = False
    Cons = "Select * from Compra, GastoImportacion, Embarque, Carpeta " _
        & " Where GImIDSubRubro = " & paSubrubroDivisa _
        & " And GImNivelFolder = " & Folder.cFEmbarque _
        & " And CarID = " & RsCarpeta!CarID _
        & " And GImFolder = EmbID " _
        & " And ComTipoDocumento = " & TipoDocumento.CompraCredito _
        & " And GImIDCompra = ComCodigo And EmbCarpeta = CarID"
    
    Set rsEmb = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsEmb.EOF Then ExisteGastoDivisa = True
    rsEmb.Close

End Function

Private Sub IngresoGastoDivisa()
Dim rsEmb As rdoResultset
        
    'Veo si ya hay gastos ingresados.
    Cons = "Select * from Compra, GastoImportacion, Embarque, Carpeta " _
        & " Where GImIDSubRubro = " & paSubrubroDivisa _
        & " And GImNivelFolder = " & Folder.cFEmbarque _
        & " And CarID = " & RsCarpeta!CarID _
        & " And GImFolder = EmbID " _
        & " And ComTipoDocumento = " & TipoDocumento.CompraCredito _
        & " And GImIDCompra = ComCodigo And EmbCarpeta = CarID"
    
    Set rsEmb = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If rsEmb.EOF Then
        rsEmb.Close
        
        Dim iCta2 As Long
        If cBcoEmisor.ListIndex > -1 Then
            iCta2 = cBcoEmisor.ItemData(cBcoEmisor.ListIndex)
        Else
            iCta2 = 0
        End If
        
        Cons = "Select * From Embarque Where EmbCarpeta = " & RsCarpeta!CarID
        Set rsEmb = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsEmb.EOF
            
            If cFormaPago.ItemData(cFormaPago.ListIndex) = FormaPago.cFPAnticipado Or cFormaPago.ItemData(cFormaPago.ListIndex) = FormaPago.cFPCobranza Then
                InsertoGastoImportacionZureo rsEmb!EmbID, Format(rsEmb!EmbDivisa, FormatoMonedaP) / Format(rsEmb!EmbArbitraje, FormatoMonedaP), Format(gFechaServidor, FormatoFP), 0, RsCarpeta("CarCodigo") & "." & rsEmb!EmbCodigo, rsEmb!EmbCodigo, "C" + rsEmb!EmbCodigo & " " & RsCarpeta!CarCodigo, TipoDocumento.CompraCredito, rsEmb!EmbArbitraje, cBcoEmisor.Text, Trim(tLC.Text), cFormaPago.Text, paSubrubroDivisa, iCta2
                'IngresoGastoAutomatico rsEmb!EmbID, Format(rsEmb!EmbDivisa, FormatoMonedaP) / Format(rsEmb!EmbArbitraje, FormatoMonedaP), Format(gFechaServidor, FormatoFP), 0, RsCarpeta!CarCodigo & "." & rsEmb!EmbCodigo, rsEmb!EmbCodigo, "C" + rsEmb!EmbCodigo, RsCarpeta!CarCodigo, TipoDocumento.CompraCredito, rsEmb!EmbArbitraje, cBcoEmisor.Text, Trim(tLC.Text), cFormaPago.Text, paSubrubroDivisa
            Else
                'IngresoGastoAutomatico rsEmb!EmbID, Format(rsEmb!EmbDivisa, FormatoMonedaP) / Format(rsEmb!EmbArbitraje, FormatoMonedaP), Format(gFechaServidor, FormatoFP), cBcoEmisor.ItemData(cBcoEmisor.ListIndex), RsCarpeta!CarCodigo & "." & rsEmb!EmbCodigo, rsEmb!EmbCodigo, "LC", tLC.Text, TipoDocumento.CompraCredito, rsEmb!EmbArbitraje, cBcoEmisor.Text, Trim(tLC.Text), cFormaPago.Text, paSubrubroDivisa
                InsertoGastoImportacionZureo rsEmb!EmbID, Format(rsEmb!EmbDivisa, FormatoMonedaP) / Format(rsEmb!EmbArbitraje, FormatoMonedaP), Format(gFechaServidor, FormatoFP), cBcoEmisor.ItemData(cBcoEmisor.ListIndex), RsCarpeta!CarCodigo & "." & rsEmb!EmbCodigo, rsEmb!EmbCodigo, "LC " & tLC.Text, TipoDocumento.CompraCredito, rsEmb!EmbArbitraje, cBcoEmisor.Text, Trim(tLC.Text), cFormaPago.Text, paSubrubroDivisa, iCta2
            End If
            rsEmb.MoveNext
        Loop
        rsEmb.Close
    Else
        rsEmb.Close
        'Ya los ingreso
    End If
    
End Sub
Private Sub CargoCamposBD()
    
    RsCarpeta!CarCodigo = tCodigo.Text
    RsCarpeta!CarFApertura = Format(tFApertura.Text, "mm/dd/yyyy")
    If cProveedor.ListIndex > -1 Then
        RsCarpeta!CarProveedor = cProveedor.ItemData(cProveedor.ListIndex)
    Else
        RsCarpeta!CarProveedor = Null
    End If
    If cBcoEmisor.ListIndex > -1 Then
        RsCarpeta!CarBcoEmisor = cBcoEmisor.ItemData(cBcoEmisor.ListIndex)
    Else
        RsCarpeta!CarBcoEmisor = Null
    End If
    If cBcoCorresponsal.ListIndex > -1 Then
        RsCarpeta!CarBcoCorresponsal = cBcoCorresponsal.ItemData(cBcoCorresponsal.ListIndex)
    Else
        RsCarpeta!CarBcoCorresponsal = Null
    End If
    If Trim(tFactura.Text) <> vbNullString Then RsCarpeta!CarFactura = tFactura.Text Else RsCarpeta!CarFactura = Null
    If Trim(tLC.Text) <> vbNullString Then RsCarpeta!CarCartaCredito = tLC.Text Else RsCarpeta!CarCartaCredito = Null
    If cFormaPago.ListIndex > -1 Then
        RsCarpeta!CarFormaPago = cFormaPago.ItemData(cFormaPago.ListIndex)
    Else
        RsCarpeta!CarFormaPago = Null
    End If
    If Trim(tPlazo.Text) <> vbNullString Then RsCarpeta!CarPlazo = tPlazo.Text Else RsCarpeta!CarPlazo = Null
    If cIncoterm.ListIndex > -1 Then
        RsCarpeta!CarIncoterm = cIncoterm.ItemData(cIncoterm.ListIndex)
    Else
        RsCarpeta!CarIncoterm = Null
    End If
    If sNuevo Then RsCarpeta!CarCosteada = 0
    If Trim(tComentario.Text) <> vbNullString Then RsCarpeta!CarComentario = tComentario.Text Else RsCarpeta!CarComentario = Null
    RsCarpeta!CarFModificacion = Format(Now, "mm/dd/yyyy hh:mm:ss")
    If chAnular.Value = 0 Then
        RsCarpeta!CarFAnulada = Null
    Else
        If IsNull(RsCarpeta!CarFAnulada) Then RsCarpeta!CarFAnulada = Format(Now, "mm/dd/yyyy hh:mm:ss")
    End If
    
End Sub

Private Function ValidoCampos() As Boolean

    If Not IsNumeric(tCodigo.Text) Then
        tCodigo.SetFocus
        ValidoCampos = False
        Exit Function
    End If
    If Not IsDate(tFApertura.Text) Then
        tFApertura.SetFocus
        ValidoCampos = False
        Exit Function
    End If
    If cProveedor.ListIndex = -1 Then
        cProveedor.SetFocus
        ValidoCampos = False
        Exit Function
    End If
    If cFormaPago.ListIndex = -1 Then
        MsgBox "La forma de pago es un campo obligatorio.", vbInformation, "Atención"
        cFormaPago.SetFocus
        ValidoCampos = False
        Exit Function
    End If
    
    If Trim(tPlazo.Text) <> vbNullString Then
        If Not IsNumeric(tPlazo.Text) Or Val(tPlazo.Text) < 0 Then
            tPlazo.SetFocus
            ValidoCampos = False
            Exit Function
        End If
    End If
    If Trim(tLC.Text) <> "" Then
        If Not IsNumeric(tLC.Text) Then Foco tLC: ValidoCampos = False: Exit Function
        If cFormaPago.ItemData(cFormaPago.ListIndex) = cFPAnticipado Then
            MsgBox "No puede seleccionar forma de pago anticipado con LC", vbExclamation, "Posible error"
            cFormaPago.SetFocus
            ValidoCampos = False
            Exit Function
        End If
    End If
    If Not clsGeneral.TextoValido(tComentario.Text) Then
        tComentario.SetFocus
        ValidoCampos = False
        Exit Function
    End If
    ValidoCampos = True
    
End Function

Private Sub CargoInformacionBoquilla(CodPedido As Long)
On Error GoTo ErrCIB
    
    Screen.MousePointer = 11
    Cons = "Select * From Pedido Where PedCodigo = " & CodPedido
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        BuscoCodigoEnCombo cProveedor, RsAux!PedProveedor
        If Not IsNull(RsAux!PedComentario) Then tComentario.Text = Trim(RsAux!PedComentario)
        RsAux.Close
        BuscoUltimosDatos
    Else
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "Otra terminal pudo eliminar el pedido, verifique.", vbExclamation, "ATENCIÓN"
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrCIB:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al levantar la información del pedido de boquilla."
End Sub

Private Sub BuscoUltimosDatos()
On Error GoTo ErrBUD

    Cons = " Select * from Carpeta" _
            & " Where CarProveedor = " & cProveedor.ItemData(cProveedor.ListIndex) _
            & " And CarFApertura = (Select Max(CarFApertura) From Carpeta Where CarProveedor = " & cProveedor.ItemData(cProveedor.ListIndex) & ")"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic)
    If Not RsAux.EOF Then
        'Cargo forma de pago.------------------------
        If Not IsNull(RsAux!CarFormaPago) Then
            BuscoCodigoEnCombo cFormaPago, RsAux!CarFormaPago
            If RsAux!CarFormaPago = FormaPago.cFPPlazoBL And Not IsNull(RsAux!CarPlazo) Then tPlazo.Text = RsAux!CarPlazo
        End If
        'Cargo Incoterm.-------------------------------
        If Not IsNull(RsAux!CarIncoterm) Then
            BuscoCodigoEnCombo cIncoterm, RsAux!CarIncoterm
        Else
            cIncoterm.Text = "FOB"
        End If
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
ErrBUD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al buscar los últimos datos de carpeta para el proveedor del pedido de boquilla."
End Sub
Private Sub VefificoPedidos()
On Error GoTo ErrVP
    Screen.MousePointer = 11
    AccionPedidoBoquilla
    If cProveedor.ListIndex > -1 Then BuscoUltimosDatos
    Screen.MousePointer = 0
    Exit Sub
ErrVP:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error"
End Sub


