VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.0#0"; "AACOMBO.OCX"
Begin VB.Form InPedidosRep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agenda de Pedidos "
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7455
   Icon            =   "InPedidosRep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin AACombo99.AACombo cEstado 
      Height          =   315
      Left            =   840
      TabIndex        =   22
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
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
   Begin AACombo99.AACombo cAgencia 
      Height          =   315
      Left            =   840
      TabIndex        =   21
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
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
   Begin AACombo99.AACombo cTransporte 
      Height          =   315
      Left            =   5160
      TabIndex        =   20
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   4080
      TabIndex        =   19
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
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
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
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
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   3
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
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   4
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
            Object.Width           =   4750
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox tCarpeta 
      Alignment       =   1  'Right Justify
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
      Left            =   6360
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.CheckBox chArribo 
      Caption         =   "Arrib�"
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
      Left            =   6480
      TabIndex        =   14
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox chPago 
      Caption         =   "Pago"
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
      Left            =   6480
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox tFlete 
      Alignment       =   1  'Right Justify
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
      Left            =   5160
      MaxLength       =   12
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox tDetalle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox tComentario 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2880
      Width           =   7215
   End
   Begin VB.TextBox tImporte 
      Alignment       =   1  'Right Justify
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
      Left            =   5160
      MaxLength       =   12
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox tFPedido 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   840
      MaxLength       =   9
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin ComctlLib.ListView lPedido 
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Fecha"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Detalle"
         Object.Width           =   3794
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Estado"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Importe"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Arrib�"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Proveedor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Agencia"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Flete"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Pago"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Comentario"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   10
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Carpeta"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   11
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Transporte"
         Object.Width           =   0
      EndProperty
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   5475
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5398
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Transporte"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   975
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Abierto en &Carpeta"
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
      Left            =   4920
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fle&te"
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
      Left            =   4560
      TabIndex        =   12
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "E&stado"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Agencia"
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
      TabIndex        =   3
      Top             =   975
      Width           =   735
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      Caption         =   " Detalle del Pedido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   7215
   End
   Begin VB.Label lProveedor 
      BackStyle       =   0  'Transparent
      Caption         =   "Pro&veedor"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lFEmbarque 
      BackStyle       =   0  'Transparent
      Caption         =   "&Importe"
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
      Left            =   4560
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha "
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
      Top             =   600
      Width           =   615
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InPedidosRep.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InPedidosRep.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InPedidosRep.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InPedidosRep.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InPedidosRep.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InPedidosRep.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InPedidosRep.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InPedidosRep.frx":0DC8
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
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario   Alt+F4"
      End
   End
End
Attribute VB_Name = "InPedidosRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sNuevo As Boolean, sModificar As Boolean

Private RsPedido As rdoResultset

Private Sub cAgencia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.Panels(4).Text = "Seleccione la agencia asociada al pedido."
    
End Sub

Private Sub cEstado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Seleccione el estado del pedido."
    
End Sub

Private Sub cPago_KeyPress(KeyAscii As Integer)

    If KeyAscii = kEnt Then
        tFlete.SetFocus
    End If
    
End Sub

Private Sub cPago_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Indique si el pedido est� pago."
    
End Sub

Private Sub cProveedor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Seleccione el proveedor de art�culos del pedido."
    
End Sub

Private Sub cTransporte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Seleccione el medio de transporte."
    
End Sub

Private Sub chArribo_GotFocus()
    Status.Panels(4).Text = "Indique si el pedido arrib�."
End Sub

Private Sub chArribo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tComentario.SetFocus
End Sub

Private Sub Form_Activate()

    Screen.MousePointer = 0
    DoEvents
    
End Sub

Private Sub Form_Load()

    SetearLView lvValores.FullRow Or lvValores.UnClickIcono, lPedido
    
    sNuevo = False
    sModificar = False
    
    CargoProveedor
    CargoAgencia
    CargoEstado
    CargoMediosTransporte
    
    CargoPedidos
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    RsPedido.Close
    
End Sub

Private Sub Label1_Click()
    
    If cAgencia.Enabled Then
        cAgencia.SetFocus
        cAgencia.SelStart = 0
        cAgencia.SelLength = Len(cAgencia.Text)
    End If
    
End Sub

Private Sub Label2_Click()

    cEstado.SetFocus
    cEstado.SelStart = 0
    cEstado.SelLength = Len(cEstado.Text)
        
End Sub

Private Sub Label3_Click()

    tFlete.SetFocus
    tFlete.SelStart = 0
    tFlete.SelLength = Len(tFlete.Text)
    
End Sub

Private Sub Label4_Click()
    Foco tfecha
End Sub



Private Sub Label5_Click()
    
    tCarpeta.SetFocus
    tCarpeta.SelStart = 0
    tCarpeta.SelLength = Len(tCarpeta.Text)
    
End Sub

Private Sub Label6_Click()
    
    cTransporte.SetFocus
    cTransporte.SelStart = 0
    cTransporte.SelLength = Len(cTransporte.Text)

End Sub

Private Sub lFEmbarque_Click()

    tImporte.SetFocus
    tImporte.SelStart = 0
    tImporte.SelLength = Len(tImporte.Text)
    
End Sub

Private Sub lPedido_ItemClick(ByVal Item As ComctlLib.ListItem)
    
    If lPedido.SelectedItem.Index <> -1 Then
        CargoDatosDesdeLista
        Call Botones(True, True, True, False, False, Toolbar1, Me)
    Else
        Call Botones(True, False, False, False, False, Toolbar1, Me)
    End If
    
End Sub

Private Sub lPedido_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Lista de pedidos realizados. (para ver detalles presione <Enter>) "
    
End Sub

Private Sub lProveedor_Click()

    If cProveedor.Enabled Then
        cProveedor.SetFocus
        cProveedor.SelStart = 0
        cProveedor.SelLength = Len(cProveedor.Text)
    End If
    
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

Sub AccionNuevo()
    
    'Prendo Se�al que es uno nuevo.
    sNuevo = True
    
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    LimpioFicha
    
    Cons = "Select * from PedidoRepuesto Where PReCodigo = 0"
    Set RsPedido = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    tFPedido.SetFocus
    
End Sub

Sub AccionModificar()
    
    sModificar = True
    
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    LimpioFicha
    Cons = "Select * from PedidoRepuesto Where PReCodigo = " & Mid(lPedido.SelectedItem.Key, 2, Len(lPedido.SelectedItem.Key))
    Set RsPedido = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    CargoCamposDesdeBD

End Sub

Sub AccionGrabar()

    If Not ValidoCampos Then
        MsgBox "Los datos ingresados no son correctos o la ficha est� incompleta.", vbExclamation, "ATENCI�N"
        Exit Sub
    End If
    
    If MsgBox("�Confirma almacenar la informaci�n ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
    
        If sNuevo Then                  'Nuevo----------
            Screen.MousePointer = 11
            On Error GoTo errorBT
            
            'Cargo tabla: Pedido
            RsPedido.AddNew
            CargoCamposBDPedido
            RsPedido.Update

            sNuevo = False
            
        Else                                    'Modificar----
        
            On Error GoTo errorBT
            'Cargo tabla: Pedido
            RsPedido.Edit
            CargoCamposBDPedido
            RsPedido.Update
            
            sModificar = False
        
        End If
        RsPedido.Close
        lPedido.SetFocus
    End If
    LimpioFicha
    CargoPedidos
    If lPedido.ListItems.Count > 0 Then
        Call Botones(True, True, True, False, False, Toolbar1, Me)
    Else
        Call Botones(True, False, False, False, False, Toolbar1, Me)
    End If

    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    MsgBox "No se ha podido inicializar la transacci�n. Reintente la operaci�n." & MapeoError, 48, "ATENCI�N"
    Exit Sub
End Sub

Sub AccionEliminar()

    If lPedido.SelectedItem.Index > 0 Then
        If MsgBox("�Confirma eliminar el pedido seleccionado?", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
            Screen.MousePointer = 11
            On Error GoTo errorBT
    
            'Borro los datos de la tabla: ArticuloFolder
            Cons = "Delete PedidoRepuesto Where PReCodigo = " & Mid(lPedido.SelectedItem.Key, 2, Len(lPedido.SelectedItem.Key))
            cBase.Execute (Cons)
            
            LimpioFicha
            CargoPedidos
            Call Botones(True, False, False, False, False, Toolbar1, Me)
            Screen.MousePointer = 0
        End If
    End If
    Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    MsgBox "No se ha podido inicializar la transacci�n. Reintente la operaci�n." & MapeoError, 48, "ATENCI�N"

End Sub

Sub AccionCancelar()

    LimpioFicha
    If sModificar Then
        Call Botones(True, True, True, False, False, Toolbar1, Me)
    Else
        Call Botones(True, False, False, False, False, Toolbar1, Me)
    End If
    sNuevo = False
    sModificar = False
    RsPedido.Close

End Sub

Private Sub tCarpeta_KeyPress(KeyAscii As Integer)

    If KeyAscii = kEnt Then
        tDetalle.SetFocus
    End If
    
End Sub

Private Sub tCarpeta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Ingrese la carpeta en la que fue abierto el pedido."
    
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)

    If KeyAscii = kEnt And (sNuevo Or sModificar) Then
        AccionGrabar
    End If
    
End Sub

Private Sub tComentario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Comentarios generales del pedido."
    
End Sub

Private Sub tDetalle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Detalle de art�culos del pedido."
    
End Sub

Private Sub tFlete_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = kEnt Then
        cArribo.SetFocus
    End If
    
End Sub

Private Sub tFlete_LostFocus()

    If IsNumeric(tFlete.Text) Then
        tFlete.Text = Format(tFlete.Text, "##,##0.00")
    Else
        tFlete.Text = ""
    End If
    
End Sub

Private Sub tFlete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Ingrese el costo del flete."
    
End Sub

Private Sub tFPedido_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = kDow Then
        Call MuestroCalendario(Me, False, tFPedido.Top + tFPedido.Height, tFPedido.Left, tFPedido)
    End If

End Sub

Private Sub tFPedido_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cProveedor.SetFocus
    End If

End Sub

Private Sub tFPedido_LostFocus()

    If IsDate(tFPedido.Text) Then
        tFPedido.Text = Format(tFPedido.Text, "d-mmm yy")
    Else
        tFPedido.Text = ""
    End If
    
End Sub

Private Sub tFPedido_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Fecha de realizaci�n del pedido."
    
End Sub

Private Sub tImporte_KeyPress(KeyAscii As Integer)

    If KeyAscii = kEnt Then
        cPago.SetFocus
    End If
    
End Sub

Private Sub tImporte_LostFocus()

    If IsNumeric(tImporte.Text) Then
        tImporte.Text = Format(tImporte.Text, "##,##0.00")
    Else
        tImporte.Text = ""
    End If
    
End Sub

Private Sub tImporte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Ingrese el importe total del pedido."
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        
        Case "nuevo"
            AccionNuevo
        
        Case "modificar"
            AccionModificar
        
        Case "eliminar"
            AccionEliminar
        
        Case "grabar"
            AccionGrabar
        
        Case "cancelar"
            AccionCancelar
        
        Case "salir"
            Unload Me
            
    End Select

End Sub

Private Sub CargoDatosDesdeLista()

    LimpioFicha
    
    tFPedido.Text = lPedido.SelectedItem
    tDetalle.Text = lPedido.SelectedItem.SubItems(1)
    cEstado.Text = lPedido.SelectedItem.SubItems(2)
    tImporte.Text = lPedido.SelectedItem.SubItems(3)
    If lPedido.SelectedItem.SubItems(4) = "Si" Then
        cArribo.Value = 1
    Else
        cArribo.Value = 0
    End If
    
    cProveedor.TextColumn = 1
    cProveedor.Text = lPedido.SelectedItem.SubItems(5)
    cProveedor.TextColumn = 2
    
    cAgencia.TextColumn = 1
    cAgencia.Text = lPedido.SelectedItem.SubItems(6)
    cAgencia.TextColumn = 2
    
    tFlete.Text = lPedido.SelectedItem.SubItems(7)
    cPago.Value = lPedido.SelectedItem.SubItems(8)
    tComentario.Text = lPedido.SelectedItem.SubItems(9)
    tCarpeta.Text = lPedido.SelectedItem.SubItems(10)
    
    cTransporte.TextColumn = 1
    cTransporte.Text = lPedido.SelectedItem.SubItems(11)
    cTransporte.TextColumn = 2
    
End Sub

Private Sub CargoCamposBDPedido()
        
        'Cargo datos tabla: Pedido
        RsPedido!PReFecha = Format(tFPedido.Text, "mm/dd/yy")
        RsPedido!PReProveedor = cProveedor.Column(0)
        RsPedido!PReAgencia = cAgencia.Column(0)
        RsPedido!PReEstado = cEstado.Column(0)
        RsPedido!PReTransporte = cTransporte.Column(0)
        If Trim(tCarpeta.Text) <> "" Then
            RsPedido!PReCarpeta = Trim(tCarpeta.Text)
        Else
            RsPedido!PReCarpeta = Null
        End If
        
        If Trim(tFlete.Text) <> "" Then
            RsPedido!PReFlete = CCur(tFlete.Text)
        Else
            RsPedido!PReFlete = Null
        End If
        
        If Trim(tImporte.Text) <> "" Then
            RsPedido!PReImporte = CCur(tImporte.Text)
        Else
            RsPedido!PReImporte = Null
        End If
        If cPago.Value = 0 Then
            RsPedido!PRePago = False
        Else
            RsPedido!PRePago = True
        End If
        If cArribo.Value = 0 Then
            RsPedido!PReArribo = False
        Else
            RsPedido!PReArribo = True
        End If
        
        If Trim(tDetalle.Text) <> "" Then
            RsPedido!PReDetalle = Trim(tDetalle.Text)
        Else
            RsPedido!PReDetalle = Null
        End If
        
        If Trim(tComentario.Text) <> "" Then
            RsPedido!PReComentario = Trim(tComentario.Text)
        Else
            RsPedido!PReComentario = Null
        End If
               
End Sub

Private Sub CargoProveedor()
    
    cProveedor.Clear
    'Cargo los proveedores en el combo
    Cons = "Select * from Empresa Where EmpGiro = " & cGProveedor & " Order by EmpNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    I = 0
    Do While Not RsAux.EOF
        cProveedor.AddItem RsAux!EmpCodigo
        cProveedor.List(I, 1) = Trim(RsAux!EmpNombre)
        RsAux.MoveNext
        I = I + 1
    Loop
    RsAux.Close
    
End Sub

Private Sub CargoAgencia()
    
    cAgencia.Clear
    Cons = "Select * from Empresa Where EmpGiro = " & cGAgencia & " Order by EmpNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    I = 0
    Do While Not RsAux.EOF
        cAgencia.AddItem RsAux!EmpCodigo
        cAgencia.List(I, 1) = Trim(RsAux!EmpNombre)
        RsAux.MoveNext
        I = I + 1
    Loop
    RsAux.Close
    
End Sub

Sub CargoMediosTransporte()

    cTransporte.AddItem CVAereo
    cTransporte.List(0, 1) = cTAereo
    cTransporte.AddItem CVMaritimo
    cTransporte.List(1, 1) = cTMaritimo
    cTransporte.AddItem CvTerrestre
    cTransporte.List(2, 1) = cTTerrestre
    
End Sub

Private Sub CargoEstado()

    cEstado.AddItem 1
    cEstado.List(0, 1) = "Pedir Cotizaci�n"
    cEstado.AddItem 2
    cEstado.List(1, 1) = "Esperando Cotizaci�n"
    cEstado.AddItem 3
    cEstado.List(2, 1) = "Esperando Rebaja"
    cEstado.AddItem 4
    cEstado.List(3, 1) = "Confirmado"
    cEstado.AddItem 5
    cEstado.List(4, 1) = "Pendiente de Embarque"
    cEstado.AddItem 6
    cEstado.List(5, 1) = "En Viaje"
    cEstado.AddItem 7
    cEstado.List(6, 1) = "En Puerto"
    
End Sub

Private Sub CargoCamposDesdeBD()
    
    If Not RsPedido.EOF Then
        LimpioFicha
        'Cargo datos desde tabla: Pedido
        tFPedido.Text = Format(RsPedido!PReFecha, "d-mmm yy")
        
        If Not IsNull(RsPedido!PReImporte) Then
            tImporte.Text = Format(RsPedido!PReImporte, "##,##0.00")
        End If
        If Not IsNull(RsPedido!PReFlete) Then
            tFlete.Text = Format(RsPedido!PReFlete, "##,##0.00")
        End If
        
        If RsPedido!PReArribo Then
            cArribo.Value = 1
        End If
        
        If RsPedido!PRePago Then
            cPago.Value = 1
        End If
    
        cAgencia.TextColumn = 1
        cAgencia.Text = RsPedido!PReAgencia
        cAgencia.TextColumn = 2
        
        cEstado.TextColumn = 1
        cEstado.Text = RsPedido!PReEstado
        cEstado.TextColumn = 2
    
        cTransporte.TextColumn = 1
        cTransporte.Text = RsPedido!PReTransporte
        cTransporte.TextColumn = 2
    
        cProveedor.TextColumn = 1
        cProveedor.Text = RsPedido!PReProveedor
        cProveedor.TextColumn = 2
        
        If Not IsNull(RsPedido!PReDetalle) Then
            tDetalle.Text = Trim(RsPedido!PReDetalle)
        End If
    
        If Not IsNull(RsPedido!PReComentario) Then
            tComentario.Text = Trim(RsPedido!PReComentario)
        End If
        
        If Not IsNull(RsPedido!PReCarpeta) Then
            tCarpeta.Text = RsPedido!PReCarpeta
        End If
        
    Else
        RsAux.Close
        MsgBox "El pedido seleccionado ha sido eliminado desde otra terminal.", vbExclamation, "ATENCI�N"
        CargoPedidos
        Call Botones(True, False, False, False, False, Toolbar1, Me)
    End If
    
End Sub

Private Function ValidoCampos()

    ValidoCampos = True
    
    If Trim(tFPedido.Text) <> "" And Not IsDate(tFPedido.Text) Then
        ValidoCampos = False
    End If
    
    If Trim(tImporte.Text) <> "" Then
        If Not IsNumeric(tImporte.Text) Then
            ValidoCampos = False
        End If
    End If
    
    If Trim(tFlete.Text) <> "" Then
        If Not IsNumeric(tFlete.Text) Then
            ValidoCampos = False
        End If
    End If
    
    If Trim(tCarpeta.Text) <> "" Then
        If Not IsNumeric(tCarpeta.Text) Then
            ValidoCampos = False
        End If
    End If
    
    If Trim(tFPedido.Text) = "" Or cProveedor.ListIndex = -1 Or cAgencia.ListIndex = -1 _
    Or cTransporte.ListIndex = -1 Or cEstado.ListIndex = -1 Then
        ValidoCampos = False
    End If
        
End Function

Private Sub CargoPedidos()

    lPedido.ListItems.Clear
    Cons = "Select * from PedidoRepuesto"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
    
        Set itmx = lPedido.ListItems.Add(, "A" + Str(RsAux!PReCodigo), Format(RsAux!PReFecha, "d-mmm-yy"))
        If Not IsNull(RsAux!PReDetalle) Then
            itmx.SubItems(1) = Trim(RsAux!PReDetalle)
        Else
            itmx.SubItems(1) = ""
        End If
        
        'Estado del pedido
        cEstado.TextColumn = 1
        cEstado.Text = RsAux!PReEstado
        itmx.SubItems(2) = Trim(cEstado.Column(1))
        cEstado.TextColumn = 2
        cEstado.Text = ""
        
        If Not IsNull(RsAux!PReImporte) Then
            itmx.SubItems(3) = Format(RsAux!PReImporte, "##,##0.00")
        Else
            itmx.SubItems(3) = ""
        End If
        If RsAux!PReArribo Then
            itmx.SubItems(4) = "Si"
        Else
            itmx.SubItems(4) = "No"
        End If
        itmx.SubItems(5) = RsAux!PReProveedor
        itmx.SubItems(6) = RsAux!PReAgencia
        If Not IsNull(RsAux!PReFlete) Then
            itmx.SubItems(7) = Format(RsAux!PReFlete, "##,##0.00")
        Else
            itmx.SubItems(7) = ""
        End If
        If RsAux!PRePago Then
            itmx.SubItems(8) = "1"
        Else
            itmx.SubItems(8) = "0"
        End If
        
        If Not IsNull(RsAux!PReComentario) Then
            itmx.SubItems(9) = Trim(RsAux!PReComentario)
        Else
            itmx.SubItems(9) = ""
        End If
        If Not IsNull(RsAux!PReCarpeta) Then
            itmx.SubItems(10) = Trim(RsAux!PReCarpeta)
        Else
            itmx.SubItems(10) = ""
        End If
        itmx.SubItems(11) = RsAux!PReTransporte
        
        RsAux.MoveNext
    Loop
    RsAux.Close

End Sub

Private Sub LimpioFicha()

    tFPedido.Text = ""
    cProveedor.Text = ""
    cAgencia.Text = ""
    cTransporte.Text = ""
    cEstado.Text = ""
    tCarpeta.Text = ""
    
    tDetalle.Text = ""
    tComentario.Text = ""
    tImporte.Text = ""
    tFlete.Text = ""
    tComentario.Text = ""
    
    cPago.Value = 0
    cArribo.Value = 0
                
End Sub

