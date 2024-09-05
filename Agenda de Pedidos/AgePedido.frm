VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.0#0"; "AACOMBO.OCX"
Begin VB.Form AgePedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda de Pedidos "
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7455
   Icon            =   "AgePedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   21
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
   Begin AACombo99.AACombo cEstado 
      Height          =   315
      Left            =   840
      TabIndex        =   9
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      BackColor       =   12648447
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
      TabIndex        =   5
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      BackColor       =   12648447
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
      TabIndex        =   7
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      BackColor       =   12648447
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
      TabIndex        =   3
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      BackColor       =   12648447
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
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.CheckBox chArribo 
      Caption         =   "Arribó"
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
      TabIndex        =   18
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
      TabIndex        =   15
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
      TabIndex        =   17
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
      TabIndex        =   12
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
      TabIndex        =   19
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
      TabIndex        =   14
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
      MaxLength       =   12
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin ComctlLib.ListView lPedido 
      Height          =   1695
      Left            =   120
      TabIndex        =   20
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
         Object.Width           =   1340
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Detalle"
         Object.Width           =   4147
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Estado"
         Object.Width           =   2469
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
         Text            =   "Arribó"
         Object.Width           =   0
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
            Key             =   ""
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
      TabIndex        =   6
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
      TabIndex        =   10
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
      TabIndex        =   16
      Top             =   2520
      Width           =   615
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
      TabIndex        =   8
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
      TabIndex        =   4
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
      TabIndex        =   22
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
      TabIndex        =   13
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
            Picture         =   "AgePedido.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AgePedido.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AgePedido.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AgePedido.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AgePedido.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AgePedido.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AgePedido.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AgePedido.frx":0DC8
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
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuFiltros 
      Caption         =   "Filtros"
      Visible         =   0   'False
      Begin VB.Menu MnuTodos 
         Caption         =   "Sin filtros"
      End
      Begin VB.Menu MnuPendientes 
         Caption         =   "Pedidos Pendientes"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuRealizados 
         Caption         =   "Pedidos Realizados"
      End
   End
End
Attribute VB_Name = "AgePedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sNuevo As Boolean, sModificar As Boolean
Private FPedido As Date

Private Sub cAgencia_GotFocus()
    RelojA
    With cAgencia
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione la agencia asociada al pedido."
    RelojD
End Sub

Private Sub cAgencia_KeyPress(KeyAscii As Integer)
    If vbKeyReturn = KeyAscii Then Foco cTransporte
End Sub

Private Sub cAgencia_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda "Seleccione la agencia asociada al pedido."
End Sub

Private Sub cEstado_GotFocus()
    RelojA
    With cEstado
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione el estado del pedido."
    RelojD
End Sub

Private Sub cEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tCarpeta
End Sub

Private Sub cEstado_LostFocus()
    cEstado.SelStart = 0
    Ayuda ""
End Sub

Private Sub cEstado_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda "Seleccione el estado del pedido."
End Sub

Private Sub cProveedor_GotFocus()
    Ayuda "Seleccione el proveedor de artículos del pedido."
    With cProveedor
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cAgencia
End Sub

Private Sub cProveedor_LostFocus()
    cProveedor.SelStart = 0
    Ayuda ""
End Sub

Private Sub cProveedor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda "Seleccione el proveedor de artículos del pedido."
End Sub

Private Sub cTransporte_GotFocus()
    With cTransporte
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione el medio de transporte."
End Sub

Private Sub cTransporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cEstado
End Sub

Private Sub cTransporte_LostFocus()
    cTransporte.SelStart = 0
    Ayuda ""
End Sub

Private Sub cTransporte_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda "Seleccione el medio de transporte."
End Sub

Private Sub chArribo_GotFocus()
    Ayuda "Indique si el pedido arribó."
End Sub

Private Sub chArribo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tComentario.SetFocus
End Sub

Private Sub chArribo_LostFocus()
    Ayuda ""
End Sub

Private Sub chPago_GotFocus()
    Ayuda "Indique si el pedido está pago."
End Sub

Private Sub chPago_KeyPress(KeyAscii As Integer)
    If vbKeyReturn = KeyAscii Then Foco tFlete
End Sub

Private Sub chPago_LostFocus()
    Ayuda ""
End Sub

Private Sub Form_Activate()
    RelojD
    Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    SetearLView lvValores.FullRow Or lvValores.UnClickIcono, lPedido
    
    sNuevo = False: sModificar = False
    FechaDelServidor
    FPedido = gFechaServidor
    CargoProveedor
    CargoAgencia
    CargoEstado
    CargoMediosTransporte
    CargoPedidos
    
    Exit Sub
ErrLoad:
    msgError.MuestroError "Ocurrio un error al iniciar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set msgError = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco cAgencia
End Sub

Private Sub Label2_Click()
    Foco cEstado
End Sub

Private Sub Label3_Click()
    Foco tFlete
End Sub

Private Sub Label4_Click()
    Foco tFPedido
End Sub

Private Sub Label5_Click()
    Foco tCarpeta
End Sub

Private Sub Label6_Click()
    Foco cTransporte
End Sub

Private Sub lFEmbarque_Click()
    Foco tImporte
End Sub

Private Sub lPedido_ItemClick(ByVal Item As ComctlLib.ListItem)
    
    If lPedido.SelectedItem.Index <> -1 Then
        CargoDatosDesdeLista
        Call Botones(True, True, True, False, False, Toolbar1, Me)
    Else
        Call Botones(True, False, False, False, False, Toolbar1, Me)
    End If
    
End Sub

Private Sub lPedido_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu MnuFiltros, x:=x + lPedido.Left, y:=y + lPedido.Top
End Sub

Private Sub lPedido_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda "Lista de pedidos."
End Sub

Private Sub lProveedor_Click()
    Foco cProveedor
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

Private Sub MnuPendientes_Click()
On Error GoTo ErrMP
    RelojA
    LimpioFicha
    MnuTodos.Checked = False
    MnuPendientes.Checked = True
    MnuRealizados.Checked = False
    CargoPedidos
    RelojD
    Exit Sub
ErrMP:
    msgError.MuestroError "Ocurrio un error al cargar los pedidos pendientes."
    RelojD
End Sub

Private Sub MnuRealizados_Click()
On Error GoTo ErrMR
    RelojA
    LimpioFicha
    MnuTodos.Checked = False
    MnuPendientes.Checked = False
    MnuRealizados.Checked = True
    CargoPedidos
    RelojD
    Exit Sub
ErrMR:
    msgError.MuestroError "Ocurrio un error al cargar los pedidos pendientes."
    RelojD
End Sub

Private Sub MnuTodos_Click()
On Error GoTo ErrMT
    RelojA
    LimpioFicha
    MnuTodos.Checked = True
    MnuPendientes.Checked = False
    MnuRealizados.Checked = False
    CargoPedidos
    RelojD
    Exit Sub
ErrMT:
    msgError.MuestroError "Ocurrio un error al cargar los pedidos pendientes."
    RelojD

End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Sub AccionNuevo()
    
    'Prendo Señal que es uno nuevo.
    sNuevo = True
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    LimpioFicha
    tFPedido.SetFocus
    
End Sub

Sub AccionModificar()
    
    sModificar = True
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    LimpioFicha
    lPedido.Enabled = False
    CargoCamposDesdeBD Mid(lPedido.SelectedItem.Key, 2, Len(lPedido.SelectedItem.Key))

End Sub

Sub AccionGrabar()

    If Not ValidoCampos Then
        MsgBox "Los datos ingresados no son correctos o la ficha está incompleta.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
        RelojA
        If sNuevo Then                  'Nuevo----------
            On Error GoTo errorBT
            
            'Cargo tabla: Pedido
            Cons = "Select * From PedidoRepuesto Where PReCodigo = 0"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            RsAux.AddNew
            CargoCamposBDPedido
            RsAux.Update
            RsAux.Close
            sNuevo = False
        Else                                    'Modificar----
            On Error GoTo errorBT
            Cons = "Select * From PedidoRepuesto Where PReCodigo = " & Mid(lPedido.SelectedItem.Key, 2, Len(lPedido.SelectedItem.Key))
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            'Cargo tabla: Pedido
            If Not RsAux.EOF Then
                If RsAux!PReFModificacion = FPedido Then
                    RsAux.Edit
                    CargoCamposBDPedido
                    RsAux.Update
                Else
                    MsgBox "Otra terminal modificícó el pedido con anterioridad, verifique.", vbExclamation, "ATENCIÓN"
                End If
            Else
                MsgBox "El pedido fue eliminado, verifique.", vbExclamation, "ATENCIÓN"
            End If
            RsAux.Close
            sModificar = False
        End If
        lPedido.Enabled = True
        lPedido.SetFocus
    End If
    LimpioFicha
    CargoPedidos
    If lPedido.ListItems.Count > 0 Then
        Call Botones(True, True, True, False, False, Toolbar1, Me)
    Else
        Call Botones(True, False, False, False, False, Toolbar1, Me)
    End If
    RelojD
    Exit Sub
    
errorBT:
    RelojD
    msgError.MuestroError "No se ha podido inicializar la transacción. Reintente la operación.", Trim(Err.Description)
    Exit Sub
End Sub

Sub AccionEliminar()
    If lPedido.SelectedItem.Index > 0 Then
        If MsgBox("¿Confirma eliminar el pedido seleccionado?", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
            RelojA
            On Error GoTo errorBT
            'Borro los datos de la tabla: ArticuloFolder
            Cons = "Delete PedidoRepuesto Where PReCodigo = " & Mid(lPedido.SelectedItem.Key, 2, Len(lPedido.SelectedItem.Key))
            cBase.Execute (Cons)
            LimpioFicha
            CargoPedidos
            Call Botones(True, False, False, False, False, Toolbar1, Me)
            RelojD
        End If
    End If
    Exit Sub
errorBT:
    RelojD
    msgError.MuestroError "No se ha podido inicializar la transacción. Reintente la operación.", Trim(Err.Description)
End Sub

Sub AccionCancelar()

    RelojA
    LimpioFicha
    If sModificar Then
        Call Botones(True, True, True, False, False, Toolbar1, Me)
    Else
        Call Botones(True, False, False, False, False, Toolbar1, Me)
    End If
    sNuevo = False
    sModificar = False
    lPedido.Enabled = True
    RelojD

End Sub

Private Sub tCarpeta_GotFocus()
    Ayuda "Ingrese la carpeta en la que fue abierto el pedido."
    With tCarpeta
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCarpeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDetalle
End Sub

Private Sub tCarpeta_LostFocus()
    Ayuda ""
    tCarpeta.SelStart = 0
End Sub

Private Sub tCarpeta_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda "Ingrese la carpeta en la que fue abierto el pedido."
End Sub

Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Comentarios generales del pedido."
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And (sNuevo Or sModificar) Then AccionGrabar
End Sub
Private Sub tComentario_LostFocus()
    Ayuda ""
    tComentario.SelStart = 0
End Sub
Private Sub tComentario_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda "Comentarios generales del pedido."
End Sub
Private Sub tDetalle_GotFocus()
    Ayuda "Detalle de artículos del pedido."
    With tDetalle
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tDetalle_LostFocus()
    tDetalle.SelStart = 0
    Ayuda ""
End Sub
Private Sub tDetalle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda "Detalle de artículos del pedido."
End Sub
Private Sub tFlete_GotFocus()
    Ayuda "Ingrese el costo del flete."
    With tFlete
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tFlete_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chArribo.SetFocus
End Sub
Private Sub tFlete_LostFocus()
    If IsNumeric(tFlete.Text) Then
        tFlete.Text = Format(tFlete.Text, FormatoMonedaP)
    Else
        tFlete.Text = ""
    End If
    Ayuda ""
End Sub
Private Sub tFlete_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda "Ingrese el costo del flete."
End Sub
Private Sub tFPedido_GotFocus()
    Ayuda "Fecha de realización del pedido."
    With tFPedido
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tFPedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cProveedor
End Sub
Private Sub tFPedido_LostFocus()

    If IsDate(tFPedido.Text) Then
        tFPedido.Text = Format(tFPedido.Text, "d-mmm yy")
    Else
        tFPedido.Text = ""
    End If
    Ayuda ""
    
End Sub
Private Sub tFPedido_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda "Fecha de realización del pedido."
End Sub
Private Sub tImporte_GotFocus()
    Ayuda "Ingrese el importe total del pedido."
    With tImporte
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chPago.SetFocus
End Sub
Private Sub tImporte_LostFocus()

    If IsNumeric(tImporte.Text) Then
        tImporte.Text = Format(tImporte.Text, FormatoMonedaP)
    Else
        tImporte.Text = ""
    End If
    Ayuda ""
    
End Sub
Private Sub tImporte_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Ayuda "Ingrese el importe total del pedido."
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
On Error GoTo ErrCDDL
    
    RelojA
    LimpioFicha
    
    tFPedido.Text = lPedido.SelectedItem
    tDetalle.Text = lPedido.SelectedItem.SubItems(1)
    cEstado.Text = lPedido.SelectedItem.SubItems(2)
    tImporte.Text = lPedido.SelectedItem.SubItems(3)
    
    If lPedido.SelectedItem.SubItems(4) = "Si" Then
        chArribo.Value = 1
    Else
        chArribo.Value = 0
    End If
    
    BuscoCodigoEnCombo cProveedor, Val(lPedido.SelectedItem.SubItems(5))
    BuscoCodigoEnCombo cAgencia, Val(lPedido.SelectedItem.SubItems(6))
    
    tFlete.Text = lPedido.SelectedItem.SubItems(7)
    chPago.Value = Val(lPedido.SelectedItem.SubItems(8))
    tComentario.Text = lPedido.SelectedItem.SubItems(9)
    tCarpeta.Text = lPedido.SelectedItem.SubItems(10)
    BuscoCodigoEnCombo cTransporte, Val(lPedido.SelectedItem.SubItems(11))
    RelojD
    Exit Sub
ErrCDDL:
    msgError.MuestroError "Ocurrio un error al cargar los datos del pedido.", Trim(Err.Description)
    RelojD
End Sub

Private Sub CargoCamposBDPedido()
On Error GoTo ErrCCBP
    
    'Cargo datos tabla: Pedido
    RsAux!PReFecha = Format(tFPedido.Text, "mm/dd/yy")
    RsAux!PReProveedor = cProveedor.ItemData(cProveedor.ListIndex)
    RsAux!PReAgencia = cAgencia.ItemData(cAgencia.ListIndex)
    RsAux!PreEstado = cEstado.ItemData(cEstado.ListIndex)
    RsAux!PReTransporte = cTransporte.ItemData(cTransporte.ListIndex)
    If Trim(tCarpeta.Text) <> "" Then
        RsAux!PReCarpeta = Trim(tCarpeta.Text)
    Else
        RsAux!PReCarpeta = Null
    End If
    
    If Trim(tFlete.Text) <> "" Then
        RsAux!PReFlete = CCur(tFlete.Text)
    Else
        RsAux!PReFlete = Null
    End If
    
    If Trim(tImporte.Text) <> "" Then
        RsAux!PReImporte = CCur(tImporte.Text)
    Else
        RsAux!PReImporte = Null
    End If
    If chPago.Value = 0 Then
        RsAux!PRePago = False
    Else
        RsAux!PRePago = True
    End If
    If chArribo.Value = 0 Then
        RsAux!PReArribo = False
    Else
        RsAux!PReArribo = True
    End If
    
    If Trim(tDetalle.Text) <> "" Then
        RsAux!PReDetalle = Trim(tDetalle.Text)
    Else
        RsAux!PReDetalle = Null
    End If
    
    If Trim(tComentario.Text) <> "" Then
        RsAux!PReComentario = Trim(tComentario.Text)
    Else
        RsAux!PReComentario = Null
    End If
    RsAux!PReFModificacion = Format(Now, sqlFormatoFH)
    Exit Sub
ErrCCBP:
    msgError.MuestroError "Ocurrio un error al cargar los datos del pedido.", Trim(Err.Description)
End Sub
Private Sub CargoProveedor()
    Cons = "Select PExCodigo, PExNombre From ProveedorExterior" _
        & " Order by PExNombre"
    CargoCombo Cons, cProveedor, ""
End Sub

Private Sub CargoAgencia()
    Cons = "Select ATrCodigo, ATrNombre From AgenciaTransporte Order by ATrNombre"
    CargoCombo Cons, cAgencia, ""
End Sub

Sub CargoMediosTransporte()
    cTransporte.AddItem RetornoMedioTransporte(CVAereo)
    cTransporte.ItemData(cTransporte.NewIndex) = CVAereo
    cTransporte.AddItem RetornoMedioTransporte(CVMaritimo)
    cTransporte.ItemData(cTransporte.NewIndex) = CVMaritimo
    cTransporte.AddItem RetornoMedioTransporte(CvTerrestre)
    cTransporte.ItemData(cTransporte.NewIndex) = CvTerrestre
End Sub

Private Sub CargoEstado()
    cEstado.AddItem RetornoEstado(1)
    cEstado.ItemData(cEstado.NewIndex) = 1
    cEstado.AddItem RetornoEstado(2)
    cEstado.ItemData(cEstado.NewIndex) = 2
    cEstado.AddItem RetornoEstado(3)
    cEstado.ItemData(cEstado.NewIndex) = 3
    cEstado.AddItem RetornoEstado(4)
    cEstado.ItemData(cEstado.NewIndex) = 4
    cEstado.AddItem RetornoEstado(5)
    cEstado.ItemData(cEstado.NewIndex) = 5
    cEstado.AddItem RetornoEstado(6)
    cEstado.ItemData(cEstado.NewIndex) = 6
    cEstado.AddItem RetornoEstado(7)
    cEstado.ItemData(cEstado.NewIndex) = 7
End Sub
Private Function RetornoEstado(CodEstado As Integer) As String
    Select Case CodEstado
        Case 1: RetornoEstado = "Pedir Cotización"
        Case 2: RetornoEstado = "Esperando Cotización"
        Case 3: RetornoEstado = "Esperando Rebaja"
        Case 4: RetornoEstado = "Confirmado"
        Case 5: RetornoEstado = "Pendiente de Embarque"
        Case 6: RetornoEstado = "En Viaje"
        Case 7: RetornoEstado = "En Puerto"
        Case Else: RetornoEstado = ""
    End Select
End Function
Private Sub CargoCamposDesdeBD(CodPedido As Long)
On Error GoTo ErrCCDB

    Cons = "Select * from PedidoRepuesto Where PReCodigo = " & CodPedido
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        RelojA
        LimpioFicha
        
        'Cargo datos desde tabla: Pedido
        tFPedido.Text = Format(RsAux!PReFecha, FormatoFP)
        
        If Not IsNull(RsAux!PReImporte) Then tImporte.Text = Format(RsAux!PReImporte, FormatoMonedaP)
        
        If Not IsNull(RsAux!PReFlete) Then tFlete.Text = Format(RsAux!PReFlete, FormatoMonedaP)
        
        If RsAux!PReArribo Then chArribo.Value = 1
        
        If RsAux!PRePago Then chPago.Value = 1
    
        BuscoCodigoEnCombo cAgencia, RsAux!PReAgencia
        BuscoCodigoEnCombo cEstado, RsAux!PreEstado
        BuscoCodigoEnCombo cTransporte, RsAux!PReTransporte
        BuscoCodigoEnCombo cProveedor, RsAux!PReProveedor
        
        If Not IsNull(RsAux!PReDetalle) Then tDetalle.Text = Trim(RsAux!PReDetalle)
    
        If Not IsNull(RsAux!PReComentario) Then tComentario.Text = Trim(RsAux!PReComentario)
        
        If Not IsNull(RsAux!PReCarpeta) Then tCarpeta.Text = RsAux!PReCarpeta
        
        FPedido = RsAux!PReFModificacion
        
        RelojD
        
    Else
    
        MsgBox "El pedido seleccionado ha sido eliminado desde otra terminal.", vbExclamation, "ATENCIÓN"
        CargoPedidos
        Call Botones(True, False, False, False, False, Toolbar1, Me)
        lPedido.Enabled = True
    End If
    Exit Sub
ErrCCDB:
    msgError.MuestroError "Ocurrio un error al cargar los campos.", Trim(Err.Description)
    RelojD
End Sub

Private Function ValidoCampos()

    ValidoCampos = True
    
    If Trim(tFPedido.Text) <> "" And Not IsDate(tFPedido.Text) Then ValidoCampos = False: Exit Function
    
    If Trim(tImporte.Text) <> "" Then
        If Not IsNumeric(tImporte.Text) Then ValidoCampos = False: Exit Function
    End If
    
    If Trim(tFlete.Text) <> "" Then
        If Not IsNumeric(tFlete.Text) Then ValidoCampos = False: Exit Function
    End If
    
    If Trim(tCarpeta.Text) <> "" Then
        If Not IsNumeric(tCarpeta.Text) Then ValidoCampos = False: Exit Function
    End If
    
    If Trim(tFPedido.Text) = "" Or cProveedor.ListIndex = -1 Or cAgencia.ListIndex = -1 _
    Or cTransporte.ListIndex = -1 Or cEstado.ListIndex = -1 Then ValidoCampos = False: Exit Function
    
End Function

Private Sub CargoPedidos()
On Error GoTo ErrCP

    RelojA
    lPedido.ListItems.Clear
    Cons = "Select * From PedidoRepuesto"

    If MnuPendientes.Checked Then Cons = Cons & " Where PReCarpeta IS Null"
    If MnuRealizados.Checked Then Cons = Cons & " Where PReCarpeta IS Not Null"
    
    Cons = Cons & " Order by PReFecha"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
    
        Set itmx = lPedido.ListItems.Add(, "A" + Str(RsAux!PReCodigo), Format(RsAux!PReFecha, FormatoFP))
        If Not IsNull(RsAux!PReDetalle) Then
            itmx.SubItems(1) = Trim(RsAux!PReDetalle)
        Else
            itmx.SubItems(1) = ""
        End If
        
        'Estado del pedido
        If Not IsNull(RsAux!PreEstado) Then itmx.SubItems(2) = RetornoEstado(RsAux!PreEstado)
        
        If Not IsNull(RsAux!PReImporte) Then
            itmx.SubItems(3) = Format(RsAux!PReImporte, FormatoMonedaP)
        Else
            itmx.SubItems(3) = ""
        End If
        If RsAux!PReArribo Then
            itmx.SubItems(4) = "Si"
        Else
            itmx.SubItems(4) = "No"
        End If
        If Not IsNull(RsAux!PReProveedor) Then itmx.SubItems(5) = RsAux!PReProveedor
        If Not IsNull(RsAux!PReAgencia) Then itmx.SubItems(6) = RsAux!PReAgencia
        If Not IsNull(RsAux!PReFlete) Then
            itmx.SubItems(7) = Format(RsAux!PReFlete, FormatoMonedaP)
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
        If Not IsNull(RsAux!PReTransporte) Then itmx.SubItems(11) = RsAux!PReTransporte
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    RelojD
    Exit Sub
ErrCP:
    msgError.MuestroError "Ocurrio un error al cargar la lista de pedidos.", Trim(Err.Description)
    RelojD
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
    
    chPago.Value = 0
    chArribo.Value = 0
                
End Sub

Private Sub Ayuda(Texto As String)
    Status.Panels(4).Text = Texto
End Sub

