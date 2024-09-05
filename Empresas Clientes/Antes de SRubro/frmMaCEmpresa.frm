VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.0#0"; "AACOMBO.OCX"
Begin VB.Form frmMaCEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresas Proveedores"
   ClientHeight    =   5850
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaCEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   9330
      _ExtentX        =   16457
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
            Style           =   4
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "bBotones"
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   6300
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
   Begin VB.TextBox lAlta 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   220
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   720
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      Caption         =   "Otros Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1365
      Left            =   120
      TabIndex        =   37
      Top             =   4200
      Width           =   9135
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   33
         Top             =   960
         Width           =   7215
      End
      Begin VB.CheckBox cEstatal 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa E&statal:"
         Height          =   255
         Left            =   3960
         TabIndex        =   28
         Top             =   600
         Width           =   1875
      End
      Begin VB.CheckBox cTC 
         Alignment       =   1  'Right Justify
         Caption         =   "&T/C del día anterior:"
         Height          =   255
         Left            =   6360
         TabIndex        =   29
         Top             =   240
         Width           =   1875
      End
      Begin VB.CheckBox cCheque 
         Alignment       =   1  'Right Justify
         Caption         =   "O&pera con Cheques:"
         Height          =   255
         Left            =   3960
         TabIndex        =   27
         Top             =   240
         Width           =   1875
      End
      Begin VB.TextBox tAfiliado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   8040
         TabIndex        =   31
         Top             =   600
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
      Begin AACombo99.AACombo cCategoria 
         Height          =   315
         Left            =   1800
         TabIndex        =   24
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Com&entarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda por &Defecto:"
         Height          =   375
         Left            =   6360
         TabIndex        =   30
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ca&tegoría de Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lAfiliado 
         BackStyle       =   0  'Transparent
         Caption         =   "&Afiliado al Clearing Nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ficha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1755
      Left            =   120
      TabIndex        =   35
      Top             =   480
      Width           =   9135
      Begin VB.TextBox tCargoC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         MaxLength       =   30
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox tNombreC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox tRazonSocial 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   3
         Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox tNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   5
         Top             =   960
         Width           =   4215
      End
      Begin MSMask.MaskEdBox tRuc 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327681
         Appearance      =   0
         ForeColor       =   12582912
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99 999 999 9999"
         PromptChar      =   "_"
      End
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   6840
         TabIndex        =   7
         Top             =   600
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
      Begin AACombo99.AACombo cRamo 
         Height          =   315
         Left            =   6840
         TabIndex        =   9
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "T&ipo Proveedor:"
         Height          =   255
         Left            =   5640
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Carg&o:"
         Height          =   285
         Left            =   4920
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Contacto:"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lModificacion 
         Caption         =   "Mar 12-Dic 1998 15:45"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   6000
         TabIndex        =   36
         Top             =   240
         Width           =   3000
      End
      Begin VB.Label lRuc 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº de R.U.C.:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lRazonSocial 
         BackStyle       =   0  'Transparent
         Caption         =   "&Razón Social:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lRamo 
         BackStyle       =   0  'Transparent
         Caption         =   "&Giro / Ramo:"
         Height          =   255
         Left            =   5640
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lNombre 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom. &Fantasía:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   5595
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8705
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2040
      Left            =   120
      TabIndex        =   38
      Top             =   2160
      Width           =   9135
      Begin AACombo99.AACombo cTipoTelefono 
         Height          =   315
         Left            =   5160
         TabIndex        =   19
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.TextBox tEMail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         MaxLength       =   40
         TabIndex        =   17
         Top             =   1650
         Width           =   4335
      End
      Begin ComctlLib.ListView lvTelefono 
         Height          =   1095
         Left            =   5160
         TabIndex        =   22
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1931
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   1481
         EndProperty
      End
      Begin VB.CommandButton bDireccion 
         Caption         =   "Dirección&..."
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tDireccion 
         Appearance      =   0  'Flat
         Height          =   1335
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox tNroTelefono 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6540
         MaxLength       =   12
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox tInterno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "E-&Mail:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1650
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Teléfonos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5160
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCEmpresa.frx":10E2
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
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefrescar 
         Caption         =   "&Refrescar"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMaCEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gIdEmpresa As Long      'Propiedad para Setear el Cliente Seleccionado
Dim gFModificacion As Date  'Guardo la fecha para contrlar modificaciones

Dim sNuevo As Boolean, sModificar As Boolean
Dim aTexto As String

Dim RsCEm As rdoResultset       'BD Cliente
Dim RsEmp As rdoResultset       'BD CEmpresa

Dim aEmpresa As Long      'Guardo el id de cliente (empresa) p/almacenar en Tabla: CEmpresa

Dim aDireccion As Long      'Guardo el id de direccion p/almacenar en Tabla: Cliente
Dim aCopia As Long

Private Sub bDireccion_Click()

    On Error GoTo errDirecccion
    Screen.MousePointer = 11

    frmDireccion.pCodigoDireccion = aDireccion
    frmDireccion.pCopiaDireccion = aCopia
    frmDireccion.pTextoDireccion = ""
    frmDireccion.Show vbModal, Me
    Me.Refresh
    
    aCopia = frmDireccion.pCopiaDireccion
    
    'Restauro los Valores de Direccion
    If aCopia = -1 Then
        tDireccion.Text = ""
    Else
        If frmDireccion.pTextoDireccion <> "" Then tDireccion.Text = frmDireccion.pTextoDireccion
    End If
        
    Screen.MousePointer = 0
    cTipoTelefono.SetFocus
    Exit Sub

errDirecccion:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al cargar la dirección.", Err.Description
End Sub

Private Sub cCategoria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tAfiliado
End Sub

Private Sub cCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cEstatal.SetFocus
End Sub

Private Sub cEstatal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cTC.SetFocus
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentario
End Sub

Private Sub cRamo_GotFocus()
    cRamo.SelStart = 0: cRamo.SelLength = Len(cRamo.Text)
End Sub

Private Sub cRamo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tNombreC
End Sub

Private Sub cTC_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cMoneda
End Sub

Private Sub cTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cRamo
End Sub

Private Sub cTipoTelefono_GotFocus()
    cTipoTelefono.SelStart = 0: cTipoTelefono.SelLength = Len(cTipoTelefono.Text)
End Sub

Private Sub cTipoTelefono_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(cTipoTelefono.Text) = "" Then cCategoria.SetFocus Else tNroTelefono.SetFocus
    End If
    
End Sub

Private Sub cTipoTelefono_LostFocus()
    cTipoTelefono.SelLength = 0
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error Resume Next
    sNuevo = False: sModificar = False
    SetearLView lvValores.Grilla Or lvValores.FullRow, lvTelefono
    LimpioFicha
    Botones True, False, False, False, False, Toolbar1, Me
    
    
    Cons = "Select TTeCodigo, TTeNombre From TipoTelefono Order by TTeNombre"   'Cargo los TiposTelefono
    CargoCombo Cons, cTipoTelefono, ""
    Cons = "Select RamCodigo, RamNombre From Ramo Order by RamNombre"   'Cargo los RAMOS DE EMPRESAS
    CargoCombo Cons, cRamo, ""
    Cons = "Select CClCodigo, CClNombre From CategoriaCliente Order by CClNombre"    'Cargo las CategoriaCliente
    CargoCombo Cons, cCategoria, ""
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"   'Cargo las monedas
    CargoCombo Cons, cMoneda, ""

    'Cargo los tipos de proveedores de importacion
    cTipo.AddItem RetornoGiroEmpresa(GiroEmpresa.AgenciaDeTransporte)
    cTipo.ItemData(cTipo.NewIndex) = GiroEmpresa.AgenciaDeTransporte
    cTipo.AddItem RetornoGiroEmpresa(GiroEmpresa.Banco)
    cTipo.ItemData(cTipo.NewIndex) = GiroEmpresa.Banco
    cTipo.AddItem RetornoGiroEmpresa(GiroEmpresa.Aseguradoras)
    cTipo.ItemData(cTipo.NewIndex) = GiroEmpresa.Aseguradoras
    cTipo.AddItem RetornoGiroEmpresa(GiroEmpresa.Despachantes)
    cTipo.ItemData(cTipo.NewIndex) = GiroEmpresa.Despachantes
    cTipo.AddItem RetornoGiroEmpresa(GiroEmpresa.ProveedorDeMercaderia)
    cTipo.ItemData(cTipo.NewIndex) = GiroEmpresa.ProveedorDeMercaderia
    
    DeshabilitoIngreso
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If sNuevo Or sModificar Then
        If MsgBox("Ud. realizó modificaciones en la ficha y no ha grabado." & Chr(13) _
            & "Desea almacenar la información ingresada.", vbYesNo + vbExclamation, "ATENCIÓN") = vbYes Then
            
            AccionGrabar
            
            If sNuevo Or sModificar Then
                Cancel = True
                Exit Sub
            End If
        Else
            If aDireccion <> aCopia And aCopia > 0 Then
                'Borro los datos de la direccion copia
                Cons = "Delete Direccion Where DirCodigo = " & aCopia
                cBase.Execute Cons
            End If
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set msgError = Nothing
End Sub

Private Sub Label1_Click()
    Foco cTipoTelefono
End Sub


Private Sub Label2_Click()
    Foco tComentario
End Sub

Private Sub Label3_Click()
    Foco cCategoria
End Sub

Private Sub Label9_Click()
    Foco tEMail
End Sub

Private Sub lAfiliado_Click()
    Foco tAfiliado
End Sub

Private Sub lNombre_Click()
    Foco tNombre
End Sub

Private Sub lRamo_Click()
    Foco cRamo
End Sub

Private Sub lRazonSocial_Click()
    Foco tRazonSocial
End Sub

Private Sub lRuc_Click()
    tRuc.SelStart = 0: tRuc.SelLength = 15: tRuc.SetFocus
End Sub

Private Sub lvTelefono_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not sNuevo And Not sModificar Then Exit Sub
    
    If KeyCode = vbKeyDelete And lvTelefono.ListItems.Count > 0 Then
        lvTelefono.ListItems.Remove lvTelefono.SelectedItem.Index
    End If

End Sub

Private Sub lvTelefono_KeyPress(KeyAscii As Integer)

    If Not sNuevo And Not sModificar Then Exit Sub
    
    If KeyAscii = vbKeyReturn And lvTelefono.ListItems.Count > 0 Then
        BuscoCodigoEnCombo cTipoTelefono, Mid(lvTelefono.SelectedItem.Key, 2, Len(lvTelefono.SelectedItem.Key) - 1)
        tNroTelefono.Text = lvTelefono.SelectedItem.SubItems(1)
        If lvTelefono.SelectedItem.SubItems(2) <> "" Then
            tInterno.Text = lvTelefono.SelectedItem.SubItems(2)
        End If
        lvTelefono.ListItems.Remove lvTelefono.SelectedItem.Index
        cTipoTelefono.SetFocus
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
    tRuc.SetFocus
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub tAfiliado_GotFocus()
    tAfiliado.SelStart = 0: tAfiliado.SelLength = Len(tAfiliado.Text)
End Sub

Private Sub tAfiliado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cCheque.SetFocus
End Sub

Private Sub tCargoC_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If Trim(tDireccion.Text) = "" Then bDireccion.SetFocus Else cTipoTelefono.SetFocus
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tDireccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And (sNuevo Or sModificar) Then cTipoTelefono.SetFocus
End Sub

Private Sub tEMail_GotFocus()
    tEMail.SelStart = 0: tEMail.SelLength = Len(tEMail.Text)
End Sub

Private Sub tEMail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cTipoTelefono
End Sub

Private Sub tInterno_GotFocus()
    Foco tInterno
End Sub

Private Sub tInterno_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cTipoTelefono.ListIndex = -1 Then
            MsgBox "Se debe seleccionar un tipo de teléfono.", vbExclamation, "ATENCIÓN"
            cTipoTelefono.SetFocus
            Exit Sub
        End If
        If Trim(tNroTelefono.Text) = "" Then
            MsgBox "Se debe ingresar un número de teléfono.", vbExclamation, "ATENCIÓN"
            tNroTelefono.SetFocus
            Exit Sub
        End If
        
        Set itmx = lvTelefono.ListItems.Add(, "A" & cTipoTelefono.ItemData(cTipoTelefono.ListIndex), cTipoTelefono.Text)
        itmx.SubItems(1) = Trim(tNroTelefono.Text)
        If Trim(tInterno.Text) <> "" Then
            itmx.SubItems(2) = tInterno.Text
        End If
        cTipoTelefono.ListIndex = -1
        tNroTelefono.Text = ""
        tInterno.Text = ""
        cTipoTelefono.SetFocus
    End If

End Sub

Private Sub tNombre_GotFocus()
    Foco tNombre
End Sub

Private Sub tNombre_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not sNuevo And Not sModificar Then
        If KeyCode = vbKeyF1 And Trim(tNombre.Text) <> "" Then
            If Not TextoValido(tNombre.Text) Then MsgBox "Se han ingresado caracteres no válidos.", vbExclamation, "ATENCIÓN": Exit Sub
            
            Cons = " Select ID_Cliente = CEmCliente, 'Razón Social' = CEmNombre, 'Nombre Fantasía' = CEmFantasia, Ramo = RamNombre " _
                    & " From CEmpresa, Ramo" _
                    & " Where CEmFantasia like '" & Trim(tNombre.Text) & "%'" _
                    & " And CEmRamo *= RamCodigo " _
                    & " Order by CEmFantasia"
            AyudaEmpresa Cons
        End If
    End If

End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        tNombre.Text = NombreEmpresa(tNombre.Text, False)
        If cTipo.Enabled Then Foco cTipo Else Foco tNroTelefono
    End If
    
End Sub

Private Sub tNombreC_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tCargoC
End Sub

Private Sub tNroTelefono_GotFocus()

    If cTipoTelefono.ListIndex > -1 And (sNuevo Or sModificar) Then
        If lvTelefono.ListItems.Count > 0 Then
            For I = 1 To lvTelefono.ListItems.Count
                If Mid(lvTelefono.ListItems(I).Key, 2, Len(lvTelefono.ListItems(I).Key) - 1) = cTipoTelefono.ItemData(cTipoTelefono.ListIndex) Then
                    
                    tNroTelefono.Text = lvTelefono.ListItems(I).SubItems(1)
                    If lvTelefono.ListItems(I).SubItems(2) <> vbNullString Then tInterno.Text = lvTelefono.ListItems(I).SubItems(2)
                    lvTelefono.ListItems.Remove (I)
                    Exit Sub
                End If
            Next I
        End If
    End If
    
    tNroTelefono.SelStart = 0: tNroTelefono.SelLength = Len(tNroTelefono)

End Sub

Private Sub tNroTelefono_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not sNuevo And Not sModificar Then
        If KeyCode = vbKeyF1 And Trim(tNroTelefono.Text) <> "" Then
            
            'Valido el formato del Nro de Telefono------------------------------------------
            If Not TextoValido(tNroTelefono.Text) Then MsgBox "Se han ingresado caracteres no válidos. Verifique.", vbExclamation, "ATENCIÓN": Exit Sub
            
            aTexto = RetornoFormatoTelefono(tNroTelefono.Text, 0)
            If aTexto <> "" Then
                tNroTelefono.Text = aTexto
            Else
                MsgBox "El teléfono ingresado no coincide con los formatos establecidos.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------
            
            Cons = " Select ID_Cliente = CEmCliente, 'Razón Social' = CEmNombre, 'Nombre Fantasía' = CEmFantasia, Ramo = RamNombre " _
                    & " From Telefono, CEmpresa, Ramo" _
                    & " Where TelCliente = CEmCliente " _
                    & " And CEmRamo *= RamCodigo" _
                    & " And TelNumero = '" & Trim(tNroTelefono.Text) & "'" _
                    & " Order by CEmNombre"
            AyudaEmpresa Cons
        End If
    End If
    
End Sub

Private Sub tNroTelefono_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
    
        If sNuevo Or sModificar Then ValidoClienteTelefono Trim(tNroTelefono.Text)
        
        If tInterno.Enabled Then
            If Trim(tNroTelefono.Text) <> "" Then Foco tInterno Else: Foco cCategoria
        Else
            tRuc.SetFocus
        End If
    End If

End Sub

Private Sub tNroTelefono_LostFocus()

    If Trim(tNroTelefono.Text) <> "" Then
        aTexto = RetornoFormatoTelefono(tNroTelefono.Text, aCopia)
        If aTexto <> "" Then
            tNroTelefono.Text = aTexto
        Else
            MsgBox "El teléfono ingresado no coincide con los formatos establecidos.", vbExclamation, "ATENCIÓN"
            Foco tNroTelefono
        End If
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

    Select Case Button.Key
        
        Case "nuevo": AccionNuevo: tRuc.SetFocus
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
            
    End Select

End Sub

Private Sub AccionGrabar()

Dim CodigoDireccion As Long, aEmpresa As Long
Dim aError As String: aError = ""

    If Not ValidoCampos Then Exit Sub
    
    'PREGUNTO PARA GRABAR----------------------------------------------------------------------------------
    If MsgBox("Confirma almacenar los datos ingresados en la ficha.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    FechaDelServidor
    
    If sNuevo Then      'GRABAR EMPRESA NUEVO----------------------------------
        On Error GoTo errorBT
                
        cBase.BeginTrans    'COMIENZO TRANSACCION---------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!
        On Error GoTo errorET
        
        'Copia = Nueva Direccion
        If aCopia > 0 Then CodigoDireccion = aCopia Else CodigoDireccion = 0
        
        Cons = "Select * from Cliente Where CliCodigo = 0"                  'Tabla Cliente
        Set RsCEm = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsCEm.AddNew
        CargoCamposBDCliente CodigoDireccion
        RsCEm.Update: RsCEm.Close
        
        'Saco el id del nuevo cliente-------------------------------------------------------------------
        Cons = "Select Max(CliCodigo) from Cliente"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        aEmpresa = RsAux(0)
        RsAux.Close
        '------------------------------------------------------------------------------------------------
        
        CargoCamposBDCEmpresa aEmpresa
        CargoCamposBDTelefono aEmpresa
        CargoCamposBDEmpresaDato aEmpresa
        
        cBase.CommitTrans   'FINALIZO TRANSACCION---------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!
        
        gIdEmpresa = aEmpresa
        lAlta.Text = "Ingresado: " & Format(gFechaServidor, "Ddd dd/mm/yyyy hh:mm") & " "
        Botones True, True, True, False, False, Toolbar1, Me
                
    Else        'GRABAR EMPRESA MODIFICAR----------------------------------
        
        On Error GoTo errorBT
        cBase.BeginTrans    'COMIENZO TRANSACCION-------------------------------------------------------------------------
        On Error GoTo errorET
        
        Cons = "Select * from Cliente Where CliCodigo = " & gIdEmpresa                  'Tabla Cliente
        Set RsCEm = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        'Controlo Modificacion Multiusuario-----------------
        If gFModificacion <> RsCEm!CliModificacion Then
            aError = "La ficha ha sido modificada por otro usuario. Verifique los datos antes de grabar."
            GoTo errorET: Exit Sub
        End If
        
        If aCopia = -1 Then
            CodigoDireccion = 0                 'Eliminar Direccion
        ElseIf aDireccion = aCopia Then
            CodigoDireccion = aDireccion    'La misma Direccion
        Else
            CodigoDireccion = aCopia         'Nueva Direccion
        End If
        
        RsCEm.Edit
        CargoCamposBDCliente CodigoDireccion
        RsCEm.Update: RsCEm.Close
        '------------------------------------------------------------------------------------------------
        
        CargoCamposBDCEmpresa gIdEmpresa        'Cargo Datos Empresa
        CargoCamposBDTelefono gIdEmpresa
        CargoCamposBDEmpresaDato gIdEmpresa
        
        If aDireccion <> aCopia And aDireccion <> 0 Then        'Borro los datos de la direccion original
            Cons = "Delete Direccion Where DirCodigo = " & aDireccion
            cBase.Execute Cons
        End If
        
        cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
        
        Botones True, True, True, False, False, Toolbar1, Me
    End If
    aCopia = CodigoDireccion: aDireccion = CodigoDireccion
    sNuevo = False: sModificar = False
    gFModificacion = gFechaServidor
    lModificacion.Caption = "Modificado: " & Format(gFechaServidor, "Ddd dd/mm/yyyy hh:mm")
        
    DeshabilitoIngreso
    Screen.MousePointer = 0
    tRuc.SetFocus
    Exit Sub

errorBT:
    Screen.MousePointer = 0
    If aError = "" Then aError = "No se ha podido inicializar la transacción. Reintente la operación."
    msgError.MuestroError aError, Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    If aError = "" Then aError = "No se ha podido inicializar la transacción. Reintente la operación."
    msgError.MuestroError aError, Err.Description
End Sub


Private Sub AccionEliminar()
Dim aResultado As Integer

    'PREGUNTO PARA ELIIMINAR----------------------------------------------------------------------------------
    'Valido eliminación del cliente
    If Not ValidoEliminacion(gIdEmpresa) Then Exit Sub
    
    aResultado = MsgBox("Confirma eliminar la empresa seleccionada." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
            & "Si- Eliminar el cliente." & Chr(vbKeyReturn) _
            & "No- Eliminarlo de las tablas de importaciones.", vbQuestion + vbYesNoCancel + vbDefaultButton3, "GRABAR")
    
    If aResultado = vbCancel Then Exit Sub

    Screen.MousePointer = 11
    On Error GoTo errorBT
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    If aResultado = vbYes Then
        Cons = "Select * from Cliente Where CliCodigo = " & gIdEmpresa
        Set RsCEm = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        'Borro los telefonos del cliente
        Cons = "Delete TELEFONO Where TelCliente = " & gIdEmpresa
        cBase.Execute (Cons)
        
        'Borro los datos de la tabla CEmpresa
        Cons = "Delete CEmpresa Where CEmCliente = " & gIdEmpresa
        cBase.Execute (Cons)
        
        'Borro los datos de la tabla EmpresaDato
        Cons = "Delete EmpresaDato Where EDaTipoEmpresa = " & TipoEmpresa.Cliente & " And EDaCodigo = " & gIdEmpresa
        cBase.Execute (Cons)
        
        'Borro los datos de la tabla cliente
        If Not IsNull(RsCEm!CliDireccion) Then aDireccion = RsCEm!CliDireccion Else aDireccion = 0
            
        RsCEm.Delete
        RsCEm.Close
        
        If aDireccion <> 0 Then      'Tiene Direccion   --> Borro los datos de la direccion
            Cons = "Delete Direccion Where DirCodigo = " & aDireccion
            cBase.Execute Cons
        End If
    End If
    
    If aResultado = vbNo Then
        'Borro los datos de la tabla EmpresaDato
        Cons = "Delete EmpresaDato Where EDaTipoEmpresa = " & TipoEmpresa.Cliente & " And EDaCodigo = " & gIdEmpresa
        cBase.Execute (Cons)
    End If
    
    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
    
    Botones True, False, False, False, False, Toolbar1, Me
    
    LimpioFicha
    Screen.MousePointer = 0
    Exit Sub

errorBT:
    Screen.MousePointer = 0
    msgError.MuestroError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    msgError.MuestroError "No se ha podido realizar la transacción. Reintente la operación."
End Sub

Private Function ValidoEliminacion(idCliente As Long) As Boolean
Dim bHay As Boolean: bHay = False

    Screen.MousePointer = 11
    On Error Resume Next
    ValidoEliminacion = False
    'Tabla Compra
    Cons = "Select * from Compra where ComProveedor = " & idCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then bHay = True: RsAux.Close
    If bHay Then
        MsgBox "Hay compras o gastos que hacen referencia al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    'Tabla RemitoCompra
    Cons = "Select * from RemitoCompra where RCoProveedor = " & idCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then bHay = True: RsAux.Close
    If bHay Then
        MsgBox "Hay compras o gastos que hacen referencia al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    'Tabla Sucursales y cheques
    Cons = "Select * from SucursalDeBanco where SBaBanco = " & idCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then bHay = True: RsAux.Close
    If bHay Then
        MsgBox "Hay sucursales de banco referenciadas al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    Cons = "Select * from ChequeDiferido where CDiBanco = " & idCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then bHay = True: RsAux.Close
    If bHay Then
        MsgBox "Hay cheques diferidos referenciados al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    'Informacion de carpetas
    Cons = "Select * from Carpeta Where CarBcoEmisor = " & idCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then bHay = True: RsAux.Close
    If bHay Then
        MsgBox "Hay carpetas con bancos emisores que hacen referencia al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    Cons = "Select * from Embarque Where EmbAgencia = " & idCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then bHay = True: RsAux.Close
    If bHay Then
        MsgBox "Hay embarques con agencias que hacen referencia al cliente que ud. desea eliminar.", vbExclamation, "ATENCIÓN"
        GoTo Salir
    End If
    
    ValidoEliminacion = True
Salir:
    Screen.MousePointer = 0
End Function

Public Sub AccionNuevo()

    LimpioFicha
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    sNuevo = True
    
    aDireccion = 0: aCopia = 0
    gIdEmpresa = 0
    
    BuscoCodigoEnCombo cCategoria, paCategoriaCliente
    BuscoCodigoEnCombo cTipoTelefono, paTipoTelefonoE

End Sub

Private Sub AccionModificar()

On Error GoTo errModificar

    Screen.MousePointer = 11
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    
    sModificar = True
    
    tRuc.SetFocus
    Screen.MousePointer = 0
    Exit Sub

errModificar:
    msgError.MuestroError "Ha ocurrido un error al cargar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionCancelar()

    On Error Resume Next
    Screen.MousePointer = 11
    DeshabilitoIngreso

    If aDireccion <> aCopia And aCopia > 0 Then 'Borro los datos de la direccion copia
        Cons = "Delete Direccion Where DirCodigo = " & aCopia
        cBase.Execute Cons
    End If
    
    If sNuevo Then
        LimpioFicha
        Botones True, False, False, False, False, Toolbar1, Me
    Else
        CargoDatosEmpresa gIdEmpresa
    End If
    
    sNuevo = False: sModificar = False
    tRuc.SetFocus
    Screen.MousePointer = 0
    
End Sub

Private Sub CargoCamposBDCliente(CodigoDireccion As Long)

    RsCEm!CliTipo = TipoCliente.Empresa
    
    If Trim(tRuc.Text) <> "" Then RsCEm!CliCIRuc = Trim(tRuc.Text) Else: RsCEm!CliCIRuc = Null
    If CodigoDireccion <> 0 Then RsCEm!CliDireccion = CodigoDireccion Else: RsCEm!CliDireccion = Null
    
    If sNuevo Then RsCEm!CliAlta = Format(gFechaServidor, sqlFormatoFH)
    RsCEm!CliModificacion = Format(gFechaServidor, sqlFormatoFH)
    
    If cCheque.Value = 0 Then RsCEm!CliCheque = "N" Else: RsCEm!CliCheque = "S"
    If cCategoria.ListIndex <> -1 Then RsCEm!CliCategoria = cCategoria.ItemData(cCategoria.ListIndex) Else: RsCEm!CliCategoria = Null
    If Trim(tEMail.Text) <> "" Then RsCEm!CliEMail = Trim(tEMail.Text) Else: RsCEm!CliEMail = Null

End Sub

Private Sub CargoCamposBDCEmpresa(idEmpresa As Long)

    Cons = "Select * from CEmpresa Where CEmCliente = " & idEmpresa
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then RsAux.AddNew Else RsAux.Edit
    
    RsAux!CEmCliente = idEmpresa
    RsAux!CEmFantasia = Trim(tNombre.Text)
    If cRamo.ListIndex <> -1 Then RsAux!CEmRamo = cRamo.ItemData(cRamo.ListIndex) Else RsAux!CEmRamo = Null
    If cEstatal.Value = 0 Then RsAux!CEmEstatal = False Else RsAux!CEmEstatal = True
    If Trim(tAfiliado.Text) <> "" Then RsAux!CEmAfiliado = Trim(tAfiliado.Text) Else RsAux!CEmAfiliado = Null
    If Trim(tRazonSocial.Text) <> "" Then RsAux!CEmNombre = Trim(tRazonSocial.Text) Else RsAux!CEmNombre = Null

    RsAux.Update
    RsAux.Close
    
End Sub

Private Sub CargoCamposBDEmpresaDato(idEmpresa As Long)

    Cons = "Select * from EmpresaDato" _
           & " Where EDaTipoEmpresa = " & TipoEmpresa.Cliente _
           & " And EDaCodigo = " & idEmpresa
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then RsAux.AddNew Else RsAux.Edit
    
    RsAux!EDaTipoEmpresa = TipoEmpresa.Cliente
    RsAux!EDaCodigo = idEmpresa
    RsAux!EDaRubro = cTipo.ItemData(cTipo.ListIndex)
    
    If Trim(tNombreC.Text) <> "" Then RsAux!EDaContacto = Trim(tNombreC.Text) Else RsAux!EDaContacto = Null
    If Trim(tCargoC.Text) <> "" Then RsAux!EDaCargoContacto = Trim(tCargoC.Text) Else RsAux!EDaCargoContacto = Null
    
    RsAux!EDaMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    If cTC.Value = vbChecked Then RsAux!EDaTCAnterior = True Else RsAux!EDaTCAnterior = False
    If Trim(tComentario.Text) <> "" Then RsAux!EDaComentario = Trim(tComentario.Text) Else RsAux!EDaComentario = Null
    
    RsAux.Update
    RsAux.Close
    
End Sub

Private Function ValidoCampos()

    ValidoCampos = False
    
    If (cRamo.ListIndex = -1 And Trim(cRamo.Text) <> "") Or (cCategoria.ListIndex = -1 And Trim(cCategoria.Text) <> "") Then
        MsgBox "Los datos ingresados no son correctos o la ficha está incompleta. Verifique", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If Trim(tNombre.Text) = "" Then
        MsgBox "El campo nombre de empresa es obligatorio.", vbExclamation, "ATENCIÓN"
        Foco tNombre: Exit Function
    End If
    
    If Trim(tRazonSocial.Text) = "" Then
        MsgBox "El campo Razón Social es obligatorio.", vbExclamation, "ATENCIÓN"
        Foco tRazonSocial: Exit Function
    End If
    
    If Not TextoValido(tRazonSocial.Text) Then MsgBox "Se han ingresado caracteres no válidos en el campo Razón Social.", vbExclamation, "ATENCIÓN":: Exit Function
    If Not TextoValido(tNombre.Text) Then MsgBox "Se han ingresado caracteres no válidos en el campo Nombre Fantasía.", vbExclamation, "ATENCIÓN": Exit Function
    
    If cTipo.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de proveedor de importaciones (campo obligatorio).", vbExclamation, "ATENCIÓN"
        Foco cTipo: Exit Function
    End If
    
    If Trim(tAfiliado.Text) <> "" Then
        If Not IsNumeric(tAfiliado.Text) Then
            MsgBox "El campo número de afiliado no es numérico.", vbExclamation, "ATENCIÓN"
            Foco tAfiliado: Exit Function
        End If
    End If
    
    If cCategoria.ListIndex = -1 Then
        MsgBox "Debe seleccionar la categoría de cliente (campo obligatorio).", vbExclamation, "ATENCIÓN"
        Foco cCategoria: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda por defecto para los comprobantes de la empresa.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    
    'Verifico si Existe una empresa para el mismo RUC-----------------------------------------------------------
    If Trim(tRuc.Text) <> "" Then
        Cons = "Select * from Cliente Where CliCiRuc = '" & Trim(tRuc.Text) & "' And CliCodigo <> " & gIdEmpresa
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            MsgBox "Existe una empresa registrada para el número de R.U.C.: " & Trim(tRuc.FormattedText), vbExclamation, "ATENCIÓN"
            RsAux.Close: Exit Function
        Else
            RsAux.Close
        End If
    End If
    
    'Verifico si Existe una empresa para el mismo Nombre-------------------------------------------------------
    If sNuevo Then
        Cons = "Select * from CEmpresa " _
                & " Where CEmFantasia = '" & Trim(tNombre.Text) & "'" _
                & " OR CEmNombre = '" & Trim(tRazonSocial.Text) & "'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aTexto = "Existe una empresa registrada para el nombre ingresado." & Chr(vbKeyReturn) _
                        & "Nombre: " & Trim(RsAux!CEmFantasia) & Chr(vbKeyReturn) _
                        & "Razón Social: "
            If Not IsNull(RsAux!CEmNombre) Then aTexto = aTexto & Trim(RsAux!CEmNombre) Else aTexto = aTexto & "S/D"
            aTexto = aTexto & Chr(vbKeyReturn) & Chr(vbKeyReturn)
            aTexto = aTexto & "Continua con el alta de datos."
            If MsgBox(aTexto, vbInformation + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then RsAux.Close: Exit Function
        End If
        RsAux.Close
    End If
        
    ValidoCampos = True
    
End Function

Private Sub DeshabilitoIngreso()

    tNombre.BackColor = Blanco
    tRazonSocial.BackColor = Blanco
    cRamo.Enabled = False: cRamo.BackColor = Inactivo
    cTipo.Enabled = False: cTipo.BackColor = Inactivo
    tNombreC.Enabled = False: tNombreC.BackColor = Inactivo
    tCargoC.Enabled = False: tCargoC.BackColor = Inactivo

    bDireccion.Enabled = False
    tDireccion.BackColor = Inactivo
    tEMail.Enabled = False: tEMail.BackColor = Inactivo
    
    lvTelefono.BackColor = Inactivo
    cTipoTelefono.Enabled = False: cTipoTelefono.BackColor = Inactivo
    tInterno.Enabled = False: tInterno.BackColor = Inactivo
    tNroTelefono.BackColor = Blanco
    
    cCategoria.Enabled = False: cCategoria.BackColor = Inactivo
    tAfiliado.Enabled = False: tAfiliado.BackColor = Inactivo
    cCheque.Enabled = False
    cEstatal.Enabled = False
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tComentario.Enabled = False: tComentario.BackColor = Inactivo
    cTC.Enabled = False

End Sub

Private Sub HabilitoIngreso()

    tNombre.BackColor = Obligatorio
    tRazonSocial.BackColor = Obligatorio
    cRamo.Enabled = True: cRamo.BackColor = Blanco
    cTipo.Enabled = True: cTipo.BackColor = Obligatorio
    tNombreC.Enabled = True: tNombreC.BackColor = Blanco
    tCargoC.Enabled = True: tCargoC.BackColor = Blanco
    
    bDireccion.Enabled = True
    lvTelefono.BackColor = Blanco
    cTipoTelefono.Enabled = True: cTipoTelefono.BackColor = Obligatorio
    tInterno.Enabled = True: tInterno.BackColor = Blanco
    tEMail.Enabled = True: tEMail.BackColor = Blanco
    
    cCategoria.Enabled = True: cCategoria.BackColor = Obligatorio
    tAfiliado.Enabled = True: tAfiliado.BackColor = Blanco
    cCheque.Enabled = True
    cEstatal.Enabled = True
    cMoneda.Enabled = True: cMoneda.BackColor = Obligatorio
    tComentario.Enabled = True: tComentario.BackColor = Blanco
    cTC.Enabled = True
    
End Sub

Private Sub LimpioFicha()

    tRuc.Text = ""
    tRazonSocial.Text = ""
    tNombre.Text = ""
    cRamo.Text = ""
    cTipo.Text = ""
    tNombreC.Text = ""
    tCargoC.Text = ""
    
    tDireccion.Text = ""
    tEMail.Text = ""
    
    cTipoTelefono.Text = ""
    tNroTelefono.Text = ""
    tInterno.Text = ""
    lvTelefono.ListItems.Clear
        
    cCheque.Value = vbUnchecked
    cEstatal.Value = vbUnchecked
    cTC.Value = vbUnchecked
    cMoneda.Text = ""
    cCategoria.Text = ""
    tAfiliado.Text = ""
    tComentario.Text = ""
    lModificacion.Caption = "Modificado: N/D"
    lAlta.Text = "Ingresado: N/D"
    
End Sub

Private Sub tRazonSocial_GotFocus()
    tRazonSocial.SelStart = 0: tRazonSocial.SelLength = Len(tRazonSocial.Text)
End Sub

Private Sub tRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not sNuevo And Not sModificar Then
        If KeyCode = vbKeyF1 And Trim(tRazonSocial.Text) <> "" Then
            If Not TextoValido(tRazonSocial.Text) Then MsgBox "Se han ingresado caracteres no válidos.", vbExclamation, "ATENCIÓN": Exit Sub
            
            Cons = " Select ID_Cliente = CEmCliente, 'Razón Social' = CEmNombre, 'Nombre Fantasía' = CEmFantasia, Ramo = RamNombre " _
                    & " From CEmpresa, Ramo" _
                    & " Where CEmNombre like '" & Trim(tRazonSocial.Text) & "%'" _
                    & " And CEmRamo *= RamCodigo" _
                    & " Order by CEmNombre"
            AyudaEmpresa Cons
        End If
    End If
    
End Sub

Private Sub tRazonSocial_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        On Error GoTo errX
        If Trim(tRazonSocial.Text) <> "" Then
            tRazonSocial.Text = NombreEmpresa(tRazonSocial.Text, True)
            tRazonSocial.Tag = NombreEmpresa(tRazonSocial.Text, False)
            
            If sNuevo Then
                'Busco x el nombre ingresado-------------------------------------------
                Cons = "Select CEmCliente, CEmNombre, CEmFantasia from CEmpresa" _
                        & " Where CEmNombre like '" & Trim(tRazonSocial.Tag) & "%'" _
                        & " Or CEmFantasia like '" & Trim(tRazonSocial.Tag) & "%'" _
                        & " Order by CEmNombre"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
                If Not RsAux.EOF Then
                    RsAux.Close
                    Screen.MousePointer = 0
                    If MsgBox("Existen empresas para el nombre ingresado. Desea visualizar la lista.", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                        Cons = "Select CEmCliente, 'Razón Social' = CEmNombre, 'Nombre Fantasía' = CEmFantasia, Ramo = RamNombre" _
                                & " From CEmpresa, Ramo" _
                                & " Where (CEmNombre like '" & Trim(tRazonSocial.Tag) & "%'" _
                                & " Or CEmFantasia like '" & Trim(tRazonSocial.Tag) & "%')" _
                                & " And CEmRamo *= RamCodigo" _
                                & " Order by CEmNombre"
                        If ListaDeAyuda(Cons) Then Exit Sub
                    End If
                Else
                    RsAux.Close
                End If
                
                tNombre.Text = Trim(tRazonSocial.Tag)
            End If
        End If
        tNombre.SetFocus
    End If
    Exit Sub
    
errX:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al procesar la información."
End Sub

Private Sub tRuc_GotFocus()
    tRuc.SelStart = 0: tRuc.SelLength = 15
End Sub

Private Sub tRuc_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Not sNuevo And Not sModificar Then
            'Busco la Empresa------------------------
            If Trim(tRuc.Text) <> "" Then BuscoEmpresaRuc Trim(tRuc.Text) Else tRazonSocial.SetFocus
        Else
            'Valido que no exista------------------------
            If Trim(tRuc.Text) <> "" Then
                If sNuevo Then
                    ValidoEmpresaRuc tRuc.Text
                Else
                    If Trim(tRuc.Text) <> tRuc.Tag Then ValidoEmpresaRuc tRuc.Text Else tRazonSocial.SetFocus
                End If
            Else
                tRazonSocial.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub BuscoEmpresaRuc(Codigo As String)

On Error GoTo errBuscar

    Screen.MousePointer = 11
    If Codigo <> "" Then
        Cons = "Select * from Cliente" _
                & " Where CliCiRuc = '" & Codigo & "'" _
                & " And CliTipo = " & TipoCliente.Empresa
        Set RsCEm = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        LimpioFicha
        If Not RsCEm.EOF Then
            CargoCamposDesdeBDCliente
            If Not IsNull(RsCEm!CliDireccion) Then CargoCamposDesdeBDDireccion RsCEm!CliDireccion
            CargoCamposDesdeBDCEmpresa RsCEm!CliCodigo
            CargoCamposDesdeBDTelefono RsCEm!CliCodigo
            CargoCamposDesdeBDEmpresaDato RsCEm!CliCodigo
            Botones True, True, True, False, False, Toolbar1, Me
            
        Else
            Screen.MousePointer = 0
            MsgBox "No existe una empresa para el número de R.U.C. ingresado.", vbExclamation, "ATENCIÓN"
            Botones True, False, False, False, False, Toolbar1, Me
            gIdEmpresa = 0
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al cargar los datos.", Err.Description
    
    Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub CargoCamposDesdeBDCliente()

    gIdEmpresa = RsCEm!CliCodigo
    gFModificacion = RsCEm!CliModificacion
    
    If Not IsNull(RsCEm!CliCIRuc) Then tRuc.Text = Trim(RsCEm!CliCIRuc): tRuc.Tag = Trim(RsCEm!CliCIRuc)
    
    If Not IsNull(RsCEm!CliCheque) Then If RsCEm!CliCheque = "S" Then cCheque.Value = vbChecked
    If Not IsNull(RsCEm!CliCategoria) Then BuscoCodigoEnCombo cCategoria, RsCEm!CliCategoria
    
    lAlta.Text = "Ingresado: " & Format(RsCEm!CliAlta, "Ddd dd/mm/yyyy hh:mm")
    lModificacion.Caption = "Modificado: " & Format(RsCEm!CliModificacion, "Ddd dd/mm/yyyy hh:mm")

    If Not IsNull(RsCEm!CliEMail) Then tEMail = Trim(RsCEm!CliEMail)
    
End Sub

Private Sub CargoCamposDesdeBDCEmpresa(idEmpresa As Long)

    Cons = "Select * From CEmpresa Where CEmCliente = " & idEmpresa
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    tNombre.Text = Trim(RsAux!CEmFantasia)
    If Not IsNull(RsAux!CEmRamo) Then BuscoCodigoEnCombo cRamo, RsAux!CEmRamo
    If Not IsNull(RsAux!CEmNombre) Then tRazonSocial.Text = Trim(RsAux!CEmNombre)
    
    If RsAux!CEmEstatal Then cEstatal.Value = vbChecked
    If Not IsNull(RsAux!CEmAfiliado) Then tAfiliado.Text = Trim(RsAux!CEmAfiliado)
    
    RsAux.Close

End Sub

Private Sub CargoCamposDesdeBDEmpresaDato(idEmpresa As Long)

    Cons = " Select * From EmpresaDato " _
            & " Where EDaCodigo = " & idEmpresa _
            & " And EDaTipoEmpresa = " & TipoEmpresa.Cliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!EDaRubro) Then BuscoCodigoEnCombo cTipo, RsAux!EDaRubro
        If Not IsNull(RsAux!EDaContacto) Then tNombreC.Text = Trim(RsAux!EDaContacto)
        If Not IsNull(RsAux!EDaCargoContacto) Then tCargoC.Text = Trim(RsAux!EDaCargoContacto)
        
        If RsAux!EDaTCAnterior Then cTC.Value = vbChecked
        If Not IsNull(RsAux!EDaMoneda) Then BuscoCodigoEnCombo cMoneda, RsAux!EDaMoneda
        If Not IsNull(RsAux!EDaComentario) Then tComentario.Text = Trim(RsAux!EDaComentario)
    End If
    
    RsAux.Close

End Sub

Private Sub CargoCamposDesdeBDDireccion(idDireccion As Long)

    If idDireccion <> 0 Then tDireccion.Text = DireccionATexto(idDireccion) Else tDireccion.Text = ""
    aDireccion = idDireccion
    aCopia = aDireccion
    
End Sub

Private Sub AyudaEmpresa(Consulta As String)
    
    On Error GoTo errBuscar
    Screen.MousePointer = 11
    Dim aLista As New clsListadeAyuda
    aLista.ActivoListaAyudaSQL Consulta, miConexion.TextoConexion(logImportaciones)
    Me.Refresh
    If IsNumeric(aLista.ItemSeleccionadoSQL) Then gIdEmpresa = CCur(aLista.ItemSeleccionadoSQL) Else gIdEmpresa = 0
    Set aLista = Nothing
    
    If gIdEmpresa > 0 Then
        CargoDatosEmpresa gIdEmpresa
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al cargar los datos de la empresa.", Err.Description
    If Not sNuevo And Not sModificar Then Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub CargoDatosEmpresa(Codigo As Long)

    On Error GoTo errCargar
    Screen.MousePointer = 11
    Cons = "Select * from Cliente Where CliCodigo = " & Codigo
    Set RsCEm = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    LimpioFicha
    If Not RsCEm.EOF Then
        CargoCamposDesdeBDCliente
        If Not IsNull(RsCEm!CliDireccion) Then CargoCamposDesdeBDDireccion RsCEm!CliDireccion Else CargoCamposDesdeBDDireccion 0
        CargoCamposDesdeBDCEmpresa Codigo
        CargoCamposDesdeBDTelefono Codigo
        CargoCamposDesdeBDEmpresaDato Codigo
        Botones True, True, True, False, False, Toolbar1, Me
    Else
        gIdEmpresa = 0
        Screen.MousePointer = 0
        MsgBox "La empresa seleccionada ha sido eliminada. Verifique", vbExclamation, "ATENCIÓN"
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    RsCEm.Close
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al cargar los datos de la empresa.", Err.Description
End Sub

Private Sub CargoCamposBDTelefono(Cliente As Long)

    Cons = "Delete TELEFONO Where TelCliente = " & Cliente
    cBase.Execute (Cons)
    If lvTelefono.ListItems.Count > 0 Then
        For I = 1 To lvTelefono.ListItems.Count
            Cons = "Insert Into TELEFONO (TelCliente, TelTipo, TelNumero, TelInterno)" _
                    & " Values (" & Cliente & ", " & Mid(lvTelefono.ListItems(I).Key, 2, Len(lvTelefono.ListItems(I).Key) - 1) _
                    & " , '" & lvTelefono.ListItems(I).SubItems(1) & "', "
                    
            If lvTelefono.ListItems(I).SubItems(2) <> "" Then
                Cons = Cons & "'" & lvTelefono.ListItems(I).SubItems(2) & "'"
            Else
                Cons = Cons & " Null"
            End If
            Cons = Cons & ")"
            cBase.Execute (Cons)
        Next I
    End If
    
End Sub

Private Sub CargoCamposDesdeBDTelefono(idCliente As Long)

On Error GoTo ErrNT

    lvTelefono.ListItems.Clear
    Cons = "Select TelNumero, TelInterno, TTeCodigo, TTeNombre From Telefono, TipoTelefono " _
            & " Where TelCliente = " & idCliente _
            & " And TelTipo = TTeCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        Set itmx = lvTelefono.ListItems.Add(, "A" & RsAux!TTeCodigo, Trim(RsAux!TTeNombre))
        itmx.SubItems(1) = Trim(RsAux!TelNumero)
        If Not IsNull(RsAux!TelInterno) Then itmx.SubItems(2) = Trim(RsAux!TelInterno)
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Sub
        
ErrNT:
    msgError.MuestroError "Ocurrió un error al cargar los números de teléfonos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub ValidoEmpresaRuc(Codigo As String)

On Error GoTo errBuscar

    Dim aCodCli As Long: aCodCli = 0
    
    If Codigo = "" Then Exit Sub
    Screen.MousePointer = 11

    Cons = "Select * from Cliente " _
            & " Where CliCiRuc = '" & Codigo & "'" _
            & " And CliTipo = " & TipoCliente.Empresa
    If sModificar Then Cons = Cons & " And CliCodigo <> " & gIdEmpresa
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then aCodCli = RsAux!CliCodigo
    RsAux.Close
    
    If aCodCli <> 0 Then
        sNuevo = False: sModificar = False
        DeshabilitoIngreso
        CargoDatosEmpresa aCodCli
    Else
        Foco tRazonSocial
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al cargar los datos."
    Botones True, False, False, False, False, Toolbar1, Me
End Sub


Private Sub ValidoClienteTelefono(Telefono As String)

Dim aCodCli As Long
    
    If Telefono = "" Then Exit Sub
    On Error GoTo errBuscar
    Screen.MousePointer = 11

    Cons = "Select * from Telefono, Cliente " _
            & " Where TelNumero = '" & Telefono & "'" _
            & " And TelCliente = CliCodigo And CliTipo = " & TipoCliente.Empresa
    If sModificar Then Cons = Cons & " And TelCliente <> " & gIdEmpresa
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    If Not RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        If MsgBox("Existen clientes registrados para el teléfono ingresado. Desea visualizar la lista de clientes.", vbQuestion + vbOKCancel, "ATENCIÓN") = vbOK Then
            
            Cons = "Select CEmCliente, CEmNombre, CEmFantasia from Telefono, CEmpresa" _
                    & " Where TelCliente = CEmCliente " _
                    & " And TelNumero = '" & Trim(tNroTelefono.Text) & "'"
            ListaDeAyuda Cons
            
        End If
    Else
        RsAux.Close
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al cargar los datos."
    Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Function ListaDeAyuda(Consulta As String) As Boolean

Dim iSeleccionado As Long
    
    On Error GoTo errLista
    Screen.MousePointer = 11
    ListaDeAyuda = False
    
    Dim aLista As New clsListadeAyuda
    aLista.ActivoListaAyuda Consulta, False, miConexion.TextoConexion(logImportaciones), 6500
    
    iSeleccionado = aLista.ValorSeleccionado
    Me.Refresh
    Set aLista = Nothing
        
    If iSeleccionado > 0 Then
        ListaDeAyuda = True
        sNuevo = False: sModificar = False
        DeshabilitoIngreso
        CargoDatosEmpresa iSeleccionado
    End If
    Screen.MousePointer = 0
    Exit Function

errLista:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al activar la lista de ayuda.", Err.Description
End Function
