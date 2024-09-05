VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D9D9E0F6-C86B-4B3A-BFD9-06B9B5B7A222}#2.1#0"; "orUserDigit.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Artículos en Garantía"
   ClientHeight    =   6525
   ClientLeft      =   2580
   ClientTop       =   1995
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7335
   Begin VB.CommandButton bDevolver 
      Caption         =   "Devolver"
      Height          =   315
      Left            =   2880
      TabIndex        =   59
      Top             =   5520
      Width           =   915
   End
   Begin VB.TextBox tSIDServicio 
      Height          =   300
      Left            =   5700
      MaxLength       =   7
      TabIndex        =   22
      Top             =   2700
      Width           =   1035
   End
   Begin VB.TextBox tID 
      Height          =   300
      Left            =   900
      MaxLength       =   7
      TabIndex        =   1
      Top             =   480
      Width           =   1035
   End
   Begin VB.CommandButton bEntregar 
      Caption         =   "&Entregar"
      Height          =   315
      Left            =   2880
      TabIndex        =   40
      Top             =   5880
      Width           =   915
   End
   Begin orUserDigit.UserDigit tUsuario 
      Height          =   285
      Left            =   4980
      TabIndex        =   44
      Top             =   5880
      Width           =   1935
      _ExtentX        =   2805
      _ExtentY        =   503
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
   Begin VB.TextBox tFEntrega 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   900
      MaxLength       =   16
      TabIndex        =   37
      Text            =   "10/10/2002"
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox tCArticulo 
      Height          =   300
      Left            =   900
      TabIndex        =   32
      Top             =   4500
      Width           =   3735
   End
   Begin VB.TextBox tCNumero 
      Height          =   300
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   35
      Top             =   4860
      Width           =   1035
   End
   Begin VB.TextBox tCSerie 
      Height          =   300
      Left            =   900
      MaxLength       =   2
      TabIndex        =   34
      Top             =   4860
      Width           =   375
   End
   Begin VB.TextBox tDCliente 
      Height          =   300
      Left            =   4440
      MaxLength       =   40
      TabIndex        =   18
      Text            =   "AAAAAAAAAAAAAAAAAAAAAABBBBBBBBBBBBBBBBBC"
      Top             =   2220
      Width           =   2715
   End
   Begin VB.TextBox tDCompra 
      Height          =   300
      Left            =   900
      MaxLength       =   10
      TabIndex        =   14
      Text            =   "10/10/2002"
      Top             =   2220
      Width           =   975
   End
   Begin VB.TextBox tDBoleta 
      Height          =   300
      Left            =   2820
      MaxLength       =   10
      TabIndex        =   16
      Text            =   "AD 000000"
      Top             =   2220
      Width           =   855
   End
   Begin VB.CheckBox cDistribuidor 
      BackColor       =   &H00808080&
      Caption         =   "Factura de &Distribuidor"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5160
      MaskColor       =   &H00808080&
      TabIndex        =   12
      Top             =   1905
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.TextBox tSerie 
      Height          =   300
      Left            =   900
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "S"
      Top             =   1860
      Width           =   375
   End
   Begin VB.TextBox tNumero 
      Height          =   300
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   11
      Top             =   1860
      Width           =   1035
   End
   Begin VB.TextBox tSSerie 
      Height          =   300
      Left            =   900
      MaxLength       =   14
      TabIndex        =   28
      Top             =   3780
      Width           =   1695
   End
   Begin VB.TextBox tSFServicio 
      Height          =   300
      Left            =   4080
      TabIndex        =   30
      Top             =   3780
      Width           =   1455
   End
   Begin VB.TextBox tSAbonado 
      Height          =   300
      Left            =   900
      MaxLength       =   7
      TabIndex        =   24
      Top             =   3060
      Width           =   1035
   End
   Begin VB.TextBox tSCliente 
      Height          =   300
      Left            =   900
      MaxLength       =   7
      TabIndex        =   26
      Top             =   3420
      Width           =   1035
   End
   Begin MSComctlLib.StatusBar sbHelp 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   45
      Top             =   6255
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11536
            Key             =   "help"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "ERR !"
            TextSave        =   "ERR !"
            Key             =   "rydesul"
            Object.ToolTipText     =   "Error de conexón (Rydesul)."
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cLocal 
      Height          =   315
      Left            =   900
      TabIndex        =   39
      Top             =   5880
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.TextBox tArticulo 
      Height          =   300
      Left            =   900
      TabIndex        =   6
      Top             =   1500
      Width           =   3735
   End
   Begin MSMask.MaskEdBox tCi 
      Height          =   300
      Left            =   900
      TabIndex        =   3
      Top             =   1140
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      ForeColor       =   12582912
      PromptInclude   =   0   'False
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#.###.###-#"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tRuc 
      Height          =   300
      Left            =   900
      TabIndex        =   4
      Top             =   1140
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
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
      Mask            =   "## ### ### ####"
      PromptChar      =   "_"
   End
   Begin AACombo99.AACombo cEntregaA 
      Height          =   315
      Left            =   4980
      TabIndex        =   42
      Top             =   5520
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
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
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
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cLocalD 
      Height          =   315
      Left            =   5220
      TabIndex        =   8
      Top             =   1500
      Width           =   1935
      _ExtentX        =   3413
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
   Begin AACombo99.AACombo cSTipo 
      Height          =   315
      Left            =   1980
      TabIndex        =   20
      Top             =   2700
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "&ID Servicio"
      Height          =   255
      Left            =   4860
      TabIndex        =   21
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "&Local:"
      Height          =   255
      Left            =   4740
      TabIndex        =   7
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label dCliente 
      Caption         =   "idCliente"
      Height          =   255
      Left            =   6420
      TabIndex        =   58
      Top             =   4500
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "&ID Cambio"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   795
   End
   Begin VB.Label lLEntrega 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   15
      Left            =   2580
      TabIndex        =   57
      Top             =   5385
      Width           =   4755
   End
   Begin VB.Label lTEntrega 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de la Entrega"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   56
      Top             =   5280
      Width           =   1425
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   4260
      TabIndex        =   43
      Top             =   5940
      Width           =   675
   End
   Begin VB.Label lSCliente 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1980
      TabIndex        =   54
      Top             =   3420
      Width           =   4755
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "A &Quien:"
      Height          =   255
      Left            =   4260
      TabIndex        =   41
      Top             =   5580
      Width           =   675
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   5580
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   4560
      Width           =   675
   End
   Begin VB.Label lTFactura 
      BackStyle       =   0  'Transparent
      Caption         =   "&Factura:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label lCFactura 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2400
      TabIndex        =   53
      Top             =   4860
      Width           =   4755
   End
   Begin VB.Label lSProducto 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1980
      TabIndex        =   52
      Top             =   3060
      Width           =   4755
   End
   Begin VB.Label lTCambio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de la Factura para el Cambio"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   51
      Top             =   4260
      Width           =   2520
   End
   Begin VB.Label lLCambio 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   15
      Left            =   2640
      TabIndex        =   50
      Top             =   4365
      Width           =   4755
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Da&tos del Service ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "F/Co&mpra:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Boleta:"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label lFactura 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2400
      TabIndex        =   49
      Top             =   1860
      Width           =   4755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Factura:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lCliente 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2400
      TabIndex        =   48
      Top             =   1140
      Width           =   4755
   End
   Begin VB.Label lLDevolucion 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   15
      Left            =   2640
      TabIndex        =   47
      Top             =   1005
      Width           =   4755
   End
   Begin VB.Label lTDevolucion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de la Factura de Devolución"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   46
      Top             =   900
      Width           =   2460
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. &Serie:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Ser&vicio:"
      Height          =   255
      Left            =   2940
      TabIndex        =   29
      Top             =   3840
      Width           =   1275
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "A&bonado:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lTCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Local:"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   5940
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   675
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
      Begin VB.Menu MnuOpL1 
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
   End
   Begin VB.Menu MnuExit 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalir 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "&?"
      Begin VB.Menu MnuHelp 
         Caption         =   "&Ayuda"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim prmModificar As Boolean
Dim prmEntregadoEnFactura As Boolean
Dim prmLocalAQuien As Long, prmTipoLocalAQuien As Integer
Dim prmIdTraslado As Long
Dim prmEnviarMensaje As Boolean
Dim prmModificarSinMovs As Boolean 'Solo para modificar datos irrelevantes
Dim prmVaSucesoPorCuotas As Boolean 'Suceso para Registrar atraso de Cuotas

Dim prmCargando As Boolean

Public Function gbl_CargaConParametros(XidCambio As Long, XidService As Long)

    If XidCambio = 0 And XidService = 0 Then Exit Function
    
    If XidCambio <> 0 Then
        tID.Text = XidCambio
        CargoCambio XidCambio
        Exit Function
    End If
    
    If XidService <> 0 Then
        NuevoCambioCGSA XidService
    End If
    
End Function

Private Function NuevoCambioCGSA(idServicioCGSA As Long)
On Error Resume Next
Dim mIdCliente As Long
    
    AccionNuevo
    Me.Refresh
    
    prmCargando = True
    
    '1) x id de Servico de CGSA
    cons = "Select SerCodigo, SerFecha, SerLocalIngreso, ArtID, ArtCodigo, ArtNombre, Producto.*" & _
                " From Servicio, Producto, Articulo " & _
                " Where SerCodigo = " & idServicioCGSA & _
                " And SerProducto = ProCodigo And ProArticulo = ArtID"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        mIdCliente = rsAux!ProCliente
        
        tArticulo.Text = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
        tArticulo.Tag = rsAux!ArtID
        
        If Not IsNull(rsAux!SerLocalIngreso) Then BuscoCodigoEnCombo cLocalD, rsAux!SerLocalIngreso
        
        If Not IsNull(rsAux!ProFacturaS) Then tSerie.Text = Trim(rsAux!ProFacturaS)
        If Not IsNull(rsAux!ProFacturaN) Then tNumero.Text = Trim(rsAux!ProFacturaN)
        
    End If
    rsAux.Close
    
    BuscoDatosServiceCGSA idServicioCGSA
    BuscarCliente miId:=mIdCliente
    
    Foco cLocalD
    
    prmCargando = False
    
End Function

Private Sub bDevolver_Click()
    AccionEliminar
End Sub

Private Sub bEntregar_Click()
    On Error Resume Next
    
    If Trim(tFEntrega.Text) = "" Then
        FechaDelServidor
        tFEntrega.Text = Format(gFechaServidor, "dd/mm/yyyy hh:mm")
        If cLocal.ListIndex = -1 Then BuscoCodigoEnCombo cLocal, paCodigoDeSucursal
    Else
        tFEntrega.Text = ""
    End If
    
    Foco cEntregaA
    
End Sub

Private Sub bEntregar_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If cEntregaA.Enabled Then Foco cEntregaA: Exit Sub
        Foco tUsuario
    End If
    
End Sub

Private Sub cDistribuidor_Click()

    Dim bState As Boolean, bBkColor As Long
        
    If cDistribuidor.Value = vbChecked Then
        bState = True: bBkColor = vbWindowBackground
    Else
        bState = False: bBkColor = lFactura.BackColor
        'tDBoleta.Text = "": tDCompra.Text = "": tDCliente.Text = ""
    End If
    
    If cDistribuidor.Enabled Then
        tDBoleta.Enabled = bState: tDBoleta.BackColor = bBkColor
        tDCompra.Enabled = bState: tDCompra.BackColor = bBkColor
        tDCliente.Enabled = bState: tDCliente.BackColor = bBkColor
    End If
    
End Sub

Private Sub cDistribuidor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tDCompra.Enabled Then Foco tDCompra Else Foco cSTipo
    End If
End Sub

Private Sub cEntregaA_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Foco tUsuario
    End If
End Sub


Private Sub cLocal_Click()

    If Not cLocal.Enabled Or Not cLocalD.Enabled Then Exit Sub
    On Error Resume Next
    If cLocal.ListIndex <> -1 Then
        If cLocalD.ListIndex <> -1 Then
            If cLocalD.ItemData(cLocalD.ListIndex) <> cLocal.ItemData(cLocal.ListIndex) Then
                cLocalD.ForeColor = vbWhite
                cLocalD.BackColor = Colores.Rojo
                cLocalD.Font.Bold = True: cLocalD.SelLength = 0
            Else
                cLocalD.ForeColor = vbWindowText
                cLocalD.BackColor = vbWindowBackground
                cLocalD.Font.Bold = False: cLocalD.SelLength = 0
            End If
        End If
        
    End If
    
End Sub

Private Sub cLocal_GotFocus()
    sbHelp.Panels("help").Text = "Local donde se entrega la nueva mercadería."
End Sub

Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        If Trim(tFEntrega.Text) = "" Then
            If bEntregar.Enabled Then bEntregar.SetFocus: Exit Sub
        End If
        
        If cEntregaA.Enabled Then Foco cEntregaA: Exit Sub
        Foco tUsuario
    End If
End Sub

Private Sub cLocal_LostFocus()
    sbHelp.Panels("help").Text = ""
End Sub

Private Sub cLocalD_Click()
    
    If Not cLocal.Enabled Or Not cLocalD.Enabled Then Exit Sub
    On Error Resume Next
    If cLocal.ListIndex <> -1 Then
        If cLocalD.ListIndex <> -1 Then
            If cLocalD.ItemData(cLocalD.ListIndex) <> cLocal.ItemData(cLocal.ListIndex) Then
                cLocalD.ForeColor = vbWhite
                cLocalD.BackColor = Colores.Rojo
                cLocalD.Font.Bold = True: cLocalD.SelLength = 0
            Else
                cLocalD.ForeColor = vbWindowText
                cLocalD.BackColor = vbWindowBackground
                cLocalD.Font.Bold = False: cLocalD.SelLength = 0
            End If
        End If
        
    End If

End Sub

Private Sub cLocalD_GotFocus()
    sbHelp.Panels("help").Text = "Local donde se recibe el artículo roto."
End Sub

Private Sub cLocalD_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        On Error Resume Next
        If cLocalD.ListIndex = -1 Then Exit Sub
        If Val(tArticulo.Tag) <> 0 Then
            If cDistribuidor.Enabled Then
                If cDistribuidor.Value = vbChecked Then Foco tDCompra Else Foco tSerie
            Else
                If cLocal.Enabled Then cLocal.SetFocus Else Foco tUsuario
            End If
            Exit Sub
        End If
    End If
    
End Sub

Private Sub cLocalD_LostFocus()
    sbHelp.Panels("help").Text = ""
End Sub

Private Sub cSTipo_Change()
    If tSIDServicio.Text <> "" Then tSIDServicio.Text = ""
End Sub

Private Sub cSTipo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If cSTipo.ListIndex = -1 Then Foco tCArticulo: Exit Sub
        If cSTipo.ItemData(cSTipo.ListIndex) = TipoService.CGSA Then
            If Val(tSIDServicio.Text) <> 0 And Trim(tSAbonado.Text) <> "" Then
                Foco tCArticulo
            Else
                Foco tSIDServicio
            End If
        Else
            Foco tSAbonado
        End If
    End If
    
End Sub

Private Sub Form_Load()
On Error Resume Next
    ObtengoSeteoForm Me
    InicializoForm
    
    FechaDelServidor
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    lLDevolucion.Left = lTDevolucion.Left + lTDevolucion.Width + 80
    lLDevolucion.Width = Me.ScaleWidth - lLDevolucion.Left - 80
    
    lLCambio.Left = lTCambio.Left + lTCambio.Width + 80
    lLCambio.Width = Me.ScaleWidth - lLCambio.Left - 80
    
    lLEntrega.Left = lTEntrega.Left + lTEntrega.Width + 80
    lLEntrega.Width = Me.ScaleWidth - lLEntrega.Left - 80
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    EndMain
End Sub

Private Sub InicializoForm()
    
    On Error Resume Next
    
    sbHelp.Panels("rydesul").Visible = Not prmBDRydesul
    
    LimpioFicha
    
    cons = "Select LocCodigo, LocNombre from Local " & _
               " Where LocTipo = " & TipoLocal.Deposito & _
               " Order by LocNombre"
    CargoCombo cons, cLocal
    CargoCombo cons, cLocalD
    
    BuscoCodigoEnCombo cLocal, paCodigoDeSucursal
        
    cons = "Select AQECodigo, AQENombre from AQuienEntrega Order by AQENombre"
    CargoCombo cons, cEntregaA
    
    With tUsuario
        .Connect cBase
        .UserID = 0
        .EnabledButton = False
    End With
    
    With cSTipo
        .AddItem "CGSA": .ItemData(.NewIndex) = TipoService.CGSA
        .AddItem "Rydesul": .ItemData(.NewIndex) = TipoService.Rydesul
    End With

    HabilitoCampos
    
End Sub

Private Sub LimpioFicha()
    
    prmVaSucesoPorCuotas = False
    bDevolver.Visible = False
    
    prmIdTraslado = 0
    prmEntregadoEnFactura = False
    prmLocalAQuien = 0

    tCi.Text = "": tRuc.Text = ""
    tArticulo.Text = ""
    
    tSerie.Text = "": tNumero.Text = ""
    cDistribuidor.Value = vbUnchecked
    tDCompra.Text = "": tDBoleta.Text = "": tDCliente.Text = ""
    
    cSTipo.Text = "": tSIDServicio.Text = ""
    tSAbonado.Text = "": lSProducto.Caption = ""
    tSCliente.Text = "": lSCliente.Caption = ""
    tSFServicio.Text = "": tSSerie.Text = ""
    
    tCArticulo.Text = ""
    tCSerie.Text = "": tCNumero.Text = "": lCFactura.Caption = ""
    
    tFEntrega.Text = "": cLocal.Text = "": cEntregaA.Text = ""
    tUsuario.UserName = "": tUsuario.UserID = 0
    
    tFEntrega.Tag = ""      'Fecha de entrega grabada
    lTFactura.Tag = 0       'Factura grabada
    
    'cLocalD.ForeColor = vbWindowText
    'cLocalD.BackColor = vbWindowBackground
    'cLocalD.Font.Bold = False
    'cLocalD.SelLength = 0
    cLocalD.Text = ""
End Sub

Private Sub lTCliente_Click()
    'If tCi.ZOrder Then tCi.SetFocus Else tRuc.SetFocus
    tCi.SetFocus
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuHelp_Click()
    AccionMenuHelp
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = 0
    tSerie.Text = "": tNumero.Text = ""
End Sub

Private Sub tArticulo_GotFocus()
    tArticulo.SelStart = 0: tArticulo.SelLength = Len(tArticulo.Text)
    sbHelp.Panels("help").Text = ""
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tArticulo.Text) = "" Then Exit Sub
        If Val(tArticulo.Tag) <> 0 Then
            Foco cLocalD: Exit Sub
        End If
        
        BuscoArticulo tArticulo
    End If
        
End Sub

Private Sub tArticulo_LostFocus()
    sbHelp.Panels("help").Text = ""
End Sub

Private Function ValidoGrabar() As Boolean

    ValidoGrabar = False
    
    If Val(lCliente.Tag) = 0 Then
        MsgBox "Falta ingresar el cliente de la factura de devolución.", vbExclamation, "Faltan datos de Devolución"
        tCi.SetFocus: Exit Function
    End If
    If Val(tArticulo.Tag) = 0 Then
        MsgBox "Falta ingresar el artículo que se devuelve.", vbExclamation, "Faltan datos de Devolución"
        tArticulo.SetFocus: Exit Function
    End If
    If Val(lFactura.Tag) = 0 Then
        MsgBox "Falta ingresar la factura de devolución.", vbExclamation, "Faltan datos de Devolución"
        tSerie.SetFocus: Exit Function
    End If
    If cLocalD.ListIndex = -1 Then
        MsgBox "Falta ingresar el local en donde inrgesa la mercadería devuelta.", vbExclamation, "Faltan datos de Devolución"
        cLocalD.SetFocus: Exit Function
    End If
    
    If Val(lFactura.Tag) = Val(lCFactura.Tag) Then
        MsgBox "El documento para el cambio no debe ser el mismo de la devolución.", vbExclamation, "Documento No Válido"
        tCSerie.SetFocus: Exit Function
    End If
    
    'If cSTipo.ListIndex = -1 Then
    '    MsgBox "Falta ingresar el local del service.", vbExclamation, "Faltan datos del Service"
    '    cSTipo.SetFocus: Exit Function
    'End If
    
    If Val(tCArticulo.Tag) = 0 Then
        MsgBox "Falta ingresar el artículo que se entrega.", vbExclamation, "Faltan datos del Cambio"
        tCArticulo.SetFocus: Exit Function
    End If
    
    If Trim(tFEntrega.Text) <> "" Then
        If cLocal.ListIndex = -1 Then
            MsgBox "Falta seleccionar el local donde se entrega la mercadería.", vbExclamation, "Faltan datos de la Entrega"
            cLocal.SetFocus: Exit Function
        End If
        If cEntregaA.ListIndex = -1 Then
            MsgBox "Falta seleccionar a quien se entrega la mercadería.", vbExclamation, "Faltan datos de la Entrega"
            cEntregaA.SetFocus: Exit Function
        End If
    End If
    If tUsuario.UserID = 0 Then
        MsgBox "Falta ingresar el usuario que realiza la entrega.", vbExclamation, "Faltan datos de Devolución"
        tUsuario.SetFocus: Exit Function
    End If
    
    Screen.MousePointer = 11
    prmEnviarMensaje = False
    prmLocalAQuien = 0
    
    If cEntregaA.ListIndex <> -1 Then
        'Valido si se le puede entregar sin factura
        cons = "Select isNull(AQEConFactura, 0) as AQEConFactura, isNull(AQELocal, 0) as AQELocal, LocTipo " & _
                    " From AQuienEntrega " & _
                            " Left Outer Join Local On AQELocal = LocCodigo" & _
                   " Where AQECodigo = " & cEntregaA.ItemData(cEntregaA.ListIndex)
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            prmLocalAQuien = rsAux!AQELocal
            If Not IsNull(rsAux!LocTipo) Then prmTipoLocalAQuien = rsAux!LocTipo Else prmTipoLocalAQuien = 0
            
            If rsAux!AQEConFactura = 1 And Val(lCFactura.Tag) = 0 And Trim(tFEntrega.Text) <> "" Then
                
                If MsgBox("A " & Trim(cEntregaA.Text) & " no se le puede entregar mercadería sin factura." & vbCrLf & "Desea grabar igualmente ?", vbQuestion + vbYesNo, "No se puede Entregar sin Factura") = vbNo Then
                    rsAux.Close
                    tCSerie.SetFocus: Screen.MousePointer = 0
                    Exit Function
                End If
                tFEntrega.Text = ""
                
                If MsgBox("Ud. quiere enviar un mensaje automático solicitando la facturación del cambio.", vbQuestion + vbYesNo, "Enviar Mensaje Automático") = vbYes Then
                    prmEnviarMensaje = True
                End If
            End If
            
        End If
        rsAux.Close
    End If
    
    'Si no hay factura y hay fecha de entrega --> tengo q hacer traslado del local al local AQuien
    'Y nunca se hizo el traslado
    If Trim(tFEntrega.Text) <> "" And Val(lCFactura.Tag) = 0 And prmIdTraslado = 0 Then
        If prmLocalAQuien = 0 Then
            MsgBox "El dato 'A Quien Entrega', no tiene asociado un local para hacer el traslado." & vbCrLf & _
                        "Cuando se da por entregada la mercadería y no se ingresa la factura, el sistema hace un traslado " & _
                        "desde el local de entrega al local 'a quien'." & vbCrLf & vbCrLf & _
                        "El proceso no puede continuar. Consulte con el Administrador.", vbInformation, "Falta Local para Traslado"
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    
    If Not prmModificar Then
        'Valido cantidad de artículos devueltos para la factura ingresada.      ---------------------------------------------------------------
        cons = "Select (RenCantidad - RenARetirar) as Q, Count(CArDArticulo) as QDev " & _
                    " From Renglon " & _
                    " Left Outer Join CambioArticulo On RenDocumento = CArDDocumento And RenArticulo = CArDArticulo" & _
                    " Where RenDocumento = " & Val(lFactura.Tag) & _
                    " And RenArticulo = " & Val(tArticulo.Tag) & _
                    " Group by RenDocumento, RenCantidad, RenARetirar"
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            If rsAux!Q = 0 Then
                MsgBox "El decumento ingresado para la devolución, no tiene artículos disponibles para devolver." & vbCrLf & _
                            "Verifique los datos en Detalle de Factura.", vbExclamation, "Documento sin Artículos para devolver"
                rsAux.Close
                Screen.MousePointer = 0: Exit Function
            End If
            If Not IsNull(rsAux!QDev) Then
                If rsAux!Q <= rsAux!QDev Then
                    MsgBox "El decumento ingresado para la devolución, no tiene artículos disponibles para devolver." & vbCrLf & _
                                "Controle las devoluciones ya ingresadas para el documento.", vbExclamation, "Documento sin Artículos para devolver"
                    rsAux.Close
                    Screen.MousePointer = 0: Exit Function
                End If
            End If
        Else
            MsgBox "El decumento ingresado para la devolución, no corresponde al artículo ingresado para devolver.", vbExclamation, "Documento sin Artículo Ingresado"
            rsAux.Close
            Screen.MousePointer = 0: Exit Function
        End If
        rsAux.Close
        '----------------------------------------------------------------------------------------------------------------------------------------------
    End If
    
    Screen.MousePointer = 0
    
    'Si se ingresa factura hay que darlo como entregado.
    If Val(lCFactura.Tag) <> 0 And Trim(tFEntrega.Text) = "" Then
        MsgBox "Al ingresar una factura, debe dar el artículo como entregado." & vbCrLf & _
                    "Presione el botón 'Entregar'", vbInformation, "Falta dar como Entregado"
        bEntregar.SetFocus: Exit Function
    End If
    
    'Si esta entregado c/factura y se hizo un traslado no se puede xq el a retirar de la total lo baja la factura.
    'Se debe hacer devolucion de mercadería.
    
    If prmEntregadoEnFactura And prmIdTraslado <> 0 Then
        MsgBox "El artículo se entregó con la factura por entrega de mercadería." & vbCrLf & _
                    "En los pasos anteriores el sistema trasladó la mercadería desde '" & Trim(cLocal.Text) & "', al local asociado a '" & _
                    Trim(cEntregaA.Text) & "'" & vbCrLf & vbCrLf & _
                    "Para continuar, la mercadería debe estar disponible para entregar en la factura." & vbCrLf & _
                    "Realice la devolución de la mercadería para entregarla desde acá.", vbInformation, "Mercadería Entregada por Factura"
        Exit Function
    End If
    
    ValidoGrabar = True
                
End Function

Private Sub GrabarDatosIrrelevantes()

On Error GoTo errGrabar
    
    If MsgBox("Confirma grabar los datos del cambio de artículo ?", vbQuestion + vbYesNo, "Grabar Datos") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    
    Dim mID As Long
    mID = Trim(tID.Text)
    
    cons = "Select * from CambioArticulo Where CArID = " & mID
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    rsAux.Edit
    
    If Trim(tSSerie.Text) <> "" Then rsAux!CArSSerie = Trim(tSSerie.Text) Else rsAux!CArSSerie = Null
    
    If tUsuario.UserID <> 0 Then rsAux!CArCUsuario = tUsuario.UserID Else rsAux!CArUsuario = Null
    
    rsAux.Update: rsAux.Close

        
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub

errGrabar:
    clsGeneral.OcurrioError "Error al grabar los datos del cambio.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub

Private Sub AccionGrabar()

On Error GoTo errValidar
        
        
    If Trim(tID.Text) = "" Then prmModificar = False Else prmModificar = True
    If prmModificar And prmModificarSinMovs Then
        GrabarDatosIrrelevantes
        Exit Sub
    End If
    
    sbHelp.Panels("help").Text = "Validando datos, espere ...": sbHelp.Refresh
    If Not ValidoGrabar Then
        sbHelp.Panels("help").Text = "": sbHelp.Refresh
        Exit Sub
    End If
    sbHelp.Panels("help").Text = "": sbHelp.Refresh
    
    Dim mInfo As String: mInfo = ""
    If Val(lCFactura.Tag) <> 0 Then
        If Not prmEntregadoEnFactura Then
            mInfo = "Se ha ingresado la factura para realizar el cambio de Mercadería." & vbCrLf & _
                        "El sistema realizará los movimientos de stock para entregar un " & Trim(tCArticulo.Text) & "."
        Else
            mInfo = "Se ha ingresado la factura para realizar el cambio de Mercadería." & vbCrLf & _
                        "El artículo ya fue entregado por la factura, no se realizarán movimientos de stock."
        End If
    Else
        If Trim(tFEntrega.Text) <> "" Then
            mInfo = "Se ha ingresado la fecha de entrega de la Mercadería." & vbCrLf & _
                        "El sistema no moverá el stock hasta que ud. asigne una factura al cambio de mercadería." & vbCrLf & _
                        "Si se realizará un traslado desde el local a el local 'A Quien'." & vbCrLf & vbCrLf & _
                        "El proceso quedará pendiente hasta asignar la factura."
        End If
    End If
    
    If mInfo <> "" Then MsgBox mInfo, vbInformation, "Información"
            
    If MsgBox("Confirma grabar los datos del cambio de artículo ?", vbQuestion + vbYesNo, "Grabar Datos") = vbNo Then Exit Sub
    
    Dim sc_Usuario As Long, sc_Defensa As String, sc_Autoriza As Long
    
    If prmVaSucesoPorCuotas Then        'Pido Suceso por Cambio con Ctas Arasadas   -------------------------------------------------
        Dim objSuceso As New clsSuceso
        With objSuceso
            .TipoSuceso = prmSucesoMovStock
            .ActivoFormulario paCodigoDeUsuario, "Cliente con Cuotas Atrasadas", cBase
        
            Me.Refresh
            sc_Usuario = .RetornoValor(Usuario:=True)
            sc_Defensa = .RetornoValor(Defensa:=True)
            sc_Autoriza = .Autoriza
        End With
        
        Set objSuceso = Nothing
        If sc_Usuario = 0 Then Screen.MousePointer = 0: Exit Sub 'Abortó el ingreso del suceso
    End If      '------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errorBT
    Dim bDarIngresoRoto As Boolean: bDarIngresoRoto = False
    
    FechaDelServidor

    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    Screen.MousePointer = 11
    
    Dim mID As Long
    If Trim(tID.Text) <> "" Then mID = Trim(tID.Text) Else mID = 0
    
    cons = "Select * from CambioArticulo Where CArID = " & mID
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then rsAux.AddNew Else rsAux.Edit
    
    If Not prmModificar Then
        'Datos del la Devolución  --------------------------------------------------------------
        rsAux!CArDCliente = Val(lCliente.Tag)
        rsAux!CArDArticulo = Val(tArticulo.Tag)
        rsAux!CArDDocumento = Val(lFactura.Tag)
        rsAux!CArDLocal = cLocalD.ItemData(cLocalD.ListIndex)
        
        If cDistribuidor.Value = vbChecked Then
            If Trim(tDCompra.Text) <> "" Then rsAux!CArDFCompra = Format(tDCompra.Text, "mm/dd/yyyy") Else rsAux!CArDFCompra = Null
            If Trim(tDBoleta.Text) <> "" Then rsAux!CArDBoleta = Trim(tDBoleta.Text) Else rsAux!CArDBoleta = Null
            If Trim(tDCliente.Text) <> "" Then rsAux!CArDNombre = Trim(tDCliente.Text) Else rsAux!CArDNombre = Null
        Else
            rsAux!CArDFCompra = Null
            rsAux!CArDBoleta = Null
            rsAux!CArDNombre = Null
        End If
    
        'Datos del Service          --------------------------------------------------------------
        If cSTipo.ListIndex <> -1 Then rsAux!CArSTipoService = cSTipo.ItemData(cSTipo.ListIndex) Else rsAux!CArSTipoService = Null
        If Trim(tSIDServicio.Text) <> "" Then rsAux!CArSIDService = Trim(tSIDServicio.Text) Else rsAux!CArSIDService = Null
        
        If Trim(tSAbonado.Text) <> "" Then rsAux!CArSIdAbonado = Trim(tSAbonado.Text) Else rsAux!CArSIdAbonado = Null
        If Trim(tSCliente.Text) <> "" Then rsAux!CArSIdCliente = Trim(tSCliente.Text) Else rsAux!CArSIdCliente = Null
        If Trim(tSFServicio.Text) <> "" Then rsAux!CArSFecha = Format(tSFServicio.Text, "mm/dd/yyyy hh:mm") Else rsAux!CArSFecha = Null
        
    End If
    
    If Trim(tSSerie.Text) <> "" Then rsAux!CArSSerie = Trim(tSSerie.Text) Else rsAux!CArSSerie = Null
    
    'Datos del la Factura del Cambio    ---------------------------------------------------
    rsAux!CArCArticulo = Val(tCArticulo.Tag)
        
    If Trim(tFEntrega.Text) <> "" Then
        rsAux!CArCFEntrega = Format(tFEntrega.Text, "mm/dd/yyyy hh:mm")
    Else
        rsAux!CArCFEntrega = Null
    End If
    
    If Val(lCFactura.Tag) <> 0 Then rsAux!CArCDocumento = Val(lCFactura.Tag) Else rsAux!CArCDocumento = Null
    
    If cLocal.ListIndex <> -1 Then rsAux!CArCLocal = cLocal.ItemData(cLocal.ListIndex) Else rsAux!CArCLocal = Null
    If cEntregaA.ListIndex <> -1 Then rsAux!CArCAQuien = cEntregaA.ItemData(cEntregaA.ListIndex) Else rsAux!CArCAQuien = Null
    
    If tUsuario.UserID <> 0 Then rsAux!CArCUsuario = tUsuario.UserID Else rsAux!CArUsuario = Null
    
    rsAux.Update: rsAux.Close
    
    If Not prmModificar Then
        cons = "Select Top 1 * from CambioArticulo " & _
                   " Where CArDDocumento = " & Val(lFactura.Tag) & _
                   " And CArDCliente = " & Val(lCliente.Tag) & _
                   " And CArDArticulo = " & Val(tArticulo.Tag) & _
                   " And CArCArticulo = " & Val(tCArticulo.Tag) & _
                   " Order by CArID desc"
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then mID = rsAux!CArID
        rsAux.Close
    End If

    'Updateo comentario de Visita Servicio CGSA
    If Not prmModificar And cSTipo.ListIndex <> -1 And Val(tSIDServicio.Text) <> 0 Then
        If cSTipo.ItemData(cSTipo.ListIndex) = TipoService.CGSA Then
            
            cons = "Select * from ServicioVisita" & _
                       " Where VisServicio = " & Val(tSIDServicio.Text) & " And VisTipo = 1"
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then
                rsAux.Edit
                If IsNull(rsAux!VisComentario) Then
                    rsAux!VisComentario = "ID CambioArt=" & mID
                Else
                    rsAux!VisComentario = "ID CambioArt=" & mID & "; " & Trim(rsAux!VisComentario)
                End If
                rsAux.Update
            End If
            rsAux.Close
        End If
    End If

    If Trim(tFEntrega.Text) <> "" Then  'Se Entregó
        If Val(lCFactura.Tag) = 0 Then
            If prmIdTraslado = 0 Then
                'Si no hay factura y nunca se hizo un traslado
                '1)  Pasar la mercaeria del local al Local a Quien
                '2)  Hacer Traslado
                Dim mIDTraslado As Long
                mIDTraslado = MuevoTrasladoMercaderia(mID, cLocal.ItemData(cLocal.ListIndex), prmLocalAQuien)
                MuevoStockLocalALocal TipoDocumento.Traslados, mIDTraslado, Val(tCArticulo.Tag), _
                                TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), prmTipoLocalAQuien, prmLocalAQuien
                
                'Updateo con ID de Traslado
                cons = "Select * from CambioArticulo Where CArID = " & mID
                Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                rsAux.Edit
                rsAux!CArCIDTraslado = mIDTraslado
                rsAux.Update: rsAux.Close
            
                bDarIngresoRoto = True  'No hay Factura, tengo que hacer traslado y doy ingreso del roto
            End If
        Else
            '1) Si hay factura ---> hay Fecha Entrega
            'Para mover stock, veo si ya fue movido por la factura
            If Not prmEntregadoEnFactura Then
                'Esto es el "A Retirar -1"
                MuevoStockDocumento Val(tCSerie.Tag), Val(lCFactura.Tag), Val(tCArticulo.Tag), cLocal.ItemData(cLocal.ListIndex)
            End If
            
            'Si nunca se hizo traslado bajo la mercaderia del local
             If prmIdTraslado = 0 Then
                bDarIngresoRoto = True  'Como no se hizo traslado no inrgeso nunca el roto
                
                If Not prmEntregadoEnFactura Then
                    MuevoStockBajaLocal Val(tCSerie.Tag), Val(lCFactura.Tag), Val(tCArticulo.Tag), _
                                                cLocal.ItemData(cLocal.ListIndex), TipoLocal.Deposito
                End If
             Else
                    'Si no la bajo del local a quien
                    'Como Anteriormente se hizo un traslado desde el local hasta el aquien --> bajar del stock del Aquien
                    MuevoStockBajaLocal Val(tCSerie.Tag), Val(lCFactura.Tag), Val(tCArticulo.Tag), _
                                                     prmLocalAQuien, prmTipoLocalAQuien
            End If
        End If
    End If
            
    'If Not prmModificar Then
    If bDarIngresoRoto Then
        MuevoStockIngreso Val(tSerie.Tag), Val(lFactura.Tag), Val(tArticulo.Tag), cLocalD.ItemData(cLocalD.ListIndex)
    End If
    
    'Si el Proceso está terminado grabo los cometarios
    If Val(lCFactura.Tag) <> 0 Then GraboComentarios (mID)
    
    If prmVaSucesoPorCuotas Then
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, prmSucesoMovStock, paCodigoDeTerminal, sc_Usuario, lFactura.Tag, tCArticulo.Tag, _
                             Descripcion:="Cambio de Artículo (" & mID & ") / Cliente debe ctas.", Defensa:=Trim(sc_Defensa), idCliente:=lCliente.Tag, idAutoriza:=sc_Autoriza
    End If
    
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------
    
    If prmEnviarMensaje And prmMenUsuarioCambioArticulo <> "" Then EnviarMensaje mID, tCArticulo.Text
    
    If Trim(tID.Text) = "" Then tID.Text = mID
    
    AccionCancelar
    
    Screen.MousePointer = 0

    Exit Sub

errValidar:
    clsGeneral.OcurrioError "Error al procesar los datos para grabar.", Err.Description
    Screen.MousePointer = 0: Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al grabar los datos del cambio.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub

Private Sub MuevoStockDocumento(mTipoDoc As Integer, mDocumento As Long, mArticulo As Long, mLocal As Long)
    
    sbHelp.Panels("help").Text = "Actualizando documento...": sbHelp.Refresh
    
    cons = "Select * from Renglon " & _
                " Where RenDocumento = " & mDocumento & _
                " And RenArticulo = " & mArticulo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsAux.Edit
    rsAux!RenARetirar = rsAux!RenARetirar - 1
    rsAux.Update: rsAux.Close
    
    'Marco la Baja del STOCK AL LOCAL
    'Genero Movimiento
'    MarcoMovimientoStockFisico tUsuario.UserID, TipoLocal.Deposito, mLocal, mArticulo, 1, paEstadoSano, -1, mTipoDoc, mDocumento
    'Bajo del Stock en Local
'    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, mLocal, mArticulo, 1, paEstadoSano, -1
    
    'Marco el Movimiento del STOCK VIRTUAL
    'Genero Movimiento
    MarcoMovimientoStockEstado tUsuario.UserID, mArticulo, 1, TipoMovimientoEstado.ARetirar, -1, mTipoDoc, mDocumento, mLocal
    'Bajo del Stock Total
    MarcoMovimientoStockTotal mArticulo, TipoEstadoMercaderia.Virtual, TipoMovimientoEstado.ARetirar, 1, -1

    sbHelp.Panels("help").Text = "": sbHelp.Refresh
End Sub

Private Sub MuevoStockLocalALocal(mTipoDoc As Integer, mDocumento As Long, mArticulo As Long, _
                mBajaTL As Integer, mBajaL As Long, mAltaTL As Integer, mAltaL As Long)

    sbHelp.Panels("help").Text = "Grabando Traslado (local a local)... ": sbHelp.Refresh
    '1) Marco la Baja del STOCK AL LOCAL   -------------
    'Genero Movimiento
    MarcoMovimientoStockFisico tUsuario.UserID, mBajaTL, mBajaL, mArticulo, 1, paEstadoSano, -1, mTipoDoc, mDocumento
    'Bajo del Stock en Local
    MarcoMovimientoStockFisicoEnLocal mBajaTL, mBajaL, mArticulo, 1, paEstadoSano, -1
    
    '2) Marco la Alta del STOCK AL LOCAL   -------------
    'Genero Movimiento
    MarcoMovimientoStockFisico tUsuario.UserID, mAltaTL, mAltaL, mArticulo, 1, paEstadoSano, 1, mTipoDoc, mDocumento
    'Bajo del Stock en Local
    MarcoMovimientoStockFisicoEnLocal mAltaTL, mAltaL, mArticulo, 1, paEstadoSano, 1
    
    sbHelp.Panels("help").Text = "": sbHelp.Refresh
End Sub

Private Sub MuevoStockBajaLocal(mTipoDoc As Integer, mDocumento As Long, mArticulo As Long, mLocal As Long, mLocalTipo As Integer)

    sbHelp.Panels("help").Text = "Grabando Baja de Stock...": sbHelp.Refresh
    'Marco la Baja del STOCK AL LOCAL
    'Genero Movimiento
    MarcoMovimientoStockFisico tUsuario.UserID, mLocalTipo, mLocal, mArticulo, 1, paEstadoSano, -1, mTipoDoc, mDocumento
    
    'Bajo del Stock en Local
    MarcoMovimientoStockFisicoEnLocal mLocalTipo, mLocal, mArticulo, 1, paEstadoSano, -1

    '!!! Esto no va porque lo mueve el documento --> A Retirar -1 y no el físico
    'Bajo del Stock Total
    'MarcoMovimientoStockTotal mArticulo, TipoEstadoMercaderia.Fisico, paEstadoSano, 1, -1

    sbHelp.Panels("help").Text = "": sbHelp.Refresh
End Sub
    
Private Sub MuevoStockIngreso(mTipoDoc As Integer, mDocumento As Long, mArticulo As Long, mLocal As Long, Optional mEsAlta As Integer = 1)

    sbHelp.Panels("help").Text = "Grabando Alta a Stock...": sbHelp.Refresh
    'Marco la Alta A Recuperar en STOCK
    
    'Genero Movimiento
    MarcoMovimientoStockFisico tUsuario.UserID, TipoLocal.Deposito, mLocal, mArticulo, 1, paEstadoRoto, mEsAlta, mTipoDoc, mDocumento
    
    'Alta en Local
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, mLocal, mArticulo, 1, paEstadoRoto, mEsAlta
    
    'Alta ARec. al Stock Total
    MarcoMovimientoStockTotal mArticulo, TipoEstadoMercaderia.Fisico, paEstadoRoto, 1, mEsAlta

    sbHelp.Panels("help").Text = "": sbHelp.Refresh
    
End Sub

Private Function MuevoTrasladoMercaderia(idCambio As Long, idOrigen As Long, idDestino As Long) As Long

Dim rsTra As rdoResultset
Dim mIDT As Long
    
    sbHelp.Panels("help").Text = "Grabando datos Traslado ...": sbHelp.Refresh
    
    cons = "Select * from Traspaso Where TraCodigo = 0"
    Set rsTra = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsTra.AddNew
    rsTra!TraFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    rsTra!TraLocalOrigen = idOrigen
    rsTra!TraLocalDestino = idDestino
    rsTra!TraComentario = "Cambio Artículo en Garantía. ID:" & idCambio
    
    rsTra!TraFechaEntregado = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    rsTra!TraUsuarioInicial = tUsuario.UserID
    rsTra!TraUsuarioFinal = tUsuario.UserID
    
    rsTra.Update: rsTra.Close
    
    'Saco el código del insertado.--------------------------------------------
    cons = "Select MAX(TraCodigo) From Traspaso" & _
               " Where TraLocalOrigen = " & idOrigen & _
               " And TraLocalDestino = " & idDestino & _
               " And TraFecha = '" & Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss") & "'" & _
               " And TraUsuarioInicial = " & tUsuario.UserID & _
               " And TraUsuarioFinal = " & tUsuario.UserID
               
    Set rsTra = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    mIDT = rsTra(0)
    rsTra.Close
            
    cons = "Select * from RenglonTraspaso Where RTrTraspaso = " & mIDT
    Set rsTra = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsTra.AddNew
    rsTra!RTrTraspaso = mIDT
    rsTra!RTrArticulo = Val(tCArticulo.Tag)
    rsTra!RTrEstado = paEstadoSano
    rsTra!RTrCantidad = 1
    rsTra!RTrPendiente = 0
    rsTra.Update
    rsTra.Close
    
    MuevoTrasladoMercaderia = mIDT
    
    sbHelp.Panels("help").Text = "": sbHelp.Refresh
    
End Function

Private Sub tCArticulo_Change()
    On Error Resume Next
    If tCArticulo.Enabled And Val(tCArticulo.Tag) <> 0 Then
        tCSerie.Text = "": tCNumero.Text = ""
        tFEntrega.Text = "": cLocal.Text = ""
    End If
    tCArticulo.Tag = 0
End Sub

Private Sub tCArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tCArticulo.Text) = "" Then Exit Sub
        If Val(tCArticulo.Tag) <> 0 Then Foco tCSerie: Exit Sub
        
        BuscoArticulo tCArticulo
    End If
    
End Sub

Private Sub tCi_Change()

    If prmCargando Then Exit Sub
    lCliente.Tag = 0
    lCliente.Caption = ""
    If Trim(tArticulo.Text) <> "" And tArticulo.Enabled Then tArticulo.Text = ""
    
End Sub

Private Sub tCi_GotFocus()
    tCi.SelStart = 0: tCi.SelLength = (Len(tCi.FormattedText))
    sbHelp.Panels("help").Text = "Cliente asociado a la factuara de devolución.  [C]- Cambia CI/Ruc.   [F4]- Buscar."
End Sub

Private Sub tCi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2, vbKeyR, vbKeyE, vbKeyC: tCi.Visible = False: tRuc.Visible = True: tRuc.SetFocus: lCliente.Tag = 0: lCliente.Caption = ""
        
        Case vbKeyReturn
                If Val(lCliente.Tag) = 0 And Trim(tCi.Text) <> "" Then
                    If Len(tCi.Text) < 7 Then Exit Sub
                    If Len(tCi.Text) = 7 Then tCi.Text = clsGeneral.AgregoDigitoControlCI(Trim(tCi.Text))
                    BuscarCliente miCi:=Trim(tCi.Text)
                Else
                    tArticulo.SetFocus
                End If
        
        Case vbKeyF4: BuscarClientes TipoCliente.Cliente
    End Select
    
End Sub

Private Sub BuscarClientes(aTipoCliente As Integer)
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    Dim objBuscar As New clsBuscarCliente
    Dim aTipo As Integer, aCliente As Long
    
    If aTipoCliente = TipoCliente.Cliente Then objBuscar.ActivoFormularioBuscarClientes cBase, Persona:=True
    If aTipoCliente = TipoCliente.Empresa Then objBuscar.ActivoFormularioBuscarClientes cBase, Empresa:=True
    Me.Refresh
    aTipo = objBuscar.BCTipoClienteSeleccionado
    aCliente = objBuscar.BCClienteSeleccionado
    Set objBuscar = Nothing
    
    If aCliente <> 0 Then BuscarCliente miId:=aCliente, miTipo:=aTipo
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCi_LostFocus()
    sbHelp.Panels("help").Text = ""
End Sub

Private Sub tCNumero_Change()
    
    If Val(lCFactura.Tag) <> 0 And tCNumero.Enabled Then
        lCFactura.Caption = ""
        If Trim(tFEntrega.Tag) = "" Then 'No estaba entregado ...
            tFEntrega.Text = ""
            bEntregar.Enabled = True
        End If
    End If
    lCFactura.Tag = 0

End Sub

Private Sub tCNumero_GotFocus()
    tCNumero.SelStart = 0: tCNumero.SelLength = Len(tCNumero.Text)
End Sub

Private Sub tCNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(lCFactura.Tag) <> 0 Or (Trim(tCSerie.Text) = "" And Trim(tCNumero.Text) = "") Then
            If cLocal.Enabled Then cLocal.SetFocus: Exit Sub
            If cEntregaA.Enabled Then cEntregaA.SetFocus: Exit Sub
            tUsuario.SetFocus
        Else
            BuscoDocumentoCambio
        End If
    End If
    
End Sub

Private Sub tCSerie_Change()
 
    If Val(lCFactura.Tag) <> 0 And tCSerie.Enabled Then
        lCFactura.Caption = ""
        If Trim(tFEntrega.Tag) = "" Then 'No estaba entregado ...
            tFEntrega.Text = ""
            bEntregar.Enabled = True
        End If
    End If
    lCFactura.Tag = 0
    
End Sub

Private Sub tCSerie_GotFocus()
    tCSerie.SelStart = 0: tCSerie.SelLength = Len(tCSerie.Text)
End Sub

Private Sub tCSerie_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Val(lCFactura.Tag) <> 0 Then Foco tFEntrega Else Foco tCNumero
    End If
    
End Sub

Private Sub tDBoleta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDCliente
End Sub

Private Sub tDCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(lFactura.Tag) = 0 Then Foco tSerie Else Foco cSTipo
    End If
End Sub

Private Sub tDCompra_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tDCompra.Text) <> "" Then
            If IsDate(tDCompra.Text) Then
                tDCompra.Text = Format(tDCompra.Text, "dd/mm/yyyy")
            Else
                Exit Sub
            End If
        End If
        Foco tDBoleta
    End If

End Sub

Private Sub tFEntrega_Change()
    On Error Resume Next
    If tFEntrega.Text = "" Then
        tFEntrega.BackColor = Colores.Inactivo
    Else
        tFEntrega.BackColor = Colores.clVerde
    End If
End Sub

Private Sub tFEntrega_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tFEntrega.Text) <> "" Then
            If IsDate(tFEntrega.Text) Then
                tFEntrega.Text = Format(tFEntrega.Text, "dd/mm/yyyy hh:mm")
            Else
                Exit Sub
            End If
        End If
        Foco cLocal
    End If
    
End Sub

Private Sub tID_Change()
    
    If tID.Enabled Then
        If Val(tID.Tag) <> 0 Then
            Botones True, False, False, False, False, Toolbar1, Me
            tID.Tag = 0
            bDevolver.Visible = False
        End If
    End If
    
End Sub

Private Sub tID_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tID.Text) Then
            If Trim(tID.Tag) <> Trim(tID.Text) Then CargoCambio tID.Text
        End If
    End If
    
End Sub

Private Sub tNumero_Change()
    lFactura.Tag = 0
    lFactura.Caption = ""
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(lFactura.Tag) <> 0 Then
            If cDistribuidor.Value = vbChecked Then
                If Trim(tDCompra.Text) = "" Then cDistribuidor.SetFocus: Exit Sub
            End If
            Foco cSTipo
            Exit Sub
        End If
        
        Dim mSerie As String, mNumero As String, mFecha As String, mTipoD As Integer
        Dim mID As Long
        mSerie = Trim(tSerie.Text): mNumero = Trim(tNumero.Text)
        
        mID = BuscoDocumento(Val(lCliente.Tag), Val(tArticulo.Tag), Trim(tDCompra.Text), mSerie, mNumero, mFecha, mTipoD)
        If mID <> 0 Then
            tSerie.Text = Trim(UCase(mSerie))
            tNumero.Text = Trim(mNumero)
            lFactura.Caption = " del " & Format(mFecha, "d/mmm/yyyy hh:mm")
            lFactura.Tag = mID
            tSerie.Tag = mTipoD
            Foco tNumero
        End If

    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        
        Case "grabar": AccionGrabar
        Case "eliminar":
        
        Case "cancelar": AccionCancelar
        
        Case "salir": Unload Me
    End Select
    
End Sub

Private Sub tRuc_Change()

    lCliente.Tag = 0
    lCliente.Caption = ""
    If Trim(tArticulo.Text) <> "" And tArticulo.Enabled Then tArticulo.Text = ""
    
End Sub

Private Sub tRuc_GotFocus()
    tRuc.SelStart = 0: tRuc.SelLength = (Len(tRuc.FormattedText))
    sbHelp.Panels("help").Text = "Cliente asociado a la factuara de devolución.  [C]- Cambia CI/Ruc.   [F4]- Buscar."
End Sub

Private Sub tRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF2, vbKeyP, vbKeyC
                tCi.Visible = True: tRuc.Visible = False: tCi.SetFocus: lCliente.Tag = 0: lCliente.Caption = ""
                
        Case vbKeyReturn
                If Val(lCliente.Tag) = 0 And Trim(tRuc.Text) <> "" Then
                    BuscarCliente miRuc:=Trim(tRuc.Text)
                Else
                    tArticulo.SetFocus
                End If
                
        Case vbKeyF4: BuscarClientes TipoCliente.Empresa
    End Select
    
End Sub

Private Sub tRuc_LostFocus()
    sbHelp.Panels("help").Text = ""
End Sub

Private Sub BuscarCliente(Optional miCi As String = "", Optional miRuc As String = "", Optional miId As Long = 0, Optional miTipo As Integer = 0)
    
    On Error GoTo errBCliente
    Screen.MousePointer = 11
    cDistribuidor.Value = vbUnchecked
    prmVaSucesoPorCuotas = False
    
    If miCi <> "" Then
        cons = "Select Cliente.*, (RTrim(CPeNombre1) + ' ' + RTrim(isnull(CPeNombre2, '')) + ' ' + RTrim(CPeApellido1) + ' ' + RTrim(isnull(CPeApellido2, '')))  as NombreP" & _
                   " From CPersona, Cliente" & _
                   " Where CliCodigo = CPeCliente And CliCiRuc = '" & Trim(miCi) & "'"
    End If
    
    If miRuc <> "" Then
        cons = "Select Cliente.*, (RTrim(CEmNombre) + ' (' + RTrim(isnull(CEmFantasia, '')) + ')')  as NombreE" & _
                   " From CEmpresa, Cliente" & _
                   " Where CliCodigo = CEmCliente And CliCiRuc = '" & Trim(miRuc) & "'"
    End If
    
    If miId <> 0 Then
        cons = "Select Cliente.*, " & _
                            " (RTrim(CPeNombre1) + ' ' + RTrim(isnull(CPeNombre2, '')) + ' ' + RTrim(CPeApellido1) + ' ' + RTrim(isnull(CPeApellido2, '')))  as NombreP, " & _
                            " (RTrim(CEmNombre) + ' (' + RTrim(isnull(CEmFantasia, '')) + ')')  as NombreE " & _
                    " From Cliente " & _
                            " Left Outer Join CPersona On CliCodigo = CPeCliente " & _
                            " Left Outer Join CEmpresa On CliCodigo = CEmCliente " & _
                    " Where CliCodigo = " & miId
    End If
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If rsAux!CliTipo = TipoCliente.Cliente Then
            If Not IsNull(rsAux!CliCiRuc) Then tCi.Text = Trim(rsAux!CliCiRuc)
            tRuc.Visible = False: tCi.Visible = True: tCi.SetFocus
            lCliente.Caption = " " & Trim(rsAux!NombreP)
        End If
        If rsAux!CliTipo = TipoCliente.Empresa Then
            If Not IsNull(rsAux!CliCiRuc) Then tRuc.Text = Trim(rsAux!CliCiRuc)
            tCi.Visible = False: tRuc.Visible = True: tRuc.SetFocus
            lCliente.Caption = " " & Trim(rsAux!NombreE)
        End If
        
        lCliente.Tag = rsAux!CliCodigo
        
        If Not IsNull(rsAux!CliCategoria) Then
            If InStr("," & prmCCDistribuidor & ",", "," & rsAux!CliCategoria & ",") <> 0 Then
                cDistribuidor.Value = vbChecked
            End If
        End If
    Else
        lCliente.Caption = " No Existe !!"
    End If
    rsAux.Close
    
    If Val(lCliente.Tag) <> 0 Then
        Dim mMaxAtraso As Long
        mMaxAtraso = 0
        
        'Valido Cuotas vencidas: Si atraso > 20 dias no dejo seguir.
        cons = "Select Min(CreProximoVto) " & _
                    " From Documento (index = iClienteTipo), Credito" & _
                    " Where DocCliente = " & Val(lCliente.Tag) & _
                    " And DocCodigo = CreFactura " & _
                    " And DocTipo = " & TipoDocumento.Credito & _
                    " And DocAnulado = 0 " & _
                    " And CreSaldoFactura > 0 "
        
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            If Not IsNull(rsAux(0)) Then mMaxAtraso = DateDiff("d", rsAux(0), gFechaServidor)
        End If
        
        Select Case mMaxAtraso
            Case Is > 20
                    MsgBox "El cliente seleccionado no está al día." & vbCrLf & _
                                "Tiene coutas vencidas con más de 20 días." & vbCrLf & vbCrLf & _
                                "Consulte para realizar el cambio", vbExclamation, "Cliente con Ctas. Vencidas"
                    'lCliente.Tag = 0
                    prmVaSucesoPorCuotas = True
            Case Is > 5
                    MsgBox "El cliente seleccionado no está al día. Tiene coutas vencidas." & vbCrLf & _
                                "Consulte antes de realizar el cambio de artículo.", vbExclamation, "Cliente con Ctas. Vencidas"
        End Select
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBCliente:
    clsGeneral.OcurrioError "Error al buscar el cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tSAbonado_Change()

    lSProducto.Caption = ""
    If Not prmBDRydesul Then Exit Sub
        
    tSCliente.Text = "": lSCliente.Caption = ""
    tSSerie.Text = ""
    tSFServicio.Text = ""
    
End Sub

Private Sub tSAbonado_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If cSTipo.ListIndex = -1 Then cSTipo.SetFocus: Exit Sub
        
        If Trim(tSAbonado.Text) <> "" And Val(tSAbonado.Text) <> 0 Then
            If cSTipo.ItemData(cSTipo.ListIndex) = TipoService.Rydesul Then BuscoDatosService idAbonado:=Val(tSAbonado.Text)
        End If
        
        If Trim(tSCliente.Text) = "" Then Foco tSCliente: Exit Sub
        If Trim(tSSerie.Text) = "" Then Foco tSSerie: Exit Sub
        If Trim(tSFServicio.Text) = "" Then Foco tSFServicio: Exit Sub
        
        Foco tCArticulo
        
    End If
    
End Sub

Private Sub tSCliente_Change()
     lSCliente.Caption = ""
End Sub

Private Sub tSCliente_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then Foco tSSerie
End Sub

Private Sub tSerie_Change()
    lFactura.Tag = 0
    lFactura.Caption = ""
End Sub

Private Sub tSerie_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF1
            Call tNumero_KeyPress(vbKeyReturn)
    End Select
    
End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(lFactura.Tag) = 0 Then Foco tNumero Else cDistribuidor.SetFocus
    End If
End Sub


Private Function BuscoDocumento(idCliente As Long, idArticulo As Long, fABuscar As String, _
            Optional rSerie As String = "", Optional rNumero As String = "", Optional rFecha As String = "", Optional rTipoDoc As Integer = 0) As Long

    BuscoDocumento = 0
    If idArticulo = 0 Then Exit Function
    If idCliente = 0 And Trim(rSerie) = "" And Trim(rNumero) = "" Then Exit Function
        
    On Error GoTo errBuscaG
    Screen.MousePointer = 11
    
    Dim rIDCliente As Long
    
    cons = "Select Top 20 DocCodigo, DocTipo, DocCliente, DocSerie as 'Serie', DocNumero as 'Número', DocFecha as Fecha" & _
                " From Documento Left Outer Join Renglon On DocCodigo = RenDocumento " & _
                " Where DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ")" & _
                " And DocAnulado = 0" & _
                " And RenArticulo = " & idArticulo & " And RenARetirar <> RenCantidad"
    
    If idCliente <> 0 Then cons = cons & " And DocCliente = " & idCliente
    If Trim(rSerie) <> "" Then cons = cons & " And DocSerie = '" & Trim(rSerie) & "'"
    If Trim(rNumero) <> "" Then cons = cons & " And DocNumero = " & Trim(rNumero)
    
    If IsDate(fABuscar) Then cons = cons & " And DocFecha <= '" & Format(fABuscar, "mm/dd/yyyy") & " 23:59'"
    
    'cons = cons " Validar Doc con Q = 1 y ya estan en la tabla
    cons = cons & " And Not (RenCantidad = 1 And DocCodigo In (Select CArDDocumento from CambioArticulo Where CArDArticulo = " & idArticulo & "))"
    
    cons = cons & " Order by DocFecha desc"
        
    Dim aQ As Integer, aIdSel As Long
    aQ = 0: aIdSel = 0
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        aQ = 1
        aIdSel = rsAux!DocCodigo
        rSerie = Trim(rsAux!serie): rNumero = rsAux(4): rFecha = rsAux!Fecha
        rTipoDoc = rsAux!DocTipo: rIDCliente = rsAux!DocCliente
        rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
    End If
    rsAux.Close
        
    Select Case aQ
        Case 0:
                    Dim mMsg As String
                    mMsg = "No hay datos que coincidan con los valores ingresados."
                    
                    If Trim(rSerie) <> "" And Trim(rNumero) <> "" Then
                        cons = "Select * from CambioArticulo " & _
                                    " Where CArDArticulo = " & idArticulo & _
                                    " And CArDDocumento IN ( " & _
                                            " Select DocCodigo from Documento, Renglon " & _
                                            " Where DocCodigo = RenDocumento " & _
                                            " And DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ")" & _
                                            " And DocAnulado = 0" & _
                                            " And DocSerie = '" & Trim(rSerie) & "'" & _
                                            " And DocNumero = " & Trim(rNumero) & _
                                            " And RenArticulo = " & idArticulo & " And RenARetirar <> RenCantidad )"
                        
                        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                        If Not rsAux.EOF Then
                            mMsg = mMsg & vbCrLf & "Documento y Artículo ya asignados en el Cambio ID: " & rsAux!CArID
                        End If
                        rsAux.Close
                    End If
                    
                    MsgBox mMsg, vbExclamation, "No hay datos"
        
        Case 2:
                    Dim miLista As New clsListadeAyuda
                    aIdSel = miLista.ActivarAyuda(cBase, cons, 4000, 3, "Últimas Compras del Cliente")
                    Me.Refresh
                    If aIdSel > 0 Then
                        aIdSel = miLista.RetornoDatoSeleccionado(0)
                        rTipoDoc = miLista.RetornoDatoSeleccionado(1)
                        rIDCliente = rTipoDoc = miLista.RetornoDatoSeleccionado(2)
                        rSerie = miLista.RetornoDatoSeleccionado(3)
                        rNumero = miLista.RetornoDatoSeleccionado(4)
                        rFecha = miLista.RetornoDatoSeleccionado(5)
                    End If
                    Set miLista = Nothing
    End Select
    
    If aIdSel > 0 Then
        BuscoDocumento = aIdSel
        If Val(lCliente.Tag) = 0 Then BuscarCliente miId:=rIDCliente
    End If
    Screen.MousePointer = 0
   
    Exit Function
errBuscaG:
    clsGeneral.OcurrioError "Error al buscar los datos.", Err.Description
    Screen.MousePointer = 0
End Function


Private Function BuscoDocumentoCambio()

    On Error GoTo errBuscaG
    If Val(tCArticulo.Tag) = 0 Then Foco tCArticulo: Exit Function
    If Trim(tCNumero.Text) = "" Then Foco tCNumero: Exit Function
    
    Screen.MousePointer = 11
    
    cons = "Select DocCodigo, DocSerie as 'Serie', DocNumero as 'Número', DocFecha as Fecha" & _
                " From Documento Left Outer Join Renglon On DocCodigo = RenDocumento " & _
                " Where DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ")" & _
                " And DocAnulado = 0" & _
                " And RenArticulo = " & Val(tCArticulo.Tag)
                
    If Trim(tCSerie.Text) <> "" Then cons = cons & " And DocSerie = '" & Trim(tCSerie.Text) & "'"
    If Trim(tCNumero.Text) <> "" Then cons = cons & " And DocNumero = " & Trim(tCNumero.Text)
    
    Dim aQ As Integer, aIdSel As Long, rSerie As String, rNumero As String, mTexto As String
    aQ = 0: aIdSel = 0
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        aQ = 1
        aIdSel = rsAux!DocCodigo
        rSerie = Trim(rsAux!serie): rNumero = rsAux(2)
        rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
    End If
    rsAux.Close
        
    Select Case aQ
        Case 0: MsgBox "No hay datos que coincidan con los valores ingresados.", vbExclamation, "No hay datos"
        
        Case 2:
                    Dim miLista As New clsListadeAyuda
                    aIdSel = miLista.ActivarAyuda(cBase, cons, 4000, 1, "Lista de Documentos")
                    Me.Refresh
                    If aIdSel > 0 Then
                        aIdSel = miLista.RetornoDatoSeleccionado(0)
                        
                        rSerie = miLista.RetornoDatoSeleccionado(1)
                        rNumero = miLista.RetornoDatoSeleccionado(2)
                    End If
                    Set miLista = Nothing
    End Select
    
    If aIdSel > 0 Then
        prmEntregadoEnFactura = False
        
        sbHelp.Panels("help").Text = "Procesando datos documento del cambio ...": sbHelp.Refresh
        
        tCSerie.Text = rSerie
        tCNumero.Text = rNumero
        lCFactura.Tag = aIdSel
        
        cons = "Select * from Documento, Cliente " & _
                        " left Outer Join CPersona On CliCodigo = CPeCliente " & _
                        " left Outer Join CEmpresa On CliCodigo = CEmCliente " & _
                    " Where DocCliente = CliCodigo " & _
                    " And DocCodigo = " & aIdSel
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            dCliente.Caption = rsAux!DocCliente
            tCSerie.Tag = rsAux!DocTipo
            mTexto = " del " & Format(rsAux!DocFecha, "d/mmm/yyyy hh:mm") & " "
            If Not IsNull(rsAux!CEmCliente) Then mTexto = mTexto & "(" & Trim(rsAux!CEmNombre) & ")"
            If Not IsNull(rsAux!CPeCliente) Then mTexto = mTexto & "(" & Trim(rsAux!CPeNombre1) & " " & Trim(rsAux!CPeApellido1) & ")"
        End If
        rsAux.Close
        
        lCFactura.Caption = mTexto
        
        'Valido datos a entregar (movs. stock)      ------------------------------------------------
        Dim mQTotal As Long, mQRetirar As Long
        
        cons = "Select * from Renglon " & _
                " Where RenDocumento = " & aIdSel & _
                " And RenArticulo = " & Val(tCArticulo.Tag)
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            mQTotal = rsAux!RenCantidad
            mQRetirar = rsAux!RenARetirar
        End If
        rsAux.Close
        
        Dim mQAsignado As Long
        mQAsignado = QCambiosDocumento(aIdSel, Val(tCArticulo.Tag))
        
        If mQAsignado = mQTotal Then
            MsgBox "El documento no es válido." & vbCrLf & _
                        "Ya se registraron cambios por la totalidad de los artículos.", vbExclamation, "No hay Artículos para Asignar"
            lCFactura.Tag = 0
            GoTo etFin
        End If
        
        Dim mFecha As String, mLocal As Long
        If mQRetirar <> mQTotal Then
            If mQRetirar = 0 Then prmEntregadoEnFactura = True
            
            If ValidoMovimientosStock(aIdSel, Val(tCArticulo.Tag), mFecha, mLocal) Then
                If mQTotal > 1 And mQRetirar > 0 Then
                    Dim mRet As Integer
                    mRet = MsgBox("El documento seleccionado tiene " & mQTotal & " artículos." & vbCrLf & _
                                    "Ya fueron entregados " & mQTotal - mQRetirar & " artículos." & vbCrLf & vbCrLf & _
                                    "Éste artículo fue entregado ?.", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Artículo Entregado ?")
                    
                    Select Case mRet
                        Case vbCancel: lCFactura.Tag = 0: GoTo etFin
                        Case vbYes:  prmEntregadoEnFactura = True
                        Case vbNo: mFecha = "": mLocal = 0 'paCodigoDeSucursal
                    End Select
                End If
                If Trim(tFEntrega.Tag) = "" Then
                    tFEntrega.Text = mFecha
                    BuscoCodigoEnCombo cLocal, mLocal
                End If
                
            End If
            
        End If
        
        If prmEntregadoEnFactura Then
            bEntregar.Enabled = False
            'Si esta entregado en factura y no hay movs de stock pongo como f entrega ahora
            tFEntrega.Text = Format(gFechaServidor, "dd/mm/yyyy hh:mm")
        Else
            If Val(lCFactura.Tag) <> 0 Then bEntregar.Enabled = True
        End If
        
        

etFin:
        sbHelp.Panels("help").Text = ""
    End If
    Screen.MousePointer = 0
   
    Exit Function
errBuscaG:
    clsGeneral.OcurrioError "Error al buscar los datos.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function BuscoArticulo(mControl As TextBox)

    On Error GoTo errBuscaG
    Screen.MousePointer = 11
    
    cons = "Select ArtID, ArtCodigo as Codigo, ArtNombre  as Nombre from Articulo "
    If IsNumeric(mControl.Text) Then
        cons = cons & " Where ArtCodigo = " & Val(mControl.Text)
    Else
        cons = cons & "Where ArtNombre like '" & Replace(Trim(mControl.Text), " ", "%") & "%'"
    End If
    cons = cons & " Order by ArtNombre"
    
    Dim aQ As Integer, aIdArticulo As Long, aTexto As String
    aQ = 0: aIdArticulo = 0
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        aQ = 1
        aIdArticulo = rsAux!ArtID: aTexto = Format(rsAux!Codigo, "(#,000,000)") & " " & Trim(rsAux!Nombre)
        rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
    End If
    rsAux.Close
        
    Select Case aQ
        Case 0: MsgBox "No hay datos que coincidan con el texto ingersado.", vbExclamation, "No hay datos"
        
        Case 2:
                    Dim miLista As New clsListadeAyuda
                    aIdArticulo = miLista.ActivarAyuda(cBase, cons, 4000, 1, "Lista de Articulos")
                    Me.Refresh
                    If aIdArticulo > 0 Then
                        aIdArticulo = miLista.RetornoDatoSeleccionado(0)
                        
                        aTexto = Format(miLista.RetornoDatoSeleccionado(1), "(#,000,000)") & " "
                        aTexto = aTexto & miLista.RetornoDatoSeleccionado(2)
                    End If
                    Set miLista = Nothing
    End Select
        
    If aIdArticulo > 0 Then
        mControl.Text = aTexto
        mControl.Tag = aIdArticulo
        Foco mControl
    End If
    
    Screen.MousePointer = 0
   
    Exit Function
errBuscaG:
    clsGeneral.OcurrioError "Error al buscar los datos.", Err.Description
    Screen.MousePointer = 0
End Function


Private Function BuscoDatosService(Optional idAbonado As Long = 0) As Boolean
    
    BuscoDatosService = True
    If Not prmBDRydesul Then Exit Function
    On Error GoTo errRydesul
    
    lSProducto.Caption = ""
    tSCliente.Text = "": lSCliente.Caption = ""
    tSSerie.Text = ""
    tSFServicio.Text = ""
    
    Screen.MousePointer = 11
    sbHelp.Panels("help").Text = "Buscando datos service.  Espere ..."
    Me.Refresh
    
    cons = "Select CliNombre, Producto.*, TipNombre, STiNombre, ModNombre " & _
            " From Producto, Modelo, SubTipo, Tipo, Cliente" & _
            " Where ProCodigo = " & idAbonado & _
            " And ProCliente = CliCodigo " & _
            " And ProModelo = ModCodigo And ModSubTipo = STiCodigo And STiTipo = TipCodigo"
    
    Set rsAux = cBaseRD.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        tSAbonado.Text = rsAux!ProCodigo
        lSProducto.Caption = " " & StrConv(Trim(rsAux!TipNombre) & " " & Trim(rsAux!STiNombre) & " " & Trim(rsAux!ModNombre), vbProperCase)
        
        tSCliente.Text = rsAux!ProCliente
        lSCliente.Caption = " " & Trim(rsAux!CliNombre)
        
        If Not IsNull(rsAux!ProNroSerie) Then tSSerie.Text = Trim(rsAux!ProNroSerie)
    Else
        BuscoDatosService = False
    End If
    rsAux.Close
    
    If BuscoDatosService Then
        cons = "Select Max(HisFechaHora) from Historia Where HisProducto = " & idAbonado
        Set rsAux = cBaseRD.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            If Not IsNull(rsAux(0)) Then tSFServicio.Text = Format(rsAux(0), "dd/mm/yyyy hh:mm")
        End If
        rsAux.Close
    End If
    
    sbHelp.Panels("help").Text = ""
    Screen.MousePointer = 0
    Exit Function

errRydesul:
    clsGeneral.OcurrioError "Error al buscar los datos del service.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function BuscoDatosServiceCGSA(idServicio As Long) As Boolean
    
    BuscoDatosServiceCGSA = True
    On Error GoTo errBService
    Dim mIdCliente As Long: mIdCliente = 0
    
    BuscoCodigoEnCombo cSTipo, TipoService.CGSA
    lSProducto.Caption = ""
    tSCliente.Text = "": lSCliente.Caption = ""
    'tSSerie.Text = ""
    'tSFServicio.Text = ""
    
    Screen.MousePointer = 11
    sbHelp.Panels("help").Text = "Buscando datos service.  Espere ..."
    Me.Refresh
    
    cons = "Select SerCodigo, SerFecha, ArtID, ArtCodigo, ArtNombre, Producto.*" & _
                " From Servicio, Producto, Articulo " & _
                " Where SerCodigo = " & idServicio & _
                " And SerProducto = ProCodigo And ProArticulo = ArtID"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        tSIDServicio.Text = rsAux!SerCodigo
        mIdCliente = rsAux!ProCliente
              
        tSAbonado.Text = rsAux!ProCodigo
        lSProducto.Caption = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
        
        If Not IsNull(rsAux!ProNroSerie) Then tSSerie.Text = Trim(rsAux!ProNroSerie)
        If Not IsNull(rsAux!SerFecha) Then tSFServicio.Text = Format(rsAux!SerFecha, "dd/mm/yyyy hh:mm")
    End If
    rsAux.Close
    
    If mIdCliente <> 0 Then
        cons = "Select Cliente.*, " & _
                            " (RTrim(CPeNombre1) + ' ' + RTrim(isnull(CPeNombre2, '')) + ' ' + RTrim(CPeApellido1) + ' ' + RTrim(isnull(CPeApellido2, '')))  as NombreP, " & _
                            " (RTrim(CEmNombre) + ' (' + RTrim(isnull(CEmFantasia, '')) + ')')  as NombreE " & _
                    " From Cliente " & _
                            " Left Outer Join CPersona On CliCodigo = CPeCliente " & _
                            " Left Outer Join CEmpresa On CliCodigo = CEmCliente " & _
                    " Where CliCodigo = " & mIdCliente
        
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            tSCliente.Text = rsAux!CliCodigo
            
            If rsAux!CliTipo = TipoCliente.Cliente Then
                lSCliente.Caption = " " & Trim(rsAux!NombreP)
            End If
            If rsAux!CliTipo = TipoCliente.Empresa Then
                lSCliente.Caption = " " & Trim(rsAux!NombreE)
            End If
            
            lCliente.Tag = rsAux!CliCodigo
        Else
            lSCliente.Caption = " No Existe !!"
        End If
        
        rsAux.Close
    End If
    
    If mIdCliente = 0 Then
        MsgBox "No existe un código de servicio para " & Trim(cSTipo.Text), vbInformation, "No hay datos del Servicio"
        BuscoDatosServiceCGSA = False
    End If
    
    sbHelp.Panels("help").Text = ""
    Screen.MousePointer = 0
    Exit Function

errBService:
    clsGeneral.OcurrioError "Error al buscar los datos del service.", Err.Description
    BuscoDatosServiceCGSA = False
    Screen.MousePointer = 0
End Function
Private Sub tSFServicio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tSFServicio.Text) <> "" Then
            If IsDate(tSFServicio.Text) Then
                tSFServicio.Text = Format(tSFServicio.Text, "dd/mm/yyyy hh:mm")
            Else
                Exit Sub
            End If
        End If
        Foco tCArticulo
    End If
    
End Sub


Private Sub tSIDServicio_Change()
    If Trim(tSAbonado.Text) <> "" Then tSAbonado.Text = ""
End Sub

Private Sub tSIDServicio_GotFocus()
    tSIDServicio.SelStart = 0: tSIDServicio.SelLength = Len(tSIDServicio.Text)
End Sub

Private Sub tSIDServicio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        On Error Resume Next
        If cSTipo.ListIndex = -1 Then Foco cSTipo: Exit Sub
        
        Select Case cSTipo.ItemData(cSTipo.ListIndex)
            Case TipoService.CGSA
                        If Trim(tSIDServicio.Text) = "" Then Foco tCArticulo: Exit Sub
                        If BuscoDatosServiceCGSA(Val(tSIDServicio.Text)) Then Foco tCArticulo
                
            Case TipoService.Rydesul: Foco tSAbonado
        End Select
        
    End If
    
End Sub

Private Sub tSSerie_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
        If prmModificarSinMovs Then tUsuario.SetFocus: Exit Sub
        Foco tSFServicio
    End If
End Sub


Private Sub HabilitoCampos(Optional sNuevo As Boolean = False, Optional sModificar As Boolean = False)

Dim bState As Boolean
Dim bkColor As Long, bkColorNot As Long
    
    If sNuevo Or sModificar Then
        bDevolver.Visible = False
    End If
    
    If Not sNuevo And Not sModificar Then
        bState = False
        bkColor = Colores.Inactivo
        bkColorNot = vbWindowBackground
    Else
        bState = True
        bkColor = vbWindowBackground
        bkColorNot = Colores.Inactivo
    End If
    
    tID.Enabled = Not bState: tID.BackColor = bkColorNot
    
    tSSerie.Enabled = bState: tSSerie.BackColor = bkColor
    tUsuario.Enabled = bState: tUsuario.BackColor = bkColor
    
    If prmModificarSinMovs Then Exit Sub
    
    tCSerie.Enabled = bState: tCSerie.BackColor = bkColor
    tCNumero.Enabled = bState: tCNumero.BackColor = bkColor
    
    cLocal.Enabled = bState: cLocal.BackColor = bkColor
    cEntregaA.Enabled = bState: cEntregaA.BackColor = bkColor
    cLocalD.Enabled = bState: cLocalD.BackColor = bkColor
     
    'Modificar --> se habilita para ingresar factura y datos de entrega.
    If sModificar Then
        
        bEntregar.Enabled = (Trim(tFEntrega.Tag) = "")
        If Val(lCFactura.Tag) <> 0 Then     'No dejo cambiar la factura
            tCNumero.Enabled = Not bState: tCNumero.BackColor = bkColorNot
            tCSerie.Enabled = Not bState: tCSerie.BackColor = bkColorNot
        End If
        
        If Trim(tFEntrega.Text) <> "" Then
            cLocal.Enabled = Not bState: cLocal.BackColor = bkColorNot
            cEntregaA.Enabled = Not bState: cEntregaA.BackColor = bkColorNot
            
            cLocalD.Enabled = Not bState: cLocalD.BackColor = bkColorNot
        Else
            'Cambio pedido x carlos. 13-9-02
            If Val(lCFactura.Tag) = 0 Then tCArticulo.Enabled = bState: tCArticulo.BackColor = bkColor
        End If
        
        Exit Sub
    Else
        bEntregar.Enabled = bState
    End If
    
    tCi.Enabled = bState: tCi.BackColor = bkColor
    tRuc.Enabled = bState: tRuc.BackColor = bkColor
    tArticulo.Enabled = bState: tArticulo.BackColor = bkColor
    cLocalD.Enabled = bState: cLocalD.BackColor = bkColor
    cLocalD.Font.Bold = False: cLocalD.ForeColor = vbWindowText
    
    tSerie.Enabled = bState: tSerie.BackColor = bkColor
    tNumero.Enabled = bState: tNumero.BackColor = bkColor
    cDistribuidor.Value = vbUnchecked
    cDistribuidor.Enabled = bState
    
    cSTipo.Enabled = bState: cSTipo.BackColor = bkColor
    tSIDServicio.Enabled = bState: tSIDServicio.BackColor = bkColor
    tSAbonado.Enabled = bState: tSAbonado.BackColor = bkColor
    tSCliente.Enabled = bState: tSCliente.BackColor = bkColor
    tSFServicio.Enabled = bState: tSFServicio.BackColor = bkColor
    
    
    tCArticulo.Enabled = bState: tCArticulo.BackColor = bkColor
       
    
    If bState Then Exit Sub
    tFEntrega.Enabled = bState: tFEntrega.BackColor = bkColor

'    cEntregaA.Enabled = bState: cEntregaA.BackColor = bkColor
            
End Sub

Private Sub AccionNuevo()
    On Error Resume Next
    prmModificarSinMovs = False
    
    tID.Text = ""
    LimpioFicha
    HabilitoCampos sNuevo:=True
    tCi.SetFocus
    
    Botones False, False, False, True, True, Toolbar1, Me
'    BuscoCodigoEnCombo cLocalD, paCodigoDeSucursal
    
End Sub

Private Sub AccionModificar()

    prmModificarSinMovs = False
    If Val(lCFactura.Tag) <> 0 And Trim(tFEntrega.Text) <> "" Then
        MsgBox "En éste registro no se pueden modificar los datos del cambio." & vbCrLf & _
                    "Tiene asignada una factura y ya fue entregado.", vbInformation, "Cambio Entregado y Asignado"
                    
        prmModificarSinMovs = True
        
    End If
    
    HabilitoCampos sModificar:=True
    If Not prmModificarSinMovs Then Foco tCSerie Else Foco tSSerie
    
    Botones False, False, False, True, True, Toolbar1, Me
    
End Sub

Private Sub AccionCancelar()
On Error Resume Next
    Screen.MousePointer = 11
    
    Dim mOldID As Long
    mOldID = 0
    If Trim(tID.Text) <> "" Then mOldID = Val(tID.Text)
    
    HabilitoCampos
    LimpioFicha
    
    Botones True, False, False, False, False, Toolbar1, Me
    CargoCambio mOldID
    
    tID.SetFocus
    
    Screen.MousePointer = 0
    
End Sub

Private Function CargoCambio(mID As Long)
    Screen.MousePointer = 11
Dim mIdCliente As Long, mIDDevolucion As Long, mIdCambio As Long
    
    LimpioFicha
    prmIdTraslado = 0
    
    cons = "Select CambioArticulo.*, ArticuloD.ArtCodigo DArtCodigo, ArticuloD.ArtNombre DArtNombre,  " & _
                        " ArticuloC.ArtCodigo CArtCodigo, ArticuloC.ArtNombre CArtNombre " & _
               " From CambioArticulo " & _
                    " Left Outer Join Articulo ArticuloD On ArticuloD.ArtID = CArDArticulo " & _
                    " Left Outer Join Articulo ArticuloC On ArticuloC.ArtID = CArCArticulo " & _
                " Where CArID = " & mID
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    mID = 0: mIdCambio = 0: mIDDevolucion = 0
    If Not rsAux.EOF Then
        mID = rsAux!CArID
        
        'Datos del la Devolución  --------------------------------------------------------------
        mIdCliente = rsAux!CArDCliente
        mIDDevolucion = rsAux!CArDDocumento
        If Not IsNull(rsAux!CArCIDTraslado) Then prmIdTraslado = rsAux!CArCIDTraslado
        
        tArticulo.Text = Format(rsAux!DArtCodigo, "(#,000,000)") & " " & Trim(rsAux!DArtNombre)
        tArticulo.Tag = rsAux!CArDArticulo
        
        If Not IsNull(rsAux!CArDLocal) Then BuscoCodigoEnCombo cLocalD, rsAux!CArDLocal
        
        If Not IsNull(rsAux!CArDFCompra) Then tDCompra.Text = Format(rsAux!CArDFCompra, "dd/mm/yyyy")
        If Not IsNull(rsAux!CArDBoleta) Then tDBoleta.Text = Trim(rsAux!CArDBoleta)
        If Not IsNull(rsAux!CArDNombre) Then tDCliente.Text = Trim(rsAux!CArDNombre)
        If Trim(tDCompra.Text) <> "" Or Trim(tDBoleta.Text) <> "" Or Trim(tDCliente.Text) <> "" Then
            cDistribuidor.Value = vbChecked
        End If
        
        If Not IsNull(rsAux!CArSTipoService) Then BuscoCodigoEnCombo cSTipo, rsAux!CArSTipoService
        If Not IsNull(rsAux!CArSIDService) Then tSIDServicio.Text = rsAux!CArSIDService
        
        'Datos del Service          --------------------------------------------------------------
        If Not IsNull(rsAux!CArSIdAbonado) Then tSAbonado.Text = rsAux!CArSIdAbonado
        If Not IsNull(rsAux!CArSIdCliente) Then tSCliente.Text = rsAux!CArSIdCliente
        If Not IsNull(rsAux!CArSFecha) Then tSFServicio.Text = Format(rsAux!CArSFecha, "dd/mm/yyyy hh:mm")
        If Not IsNull(rsAux!CArSSerie) Then tSSerie.Text = Trim(rsAux!CArSSerie)
    
        'Datos del la Factura del Cambio    ---------------------------------------------------
        tCArticulo.Text = Format(rsAux!CArtCodigo, "(#,000,000)") & " " & Trim(rsAux!CArtNombre)
        tCArticulo.Tag = rsAux!CArCArticulo
        
        If Not IsNull(rsAux!CArCFEntrega) Then tFEntrega.Text = Format(rsAux!CArCFEntrega, "dd/mm/yyyy hh:mm")
        tFEntrega.Tag = tFEntrega.Text
        
        If Not IsNull(rsAux!CArCDocumento) Then mIdCambio = rsAux!CArCDocumento
        lTFactura.Tag = mIdCambio
        
        If Not IsNull(rsAux!CArCLocal) Then BuscoCodigoEnCombo cLocal, rsAux!CArCLocal
        If Not IsNull(rsAux!CArCAQuien) Then BuscoCodigoEnCombo cEntregaA, rsAux!CArCAQuien
        If Not IsNull(rsAux!CArCUsuario) Then tUsuario.UserID = rsAux!CArCUsuario
        
    End If
    rsAux.Close
    
    If mID <> 0 Then
        'Cargo datos del Cliente Devolcuion ------------------------------------------------------
        sbHelp.Panels("help").Text = "Cargando datos cliente ...": sbHelp.Refresh
        
        cons = "Select Cliente.*, CPeCliente, " & _
                            " (RTrim(isnull(CEmNombre, '')) + ' (' + RTrim(isnull(CEmFantasia, '')) + ')')  as NombreE, " & _
                            " (RTrim(isnull(CPeNombre1, '')) + ' ' + RTrim(isnull(CPeNombre2, '')) + ' ' + RTrim(isnull(CPeApellido1, '')) + ' ' + RTrim(isnull(CPeApellido2, '')))  as NombreP" & _
                    " From Cliente " & _
                        " Left Outer Join CPersona On CliCodigo = CPeCliente " & _
                        " Left Outer Join CEmpresa On CliCodigo = CEmCliente " & _
                    " Where CliCodigo = " & mIdCliente
                    
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            If Not IsNull(rsAux!CPeCliente) Then
                If Not IsNull(rsAux!CliCiRuc) Then tCi.Text = Trim(rsAux!CliCiRuc)
                tCi.ZOrder 0
                lCliente.Caption = Trim(rsAux!NombreP)
            Else
                If Not IsNull(rsAux!CliCiRuc) Then tRuc.Text = Trim(rsAux!CliCiRuc)
                tRuc.ZOrder 0
                lCliente.Caption = Trim(rsAux!NombreE)
            End If
            lCliente.Tag = mIdCliente
        End If
        rsAux.Close
    
        'Cargo datos del documento Devolucion ------------------------------------------------------
        sbHelp.Panels("help").Text = "Cargando datos documentos ...": sbHelp.Refresh
        
        cons = "Select * from Documento Where DocCodigo = " & mIDDevolucion
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            tSerie.Text = Trim(UCase(rsAux!DocSerie))
            tNumero.Text = Trim(rsAux!DocNumero)
            lFactura.Caption = " del " & Format(rsAux!DocFecha, "d/mmm/yyyy hh:mm")
            lFactura.Tag = mIDDevolucion
            tSerie.Tag = rsAux!DocTipo
        End If
        rsAux.Close
                
        'Cargo datos del documento del Cambio ------------------------------------------------------
        If mIdCambio <> 0 Then
            Dim mTexto As String
            cons = "Select * from Documento, Cliente " & _
                            " left Outer Join CPersona On CliCodigo = CPeCliente " & _
                            " left Outer Join CEmpresa On CliCodigo = CEmCliente " & _
                        " Where DocCliente = CliCodigo " & _
                        " And DocCodigo = " & mIdCambio
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then
                mTexto = " del " & Format(rsAux!DocFecha, "d/mmm/yyyy hh:mm") & " "
                If Not IsNull(rsAux!CEmCliente) Then mTexto = mTexto & "(" & Trim(rsAux!CEmNombre) & ")"
                If Not IsNull(rsAux!CPeCliente) Then mTexto = mTexto & "(" & Trim(rsAux!CPeNombre1) & " " & Trim(rsAux!CPeApellido1) & ")"
                
                tCSerie.Text = Trim(rsAux!DocSerie)
                tCNumero.Text = Trim(rsAux!DocNumero)
                
                lCFactura.Caption = mTexto
                lCFactura.Tag = mIdCambio
            End If
            rsAux.Close
        End If
        
        If cSTipo.ListIndex <> -1 Then
            If cSTipo.ItemData(cSTipo.ListIndex) = TipoService.CGSA And Val(tSIDServicio.Text) <> 0 Then BuscoDatosServiceCGSA Val(tSIDServicio.Text)
        End If
        
        sbHelp.Panels("help").Text = ""
        Botones True, True, True, False, False, Toolbar1, Me
        tID.Tag = mID
    End If
    
    bDevolver.Visible = SePuedeDevolver
    Screen.MousePointer = 0

    Exit Function

errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del registro.", Err.Description
    Screen.MousePointer = 0
    Exit Function
End Function

Private Function QCambiosDocumento(mDocumento As Long, mArticulo As Long) As Long

    sbHelp.Panels("help").Text = "Validando asignaciones del documento ...": sbHelp.Refresh
    QCambiosDocumento = 0
    
    cons = "Select Count(*) from CambioArticulo " & _
            " Where CArCDocumento = " & mDocumento & _
            " And CArCArticulo = " & mArticulo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsAux.EOF Then
        If Not IsNull(rsAux(0)) Then QCambiosDocumento = rsAux(0)
    End If
    rsAux.Close
    
    sbHelp.Panels("help").Text = "": sbHelp.Refresh
    
End Function

Private Function ValidoMovimientosStock(mDocumento As Long, mArticulo As Long, retFecha As String, retLocal As Long) As Boolean

    sbHelp.Panels("help").Text = "Validando movimientos de stock ...": sbHelp.Refresh
    ValidoMovimientosStock = False
    
    cons = "Select Top 1 * from MovimientoStockFisico " & _
            " Where MSFDocumento = " & mDocumento & _
            " And MSFTipoDocumento IN (1, 2) " & _
            " And MSFArticulo = " & mArticulo & _
            " And MSFTipoLocal = 2" & _
            " Order by MSFFecha Desc"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsAux.EOF Then
        ValidoMovimientosStock = True
        retFecha = Format(rsAux!MSFFecha, "dd/mm/yyyy hh:mm")
        retLocal = rsAux!MSFLocal
    End If
    rsAux.Close
    
    sbHelp.Panels("help").Text = "": sbHelp.Refresh

End Function

Private Sub tUsuario_AfterDigit()
    If tUsuario.Enabled Then AccionGrabar
End Sub


Private Function GraboComentarios(idCambio As Long)

Dim rsCom As rdoResultset
Dim mTxt As String, mCom As String

    mTxt = "ID:" & idCambio & " "
    
    cons = "Select * from Comentario Where ComCodigo = 0"
    Set rsCom = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    rsCom.AddNew
    rsCom!ComCliente = Val(lCliente.Tag)
    rsCom!ComFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    
    mCom = mTxt & "Cambio de Artículos en Garantía (doc de devolución)"
    mCom = mCom & vbCrLf & Trim(tArticulo.Text) & " x " & Trim(tCArticulo.Text)
    If Val(lCFactura.Tag) <> 0 Then mCom = mCom & vbCrLf & "Doc. cambio: " & Trim(tCSerie.Text) & " " & Trim(tCNumero.Text)
    rsCom!ComComentario = mCom
    
    rsCom!ComTipo = prmTCCAmbio
    rsCom!ComUsuario = tUsuario.UserID
    rsCom!ComDocumento = Val(lFactura.Tag)
    rsCom.Update
    
    rsCom.AddNew
    rsCom!ComCliente = Val(dCliente.Caption)
    rsCom!ComFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    
    mCom = mTxt & "Cambio de Artículos en Garantía (doc del cambio)"
    mCom = mCom & vbCrLf & Trim(tArticulo.Text) & " x " & Trim(tCArticulo.Text)
    If Val(lCFactura.Tag) <> 0 Then mCom = mCom & vbCrLf & "Doc. dev: " & Trim(tSerie.Text) & " " & Trim(tNumero.Text)
    rsCom!ComComentario = mCom
    
    rsCom!ComTipo = prmTCCAmbio
    rsCom!ComUsuario = tUsuario.UserID
    rsCom!ComDocumento = Val(lCFactura.Tag)
    rsCom.Update
    
    rsCom.Close

End Function

Private Sub AccionMenuHelp()
    On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    cons = "Select * from Aplicacion Where AplNombre = 'Cambio en Garantia'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux!AplHelp) Then aFile = Trim(rsAux!AplHelp)
    rsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function EnviarMensaje(mIdCambio As Long, mArticulo As String) As Boolean
    On Error GoTo errMsg
    
    If prmMenUsuarioCambioArticulo = "" Then Exit Function
    
    Dim aMsg As String
    aMsg = "Facturar Cambio de Artículo en Garantía" & vbCrLf & vbCrLf & _
                "Código de Cambio: " & mIdCambio & vbCrLf & _
                "Cliente Devolución: " & lCliente.Caption & vbCrLf & _
                "Artículo que Devuelve: " & tArticulo.Text & vbCrLf & _
                "Factura Devolución: " & Trim(tSerie.Text) & "-" & Trim(tNumero.Text) & vbCrLf & _
                "Local de Ingreso: " & cLocalD.Text & vbCrLf & vbCrLf & _
                "Artículo del Cambio: " & mArticulo & vbCrLf & _
                "Usuario que Entrega: " & Trim(tUsuario.UserName)
                
    miConexion.EnviaMensaje prmMenUsuarioCambioArticulo, "Facturar Cambio de Artículo", aMsg, _
                                        DateAdd("s", 10, gFechaServidor), 0, prmMenUsuarioSistema
        
    EnviarMensaje = True
    Exit Function

errMsg:
    clsGeneral.OcurrioError "Error al enviar mensaje automático.", Err.Description
End Function

Private Sub AccionEliminar()

    If Val(tID.Tag) = 0 Then Exit Sub
    
    If Trim(tFEntrega.Text) = "" Then
        MsgBox "El cambio no está entregado, no se puede eliminar" & vbCrLf & vbCrLf & _
                    "Ésta acción permite eliminar el traslado al local 'A Quien'.", vbInformation, "Cambio No Entregado"
        Exit Sub
    End If
    
    If Val(lCFactura.Tag) <> 0 Then
        MsgBox "En éste registro no se puede eliminar el traslado al local 'A Quien'." & vbCrLf & _
                    "Ya tiene asignada una factura.", vbInformation, "Cambio Asignado"
                    
        Exit Sub
    End If
    
    If prmIdTraslado = 0 Then
        MsgBox "El registro no se puede eliminar." & vbCrLf & _
                    "El local 'A Quien' no es un local intermediario o  no hay un traslado de mercadería al local.", vbInformation, "No se puede Eliminar"
        Exit Sub
    End If
    
    If MsgBox("Confirma eliminar el traslado al local A Quien y quitar la fecha de entrega.", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Traslado") = vbYes Then
        EliminoEntregaAQuien
    End If
    
End Sub

Private Sub EliminoEntregaAQuien()

    On Error GoTo errorBT
    Screen.MousePointer = 11
    prmTipoLocalAQuien = 0
    Dim mID As Long
    mID = Val(tID.Text)
    
    'Cargo Tipo y Local a Quien
    cons = "Select isNull(AQEConFactura, 0) as AQEConFactura, isNull(AQELocal, 0) as AQELocal, LocTipo " & _
                " From AQuienEntrega " & _
                        " Left Outer Join Local On AQELocal = LocCodigo" & _
               " Where AQECodigo = " & cEntregaA.ItemData(cEntregaA.ListIndex)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        prmLocalAQuien = rsAux!AQELocal
        If Not IsNull(rsAux!LocTipo) Then prmTipoLocalAQuien = rsAux!LocTipo Else prmTipoLocalAQuien = 0
    End If
    rsAux.Close
    
    If prmTipoLocalAQuien = 0 Then
        MsgBox "Error al validar el tipo de local 'A Quien'", vbExclamation, "Error al Validar Local"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    FechaDelServidor
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    Dim mIDTraslado As Long
    mIDTraslado = MuevoTrasladoMercaderia(mID, prmLocalAQuien, cLocal.ItemData(cLocal.ListIndex))
    
    MuevoStockLocalALocal TipoDocumento.Traslados, mIDTraslado, Val(tCArticulo.Tag), _
                     prmTipoLocalAQuien, prmLocalAQuien, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex)
    
    'Updateo con ID de Traslado
    cons = "Select * from CambioArticulo Where CArID = " & mID
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsAux.Edit
    rsAux!CArCIDTraslado = Null
    rsAux!CArCFEntrega = Null
    rsAux.Update: rsAux.Close
    

    MuevoStockIngreso Val(tSerie.Tag), Val(lFactura.Tag), Val(tArticulo.Tag), cLocalD.ItemData(cLocalD.ListIndex), mEsAlta:=-1
    
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------
   
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub

errValidar:
    clsGeneral.OcurrioError "Error al procesar los datos para grabar.", Err.Description
    Screen.MousePointer = 0: Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al grabar los datos del cambio.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub

Private Function SePuedeDevolver() As Boolean
    
    'Tiene que haber F/Entrega, No haber factura y haber traslado.
    
    SePuedeDevolver = False
    
    If Trim(tFEntrega.Text) = "" Then Exit Function
    If Val(lCFactura.Tag) <> 0 Then Exit Function
    
    If prmIdTraslado = 0 Then Exit Function
    
    SePuedeDevolver = True

End Function
