VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5694326E-AE1E-40BF-B7B0-0E8918015F0D}#1.1#0"; "orChequeCtrl.ocx"
Begin VB.Form frmFacturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos"
   ClientHeight    =   5115
   ClientLeft      =   2685
   ClientTop       =   3285
   ClientWidth     =   8085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8085
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
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
            ImageIndex      =   6
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
            Style           =   4
            Object.Width           =   400
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "pagos"
            Object.ToolTipText     =   "Con Qué paga"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "plazos"
            Object.ToolTipText     =   "Vencimientos"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "dolar"
            Object.ToolTipText     =   "Tipos de cambio"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nota"
            Object.ToolTipText     =   "Asignar nota"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   3300
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Comprobante"
      ForeColor       =   &H00000080&
      Height          =   2055
      Left            =   60
      TabIndex        =   34
      Top             =   480
      Width           =   7935
      Begin VB.TextBox tCofis 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5580
         MaxLength       =   13
         TabIndex        =   17
         Text            =   "9999.99"
         Top             =   1335
         Width           =   735
      End
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3960
         MaxLength       =   40
         TabIndex        =   5
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox tID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox tTCDolar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   19
         Text            =   "999.99"
         Top             =   1695
         Width           =   795
      End
      Begin VB.TextBox tIva 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         MaxLength       =   13
         TabIndex        =   15
         Text            =   "9999.99"
         Top             =   1335
         Width           =   915
      End
      Begin VB.TextBox tIOriginal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2040
         MaxLength       =   13
         TabIndex        =   13
         Text            =   "1,000,000.00"
         Top             =   1320
         Width           =   1035
      End
      Begin VB.TextBox tNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         MaxLength       =   9
         TabIndex        =   10
         Top             =   980
         Width           =   975
      End
      Begin VB.TextBox tSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   9
         Top             =   980
         Width           =   435
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   1320
         Width           =   795
         _ExtentX        =   1402
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
      Begin AACombo99.AACombo cComprobante 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   1875
         _ExtentX        =   3307
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
      Begin MSComCtl2.DTPicker dFecha 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   90374145
         CurrentDate     =   37543
      End
      Begin VB.Label lModificado 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Mod. 00/00/00 00:00 x UsuarioXXXX"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2520
         TabIndex        =   39
         Top             =   300
         Width           =   3135
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Co&fis:"
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Top             =   1365
         Width           =   495
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Total NETO:"
         Height          =   255
         Left            =   5760
         TabIndex        =   38
         Top             =   285
         Width           =   975
      End
      Begin VB.Label lTotalGasto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,000,000.00"
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
         Height          =   285
         Left            =   6660
         TabIndex        =   37
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "I&d Compra:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   280
         Width           =   855
      End
      Begin VB.Label lTC 
         BackStyle       =   0  'Transparent
         Caption         =   "21/08/2000"
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
         Left            =   2040
         TabIndex        =   36
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "T/&C Dólar:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1725
         Width           =   1035
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Total BRUTO:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1365
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&I.V.A.:"
         Height          =   255
         Left            =   3420
         TabIndex        =   14
         Top             =   1365
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Compro&bante:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   2940
         TabIndex        =   4
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha &Gasto:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   615
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Número:"
         Height          =   255
         Left            =   3300
         TabIndex        =   8
         Top             =   1020
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Distribución del Gasto"
      ForeColor       =   &H00000080&
      Height          =   2160
      Left            =   60
      TabIndex        =   33
      Top             =   2640
      Width           =   7935
      Begin orChequeCtrl.orCheque orCheque 
         Height          =   315
         Left            =   3900
         TabIndex        =   26
         Top             =   1020
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   556
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
      Begin VB.CommandButton bSplitR 
         Caption         =   "&Varios Rubros"
         Height          =   325
         Left            =   4740
         TabIndex        =   40
         Top             =   220
         Width           =   1155
      End
      Begin VB.TextBox tRubro 
         Appearance      =   0  'Flat
         Height          =   305
         Left            =   1140
         TabIndex        =   21
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox tSubRubro 
         Appearance      =   0  'Flat
         Height          =   305
         Left            =   1140
         TabIndex        =   23
         Top             =   600
         Width           =   3495
      End
      Begin VB.CheckBox chVerificado 
         Appearance      =   0  'Flat
         Caption         =   "Autoriza&do"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   31
         Top             =   1785
         Width           =   1395
      End
      Begin VB.TextBox tAutoriza 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   80
         TabIndex        =   30
         Top             =   1740
         Width           =   2055
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   80
         TabIndex        =   28
         Top             =   1440
         Width           =   6675
      End
      Begin AACombo99.AACombo cDisponibilidad 
         Height          =   315
         Left            =   1140
         TabIndex        =   25
         Top             =   1020
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         ForeColor       =   0
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
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Se P&aga Con:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lAutoriza 
         BackStyle       =   0  'Transparent
         Caption         =   "Autori&za:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1780
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Rubro:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario&s:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Subrubro:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   4860
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
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
            Object.Width           =   6509
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":0BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":0EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":0FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":12EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":1604
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":191E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":1C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":1F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":226C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":2446
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturas.frx":2760
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
      Begin VB.Menu MnuL1 
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
   Begin VB.Menu MnuAcciones 
      Caption         =   "&Acciones"
      Begin VB.Menu mnuDocsRel 
         Caption         =   "Documentos Relacionados"
      End
      Begin VB.Menu MnuAConQuePaga 
         Caption         =   "Con Que &Paga"
      End
      Begin VB.Menu MnuAVencimientos 
         Caption         =   "Ingresar &Vencimientos"
      End
      Begin VB.Menu MnuAcL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuATipoCambio 
         Caption         =   "Tipos de &Cambio"
      End
      Begin VB.Menu MnuAAsignarNota 
         Caption         =   "Asignar &Nota"
      End
      Begin VB.Menu MnuAcL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAcProveedores 
         Caption         =   "Ingresar Nuevos Proveedores"
         Shortcut        =   {F3}
      End
      Begin VB.Menu MnuAcVOpe 
         Caption         =   "Visualización de Operaciones"
         Shortcut        =   {F12}
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
      Begin VB.Menu MnuExit 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "&?"
      Begin VB.Menu MnuHelp 
         Caption         =   "Ayuda ..."
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean
Dim prmIdCompra As Long
Dim prmFModificacion As Date

Dim bIngresarVencimientos As Boolean        'Señal para ingresar los vencimientos

Dim rsCom As rdoResultset
Dim aTiposDocs As String

Private Type typReg
    oFechaCompra As Date
    oIDMovimiento As Long       'old Id Movimiento de Disponibilidad (caso de modificacion)
    oTotalBruto As Currency     'Total Bruto, p/controlar cambios de importe cdo está pago
    oPesos As Currency              'Total bruto en pesos (p/ cambios de moneda)
    oProveedor As Long          'Para sucesos por modificacion al cerrar disponibilidad
    oSubRubro As Long           '   ""
    oUsuario As Long              '   ""
    
    Disponibilidad As Long
    ImporteCompra As Currency
    ImporteDisponibilidad As Currency
    ImportePesos As Currency
    HaceSalidaCaja As Boolean
    
    cndHayRelCompraPago As Boolean                    'Relaciones c/Tablas
    cndPagoConOtras As Boolean                            'Pago con otras disponibilidades
    cndFCierreDisponibilidad As Date         'Fecha de Cierre de la Disponibilidad
    
    flgSucesoXMod As Boolean            'Si hay suceso por modificacion de datos
End Type

Dim mData As typReg

Private Sub bSplitR_Click()
    
    If Not IsNumeric(lTotalGasto.Caption) Then Exit Sub
    
    frmSplitRubros.prmTotal = CCur(lTotalGasto.Caption)
    frmSplitRubros.prmEdit = sNuevo Or sModificar
    frmSplitRubros.Show vbModal, Me
    Me.Refresh
    
    If frmSplitRubros.prmOK Then CargoRubrosDelArray
    If cDisponibilidad.Enabled Then Foco cDisponibilidad Else Foco tComentario
    
End Sub

Private Sub cComprobante_GotFocus()
    cComprobante.SelStart = 0: cComprobante.SelLength = Len(cComprobante.Text)
End Sub

Private Sub cComprobante_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        On Error Resume Next
        If cComprobante.ListIndex = -1 Then Exit Sub
        
        Dim mTipoC As Integer
        mTipoC = cComprobante.ItemData(cComprobante.ListIndex)
        If mTipoC = TipoDocumento.CompraCredito Or mTipoC = TipoDocumento.CompraNotaCredito Then
            cDisponibilidad.Text = ""
            orCheque.fnc_BlankControls
        End If
        Foco tSerie
    End If
    
End Sub

Private Sub cDisponibilidad_Change()
    If Val(cDisponibilidad.Tag) = 0 Then
        orCheque.fnc_BlankControls
        orCheque.Enabled = False: orCheque.BackColor = Colores.Inactivo
        cDisponibilidad.Tag = 1
    End If
End Sub

Private Sub cDisponibilidad_Click()
    If Val(cDisponibilidad.Tag) = 0 Then
        orCheque.fnc_BlankControls
        orCheque.Enabled = False: orCheque.BackColor = Colores.Inactivo
        cDisponibilidad.Tag = 1
    End If
End Sub

Private Sub cDisponibilidad_KeyDown(KeyCode As Integer, Shift As Integer)
 
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDivide
                cDisponibilidad.ListIndex = 0
                KeyCode = 0
    
        Case vbKeyReturn
                If Not IsNumeric(tIOriginal.Text) Then Foco tIOriginal: Exit Sub
                If cMoneda.ListIndex = -1 Then Foco cMoneda: Exit Sub
                If cDisponibilidad.ListIndex = -1 Then Exit Sub
                cDisponibilidad.Tag = 0
                'Veo si la seleccionada es bancaria
                Dim mID As Long, mIdx As Integer
                mID = cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
                If mID > 0 Then
                    mIdx = dis_IdxArray(mID)
                    If mIdx <> -1 Then
                        If arrDisp(mIdx).Bancaria Then
                            orCheque.Enabled = True: orCheque.BackColor = vbWindowBackground
                            With orCheque
                                .prp_IdDisponibilidad = mID
                                .prp_IdGasto = prmIdCompra
                                .prp_ValorAAsignar = CCur(tIOriginal.Text)
                                .prp_IdCheque = Val(orCheque.Tag)
                                .prp_ValorInicial = 1
                                .fnc_Show
                            End With
                        End If
                    End If
                End If
                
                If orCheque.Enabled Then orCheque.SetFocus Else Foco tComentario
        
    End Select

End Sub

Private Sub cDisponibilidad_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "/" Then KeyAscii = 0
End Sub

Private Sub cDisponibilidad_LostFocus()
    cDisponibilidad.Tag = 0
End Sub

Private Sub chVerificado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then AccionGrabar
    If KeyCode = vbKeyDelete Then chVerificado.Value = vbGrayed
End Sub

Private Sub cMoneda_Change()
    If cDisponibilidad.ListCount > 0 Then cDisponibilidad.Clear
End Sub

Private Sub cMoneda_Click()
    If cDisponibilidad.ListCount > 0 Then cDisponibilidad.Clear
End Sub

Private Sub dFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF1:
            If Not sNuevo And Not sModificar Then
                tProveedor.SetFocus: dFecha.SetFocus
                AccionListaDeAyuda
            End If
            
        Case vbKeyReturn: If tProveedor.Enabled Then Foco tProveedor Else Foco tSerie
    End Select
    
End Sub

Private Sub MnuAAsignarNota_Click()
    Call Toolbar1_ButtonClick(Toolbar1.Buttons("nota"))
End Sub

Private Sub MnuAConQuePaga_Click()
    Call Toolbar1_ButtonClick(Toolbar1.Buttons("pagos"))
End Sub

Private Sub MnuAcProveedores_Click()
    If sNuevo Or sModificar Then EjecutarApp prmPathApp & "Empresas Clientes.exe"
End Sub

Private Sub MnuAcVOpe_Click()
    If Val(tProveedor.Tag) <> 0 Then EjecutarApp prmPathApp & "Visualizacion de Operaciones.exe", CStr(tProveedor.Tag)
End Sub

Private Sub MnuATipoCambio_Click()
    Call Toolbar1_ButtonClick(Toolbar1.Buttons("dolar"))
End Sub

Private Sub MnuAVencimientos_Click()
    Call Toolbar1_ButtonClick(Toolbar1.Buttons("plazos"))
End Sub

Private Sub MnuBx_Click(Index As Integer)

'On Error Resume Next

    If Not AccionCambiarBase(MnuBx(Index).Tag, MnuBx(Index).Caption) Then Exit Sub
    Screen.MousePointer = 11
    
    CargoParametrosImportaciones
    CargoParametrosComercio
    CargoParametrosSucursal
    LoadME
   
    'Cambio el Color del fondo de controles ----------------------------------------------------------------------------------------
    Dim arrC() As String
    arrC = Split(MnuBases.Tag, "|")
    If arrC(Index) <> "" Then Me.BackColor = arrC(Index) Else Me.BackColor = vbButtonFace
    
    Frame1.BackColor = Me.BackColor
    Frame2.BackColor = Me.BackColor
    chVerificado.BackColor = Me.BackColor
    '-------------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
End Sub

Private Sub mnuDocsRel_Click()

    On Error GoTo errAyuda
    
    If prmIdCompra = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    Dim aLista As New clsListadeAyuda
    Dim mIDSel As Long: mIDSel = 0
    
    Cons = " Select ID_Compra = ComCodigo, Fecha = ComFecha, Proveedor = PClFantasia, Comprobante = ComSerie + Convert(char(10), ComNumero), Moneda = MonSigno , Importe = ComImporte, Comentarios = ComComentario" _
            & " from Compra, ProveedorCliente, Moneda" _
            & " Where ComProveedor = PClCodigo" _
            & " And ComMoneda = MonCodigo" _
            & " And ( ComCodigo In (Select CPaDocQSalda from CompraPago Where CPaDocASaldar = " & prmIdCompra & ")" _
            & "     OR ComCodigo In (Select CPaDocASaldar from CompraPago Where CPaDocQSalda = " & prmIdCompra & ") ) "
    
    aLista.ActivoListaAyudaSQL cBase, Cons
    
    Me.Refresh
    DoEvents
    
    If IsNumeric(aLista.ItemSeleccionadoSQL) Then mIDSel = CLng(aLista.ItemSeleccionadoSQL)
    Set aLista = Nothing
    
    If mIDSel <> 0 And Not (sNuevo Or sModificar) Then
        LimpioFicha
        CargoCamposDesdeBD mIDSel
        If prmIdCompra <> 0 Then Botones True, True, True, False, False, Toolbar1, Me
    End If
    Screen.MousePointer = 0
    Exit Sub
        
errAyuda:
    clsGeneral.OcurrioError "Error al activar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub orCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentario
End Sub

Private Sub tRubro_Change()
    
    If Val(tRubro.Tag) <> 0 Then
        tRubro.Tag = 0
        tSubRubro.Text = "": tSubRubro.Tag = 0
        'If Val(tSubRubro.Tag) <> 0 Then tSubRubro.Text = ""
    End If
    
End Sub

Private Sub cMoneda_GotFocus()
    cMoneda.SelStart = 0: cMoneda.SelLength = Len(cMoneda.Text)
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        
        If cMoneda.ListIndex = -1 Then Exit Sub
        If IsDate(dFecha.Value) And cComprobante.ListIndex <> -1 Then
            Dim aFechaTC As String: aFechaTC = ""
            tTCDolar.ToolTipText = "": lTC.Caption = ""
            
            If cComprobante.ItemData(cComprobante.ListIndex) = TipoDocumento.CompraCredito Then ' DEL dia anterior
                'La maxima fecha menor a la del dia anterior a ultima hora
                tTCDolar.Text = TasadeCambio(paMonedaDolar, paMonedaPesos, dFecha.Value - 1, aFechaTC)
                lTC.Caption = aFechaTC
            
            Else        'TC del ultimo dia del mes anterior
                tTCDolar.Text = TasadeCambio(paMonedaDolar, paMonedaPesos, UltimoDia(DateAdd("m", -1, dFecha.Value)), aFechaTC)
                lTC.Caption = aFechaTC
            End If
        End If
        
        If cDisponibilidad.ListCount = 0 Then
            dis_CargoDisponibilidades cDisponibilidad, cMoneda.ItemData(cMoneda.ListIndex)
        End If
        If cComprobante.ItemData(cComprobante.ListIndex) = TipoDocumento.CompraCredito Then
            cDisponibilidad.Text = ""
            orCheque.fnc_BlankControls
        End If
        
        Foco tIOriginal
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    FechaDelServidor
    
    orCheque.fnc_Start cBase
    dis_StartArray
    
    aTiposDocs = TipoDocumento.CompraContado & ", " & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraRecibo & ", " _
                    & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ", " _
                    & TipoDocumento.CompraEntradaCaja & ", " & TipoDocumento.CompraSalidaCaja
                    
    LoadME
    
    If Trim(Command()) <> "" Then CargoCamposDesdeBD Val(Command())
    If prmIdCompra <> 0 Then Botones True, True, True, False, False, Toolbar1, Me
    
    'Me.Show
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
End Sub

Private Sub LoadME()

    On Error Resume Next
    sNuevo = False: sModificar = False
    
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    Status.Panels("bd") = "BD: " & PropiedadesConnect(txtConexion, Database:=True) & " "
        
    CargoDatosCombos
    DeshabilitoIngreso
    LimpioFicha
    
    Botones True, False, False, False, False, Toolbar1, Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco tSubRubro
End Sub

Private Sub Label10_Click()
    Foco tComentario
End Sub

Private Sub Label2_Click()
    Foco tRubro
End Sub

Private Sub Label3_Click()
    Foco tProveedor
End Sub

Private Sub Label4_Click()
    Foco dFecha
End Sub

Private Sub Label6_Click()
    Foco cComprobante
End Sub

Private Sub Label7_Click()
    Foco tIva
End Sub

Private Sub Label8_Click()
    Foco cMoneda
End Sub

Private Sub Label9_Click()
    Foco tTCDolar
End Sub


Private Sub lTotalGasto_Change()
    'If Trim(lTotalGasto.Caption) <> "" And IsNumeric(lTotalGasto.Caption) Then lTotalGasto.Caption = Format(Abs(lTotalGasto.Caption), FormatoMonedaP)
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuExit_Click()
    Unload Me
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuHelp_Click()
On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    Cons = "Select * from Aplicacion Where AplNombre = '" & Trim(prmKeyApp) & "'"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux!AplHelp) Then aFile = Trim(rsAux!AplHelp)
    rsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Sub AccionNuevo(Optional DesdeNuevo As Boolean = False)

    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso

    If DesdeNuevo Then
        InicializoMData
        
        lTotalGasto.Caption = ""
        tNumero.Text = "": tIOriginal.Text = "": tIva.Text = "": tCofis.Text = ""
        tComentario.Text = ""
        
        If Val(tSubRubro.Tag) = 0 Then tRubro.Text = "": tSubRubro.Text = ""
                
        Foco tProveedor
        tAutoriza.Text = ""
        
        cDisponibilidad.Tag = 0
        orCheque.fnc_BlankControls
        orCheque.Tag = 0
    Else
        LimpioFicha
        dFecha.Value = Format(gFechaServidor, "dd/mm/yyyy"): dFecha.SetFocus
    End If
    
    tRubro.Locked = False: tSubRubro.Locked = False
    
    prmIdCompra = 0
    ReDim arrRubros(0)
    arrRubros(0).IdRubro = 0
    
End Sub

Private Sub AccionModificar()

    On Error Resume Next
    Screen.MousePointer = 11
    '1) Si es crédito, Si hay pagos asociados ... No dejar modificar el importe del Gasto
    
    If Not ValidoCompraImportacion Then Exit Sub
    
    CargoCamposDesdeBD prmIdCompra
    
    ChequeoCondicionesModificar
    
    sModificar = True
    Botones False, False, False, True, True, Toolbar1, Me
    
    HabilitoIngreso
    Screen.MousePointer = 0
    
End Sub

Private Function ChequeoCondicionesModificar()

    '>>>  Condición: Hay pagos ingresados
    '1) Cuando esta asignado como nota o recibo a una factura.
    '2) Cuando tiene pagos ingresados.
    mData.cndHayRelCompraPago = False
    'Valido los campos de la tabla CompraPago-------------------------------------------------------------------------------------------
    Cons = "Select * from CompraPago " & _
                " Where CPaDocASaldar = " & prmIdCompra & _
                " OR CPaDocQSalda = " & prmIdCompra
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then mData.cndHayRelCompraPago = True
    rsAux.Close

    '>>>  Condición: Pago con otras Disponibilidades
    mData.cndPagoConOtras = False
    If cDisponibilidad.ListIndex <> -1 Then
        If cDisponibilidad.ItemData(cDisponibilidad.ListIndex) = 0 Then mData.cndPagoConOtras = True
    End If
    
End Function

Private Sub AccionGrabar()
   
Dim aError As String: aError = ""
Dim bNuevoIngreso As Boolean: bNuevoIngreso = False
Dim bDispEnabled As Boolean

    bDispEnabled = cDisponibilidad.Enabled
    
    bIngresarVencimientos = False
    Screen.MousePointer = 11
    
    If Not ValidoCampos Then Screen.MousePointer = 0: Exit Sub
    If Not ValidoDocumento Then Screen.MousePointer = 0: Exit Sub
    If Not ValidoCompraImportacion Then Screen.MousePointer = 0: Exit Sub
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    If paCodigoDeUsuario = 0 Then
        If miConexion.AccesoAlMenu(prmKeyApp) Then
            paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        Else
            MsgBox "Sin acceso a la aplicación.", vbExclamation, "ATENCIÓN"
        End If
        If paCodigoDeUsuario = 0 Then Screen.MousePointer = 0: Exit Sub
    End If
    
    Screen.MousePointer = 0
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "Grabar Gasto") = vbNo Then Exit Sub
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    FechaDelServidor
    
    Dim mCompra As Long: mCompra = 0
    
    If sNuevo Then bNuevoIngreso = True
    If sModificar Then mCompra = prmIdCompra
    If sModificar And bDispEnabled Then
        If Not ValidoIngresoDeVencimientos Then Screen.MousePointer = 0: Exit Sub
    End If
    
    cBase.BeginTrans    'COMIENZO TRANSACCION----------------------------------------------!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    'Tabla Compra---------------------------------------------------------------------------------------------------------------
    Cons = "Select * from Compra Where ComCodigo = " & mCompra
    Set rsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If sModificar Then
         If prmFModificacion <> rsCom!ComFModificacion Then
            aError = "El comprobante ha sido modificado por otro usuario. Vuelva a cargar los datos"
            GoTo errorET: Exit Sub
        End If
    End If
    If rsCom.EOF Then rsCom.AddNew Else rsCom.Edit
    CargoCamposBDComprobante
    rsCom.Update: rsCom.Close
    '--------------------------------------------------------------------------------------------------------------------------------
    
    If sNuevo Then
        Cons = "Select Max(ComCodigo) from Compra" & _
                " Where ComFecha = " & Format(dFecha.Value, "'mm/dd/yyyy'") & _
                " And ComTipoDocumento = " & cComprobante.ItemData(cComprobante.ListIndex) & _
                " And ComProveedor = " & Val(tProveedor.Tag) & _
                " And ComMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
        Set rsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        mCompra = rsCom(0)
        rsCom.Close
    End If

    'Tabla  GastosSubRubro
    CargoCamposBDGastos mCompra, sModificar
        
    If bDispEnabled Then GraboElPago mCompra
    
    If dSuceso.Tipo <> 0 Then
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, dSuceso.Tipo, paCodigoDeTerminal, _
                        dSuceso.Usuario, 0, _
                        Descripcion:=dSuceso.Titulo, Defensa:=dSuceso.Defensa, _
                        Valor:=dSuceso.Valor, idCliente:=dSuceso.Cliente, idAutoriza:=dSuceso.Autoriza
    End If
    
    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
    
    If sNuevo Then
        If cComprobante.ItemData(cComprobante.ListIndex) = TipoDocumento.CompraCredito Then bIngresarVencimientos = True
    End If
    
    prmIdCompra = mCompra: prmFModificacion = gFechaServidor
    If sNuevo Then AccionIrANota prmIdCompra
    If sModificar Then lModificado.Caption = "Mod. " & Format(gFechaServidor, "dd/mm/yy hh:mm") & " x " & miConexion.UsuarioLogueado(Nombre:=True)
    
    sNuevo = False: sModificar = False
    DeshabilitoIngreso
    Botones True, True, True, False, False, Toolbar1, Me
    dFecha.SetFocus
    
    If mData.Disponibilidad = 0 And mData.oIDMovimiento = 0 Then AccionIrA ConQuePaga:=True
    If bIngresarVencimientos Then AccionIrA Vencimientos:=True
    If bNuevoIngreso Then AccionNuevo True
    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    If Trim(aError) = "" Then aError = "No se ha podido inicializar la transacción. Reintente la operación."
    Screen.MousePointer = 0
    clsGeneral.OcurrioError aError, Err.Description
End Sub

Private Sub AccionEliminar()
Dim aError As String

    If Not ValidoCompraImportacion Then Exit Sub
    
    If Not ValidoDatosMovimientos(prmIdCompra, paraEliminar:=True) Then Exit Sub
    If Not ValidoDatosEliminar Then Screen.MousePointer = 0: Exit Sub
    
    If Not zPidoSuceso(prmSucesoGastos, "Eliminar Gastos") Then Exit Sub
    
    If MsgBox("Confirma eliminar el gasto seleccionado", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        FechaDelServidor
        Screen.MousePointer = 11
        
        On Error GoTo errorBT
        cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
        On Error GoTo errorET
        
        Cons = "Select * from Compra Where ComCodigo = " & prmIdCompra
        Set rsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If prmFModificacion <> rsCom!ComFModificacion Then
            aError = "El comprobante ha sido modificado recientemente por otro usuario. Vuelva a cargar los datos"
            GoTo errorET: Exit Sub
        End If
        
        'Elimino tabla: Gastos Subrubro e Importacion
        Cons = "Delete GastoSubrubro Where GSrIDCompra = " & prmIdCompra
        cBase.Execute Cons
                
        rsCom.Delete: rsCom.Close
        
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, dSuceso.Tipo, paCodigoDeTerminal, _
                        dSuceso.Usuario, 0, _
                        Descripcion:=dSuceso.Titulo, Defensa:=dSuceso.Defensa, _
                        Valor:=dSuceso.Valor, idCliente:=dSuceso.Cliente, idAutoriza:=dSuceso.Autoriza
                
        cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
        LimpioFicha
        DeshabilitoIngreso
        Botones True, False, False, False, False, Toolbar1, Me
        prmIdCompra = 0
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación.", Err.Description
End Sub

Private Sub AccionCancelar()

    LimpioFicha
    If sModificar Then
        Botones True, True, True, False, False, Toolbar1, Me
        CargoCamposDesdeBD prmIdCompra
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    
    DeshabilitoIngreso
    sNuevo = False: sModificar = False
    dFecha.SetFocus

End Sub

Private Sub CargoCamposBDComprobante()

    
    If sNuevo Then      'Si es Nuevo y es Credito ---> pongo el saldo
        Select Case cComprobante.ItemData(cComprobante.ListIndex)
            Case TipoDocumento.CompraCredito: rsCom!ComSaldo = CCur(tIOriginal.Text)
        End Select
    End If
        
    If sModificar Then  'Si es modificar  solo toco el saldo si (Importe + Iva) = Saldo ---> no hay pagos ingresados.
        Select Case cComprobante.ItemData(cComprobante.ListIndex)
            Case TipoDocumento.CompraCredito
                If rsCom!ComTipoDocumento = TipoDocumento.CompraCredito Then  'Si antes era credito
                    If Not IsNull(rsCom!ComSaldo) Then
                        Dim aImporte As Currency
                        aImporte = rsCom!ComImporte: If Not IsNull(rsCom!ComIva) Then aImporte = aImporte + rsCom!ComIva
                         If Not IsNull(rsCom!ComCofis) Then aImporte = aImporte + rsCom!ComCofis
                        If rsCom!ComSaldo = aImporte Then
                            rsCom!ComSaldo = CCur(tIOriginal.Text)
                        End If
                    End If
                'End If
                'If rsCom!ComTipoDocumento = TipoDocumento.CompraContado Then  'Si antes era contado
                Else
                    rsCom!ComSaldo = CCur(tIOriginal.Text)
                End If
            
            Case Else
                    rsCom!ComSaldo = 0
        End Select
    End If
        
    rsCom!ComTipoDocumento = cComprobante.ItemData(cComprobante.ListIndex)
    rsCom!ComFecha = Format(dFecha.Value, "mm/dd/yyyy")
    
    If cMoneda.Enabled Then rsCom!ComMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    rsCom!ComImporte = CCur(lTotalGasto.Caption)
    If Trim(tIva.Text) <> "" Then rsCom!ComIva = CCur(tIva.Text) Else rsCom!ComIva = Null
    If Trim(tCofis.Text) <> "" Then rsCom!ComCofis = CCur(tCofis.Text) Else rsCom!ComCofis = Null
    rsCom!ComTC = CCur(tTCDolar.Text)
    
    rsCom!ComProveedor = Val(tProveedor.Tag)
    If Trim(tSerie.Text) <> "" Then rsCom!ComSerie = Trim(tSerie.Text) Else rsCom!ComSerie = Null
    If Trim(tNumero.Text) <> "" Then rsCom!ComNumero = tNumero.Text Else rsCom!ComNumero = Null
    
    If Trim(tComentario.Text) <> "" Then rsCom!ComComentario = Trim(tComentario.Text) Else rsCom!ComComentario = Null
    
    rsCom!ComFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    rsCom!ComUsuario = paCodigoDeUsuario
    
    If sNuevo Then
        rsCom!ComUsrAutoriza = Val(tAutoriza.Tag)
        'If Val(tAutoriza.Tag) = tUsuario.UserID Then rsCom!ComVerificado = 1 Else rsCom!ComVerificado = Null
        rsCom!ComVerificado = Null
    Else
        If tAutoriza.Enabled Then
            rsCom!ComUsrAutoriza = Val(tAutoriza.Tag)
            'If chVerificado.Value = vbChecked Then rsCom!ComVerificado = 1 Else rsCom!ComVerificado = 0
            'If Val(tAutoriza.Tag) = tUsuario.UserID Then rsCom!ComVerificado = 1
            Select Case chVerificado.Value
                Case vbChecked: rsCom!ComVerificado = 1
                Case vbUnchecked: rsCom!ComVerificado = 0
                Case vbGrayed: rsCom!ComVerificado = Null
            End Select
            
            If (Val(tAutoriza.Tag) <> paCodigoDeUsuario) Then rsCom!ComVerificado = Null
        End If
    End If
    'If Not rsCom!ComVerificado Then chVerificado.Value = vbUnchecked Else chVerificado.Value = vbChecked
    If IsNull(rsCom!ComVerificado) Then
        chVerificado.Value = vbGrayed
    Else
        chVerificado.Value = IIf(rsCom!ComVerificado = 1, vbChecked, vbUnchecked)
    End If

End Sub

Private Sub CargoCamposBDGastos(idCompra As Long, bBorrarAnterior As Boolean)

    If bBorrarAnterior Then
        Cons = "Delete GastoSubrubro Where GSrIDCompra = " & idCompra
        cBase.Execute Cons
    End If
    
    Cons = "Select * from GastoSubrubro Where GSrIDCompra = " & idCompra
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    For I = LBound(arrRubros) To UBound(arrRubros)
        rsAux.AddNew
        rsAux!GSrIDCompra = idCompra
        rsAux!GSrIDSubrubro = arrRubros(I).IdSRubro
        rsAux!GSrImporte = arrRubros(I).Importe
        rsAux.Update
    Next I

    rsAux.Close
        
End Sub

Private Sub CargoCamposDesdeBD(idCompra As Long)

Dim aValor As Long

    Screen.MousePointer = 11
    
    InicializoMData
    
    On Error GoTo errCargar
    'Cargo los datos desde la tabla COMPRA-----------------------------------------------------------------------------------------
    Cons = " Select Compra.*, UsrA.UsuCodigo as UsuACodigo, UsrA.UsuIdentificacion as UsuAIdentificacion, " & _
                    " UsrC.UsuCodigo as UsuCCodigo, UsrC.UsuIdentificacion as UsuCIdentificacion " & _
                "FROM Compra " & _
                    " Left Outer Join ZUREOCGSA.dbo.admUsuarios UsrA On ComUsrAutoriza = UsrA.UsuCodigo " & _
                    " Left Outer Join ZUREOCGSA.dbo.admUsuarios UsrC On ComUsuario = UsrC.UsuCodigo " & _
                " Where ComCodigo = " & idCompra & _
                " And ComTipoDocumento In (" & aTiposDocs & ")"
    
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        rsAux.Close
        MsgBox "No existe registro de compra para el id ingresado.", vbExclamation, "No hay Datos"
        Botones True, False, False, False, False, Toolbar1, Me
        prmIdCompra = 0
        Screen.MousePointer = 0: Exit Sub
    End If
        
    prmIdCompra = rsAux!ComCodigo
    prmFModificacion = rsAux!ComFModificacion
    
    tID.Text = Format(rsAux!ComCodigo, "#,##0")
    dFecha.Value = Format(rsAux!ComFecha, "dd/mm/yyyy")
    mData.oFechaCompra = rsAux!ComFecha
    
    Dim rs1 As rdoResultset
    Cons = "Select * from ProveedorCliente Where PClCodigo = " & rsAux!ComProveedor
    Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rs1.EOF Then
        tProveedor.Text = Trim(rs1!PClFantasia)
        tProveedor.Tag = rsAux!ComProveedor
    End If
    rs1.Close
    mData.oProveedor = Val(tProveedor.Tag)
    
    BuscoCodigoEnCombo cComprobante, rsAux!ComTipoDocumento
    If Not IsNull(rsAux!ComSerie) Then tSerie.Text = Trim(rsAux!ComSerie)
    If Not IsNull(rsAux!ComNumero) Then tNumero.Text = rsAux!ComNumero
    
    BuscoCodigoEnCombo cMoneda, rsAux!ComMoneda
    dis_CargoDisponibilidades cDisponibilidad, cMoneda.ItemData(cMoneda.ListIndex)
    
    lTotalGasto.Caption = Format(rsAux!ComImporte, "##,##0.00")
    Dim aImp As Currency
    
    aImp = rsAux!ComImporte
    tIva.Text = "": tCofis.Text = ""
    If Not IsNull(rsAux!ComIva) Then
        tIva.Text = Format(rsAux!ComIva, FormatoMonedaP)
        aImp = aImp + rsAux!ComIva
    End If
    If Not IsNull(rsAux!ComCofis) Then
        tCofis.Text = Format(rsAux!ComCofis, FormatoMonedaP)
        aImp = aImp + rsAux!ComCofis
    End If
    tIOriginal.Text = Format(aImp, FormatoMonedaP)
    mData.oTotalBruto = Format(aImp, FormatoMonedaP)
    
    If Not IsNull(rsAux!ComTC) Then If rsAux!ComTC <> 1 Then tTCDolar.Text = Format(rsAux!ComTC, "0.000")
    
    If rsAux!ComMoneda = paMonedaPesos Then
        mData.oPesos = mData.oTotalBruto
    Else
        mData.oPesos = mData.oTotalBruto * CCur(tTCDolar.Text)
    End If
        
    If Not IsNull(rsAux!ComComentario) Then tComentario.Text = Trim(rsAux!ComComentario)
    
    If Not IsNull(rsAux!UsuACodigo) Then
        tAutoriza.Text = Trim(rsAux!UsuAIdentificacion)
        tAutoriza.Tag = rsAux!UsuACodigo
    End If
    lAutoriza.Tag = tAutoriza.Tag
    
    If Not IsNull(rsAux!ComVerificado) Then
        chVerificado.Value = IIf(rsAux!ComVerificado = 1, vbChecked, vbUnchecked)
    Else
        chVerificado.Value = vbGrayed
    End If
    
    If Not IsNull(rsAux!ComFModificacion) Or Not IsNull(rsAux!UsuCIdentificacion) Then
        Dim sText As String
        sText = "Mod. "
        If Not IsNull(rsAux!ComFModificacion) Then sText = sText & Format(rsAux!ComFModificacion, "dd/mm/yy hh:mm") & " "
        If Not IsNull(rsAux!UsuCIdentificacion) Then sText = sText & "x " & Trim(rsAux!UsuCIdentificacion)
        lModificado.Caption = sText
    End If
    If Not IsNull(rsAux!ComUsuario) Then
'        tUsuario.UserID = rsAux!ComUsuario
        mData.oUsuario = rsAux!ComUsuario
    End If
    rsAux.Close
    
    'Cargo los datos desde la BD GastosSubRubro-----------------------------------------------------------------------------------------
    Dim idX As Integer: idX = 0
    ReDim arrRubros(0)
    Cons = "Select * from GastoSubrubro, SubRubro, Rubro " _
           & " Where GSrIDCompra = " & idCompra _
           & " And GSrIDSubrubro = SRuID And SRuRubro = RubID"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        ReDim Preserve arrRubros(idX)
        
        With arrRubros(idX)
            .IdRubro = rsAux!RubID
            .TextoRubro = Trim(rsAux!RubNombre)
            .IdSRubro = rsAux!SRuID
            .TextoSRubro = Trim(rsAux!SRuNombre)
            .Importe = Format(rsAux!GSrImporte, "#,##0.00")
        End With
        idX = idX + 1
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    If UBound(arrRubros) = 0 And arrRubros(0).IdRubro <> 0 Then mData.oSubRubro = arrRubros(0).IdSRubro
    
    CargoRubrosDelArray
    
    'Busco los movimientos de Disponibilidades para Ver con que se Pagó --------------------------------------------------------
    Dim mIDDPago As Long, mIDCheque As Long
    mIDDPago = -1: mIDCheque = 0
    Cons = "Select * from MovimientoDisponibilidad, MovimientoDisponibilidadRenglon" & _
               " Where MDiIDCompra = " & idCompra & _
               " And MDiID = MDRIDMovimiento "
    Set rsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not rsAux.EOF Then
        mIDDPago = rsAux!MDRIdDisponibilidad
        mData.oIDMovimiento = rsAux!MDiID
        mData.cndFCierreDisponibilidad = CDate("1/1/1900")
        If Not IsNull(rsAux!MDRIdCheque) Then mIDCheque = rsAux!MDRIdCheque
        
        Dim aDate As Date
        Do While Not rsAux.EOF
            aDate = dis_FechaCierre(rsAux!MDRIdDisponibilidad, mData.oFechaCompra)
            If aDate > mData.cndFCierreDisponibilidad Then mData.cndFCierreDisponibilidad = aDate
            rsAux.MoveNext
            If Not rsAux.EOF Then mIDDPago = 0
        Loop
    End If
    rsAux.Close
    
    orCheque.Tag = 0
    If mIDDPago >= 0 Then
        BuscoCodigoEnCombo cDisponibilidad, mIDDPago
        If cDisponibilidad.ListIndex = -1 Then 'Se pago con otra Disponibilidad,  <> a la MonedadelGasto o Varias Disp.
            BuscoCodigoEnCombo cDisponibilidad, 0
        Else
            If mIDDPago > 0 And mIDCheque <> 0 Then
                With orCheque
                    .prp_IdCheque = mIDCheque
                    .prp_IdDisponibilidad = mIDDPago
                    .prp_IdGasto = idCompra
                    .prp_ValorAAsignar = CCur(tIOriginal.Text)
                    .fnc_Show
                    .Tag = mIDCheque
                End With
            End If
        End If
        
    Else
        cDisponibilidad.Text = ""
        orCheque.fnc_BlankControls
    End If
    cDisponibilidad.Tag = 0
    '-------------------------------------------------------------------------------------------------------------------------------------------
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del comprobante.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tAutoriza_Change()
    tAutoriza.Tag = 0
End Sub

Private Sub tAutoriza_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Val(tAutoriza.Tag) <> 0 Then
            If sModificar Then
                If Val(tAutoriza.Tag) <> Val(lAutoriza.Tag) Then chVerificado.Enabled = False: chVerificado.Value = vbUnchecked
            End If
            If chVerificado.Enabled Then chVerificado.SetFocus Else AccionGrabar
            Exit Sub
        End If
        
        If Trim(tAutoriza.Text) = "" Then Exit Sub
        On Error GoTo errBsucarUsr
        Screen.MousePointer = 11
        
        Cons = "Select UsuCodigo, UsuIdentificacion as 'Identificación', Cast(UsuDigito as Char(4)) as 'Dígito' From Usuario " & _
                   " Where UsuHabilitado = 1"
        If IsNumeric(tAutoriza.Text) Then
            Cons = Cons & " And UsuDigito = " & Val(tAutoriza.Text)
        Else
            Cons = Cons & " And UsuIdentificacion like '" & Replace(Trim(tAutoriza.Text), " ", "%") & "%'"
        End If
        Cons = Cons & " Order by UsuIdentificacion"
        
        Dim aQ As Integer, aIdUsr As Long, aName As String
        aQ = 0
        
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            aQ = 1: aIdUsr = rsAux!UsuCodigo: aName = Trim(rsAux(1))
            rsAux.MoveNext
            If Not rsAux.EOF Then aQ = 2: aIdUsr = 0
        End If
        rsAux.Close
        
        Select Case aQ
            Case 0: MsgBox "No hay usuarios que coincidan con la búsqueda ingresada.", vbInformation, "No Hay Datos"
            
            Case 2:
                    Dim miHelp As New clsListadeAyuda
                    aIdUsr = miHelp.ActivarAyuda(cBase, Cons, 4100, 1, "Usuarios")
                    Me.Refresh
                    If aIdUsr <> 0 Then
                        aIdUsr = miHelp.RetornoDatoSeleccionado(0)
                        aName = Trim(miHelp.RetornoDatoSeleccionado(1))
                    End If
                    Set miHelp = Nothing
        End Select
        
        If aIdUsr > 0 Then
            tAutoriza.Text = aName
            tAutoriza.Tag = aIdUsr
        End If
        
        Screen.MousePointer = 0
        
        If Val(tAutoriza.Tag) <> 0 Then
            If sModificar Then
                If Val(tAutoriza.Tag) <> Val(lAutoriza.Tag) Then chVerificado.Enabled = False: chVerificado.Value = vbUnchecked
            End If
            If chVerificado.Enabled Then chVerificado.SetFocus Else AccionGrabar
            Exit Sub
        End If
        
    End If
    Exit Sub

errBsucarUsr:
    clsGeneral.OcurrioError "Error al buscar el usuario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCofis_GotFocus()
    With tCofis: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tCofis_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tTCDolar.Enabled Then Foco tTCDolar Else Foco tSubRubro
    End If
End Sub

Private Sub tCofis_LostFocus()

    If Not IsNumeric(tCofis.Text) Then
        tCofis.Text = ""
    Else
        tCofis.Text = Format(tCofis.Text, "##,##0.00")
        
        If cComprobante.ListIndex <> -1 Then
            Select Case cComprobante.ItemData(cComprobante.ListIndex)
                Case TipoDocumento.CompraNotaCredito, TipoDocumento.CompraNotaDevolucion, TipoDocumento.CompraEntradaCaja
                    tCofis.Text = Format(Abs(CCur(tCofis.Text)) * -1, "##,##0.00")
                Case Else
                    tCofis.Text = Format(Abs(CCur(tCofis.Text)), "##,##0.00")
            End Select
        End If
    End If
    
    Dim aImp As Currency, aIva As Currency, aCofis As Currency
    aImp = 0: aIva = 0: aCofis = 0
    If Val(tIOriginal.Text) <> 0 Then
        If Val(tIva.Text) <> 0 Then aIva = CCur(tIva.Text)
        If Val(tCofis.Text) <> 0 Then aCofis = CCur(tCofis.Text)
        aImp = CCur(tIOriginal.Text)
         lTotalGasto.Caption = Format(aImp - aIva - aCofis, FormatoMonedaP)
    End If
    
End Sub

Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0: tComentario.SelLength = Len(tComentario.Text)
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tAutoriza.Enabled Then Foco tAutoriza Else AccionGrabar
    End If
End Sub

Private Sub AccionListaDeAyuda()

    On Error GoTo errAyuda
    
    If Not IsDate(dFecha.Value) And Val(tProveedor.Tag) = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    Dim aLista As New clsListadeAyuda
    Dim aSeleccionado As Long: aSeleccionado = 0
    
    Cons = " Select ID_Compra = ComCodigo, Fecha = ComFecha, Proveedor = PClFantasia, Comprobante = ComSerie + Convert(char(10), ComNumero), Moneda = MonSigno , Importe = ComImporte, Comentarios = ComComentario" _
            & " from Compra, ProveedorCliente, Moneda" _
            & " Where ComProveedor = PClCodigo" _
            & " And ComMoneda = MonCodigo" _
            & " And ComTipoDocumento In (" & aTiposDocs & ")"
            
    If IsDate(dFecha.Value) Then Cons = Cons & " And ComFecha >= '" & Format(dFecha.Value, sqlFormatoF) & "'"
    If Val(tProveedor.Tag) <> 0 Then Cons = Cons & " And ComProveedor = " & Val(tProveedor.Tag)
    
    Cons = Cons & " Order by ComFecha DESC"
    
    aLista.ActivoListaAyudaSQL cBase, Cons
    
    Me.Refresh
    DoEvents
    
    If IsNumeric(aLista.ItemSeleccionadoSQL) Then aSeleccionado = CLng(aLista.ItemSeleccionadoSQL)
    Set aLista = Nothing
    
    If aSeleccionado <> 0 Then LimpioFicha: CargoCamposDesdeBD aSeleccionado
    If prmIdCompra <> 0 Then Botones True, True, True, False, False, Toolbar1, Me
    
    Screen.MousePointer = 0
    Exit Sub
        
errAyuda:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al activar la lista de ayuda.", Err.Description
End Sub

Private Sub tID_Change()
    If tID.Enabled Then Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub tID_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tID.Text) = "" Then dFecha.SetFocus: Exit Sub
        If Not IsNumeric(tID.Text) Then MsgBox "El id ingresado no es correcto. Verifique.", vbExclamation, "Posible Error": Exit Sub
        prmIdCompra = CLng(tID.Text)
        LimpioFicha
        CargoCamposDesdeBD prmIdCompra
        If prmIdCompra <> 0 Then Botones True, True, True, False, False, Toolbar1, Me
    End If
    
End Sub

Private Sub tNumero_GotFocus()
    tNumero.SelStart = 0: tNumero.SelLength = Len(tNumero.Text)
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cMoneda.Enabled Then Foco cMoneda: Exit Sub
        If tIOriginal.Enabled Then Foco tIOriginal Else Foco tIva
    End If
End Sub

Private Sub tIOriginal_GotFocus()
    tIOriginal.SelStart = 0: tIOriginal.SelLength = Len(tIOriginal.Text)
End Sub

Private Sub tIOriginal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tIva
End Sub

Private Sub tIOriginal_LostFocus()
    If Not sNuevo And Not sModificar Then Exit Sub
    
    If Not IsNumeric(tIOriginal.Text) Then
        tIOriginal.Text = ""
    Else
        tIOriginal.Text = Format(tIOriginal.Text, "##,##0.00")
            
        If cComprobante.ListIndex <> -1 Then
            Select Case cComprobante.ItemData(cComprobante.ListIndex)
                Case TipoDocumento.CompraNotaCredito, TipoDocumento.CompraNotaDevolucion, TipoDocumento.CompraEntradaCaja
                    tIOriginal.Text = Format(Abs(CCur(tIOriginal.Text)) * -1, "##,##0.00")
                Case Else
                    tIOriginal.Text = Format(Abs(CCur(tIOriginal.Text)), "##,##0.00")
            End Select
        End If
    End If
    
    Dim aImp As Currency: aImp = 0
    lTotalGasto.Caption = "0.00"
    If Val(tIOriginal.Text) <> 0 Then
        If cComprobante.ListIndex <> -1 Then
            Select Case cComprobante.ItemData(cComprobante.ListIndex)
                Case TipoDocumento.CompraContado, TipoDocumento.CompraCredito, TipoDocumento.CompraNotaCredito, TipoDocumento.CompraNotaDevolucion
                    Dim aNeto As Currency
                    tIva.Text = Format(CCur(tIOriginal.Text) - CCur(tIOriginal.Text) / (paIvaMora / 100 + 1), FormatoMonedaP)
                    aNeto = Format(CCur(tIOriginal.Text) - CCur(tIva.Text), "##,##0.00")
                    tCofis.Text = Format(aNeto - (aNeto / ((paCofis / 100) + 1)), "##,##0.00")
                    lTotalGasto.Caption = Format(aNeto - CCur(tCofis.Text), "##,##0.00")
                    
                Case Else
                    tIva.Text = "": tCofis.Text = ""
                    lTotalGasto.Caption = Format(CCur(tIOriginal.Text), "##,##0.00")
            End Select
        End If
    End If
    
End Sub

Private Sub tIva_GotFocus()
    tIva.SelStart = 0: tIva.SelLength = Len(tIva.Text)
End Sub

Private Sub tIva_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tIva.Text) <> 0 Then
            If IsNumeric(tIOriginal.Text) Then
                If ((CCur(tIva.Text) * (100 + paIvaMora) / CCur(tIOriginal.Text))) > paIvaMora + 0.5 Then
                    If MsgBox("El importe que ud. ingresó de IVA puede no ser correcto." & vbCrLf & _
                                "Quiere verificarlo.", vbQuestion + vbYesNo, "Posible Error ") = vbYes Then Exit Sub
                End If
            End If
            Foco tCofis: Exit Sub
        End If
        If tTCDolar.Enabled Then Foco tTCDolar Else Foco tSubRubro
    End If
End Sub

Private Sub tIva_LostFocus()
    
    If Not IsNumeric(tIva.Text) Then
        tIva.Text = "": tCofis.Text = ""
    Else
        tIva.Text = Format(tIva.Text, "##,##0.00")
        
        If cComprobante.ListIndex <> -1 Then
            Select Case cComprobante.ItemData(cComprobante.ListIndex)
                Case TipoDocumento.CompraNotaCredito, TipoDocumento.CompraNotaDevolucion, TipoDocumento.CompraEntradaCaja
                    tIva.Text = Format(Abs(CCur(tIva.Text)) * -1, "##,##0.00")
                Case Else
                    tIva.Text = Format(Abs(CCur(tIva.Text)), "##,##0.00")
            End Select
        End If
    End If
    
    Dim aImp As Currency: aImp = 0
    If Val(tIva.Text) = 0 Then
        If Val(tIOriginal.Text) <> 0 Then lTotalGasto.Caption = tIOriginal.Text
    Else
        If Val(tIOriginal.Text) <> 0 Then
            lTotalGasto.Caption = Format(CCur(tIOriginal.Text) - CCur(tIva.Text), FormatoMonedaP)
            If Val(tCofis.Text) <> 0 Then lTotalGasto.Caption = Format(CCur(tIOriginal.Text) - CCur(tIva.Text) - CCur(tCofis.Text), FormatoMonedaP)
        End If
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        
        Case "pagos": EjecutarApp prmPathApp & "Con Que Paga.exe", Str(prmIdCompra) & "|" & prmKeyConnect
        Case "plazos": EjecutarApp prmPathApp & "Vencimiento de Pagos.exe", Str(prmIdCompra) & "|" & prmKeyConnect
        Case "dolar": EjecutarApp prmPathApp & "Tasa de Cambio"
        
        Case "nota": If prmIdCompra <> 0 Then AccionIrANota prmIdCompra, True
    
    End Select

End Sub

Private Sub AccionIrA(Optional Vencimientos As Boolean = False, Optional ConQuePaga As Boolean = False)

    If Vencimientos Then
        If MsgBox("Desea ingresar los vencimientos de las cuotas.", vbQuestion + vbYesNo + vbDefaultButton2, "Ingreso de Vencimientos") = vbNo Then Exit Sub
        EjecutarApp prmPathApp & "Vencimiento de Pagos.exe", Str(prmIdCompra)
    End If
    
    If ConQuePaga Then
        If MsgBox("Desea ingresar Con Que Paga el Gasto ?.", vbQuestion + vbYesNo, "Ingresa el Pago ?") = vbNo Then Exit Sub
        EjecutarApp prmPathApp & "Con Que Paga.exe", Str(prmIdCompra) & "|" & prmKeyConnect
    End If
    
End Sub

Private Sub AccionIrANota(idNota As Long, Optional Validar As Boolean = False)
    
    If Validar Then
        Dim sSalir As Boolean: sSalir = True
        Cons = " Select * from Compra Where ComCodigo = " & idNota
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            Select Case rsAux!ComTipoDocumento
                Case TipoDocumento.CompraNotaCredito, TipoDocumento.CompraNotaDevolucion: sSalir = False
            End Select
        End If
        rsAux.Close
        
        If sSalir Then
            MsgBox "El comprobante seleccionado no es un nota. " & Chr(vbKeyReturn) & "Esta acción permite relacionar la nota con la factura original.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
    End If
    
    Select Case cComprobante.ItemData(cComprobante.ListIndex)
        Case TipoDocumento.CompraNotaCredito, TipoDocumento.CompraNotaDevolucion
            frmAsignarNota.pComprobante = idNota
            frmAsignarNota.Show vbModal, Me
    End Select
    
End Sub

Private Sub DeshabilitoIngreso()

    tID.BackColor = vbWindowBackground: tID.Enabled = True
    dFecha.Enabled = True ': dFecha.BackColor = Blanco
    tProveedor.Enabled = True: tProveedor.BackColor = vbWindowBackground
    cComprobante.Enabled = False: cComprobante.BackColor = Inactivo
    tSerie.Enabled = False: tSerie.BackColor = Inactivo
    tNumero.Enabled = False: tNumero.BackColor = Inactivo
    
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tIOriginal.Enabled = False: tIOriginal.BackColor = Inactivo
    tIva.Enabled = False: tIva.BackColor = Inactivo
    tTCDolar.Enabled = False: tTCDolar.BackColor = Inactivo
    tCofis.Enabled = False: tCofis.BackColor = Inactivo
    
    tRubro.Enabled = False: tRubro.BackColor = Inactivo
    tSubRubro.Enabled = False: tSubRubro.BackColor = Inactivo
        
    tComentario.Enabled = False: tComentario.BackColor = Inactivo
        
    tAutoriza.Enabled = False: tAutoriza.BackColor = Inactivo
    chVerificado.Enabled = False
    'tUsuario.Enabled = False: tUsuario.BackColor = Colores.Inactivo
    
    cDisponibilidad.Enabled = False: cDisponibilidad.BackColor = Colores.Inactivo
    orCheque.Enabled = False: orCheque.BackColor = Colores.Inactivo
    
End Sub

Private Sub HabilitoIngreso()
    
    tID.BackColor = Inactivo: tID.Enabled = False
    mData.flgSucesoXMod = False
    
    If sModificar Then
        'Accion Modificar   ----------------------------------------------------------------------------
        'Fecha  --------
        dFecha.Enabled = True
        Dim mTipoC As Integer
        mTipoC = cComprobante.ItemData(cComprobante.ListIndex)
        Select Case mTipoC
            Case TipoDocumento.CompraCredito
                    If mData.oFechaCompra <= prmFCierreIVA Then dFecha.Enabled = False
            
            Case Else
                    If mData.oFechaCompra <= prmFCierreIVA Then dFecha.Enabled = False
                    If mData.oFechaCompra <= mData.cndFCierreDisponibilidad Then dFecha.Enabled = False
        End Select
        
        Dim bADM As Boolean
        
        If mData.oFechaCompra < prmFCierreIVA Then      'ANTES DEL CIERRE DEL IVA
        
            mData.flgSucesoXMod = True  'Si se hizo el cierre del iva va suceso siempre
            bADM = miConexion.AccesoAlMenu(prmKeyAppADM)
            
            If bADM Then
                tProveedor.Enabled = True: tProveedor.BackColor = Colores.Obligatorio
                tSerie.Enabled = True: tSerie.BackColor = vbWindowBackground
                tNumero.Enabled = True: tNumero.BackColor = vbWindowBackground
                tTCDolar.Enabled = True: tTCDolar.BackColor = Colores.Obligatorio
                tComentario.Enabled = True: tComentario.BackColor = vbWindowBackground
                
                tRubro.Enabled = True: tRubro.BackColor = vbWindowBackground
                tSubRubro.Enabled = True: tSubRubro.BackColor = vbWindowBackground
            Else
                bSplitR.Enabled = False
                tProveedor.Enabled = False: tProveedor.BackColor = Colores.Inactivo
            End If
            
        ElseIf mData.oFechaCompra > mData.cndFCierreDisponibilidad Then     'MAYOR AL CIERRE DE LA DISP.
                    
            tProveedor.Enabled = True: tProveedor.BackColor = Colores.Obligatorio
            cComprobante.Enabled = True: cComprobante.BackColor = Colores.Obligatorio
            cMoneda.Enabled = True: cMoneda.BackColor = Colores.Obligatorio
            tSerie.Enabled = True: tSerie.BackColor = vbWindowBackground
            tNumero.Enabled = True: tNumero.BackColor = vbWindowBackground
            
            tIOriginal.Enabled = True: tIOriginal.BackColor = Colores.Obligatorio
            tIva.Enabled = True: tIva.BackColor = vbWindowBackground
            tCofis.Enabled = True: tCofis.BackColor = vbWindowBackground
            tTCDolar.Enabled = True: tTCDolar.BackColor = Colores.Obligatorio
        
            tRubro.Enabled = True: tRubro.BackColor = vbWindowBackground
            tSubRubro.Enabled = True: tSubRubro.BackColor = vbWindowBackground
            
            tComentario.Enabled = True: tComentario.BackColor = vbWindowBackground
            cDisponibilidad.Enabled = True: cDisponibilidad.BackColor = vbWindowBackground
        
        Else        'ANTES DEL CIERRE DE LA DISPONIBILIDAD
'            bADM = miConexion.AccesoAlMenu(prmKeyAppADM)
            
            tProveedor.Enabled = True: tProveedor.BackColor = Colores.Obligatorio
            tSerie.Enabled = True: tSerie.BackColor = vbWindowBackground
            tNumero.Enabled = True: tNumero.BackColor = vbWindowBackground
            tIva.Enabled = True: tIva.BackColor = vbWindowBackground
            tCofis.Enabled = True: tCofis.BackColor = vbWindowBackground
            tTCDolar.Enabled = True: tTCDolar.BackColor = Colores.Obligatorio
            
'            If bADM Then
                tRubro.Enabled = True: tRubro.BackColor = vbWindowBackground
                tSubRubro.Enabled = True: tSubRubro.BackColor = vbWindowBackground
'            Else
'                bSplitR.Enabled = False
'            End If
            tComentario.Enabled = True: tComentario.BackColor = vbWindowBackground
        End If
        
        'Si se pagó con una Disp con moneda dif a la del gasto o con varias ...
        '1) No dejo modificar Disponibilidad, Comprobante y Monedas
        If mData.cndPagoConOtras Or (mData.oFechaCompra <= mData.cndFCierreDisponibilidad) Then
            cDisponibilidad.Enabled = False: cDisponibilidad.BackColor = Colores.Inactivo
            cMoneda.Enabled = False: cMoneda.BackColor = Colores.Inactivo
            cComprobante.Enabled = False: cComprobante.BackColor = Colores.Inactivo
        End If
    
        If mData.cndHayRelCompraPago Then
            If tProveedor.Enabled Then tProveedor.Enabled = False: tProveedor.BackColor = Colores.Inactivo
            If cMoneda.Enabled Then cMoneda.Enabled = False: cMoneda.BackColor = Colores.Inactivo
            cComprobante.Enabled = False: cComprobante.BackColor = Colores.Inactivo
        End If
        
    Else
        'Accion Nuevo   ----------------------------------------------------------------------------
        dFecha.Enabled = True
        
        cDisponibilidad.Enabled = True: cDisponibilidad.BackColor = vbWindowBackground
        cComprobante.Enabled = True: cComprobante.BackColor = Obligatorio
        cMoneda.Enabled = True: cMoneda.BackColor = Obligatorio
        
        tIOriginal.Enabled = True: tIOriginal.BackColor = Obligatorio
        tIva.Enabled = True: tIva.BackColor = vbWindowBackground
        tCofis.Enabled = True: tCofis.BackColor = vbWindowBackground
        tTCDolar.Enabled = True: tTCDolar.BackColor = Obligatorio
    
        tProveedor.BackColor = Obligatorio
        tSerie.Enabled = True: tSerie.BackColor = vbWindowBackground
        tNumero.Enabled = True: tNumero.BackColor = vbWindowBackground
            
        tRubro.Enabled = True: tRubro.BackColor = vbWindowBackground
        tSubRubro.Enabled = True: tSubRubro.BackColor = vbWindowBackground
                
        tComentario.Enabled = True: tComentario.BackColor = vbWindowBackground
    End If
    
    If sNuevo Then
        tAutoriza.Enabled = True: tAutoriza.BackColor = Obligatorio
        chVerificado.Enabled = False: chVerificado.Value = vbGrayed
    Else
        If Val(tAutoriza.Tag) <> 0 Then
            tAutoriza.Enabled = True: tAutoriza.BackColor = Obligatorio
            If Val(tAutoriza.Tag) = paCodigoDeUsuario Then chVerificado.Enabled = True
        End If
    End If
    
    'tUsuario.Enabled = True: tUsuario.BackColor = vbWindowBackground
    
End Sub

Private Sub LimpioFicha()
    
    InicializoMData
    
    tID.Text = ""
    dFecha.Value = gFechaServidor
    tProveedor.Text = ""
    cComprobante.Text = "": tNumero.Text = "": tSerie.Text = ""
    cMoneda.Text = "": tIOriginal.Text = "": tIva.Text = ""
    tCofis.Text = ""
    
    tTCDolar.Text = "": lTC.Caption = ""
    
    tRubro.Text = "": tSubRubro.Text = ""
    
    cDisponibilidad.Text = "": cDisponibilidad.Tag = 0
    orCheque.fnc_BlankControls
    orCheque.Tag = 0
    
    tComentario.Text = ""
    lTotalGasto.Caption = ""
    tAutoriza.Text = "": chVerificado.Value = vbGrayed
 '   tUsuario.UserID = 0
    lModificado.Caption = ""
    
    ReDim arrRubros(0)
    arrRubros(0).IdRubro = 0
    
End Sub

Private Sub tProveedor_Change()
    On Error Resume Next
    tProveedor.Tag = 0
    
    If sNuevo Then
        tTCDolar.Tag = "": tTCDolar.Text = "": lTC.Caption = ""
        tSubRubro.Text = "": tRubro.Text = ""
        cMoneda.Text = ""
    End If
    
End Sub

Private Sub tProveedor_GotFocus()
    tProveedor.SelStart = 0: tProveedor.SelLength = Len(tProveedor.Text)
End Sub

Private Sub tProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF1:
            If sNuevo Or sModificar Then Exit Sub
            If Val(tProveedor.Tag) <> 0 Then AccionListaDeAyuda
        
    End Select
    
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then
            If tID.Enabled Then
                Foco tID
            ElseIf cComprobante.Enabled Then Foco cComprobante
            Else
                Foco tSerie
            End If
            Exit Sub
        End If
        Screen.MousePointer = 11
        tProveedor.Text = Replace(tProveedor.Text, " ", "%")
        Dim aQ As Long, aIdProveedor As Long, aTexto As String
        aQ = 0
        Cons = "Select PClCodigo, PClFantasia as Nombre, PClNombre as 'Razón Social' from ProveedorCliente " _
                & " Where PClNombre like '" & Trim(tProveedor.Text) & "%' Or PClFantasia like '" & Trim(tProveedor.Text) & "%'"
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            aQ = 1: aIdProveedor = rsAux!PClCodigo: aTexto = Trim(rsAux!Nombre)
            rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
        End If
        rsAux.Close
        
        Select Case aQ
            Case 0:
                    MsgBox "No existe una empresa para el con el nombre ingresado.", vbExclamation, "No existe Empresa"
            Case 1:
                    tProveedor.Text = aTexto: tProveedor.Tag = aIdProveedor
        
            Case 2:
                    Dim aLista As New clsListadeAyuda, mID As Long
                    mID = aLista.ActivarAyuda(cBase, Cons, 5500, 1, "Proveedores")
                    If mID <> 0 Then
                        tProveedor.Text = Trim(aLista.RetornoDatoSeleccionado(1))
                        tProveedor.Tag = aLista.RetornoDatoSeleccionado(0)
                    End If
                    Set aLista = Nothing
        End Select
        
        If Val(tProveedor.Tag) <> 0 Then
            If sNuevo Then CargoValoresProveedor CLng(tProveedor.Tag)
            If cComprobante.Enabled Then Foco cComprobante Else Foco tSerie
        End If
        
        Screen.MousePointer = 0
    End If
    
    Exit Sub
 
errBuscar:
    clsGeneral.OcurrioError "Error al procesar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tRubro_GotFocus()
    tRubro.SelStart = 0: tRubro.SelLength = Len(tRubro.Text)
End Sub

Private Sub tRubro_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyDivide
                If bSplitR.Enabled And (sNuevo Or sModificar) Then Call bSplitR_Click
                KeyCode = 0
    End Select
    
End Sub

Private Sub tRubro_KeyPress(KeyAscii As Integer)
On Error GoTo errBS
    
    If Chr(KeyAscii) = "/" Then KeyAscii = 0
    
    If KeyAscii = vbKeyReturn Then
        If Not tRubro.Locked Then
            If Val(tRubro.Tag) <> 0 Then Foco tSubRubro: Exit Sub
            If Trim(tRubro.Text) = "" Then Foco tSubRubro: Exit Sub
        
            ing_BuscoRubro tRubro
            
            Foco tSubRubro
            Exit Sub
        End If
        
        If cDisponibilidad.Enabled Then Foco cDisponibilidad Else Foco tComentario
    End If
    Exit Sub

errBS:
    clsGeneral.OcurrioError "Error al buscar el rubro.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tSerie_GotFocus()
    tSerie.SelStart = 0: tSerie.SelLength = Len(tSerie.Text)
End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then Foco tNumero
End Sub

Private Sub tSubRubro_Change()
    tSubRubro.Tag = 0
End Sub

Private Sub tSubRubro_GotFocus()
    tSubRubro.SelStart = 0: tSubRubro.SelLength = Len(tSubRubro.Text)
End Sub

Private Sub tSubRubro_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyDivide
                If bSplitR.Enabled And (sNuevo Or sModificar) Then Call bSplitR_Click
                KeyCode = 0
    End Select
    
End Sub

Private Sub tSubRubro_KeyPress(KeyAscii As Integer)
On Error GoTo errBS
    
    If Chr(KeyAscii) = "/" Then KeyAscii = 0
    
    If KeyAscii = vbKeyReturn Then
        If Not tSubRubro.Locked Then
            If Trim(lTotalGasto.Caption) = "" Then Foco tIOriginal: Exit Sub
            If Not IsNumeric(lTotalGasto.Caption) Then Foco tIOriginal: Exit Sub
            
            If Trim(tSubRubro.Text) <> "" And Val(tSubRubro.Tag) = 0 Then
                ing_BuscoSubrubro tRubro, tSubRubro
            End If
            
            If Val(tSubRubro.Tag) = 0 Then Exit Sub
                
                Dim aImp As Currency: aImp = 0
                If Val(tIOriginal.Text) <> 0 Then aImp = aImp + CCur(tIOriginal.Text)
                If Val(tIva.Text) <> 0 Then aImp = aImp - CCur(tIva.Text)
                If Val(tCofis.Text) <> 0 Then aImp = aImp - CCur(tCofis.Text)
            
                If Abs(CCur(lTotalGasto.Caption)) <> Abs(aImp) Then
                    MsgBox "El total Bruto - I.V.A - Cofis, no es igual al total Neto del gasto. Verifique.", vbExclamation, "Error de Ingreso"
                    Exit Sub
                End If
        
                'Agregoel Gasto Al array----------------------------------------
                ReDim arrRubros(0)
                With arrRubros(0)
                    .TextoRubro = Trim(tRubro.Text)
                    .IdRubro = Val(tRubro.Tag)
                    
                    .TextoSRubro = Trim(tSubRubro.Text)
                    .IdSRubro = Val(tSubRubro.Tag)
                    
                    .Importe = Format(lTotalGasto.Caption, FormatoMonedaP)
                End With
                
        End If
        
        If cComprobante.ListIndex = -1 Then cComprobante.SetFocus: Exit Sub
        If cComprobante.ItemData(cComprobante.ListIndex) <> TipoDocumento.CompraCredito And _
            cComprobante.ItemData(cComprobante.ListIndex) <> TipoDocumento.CompraNotaCredito Then
                If cDisponibilidad.Enabled Then Foco cDisponibilidad Else Foco tComentario
        Else
                Foco tComentario
        End If
        
    End If
    Exit Sub

errBS:
    clsGeneral.OcurrioError "Error al buscar el subrubro.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tTCDolar_Change()
    lTC.Caption = "manual"
End Sub

Private Sub tTCDolar_GotFocus()
    tTCDolar.SelStart = 0: tTCDolar.SelLength = Len(tTCDolar.Text)
End Sub

Private Sub tTCDolar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And tTCDolar.Enabled Then Foco tRubro
End Sub

Private Sub CargoDatosCombos()

    On Error Resume Next
    
    Cons = "Select MonCodigo, MonSigno from Moneda Where MonCodigo In (" & paMonedaDolar & ", " & paMonedaPesos & ")"
    CargoCombo Cons, cMoneda
        
    'Cargo los valores para los comprobantes de pago
    cComprobante.Clear
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraContado)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraContado
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraCredito)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraCredito
    
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraEntradaCaja)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraEntradaCaja
    
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaCredito)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraNotaCredito
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaDevolucion)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraNotaDevolucion
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraRecibo)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraRecibo
    
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraSalidaCaja)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraSalidaCaja
        
End Sub

Private Sub CargoValoresProveedor(idProveedor As Long)
    
    On Error GoTo errCargo
    
    Screen.MousePointer = 11
    'Cargo Valores por defecto del Proveedor   ----------------------------------------
    Cons = "Select * from EmpresaDato " _
           & " Where EDaTipoEmpresa = " & TipoEmpresa.Cliente _
           & " And EDaCodigo = " & idProveedor
    Set rsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        tTCDolar.Tag = rsAux!EDaTCAnterior
    End If
    rsAux.Close

    'Cargo Valores por defecto del ultimo Gasto   ----------------------------------------
    Dim mIDR As Long, mIDSR As Long
    Dim mNombreR As String, mNombreSR As String
    Dim mIDCompraA As Long
    mIDR = 0

    Cons = "Select Top 1 ComCodigo, ComMoneda, ComTipoDocumento, ComSerie, SubRubro.*, Rubro.* " & _
                " From Compra, GastoSubrubro, SubRubro, Rubro " & _
                " Where ComProveedor = " & idProveedor & _
                " And ComCodigo = GSrIDCompra  " & _
                " And GSrIDSubrubro = SRuID And SRuRubro = RubID" & _
                " And ComTipoDocumento Not In ( " & _
                    TipoDocumento.CompraReciboDePago & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ")"
                
    Cons = Cons & " Order by ComCodigo Desc"
    
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        mIDCompraA = rsAux!ComCodigo
        mNombreR = Trim(rsAux!RubNombre)
        mIDR = rsAux!SRuRubro
        mNombreSR = Trim(rsAux!SRuNombre)
        mIDSR = rsAux!SRuID
            
        BuscoCodigoEnCombo cMoneda, rsAux!ComMoneda
        BuscoCodigoEnCombo cComprobante, rsAux!ComTipoDocumento
        If Not IsNull(rsAux!ComSerie) Then tSerie.Text = Trim(rsAux!ComSerie)
        
        rsAux.MoveNext
        If Not rsAux.EOF Then mIDR = 0
    End If
    rsAux.Close
        
    tRubro.Text = mNombreR: tRubro.Tag = mIDR
    tSubRubro.Text = mNombreSR: tSubRubro.Tag = mIDSR
    
    
    If cMoneda.ListIndex <> -1 Then     'Busco Pago Anterior
        dis_CargoDisponibilidades cDisponibilidad, cMoneda.ItemData(cMoneda.ListIndex)
        Cons = "Select * from MovimientoDisponibilidad, MovimientoDisponibilidadRenglon " & _
                    " Where MDiId = MDRIDMovimiento " & _
                    " And MDiIdCompra = " & mIDCompraA
        mIDCompraA = 0
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            mIDCompraA = rsAux!MDRIdDisponibilidad
            rsAux.MoveNext: If Not rsAux.EOF Then mIDCompraA = 0
        End If
        rsAux.Close
        If mIDCompraA > 0 Then BuscoCodigoEnCombo cDisponibilidad, mIDCompraA
        
    End If
    Screen.MousePointer = 0

Exit Sub
errCargo:
    clsGeneral.OcurrioError "Error al cargar los valores del proveedor.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValidoCampos() As Boolean

Dim aTotal As Currency
    
    On Error GoTo errValido
    ValidoCampos = False
    
'    If tUsuario.UserID = 0 Then
'        MsgBox "Falta ingresar el usuario.", vbExclamation, "Ingrese su Usuario"
'        tUsuario.SetFocus: Exit Function
'    End If
    
    If Not IsDate(dFecha.Value) Then
        MsgBox "La fecha ingresada para el registro del gasto no es correcta.", vbExclamation, "ATENCIÓN"
        dFecha.SetFocus: Exit Function
    End If
    If Val(tProveedor.Tag) = 0 Then
        MsgBox "Debe seleccionar el proveedor del gasto.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Function
    End If
    If cComprobante.ListIndex = -1 Then
        MsgBox "Debe seleccionar el comprobante para el registro del gasto.", vbExclamation, "ATENCIÓN"
        Foco cComprobante: Exit Function
    End If
    If Trim(tNumero.Text) <> "" Then
        If Not IsNumeric(tNumero.Text) Then
            MsgBox "Debe ingresar la numeración del comprobante.", vbExclamation, "ATENCIÓN"
            Foco tNumero: Exit Function
        End If
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda para el registro del gasto.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    If Not IsNumeric(tIOriginal.Text) Then
        MsgBox "Debe ingresar el importe total del gasto.", vbExclamation, "ATENCIÓN"
        Foco tIOriginal: Exit Function
    End If
    
    If Not IsNumeric(tTCDolar.Text) Then
        MsgBox "Debe ingresar el valor del dólar para la fecha ingresada (tasa de cambio).", vbExclamation, "ATENCIÓN"
        Foco tNumero: Exit Function
    End If
        
    If arrRubros(0).IdRubro = 0 Then
        MsgBox "Debe ingresar los rubros a los que va al gasto.", vbExclamation, "Falta Asignar Rubro"
        Foco tSubRubro: Exit Function
    End If
    
    If Not tSubRubro.Locked And (tSubRubro.Tag) = 0 Then
        MsgBox "Debe ingresar el rubro al que va al gasto.", vbExclamation, "Falta Asignar Rubro"
        Foco tSubRubro: Exit Function
    End If
    
    'Valido importe de los gastos contra el importe original
    aTotal = 0
    For I = LBound(arrRubros) To UBound(arrRubros)
        aTotal = aTotal + arrRubros(I).Importe
    Next
    If Abs(aTotal) <> Abs(CCur(lTotalGasto.Caption)) Then
        MsgBox "El importe del gasto (" & lTotalGasto.Caption & ") no coincide con la suma de los rubros (" & Format(aTotal, FormatoMonedaP) & ").", vbExclamation, "Diferencia en Asignación"
        Foco tIOriginal: Exit Function
    End If
    
    'Chequeo los importes, suma Total   ----
    Dim mTotalG As Currency
    mTotalG = CCur(lTotalGasto.Caption)
    If IsNumeric(tIva.Text) Then mTotalG = mTotalG + CCur(tIva.Text)
    If IsNumeric(tCofis.Text) Then mTotalG = mTotalG + CCur(tCofis.Text)
    If CCur(tIOriginal.Text) <> mTotalG Then
        MsgBox "El importe del gasto no coincide con el Neto + Impuestos.", vbExclamation, "Posible Error"
        Foco tIOriginal: Exit Function
    End If
                    
    If Val(tAutoriza.Tag) = 0 And tAutoriza.Enabled Then
        MsgBox "Debe ingresar el usuario que autoriza el ingreso del gasto.", vbExclamation, "Falta Usuario que Autoriza el Gasto"
        Foco tAutoriza: Exit Function
    End If
    
    Dim mTipoCG As Integer
    mTipoCG = cComprobante.ItemData(cComprobante.ListIndex)
    If mTipoCG = TipoDocumento.CompraContado Or mTipoCG = TipoDocumento.CompraCredito Or _
        mTipoCG = TipoDocumento.CompraNotaCredito Or mTipoCG = TipoDocumento.CompraNotaDevolucion Then  'TipoDocumento.CompraRecibo
        
        If Trim(tSerie.Text) = "" Or Trim(tNumero.Text) = "" Then
            MsgBox "Para el tipo de comprobante seleccionado se deben ingresar la serie y el número.", vbExclamation, "Faltan Datos"
            Foco tSerie: Exit Function
        End If
    End If
    
    If mTipoCG = TipoDocumento.CompraCredito Then
        cDisponibilidad.Text = "": cDisponibilidad.Tag = 0
        orCheque.fnc_BlankControls
    End If
    
    'Controlo la fecha del Gasto -------------------------------------------------------------------------------------------------
    If dFecha.Enabled Then
        If cDisponibilidad.ListIndex > 0 Then
            Dim mDate As Date, mDateG As Date
            Dim bCerrada As Boolean, mIdx As Integer
            bCerrada = False
            mDate = dis_FechaCierre(cDisponibilidad.ItemData(cDisponibilidad.ListIndex), dFecha.Value)
            mDateG = dFecha.Value
            
            mIdx = dis_IdxArray(cDisponibilidad.ItemData(cDisponibilidad.ListIndex))
            If arrDisp(mIdx).Bancaria Then
                If orCheque.fnc_GetValorData("") Then   'Hay Cheque, si no a la orden
                    If Trim(orCheque.fnc_GetValorData("CheVencimiento")) <> "" Then
                        mDateG = CDate(orCheque.fnc_GetValorData("CheVencimiento"))
                    Else
                        mDateG = CDate(orCheque.fnc_GetValorData("CheLibrado"))
                    End If
                End If
            End If
            If mDate >= mDateG Then bCerrada = True
            
            If bCerrada Then
                MsgBox "La disponibilidad " & cDisponibilidad.Text & " está cerrada." & vbCrLf & _
                            "No se pueden realizar movimientos para la fecha del gasto.", vbExclamation, "Disponibilidad Cerrada"
                Exit Function
            End If
        End If
        
        If prmFCierreIVA >= dFecha.Value Then
            MsgBox "El pago de impuestos está cerrado al " & Format(prmFCierreIVA, "d/mm/yyyy") & "." & vbCrLf & _
                        "No se pueden ingresar gastos menores a esa fecha.", vbExclamation, "Pago de Impuestos Cerrado"
            Exit Function
        End If
    End If
    
    '-----------------------------------------------------------------------------------------------------------------------------------
    
    If sModificar Then
        If (mData.oTotalBruto <> CCur(tIOriginal.Text)) Then
            If mData.cndHayRelCompraPago Or mData.cndPagoConOtras Then
                MsgBox "El importe Total del Gasto no se puede modificar debido a: " & vbCrLf & vbCrLf & _
                            "1) Tiene pagos asignados (por notas o recibos)." & vbCrLf & _
                            "2) Está asignado a una factura (saldando parte de ella), como recibo o nota." & vbCrLf & _
                            "3) Está pago con otras disponibilidades.", vbInformation, "Control de Datos"
                Foco tIOriginal: Exit Function
            End If
        End If
    End If
    
    If mTipoCG <> TipoDocumento.CompraCredito And mTipoCG <> TipoDocumento.CompraNotaCredito Then
        If cDisponibilidad.ListIndex = -1 Then
            MsgBox "Para el tipo de comprobante seleccionado se debe ingresar con que disponibilidad se paga.", vbExclamation, "Falta con Que se Paga"
            Foco cDisponibilidad: Exit Function
        End If
    End If
    
    'Cargo Estructura mData con valores para usar al grabar ---------------------------------------------------------------------------
    With mData
        .ImporteCompra = CCur(tIOriginal.Text)
        .ImporteDisponibilidad = CCur(tIOriginal.Text)
        .ImportePesos = 0
        
        .Disponibilidad = -1
        If cDisponibilidad.ListIndex <> -1 Then .Disponibilidad = cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
        
        If .Disponibilidad <> -1 And .Disponibilidad <> 0 Then      '-1= Nada; 0= Split
            Dim mMonedaG As Integer    'Voy a sacar el importe en pesos
            mMonedaG = cMoneda.ItemData(cMoneda.ListIndex)
        
            If mMonedaG = paMonedaPesos Then
                .ImportePesos = .ImporteCompra
            Else
                
                If mMonedaG = paMonedaDolar Then
                    .ImportePesos = .ImporteCompra * CCur(tTCDolar.Text)
                Else
                    .ImportePesos = TasadeCambio(mMonedaG, paMonedaPesos, dFecha.Value, "")
                    .ImportePesos = .ImporteDisponibilidad * .ImportePesos
                End If
                
            End If
        End If

        .ImporteCompra = Abs(.ImporteCompra)
        .ImporteDisponibilidad = Abs(.ImporteDisponibilidad)
        .ImportePesos = Abs(.ImportePesos)

        
        Select Case cComprobante.ItemData(cComprobante.ListIndex)
            Case TipoDocumento.CompraNotaCredito, TipoDocumento.CompraNotaDevolucion, _
                    TipoDocumento.CompraEntradaCaja: .HaceSalidaCaja = False
                    
            Case Else: .HaceSalidaCaja = True
        End Select
    
    End With
    '------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Valido los datos del Cheque 'Importe del Gasto Contra Importe Disponible
    If orCheque.fnc_GetValorData("") Then
        Dim mIDC As Long, mAsignado As Currency
        mIDC = orCheque.fnc_GetValorData("CheID")
        mAsignado = orCheque.fnc_GetValorData("CheImporte")
        
        If mIDC <> 0 Then
            mAsignado = mAsignado - dis_ImporteAsignadoCheque(mIDC, prmIdCompra)
        End If
        If CCur(tIOriginal.Text) > mAsignado Then
            MsgBox "El valor del gasto no debe superar el importe disponible del cheque.", vbInformation, "Importe Gasto > al del Cheque"
            Foco tIOriginal: Exit Function
        End If
    End If
            
    dSuceso.Tipo = 0
    If sModificar Then
    
        Dim mDiff As Currency
        'Cambio en Importe      cambios de importe > a 1 $
        If cMoneda.ItemData(cMoneda.ListIndex) = paMonedaPesos Then
            mDiff = mData.oPesos - CCur(tIOriginal.Text)
        Else
            mDiff = mData.oPesos - (CCur(tIOriginal.Text) * CCur(tTCDolar.Text))
        End If
            
        If Abs(mDiff) > 0.99 Then mData.flgSucesoXMod = True
        If Not mData.flgSucesoXMod Then
            If mData.oProveedor <> Val(tProveedor.Tag) Then mData.flgSucesoXMod = True
        End If
        If Not mData.flgSucesoXMod Then
            If mData.oSubRubro <> Val(tSubRubro.Tag) Then mData.flgSucesoXMod = True
        End If
        If Not mData.flgSucesoXMod Then
            If mData.oUsuario <> paCodigoDeUsuario Then mData.flgSucesoXMod = True
        End If
        
        If mData.flgSucesoXMod Then
            If Not zPidoSuceso(prmSucesoModGastos, "Modificación de Gastos") Then Exit Function
                        
            If Abs(mDiff) > 0.99 Then dSuceso.Defensa = "Var. Importe $ " & Format(mDiff, "0.00") & vbCrLf & Trim(dSuceso.Defensa)
        End If
    End If
    
    ValidoCampos = True
    Exit Function

errValido:
    clsGeneral.OcurrioError "Error al validar los datos.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function ValidoDocumento() As Boolean

Dim bMsg As Boolean: bMsg = False

    On Error Resume Next
    ValidoDocumento = False
    
    If UCase(Trim(tProveedor.Text)) = "ND" Or UCase(Trim(tProveedor.Text)) = "N/D" Then
        ValidoDocumento = True: Exit Function
    End If
    
    Cons = "Select * from Compra Where ComCodigo <> " & prmIdCompra
           
    If Trim(tNumero.Text) <> "" Then Cons = Cons & " And ComNumero = " & Trim(tNumero.Text)
    
    Cons = Cons & " And ComProveedor = " & Val(tProveedor.Tag) _
                       & " And ComMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                       & " And ComImporte = " & CCur(lTotalGasto.Caption)
           
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If IsNull(rsAux!ComNumero) Then
            If Format(rsAux!ComFecha, sqlFormatoF) = Format(dFecha.Value, sqlFormatoF) Then bMsg = True
        Else
            bMsg = True
        End If
        If bMsg Then
            Screen.MousePointer = 0
            If MsgBox("Existen gastos ingresados con el mismo documento y proveedor." & Chr(vbKeyReturn) _
                & "Fecha: " & Format(rsAux!ComFecha, "d-mmm yyyy") & Chr(vbKeyReturn) _
                & "Importe: " & Format(rsAux!ComImporte, "##,##0.00") & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                & "Desea proseguir con el ingreso del gasto.", vbInformation + vbYesNo + vbDefaultButton2, "Gastos Ingresados") = vbNo Then
                    rsAux.Close
                    Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    rsAux.Close
    
    '2do caso, Valido Proveedor y Numeración del comprobante    8/2/2003    -----------------------------------------------------
    Dim mTipoCG As Integer
    mTipoCG = cComprobante.ItemData(cComprobante.ListIndex)
    If mTipoCG = TipoDocumento.CompraContado Or mTipoCG = TipoDocumento.CompraCredito Or _
        mTipoCG = TipoDocumento.CompraNotaCredito Or mTipoCG = TipoDocumento.CompraNotaDevolucion Then
        
        If Trim(tNumero.Text) <> "" Then
           
            Cons = "Select * from Compra " & _
                        " Where ComCodigo <> " & prmIdCompra & _
                        " And ComNumero = " & Trim(tNumero.Text) & _
                        " And ComProveedor = " & Val(tProveedor.Tag) & _
                        " And ComTipoDocumento = " & mTipoCG
            Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then
                
                If MsgBox("La numeración del comprobante ya se ingresó." & vbCrLf _
                            & "Fecha: " & Format(rsAux!ComFecha, "d-mmm yyyy") & vbCrLf _
                            & "Importe: " & Format(rsAux!ComImporte, "##,##0.00") & vbCrLf & vbCrLf _
                            & "Desea proseguir con el ingreso del gasto.", vbInformation + vbYesNo + vbDefaultButton2, "Numeración ya Ingresada") = vbNo Then
                            
                    rsAux.Close
                    Screen.MousePointer = 0: Exit Function
                End If
            End If
            rsAux.Close
        End If
    End If
    '-------------------------------------------------------------------------------------------------------------------------------------------
    
    Screen.MousePointer = 0
    ValidoDocumento = True
                   
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Controla los datos de la tabla CompraVencimiento
'       -> Cambios de importes
'       -> Cambio de tipo de documento (de credito a otro)
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ValidoIngresoDeVencimientos() As Boolean

Dim aImporteI, aImporteBD As Currency

    On Error GoTo errValidar
    ValidoIngresoDeVencimientos = False
    
    'Valido los campos de la tabla vencimiento-------------------------------------------------------------------------------------------
    Cons = "Select * from CompraVencimiento Where CVeIdCompra = " & prmIdCompra
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then rsAux.Close: ValidoIngresoDeVencimientos = True: Exit Function
    rsAux.Close
    
    '----> Si hay vencimientos ingresados quiere decir que el comprobante era crédito.......Hay que ver si cambió
    If TipoDocumento.CompraCredito <> cComprobante.ItemData(cComprobante.ListIndex) Then
        If MsgBox("El tipo comprobante ha cambiado a " & Trim(cComprobante.Text) & ". Desea continuar con la modificación." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                    & "Si- Eliminar los vencimientos ingresados y continuar." & Chr(vbKeyReturn) _
                    & "No- Cancela la modificación de datos." & Chr(vbKeyReturn) & Chr(vbKeyReturn), vbYesNo + vbQuestion + vbDefaultButton2, "Cambio de Importe") = vbNo Then Exit Function
        'Borro los pagos
        Cons = "Delete CompraVencimiento Where CVeIDCompra = " & prmIdCompra
        cBase.Execute Cons
        ValidoIngresoDeVencimientos = True: Exit Function
    End If
    
    Cons = "Select * from Compra Where ComCodigo = " & prmIdCompra
    Set rsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
       
    aImporteBD = rsCom!ComImporte: If Not IsNull(rsCom!ComIva) Then aImporteBD = aImporteBD + rsCom!ComIva
    If Not IsNull(rsCom!ComCofis) Then aImporteBD = aImporteBD + rsCom!ComCofis
    aImporteI = CCur(tIOriginal.Text)
    rsCom.Close
    
    If aImporteI <> aImporteBD Then
        If MsgBox("El importe del comprobante ha cambiado. Desea continuar con la modificación." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                    & "Si- Eliminar los vencimientos ingresados y continuar." & Chr(vbKeyReturn) _
                    & "No- Cancela la modificación de datos." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                    & "(*) Si continúa recuerde ingresar los nuevos vencimientos.", vbYesNo + vbQuestion + vbDefaultButton2, "Cambio de Importe") = vbNo Then Exit Function
        
        'Borro los pagos
        Cons = "Delete CompraVencimiento Where CVeIDCompra = " & prmIdCompra
        cBase.Execute Cons
        
        bIngresarVencimientos = True
    End If
             
    ValidoIngresoDeVencimientos = True
    Exit Function

errValidar:
    clsGeneral.OcurrioError "Error al validar los vencimientos de pagos.", Err.Description
End Function

Private Function ValidoDatosEliminar() As Boolean

    On Error GoTo errValidar
    ValidoDatosEliminar = False
    
    'Valido los campos de la tabla vencimiento-------------------------------------------------------------------------------------------
    Cons = "Select * from CompraVencimiento Where CVeIdCompra = " & prmIdCompra
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        '----> Si hay vencimientos ingresados
        If MsgBox("El comprobante seleccionado tiene vencimientos ingresados. Desea eliminarlos." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                    & "Si- Eliminar los vencimientos ingresados y continuar." & Chr(vbKeyReturn) _
                    & "No- Cancela la eliminación de datos." & Chr(vbKeyReturn) & Chr(vbKeyReturn), vbYesNo + vbQuestion + vbDefaultButton2, "Vencimientos Ingresados") = vbNo Then rsAux.Close: Exit Function
        'Borro los pagos
        Cons = "Delete CompraVencimiento Where CVeIDCompra = " & prmIdCompra
        cBase.Execute Cons
    End If
    rsAux.Close
    
    'Valido los campos de la tabla CompraPago-------------------------------------------------------------------------------------------
    Cons = "Select * from CompraPago Where CPaDocASaldar = " & prmIdCompra
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        '----> Si hay vencimientos pagos
        MsgBox "El comprobante seleccionado tiene pagos asignados (por notas o recibos)." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                    & "No podrá eliminar el gasto.", vbInformation, "Pagos Ingresados"
        rsAux.Close: Exit Function
    End If
    rsAux.Close

    'Valido los campos de la tabla CompraPago-------------------------------------------------------------------------------------------
    Cons = "Select * from CompraPago Where CPaDocQSalda = " & prmIdCompra
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        '----> Si hay vencimientos pagos
        MsgBox "El comprobante seleccionado está asignado a una factura (saldando parte de ella)." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                    & "Para eliminarlo, primero debe eliminar la relación.", vbInformation, "Pagos Ingresados"
        rsAux.Close: Exit Function
    End If
    rsAux.Close
    
    ValidoDatosEliminar = True
    Exit Function

errValidar:
    clsGeneral.OcurrioError "Error al validar los vencimientos de pagos.", Err.Description
End Function

Private Function ValidoDatosMovimientos(aComprobante As Long, Optional paraEliminar As Boolean = False) As Boolean

    On Error GoTo errValidar
    ValidoDatosMovimientos = True
    
    'Valido los campos de la tabla vencimiento-------------------------------------------------------------------------------------------
    Cons = "Select * from MovimientoDisponibilidad Where MDiIdCompra = " & aComprobante
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        ValidoDatosMovimientos = False
        Screen.MousePointer = 0
        If paraEliminar Then
            MsgBox "Hay movimientos de disponibilidades ingresados para el comprobante." & Chr(vbKeyReturn) & _
                        "Para continuar con la acción debe eliminarlos.", vbInformation, "Movimientos de Disponibilidades"
        End If
    End If
    rsAux.Close
    Exit Function

errValidar:
    clsGeneral.OcurrioError "Error al validar movimientos de disponibilidades.", Err.Description
End Function

Private Function ValidoCompraImportacion() As Boolean
On Error GoTo errValidar
    
    ValidoCompraImportacion = True
    Screen.MousePointer = 11
    
    For I = LBound(arrRubros) To UBound(arrRubros)
        If arrRubros(I).IdRubro = paRubroImportaciones Then
            ValidoCompraImportacion = False
            
            MsgBox "Algunos de los subrubros pertenece al rubro Importaciones." & vbCrLf & _
                        "Para ingresar, modificar o eliminar datos ejecute el Ingreso de Gastos de Importaciones.", vbInformation, "Rubro Importaciones"
                        
            Screen.MousePointer = 0: Exit Function
        End If
    Next
    
    For I = LBound(arrRubros) To UBound(arrRubros)
        If arrRubros(I).IdSRubro = paSubrubroCompraMercaderia Then
            Cons = "Select * from CompraRenglon Where CReCompra = " & prmIdCompra
            Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then ValidoCompraImportacion = False
            rsAux.Close
            
            If Not ValidoCompraImportacion Then
                
                MsgBox "El comprobante seleccionado es una compra de mercadería a proveedores." & vbCrLf & _
                            "Para trabajar con este registro acceda al formulario Compra de Mercadería.", vbInformation, "Compra de Mercadería"
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    Next
    
    Screen.MousePointer = 0
    Exit Function

errValidar:
    clsGeneral.OcurrioError "Error al validar los rubros de importaciones y compras.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function zPidoSuceso(mTipoS As Integer, mTituloS As String) As Boolean

    On Error GoTo errSuceso
    zPidoSuceso = False
    
    Dim objSuceso As New clsSuceso
    objSuceso.TipoSuceso = mTipoS
    objSuceso.ActivoFormulario paCodigoDeUsuario, mTituloS, cBase
    
    Me.Refresh
    With dSuceso
        .Usuario = objSuceso.RetornoValor(Usuario:=True)
        .Defensa = objSuceso.RetornoValor(Defensa:=True)
        .Autoriza = objSuceso.Autoriza
    End With
    
    Set objSuceso = Nothing
    If dSuceso.Usuario = 0 Or Trim(dSuceso.Defensa) = "" Then Exit Function  'Abortó el ingreso del suceso
    
    'Cargo otros datos en la estructura del suceso
    With dSuceso
        .Tipo = mTipoS
        .Titulo = "Gasto (ID:" & Trim(tID.Text) & ") " & Format(dFecha.Value, "dd/mm/yy") & " "
        If Trim(tSerie.Text) <> "" Then .Titulo = .Titulo & Trim(tSerie.Text) & "-"
        If Trim(tNumero.Text) <> "" Then .Titulo = .Titulo & Trim(tNumero.Text)
        .Valor = tIOriginal.Text
        .Cliente = Val(tProveedor.Tag)
    End With
    zPidoSuceso = True
    
    Screen.MousePointer = 0
    Exit Function
errSuceso:
    clsGeneral.OcurrioError "Error al pedir los datos del suceso.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub GraboElPago(mIDCompra As Long)
    
    If mData.oIDMovimiento = 0 And (mData.Disponibilidad = -1 Or mData.Disponibilidad = 0) Then Exit Sub
    
    Dim mIDCheque As Long
    mIDCheque = Val(orCheque.Tag)
    
    '1) Si Antes se hizo movimiento y Ahora no hay o pone otras .... Lo Borro
    If mData.oIDMovimiento <> 0 And (mData.Disponibilidad = -1 Or mData.Disponibilidad = 0) Then
        
        Cons = "Delete MovimientoDisponibilidadRenglon Where MDRIdMovimiento = " & mData.oIDMovimiento
        cBase.Execute Cons
        
        Cons = "Delete MovimientoDisponibilidad Where MDiID = " & mData.oIDMovimiento
        cBase.Execute Cons
        
        If mIDCheque <> 0 Then dis_BorroRelacionCheque mIDCheque, mIDCompra
        Exit Sub
    End If
    
Dim RsMov As rdoResultset
Dim mFechaHora As String
Dim mIDMov As Long
    
    mFechaHora = Format(dFecha.Value, "dd/mm/yyyy") & " " & Format(gFechaServidor, "hh:mm:ss")
    mIDMov = mData.oIDMovimiento
    
    'Inserto en la Tabla Movimiento-Disponibilidad--------------------------------------------------------
    Cons = "Select * from MovimientoDisponibilidad Where MDiID = " & mIDMov
    Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If mIDMov = 0 Then RsMov.AddNew Else RsMov.Edit
    
    RsMov!MDiFecha = Format(mFechaHora, "mm/dd/yyyy")
    RsMov!MDiHora = Format(mFechaHora, "hh:mm:ss")
    RsMov!MDiTipo = paMDPagoDeCompra
    RsMov!MDiIdCompra = mIDCompra
    RsMov!MDiComentario = Trim(tProveedor.Text)
    RsMov.Update: RsMov.Close
    '------------------------------------------------------------------------------------------------------------
    
    'Saco el Id de movimiento-------------------------------------------------------------------------------
    If mIDMov = 0 Then
        Cons = "Select Max(MDiID) from MovimientoDisponibilidad" & _
                  " Where MDiFecha = " & Format(mFechaHora, "'mm/dd/yyyy'") & _
                  " And MDiHora = " & Format(mFechaHora, "'hh:mm:ss'") & _
                  " And MDiTipo = " & paMDPagoDeCompra & _
                  " And MDiIdCompra = " & mIDCompra
        
        Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        mIDMov = RsMov(0)
        RsMov.Close
    End If
    '------------------------------------------------------------------------------------------------------------
    
    'Valido Si hay que ingresar Cheque --------------------------------------------------------------------
    If orCheque.fnc_GetValorData("") Then
        Dim mNewCheque As Long
        mNewCheque = orCheque.fnc_GetValorData("CheID")
        If mNewCheque = 0 Then
            If mIDCheque > 0 Then dis_BorroRelacionCheque mIDCheque, mIDCompra  'Borro rel al cheque anterior
            
            'Agrego el nuevo Cheque
            Cons = "Select * from Cheque Where CheID = 0"
            Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            RsMov.AddNew
            
            RsMov!CheIDDisponibilidad = mData.Disponibilidad
            RsMov!CheSerie = Trim(orCheque.fnc_GetValorData("CheSerie"))
            RsMov!CheNumero = Trim(orCheque.fnc_GetValorData("CheNumero"))
            RsMov!CheImporte = orCheque.fnc_GetValorData("CheImporte")
            RsMov!CheLibrado = Format(orCheque.fnc_GetValorData("CheLibrado"), "mm/dd/yyyy")
            If Trim(orCheque.fnc_GetValorData("CheVencimiento")) <> "" Then RsMov!CheVencimiento = Format(orCheque.fnc_GetValorData("CheVencimiento"), "mm/dd/yyyy")
                    
            RsMov.Update: RsMov.Close
                    
            Cons = "Select Max(CheID) from Cheque" & _
                    " Where CheIDDisponibilidad = " & mData.Disponibilidad
            Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            mIDCheque = RsMov(0)
            RsMov.Close
        
            Cons = "Select * from ChequePago Where CPaIDCheque = " & mIDCheque & " And CPaIDCompra = " & mIDCompra
            Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            RsMov.AddNew
            RsMov!CPaIDCheque = mIDCheque
            RsMov!CPaIDCompra = mIDCompra
            RsMov!CPaImporte = mData.ImporteDisponibilidad
            RsMov.Update: RsMov.Close
        
        Else
            'El nuevo cheque ya existe
            If mNewCheque = mIDCheque Then  'Si es el mismo, solametne actualizo la relacion
                Cons = "Select * from ChequePago " & _
                            " Where CPaIDCheque = " & mIDCheque & " And CPaIDCompra = " & mIDCompra
                Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsMov!CPaImporte <> mData.ImporteDisponibilidad Then
                    RsMov.Edit
                    RsMov!CPaImporte = mData.ImporteDisponibilidad
                    RsMov.Update
                End If
                RsMov.Close
                
            Else        'Es otro Cheque 1) Borro rel al viejo, 2) hago rel al nuevo
                If mIDCheque <> 0 Then dis_BorroRelacionCheque mIDCheque, mIDCompra
                
                Cons = "Select * from ChequePago " & _
                            " Where CPaIDCheque = " & mNewCheque & " And CPaIDCompra = " & mIDCompra
                Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                RsMov.AddNew
                RsMov!CPaIDCheque = mNewCheque
                RsMov!CPaIDCompra = mIDCompra
                RsMov!CPaImporte = mData.ImporteDisponibilidad
                RsMov.Update: RsMov.Close
                
                mIDCheque = mNewCheque
            End If
        End If
    Else
        If mIDCheque > 0 Then dis_BorroRelacionCheque mIDCheque, mIDCompra  'Borro rel al cheque anterior
        mIDCheque = 0
    End If
    '------------------------------------------------------------------------------------------------------------
    
    'Grabo en Tabla Movimiento-Disponibilidad-Renglon--------------------------------------------------
    Cons = "Select * from MovimientoDisponibilidadRenglon Where MDRIdMovimiento = " & mIDMov
    Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsMov.EOF Then RsMov.AddNew Else RsMov.Edit
    
    RsMov!MDRIdMovimiento = mIDMov
    RsMov!MDRIdDisponibilidad = mData.Disponibilidad
    RsMov!MDRIdCheque = mIDCheque
    
    RsMov!MDRImporteCompra = mData.ImporteCompra
    RsMov!MDRImportePesos = mData.ImportePesos
    
    If mData.HaceSalidaCaja Then
        RsMov!MDRHaber = mData.ImporteDisponibilidad
        RsMov!MDRDebe = Null
    Else
        RsMov!MDRDebe = mData.ImporteDisponibilidad
        RsMov!MDRHaber = Null
    End If
    
    RsMov.Update: RsMov.Close

End Sub

Private Function CargoRubrosDelArray()

Dim bLocked As Boolean

    If UBound(arrRubros) = 0 Then
        bLocked = False
        tRubro.Text = arrRubros(0).TextoRubro
        tRubro.Tag = arrRubros(0).IdRubro
        tSubRubro.Text = arrRubros(0).TextoSRubro
        tSubRubro.Tag = arrRubros(0).IdSRubro
    Else
        bLocked = True
        tRubro.Text = "(Varios Rubros)": tRubro.Tag = 0
        tSubRubro.Text = "(Varios Rubros)": tSubRubro.Tag = 0
    End If
    
    tRubro.Locked = bLocked
    tSubRubro.Locked = bLocked
    bSplitR.Enabled = True
    
End Function

'Private Sub tUsuario_AfterDigit()
'    If tUsuario.UserID <> 0 Then AccionGrabar
'End Sub

Private Function InicializoMData()

    With mData
        .cndFCierreDisponibilidad = CDate("1/1/1900")
        .oFechaCompra = CDate("1/1/1900")
        .oIDMovimiento = 0
        .oTotalBruto = 0
        .oPesos = 0
        .oProveedor = 0
        .oSubRubro = 0
        .oUsuario = 0
        
        .Disponibilidad = -1
        .ImporteCompra = 0
        .ImporteDisponibilidad = 0
        .ImportePesos = 0
        .HaceSalidaCaja = False
        
        .cndHayRelCompraPago = False
        .cndPagoConOtras = False
        .cndFCierreDisponibilidad = CDate("1/1/1900")
    
        .flgSucesoXMod = False
    End With
    
End Function

