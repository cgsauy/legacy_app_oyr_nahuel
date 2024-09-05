VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{B443E3A5-0B4D-4B43-B11D-47B68DC130D7}#1.7#0"; "orArticulo.ocx"
Object = "{5EA2D00A-68AC-4888-98E6-53F6035BBEE3}#1.3#0"; "CGSABuscarCliente.ocx"
Begin VB.Form FacVtaTelefonica 
   BackColor       =   &H00C0CAAA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contados a Cobrar En Domicilio"
   ClientHeight    =   5130
   ClientLeft      =   2025
   ClientTop       =   2340
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FacVtaTelefonica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8610
   Begin VB.Timer tmArticuloLimitado 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   2400
      Top             =   480
   End
   Begin prjBuscarCliente.ucBuscarCliente txtCliente 
      Height          =   285
      Left            =   960
      TabIndex        =   43
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      Text            =   "_.___.___-_"
      DocumentoCliente=   1
      QueryFind       =   "EXEC [dbo].[prg_BuscarCliente] 0, '', '', '', '', '', '[KeyQuery]', 0, 0, '', '', 7"
      KeyQuery        =   "[KeyQuery]"
      NeedCheckDigit  =   0   'False
      Comportamiento  =   1
   End
   Begin prjFindArticulo.orArticulo tArticulo 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin AACombo99.AACombo cPagaCon 
      Height          =   315
      Left            =   5280
      TabIndex        =   24
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
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
   Begin AACombo99.AACombo cMoneda 
      Height          =   315
      Left            =   7680
      TabIndex        =   9
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
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
   End
   Begin AACombo99.AACombo cTipoTelefono 
      Height          =   315
      Left            =   960
      TabIndex        =   41
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.TextBox tComentarioDocumento 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   100
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   4200
      Width           =   5535
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
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
            ImageIndex      =   8
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
            Style           =   4
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "envio"
            Object.ToolTipText     =   "Formulario de Envío"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "facturar"
            Object.ToolTipText     =   "Facturar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "validar"
            Object.ToolTipText     =   "Validar venta"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "formapago"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "verfactura"
            Object.ToolTipText     =   "Ver detalle factura"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   3600
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tNombreC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Width           =   4995
   End
   Begin VB.CheckBox chNomDireccion 
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   4140
      TabIndex        =   3
      Top             =   1320
      Width           =   195
   End
   Begin VB.ComboBox cDireccion 
      Height          =   315
      Left            =   2460
      TabIndex        =   2
      Text            =   "cDireccion"
      Top             =   1320
      Width           =   1635
   End
   Begin VB.TextBox tInterno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   22
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox tTelefono 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2235
      MaxLength       =   15
      TabIndex        =   21
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox tCodigo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   960
      MaxLength       =   8
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   960
      MaxLength       =   3
      TabIndex        =   28
      Text            =   "888"
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox tCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   13
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox tUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   7320
      MaxLength       =   12
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5160
      MaxLength       =   200
      TabIndex        =   15
      Top             =   2040
      Width           =   2175
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   4875
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8786
            MinWidth        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1799
            MinWidth        =   1482
            Text            =   "F2-Modificar "
            TextSave        =   "F2-Modificar "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1482
            MinWidth        =   1482
            Text            =   "F3-Nuevo"
            TextSave        =   "F3-Nuevo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1482
            MinWidth        =   1482
            Text            =   "F4-Buscar"
            TextSave        =   "F4-Buscar"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1508
            MinWidth        =   2
            Text            =   "F9 - Envío "
            TextSave        =   "F9 - Envío "
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvVenta 
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cant."
         Object.Width           =   1076
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Artículo"
         Object.Width           =   5381
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Comentario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Unitario"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "I.V.A."
         Object.Width           =   1375
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Sub Total"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Específico"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "CantidadPorMayor"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":10E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":13FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":1716
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacVtaTelefonica.frx":1A30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFlete 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   46
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Flete"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6600
      TabIndex        =   45
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label labIVA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   44
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblRutCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "213 025 510 019"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   42
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pago:"
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Llama del:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Có&digo:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label labUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   40
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label53 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&ment.:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "&LISTA"
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuar&io:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6600
      TabIndex        =   38
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Venta"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6600
      TabIndex        =   37
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sub total"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5160
      TabIndex        =   36
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label labSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   35
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label labTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   34
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblTotalCflete 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   33
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblArticulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Artículo"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ca&nt."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Precio Unitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Moneda:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comen&tario"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblIdentificador 
      BackStyle       =   0  'Transparent
      Caption         =   "&C.I.:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblInfoRutCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "&R.U.C.:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom&bre:"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   795
   End
   Begin VB.Label labDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Niagara 2345"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4380
      TabIndex        =   32
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   3975
   End
   Begin VB.Label labFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10-Dic-1998"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   3840
      TabIndex        =   31
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   30
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   855
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   8415
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
      Begin VB.Menu MnuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEnvio 
         Caption         =   "En&vío"
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu MnuFacturar 
         Caption         =   "Facturar"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu MnuValidarVenta 
         Caption         =   "&Validar Venta"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu MnuSalirLine 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVolver 
         Caption         =   "&Cerrar"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu MnuDetalleFactura 
         Caption         =   "&Detalle de Factura"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuVerVisualizacion 
         Caption         =   "&Visualización de Operaciones"
         Shortcut        =   {F12}
      End
      Begin VB.Menu MnuVerL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerInstalacion 
         Caption         =   "Instalaciones"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu MnuPrinter 
      Caption         =   "Impresora"
      Begin VB.Menu MnuPrintConfig 
         Caption         =   "Configurar"
      End
   End
   Begin VB.Menu MnuMoussePersona 
      Caption         =   "&MoussePersona"
      Visible         =   0   'False
      Begin VB.Menu MnuCPersona 
         Caption         =   "Menú Cliente"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuLineaMP1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuNuevoCliente 
         Caption         =   "&Nuevo Cliente            F3"
      End
      Begin VB.Menu MnuFichaCliente 
         Caption         =   "&Ir a ficha de Cliente   F2"
      End
      Begin VB.Menu MnuBuscarPresona 
         Caption         =   "&Buscar Clientes          F4"
      End
      Begin VB.Menu MnuLineaMP2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBuscarCtosPersona 
         Caption         =   "B&uscar Contados"
         Begin VB.Menu MnuCtdoDomAnuPer 
            Caption         =   "&Anulados"
         End
         Begin VB.Menu MnuCtdoDomPdtePer 
            Caption         =   "&Pendientes"
         End
         Begin VB.Menu MnuCtdoDomRealiPer 
            Caption         =   "&Realizados"
         End
      End
      Begin VB.Menu MnuCancelarMP 
         Caption         =   "&Cancelar"
      End
   End
   Begin VB.Menu MnuMousseEmpresa 
      Caption         =   "MousseEmpresa"
      Visible         =   0   'False
      Begin VB.Menu MnuCEmpresa 
         Caption         =   "Menú Empresa"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuLineaME1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuNuevaEmpresa 
         Caption         =   "&Nueva Empresa             F3"
      End
      Begin VB.Menu MnuFichaEmpresa 
         Caption         =   "&Ir a Ficha de Empresa    F2"
      End
      Begin VB.Menu MnuBuscarEmpresa 
         Caption         =   "&Buscar Empresas           F4"
      End
      Begin VB.Menu MnuLineaME2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBuscarCtosEmpresa 
         Caption         =   "B&uscar contados"
         Begin VB.Menu MnuCtdoDomAnuEmp 
            Caption         =   "&Anulados"
         End
         Begin VB.Menu MnuCtdoDomPendEmp 
            Caption         =   "&Pendiente"
         End
         Begin VB.Menu MnuCtdoDomReaEmp 
            Caption         =   "&Realizados"
         End
      End
      Begin VB.Menu MnuCancelarME 
         Caption         =   "&Cancelar"
      End
   End
End
Attribute VB_Name = "FacVtaTelefonica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modificaciones
'   08/01/2007    si elimina --> anulo la instalación.
'   12/1/2007     permito editar ventaOnline confirmada, además de incluirlas en las busquedas.
'........................

Option Explicit
'REGISTRO DE SUCESOS-------------------------------------------------
Private Enum TipoSuceso
    ModificacionDePrecios = 3
    FacturaArticuloInhabilitado = 13
    FacturaCambioNombre = 16
    ClienteNoVender = 23
    DesasociarEspecifico = 31
End Enum
'--------------------------------------------------------------------

'Definicion de Tipos de Pagos de Envio------------------------------------------------------------------------------------
Private Enum TipoPagoEnvio
    PagaAhora = 1
    PagaDomicilio = 2
    FacturaCamión = 3
End Enum

Private Enum TipoCliente
    Cliente = 1
    Empresa = 2
End Enum

Private Enum ePagaCon
    Cheque = 1
    Dolares = 2
    Efectivo = 3
    RedPagos = 4
End Enum

Private itmx As ListItem

Private Const cte_KeyFindDir = "Buscar ......?"

Private Type tRenglonFact
    idArticulo As Long
    CodArticulo As Long
    IDCombo As Long
'    ArtCombo As Long
    Tipo As Long
    Precio As Currency
    PrecioOriginal As Currency
    PrecioBonificacion As Currency
    EsInhabilitado As Boolean
    Especifico As Long
    DescuentoEspecifico As Currency
    CantidadAlXMayor As Integer
    NombreArticulo As String
End Type
Private miRenglon As tRenglonFact

Public prmIDVta As Long
Private idCliente As Long ' , sNomCliente As String
Private gDirFactura As Long

'String.----------------------------------------
Public strCodigoEnvio As String         'Cdo. vuelvo de envio si no graba tengo que borrar los que esten aca.
Private m_Patron As String
Private strArticuloFlete As String
'Booleanas.-------------------------------------
Private sNuevo As Boolean, sModificar As Boolean, sDiferencia As Boolean

'CURRENCY.------------------------------------------
Private cCambio As Currency

'RDO.------------------------------------------
Private RsAuxVta As rdoResultset
'Public tTiposArtsServicio As String     'Trama con todos los tipos que pertenecen a Servicio

Public Property Let prmIDCliente(ByVal lID As Long)
    idCliente = lID
End Property

Private Sub cDireccion_Change()
    If labDireccion.Caption <> "" And cDireccion.ListIndex = -1 Then labDireccion.Caption = ""
End Sub

Private Sub cDireccion_Click()
On Error GoTo errCargar
    If cDireccion.ListIndex <> -1 Then
        If Val(cDireccion.ItemData(cDireccion.ListIndex)) > -1 Then
            Screen.MousePointer = 11
            labDireccion.Caption = ""
            labDireccion.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, cDireccion.ItemData(cDireccion.ListIndex))
            Screen.MousePointer = 0
        Else
            labDireccion.Caption = ""
            cDireccion.SelStart = 0: cDireccion.SelLength = Len(cDireccion.Text)
        End If
    Else
        labDireccion.Caption = ""
    End If
Exit Sub
errCargar:
    Screen.MousePointer = 0
End Sub

Private Sub cDireccion_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If cDireccion.ListIndex = -1 Then
            If Val(cDireccion.ItemData(cDireccion.ListIndex)) = -1 And cte_KeyFindDir <> cDireccion.Text Then
                loc_FindDireccionAuxiliarTexto
            End If
        Else
            chNomDireccion.SetFocus
        End If
    End If
End Sub

Private Sub cMoneda_Change()
    m_Patron = ""
    LimpioRenglon
End Sub

Private Sub cMoneda_Click()
    LimpioRenglon
End Sub

Private Sub cMoneda_GotFocus()

    Foco cMoneda
    Status.Panels(1).Text = " Seleccione una moneda."

End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And txtCliente.Enabled Then Foco txtCliente
End Sub

Private Sub cMoneda_LostFocus()
    cMoneda.SelLength = 0
End Sub

Private Sub cPagaCon_Click()
    
    If cPagaCon.ListIndex > -1 Then
        If cPagaCon.ItemData(cPagaCon.ListIndex) = 4 Then
            Me.BackColor = &HE7EBE6
            Me.Caption = "Contados telefónicos a cobrar por redpagos"
        Else
            Me.BackColor = &HC0CAAA
            Me.Caption = "Contados telefónicos a cobrar en domicilio"
        End If
    Else
        Me.BackColor = &HC0CAAA
        Me.Caption = "Contados telefónicos a cobrar en domicilio"
    End If
    
End Sub

Private Sub cPagaCon_GotFocus()
    Status.Panels(1).Text = " Indique la forma de pago."
    If sNuevo And cPagaCon.ListIndex = -1 Then cPagaCon.ListIndex = 2
End Sub

Private Sub cPagaCon_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        If cPagaCon.ListIndex > -1 Then
            
            If sNuevo And cPagaCon.ItemData(cPagaCon.ListIndex) = ePagaCon.RedPagos Then
                If strCodigoEnvio <> "" Then
                    CambioFormaPagoEnvios
                    MsgBox "Se modificará la forma de pago de los envíos.", vbInformation, "ATENCIÓN"
                    cPagaCon.Enabled = False
                    lblFlete.Caption = "0.00"
                    lblTotalCflete.Caption = labTotal.Caption
                End If
            ElseIf sNuevo Then
                If LCase(Mid(tComentarioDocumento.Text, 1, 5)) = "pago:" Then
                    tComentarioDocumento.Text = Replace(tComentarioDocumento.Text, cPagaCon.Tag, "", , , vbTextCompare)
                    If LCase(Mid(tComentarioDocumento.Text, 1, 5)) = "pago:" Then
                        'no era el mismo comentario.
                        Foco tComentarioDocumento
                        Exit Sub
                    End If
                End If
                Select Case cPagaCon.ListIndex
                    Case 0: cPagaCon.Tag = "Pago: cheque "
                    Case 1: cPagaCon.Tag = "Pago: a T.C. " & TasadeCambio(paMonedaDolar, cMoneda.ItemData(cMoneda.ListIndex), Date, , paTCComME)
                    Case 2: cPagaCon.Tag = "Pago: efectivo "
                End Select
                tComentarioDocumento.Text = cPagaCon.Tag & tComentarioDocumento.Text
            End If
            
        End If
        Foco tComentarioDocumento
    End If
    
End Sub

Private Sub cPagaCon_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub cPagaCon_Validate(Cancel As Boolean)
    If Not cPagaCon.Enabled Then Exit Sub
    If sNuevo And cPagaCon.ItemData(cPagaCon.ListIndex) = ePagaCon.RedPagos And strCodigoEnvio <> "" And cPagaCon.Enabled Then
        CambioFormaPagoEnvios
        cPagaCon.Enabled = False
        lblFlete.Caption = "0.00"
        lblTotalCflete.Caption = labTotal.Caption
    End If
End Sub

Private Sub cTipoTelefono_Click()
    tTelefono.Text = ""
    tInterno.Text = ""
End Sub

Private Sub cTipoTelefono_GotFocus()
    Foco cTipoTelefono
    Status.Panels(1).Text = " Seleccione un tipo de teléfono."
End Sub

Private Sub cTipoTelefono_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(cTipoTelefono.ListIndex) <> -1 Then
            If Trim(tTelefono.Text) = "" Then
                CargoTelefonos txtCliente.Cliente.Codigo, cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
            End If
            Foco tTelefono
        Else
            If sNuevo And cPagaCon.Enabled Then
                Foco cPagaCon
            Else
                Foco tComentarioDocumento
            End If
        End If
    End If
    
End Sub

Private Sub cTipoTelefono_LostFocus()
    cTipoTelefono.SelLength = 0
End Sub

Private Sub chNomDireccion_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then If tArticulo.Enabled Then Foco tArticulo Else Foco txtCliente
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    Me.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errKD
    Select Case KeyCode
        Case vbKeyF2
            If Me.ActiveControl.Name <> "txtCliente" Then txtCliente.EditarCliente
        Case vbKeyF12
            Screen.MousePointer = 11
            EjecutarApp App.Path & "\visualizacion de operaciones.exe", txtCliente.Cliente.Codigo
        Case vbKeyC
            If Shift = vbAltMask And txtCliente.Enabled Then txtCliente.SetFocus
    End Select
    Screen.MousePointer = 0
    Exit Sub
errKD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error inesperado.", Trim(Err.Description)
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    Set tArticulo.Connect = cBase
    tArticulo.KeyQuerySP = "VtaCdo"
    
    Set txtCliente.Connect = cBase
    'EXEC [dbo].[prg_BuscarCliente] 0, '', '', '', '', '', [KeyQuery], 0, 0, '', '', 7
    txtCliente.DocumentoCliente = DC_CI
        
    'Inicializo variables locales y globales
    sNuevo = False: sModificar = False
    strCodigoEnvio = ""
    labFecha.Caption = vbNullString
    dis_CargoArrayMonedas

    LimpioDatosCliente
    LimpioRenglon
    
    With cPagaCon
        .Clear
        .AddItem "Cheque": .ItemData(.NewIndex) = 1
        .AddItem "Dolares": .ItemData(.NewIndex) = 2
        .AddItem "Efectivo": .ItemData(.NewIndex) = 3
        .AddItem "Redpagos": .ItemData(.NewIndex) = 4
    End With
    
    strArticuloFlete = CargoArticulosDeFlete
    
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda
        
    'Cargo los TiposTelefono
    Cons = "Select TTeCodigo, TTeNombre From TipoTelefono Order by TTeNombre"
    CargoCombo Cons, cTipoTelefono, ""
    
    If paMonedaFacturacion > 0 Then BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
    SetearLView lvValores.Grilla + lvValores.FullRow, lvVenta
    LabTotalesEnCero
    ctr_Enabled False
    
    'Inicio Resultset.
'    Cons = "Select * From VentaTelefonica Where VTeCodigo = 0"
'    Set RsVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If idCliente > 0 Then
        AccionNuevo
        txtCliente.CargarControl (idCliente)
    Else
        If prmIDVta > 0 Then
            tCodigo.Text = prmIDVta
            tCodigo_KeyPress vbKeyReturn
        End If
    End If
    Exit Sub
    
ErrLoad:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error inesperado.", Trim(Err.Description)
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.Panels(1).Text = vbNullString
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If strCodigoEnvio <> vbNullString And _
        lvVenta.ListItems.Count > 0 And sNuevo Then
        BorroEnvios
    End If
End Sub

Private Sub BorroEnvios()
On Error GoTo ErrBE
Dim CodEnvios As String
Dim lngCodEnvio As Long
    
    Do While strCodigoEnvio <> ""
        If InStr(1, strCodigoEnvio, ",") > 0 Then
            CodEnvios = Left(strCodigoEnvio, InStr(1, strCodigoEnvio, ","))
            lngCodEnvio = CLng(Left(CodEnvios, InStr(1, CodEnvios, ",") - 1))
            strCodigoEnvio = Right(strCodigoEnvio, Len(strCodigoEnvio) - InStr(1, strCodigoEnvio, ","))
        Else
            lngCodEnvio = CLng(strCodigoEnvio)
            strCodigoEnvio = ""
        End If
        
        cBase.BeginTrans
        On Error GoTo ErrResumo
        Cons = "DELETE RenglonEnvio Where REvEnvio = " & lngCodEnvio
        cBase.Execute (Cons)
        
        Cons = "DELETE Envio Where EnvCodigo = " & lngCodEnvio
        cBase.Execute (Cons)
        cBase.CommitTrans
    Loop
    Exit Sub
    
ErrBE:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error inesperado al intentar la transacción."
ErrResumo:
    Resume Relajo
Relajo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "No se pudo eliminar algunos de los envío.", Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'    RsVta.Close
    CierroConexion
    crCierroEngine
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Sub

Private Sub Label12_Click()
    Foco cMoneda
End Sub

Private Sub Label13_Click()
    Foco tUsuario
    Status.Panels(1).Text = " Ingrese el dígito de usuario."
End Sub

Private Sub Label16_Click()
    Foco cTipoTelefono
End Sub

Private Sub Label17_Click()
    Foco tComentario
End Sub

Private Sub Label4_Click()
    Foco cPagaCon
End Sub

Private Sub Label5_Click()
    Foco tArticulo
End Sub
Private Sub Label53_Click()
    Foco tComentarioDocumento
End Sub
Private Sub Label6_Click()
    Foco tCantidad
End Sub
Private Sub Label7_Click()
    Foco tUnitario
End Sub
Private Sub lvVenta_GotFocus()
On Error Resume Next
    Status.Panels(1).Text = " [Esp] Edita, [+/-] Agrega o Quita, [Supr] Elimina."
    If lvVenta.ListItems Is Nothing Then Exit Sub
    Set lvVenta.DropHighlight = Nothing
    Set lvVenta.DropHighlight = lvVenta.SelectedItem
End Sub

Private Sub lvVenta_ItemClick(ByVal Item As MSComCtlLib.ListItem)
    Set lvVenta.DropHighlight = Nothing
End Sub

Private Sub lvVenta_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrlvKD
   
    If lvVenta.ListItems.Count > 0 Then
        Select Case KeyCode
        
            Case vbKeySpace
                
                LimpioRenglon
                If Mid(lvVenta.SelectedItem.Key, 1, 1) = "X" Then
                    MsgBox "No se pueden editar artículos que pagan envíos.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                ElseIf Val(lvVenta.SelectedItem.SubItems(8)) > 0 Then
                    MsgBox "No se pueden editar artículos específicos.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                Else
                    
                    If Trim(lvVenta.SelectedItem.Tag) <> vbNullString Then
                        If CCur(lvVenta.SelectedItem.Tag) > 0 Then
                            MsgBox "No puede editar un artículo que fue asignado a envíos.", vbExclamation, "ATENCIÓN"
                            Exit Sub
                        End If
                    End If
                    tArticulo.LoadArticulo (Mid(lvVenta.SelectedItem.Key, 2, Len(lvVenta.SelectedItem.Key)))
                    With miRenglon
                        .idArticulo = tArticulo.prm_ArtID
                        .Tipo = tArticulo.GetField("ArtTipo")
                        If IsNull(tArticulo.GetField("ArtHabilitado")) Then
                            .EsInhabilitado = True
                        Else
                            .EsInhabilitado = UCase(tArticulo.GetField("ArtHabilitado")) = "N"
                        End If
                        If tArticulo.GetField("ArtEsCombo") Then
                            .IDCombo = .idArticulo
                        End If
                        .Precio = CCur(lvVenta.SelectedItem.SubItems(5))
                        If Not IsNull(tArticulo.GetField("ArtEnVentaXMayor")) Then .CantidadAlXMayor = tArticulo.GetField("ArtEnVentaXMayor") Else .CantidadAlXMayor = 1
                        .NombreArticulo = Trim(tArticulo.GetField("ArtNombre"))
                    End With
                    
                    tComentario.Text = Trim(lvVenta.SelectedItem.SubItems(2))
                    tCantidad.Text = lvVenta.SelectedItem.Text
                    tCantidad.Tag = lvVenta.SelectedItem.Tag
                    tUnitario.Text = lvVenta.SelectedItem.SubItems(5)
                    tUnitario.Tag = CCur(db_FindPrecioVigente(tArticulo.prm_ArtID))
                    miRenglon.PrecioOriginal = CCur(tUnitario.Tag)
                    
                    Call RestoLabTotales(CCur(lvVenta.SelectedItem.SubItems(5)), CCur(lvVenta.SelectedItem.SubItems(4)))
                    lvVenta.ListItems.Remove lvVenta.SelectedItem.Index
                    MnuEnvio.Enabled = False
                    If lvVenta.ListItems.Count = 0 Then cMoneda.Enabled = True Else: cMoneda.Enabled = False
                    AplicoCantidadLimitadaPorCantidad
                    tCantidad.SetFocus
                End If
                
            Case vbKeyDelete
                'Si es X es un artículo que paga envío.
                'Si el tag es distinto al itmx entonces tiene envíos.
                If Mid(lvVenta.SelectedItem.Key, 1, 1) = "X" Then
                    MsgBox "No se puede eliminar un artículo que paga envíos, debe ir al formulario de envío y eliminar el mismo.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                Else
                    If CCur(lvVenta.SelectedItem.Tag) > 0 Then
                        MsgBox "El artículo seleccionado fue asignado a envíos, elimine el mismo de los envíos.", vbExclamation, "ATENCIÓN"
                        Exit Sub
                    Else
                        
                        If Val(lvVenta.SelectedItem.SubItems(8)) > 0 And Val(tCodigo.Tag) > 0 Then
                            If MsgBox("ATENCIÓN!!!" & vbCrLf & vbCrLf & "Está eliminando el artículo específico de la compra, al confirmar la acción el mismo quedará libre para ser utilizado en otro documento POR MAS QUE UD NO LLEGUE A GRABAR la modificación de la venta." & _
                                vbCrLf & vbCrLf & "¿Confirma quitar de la venta el artículo específico?", vbYesNo + vbQuestion + vbDefaultButton2, "Desasignar artículo específico") = vbYes Then
                                
                                Cons = "SELECT * FROM ArticuloEspecifico WHERE AEsID = " & Val(lvVenta.SelectedItem.SubItems(8)) & " AND AEsDocumento = " & Val(tCodigo.Tag)
                                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                                RsAux.Edit
                                RsAux("AEsTipoDocumento") = Null
                                RsAux("AEsDocumento") = Null
                                RsAux.Update
                                RsAux.Close
                            Else
                                Exit Sub
                            End If
                        End If
                        
                        Call RestoLabTotales(CCur(lvVenta.SelectedItem.SubItems(5)), CCur(lvVenta.SelectedItem.SubItems(4)))
                        lvVenta.ListItems.Remove lvVenta.SelectedItem.Index
                        If lvVenta.ListItems.Count > 0 Then
                            cMoneda.Enabled = False
                            MnuEnvio.Enabled = False
                        Else
                            cMoneda.Enabled = True
                            MnuEnvio.Enabled = True
                        End If
                    End If
                End If
                
            Case vbKeyReturn
                Foco cTipoTelefono
            
            Case vbKeyAdd
                If Mid(lvVenta.SelectedItem.Key, 1, 1) = "X" Then
                    MsgBox "No se pueden agregar artículos que pagan envíos.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                ElseIf Val(lvVenta.SelectedItem.SubItems(8)) > 0 Then
                    MsgBox "No se pueden agregar artículos si es específico.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                Else
                    lvVenta.SelectedItem.Text = CCur(lvVenta.SelectedItem.Text) + 1
                    lvVenta.SelectedItem.SubItems(5) = Format(CCur(lvVenta.SelectedItem.Text) * CCur(lvVenta.SelectedItem.SubItems(3)), "#,##0.00")
                    labIVA.Caption = Format(CCur(labIVA.Caption) + CCur(lvVenta.SelectedItem.SubItems(3)) - (CCur(lvVenta.SelectedItem.SubItems(3)) / CCur(1 + (CCur(lvVenta.SelectedItem.SubItems(4)) / 100))), "#,##0.00")
                    labTotal.Caption = Format(CCur(labTotal.Caption) + CCur(lvVenta.SelectedItem.SubItems(3)), "#,##0.00")
                    labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
                    
                    If (lblFlete.Caption = "") Then lblFlete.Caption = "0.00"
                    lblTotalCflete.Caption = Format(CCur(lblFlete.Caption) + CCur(labTotal.Caption), "#,##0.00")
                    
                    ValidoVentaLimitadaPorFila lvVenta.ListItems(lvVenta.SelectedItem.Index)
                End If
                
            Case vbKeySubtract
                If Mid(lvVenta.SelectedItem.Key, 1, 1) = "X" Then
                    MsgBox "No se pueden restar artículos que pagan envíos.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                ElseIf Val(lvVenta.SelectedItem.SubItems(8)) > 0 Then
                    MsgBox "No se pueden restar artículos si es específico.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                Else
                    If CCur(lvVenta.SelectedItem.Tag) = CCur(lvVenta.SelectedItem.Text) Then
                        MsgBox "No se pueden restar más artículos debido a que están asignados a envíos.", vbExclamation, "ATENCIÓN"
                        Exit Sub
                    Else
                        If CCur(lvVenta.SelectedItem.Text) - 1 > 0 Then
                            lvVenta.SelectedItem.Text = CCur(lvVenta.SelectedItem.Text) - 1
                            
                            ValidoVentaLimitadaPorFila lvVenta.ListItems(lvVenta.SelectedItem.Index)
                            
                            lvVenta.SelectedItem.SubItems(5) = Format(CCur(lvVenta.SelectedItem.Text) * CCur(lvVenta.SelectedItem.SubItems(3)), "#,##0.00")
                            'Aca le resto uno.
                            labIVA.Caption = Format(CCur(labIVA.Caption) - (CCur(lvVenta.SelectedItem.SubItems(3)) - (CCur(lvVenta.SelectedItem.SubItems(3)) / CCur(1 + (CCur(lvVenta.SelectedItem.SubItems(4)) / 100)))), "#,##0.00")
                            labTotal.Caption = Format(CCur(labTotal.Caption) - CCur(lvVenta.SelectedItem.SubItems(3)), "#,##0.00")
                            labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
                            If lblFlete.Caption = "" Then lblFlete.Caption = "0.00"
                            lblTotalCflete.Caption = Format(CCur(lblFlete.Caption) + CCur(labTotal.Caption), "#,##0.00")
                        End If
                    End If
                End If
            
        End Select
    End If
    Exit Sub

ErrlvKD:
    clsGeneral.OcurrioError "Error inesperado.", Trim(Err.Description)
    
End Sub

Private Sub lvVenta_LostFocus()
    Set lvVenta.DropHighlight = Nothing
End Sub

Private Sub AyudaCliente()

    On Error GoTo errAyuda
    Screen.MousePointer = 11
    
    Dim aIdSeleccionado As Long
    Dim objLista As New clsListadeAyuda
    If objLista.ActivarAyuda(cBase, Cons, 8200, 1, "Ayuda de Cliente") Then
        aIdSeleccionado = objLista.RetornoDatoSeleccionado(0)
    End If
    Set objLista = Nothing
    Me.Refresh
        
    If aIdSeleccionado > 0 Then
        If sNuevo Then AccionCancelar
        BuscoVentaTelefonica aIdSeleccionado
    End If
    Screen.MousePointer = 0
    Exit Sub

errAyuda:
    clsGeneral.OcurrioError "Ocurrió un error al procesar la información", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuDetalleFactura_Click()
    AccionVerFactura
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuEnvio_Click()
On Error GoTo ErrME

    Dim itmxEnvio As ListItem
    
    If lvVenta.ListItems.Count > 0 Then
    
        If txtCliente.Cliente.Codigo = 0 Then
            MsgBox "No se pueden ingresar envíos sin seleccionar un cliente.", vbInformation, "ATENCIÓN"
            Exit Sub
        Else
            If strCodigoEnvio = vbNullString Then strCodigoEnvio = "0"
            Dim objEnvio As New clsEnvio
            If sNuevo Then
                Dim idTabla As Integer
                idTabla = NumeroAuxiliarEnvio
                If idTabla = 0 Then MsgBox "Reintente la operación.", vbExclamation, "ATENCIÓN": Exit Sub
                objEnvio.NuevoEnvio cBase, strCodigoEnvio, idTabla, txtCliente.Cliente.Codigo, cMoneda.ItemData(cMoneda.ListIndex), TipoEnvio.Cobranza, CCur(labTotal.Caption)
                For Each itmx In lvVenta.ListItems
                    itmx.Tag = 0
                Next
            Else
                objEnvio.InvocoEnvioXDocumento tCodigo.Text, TipoDocumento.ContadoDomicilio, gPathListados
            End If
            Me.Refresh
            strCodigoEnvio = objEnvio.RetornoEnvios
            Set objEnvio = Nothing
            If sNuevo Then
                CalculoArticulosEnEnvio
                If cPagaCon.ListIndex > -1 Then
                    If cPagaCon.ItemData(cPagaCon.ListIndex) = ePagaCon.RedPagos Then CambioFormaPagoEnvios
                End If
                CargarFletesEnvio
            Else
                BuscoVentaTelefonica tCodigo.Text
            End If
        End If
    End If
    Exit Sub
    
ErrME:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error inesperado.", Err.Description
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuFacturar_Click()
'    AccionFacturar
End Sub

Private Sub AccionLimpiar()

    lblArticulo.Tag = "0": lblArticulo.Caption = "&Artículo"
    tArticulo.KeyQuerySP = "VtaCdo"
    
    LimpioRenglon
    Shape2.BackColor = &HEEEEEE   '&HC0D0FA
    chNomDireccion.BackColor = Shape2.BackColor
    labFecha.Caption = vbNullString
    lvVenta.ListItems.Clear
    LimpioDatosCliente
    LabTotalesEnCero
    labFecha.Tag = ""
    lblCodigo.Tag = ""
    
    With cPagaCon
        .Text = "": .Tag = ""
    End With
    
    With cMoneda
        .Enabled = True
        .BackColor = Blanco
    End With
    
    BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
    With tUsuario
        .Text = vbNullString
        .Tag = vbNullString
    End With
    tComentarioDocumento.Text = vbNullString
    labUsuario.Caption = vbNullString
    tTelefono.Text = vbNullString
    
    MnuEnvio.Enabled = False
    Toolbar1.Buttons("envio").Enabled = False
    MnuValidarVenta.Enabled = False
    Toolbar1.Buttons("validar").Enabled = False
    MnuDetalleFactura.Enabled = False
    Toolbar1.Buttons("verfactura").Enabled = False
    MnuFacturar.Enabled = False
    Toolbar1.Buttons("facturar").Enabled = False
    Botones True, False, False, False, False, Toolbar1, Me
    Toolbar1.Buttons("formapago").Enabled = False
    
    'Siempre limpio al color de la vta telefónica.
    Me.BackColor = &HC0CAAA
    Me.Caption = "Contados telefónicos a cobrar en domicilio"
    
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

Private Sub MnuPrintConfig_Click()
    prj_LoadConfigPrint True
End Sub

Private Sub MnuValidarVenta_Click()
   AccionValidar
End Sub

Private Sub MnuVerInstalacion_Click()
On Error Resume Next
    If Val(tCodigo.Text) > 0 Then EjecutarApp App.Path & "\Instaires.exe", "VTe:" & tCodigo.Text
End Sub

Private Sub MnuVerVisualizacion_Click()
    On Error Resume Next
    If txtCliente.Cliente.Codigo > 0 Then
        EjecutarApp App.Path & "\Visualizacion de operaciones", txtCliente.Cliente.Codigo
    End If
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub tArticulo_Change()
    miRenglon.idArticulo = 0
    tmArticuloLimitado.Enabled = False
    tArticulo.ForeColor = vbBlack
    tCantidad.ForeColor = vbBlack
End Sub

Private Sub tArticulo_GotFocus()
    Status.Panels(1).Text = " Ingrese un artículo."
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errCA
    Select Case KeyCode
        Case vbKeyF1
            LimpioRenglon
            If Val(lblArticulo.Tag) = "0" Then
                lblArticulo.Tag = "1": lblArticulo.Caption = "&Artículo Específico"
                tArticulo.KeyQuerySP = "CdoArtEspecifico"
                tArticulo.EsEspecifico = True
            Else
                lblArticulo.Tag = "0": lblArticulo.Caption = "&Artículo"
                tArticulo.KeyQuerySP = "VtaCdo"
                tArticulo.EsEspecifico = False
            End If
    End Select
    Exit Sub
errCA:
    clsGeneral.OcurrioError "Ocurrió un error inesperado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then

        If Trim(tArticulo.Text) = vbNullString Then
            If lvVenta.ListItems.Count > 0 And MnuEnvio.Enabled Then
                MnuEnvio_Click
                Foco cTipoTelefono
                Screen.MousePointer = vbDefault
            Else
                lvVenta.SetFocus
            End If
        Else
            If tArticulo.prm_ArtID > 0 Then
                db_FindArticulo
            End If
        End If

    End If

End Sub

Private Sub tArticulo_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub tCantidad_GotFocus()
    Foco tCantidad
    Status.Panels(1).Text = " Ingrese la cantidad de artículos."
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsNumeric(tCantidad.Text) And tArticulo.prm_ArtID > 0 Then
        If Val(tCantidad.Text) > 0 Then tComentario.SetFocus
    End If
End Sub

Private Sub tCantidad_LostFocus()
    Status.Panels(1).Text = ""
    AplicoCantidadLimitadaPorCantidad
End Sub

Private Sub AplicoCantidadLimitadaPorCantidad()
    tmArticuloLimitado.Enabled = False
    If miRenglon.idArticulo = 0 Then Exit Sub
    If IsNumeric(tCantidad.Text) Then
        If CInt(tCantidad.Text) > 0 Then
            tCantidad.Text = CInt(tCantidad.Text)
            tmArticuloLimitado.Enabled = (miRenglon.CantidadAlXMayor = 0 And InStr(1, paCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0) Or miRenglon.CantidadAlXMayor > 1 And miRenglon.CantidadAlXMayor < Val(tCantidad.Text)
        Else
            tCantidad.Text = ""
        End If
    End If
    AplicoTextoDeVentaLimitada
End Sub

Private Sub LimpioDatosCliente()
   
    tNombreC.Text = ""
    tNombreC.Tag = vbNullString
    labDireccion.Caption = ""
    labDireccion.Tag = ""
    lblRutCliente.Caption = ""
    
    cDireccion.Clear: cDireccion.BackColor = Colores.Gris
    gDirFactura = 0
    
    chNomDireccion.Value = 0
    cTipoTelefono.Text = ""
    tTelefono.Text = ""
    tInterno.Text = ""
    
End Sub

Private Sub tCodigo_Change()
    If Val(tCodigo.Tag) > 0 Then
        AccionLimpiar
        tCodigo.Tag = ""
    End If
End Sub

Private Sub tCodigo_GotFocus()
    tCodigo.SelStart = 0
    tCodigo.SelLength = Len(tCodigo.Text)
    Status.Panels(1).Text = " Ingrese el código de Contado a domicilio."
End Sub
Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCodigo.Text) Then
            BuscoVentaTelefonica tCodigo.Text
        Else
            MsgBox "El formato ingresado no es numérico.", vbExclamation, "ATENCIÓN"
            tCodigo.SelStart = 0
            tCodigo.SelLength = Len(tCodigo.Text)
        End If
    End If
End Sub
Private Sub tCodigo_LostFocus()
    Status.Panels(1).Text = ""
End Sub
Private Sub tCodigo_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.Panels(1).Text = " Ingrese el código de Contado a domicilio."
End Sub
Private Sub tComentario_GotFocus()
    Foco tComentario
    Status.Panels(1).Text = " Ingrese un comentario para el artículo."
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And tArticulo.prm_ArtID > 0 Then tUnitario.SetFocus
End Sub

Private Sub tComentario_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub tComentarioDocumento_GotFocus()
    Foco tComentarioDocumento
End Sub

Private Sub tComentarioDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub tInterno_GotFocus()
    tInterno.SelStart = 0
    tInterno.SelLength = Len(tInterno.Text)
    Status.Panels(1).Text = " Ingrese el número de interno o descripción."
End Sub

Private Sub tInterno_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cPagaCon.ListIndex > -1 Or sModificar Then
            Foco tComentarioDocumento
        Else
            Foco cPagaCon
        End If
    End If
End Sub

Private Sub tmArticuloLimitado_Timer()

    tmArticuloLimitado.Enabled = False
    If Val(tmArticuloLimitado.Tag) = 0 Then
        tArticulo.ForeColor = &HFF&
        tmArticuloLimitado.Tag = 1
    Else
        tArticulo.ForeColor = vbBlack
        tmArticuloLimitado.Tag = 0
    End If
    tCantidad.ForeColor = tArticulo.ForeColor
    tmArticuloLimitado.Enabled = True

End Sub

Private Sub tNombreC_GotFocus()
On Error Resume Next
    tNombreC.SelStart = Len(tNombreC.Text)
End Sub

Private Sub tNombreC_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And tArticulo.Enabled Then Foco tArticulo
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        
        Case "envio": MnuEnvio_Click
        Case "validar": AccionValidar
        Case "verfactura": AccionVerFactura
        Case "formapago": CambiarLaFormaDePago
        Case "salir": Unload Me
    End Select
    
End Sub

Private Sub CambiarLaFormaDePago()
On Error GoTo errV
    
    If Val(lblCodigo.Tag) = 44 Then
    
        
        'valido que no tenga el giro.
        Cons = "SELECT * FROM comTransaccionItems WHERE TItTipoItem = 5 AND TItItem = " & Val(tCodigo.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            MsgBox "El cliente ya realizó el pago, no puede cambiar la forma de pago.", vbExclamation, "ATENCIÓN"
            RsAux.Close
            Exit Sub
        End If
        RsAux.Close
        
        Cons = "SELECT * FROM Envio WHERE EnvDocumento = " & Val(tCodigo.Tag) & " AND EnvTipo = 2 AND EnvAgencia > 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            MsgBox "El envío es por Agencia!!!, no puede cambiar la forma de pago.", vbExclamation, "ATENCIÓN"
            RsAux.Close
            Exit Sub
        End If
        RsAux.Close
        
        CambiarLaFormaDePagoADomicilio
    
    ElseIf Val(lblCodigo.Tag) = 7 Then
        CambiarLaFormaDePagoARedPagos
    End If
    
errV:
End Sub

Sub CambiarLaFormaDePagoADomicilio()

    If MsgBox("¿Confirma cambiar la forma de pago de la venta?", vbQuestion + vbYesNo, "Cambiar forma de pago") = vbNo Then Exit Sub
    
    Dim idUsuario As String
    idUsuario = InputBox("Ingrese su dígito de usuario.")
    If Not IsNumeric(idUsuario) Then
        Exit Sub
    End If
    Cons = "SELECT UsuID FROM Usuarios WHERE UsuDigito = " & idUsuario
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        idUsuario = RsAux(0)
    Else
        RsAux.Close
        Exit Sub
    End If
    RsAux.Close
    
    
On Error GoTo ErrGFR
    cBase.BeginTrans
    On Error GoTo ErrResumo
    Dim cTotal As Currency
        
    
    Cons = "UPDATE Envio SET EnvEstado = 0 WHERE EnvTipo = 3 AND EnvDocumento = " & Val(tCodigo.Tag)
    cBase.Execute Cons
    
    Dim vTotal As Currency
    Dim vIVA As Currency

    vTotal = 0
    vIVA = 0
    
    
    Dim rsR As rdoResultset
    Cons = "Select * From RenglonVtaTelefonica, Articulos Where RVTVentaTelefonica = " & Val(tCodigo.Tag) _
            & " And RVTArticulo = ArtID"
    Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsR.EOF
    
        If Not IsNull(rsR("ArtMueveStock")) Then
            If rsR("ArtMueveStock") Then
            'Hago el update en el stock del artículo.
                Dim idT As Currency
                Dim idR As Currency
                idT = rsR("RVTARetirar")
                idR = rsR("RVTCantidad") - rsR("RVTARetirar")
                MarcoStockVenta CLng(idUsuario), rsR("ArtID"), idT, idR, 0, TipoDocumento.ContadoDomicilio, Val(tCodigo.Tag), paCodigoDeSucursal
            End If
        End If
    
        If rsR("ArtID") = 4 Or rsR("ArtID") = 8 Then
            vTotal = vTotal + (rsR("RVTCantidad") * rsR("RVTPrecio"))
            vIVA = vIVA + (rsR("RVTCantidad") * rsR("RVTIva"))
        End If
    
        rsR.MoveNext
    Loop
    rsR.Close
    
    Cons = "DELETE RenglonVtaTelefonica WHERE RVTVentaTelefonica = " & Val(tCodigo.Tag) & " AND RVTArticulo in (4,8)"
    cBase.Execute Cons
    
    Cons = "SELECT * FROM VentaTelefonica WHERE VTeCodigo = " & Val(tCodigo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux("VTeTipo") = TipoDocumento.ContadoDomicilio
    RsAux("VTeTotal") = RsAux("VTeTotal") - vTotal
    RsAux("VTeIva") = RsAux("VTeIva") - vIVA
    cTotal = RsAux("VTeTotal")
    RsAux.Update
    RsAux.Close
    '"UPDATE VentaTelefonica SET VTeTipo = 7 WHERE VTeCodigo = " & Val(tCodigo.Tag)
    
    Dim bEnv As Boolean
    Cons = "SELECT * FROM Envio WHERE EnvDocumento = " & Val(tCodigo.Tag) & " AND EnvTipo = 3 And EnvReclamoCobro > 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux("EnvReclamoCobro") <> cTotal Then
            RsAux.Edit
            RsAux("EnvReclamoCobro") = cTotal
            RsAux.Update
        End If
        bEnv = True
    End If
    RsAux.Close
    
    If Not bEnv Then
        Cons = "SELECT top 1 * FROM Envio WHERE EnvDocumento = " & Val(tCodigo.Tag) & " AND EnvTipo = 3 order by envFechaPrometida"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Edit
            RsAux("EnvReclamoCobro") = cTotal
            RsAux.Update
        End If
        RsAux.Close
    End If
    
    cBase.CommitTrans
    
    On Error Resume Next
    MsgBox "Datos modificados, se cargará nuevamente la información.", vbExclamation, "ATENCIÓN"
    Call tCodigo_KeyPress(13)
    Exit Sub

ErrGFR:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción."
    Exit Sub

ErrResumo:
    Resume ErrRelajo
    
ErrRelajo:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al intentar grabar.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub

End Sub

Sub CambiarLaFormaDePagoARedPagos()

    If MsgBox("¿Confirma cambiar la forma de pago de la venta?", vbQuestion + vbYesNo, "Cambiar forma de pago") = vbNo Then Exit Sub
    
    Dim idUsuario As String
    idUsuario = InputBox("Ingrese su dígito de usuario.")
    If Not IsNumeric(idUsuario) Then
        Exit Sub
    End If
    Cons = "SELECT UsuID FROM Usuarios WHERE UsuDigito = " & idUsuario
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        idUsuario = RsAux(0)
    Else
        RsAux.Close
        Exit Sub
    End If
    RsAux.Close
    
    
On Error GoTo ErrGFR
    cBase.BeginTrans
    On Error GoTo ErrResumo
    Dim cTotal As Currency
        
    
    Cons = "UPDATE Envio SET EnvEstado = 6 WHERE EnvTipo = 3 AND EnvDocumento = " & Val(tCodigo.Tag)
    cBase.Execute Cons
    
    
    Dim vFlete As Currency
    Dim vIvaFlete As Currency
    Cons = "SELECT SUM(EnvValorFlete) Flete, SUM(EnvIvaFlete) IvaFlete FROM Envio WHERE EnvTipo = 3 AND EnvDocumento = " & Val(tCodigo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux("Flete")) Then vFlete = RsAux("Flete")
        If Not IsNull(RsAux("IvaFlete")) Then vIvaFlete = RsAux("IvaFlete")
    End If
    RsAux.Close
    
    Dim idArtFlete As Long
    Cons = "SELECT TFlArticulo FROM Envio INNER JOIN TipoFlete ON TFLCodigo = EnvTipoFlete WHERE EnvTipo = 3 AND EnvDocumento = " & Val(tCodigo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        idArtFlete = RsAux(0)
    End If
    RsAux.Close
    
    Dim vTotal As Currency
    Dim vIVA As Currency

    vTotal = 0
    vIVA = 0
    
    Dim rsR As rdoResultset
    Cons = "Select * From RenglonVtaTelefonica, Articulos Where RVTVentaTelefonica = " & Val(tCodigo.Tag) _
            & " And RVTArticulo = ArtID"
    Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsR.EOF
    
        If Not IsNull(rsR("ArtMueveStock")) Then
            If rsR("ArtMueveStock") Then
            'Hago el update en el stock del artículo.
                Dim idT As Currency
                Dim idR As Currency
                idT = rsR("RVTARetirar")
                idR = rsR("RVTCantidad") - rsR("RVTARetirar")
                MarcoStockVenta CLng(idUsuario), rsR("ArtID"), idT * -1, idR * -1, 0, TipoDocumento.ContadoDomicilio, Val(tCodigo.Tag), paCodigoDeSucursal
            End If
        End If
        rsR.MoveNext
    Loop
    rsR.Close
    
    If idArtFlete > 0 Then
        Cons = "INSERT INTO RenglonVtaTelefonica (RVTVentaTelefonica, RVTArticulo, RVTARetirar, RVTCantidad, RVTPrecio, RVTIva) VALUES (" & Val(tCodigo.Tag) & ", " & _
                idArtFlete & ", 0, 1, " & vFlete & ", " & vIvaFlete & ")"
        cBase.Execute Cons
    End If
    Cons = "SELECT * FROM VentaTelefonica WHERE VTeCodigo = " & Val(tCodigo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux("VTeTipo") = 44
    RsAux("VTeTotal") = RsAux("VTeTotal") + vFlete
    RsAux("VTeIva") = RsAux("VTeIva") + vIvaFlete
    cTotal = RsAux("VTeTotal")
    RsAux.Update
    RsAux.Close
    
    Dim bEnv As Boolean
    Cons = "SELECT * FROM Envio WHERE EnvDocumento = " & Val(tCodigo.Tag) & " AND EnvTipo = 3 And EnvReclamoCobro > 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux("EnvReclamoCobro") <> cTotal Then
            RsAux.Edit
            RsAux("EnvReclamoCobro") = cTotal
            RsAux.Update
        End If
        bEnv = True
    End If
    RsAux.Close
    
    If Not bEnv Then
        Cons = "SELECT top 1 * FROM Envio WHERE EnvDocumento = " & Val(tCodigo.Tag) & " AND EnvTipo = 3 order by envFechaPrometida"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Edit
            RsAux("EnvReclamoCobro") = cTotal
            RsAux.Update
        End If
        RsAux.Close
    End If
    
    cBase.CommitTrans
    
    On Error Resume Next
    MsgBox "Datos modificados, se cargará nuevamente la información.", vbExclamation, "ATENCIÓN"
    Call tCodigo_KeyPress(13)
    Exit Sub

ErrGFR:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción."
    Exit Sub

ErrResumo:
    Resume ErrRelajo
    
ErrRelajo:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al intentar grabar.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub

End Sub



Private Function db_FindPrecioVigente(ByVal lArticulo As Long) As String
Dim rsPV As rdoResultset
    
    On Error Resume Next
    db_FindPrecioVigente = ""
    Cons = "Select * From PrecioVigente WHere PViArticulo = " & lArticulo _
        & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
        & " And PViTipoCuota = " & paTipoCuotaContado & " And PViHabilitado = 1 And PViPrecio Is Not Null"
    Set rsPV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsPV.EOF Then
        If m_Patron = "" Then m_Patron = modMaeDisponibilidad.dis_arrMonedaProp(cMoneda.ItemData(cMoneda.ListIndex), pRedondeo)
        db_FindPrecioVigente = Redondeo(rsPV!PViPrecio, m_Patron)
    End If
    rsPV.Close
    
End Function

Private Function PrecioDelCombo(ByVal idDelCombo As Long) As Currency
Dim rsP As rdoResultset
    Cons = "SELECT PViPrecio, ACoCantidad, ACoPorcPrecio FROM ArticulosDelCombo LEFT OUTER JOIN PrecioVigente ON ACoArticulo = PViArticulo AND PViHabilitado = 1 AND PViTipoCuota = " & paTipoCuotaContado & _
            " AND PViMoneda = " & paMonedaPesos & _
            " WHERE ACoCombo = " & idDelCombo
    Set rsP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsP.EOF
        If IsNull(rsP("PViPrecio")) Then
            MsgBox "Hay artículos del combo que no poseen precio vigente, se interrumpe la carga.", vbExclamation, "ATENCIÓN"
            PrecioDelCombo = 0
            Exit Do
        Else
            If CInt(rsP("ACoPorcPrecio")) < 100 Then
                PrecioDelCombo = PrecioDelCombo + (((rsP("PViPrecio") * rsP("ACoPorcPrecio")) / 100) * rsP("ACoCantidad"))
            Else
                PrecioDelCombo = PrecioDelCombo + (rsP("PViPrecio") * rsP("ACoCantidad"))
            End If
        End If
        rsP.MoveNext
    Loop
    rsP.Close
    m_Patron = dis_arrMonedaProp(CLng(paMonedaPesos), pRedondeo)
    PrecioDelCombo = Redondeo(PrecioDelCombo, m_Patron)
End Function


Private Sub db_FindArticulo()
On Error GoTo ErrBAN

    s_InitVarRenglon
    If cMoneda.ListIndex = -1 Then MsgBox "Debe seleccionar una moneda para cargar el precio.", vbCritical, "ATENCIÓN": Exit Sub
            
    With miRenglon
        .idArticulo = tArticulo.prm_ArtID
        .Tipo = tArticulo.GetField("ArtTipo")
        If IsNull(tArticulo.GetField("ArtHabilitado")) Then
            .EsInhabilitado = True
        Else
            .EsInhabilitado = UCase(tArticulo.GetField("ArtHabilitado")) = "N"
        End If
        
        If tArticulo.GetField("ArtEsCombo") Then
            .IDCombo = .idArticulo
        End If
        
        If Val(lblArticulo.Tag) = 1 Then
            .Especifico = tArticulo.GetField("Especifico")
            .DescuentoEspecifico = tArticulo.GetField("AEsVariacionPrecio")
        End If
        If Not IsNull(tArticulo.GetField("ArtEnVentaXMayor")) Then .CantidadAlXMayor = tArticulo.GetField("ArtEnVentaXMayor") Else .CantidadAlXMayor = 1
        .NombreArticulo = Trim(tArticulo.GetField("ArtNombre"))
    End With
        
    If Ingresado(Val(tArticulo.prm_ArtID)) Then
        MsgBox "El artículo seleccionado ya fue ingresado.", vbExclamation, "ATENCIÓN"
        GoTo evSalir
    ElseIf InStr(strArticuloFlete, tArticulo.prm_ArtID & ",") > 0 Then
        MsgBox "El artículo es de flete, no podrá ingresarlo.", vbExclamation, "Atención"
        GoTo evSalir
    End If
        
    If miRenglon.Especifico = 0 Then
        If Not tArticulo.GetField("ArtEnUso") Then
            snd_ActivarSonido "c:\aa aplicaciones\sonidos\artfuerauso.wav"
            If MsgBox("El artículo ingresado no esta en uso." & vbCrLf & "¿Desea facturarlo de todas formas?", vbQuestion + vbYesNo + vbDefaultButton2, "Artículo fuera de uso") = vbNo Then
                tUnitario.Tag = ""
                Exit Sub
            End If
        ElseIf miRenglon.EsInhabilitado Then
            snd_ActivarSonido "c:\aa aplicaciones\sonidos\artnohabilitado.wav"
            If MsgBox("El artículo ingresado no esta habilitado para la venta." & vbCrLf & "¿Desea facturarlo de todas formas?", vbQuestion + vbYesNo + vbDefaultButton2, "Artículo no habilitado") <> vbYes Then
                tUnitario.Tag = ""
                Exit Sub
            End If
        End If
    End If
    
    If miRenglon.IDCombo > 0 Then
        'Es Combo
        miRenglon.Precio = PrecioDelCombo(miRenglon.IDCombo)
        miRenglon.PrecioOriginal = miRenglon.Precio
    
    Else
    
        If Not PrecioArticulo(miRenglon.idArticulo, cMoneda.ItemData(cMoneda.ListIndex), miRenglon.Precio) Then
            MsgBox "El artículo seleccionado no posee precios ingresados para la moneda seleccionada.", vbInformation, "ATENCIÓN"
            If miRenglon.Especifico > 0 Then GoTo evSalir
        Else
            If miRenglon.Especifico > 0 Then miRenglon.Precio = miRenglon.Precio + miRenglon.DescuentoEspecifico
            miRenglon.PrecioOriginal = miRenglon.Precio
            
        End If
    End If
    
    If miRenglon.Especifico > 0 Then
        
        'InsertoFila
        s_InsertArticulo miRenglon.idArticulo, miRenglon.Tipo, 1, tArticulo.Text, miRenglon.Precio, miRenglon.Precio, IDEspecifico:=miRenglon.Especifico
        
        LimpioRenglon
        lblArticulo.Tag = "0": lblArticulo.Caption = "&Artículo"
        tArticulo.KeyQuerySP = "VtaCdo"
        tArticulo.EsEspecifico = False
        
        cMoneda.Enabled = False
        If sNuevo Then
            MnuEnvio.Enabled = True
            Toolbar1.Buttons("envio").Enabled = True
        End If
        
        tArticulo.Text = ""
        Foco tArticulo
    Else
        AplicoTextoDeVentaLimitada
        Foco tCantidad
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    

ErrBAN:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del artículo.", Err.Description
    
evSalir:
    Screen.MousePointer = 0
    s_InitVarRenglon
    tArticulo.Text = ""
    
End Sub

Private Sub AplicoTextoDeVentaLimitada()
    
    Dim codAux As Long
    codAux = miRenglon.idArticulo
    If IsNumeric(tCantidad.Text) Then
    
        If miRenglon.CantidadAlXMayor = 0 And (InStr(1, paCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0 Or Val(labDireccion.Tag) = 0) Then
            tArticulo.CambiarNombreSinLimpiar (miRenglon.NombreArticulo & " (no vta. a Distr.)")
        ElseIf miRenglon.CantidadAlXMayor > 1 And miRenglon.CantidadAlXMayor < Val(tCantidad.Text) Then
            tArticulo.CambiarNombreSinLimpiar (miRenglon.NombreArticulo & " (limitado a " & miRenglon.CantidadAlXMayor & ")")
        Else
            tArticulo.CambiarNombreSinLimpiar (miRenglon.NombreArticulo)
        End If
    Else
        'Defino el nombre en base a la disponibilidad de venta.
        If miRenglon.CantidadAlXMayor = 0 And (InStr(1, paCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0 Or Val(labDireccion.Tag) = 0) Then
            tArticulo.CambiarNombreSinLimpiar (miRenglon.NombreArticulo & " (no vta. a Distr.)")
        ElseIf miRenglon.CantidadAlXMayor > 1 Then
            tArticulo.CambiarNombreSinLimpiar (miRenglon.NombreArticulo & " (limitado a " & miRenglon.CantidadAlXMayor & ")")
        Else
            tArticulo.CambiarNombreSinLimpiar (miRenglon.NombreArticulo)
        End If
    End If
    If Not tmArticuloLimitado.Enabled Then
        tmArticuloLimitado.Enabled = (tArticulo.Text <> miRenglon.NombreArticulo)
        If Not tmArticuloLimitado.Enabled Then
            tArticulo.ForeColor = vbBlack
            tCantidad.ForeColor = vbBlack
        End If
    End If
    miRenglon.idArticulo = codAux
    
End Sub

Private Sub tTelefono_GotFocus()
    tTelefono.SelStart = 0
    tTelefono.SelLength = Len(tTelefono.Text)
    Status.Panels(1).Text = " Ingrese el teléfono de donde llama el cliente."
End Sub

Private Sub tTelefono_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tInterno
End Sub

Private Sub tTelefono_LostFocus()
Dim aTexto As String
    If Trim(tTelefono.Text) <> "" Then
        aTexto = Trim(clsGeneral.RetornoFormatoTelefono(cBase, tTelefono.Text, 0))
        If aTexto <> "" Then
            tTelefono.Text = aTexto
        Else
            MsgBox "El teléfono ingresado no coincide con los formatos establecidos.", vbExclamation, "ATENCIÓN"
            Foco tTelefono
        End If
    End If

End Sub

Private Sub tUnitario_GotFocus()
On Error Resume Next
   
   If miRenglon.idArticulo > 0 Then s_PresentoPrecio
    
    Status.Panels(1).Text = " Costo unitario del artículo."
    With tUnitario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub tUnitario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And IsNumeric(tUnitario.Text) And cMoneda.ListIndex > -1 Then
        
        If tArticulo.prm_ArtID = 0 Then
            MsgBox "No hay un artículo seleccionado.", vbExclamation, "ATENCIÓN"
            LimpioRenglon
            Exit Sub
        End If
        
        If IsNumeric(tUnitario.Text) Then
            m_Patron = dis_arrMonedaProp(cMoneda.ItemData(cMoneda.ListIndex), pRedondeo)
            tUnitario.Text = Redondeo(CCur(tUnitario.Text), m_Patron)
        End If
        
        If Not IsNumeric(tCantidad.Text) Then
            MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            tCantidad.SetFocus
            Exit Sub
        ElseIf CCur(tCantidad.Text) < 0 Then
            MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            tCantidad.SetFocus
            Exit Sub
        End If
        
        InsertoFila
        
    Else
        If cMoneda.ListIndex = -1 Then
            MsgBox "No seleccionó una moneda.", vbCritical, "ATENCIÓN"
            LimpioRenglon
            cMoneda.SetFocus
            lvVenta.ListItems.Clear
        End If
    End If
        
End Sub

Private Sub LimpioRenglon()

    tArticulo.Text = ""
    tCantidad.Text = ""
    tComentario.Text = ""
    tUnitario.Text = ""
    tUnitario.Tag = ""
    
    tArticulo.ForeColor = vbBlack
    tCantidad.ForeColor = vbBlack
    tmArticuloLimitado.Enabled = False
    
End Sub

Private Function BuscoDescuentoCliente(idArticulo As Long, CatCliente As Long, curUnitario As Currency, intCantidad As Integer) As String
Dim RsDto As rdoResultset

    BuscoDescuentoCliente = curUnitario
    miRenglon.Precio = curUnitario
    
    If paTipoCuotaContado > 0 And CatCliente > 0 Then
        
        m_Patron = dis_arrMonedaProp(cMoneda.ItemData(cMoneda.ListIndex), pRedondeo)
    
        Cons = "Select CDTPorcentaje, AFaCantidadD From ArticuloFacturacion, CategoriaDescuento" _
            & " Where AfaArticulo = " & idArticulo _
            & " And AfaCategoriaD = CDtCatArticulo And CDtCatCliente = " & CatCliente _
            & " And CDtCatPlazo = " & paTipoCuotaContado
            
        Set RsDto = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Not RsDto.EOF Then
            If Not IsNull(RsDto!AFaCantidadD) Then
                If RsDto!AFaCantidadD <= intCantidad Then
                    BuscoDescuentoCliente = Redondeo(curUnitario - (curUnitario * RsDto(0)) / 100, m_Patron)
                    miRenglon.Precio = CCur(BuscoDescuentoCliente)
                Else
                    miRenglon.Precio = curUnitario
                    If MsgBox("El cliente tiene descuento para el artículo pero, no cumple con la cantidad mínima." & Chr(13) _
                        & "¿Le aplica el descuento de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                        BuscoDescuentoCliente = Redondeo(curUnitario - (curUnitario * RsDto(0)) / 100, m_Patron)
                        miRenglon.Precio = CCur(BuscoDescuentoCliente)
                    End If
                End If
            End If
        End If
        RsDto.Close
    End If
    BuscoDescuentoCliente = Format(BuscoDescuentoCliente, FormatoMonedaP)

End Function

Private Sub InsertoFila()
On Error GoTo ErrIF
    
    If Not IsNumeric(tUnitario.Text) Then
        MsgBox "El precio unitario ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        tUnitario.SetFocus
        Exit Sub
    Else
        If CCur(tUnitario.Text) < 0 Then
            MsgBox "No se puede facturar artículos con costo negativo.", vbExclamation, "ATENCIÓN"
            tUnitario.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(tComentario.Text) <> vbNullString Then
        If Not clsGeneral.TextoValido(tComentario.Text) Then
            MsgBox "Se ingreso un carácter no válido en el campo comentario.", vbExclamation, "ATENCIÓN"
            tComentario.SetFocus
            Exit Sub
        End If
    End If
    
    If miRenglon.IDCombo > 0 Then
        s_InsertCombo
    Else
        s_InsertArticulo miRenglon.idArticulo, miRenglon.Tipo, CCur(tCantidad.Text), miRenglon.NombreArticulo, miRenglon.Precio, CCur(tUnitario.Text)
    End If
    
    tArticulo.Text = ""
    Foco tArticulo
    lvVenta.Refresh
    If lvVenta.ListItems.Count > 0 Then
        cMoneda.Enabled = False
        If sNuevo Then
            MnuEnvio.Enabled = True
            Toolbar1.Buttons("envio").Enabled = True
        End If
    Else
        cMoneda.Enabled = True
        MnuEnvio.Enabled = False
        Toolbar1.Buttons("envio").Enabled = False
    End If
    LimpioRenglon
    Exit Sub
    
ErrIF:
    clsGeneral.OcurrioError "Ocurrió un error inesperado al insertar el renglon."

End Sub

Private Function IVAArticulo(lngArticulo As Long)

    IVAArticulo = 0
    Cons = "Select IVAPorcentaje From ArticuloFacturacion, TipoIva " _
        & " Where AFaArticulo = " & lngArticulo & " And AFaIVA = IVACodigo"
    Set RsAuxVta = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAuxVta.EOF Then IVAArticulo = Format(RsAuxVta(0), "#0.00") Else MsgBox "Este artículo no tiene asociado un porcentaje de I.V.A..", vbExclamation, "ATENCIÓN"
    RsAuxVta.Close

End Function

Private Sub LabTotalesEnCero()
    labTotal.Caption = "0.00": labIVA.Caption = "0.00": labSubTotal.Caption = "0.00"
    lblTotalCflete.Caption = "0.00": lblFlete.Caption = "0.00"
End Sub


Private Function Ingresado(lngCodigo As Long)

    If lvVenta.ListItems.Count > 0 Then
        Ingresado = False
        For I = 1 To lvVenta.ListItems.Count
            If InStr(1, lvVenta.ListItems(I).Key, "X") = 0 Then
                If Mid(lvVenta.ListItems(I).Key, 2, Len(lvVenta.ListItems(I).Key)) = lngCodigo Then
                    Ingresado = True
                    Exit Function
                End If
            Else
                If CLng(Mid(lvVenta.ListItems(I).Key, 2)) = lngCodigo Then
                    Ingresado = True
                    Exit Function
                End If
            End If
        Next I
    Else
        Ingresado = False
    End If

End Function

Private Sub RestoLabTotales(curTotal As Currency, curIva As Currency)

    labIVA.Caption = Format(CCur(labIVA.Caption) - Format(curTotal - (curTotal / CCur(1 + (curIva / 100))), "#,##0.00"), "#,##0.00")
    labTotal.Caption = Format(CCur(labTotal.Caption) - curTotal, "#,##0.00")
    labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
    If (lblFlete.Caption = "") Then lblFlete.Caption = "0.00"
    lblTotalCflete.Caption = Format(CCur(lblFlete.Caption) + CCur(labTotal.Caption), "#,##0.00")
    
End Sub

Private Sub VeoCambiosEnDescuentos() '(ByVal idCatAntes As Long)
        
    If Val(labDireccion.Tag) > 0 Then
        CambioPreciosEnLista Val(labDireccion.Tag)
        'Sino queda todo igual
    Else
        'Este no tiene descuentos---------------
        'Recorro la lista y modifico los costos.
        CambioPreciosEnLista (0)
    End If

End Sub

Private Sub CambioPreciosEnLista(lngCodCategoria As Long)
On Error GoTo ErrCPEL

    'ATENCION.----------------------------------------------------
    'Si lngCodCategoria = 0 then No se aplica descuento.----------
    'Pongo los labels de totales en cero.
    LabTotalesEnCero
    If cMoneda.ListIndex = -1 Then Exit Sub
    m_Patron = modMaeDisponibilidad.dis_arrMonedaProp(cMoneda.ItemData(cMoneda.ListIndex), pRedondeo)
    
    For I = 1 To lvVenta.ListItems.Count
        Cons = "Select PViPrecio From PrecioVigente" _
            & " Where PViArticulo = " & Mid(lvVenta.ListItems(I).Key, 2, Len(lvVenta.ListItems(I).Key)) _
            & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
            & " And PViHabilitado = 1" _
            & " And PViTipoCuota = " & paTipoCuotaContado
            
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)

        If RsAux.EOF Then
            lvVenta.ListItems(I).Key = "A" & Mid(lvVenta.ListItems(I).Key, 2, Len(lvVenta.ListItems(I).Key))
            lvVenta.ListItems(I).SubItems(6) = ""
        Else
            If lngCodCategoria = 0 Then
                lvVenta.ListItems(I).Key = "A" & Mid(lvVenta.ListItems(I).Key, 2, Len(lvVenta.ListItems(I).Key))
                lvVenta.ListItems(I).SubItems(3) = Format(Redondeo(RsAuxVta!PViPrecio, m_Patron), "#,##0.00")
                lvVenta.ListItems(I).SubItems(5) = Format(lvVenta.ListItems(I).Text * RsAux!PViPrecio, "#,##0.00")
            Else
                lvVenta.ListItems(I).Key = "A" & Mid(lvVenta.ListItems(I).Key, 2, Len(lvVenta.ListItems(I).Key))
                lvVenta.ListItems(I).SubItems(3) = Format(BuscoDescuentoCliente(Mid(lvVenta.ListItems(I).Key, 2, Len(lvVenta.ListItems(I).Key)), lngCodCategoria, CCur(RsAux!PViPrecio), Val(lvVenta.ListItems(I).Text)), "#,##0.00")
                lvVenta.ListItems(I).SubItems(5) = Format(lvVenta.ListItems(I).Text * lvVenta.ListItems(I).SubItems(3), "#,##0.00")
            End If
            labIVA.Caption = Format(CCur(labIVA.Caption) + CCur(lvVenta.ListItems(I).SubItems(5)) - (CCur(lvVenta.ListItems(I).SubItems(5)) / CCur(1 + (CCur(lvVenta.ListItems(I).SubItems(4)) / 100))), "#,##0.00")
            labTotal.Caption = Format(CCur(labTotal.Caption) + CCur(lvVenta.ListItems(I).SubItems(5)), "#,##0.00")
            labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
            If (lblFlete.Caption = "") Then lblFlete.Caption = "0.00"
            lblTotalCflete.Caption = Format(CCur(lblFlete.Caption) + CCur(labTotal.Caption), "#,##0.00")
        End If
        RsAux.Close
    Next I
    Exit Sub

ErrCPEL:
    clsGeneral.OcurrioError "Ocurrió un error inesperado al modificar los precios, VERIFIQUE."
    
End Sub

Private Sub tUsuario_GotFocus()
    tUsuario.SelStart = 0
    tUsuario.SelLength = Len(tUsuario.Text)
    Status.Panels(1).Text = " Ingrese el dígito de usuario."
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And IsNumeric(tUsuario.Text) Then
        tUsuario.Tag = GetUIDCodigo(Val(tUsuario.Text))
        If Val(tUsuario.Tag) = 0 Then
            tUsuario.Text = vbNullString
            tUsuario.Tag = vbNullString
            Exit Sub
        End If
        If tUsuario.Tag <> vbNullString Then AccionGrabar
    End If
    
End Sub

Private Sub AccionGrabar()

    Dim bNoVaRUT As Boolean
    If txtCliente.Cliente.Tipo = TC_Persona And txtCliente.Cliente.RutPersona <> "" Then
        Dim rsP As VbMsgBoxResult
        rsP = vbCancel
        Do While rsP = vbCancel
            rsP = MsgBox("CLIENTE UNIPERSONAL" & vbCrLf & vbCrLf & "¿El cliente desea facturar con RUT?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "FACTURAR CON RUT")
        Loop
        If rsP = vbNo Then
            bNoVaRUT = True
            rsP = vbCancel
            Do While rsP = vbCancel
                rsP = MsgBox("¿El cliente aún posee ese RUT?" & vbCrLf & vbCrLf & "Si responde NO el RUT se eliminará de la ficha del cliente.", vbQuestion + vbYesNoCancel + vbDefaultButton3, "RUT EN USO")
            Loop
            If rsP = vbNo Then
                
                'Updateo la tabla CPERSONA y registro el suceso de cambio de RUT.
                Cons = "UPDATE CPersona SET CPERuc = NULL WHERE CPeCliente = " & txtCliente.Cliente.Codigo
                cBase.Execute Cons
                
                lblRutCliente.Caption = ""
            End If
        End If
    End If
    
    If ControloDatos(bNoVaRUT) Then
    
        If Not ValidoRUT() Then
            txtCliente.SetFocus
            tUsuario.Enabled = True
            Exit Sub
        End If
    
        Dim bVTaLimitada As Boolean
        bVTaLimitada = AvisoVentaLimitada(True)
        If MsgBox("¿ Confirma grabar los datos ingresados ?", vbQuestion + vbYesNo + IIf(bVTaLimitada, vbDefaultButton2, vbDefaultButton1), "ATENCIÓN") = vbYes Then
                        
            If sNuevo Then
                If ControlStock Then GraboNuevaVenta bVTaLimitada, bNoVaRUT
            Else
                GraboModificacion
            End If
            
        End If
    End If

End Sub

Private Function ControloDatos(ByVal noVaRUT As Boolean) As Boolean
Dim Suma As Currency


    If txtCliente.Cliente.Codigo <= 0 Then
        MsgBox "No se puede facturar sin seleccionar un cliente.", vbExclamation, "ATENCIÓN"
        txtCliente.SetFocus
        ControloDatos = False
        Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Se debe seleccionar una moneda.", vbExclamation, "ATENCIÓN"
        cMoneda.Enabled = True
        cMoneda.SetFocus
        ControloDatos = False
        Exit Function
    End If
    
    If lvVenta.ListItems.Count = 0 Then
        MsgBox "Debe ingresar por los menos un artículo.", vbExclamation, "ATENCIÓN"
        Foco tArticulo
        ControloDatos = False
        Exit Function
    End If
    
    Suma = 0
    For I = 1 To lvVenta.ListItems.Count
        Suma = Suma + lvVenta.ListItems(I).SubItems(5)
    Next I
    
    If (Suma > 40000 * paValorUIUltMes) Then
        MsgBox "La venta supera las 40000 UI, no puede emitirla.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If Suma <> labTotal.Caption Then
        MsgBox "La suma total no coincide con la suma de la lista, verifique.", vbCritical, "ATENCIÓN"
        Foco tArticulo
        ControloDatos = False
        Exit Function
    End If
    
    If Not clsGeneral.TextoValido(tComentarioDocumento.Text) Then
        MsgBox "Se ingreso un carácter no válido en el comentario del documento.", vbExclamation, "ATENCIÓN"
        Foco tComentarioDocumento
        ControloDatos = False
        Exit Function
    End If
    
    If tUsuario.Tag = vbNullString Then
        MsgBox "Debe ingresar el dígito de usuario que factura.", vbExclamation, "ATENCIÓN"
        tUsuario.SetFocus
        ControloDatos = False
        Exit Function
    End If
    
    sDiferencia = False
    
    'Verifico que en la suma a reclamar de cobro en los envíos sea igual a la venta.
    If strCodigoEnvio <> "" And strCodigoEnvio <> "0" Then
        
        Dim fpago As Byte
        
        If cPagaCon.ListIndex >= 0 Then
            fpago = cPagaCon.ItemData(cPagaCon.ListIndex)
        Else
            fpago = 0
        End If
            
        If fpago <> 4 Then
            Cons = "Select EnvCodigo From Envio Where EnvCodigo IN (" & strCodigoEnvio & ") And EnvTipoFlete IN(" & paNoFletesVta & ")"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                Dim sMsgEnvios As String
                Do While Not RsAux.EOF
                    sMsgEnvios = sMsgEnvios & IIf(sMsgEnvios <> "", ", ", "") & RsAux(0)
                    RsAux.MoveNext
                Loop
            
                MsgBox "Atención el/los envíos " & sMsgEnvios & " tienen asignado un tipo de flete que no cobra ventas telefónicas, por favor corrija el dato y reintente.", vbExclamation, "ATENCIÓN"
                ControloDatos = False
                Exit Function
            End If
            RsAux.Close
        End If
        
        If sNuevo Then
            Cons = "Select Count(*), SUM(EnvReclamoCobro) From Envio Where EnvCodigo IN (" & strCodigoEnvio & ") And EnvReclamoCobro Is Not Null"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not IsNull(RsAux(0)) Then
                If RsAux(0) <> 1 Then
                    MsgBox "Sólo se puede cobrar la venta en un envío.", vbExclamation, "ATENCIÓN"
                    RsAux.Close
                    ControloDatos = False
                    Exit Function
                End If
                If RsAux(1) <> CCur(labTotal.Caption) Then
                    If RsAux(0) = 1 Then
                        RsAux.Close
                        Cons = "Select EnvReclamoCobro From Envio Where EnvCodigo IN (" & strCodigoEnvio & ") And EnvReclamoCobro Is Not Null"
                        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        RsAux.Edit
                        RsAux("EnvReclamoCobro") = CCur(labTotal.Caption)
                        RsAux.Update
                        RsAux.Close
                    Else
                        MsgBox "El valor de reclamo de cobro no coincide con el de la venta, verifique.", vbExclamation, "ATENCIÓN"
                        RsAux.Close
                        ControloDatos = False
                        Exit Function
                    End If
                End If
            Else
                RsAux.Close
                MsgBox "Debe existir por lo menos un envío de cobranza.", vbExclamation, "ATENCIÓN"
                ControloDatos = False
                Exit Function
            End If
        Else
            If CCur(labTotal.Caption) <> CCur(labTotal.Tag) Then
                'Aca le cargo al primer envío que tenga reclamo.
                MsgBox "Se ha modificado el total de la venta, se actualizará el valor de reclamo de cobro para el envío.", vbInformation, "ATENCIÓN"
                sDiferencia = True
            End If
        End If
    Else
        If MsgBox("No hay envío asignado a la venta." & vbCrLf & vbCrLf & "¿Confirma almacenar la venta de todas formas?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Posible error") <> vbYes Then
        'MsgBox "Debe existir por lo menos un envío de cobranza.", vbExclamation, "ATENCIÓN"
            Exit Function
        End If
    End If

    'Controlo el formato del número del telefono de llamado
    If Trim(tTelefono.Text) <> "" Then
        Dim aTexto As String
        aTexto = Trim(clsGeneral.RetornoFormatoTelefono(cBase, tTelefono.Text, 0))
        If aTexto <> "" Then
            tTelefono.Text = aTexto
        Else
            MsgBox "El teléfono ingresado no coincide con los formatos establecidos.", vbExclamation, "ATENCIÓN"
            Foco tTelefono
            ControloDatos = False
            Exit Function
        End If
    End If
    
    'Controles eFactura.
    If (Suma > prmImporteConInfoCliente) Then
        'Aca sí o sí tengo que pedir CI o RUT.
        If (txtCliente.Cliente.Tipo = TC_Empresa And txtCliente.Cliente.Documento = "") Or (txtCliente.Cliente.Tipo = TC_Persona And (txtCliente.Cliente.Documento = "" And txtCliente.Cliente.RutPersona = "")) Then
            MsgBox "Es necesario facturar con RUT o Cédula.", vbCritical, "EFactura"
            Exit Function
        End If
        If (labDireccion.Caption = "") Then
            MsgBox "El cliente tiene que tener dirección ingresada, de lo contrario DGI anulará el documento.", vbInformation, "ATENCIÓN"
            Exit Function
        End If
    Else
        If (txtCliente.Cliente.Tipo = TC_Empresa And txtCliente.Cliente.Documento = "") Then  'Or (txtCliente.Cliente.Tipo = TC_Persona And (Suma > prmImporteConInfoCliente Or txtCliente.Cliente.RutPersona <> "")) Then
            If MsgBox("Para facturar a una empresa es necesario facturar con RUT." & vbCrLf & vbCrLf & "¿Desea facturar de todas formas?", vbQuestion + vbYesNo, "EFactura") = vbNo Then
                Exit Function
            End If
            'MsgBox "Para facturar a una empresa es necesario facturar con RUT.", vbExclamation, "EFactura"
            'Exit Function
        'ElseIf (txtCliente.Cliente.Tipo = TC_Empresa And labDireccion.Caption = "") Then
        ElseIf labDireccion.Caption = "" And (txtCliente.Cliente.Tipo = TC_Empresa Or (txtCliente.Cliente.Tipo = TC_Persona And txtCliente.Cliente.RutPersona <> "" And Not noVaRUT)) Then
            MsgBox "La empresa debe tener domicilio fiscal para facturar.", vbExclamation, "ATENCIÓN"
            Exit Function
            
        End If
    End If
    

    ControloDatos = True
    If cPagaCon.ListIndex = 0 Then
        MsgBox "Se modificará la ficha del cliente indicando que opera con cheque.", vbInformation, "Atención"
    End If
    
    'Si paso la Validación controlo la dirección que factura-----------------------------------------------------------------
    If cDireccion.ListIndex <> -1 Then
        On Error Resume Next
        If gDirFactura <> cDireccion.ItemData(cDireccion.ListIndex) Then        'Cambio Dir Facutua
            If MsgBox("Ud. a cambiado la dirección con la que el cliente factura habitualmente." & vbCrLf & "Quiere que esta dirección quede por defecto para facturar.", vbQuestion + vbYesNo, "Dirección por Defecto al Facturar") = vbNo Then Exit Function
            
            If cDireccion.ItemData(cDireccion.ListIndex) <> Val(cDireccion.Tag) Then        'Dir. selecc. <> a la Ppal.
                
                Dim rsAD As rdoResultset
                Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & txtCliente.Cliente.Codigo & " And DAuDireccion = " & cDireccion.ItemData(cDireccion.ListIndex)
                Set rsAD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAD.EOF Then
                    rsAD.Edit: rsAD!DAuFactura = True: rsAD.Update
                End If
                rsAD.Close
            End If
            
            If gDirFactura <> Val(cDireccion.Tag) Then      'La gDirFactura Anterior no era la ppal, la desmarco
                Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & txtCliente.Cliente.Codigo & " And DAuDireccion = " & gDirFactura
                Set rsAD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAD.EOF Then
                    rsAD.Edit: rsAD!DAuFactura = False: rsAD.Update
                End If
                rsAD.Close
            End If
            gDirFactura = cDireccion.ItemData(cDireccion.ListIndex)
        End If
        
    End If
    '----------------------------------------------------------------------------------------------------------------------------------
    
End Function
Private Function ControlStock() As Boolean

    For Each itmx In lvVenta.ListItems
        If InStr(1, itmx.Key, "C") = 0 And Not ArticuloEsServicio(Mid(itmx.Key, 2, Len(itmx.Key))) Then   'Val(itmx.SubItems(7)) <> paTipoArticuloServicio Then
                            
            Cons = "Select StTCantidad From StockTotal Where StTArticulo = " & Mid(itmx.Key, 2, Len(itmx.Key))
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If RsAux.EOF Then
                If MsgBox("No existe registro de stock para el artículo " & itmx.SubItems(1) & "." _
                    & Chr(13) & "¿Desea facturar de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                    ControlStock = False
                    RsAux.Close
                    Exit Function
                End If
            Else
                If RsAux!StTCantidad < CInt(itmx.Text) Then
                    If MsgBox("No existe tanto stock para el artículo " & Trim(itmx.SubItems(1)) & "." & Chr(13) _
                        & " ¿Desea facturar de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                            RsAux.Close
                            ControlStock = False
                            Exit Function
                    End If
                End If
            End If
            RsAux.Close
        End If
    Next
    ControlStock = True

End Function

Private Function ArticuloEsServicio(ByVal idArticulo As Long) As Boolean
     'EsTipoDeServicio = (InStr(1, "," & tTiposArtsServicio & ",", "," & idTipo & ",") > 0)
     ArticuloEsServicio = False
     Cons = "SELECT ArtMueveStock FROM Articulos WHERE ArtID = " & idArticulo
     Dim rsS As rdoResultset
     Set rsS = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
     If Not rsS.EOF Then
        If Not IsNull(rsS(0)) Then
            ArticuloEsServicio = Not rsS(0)
        End If
     End If
     rsS.Close
End Function


Private Function fnc_GetNombreVenta(ByVal iTipo As Byte) As String
    Select Case iTipo
        Case 7: fnc_GetNombreVenta = "Venta telefónica"
        Case 32: fnc_GetNombreVenta = "Online a confirmar"
        Case 33: fnc_GetNombreVenta = "Online"
    End Select
End Function

Private Sub GraboNuevaVenta(ByVal bSucVtaXMayor As Boolean, ByVal noVaRUT As Boolean)
Dim lnNro As Long
Dim Control As Long
Dim aUsuario As Long, strDefensa As String
Dim idAutPrecio As Long, idAutCamName As Long
    
    Screen.MousePointer = vbHourglass
    Control = VerificoCostoArticulo
    Screen.MousePointer = 0
    
    If Control = 1 Then
        Dim objSuceso As New clsSuceso
        aUsuario = 0
        objSuceso.TipoSuceso = TipoSuceso.ModificacionDePrecios
        objSuceso.ActivoFormulario tUsuario.Tag, "Cambio de Precio", cBase
        Me.Refresh
        aUsuario = objSuceso.Usuario
        strDefensa = objSuceso.Defensa
        idAutPrecio = objSuceso.Autoriza
        Set objSuceso = Nothing
        If aUsuario = 0 Then Exit Sub
    End If
    
    Dim bSucesoNombre As Boolean
    Dim sDefCambioNombre As String
    If Trim(txtCliente.Cliente.Nombre) <> Trim(tNombreC.Text) Then
        'Cambio el nombre del cliente
        bSucesoNombre = True
        'Llamo al registro del Suceso-------------------------------------------------------------
        Set objSuceso = New clsSuceso
        aUsuario = 0
        objSuceso.TipoSuceso = TipoSuceso.FacturaCambioNombre
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "Cambio de Nombre", cBase
        Me.Refresh
        aUsuario = objSuceso.Usuario
        sDefCambioNombre = objSuceso.Defensa
        idAutCamName = objSuceso.Autoriza
        Set objSuceso = Nothing
        If aUsuario = 0 Then Screen.MousePointer = 0: Exit Sub
    Else
        bSucesoNombre = False
    End If
    
    Dim sDefNoVender As String, iUsuNoV As Long, idAutNoV As Long
    If txtCliente.Cliente.NoVender Then
        Set objSuceso = New clsSuceso
        objSuceso.TipoSuceso = TipoSuceso.ClienteNoVender
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "No vender a cliente", cBase
        Me.Refresh
        iUsuNoV = objSuceso.Usuario
        sDefNoVender = objSuceso.Defensa
        idAutNoV = objSuceso.Autoriza
        Set objSuceso = Nothing
        If iUsuNoV = 0 Then Screen.MousePointer = 0: Exit Sub
    End If
    
    Dim sDefVtaxMayor As String, iUsuVtaXMayor As Long, idAutVtaXMayor As Long
    If bSucVtaXMayor Then
        Set objSuceso = New clsSuceso
        objSuceso.TipoSuceso = TipoSuceso.FacturaArticuloInhabilitado
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "Venta por mayor inhabilitada", cBase
        Me.Refresh
        iUsuVtaXMayor = objSuceso.Usuario
        sDefVtaxMayor = objSuceso.Defensa
        idAutVtaXMayor = objSuceso.Autoriza
        Set objSuceso = Nothing
        If iUsuVtaXMayor = 0 Then Screen.MousePointer = 0: Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    On Error GoTo ErrGFR
    cBase.BeginTrans 'Comienzo la TRANSACCION-------------------------------------------------------------------------
    
    On Error GoTo ErrResumo
    FechaDelServidor
    
    Dim rsVta As rdoResultset
    Cons = "SELECT * FROM VentaTelefonica WHERE VTeCodigo = 0 AND VTeCliente = " & txtCliente.Cliente.Codigo
    Set rsVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    'Cierro el set de Venta Telefonica.
    rsVta.AddNew
    rsVta!VTeFechaLlamado = Format(gFechaServidor, sqlFormatoFH)
    
    rsVta!VTeTipo = TipoDocumento.ContadoDomicilio
    If cPagaCon.ListIndex > -1 Then
        If cPagaCon.ItemData(cPagaCon.ListIndex) = 4 Then rsVta!VTeTipo = TipoDocumento.VentaRedPagosTelefonicas
    End If
    
    rsVta!VTeCliente = txtCliente.Cliente.Codigo
    rsVta!VTeMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    rsVta!VTeTotal = CCur(labTotal.Caption)
    rsVta!VTeIva = CCur(labIVA.Caption)
    rsVta!VTeSucursal = paCodigoDeSucursal
    rsVta!VTeUsuario = tUsuario.Tag
    rsVta!VTeFModificacion = Format(gFechaServidor, sqlFormatoFH)
    If Trim(tComentarioDocumento.Text) <> vbNullString Then
        rsVta!VTeComentario = clsGeneral.SacoEnter(Trim(tComentarioDocumento.Text))
    End If
    If Trim(tTelefono.Text) <> vbNullString Then rsVta!VTeTelefonoLlamada = Trim(tTelefono.Text)
    rsVta!VTeCompleta = 1   'Indico que la venta esta completa.
    If cDireccion.ListIndex > -1 Then rsVta!VTeDireccionFactura = cDireccion.ItemData(cDireccion.ListIndex)
    If chNomDireccion.Value = 1 Then
        rsVta!VTeNombreFactura = Trim(tNombreC.Text) & " (" & Trim(cDireccion.Text) & ")"
    Else
        If Trim(txtCliente.Cliente.Nombre) <> Trim(tNombreC.Text) Then rsVta!VTeNombreFactura = Trim(tNombreC.Text)
    End If
    If noVaRUT Then
        rsVta("VTeSinRut") = 1
    End If
    rsVta.Update
    rsVta.Close
    
    Cons = "SELECT MAX(VTeCodigo) From VentaTelefonica" _
        & " WHERE VTeCliente = " & txtCliente.Cliente.Codigo _
        & " AND VTeFechaLlamado = '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
        & " AND VTeSucursal = " & paCodigoDeSucursal _
        & " And VTeUsuario = " & tUsuario.Tag _
        & " AND VTeTotal = " & CCur(labTotal.Caption)
                
    Set rsVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    lnNro = rsVta(0)
    rsVta.Close
    
    'Grabo el telefono en la BD Telefonos
    GraboDatosBDTelefono txtCliente.Cliente.Codigo
    
    db_CliCheque

    For Each itmx In lvVenta.ListItems
        
        Cons = "INSERT INTO RenglonVtaTelefonica (RVTVentaTelefonica, RVTArticulo, RVTCantidad, RVTPrecio, RVTIVA, RVTARetirar, RVTComentario)" _
            & " VALUES (" & lnNro & ", " & Mid(itmx.Key, 2, Len(itmx.Key)) _
            & ", " & itmx.Text & ", " & CCur(itmx.SubItems(3)) _
            & ", " & CCur(itmx.SubItems(3)) - (CCur(itmx.SubItems(3)) / CCur(1 + (CCur(itmx.SubItems(4) / 100))))
        
        If Trim(itmx.Tag) <> vbNullString Then
            Cons = Cons & ", " & CCur(itmx.Text) - CCur(itmx.Tag)
        Else
            If CCur(itmx.Tag) > 0 Then
                Cons = Cons & ", " & CCur(itmx.Text)
            Else
                'Es un artículo de Venta.
                Cons = Cons & ", 0"
            End If
        End If
        If Trim(itmx.SubItems(2)) = vbNullString Then
            Cons = Cons & ", Null"
        Else
            Cons = Cons & ", '" & Trim(itmx.SubItems(2)) & "'"
        End If
        Cons = Cons & ")"
        cBase.Execute (Cons)
               
        'If Val(itmx.SubItems(7)) <> paTipoArticuloServicio Then
        If Not ArticuloEsServicio(Mid(itmx.Key, 2, Len(itmx.Key))) Then
            'Hago el update en el stock del artículo.
            If Mid(itmx.Key, 1, 1) = "A" Or Mid(itmx.Key, 1, 1) = "Z" Then
                MarcoStockVenta CLng(tUsuario.Tag), CLng(Mid(itmx.Key, 2, Len(itmx.Key))), CCur(itmx.Text) - CCur(itmx.Tag), CCur(itmx.Tag), 0, TipoDocumento.ContadoDomicilio, lnNro, paCodigoDeSucursal
            Else
                MarcoStockVenta CLng(tUsuario.Tag), Mid(itmx.Key, 2, Len(itmx.Key)), CCur(itmx.Text), 0, 0, TipoDocumento.ContadoDomicilio, lnNro, paCodigoDeSucursal
            End If
        End If
        
        If Val(itmx.SubItems(8)) > 0 Then
            Cons = "SELECT * FROM ArticuloEspecifico WHERE AEsID = " & Val(itmx.SubItems(8)) '& " AND (AEsDocumento IS NULL OR AEsDocumento "
            Set rsVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            rsVta.Edit
            rsVta("AEsTipoDocumento") = 7
            rsVta("AEsDocumento") = lnNro
            rsVta.Update
            rsVta.Close
        End If
        
    Next
    
    Dim iTipoRP As Integer
    If cPagaCon.ListIndex > -1 Then
        If cPagaCon.ItemData(cPagaCon.ListIndex) = 4 Then iTipoRP = 1
    End If
    
    If strCodigoEnvio <> vbNullString Then
            
        If iTipoRP = 1 Then
            
            Cons = "UPDATE Envio Set EnvDocumento = " & lnNro _
                        & ", EnvUsuario = " & tUsuario.Tag _
                        & ", EnvEstado = 6" _
                        & ", EnvCliente = " & txtCliente.Cliente.Codigo _
                        & " WHERE EnvCodigo IN (" & strCodigoEnvio & ")"
            cBase.Execute (Cons)
            
        Else
            'Si hay envios que no los paga ahora pero los hizo.
            Cons = "UPDATE Envio Set EnvDocumento = " & lnNro _
                        & ", EnvUsuario = " & tUsuario.Tag _
                & " WHERE EnvCodigo IN (" & strCodigoEnvio & ")" _
                & " And EnvFormaPago <> " & TipoPagoEnvio.PagaAhora
            cBase.Execute (Cons)
            
            Cons = "UPDATE Envio Set EnvCliente = " & txtCliente.Cliente.Codigo _
                & " WHERE EnvCodigo IN (" & strCodigoEnvio & ")"
            cBase.Execute (Cons)
            
            Cons = "UPDATE EnvioVaCon Set EVCDocumento = " & lnNro & "WHERE EVCEnvio IN (" & strCodigoEnvio & ")"
            cBase.Execute (Cons)
        End If
        
    End If
    
    If Control <> 0 Then
        Dim aTexto As String
        aTexto = fnc_GetNombreVenta(TipoDocumento.ContadoDomicilio) & " Nro. :" & lnNro & " (" & Trim(gFechaServidor) & ")"
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, TipoSuceso.ModificacionDePrecios, paCodigoDeTerminal, aUsuario, 0, Descripcion:=aTexto, Defensa:=Trim(strDefensa), Valor:=1, idCliente:=txtCliente.Cliente.Codigo, idautoriza:=idAutPrecio
    End If
    If bSucesoNombre Then
        aTexto = "Cambio de Nombre en " & fnc_GetNombreVenta(TipoDocumento.ContadoDomicilio)
        sDefCambioNombre = "Nuevo nombre: " & Trim(tNombreC.Text) & vbCrLf & sDefCambioNombre
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, TipoSuceso.FacturaCambioNombre, paCodigoDeTerminal, aUsuario, 0, 0, aTexto, Trim(sDefCambioNombre), 1, txtCliente.Cliente.Codigo, idAutCamName
    End If
    If iUsuNoV > 0 Then
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, TipoSuceso.ClienteNoVender, paCodigoDeTerminal, iUsuNoV, 0, , "Cliente No vender", Trim(sDefNoVender), 1, txtCliente.Cliente.Codigo, idAutNoV
    End If
    If iUsuVtaXMayor > 0 Then
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, TipoSuceso.FacturaArticuloInhabilitado, paCodigoDeTerminal, iUsuVtaXMayor, 0, , "Venta por Mayor venta telefónica", Trim(sDefVtaxMayor), 1, txtCliente.Cliente.Codigo, idAutVtaXMayor
    End If
    cBase.CommitTrans

    
    On Error Resume Next
    Cons = "Select * From RenglonVtaTelefonica, Articulo Where RVTVentaTelefonica = " & lnNro _
            & " And ArtInstalador > 0 And RVTArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then EjecutarApp App.Path & "\instaires.exe", "VTe:" & CStr(lnNro)
    RsAux.Close
    
    tCodigo.Text = lnNro
    ctr_Enabled False
    BuscoVentaTelefonica lnNro
    tCodigo.SetFocus
    sNuevo = False
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrGFR:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción."
    Exit Sub

ErrResumo:
    Resume ErrRelajo
    
ErrRelajo:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al intentar grabar.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub ControloArticuloConEnvio(CodVenta As Long)
Dim RsControl As rdoResultset

    'Verifico que esten los artículos del envío igual a los de la venta.
    Cons = "Select REvArticulo From RenglonEnvio Where RevEnvio in (" & strCodigoEnvio & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    Do While Not RsAux.EOF
        Cons = "Select * from RenglonVtaTelefonica Where RVTVentaTelefonica = " & CodVenta _
            & " And RVTArticulo = " & RsAux!REvArticulo
        Set RsControl = cBase.OpenResultset(Cons, rdOpenForwardOnly)
        If RsControl.EOF Then
            MsgBox "La venta no se almacenó correctamente, los artículos de la Venta no coinciden con los del envío." & Chr(13) & "VERIFIQUE", vbCritical, "ATENCIÓN"
        End If
        RsControl.Close
        RsAux.MoveNext
    Loop
    RsAux.Close

End Sub
Private Sub ctr_Enabled(ByVal bEn As Boolean)

    With tCodigo
        .Enabled = Not bEn
        .BackColor = IIf(Not bEn, vbWhite, vbButtonFace)
    End With
    
    With cMoneda
        .Enabled = bEn And sNuevo
        .BackColor = IIf(bEn, Obligatorio, vbButtonFace)
    End With
    
    txtCliente.Enabled = bEn
    tNombreC.Enabled = bEn
    
    With tArticulo
        .Enabled = bEn
        .BackColor = IIf(bEn, Obligatorio, vbButtonFace)
    End With
    With tCantidad
        .Enabled = bEn
        .BackColor = IIf(bEn, Obligatorio, vbButtonFace)
    End With
    With tComentario
        .Enabled = bEn
        .BackColor = IIf(bEn, vbWhite, vbButtonFace)
    End With
    With tUnitario
        .Enabled = bEn
        .BackColor = IIf(bEn, Obligatorio, vbButtonFace)
    End With
    
    With tComentarioDocumento
        .Enabled = bEn
        .BackColor = IIf(bEn, vbWhite, vbButtonFace)
    End With
    
    With tUsuario
        .Enabled = bEn
        .BackColor = IIf(bEn, Obligatorio, vbButtonFace)
    End With
    
    With tTelefono
        .Enabled = bEn
        .BackColor = IIf(bEn, vbWhite, vbButtonFace)
    End With
    With cTipoTelefono
        .Enabled = bEn
        .BackColor = IIf(bEn, vbWhite, vbButtonFace)
    End With
    With tInterno
        .Enabled = bEn
        .BackColor = IIf(bEn, vbWhite, vbButtonFace)
    End With
    
    With cPagaCon
        .Enabled = bEn
        .BackColor = IIf(bEn, vbWhite, vbButtonFace)
    End With
    
    With lvVenta
        .Enabled = bEn
        .BackColor = IIf(bEn, vbWhite, vbButtonFace)
    End With
    
End Sub

Private Sub AccionNuevo()
On Error GoTo ErrAN
    
    Screen.MousePointer = vbHourglass
    sNuevo = True
    strCodigoEnvio = 0
    tCodigo.Tag = ""
    tCodigo.Text = ""
    txtCliente.Text = ""
    txtCliente.DocumentoCliente = DC_CI
    AccionLimpiar
    ctr_Enabled True
    labFecha.Caption = Format(Date, "d-Mmm-yyyy")
    BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
    Botones False, False, False, True, True, Toolbar1, Me
    MnuBuscarCtosEmpresa.Enabled = False
    MnuBuscarCtosPersona.Enabled = False
    
    On Error Resume Next
    txtCliente.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrAN:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error inesperado.", Trim(Err.Description)
End Sub

Private Sub AccionCancelar()

    ctr_Enabled False
    AccionLimpiar
    MnuEnvio.Enabled = False
    Toolbar1.Buttons("envio").Enabled = False
    MnuBuscarCtosEmpresa.Enabled = True
    MnuBuscarCtosPersona.Enabled = True
    If sNuevo Then
        If Trim(strCodigoEnvio) <> vbNullString Then
            BorroEnvios
        End If
        strCodigoEnvio = ""
        Botones True, False, False, False, False, Toolbar1, Me
    Else
        Botones True, True, True, False, False, Toolbar1, Me
        txtCliente.Enabled = True
        BuscoVentaTelefonica Val(tCodigo.Tag)    'rsVta!VTeCodigo
    End If
    sNuevo = False
    sModificar = False
    tCodigo.SetFocus
    
End Sub
Private Sub CalculoArticulosEnEnvio()
    If strCodigoEnvio = "" Then strCodigoEnvio = 0: Exit Sub
    
    Cons = "Select Sum(REvAEntregar), REvArticulo From RenglonEnvio Where REvEnvio IN (" & strCodigoEnvio & ")" _
        & "Group by REvArticulo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        For Each itmx In lvVenta.ListItems
            If CLng(Mid(itmx.Key, 2, Len(itmx.Key))) = RsAux!REvArticulo Then
                itmx.Tag = RsAux(0)
                Exit For
            End If
        Next
        RsAux.MoveNext
    Loop
    RsAux.Close
End Sub

Private Sub BuscoVentaTelefonica(lnCodigo As Long)
On Error GoTo ErrBVT

    Screen.MousePointer = vbHourglass
    strCodigoEnvio = 0
    Botones True, False, False, False, False, Toolbar1, Me
    MnuEnvio.Enabled = False
    Toolbar1.Buttons("envio").Enabled = False
    MnuFacturar.Enabled = False
    Toolbar1.Buttons("facturar").Enabled = False
    MnuBuscarCtosEmpresa.Enabled = True
    MnuBuscarCtosPersona.Enabled = True
    MnuValidarVenta.Enabled = False
    Toolbar1.Buttons("validar").Enabled = False
    MnuDetalleFactura.Enabled = False
    Toolbar1.Buttons("verfactura").Enabled = False
        
    Dim rsVta As rdoResultset
    Cons = "Select * From VentaTelefonica Where VTeCodigo = " & lnCodigo _
        & " And VTeTipo IN (" & TipoDocumento.ContadoDomicilio & ", " & TipoDocumento.VentaOnLineAConfirmar & ", " & TipoDocumento.VentaOnLineConfirmada & ", " & TipoDocumento.VentaRedPagosTelefonicas & ")"
        
    Set rsVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    AccionLimpiar
    tCodigo.Tag = ""
    cMoneda.ListIndex = -1
    
    tCodigo.Text = lnCodigo
    
    If rsVta.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "No existe una venta con ese código.", vbExclamation, "ATENCIÓN"
        Foco tCodigo
    Else
        tCodigo.Tag = lnCodigo
        lblCodigo.Tag = rsVta("VTeTipo")
        If rsVta("VTeTipo") = 32 Then
            Me.BackColor = &H9090B0
            Me.Caption = "Contados onLine a confirmar"
        ElseIf rsVta("VTeTipo") = 33 Then
            Me.BackColor = &HC0C0C0
            Me.Caption = "Contados onLine a cobrar en domicilio"
        ElseIf rsVta("VTeTipo") = 44 Then
            Me.BackColor = &HE7EBE6
            Me.Caption = "Contados telefónicos a cobrar por redpagos"
            BuscoCodigoEnCombo cPagaCon, ePagaCon.RedPagos
            Toolbar1.Buttons("formapago").Enabled = True
        Else
            Me.BackColor = &HC0CAAA
            Me.Caption = "Contados telefónicos a cobrar en domicilio"
            Toolbar1.Buttons("formapago").Enabled = True

            'BuscoCodigoEnCombo cPagaCon, ePagaCon.Efectivo
        End If
        
        txtCliente.CargarControl rsVta("VTeCliente")
        If (Not IsNull(rsVta("VTeSinRut"))) Then
            If (rsVta("VTeSinRut") = 1) Then lblRutCliente.Caption = ""
        End If
        
        If Not IsNull(rsVta!VTeNombreFactura) Then
            tNombreC.Text = Trim(rsVta!VTeNombreFactura)
            chNomDireccion.Value = 1
        End If
        
            'Datos de tabla VentaTelefonica
        BuscoCodigoEnCombo cMoneda, rsVta!VTeMoneda
        labFecha.Caption = Format(rsVta!VTeFechaLlamado, "d-Mmm-yyyy")
        labFecha.Tag = rsVta("VTeFModificacion")
        
        If Not IsNull(rsVta!VTeTelefonoLlamada) Then tTelefono.Text = Trim(rsVta!VTeTelefonoLlamada)
        If Not IsNull(rsVta!VTeComentario) Then tComentarioDocumento.Text = Trim(rsVta!VTeComentario)
        
        If Not IsNull(rsVta!VTeDocumento) Then
            CargoListaDeArticulos rsVta("VTeCodigo"), rsVta("VTeDocumento")
        Else
            CargoListaDeArticulos rsVta("VTeCodigo"), 0
        End If
        
        labIVA.Caption = Format(rsVta!VTeIva, "#,##0.00")
        labTotal.Caption = Format(rsVta!VTeTotal, "#,##0.00")
        labTotal.Tag = rsVta("VTeTotal")
        labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
        labUsuario.Caption = GetUIDIdentificacion(0, rsVta!VTeUsuario)
        
        If Not IsNull(rsVta!VTeDireccionFactura) Then
            For I = 0 To cDireccion.ListCount - 1
                If cDireccion.ItemData(I) = rsVta!VTeDireccionFactura Then cDireccion.ListIndex = I: Exit For
            Next I
        End If
        
        MnuValidarVenta.Enabled = False
        Toolbar1.Buttons("validar").Enabled = False
        
        If Not IsNull(rsVta!VTeAnulado) Then
            Shape2.BackColor = &HD0&
            chNomDireccion.BackColor = Shape2.BackColor
            MsgBox "Esta venta fue anulada.", vbInformation, "INFORMACIÓN"
            Exit Sub
        End If
        
        If Not IsNull(rsVta!VTeDocumento) Then
            MsgBox "Esta venta ya fue facturada.", vbInformation, "INFORMACIÓN"
            MnuFacturar.Enabled = False
            Toolbar1.Buttons("facturar").Enabled = False
            MnuDetalleFactura.Enabled = True
            Toolbar1.Buttons("verfactura").Enabled = True
            Toolbar1.Buttons("formapago").Enabled = False
            Exit Sub
        ElseIf rsVta("VTeTipo") <> TipoDocumento.VentaRedPagosTelefonicas Then
            MnuFacturar.Enabled = True
            Toolbar1.Buttons("facturar").Enabled = True
        End If
        
        If Not sModificar Then
            If Not IsNull(rsVta!VTeCompleta) Then
                If rsVta!VTeCompleta = 0 Then
                    MsgBox "Esta venta no posee el envío de cobranza correcto, verifique.", vbExclamation, "ATENCIÓN"
                    MnuValidarVenta.Enabled = True
                    Toolbar1.Buttons("validar").Enabled = True
                    MnuFacturar.Enabled = False
                    Toolbar1.Buttons("facturar").Enabled = False
                End If
            Else
                MsgBox "Esta venta no posee el envío de cobranza correcto, verifique. El campo esta nulo", vbExclamation, "ATENCIÓN"
                MnuValidarVenta.Enabled = True
                Toolbar1.Buttons("validar").Enabled = True
                MnuFacturar.Enabled = False
                Toolbar1.Buttons("facturar").Enabled = False
            End If
        End If
        
        Botones True, True, True, False, False, Toolbar1, Me
        
        BuscoEnviosParaLaVenta Val(tCodigo.Tag)
        If strCodigoEnvio = 0 Then
            MnuValidarVenta.Enabled = False
            Toolbar1.Buttons("validar").Enabled = False
        Else
            If cPagaCon.ListIndex > -1 Then
                If cPagaCon.ItemData(cPagaCon.ListIndex) <> ePagaCon.RedPagos Then
                    CargarFletesEnvio
                End If
            Else
                CargarFletesEnvio
            End If
        End If
        
    
        MnuEnvio.Enabled = True
        Toolbar1.Buttons("envio").Enabled = True
        
    End If
    rsVta.Close
    Screen.MousePointer = vbDefault
    Exit Sub

ErrBVT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al buscar la venta."
    
End Sub

Private Sub CargoListaDeArticulos(ByVal idVta As Long, ByVal idDocumento As Long)

    If idDocumento = 0 Then
        Cons = "SELECT R.*, ArtTipo, ISNULL(AEsNombre, ArtNombre) ArtNombre, IvaPorcentaje, IsNull(AEsID, 0) AEsID " & _
            "FROM RenglonVtaTelefonica R INNER JOIN Articulo ON RVTArticulo = ArtId " & _
            "LEFT OUTER JOIN ArticuloEspecifico ON AEsArticulo = ArtId AND AEsDocumento = R.RVTVentaTelefonica AND AEsTipoDocumento = 7 " & _
            "LEFT OUTER JOIN ArticuloFacturacion ON ArtID = AFaArticulo " & _
            "LEFT OUTER JOIN TipoIva ON AFaIva = IvaCodigo " & _
            "WHERE RVTVentaTelefonica = " & idVta
    Else
        Cons = "SELECT R.*, ArtTipo, ISNULL(AEsNombre, ArtNombre) ArtNombre, IvaPorcentaje, IsNull(AEsID, 0) AEsID " & _
            "FROM RenglonVtaTelefonica R INNER JOIN Articulo ON RVTArticulo = ArtId " & _
            "LEFT OUTER JOIN ArticuloEspecifico ON AEsArticulo = ArtId AND AEsDocumento = " & idDocumento & " AND AEsTipoDocumento = 1 " & _
            "LEFT OUTER JOIN ArticuloFacturacion ON ArtID = AFaArticulo " & _
            "LEFT OUTER JOIN TipoIva ON AFaIva = IvaCodigo " & _
            "WHERE RVTVentaTelefonica = " & idVta
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        s_InsertArticulo RsAux!RVTArticulo, RsAux!ArtTipo, RsAux!RVtCantidad, RsAux!ArtNombre, RsAux!RVtPrecio, RsAux!RVtPrecio, IIf(IsNull(RsAux!RVTComentario), "", RsAux!RVTComentario), IDEspecifico:=RsAux("AEsID")
        RsAux.MoveNext
    Loop
    RsAux.Close

End Sub

Private Sub BuscoEnviosParaLaVenta(ByVal idVta As Long)

    Cons = "Select EnvCodigo From Envio Where EnvTipo = " & TipoEnvio.Cobranza _
        & " And EnvDocumento = " & idVta
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        If strCodigoEnvio = 0 Then
            strCodigoEnvio = RsAux!EnvCodigo
        Else
            strCodigoEnvio = strCodigoEnvio & "," & RsAux!EnvCodigo
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    If strCodigoEnvio <> vbNullString Then CalculoArticulosEnEnvio
    
End Sub
Private Sub AccionModificar()
On Error GoTo errModificar

    Screen.MousePointer = vbHourglass
    sModificar = True
    BuscoVentaTelefonica Val(tCodigo.Tag)
    Botones False, False, False, True, True, Toolbar1, Me
    MnuEnvio.Enabled = False
    Toolbar1.Buttons("envio").Enabled = False
    Toolbar1.Buttons("facturar").Enabled = False
    MnuFacturar.Enabled = False
    MnuBuscarCtosEmpresa.Enabled = False
    MnuBuscarCtosPersona.Enabled = False

    ctr_Enabled True
    
    txtCliente.Enabled = False
' Para evitar más complejidad si edita una venta telefónica no la pude pasar a redpagos
' y una redpagos no puede ser vta telef.

    cPagaCon.Enabled = False
    
    'If Val(lblCodigo.Tag) = TipoDocumento.VentaRedPagosTelefonicas Then cPagaCon.Enabled = False
    
    Foco tArticulo
    Exit Sub

errModificar:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error inesperado.", Trim(Err.Description)
    
End Sub

Private Sub db_CliCheque()

    If cPagaCon.ListIndex = 0 Then
        Cons = "Update Cliente Set CliCheque = 'S' Where CliCodigo = " & txtCliente.Cliente.Codigo
        cBase.Execute (Cons)
    End If

End Sub

Private Sub GraboModificacion()
Dim Msg As String, aTexto As String
Dim Control As Integer
Dim lnNro As Long
Dim sEncontre As Boolean
Dim aUsuario As Long, strDefensa As String
    
    aUsuario = 0
    Screen.MousePointer = vbHourglass
    Control = VerificoCostoArticulo
    If Control = 1 Then
        Dim objSuceso As New clsSuceso
        objSuceso.ActivoFormulario tUsuario.Tag, "Cambio de Precio", cBase
        Me.Refresh
        aUsuario = objSuceso.RetornoValor(True)
        strDefensa = objSuceso.RetornoValor(False, True)
        Set objSuceso = Nothing
        If aUsuario = 0 Then Exit Sub
    End If
    
    Dim bSucesoNombre As Boolean
    Dim sDefCambioNombre As String
    If Trim(txtCliente.Cliente.Nombre) <> Trim(tNombreC.Text) Then
        'Cambio el nombre del cliente
        bSucesoNombre = True
        'Llamo al registro del Suceso-------------------------------------------------------------
        Set objSuceso = New clsSuceso
        aUsuario = 0
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "Cambio de Nombre", cBase
        Me.Refresh
        aUsuario = objSuceso.RetornoValor(True)
        sDefCambioNombre = objSuceso.RetornoValor(False, True)
        Set objSuceso = Nothing
        If aUsuario = 0 Then Screen.MousePointer = 0: Exit Sub
    Else
        bSucesoNombre = False
    End If
    
    Screen.MousePointer = 11

    On Error GoTo ErrGFR
    FechaDelServidor
    
    cBase.BeginTrans 'Comienzo la TRANSACCION-------------------------------------------------------------------------
    On Error GoTo ErrResumo
    
    If sDiferencia Then
        Cons = "Select * From Envio Where EnvCodigo IN (" & strCodigoEnvio & ") And EnvReclamoCobro <> Null"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            sDiferencia = True
        Else
            sDiferencia = False
            RsAux.Edit
            RsAux!EnvReclamoCobro = CCur(labTotal.Caption)
            RsAux.Update
        End If
        RsAux.Close
    End If
    
    Dim rsVta As rdoResultset
    Msg = "Ocurrió un error al intentar grabar."
    Cons = "Select * From VentaTelefonica Where VTeCodigo = " & Val(tCodigo.Tag)
    Set rsVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsVta.EOF Then
        Msg = "Otra terminal pudo eliminar la venta, verifique."
        GoTo ErrResumo
    Else
        If CDate(labFecha.Tag) <> rsVta!VTeFModificacion Then
            Msg = "Otra terminal pudo modificar la venta, verifique"
            GoTo ErrResumo
        End If
    End If
    Msg = "Ocurrió un error al intentar grabar."
    
    rsVta.Edit
    rsVta!VTeTotal = CCur(labTotal.Caption)
    rsVta!VTeIva = CCur(labIVA.Caption)
    rsVta!VTeSucursal = paCodigoDeSucursal
    rsVta!VTeUsuario = tUsuario.Tag
    rsVta!VTeFModificacion = Format(gFechaServidor, sqlFormatoFH)
    If Trim(tComentarioDocumento.Text) <> vbNullString Then
        rsVta!VTeComentario = clsGeneral.SacoEnter(Trim(tComentarioDocumento.Text))
    Else
        rsVta!VTeComentario = Null
    End If
    If Trim(tTelefono.Text) <> vbNullString Then rsVta!VTeTelefonoLlamada = Trim(tTelefono.Text) Else rsVta!VTeTelefonoLlamada = Null
    If sDiferencia Then
        'Updateo el envio de cobro por el nuevo total
        rsVta!VTeCompleta = 0
    Else
        rsVta!VTeCompleta = 1
    End If
    If cDireccion.ListIndex > -1 Then rsVta!VTeDireccionFactura = cDireccion.ItemData(cDireccion.ListIndex)
    If chNomDireccion.Value = 1 Then
        rsVta!VTeNombreFactura = Trim(tNombreC.Text) & " (" & Trim(cDireccion.Text) & ")"
    Else
        If Trim(txtCliente.Cliente.Nombre) <> Trim(tNombreC.Text) Then
            rsVta!VTeNombreFactura = Trim(tNombreC.Text)
        Else
            rsVta!VTeNombreFactura = Null
        End If
    End If
    rsVta.Update
    rsVta.Close
    
    'Grabo el telefono en la BD Telefonos
    GraboDatosBDTelefono txtCliente.Cliente.Codigo
    db_CliCheque
    
    'Verifico si elimino alguno primero.
    Cons = " Select * From RenglonVtaTelefonica Where RVTVentaTelefonica = " & Val(tCodigo.Tag)
    Set RsAuxVta = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    
    Do While Not RsAuxVta.EOF
        sEncontre = False
        For Each itmx In lvVenta.ListItems
            If RsAuxVta!RVTArticulo = CLng(Mid(itmx.Key, 2, Len(itmx.Key))) Then
                sEncontre = True
                Exit For
            End If
        Next
        
        'Si no lo encuentro lo elimino.
        If Not sEncontre Then
        
            If RsAuxVta!RVTARetirar <> 0 Then
                'LE Retorno al estado sano y elimino el a retirar.
                MarcoStockVenta CLng(tUsuario.Tag), RsAuxVta!RVTArticulo, RsAuxVta!RVTARetirar * -1, 0, 0, TipoDocumento.ContadoDomicilio, Val(tCodigo.Tag), paCodigoDeSucursal
            End If
        
            Cons = "Delete RenglonVtaTelefonica " _
                & " Where RVTVentaTelefonica = " & Val(tCodigo.Tag) & " And RVTArticulo = " & RsAuxVta!RVTArticulo
            cBase.Execute (Cons)
        End If
        RsAuxVta.MoveNext
    Loop
    
    For Each itmx In lvVenta.ListItems
        Cons = " Select * From RenglonVtaTelefonica Where RVTVentaTelefonica = " & Val(tCodigo.Tag) _
            & " And RVTArticulo = " & Mid(itmx.Key, 2, Len(itmx.Key))
        
        Set RsAuxVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If RsAuxVta.EOF Then
            RsAuxVta.AddNew
            RsAuxVta!RVTVentaTelefonica = Val(tCodigo.Tag)
            RsAuxVta!RVTArticulo = Mid(itmx.Key, 2, Len(itmx.Key))
            RsAuxVta!RVtCantidad = itmx.Text
            RsAuxVta!RVtPrecio = itmx.SubItems(3)
            RsAuxVta!RVTIVA = itmx.SubItems(3) - (itmx.SubItems(3) / (1 + (itmx.SubItems(4) / 100)))
            RsAuxVta!RVTARetirar = itmx.Text
            RsAuxVta!RVTComentario = itmx.SubItems(2)
            RsAuxVta.Update
            
            'If Val(itmx.SubItems(7)) <> paTipoArticuloServicio Then
            If Not ArticuloEsServicio(Mid(itmx.Key, 2, Len(itmx.Key))) Then
                MarcoStockVenta CLng(tUsuario.Tag), CLng(Mid(itmx.Key, 2, Len(itmx.Key))), CCur(itmx.Text), CCur(itmx.Tag), 0, TipoDocumento.ContadoDomicilio, Val(tCodigo.Tag), paCodigoDeSucursal
            End If
        Else
            'If Val(itmx.SubItems(7)) <> paTipoArticuloServicio Then
            If Not ArticuloEsServicio(Mid(itmx.Key, 2, Len(itmx.Key))) Then
                'NO es de servicio --> veo si modifico la cantidad y luego si cambio cantidad a retirar.
                If RsAuxVta!RVtCantidad <> CCur(itmx.Text) Then
                    'Al modificar solo modifica los a retirar.
                    If RsAuxVta!RVTARetirar > CInt(itmx.Text) - CInt(itmx.Tag) Then
                        MarcoStockVenta CLng(tUsuario.Tag), RsAuxVta!RVTArticulo, (RsAuxVta!RVTARetirar - (CInt(itmx.Text) - CCur(itmx.Tag))) * -1, 0, 0, TipoDocumento.ContadoDomicilio, Val(tCodigo.Tag), paCodigoDeSucursal
                    ElseIf RsAuxVta!RVTARetirar < CInt(itmx.Text) - CInt(itmx.Tag) Then
                        MarcoStockVenta CLng(tUsuario.Tag), RsAuxVta!RVTArticulo, (CCur(itmx.Text) - CCur(itmx.Tag)) - RsAuxVta!RVTARetirar, 0, 0, TipoDocumento.ContadoDomicilio, Val(tCodigo.Tag), paCodigoDeSucursal
                    End If
                End If
            End If
            RsAuxVta.Edit
            If Trim(itmx.SubItems(2)) = vbNullString Then
                RsAuxVta!RVTComentario = Null
            Else
                RsAuxVta!RVTComentario = itmx.SubItems(2)
            End If
            RsAuxVta!RVtPrecio = itmx.SubItems(3)
            RsAuxVta!RVTIVA = itmx.SubItems(3) - (itmx.SubItems(3) / (1 + (itmx.SubItems(4) / 100)))
            If RsAuxVta!RVtCantidad <> CCur(itmx.Text) Then
                'Aca pudo restar o agregar artículos.
                RsAuxVta!RVTARetirar = CCur(itmx.Text) - CCur(itmx.Tag)
            End If
            RsAuxVta!RVtCantidad = itmx.Text
            RsAuxVta.Update
        End If
        RsAuxVta.Close
    Next
    If Control <> 0 Then
        aTexto = "Modificación de " & fnc_GetNombreVenta(rsVta("VteTipo")) & " Nro. :" & lnNro & " (" & Trim(gFechaServidor) & ")"
        clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.ModificacionDePrecios, paCodigoDeTerminal, aUsuario, 0, Descripcion:=aTexto, Defensa:=Trim(strDefensa), Valor:=cCambio, idCliente:=txtCliente.Cliente.Codigo
    End If
    If bSucesoNombre Then
        clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.FacturaCambioNombre, paCodigoDeTerminal, aUsuario, 0, , aTexto, Trim(sDefCambioNombre), 1, txtCliente.Cliente.Codigo
    End If
    
    cBase.CommitTrans
    
    ControloArticuloConEnvio CLng(tCodigo.Text)
    ctr_Enabled False
    BuscoVentaTelefonica tCodigo.Text
    sModificar = False
    tCodigo.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrGFR:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción."
    Exit Sub

ErrResumo:
    Resume ErrRelajo
    
ErrRelajo:
    cBase.RollbackTrans
    clsGeneral.OcurrioError Msg, Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub


Private Sub AccionEliminar()
On Error GoTo ErrAE
Dim strFecha As String, strUsuario As String

    If MsgBox("Desea anular la venta seleccionada", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        
        Screen.MousePointer = vbHourglass
        strUsuario = vbNullString
        strUsuario = InputBox("Ingrese su código de usuario.", "Emitir Factura")
        If Trim(strUsuario) = vbNullString Then
            MsgBox "No se almacenará la información.", vbInformation, "ATENCIÓN"
            Exit Sub
        Else
            If Not IsNumeric(strUsuario) Then
                MsgBox "El formato ingresado no es numérico.", vbExclamation, "ATENCIÓN"
                Exit Sub
            Else
                strUsuario = GetUIDCodigo(CLng(strUsuario))
            End If
        End If
        
        If Val(strUsuario) = 0 Then
            MsgBox "Usuario incorrecto.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        FechaDelServidor
        
        cBase.BeginTrans
        On Error GoTo ErrResumo
        
        Dim rsVta As rdoResultset
        Cons = "Select * From VentaTelefonica Where VTeCodigo = " & Val(tCodigo.Tag)
        Set rsVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        'Se la pelaron.
        If rsVta.EOF Then
            cBase.RollbackTrans
            rsVta.Close
            Screen.MousePointer = vbDefault
            MsgBox "Otra terminal elimino la venta, verifique.", vbExclamation, "ATENCIÓN"
            BuscoVentaTelefonica tCodigo.Text: Exit Sub
        End If
        
        'Se la modificaron.
        If CDate(labFecha.Tag) <> rsVta!VTeFModificacion Then
            cBase.RollbackTrans
            rsVta.Close
            Screen.MousePointer = vbDefault
            MsgBox "Otra terminal modifico la venta, verifique.", vbExclamation, "ATENCIÓN"
            BuscoVentaTelefonica tCodigo.Text: Exit Sub
        End If
        
        'Borro si tiene envíos.-------------------------------
        'Para cada artículo que este en envío le hago el movimiento de stock
'        Cons = "Select * From RenglonEnvio Where REvEnvio IN (" _
'            & "Select EnvCodigo From Envio Where EnvTipo = " & TipoEnvio.Cobranza _
'            & " And EnvDocumento = " & rsVta!VTeCodigo & ")" _
'            & " And REvArticulo Not IN (Select ArtID From Articulo Where ArtTipo = " & paTipoArticuloServicio & ")"
'
'        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'        Do While Not RsAux.EOF
'            MarcoStockVenta CLng(strUsuario), RsAux!REvArticulo, 0, RsAux!REvAEntregar * -1, 0, rsVta("VTeTipo"), rsVta!VTeCodigo, paCodigoDeSucursal
'            RsAux.MoveNext
'        Loop
'        RsAux.Close


        If rsVta("VTeTipo") = 44 Then
            'Not ArticuloEsServicio(Mid(itmx.Key, 2, Len(itmx.Key)))
            'Cons = "Select * From RenglonVtaTelefonica " _
                & " Where RVTVentaTelefonica = " & rsVta!vtecodigo _
                & " And RVTArticulo Not IN (Select ArtID From Articulo Where ArtTipo = " & paTipoArticuloServicio & ")"
            Cons = "Select * From RenglonVtaTelefonica INNER JOIN Articulos ON ArtID = RVTArticulo AND ArtMueveStock = 1 " _
                & " Where RVTVentaTelefonica = " & rsVta!vtecodigo
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            Do While Not RsAux.EOF
                If RsAux!RVTARetirar > 0 Then
                    MarcoStockVenta CLng(strUsuario), RsAux!RVTArticulo, RsAux!RVTARetirar * -1, 0, 0, TipoDocumento.ContadoDomicilio, rsVta!vtecodigo, paCodigoDeSucursal
                End If
                RsAux.MoveNext
            Loop
            RsAux.Close
        End If
        
        Dim listaNotas As Collection
        Dim oEnvio As New clsEnvio
        Dim sError As String
        
        Cons = "SELECT EnvCodigo, EnvFModificacion FROM Envio Where  EnvTipo = " & TipoEnvio.Cobranza _
            & " And EnvDocumento = " & rsVta!vtecodigo
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            sError = oEnvio.EliminarEnvio(cBase, RsAux("EnvCodigo"), RsAux("EnvFModificacion"), Val(strUsuario), paCodigoDeSucursal, listaNotas)
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        
        If rsVta("VTeTipo") <> 44 Then
            'Cons = "Select * From RenglonVtaTelefonica " _
                & " Where RVTVentaTelefonica = " & rsVta!vtecodigo _
                & " And RVTArticulo Not IN (Select ArtID From Articulo Where ArtTipo = " & paTipoArticuloServicio & ")"
            
            Cons = "Select * From RenglonVtaTelefonica INNER JOIN Articulos ON ArtID = RVTArticulo AND ArtMueveStock = 1 " _
                & " Where RVTVentaTelefonica = " & rsVta!vtecodigo
                
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            Do While Not RsAux.EOF
                If RsAux!RVTARetirar > 0 Then
                    MarcoStockVenta CLng(strUsuario), RsAux!RVTArticulo, RsAux!RVTARetirar * -1, 0, 0, TipoDocumento.ContadoDomicilio, rsVta!vtecodigo, paCodigoDeSucursal
                End If
                RsAux.MoveNext
            Loop
            RsAux.Close
        End If
        
        
'        Cons = " Delete RenglonEnvio Where REvEnvio IN (" _
'            & "Select EnvCodigo From Envio Where EnvTipo = " & TipoEnvio.Cobranza _
'            & " And EnvDocumento = " & rsVta!VTeCodigo & ")"
'        cBase.Execute (Cons)
'
'        Cons = " Delete Envio Where  EnvTipo = " & TipoEnvio.Cobranza _
'            & " And EnvDocumento = " & rsVta!VTeCodigo
'        cBase.Execute (Cons)
        
'        Cons = "Delete Direccion Where DirCodigo IN(" _
'            & "Select EnvDireccion From Envio Where EnvTipo = " & TipoEnvio.Cobranza _
'            & " And EnvDocumento = " & rsVta!VTeCodigo & ")"
'        cBase.Execute (Cons)
        
        'Instalaciones la anulo.
        Cons = "Select * from Instalacion Where InsTipoDocumento = 2 And InsDocumento = " & rsVta!vtecodigo
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Edit
            RsAux!InsAnulada = Format(gFechaServidor, "yyyy-mm-dd hh:nn")
            RsAux!InsFechaModificacion = Format(gFechaServidor, "yyyy/mm/dd hh:mm:ss")
            RsAux.Update
        End If
        RsAux.Close
        
        'Si tiene específicos los libero.
        cBase.Execute "UPDATE ArticuloEspecifico SET AEsTipoDocumento = Null, AEsDocumento = Null WHERE AEsTipoDocumento = 7 AND AEsDocumento = " & rsVta!vtecodigo
        
        'where instipodocumento = 2
        '................................................
        
        rsVta.Close
        'Si tiene envío me edita la vta así que salta por concurrencia así que la vuelvo a cargar.
        Cons = "Select * From VentaTelefonica Where VTeCodigo = " & Val(tCodigo.Tag)
        Set rsVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        rsVta.Edit
        rsVta!VTeAnulado = Format(gFechaServidor, sqlFormatoFH)
        rsVta("VTeUsuario") = strUsuario
        rsVta.Update
        rsVta.Close
        
        cBase.CommitTrans
        
        BuscoVentaTelefonica tCodigo.Text
        Screen.MousePointer = vbDefault
    End If
    Exit Sub

ErrAE:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al intentar eliminar la venta."
    Exit Sub
    
ErrResumo:
    Resume ErrTrans
    
ErrTrans:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al intentar eliminar la venta.", Err.Description
    rsVta.Requery
End Sub

'Private Sub AccionFacturar()
'Dim strUsuario As String, aTexto As String
'Dim lnDoc As Long
'Dim dFecha As Date
'Dim cCofisG As Currency, cCofisNG As Currency, bCofis As Boolean
'
'
'    If Trim(paDContado) = "" Then   'Verifico si hay un documento asociado para facturar (está cargado el paDContado)
'        MsgBox "Su sucursal no tiene un documento asociado para emitir boletas contado.", vbCritical, "ATENCIÓN"
'        Exit Sub
'    End If
'
'    If MsgBox("¿Confirma emitir la factura contado para esta venta?", vbQuestion + vbYesNo, "FACTURAR") = vbYes Then
'        strUsuario = vbNullString
'        strUsuario = InputBox("Ingrese su dígito de usuario.", "Emitir Factura")
'        If Trim(strUsuario) = vbNullString Then
'            MsgBox "No se almacenará la información.", vbInformation, "ATENCIÓN"
'            Exit Sub
'        Else
'            If Not IsNumeric(strUsuario) Then
'                MsgBox "El formato ingresado no es numérico.", vbExclamation, "ATENCIÓN"
'                Exit Sub
'            Else
'                strUsuario = GetUIDCodigo(Val(strUsuario))
'                If Trim(strUsuario) = vbNullString Then MsgBox "No se encontro un usuario con ese dígito.", vbExclamation, "ATENCIÓN": Exit Sub
'            End If
'        End If
'        Screen.MousePointer = vbHourglass
'        On Error GoTo ErrAF
'
'        cBase.BeginTrans
'        On Error GoTo ErrResumir
'
'        FechaDelServidor
'
'        Dim rsVta As rdoResultset
'        Cons = "Select * From VentaTelefonica Where VTeCodigo = " & Val(tCodigo.Tag)
'        Set rsVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'
'        If CDate(labFecha.Tag) <> rsVta!VTeFModificacion Then
'            cBase.RollbackTrans
'            MsgBox "Otra terminal pudo modificar la venta, verifique.", vbInformation, "INFORMACIÓN"
'            rsVta.Close
'            BuscoVentaTelefonica tCodigo.Text
'            Exit Sub
'        End If
'
'        'Tengo que hacer la factura Contado.
'        lnDoc = InsertoDocumento(CCur(labTotal.Caption), CCur(labIVA.Caption), CLng(strUsuario))
'
'        'Inserto los artículos en la tabla renglon.
'        CopioTablaArticulos lnDoc, Val(tCodigo.Tag), cCofisG, bCofis
'        cCofisG = Format(cCofisG, "###0.00")
'
'        If bCofis Then
'            Cons = "Update Documento set DocCofis = " & Format(cCofisG, "###0.00") & " Where DocCodigo = " & lnDoc
'            cBase.Execute (Cons)
'        End If
'
'        'Updateo el envío y lo asigno al documento factura.
'        Cons = "Update Envio Set EnvDocumento = " & lnDoc & ", EnvTipo = " & TipoEnvio.Entrega & ", EnvReclamoCobro = Null " _
'            & " Where EnvCodigo IN (" & strCodigoEnvio & ")"
'        cBase.Execute (Cons)
'
'        'Modifico la tabla ventatelefonica, le pongo el código de documento.
'        rsVta.Edit
'        rsVta!VTeFModificacion = Format(gFechaServidor, sqlFormatoFH)
'        rsVta!VTeDocumento = lnDoc
'        rsVta.Update
'        rsVta.Close
'
'        cBase.CommitTrans '.........................
'        'Imprimo la Factura.
'        ImprimoDocumentoContado lnDoc, aTexto
'
'        BuscoVentaTelefonica tCodigo.Text
'        Screen.MousePointer = vbDefault
'    End If
'    Exit Sub
'
'ErrAF:
'    Screen.MousePointer = vbDefault
'    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transacción."
'    Exit Sub
'
'ErrResumir:
'    Resume Resumir
'
'Resumir:
'    cBase.RollbackTrans
'    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transacción.", Err.Description
'    rsVta.Requery
'    Screen.MousePointer = vbDefault
'End Sub

Private Sub AccionValidar()
On Error GoTo ErrAV

    If strCodigoEnvio <> vbNullString Then
        Cons = "Select SUM(EnvReclamoCobro) From Envio Where EnvCodigo IN (" & strCodigoEnvio & ")"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not IsNull(RsAux(0)) Then
            If RsAux(0) <> CCur(labTotal.Tag) Then
                RsAux.Close
                MsgBox "No coincide el total con lo que tienen los envíos a cobrar.", vbExclamation, "ATENCIÓN"
                Exit Sub
            Else
                Screen.MousePointer = vbHourglass
                RsAux.Close
                
                FechaDelServidor
                
                Dim rsVta As rdoResultset
    
                Cons = "Select * From VentaTelefonica Where VTeCodigo = " & Val(tCodigo.Tag)
                Set rsVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                rsVta.Edit
                rsVta!VTeCompleta = 1
                rsVta!VTeFModificacion = Format(gFechaServidor, sqlFormatoFH)
                rsVta.Update
                rsVta.Close
                
                'Verifico la cantidad de los artículos para retirar.
                BuscoVentaTelefonica Val(tCodigo.Tag)
                
            End If
        Else
            RsAux.Close
            MsgBox "No existe un envío cobranza.", vbExclamation, "ATENCIÓN"
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrAV:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error inesperado al actualizar la venta."
    rsVta.Requery
End Sub

'Private Function InsertoDocumento(cTotal As Currency, cIva As Currency, lnUsuario As Long) As Long
'Dim aTexto As String
'Dim Rs As rdoResultset
'
'    aTexto = NumeroDocumento(paDContado)
'
'    Cons = "INSERT INTO Documento " _
'        & " (DocFecha, DocTipo, DocSerie, DocNumero, DocCliente, DocMoneda, DocTotal, DocIva, DocAnulado, DocSucursal, DocUsuario, DocFModificacion)" _
'        & " Values ('" & Format(gFechaServidor, sqlFormatoFH) & "'" & ", " & TipoDocumento.Contado
'
'    Cons = Cons _
'        & ", '" & Mid(aTexto, 1, 1) & "', " & Mid(aTexto, 2, Len(aTexto)) _
'        & ", " & txtCliente.Cliente.Codigo & ", " & cMoneda.ItemData(cMoneda.ListIndex) _
'        & ", " & cTotal & ", " & cIva
'
'    Cons = Cons _
'        & ", 0, " & paCodigoDeSucursal & ", " & lnUsuario _
'        & ", '" & Format(gFechaServidor, sqlFormatoFH) & "')"
'    cBase.Execute (Cons)
'
'    Cons = "SELECT MAX(DocCodigo) From Documento" _
'        & " WHERE DocTipo = " & TipoDocumento.Contado _
'        & " AND DocSerie = '" & Mid(aTexto, 1, 1) & "'" _
'        & " AND DocNumero = " & Mid(aTexto, 2, Len(aTexto))
'
'    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'    InsertoDocumento = Rs(0)
'    Rs.Close
'
'End Function

Private Sub CopioTablaArticulos(ByVal lnDocumento As Long, ByVal idVta As Long, cCofisG As Currency, bCofis As Boolean)
Dim cAux As Currency
    
    bCofis = False 'LlevaCofis
    cCofisG = 0
    
    Cons = "Select * From RenglonVtaTelefonica, Articulo Where RVTVentaTelefonica = " & idVta _
        & " And ArtID = RVTArticulo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        
        Cons = "INSERT INTO Renglon (RenDocumento, RenArticulo, RenCantidad, RenPrecio, RenIVA, RenARetirar, RenCofis)" _
            & " VALUES (" & lnDocumento & ", " & RsAux!RVTArticulo & ", " & RsAux!RVtCantidad _
            & ", " & RsAux!RVtPrecio & ", " & RsAux!RVTIVA & ", " & RsAux!RVTARetirar
        
        
        If bCofis Then
            'Neto
            cAux = RsAux!RVtPrecio - RsAux!RVTIVA
            'Aplico cofis.
            cAux = cAux - (cAux / (1 + (paCofis / 100)))
            cAux = Format(cAux, "###0.00")
            'Total del cofis.
            cCofisG = cCofisG + (cAux * RsAux!RVtCantidad)
            
            Cons = Cons & ", " & Format(cAux, "###0.00") & ")"
        Else
            Cons = Cons & ", Null)"
        End If
        cBase.Execute (Cons)
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

Private Sub AccionVerFactura()
    EjecutarApp App.Path & "\Detalle de factura", Val(tCodigo.Tag)
End Sub

Private Sub CargoTelefonos(Cliente As Long, Tipo As Long)

Dim RsTel As rdoResultset
    
    On Error GoTo ErrNT
    Screen.MousePointer = 11
    Cons = " Select TelTipo, TelNumero, TelInterno " _
            & " From Telefono " _
            & " Where TelCliente = " & Cliente _
            & " And TelTipo = " & Tipo
    Set RsTel = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsTel.EOF Then
        BuscoCodigoEnCombo cTipoTelefono, Tipo
        tTelefono.Text = Trim(RsTel!TelNumero)
        If Not IsNull(RsTel!TelInterno) Then tInterno.Text = Trim(RsTel!TelInterno) Else: tInterno.Text = ""
    End If
    RsTel.Close
    Screen.MousePointer = 0
    Exit Sub
        
ErrNT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al cargar los números de teléfonos."
End Sub

Private Sub GraboDatosBDTelefono(Cliente As Long)

    If cTipoTelefono.ListIndex = -1 Then Exit Sub
    If Trim(tTelefono.Text) = "" Then Exit Sub
    
Dim RsTel As rdoResultset
    
    Cons = " Select * From Telefono " _
            & " Where TelCliente = " & Cliente _
            & " And TelTipo = " & cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
    Set RsTel = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsTel.EOF Then
        RsTel.Edit
        RsTel!TelNumero = Trim(tTelefono.Text)
        If Trim(tInterno.Text) <> "" Then RsTel!TelInterno = Trim(tInterno.Text) Else: RsTel!TelInterno = Null
        RsTel.Update
    Else
        RsTel.AddNew
        RsTel!TelCliente = Cliente
        RsTel!TelTipo = cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
        RsTel!TelNumero = Trim(tTelefono.Text)
        If Trim(tInterno.Text) <> "" Then RsTel!TelInterno = Trim(tInterno.Text) Else: RsTel!TelInterno = Null
        RsTel.Update
    End If
    RsTel.Close

End Sub
Private Function VerificoCostoArticulo() As Integer
Dim sControl As Boolean

    'RETORNO
        '0 = No hay cambios de precio
        
    VerificoCostoArticulo = 0
    sControl = False
    'Para cada artículo de la factura veo si posee descuento el cliente y verifico su costo.
    For Each itmx In lvVenta.ListItems
        If Mid(itmx.Key, 1, 1) = "Z" Then
            sControl = True
            cCambio = CCur(itmx.Text) * (CCur(itmx.SubItems(3)) - CCur(itmx.SubItems(6)))
        End If
    Next
    If sControl Then VerificoCostoArticulo = 1

End Function
'Private Sub ImprimoDocumentoContado(Documento As Long, ByVal NroDoc As String)
'On Error GoTo ErrCrystal
'Dim Result As Integer, JobSRep1 As Integer, JobSRep2 As Integer, jobnum As Integer
'Dim NombreFormula As String, CantForm As Integer
'Dim TableType As PETableType, LogOnInfo As PELogOnInfo
'Dim aTexto As String
'
'    Screen.MousePointer = 11
'    'Inicializo el Reporte y SubReportes
'    jobnum = crAbroReporte(gPathListados & "Contado.RPT")
'    If jobnum = 0 Then GoTo ErrCrystal
'
'    'Configuro la Impresora
'    If Trim(Printer.DeviceName) <> Trim(paIContadoN) Then SeteoImpresoraPorDefecto paIContadoN
'    If Not crSeteoImpresora(jobnum, Printer, paIContadoB) Then GoTo ErrCrystal
'
'    'Obtengo la cantidad de formulas que tiene el reporte.
'    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
'    If CantForm = -1 Then GoTo ErrCrystal
'
'    'Cargo Propiedades para el reporte Contado --------------------------------------------------------------------------------
'    For i = 0 To CantForm - 1
'        NombreFormula = crObtengoNombreFormula(jobnum, i)
'
'        Select Case LCase(NombreFormula)
'            Case "": GoTo ErrCrystal
'            Case "nombredocumento": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDContado & "'")
'            Case "cliente"
'                If Trim(tCi.Text) <> FormatoCedula Then aTexto = "(" & tCi.Text & ")" Else aTexto = ""
'                If chNomDireccion.Value = 1 And Trim(labDireccion.Caption) <> "" Then aTexto = aTexto & " (" & Trim(cDireccion.Text) & ")"
'                Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(tNombreC.Text) & " " & Trim(aTexto) & "'")
'
'            Case "direccion": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labDireccion.Caption) & "'")
'            Case "ruc": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(clsGeneral.RetornoFormatoRuc(tRuc.Text)) & "'")
'            Case "codigobarras": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.Contado, Documento) & "'")
'            Case "signomoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(cMoneda.Text) & "'")
'            Case "nombremoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'(" & BuscoNombreMoneda(cMoneda.ItemData(cMoneda.ListIndex)) & ")'")
'
'            Case "textoretira"
'                If Trim(strCodigoEnvio) = "" Then
'                    aTexto = "RETIRA "
'                Else
'                    aTexto = "HAY ENVIOS DE MERCADERIA"
'                End If
'                Result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
'
'            Case Else: Result = 1
'        End Select
'        If Result = 0 Then GoTo ErrCrystal
'    Next
'    '--------------------------------------------------------------------------------------------------------------------------------------------
'
'    'Seteo la Query del reporte-----------------------------------------------------------------
'    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
'            & " From " & paBD & ".dbo.Documento Documento " _
'            & " Where DocCodigo = " & Documento
'    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
'
'    'Subreporte srContado.rpt  y srContado.rpt - 01-----------------------------------------------------------------------------
'    JobSRep1 = crAbroSubreporte(jobnum, "srContado.rpt")
'    If JobSRep1 = 0 Then GoTo ErrCrystal
'
'    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
'            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
'                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"
'    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
'
'    JobSRep2 = crAbroSubreporte(jobnum, "srContado.rpt - 01")
'    If JobSRep2 = 0 Then GoTo ErrCrystal
'    If crSeteoSqlQuery(JobSRep2, Cons) = 0 Then GoTo ErrCrystal
'    '-------------------------------------------------------------------------------------------------------------------------------------
'
'    Result% = PEGetNthTableType(jobnum, 1, TableType)
'    Result% = crPELogOnServer("PDSDBC.DLL", "", "", "", "")
'
'
'    'If crMandoAPantalla(JobNum, "Factura Contado") = 0 Then GoTo ErrCrystal
'    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
'    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
'
'    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
'    If Not crCierroSubReporte(JobSRep2) Then GoTo ErrCrystal
'
'    'crEsperoCierreReportePantalla
'
'    Screen.MousePointer = 0
'    Exit Sub
'
'ErrCrystal:
'    Screen.MousePointer = 0
'    clsGeneral.OcurrioError crMsgErr & " Nro. Documento = " & Mid(NroDoc, 1, 1) & " " & CLng(Trim(Mid(NroDoc, 2, Len(NroDoc))))
'    On Error Resume Next
'    Screen.MousePointer = 11
'    crCierroSubReporte JobSRep1
'    crCierroSubReporte JobSRep2
'    Screen.MousePointer = 0
'    Exit Sub
'
'End Sub

Private Function CambioClienteEnvios(Cliente As Long, Envios As String) As Boolean
On Error GoTo ErrCCE
    
    CambioClienteEnvios = True
    If InStr(Envios, ",") = 0 Then
        If Envios = "0" Then Exit Function
    End If
    
    Cons = "Update Envio Set EnvCliente = " & Cliente _
            & " Where EnvCodigo IN (" & Envios & ")"
    cBase.Execute (Cons)
    Exit Function

ErrCCE:
    clsGeneral.OcurrioError "Ocurrió un error al intentar modificar el cliente en los envíos."
    CambioClienteEnvios = False
End Function

Private Function NumeroAuxiliarEnvio() As Integer
Dim idAux As Integer
    
    NumeroAuxiliarEnvio = 0
    On Error GoTo ErrBT
    idAux = Autonumerico(TAutonumerico.AuxiliarEnvio)
    cBase.BeginTrans
    On Error GoTo ErrRB
    For Each itmx In lvVenta.ListItems
        'Si tiene X en la clave es porque es un artículo que factura envío.
        If InStr(1, itmx.Key, "C") = 0 And Mid(itmx.Key, 1, 1) <> "X" And Not ArticuloEsServicio(Mid(itmx.Key, 2, Len(itmx.Key))) Then     'And Val(itmx.SubItems(7)) <> paTipoArticuloServicio Then
            Cons = "Insert into EnvioAuxiliar (EAuID, EAuArticulo, EAuCantidad) Values (" & idAux & ", " & Mid(itmx.Key, 2, Len(itmx.Key)) & ", " & itmx.Text & ")"
            cBase.Execute (Cons)
        End If
    Next
    cBase.CommitTrans
    NumeroAuxiliarEnvio = idAux
    Exit Function

ErrBT:
    clsGeneral.OcurrioError "Ocurrió un error al intentar abrir la transacción.", Err.Description
    Screen.MousePointer = 0
ErrRB:
    Resume ErrResumo
ErrResumo:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al insertar los artículos para enviar.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub CargoDireccionesAuxiliares(aIdCliente As Long)
Dim lDPcpal As Long, sNFactura As String
    On Error GoTo errCDA
    Dim rsDA As rdoResultset
    
    If gDirFactura > 0 Then lDPcpal = gDirFactura: gDirFactura = 0
    
    'Direcciones Auxiliares-----------------------------------------------------------------------
    
    'Le incluyo el order by para que me cargue si o si tiene dirección de facturación.
    
    Cons = "Select Top 16 * from DireccionAuxiliar Where DAuCliente = " & aIdCliente & _
            " Order by DAuFactura Desc"
    
    Set rsDA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDA.EOF Then
        Do While Not rsDA.EOF
            
            With cDireccion
                .AddItem Trim(rsDA!DAuNombre)
                .ItemData(.NewIndex) = rsDA!DAuDireccion
            End With
            
            'Si es la seleccionada para facturar.
            If rsDA!DAuFactura Then
                gDirFactura = rsDA!DAuDireccion
                sNFactura = Trim(rsDA!DAuNombre)
            End If
            rsDA.MoveNext
        Loop
    End If
    rsDA.Close
    
    With cDireccion
        If .ListCount > 1 Then cDireccion.BackColor = Colores.Blanco
        
        If .ListCount > IIf(lDPcpal > 0, 16, 15) Then
            'Cumple Top --> borro pongo nuevamente la dirección del cliente y la opción buscar.
            .Clear
            
            If lDPcpal > 0 Then
                .AddItem "Dirección Principal": .ItemData(.NewIndex) = lDPcpal
                .Tag = lDPcpal
            End If
                        
            If gDirFactura > 0 Then
                .AddItem sNFactura
                .ItemData(.NewIndex) = gDirFactura
            End If
            
            'Para buscar.
            .AddItem cte_KeyFindDir
            .ItemData(.NewIndex) = -1
            
        End If
        
        If gDirFactura = 0 And lDPcpal > 0 Then gDirFactura = lDPcpal

        If gDirFactura <> 0 Then BuscoCodigoEnCombo cDireccion, gDirFactura
       
    End With
    
errCDA:
End Sub

'Private Function LlevaCofis() As Boolean
'    LlevaCofis = False
'    If Trim(tRuc.Text) = "" Then
'        'Busco si el cliente es empresa si es empresa estatal.
'        If Mid(tNombreC.Tag, 1, 1) = "P" Then Exit Function
'    End If
'    LlevaCofis = True
'    '------------------------------------------------------------------------
'End Function

Private Sub s_InitVarRenglon()
    
    With miRenglon
        .PrecioBonificacion = 0
        .CodArticulo = 0
        .IDCombo = 0
        .EsInhabilitado = False
        .idArticulo = 0
        .Precio = 0
        .Tipo = 0
        .PrecioOriginal = 0
        .Especifico = 0
        .DescuentoEspecifico = 0
    End With
End Sub

Private Function PrecioArticulo(ByVal lArticulo As Long, ByVal idMoneda As Long, cPrecio As Currency) As Boolean
On Error GoTo errPA
Dim rsPrecio As rdoResultset
Dim cTC As Currency

    PrecioArticulo = False
    cPrecio = 0
    Cons = "Select * From PrecioVigente Where PViArticulo = " & lArticulo _
        & " And PViMoneda = " & idMoneda & " And PViTipoCuota = " & paTipoCuotaContado
    Set rsPrecio = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsPrecio.EOF Then
        If rsPrecio!PViHabilitado Then cPrecio = rsPrecio!PViPrecio: PrecioArticulo = True
    End If
    rsPrecio.Close
    
    
    m_Patron = dis_arrMonedaProp(idMoneda, pRedondeo)
    cPrecio = Redondeo(cPrecio, m_Patron)
    
    If PrecioArticulo Or idMoneda <> paMonedaPesos Then Exit Function
    
    'Si la moneda es pesos y no tengo precio, busco el precio en dolares.
    Cons = "Select * From PrecioVigente Where PViArticulo = " & lArticulo _
        & " And PViMoneda = " & paMonedaDolar & " And PViTipoCuota = " & paTipoCuotaContado
    Set rsPrecio = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsPrecio.EOF Then
        If rsPrecio!PViHabilitado Then
            cPrecio = rsPrecio!PViPrecio
            cTC = TasadeCambio(paMonedaDolar, CInt(idMoneda), Date)
            cPrecio = cPrecio * cTC
        End If
        PrecioArticulo = True
    End If
    rsPrecio.Close
    
    m_Patron = dis_arrMonedaProp(idMoneda, pRedondeo)
    cPrecio = Redondeo(cPrecio, m_Patron)
    Exit Function
    
errPA:
    clsGeneral.OcurrioError "Ocurrió el siguiente error al buscar el precio vigente del artículo con ID: " & lArticulo & ".", Err.Description
End Function

Private Sub s_PresentoPrecio()
On Error GoTo errPP
Dim cAuxPrecio As Currency, cSumaPrecio As Currency
    
    cSumaPrecio = 0
    cSumaPrecio = BuscoDescuentoCliente(miRenglon.idArticulo, Val(labDireccion.Tag), miRenglon.PrecioOriginal, Val(tCantidad.Text))
    tUnitario.Text = Format(CCur(cSumaPrecio), "#,##0.00")
    Exit Sub
    
errPP:
    clsGeneral.OcurrioError "Ocurrió al calcular el precio unitario.", Err.Description, "Error en Presentar Precio"
End Sub

Private Sub s_InsertCombo()
On Error GoTo errIAC
Dim cAuxPrecio As Currency, idAux As Long
Dim cComboParcial As Currency
    
    Cons = "SELECT * FROM ArticulosDelCombo INNER JOIN Articulo ON ACoArticulo = ArtID " & _
            " INNER JOIN PrecioVigente ON ACoArticulo = PViArticulo AND PViHabilitado = 1 AND PViTipoCuota = " & paTipoCuotaContado & _
            " AND PViMoneda = " & paMonedaPesos & _
            " WHERE ACoCombo = " & miRenglon.IDCombo
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If Ingresado(RsAux!ArtId) Then
            MsgBox "El artículo " & Trim(RsAux!ArtNombre) & " ya está ingresado, no podrá ingresar el combo.", vbExclamation, "ATENCIÓN"
            RsAux.Close
            Exit Sub
        End If
        RsAux.MoveNext
    Loop
    RsAux.MoveFirst
    
    cComboParcial = 0
    
    Do While Not RsAux.EOF
    
        cAuxPrecio = (((RsAux("PViPrecio") * RsAux("ACoPorcPrecio")) / 100))
        cAuxPrecio = BuscoDescuentoCliente(RsAux!ArtId, Val(labDireccion.Tag), cAuxPrecio, Val(tCantidad.Text) * RsAux!ACoCantidad)
        cAuxPrecio = CCur(cAuxPrecio) * RsAux!ACoCantidad
        cComboParcial = cComboParcial + cAuxPrecio
'        CargoArticuloEnGrilla RsAux!ArtId, RsAux!ArtTipo, CInt(tCantidad.Text) * RsAux!ACoCantidad, Articulo, miRenglon.Precio, RsAux!ArtNombre, "", cAuxPrecio, cEnvio.Text, False, RsAux!ArtCodigo, 0
        s_InsertArticulo RsAux!ArtId, RsAux!ArtTipo, CCur(tCantidad.Text) * RsAux!ACoCantidad, RsAux!ArtNombre, miRenglon.Precio, cAuxPrecio
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Sub
    
errIAC:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos del combo.", Err.Description
End Sub

Private Sub s_InsertArticulo(ByVal lIDArt As Long, ByVal lTipo As Long, ByVal cQ As Currency, ByVal sNameArt As String, ByVal cPrecioBD As Currency, ByVal cPrecioUID As Currency, Optional sComentario As String = "", Optional bEsBonifica As Boolean = False, Optional esFlete As Boolean = False, Optional IDEspecifico As Long = 0)
        
    If esFlete Then
        Set itmx = lvVenta.ListItems.Add(, "X" & lIDArt, cQ)
    Else
        If cPrecioUID <> cPrecioBD And cPrecioBD <> 0 Then
            Set itmx = lvVenta.ListItems.Add(, "Z" & lIDArt, cQ)
        Else
            If bEsBonifica Then
                Set itmx = lvVenta.ListItems.Add(, "C" & lIDArt, cQ)
            Else
                Set itmx = lvVenta.ListItems.Add(, "A" & lIDArt, cQ)
            End If
        End If
    End If
    
    With itmx
        .SubItems(6) = cPrecioBD
        .SubItems(1) = Trim(sNameArt)
        .SubItems(2) = Trim(sComentario)
        .SubItems(3) = Format(cPrecioUID, "#,#00.00")
        .SubItems(4) = IVAArticulo(lIDArt)
        .SubItems(5) = Format(CCur(.SubItems(3)) * cQ, "#,##0.00")
        .SubItems(7) = lTipo
        .SubItems(8) = IDEspecifico
        .SubItems(9) = miRenglon.CantidadAlXMayor
        itmx.Tag = "0"      'En el tag de cada item dejo la cantidad de envíos que tiene.

        labIVA.Caption = Format(CCur(labIVA.Caption) + CCur(.SubItems(5)) - (CCur(.SubItems(5)) / CCur(1 + (CCur(.SubItems(4)) / 100))), "#,##0.00")
        labTotal.Caption = Format(CCur(labTotal.Caption) + CCur(.SubItems(5)), "#,##0.00")
    End With
    labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
    If (lblFlete.Caption = "") Then lblFlete.Caption = "0.00"
    lblTotalCflete.Caption = Format(CCur(lblFlete.Caption) + CCur(labTotal.Caption), "#,##0.00")
    ValidoVentaLimitadaPorFila itmx
    
End Sub

Private Sub ValidoVentaLimitadaPorFila(ByVal lvItmx As ListItem)

    With lvItmx
        '.Bold = False
        .ForeColor = vbBlack
        'If Val(.SubItems(7)) <> paTipoArticuloServicio Then
        If Not ArticuloEsServicio(Mid(.Key, 2, Len(.Key))) Then
            If Val(.SubItems(9)) = 0 And InStr(1, paCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0 Then
                .ForeColor = &HFF&
                '.Bold = True
            ElseIf Val(.SubItems(9)) > 1 And Val(.SubItems(9)) < Val(.Text) Then
                '.Bold = True
                .ForeColor = &HFF&
            End If
        End If
    End With
    
End Sub


Private Function GetUIDIdentificacion(ByVal Digito As Long, ByVal Codigo As Long) As String
Dim RsUsr As rdoResultset
On Error GoTo ErrBUD
    GetUIDIdentificacion = ""
    Cons = "Select * from Usuario"
    If Digito > 0 Then
        Cons = Cons & " Where UsuDigito = " & Digito
    Else
        Cons = Cons & " Where UsuCodigo = " & Codigo
    End If
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsUsr.EOF Then GetUIDIdentificacion = RsUsr!UsuIdentificacion
    RsUsr.Close
    Exit Function
ErrBUD:
    MsgBox "Error al buscar el usuario." & vbCr & vbCr & "Error: " & Err.Description, vbCritical, "ATENCIÓN"
End Function

Private Function GetUIDCodigo(ByVal Digito As Long) As Long
Dim RsUsr As rdoResultset
On Error GoTo ErrBUD
    GetUIDCodigo = 0
    Cons = "Select * from Usuario Where UsuDigito = " & Digito
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsUsr.EOF Then GetUIDCodigo = RsUsr!UsuCodigo
    RsUsr.Close
    Exit Function
ErrBUD:
    MsgBox "Error al buscar el usuario." & vbCr & vbCr & "Error: " & Err.Description, vbCritical, "ATENCIÓN"
End Function
Private Function fnc_DireccionInsertada(ByVal lID As Long) As Boolean
Dim iQ As Integer
    fnc_DireccionInsertada = False
    For iQ = 0 To cDireccion.ListCount - 1
        If cDireccion.ItemData(iQ) = lID Then fnc_DireccionInsertada = True: Exit Function
    Next
End Function
Private Sub loc_FindDireccionAuxiliarTexto()
On Error GoTo errFDA
Dim rsD As rdoResultset
    Cons = "Select DAuDireccion , DAuNombre " & _
                "From DireccionAuxiliar Where DAuCliente = " & txtCliente.Cliente.Codigo & _
                " And DAuNombre Like '" & Replace(cDireccion.Text, " ", "%") & "%'"
    Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsD.EOF Then
        rsD.MoveNext
        If rsD.EOF Then
            rsD.MoveFirst
            If fnc_DireccionInsertada(rsD("DAuDireccion")) Then
                BuscoCodigoEnCombo cDireccion, rsD("DAuDireccion")
            Else
                'Inserto
                With cDireccion
                    .AddItem Trim(rsD("DAuNombre"))
                    .ItemData(.NewIndex) = rsD("DAuDireccion")
                    .ListIndex = .NewIndex
                End With
            End If
            chNomDireccion.Value = 1
        Else
            Dim objLista As New clsListadeAyuda
            With objLista
                If .ActivarAyuda(cBase, Cons, 3000, 1, "Dirección Auxiliar") > 0 Then
                    If fnc_DireccionInsertada(.RetornoDatoSeleccionado(0)) Then
                        BuscoCodigoEnCombo cDireccion, .RetornoDatoSeleccionado(0)
                    Else
                        cDireccion.AddItem Trim(.RetornoDatoSeleccionado(1))
                        cDireccion.ItemData(cDireccion.NewIndex) = .RetornoDatoSeleccionado(0)
                        cDireccion.ListIndex = cDireccion.NewIndex
                    End If
                    chNomDireccion.Value = 1
                End If
            End With
            Set objLista = Nothing
        End If
        rsD.Close
    Else
        rsD.Close
        MsgBox "No hay coincidencias.", vbInformation, "Buscar dirección auxiliar"
    End If
    
Exit Sub
errFDA:
    clsGeneral.OcurrioError "Error al buscar la dirección auxiliar.", Err.Description
End Sub


Private Function TasadeCambio(MOriginal As Integer, MDestino As Integer, Fecha As Date, Optional FechaTC As String = "", Optional TipoTC As Integer = -1) As Currency
Dim RsTC As rdoResultset

    On Error GoTo errTC
    If TipoTC = -1 Then TipoTC = 1
    TasadeCambio = 1
    Cons = "Select * from TasaCambio" _
            & " Where TCaFecha = (Select MAX(TCaFecha) from TasaCambio " _
                                          & " Where TCaFecha < '" & Format(Fecha, "mm/dd/yyyy 23:59") & "'" _
                                          & " And TCaOriginal = " & MOriginal _
                                          & " And TCaDestino = " & MDestino _
                                          & " And TCaTipo = " & TipoTC & ")" _
            & " And TCaOriginal = " & MOriginal _
            & " And TCaDestino = " & MDestino _
            & " And TCaTipo = " & TipoTC
            
    Set RsTC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsTC.EOF Then
        TasadeCambio = CCur(Format(RsTC!TCaComprador, "#.000"))
        FechaTC = Format(RsTC!TCaFecha, "dd/mm/yyyy")
    End If
    RsTC.Close
    Exit Function
    
errTC:
End Function

Private Sub EliminoRenglonesFactura()

    For Each itmx In lvVenta.ListItems
        
        If itmx.SubItems(7) = 151 And Mid(itmx.Key, 1, 1) = "X" Then
            
            Call RestoLabTotales(CCur(lvVenta.SelectedItem.SubItems(5)), CCur(lvVenta.SelectedItem.SubItems(4)))
            lvVenta.ListItems.Remove itmx.Index
            
            'Lo hago recursivo ya que el foreach rompe todo al indice.
            EliminoRenglonesFactura
            Exit Sub
            
        End If
        
    Next
    
End Sub

Private Sub CambioFormaPagoEnvios()
On Error GoTo errCFPE
    
    'Como es caja tengo que fijarme si tiene liquidar.
    If InStr(1, strCodigoEnvio, ",") > 0 Then
        Cons = "SELECT * FROM Envio " & _
            "WHERE EnvCodigo IN (" & strCodigoEnvio & ") And EnvLiquidar = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            MsgBox "ATENCIÓN!!!" & vbCrLf & "Hay envíos que no tienen cuanto debe liquidarse al camión, debe acceder al formulario de envíos y corregir dicho valor si corresponde.", vbInformation, "Posible error"
        End If
        RsAux.Close
        
        Cons = "UPDATE Envio SET EnvFormaPago = " & TipoPagoEnvio.PagaAhora & " WHERE EnvCodigo IN (" & strCodigoEnvio & ")"
        cBase.Execute Cons
    Else
        Cons = "SELECT EnvValorFlete, IsNull(PFlPrecioPpal,0),  IsNull(PFlCostoPpal, 0) FROM Envio LEFT OUTER JOIN PrecioFlete ON EnvTipoFlete = PFlTipoFlete AND PFlPrecioPpal = EnvValorFlete " & _
            "WHERE EnvCodigo = " & strCodigoEnvio & " AND EnvLiquidar = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            Cons = "UPDATE Envio SET EnvFormaPago = " & TipoPagoEnvio.PagaAhora & " WHERE EnvCodigo IN (" & strCodigoEnvio & ")"
        Else
            Dim iCosto As Currency
            If RsAux(1) > 0 And RsAux(0) <> 0 Then
                iCosto = (RsAux(0) / RsAux(1)) * RsAux(2)
            Else
                iCosto = RsAux(0)
            End If
            MsgBox "ATENCIÓN!!!" & vbCrLf & "Se cambia la forma de pago y se le agrega costo a liquidar para el camión.", vbInformation, "ATENCIÓN"
            Cons = "UPDATE Envio SET EnvFormaPago = " & TipoPagoEnvio.PagaAhora & ", EnvLiquidar = " & iCosto & " WHERE EnvCodigo IN (" & strCodigoEnvio & ")"
        End If
        RsAux.Close
        
        cBase.Execute Cons
    End If
    
    
    EliminoRenglonesFactura
    
    Dim rsE As rdoResultset
    Cons = "SELECT EnvCodigo, EnvTipoFlete, IsNull(EnvValorFlete, 0) EnvValorFlete, IsNull(EnvValorPiso, 0) EnvValorPiso FROM Envio WHERE EnvCodigo IN (" & strCodigoEnvio & ") AND EnvFormaPago = " & TipoPagoEnvio.PagaAhora
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsE.EOF
        
        Dim bInsert As Boolean
        If rsE!EnvValorFlete > 0 Then
            bInsert = False
            Dim RsArt As rdoResultset
            Cons = "Select ArtID, ArtTipo, ArtNombre From TipoFlete, Articulo, ArticuloFacturacion, TipoIva " _
                & " Where TFlCodigo = " & rsE!EnvTipoFlete _
                & " And ArtID = TFlArticulo And ArtId = AFaArticulo  And AFaIva = IvaCodigo"
            
            Set RsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsArt.EOF Then
                
                For Each itmx In lvVenta.ListItems
                    If itmx.SubItems(7) = 151 And itmx.Key = "X" & RsArt("ArtID") Then
                        
                        bInsert = True
                        itmx.Text = Val(itmx.Text) + 1
                        itmx.SubItems(5) = Format(CCur(itmx.SubItems(5)) + rsE!EnvValorFlete, "#,##0.00")
                        
                        Exit For
                    End If
                Next
                
                If Not bInsert Then
                    s_InsertArticulo RsArt("ArtID"), RsArt("ArtTipo"), 1, Trim(RsArt!ArtNombre), rsE!EnvValorFlete, rsE!EnvValorFlete, "", False, True
                End If
            End If
            RsArt.Close
        End If
        If rsE!EnvValorPiso > 0 Then
            bInsert = False
            
            Cons = "Select ArtID, ArtTipo, ArtNombre  From Articulo, ArticuloFacturacion, TipoIva" _
                        & " Where ArtId = " & paArticuloPisoAgencia _
                        & " And ArtID = AFaArticulo And AFaArticulo = " & paArticuloPisoAgencia _
                        & " And AFaIVA = IVACodigo"
            
            Set RsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsArt.EOF Then
                
                For Each itmx In lvVenta.ListItems
                    If itmx.SubItems(7) = 151 And itmx.Key = "X" & RsArt("ArtID") Then
                        
                        bInsert = True
                        itmx.Text = Val(itmx.Text) + 1
                        itmx.SubItems(5) = Format(CCur(itmx.SubItems(5)) + rsE!EnvValorPiso, "#,##0.00")
                        
                        Exit For
                    End If
                Next
                
                If Not bInsert Then
                    s_InsertArticulo RsArt("ArtID"), RsArt("ArtTipo"), 1, Trim(RsArt!ArtNombre), rsE!EnvValorPiso, rsE!EnvValorPiso, "", False, True
                End If
            End If
            RsArt.Close
        End If
        rsE.MoveNext
    Loop
    rsE.Close
    
    TotalVenta
    
    'El envío tiene el valor de la cobranza así que lo modifico.
    
    Exit Sub
    
errCFPE:
    clsGeneral.OcurrioError "Error al editar los envíos", Err.Description, "Editar envíos"
End Sub

Private Sub TotalVenta()
    
'    .SubItems(3) = Format(cPrecioUID, "#,#00.00")
'    .SubItems(4) = IVAArticulo(lIDArt)
'    .SubItems(5) = Format(CCur(.SubItems(3)) * cQ, "#,##0.00")
    labIVA.Caption = 0
    labTotal.Caption = 0
    For Each itmx In lvVenta.ListItems
        labIVA.Caption = Format(CCur(labIVA.Caption) + (itmx.SubItems(5) - (itmx.SubItems(5) / CCur(1 + (CCur(itmx.SubItems(4)) / 100)))), "#,##0.00")
        labTotal.Caption = Format(CCur(labTotal.Caption) + itmx.SubItems(5), "#,##0.00")
    Next
    If lblFlete.Caption = "" Then lblFlete.Caption = "0.00"
    lblTotalCflete.Caption = Format(CCur(lblFlete.Caption) + CCur(labTotal.Caption), "#,##0.00")
        
End Sub

Private Sub txtCliente_BorroCliente()
    LimpioDatosCliente
End Sub

Private Sub txtCliente_CambioTipoDocumento()
    lblRutCliente.Visible = (txtCliente.DocumentoCliente <> DC_RUT)
    lblInfoRutCliente.Visible = (txtCliente.DocumentoCliente <> DC_RUT)
    lblIdentificador.ForeColor = vbBlack
    Select Case txtCliente.DocumentoCliente
        Case DC_CI
            lblIdentificador.Caption = "C.I.:"
        Case DC_RUT
            lblIdentificador.Caption = "R.U.T.:"
        Case Else
            If txtCliente.Cliente.TipoDocumento.Nombre = "" Then
                lblIdentificador.Caption = "Otro:"
            Else
                lblIdentificador.Caption = txtCliente.Cliente.TipoDocumento.Abreviacion
            End If
            lblIdentificador.ForeColor = &H40C0&
    End Select
End Sub

Private Sub txtCliente_Focus()
    Status.Panels(1).Text = "Ingrese el documento del cliente (cambie opción con C, E y O."
End Sub

Private Sub txtCliente_PresionoEnter()
    If (txtCliente.Cliente.Codigo > 0) Then
        tArticulo.SetFocus
    End If
End Sub

Private Sub txtCliente_SeleccionoCliente()

'Tengo un nuevo cliente.
    On Error GoTo errBC
    LimpioDatosCliente
    tNombreC.Text = txtCliente.Cliente.Nombre
    lblRutCliente.Caption = clsGeneral.RetornoFormatoRuc(txtCliente.Cliente.RutPersona)
    'TODO: ver campo tNombreC.Tag
    Cons = "Select CliDireccion, CPERuc, CliCategoria From Cliente " _
            & " Left Outer Join CPersona On CPeCliente = CliCodigo" _
            & " Left Outer Join CEmpresa On CEmCliente = CliCodigo" _
        & " Where CliCodigo = " & txtCliente.Cliente.Codigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(RsAux!CliDireccion) Then
        cDireccion.AddItem "Dirección Principal": cDireccion.ItemData(cDireccion.NewIndex) = RsAux!CliDireccion
        cDireccion.Tag = RsAux!CliDireccion
        gDirFactura = RsAux!CliDireccion
    End If
    If Not IsNull(RsAux!CliCategoria) Then labDireccion.Tag = RsAux!CliCategoria
    RsAux.Close
    
    CargoTelefonos txtCliente.Cliente.Codigo, paTipoTelefonoP
    CargoDireccionesAuxiliares txtCliente.Cliente.Codigo
    
    prmIDCliente = txtCliente.Cliente.Codigo
    
    If txtCliente.Enabled Then
        Cons = "Select VTeCliente From VentaTelefonica " _
            & " Where VTeTipo IN (" & TipoDocumento.ContadoDomicilio & ", " & TipoDocumento.VentaOnLineConfirmada & ", " & TipoDocumento.VentaOnLineAConfirmar & ", " & TipoDocumento.VentaRedPagosTelefonicas & ")" _
            & " And VTeCliente = " & idCliente _
            & " And VTeAnulado Is Null And VTeDocumento Is Null"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Close
            MsgBox "IMPORTATE!!!! " & vbCrLf & vbCrLf & "El cliente ya tiene ventas telefónicas pendientes, verifique en visualización de operaciones.", vbInformation, "ATENCIÓN"
        Else
            RsAux.Close
            Me.Show
            Foco tArticulo
        End If
    End If
    txtCliente.BuscoComentariosAlerta txtCliente.Cliente.Codigo, True
    If txtCliente.DarMsgClienteNoVender(txtCliente.Cliente.Codigo) Then
        MsgBox "Atención: NO se puede vender sin autorización. Consultar con gerencia!", vbCritical, "ATENCIÓN"
    End If
    VeoCambiosEnDescuentos
    AvisoVentaLimitada False
    ValidoRUT
'    If sNuevo And txtCliente.Cliente.RutPersona <> "" Then
'        MsgBox "UNIPERSONAL" & vbCrLf & vbCrLf & "Consulte con el cliente si desea facturar con RUT", vbExclamation, "ATENCIÓN"
'    End If
    Exit Sub
errBC:
    clsGeneral.OcurrioError "Error al buscar el cliente por parámetro.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function AvisoVentaLimitada(ByVal doyMsg As Boolean) As Boolean

    For Each itmx In lvVenta.ListItems
        'itmx.Bold = False
        itmx.ForeColor = vbBlack
        If InStr(1, itmx.Key, "C") = 0 And Not ArticuloEsServicio(Mid(itmx.Key, 2, Len(itmx.Key))) Then  'Val(itmx.SubItems(7)) <> paTipoArticuloServicio Then
            If Val(itmx.SubItems(9)) = 0 And InStr(1, paCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0 Then
                'itmx.Bold = True
                itmx.ForeColor = &HFF&
                AvisoVentaLimitada = True
                If doyMsg Then MsgBox "Atención!!! " & vbCrLf & vbCrLf & "No está autorizada la venta del artículo " & itmx.SubItems(1) & " a distribuidores." & vbCrLf & vbCrLf & "Debe consultar para vender.", vbExclamation, "POSIBLE ERROR"
            ElseIf Val(itmx.SubItems(9)) > 1 And Val(itmx.SubItems(9)) < itmx.Text Then
                'itmx.Bold = True
                itmx.ForeColor = &HFF&
                AvisoVentaLimitada = True
                If doyMsg Then MsgBox "Atención!!! " & vbCrLf & vbCrLf & "La cantidad máxima autorizada de venta para el artículo " & itmx.SubItems(1) & " es de  " & Val(itmx.SubItems(9)) & vbCrLf & vbCrLf & "Debe consultar para exceder dicha cantidad.", vbExclamation, "POSIBLE ERROR"
            End If
        End If
    Next
    
    lvVenta.Refresh
    
End Function

Private Function ValidoRUT() As Boolean
On Error GoTo errVR
    
    ValidoRUT = True
    Dim oValida As New clsValidaRUT
    If txtCliente.Cliente.Tipo = TC_Empresa And txtCliente.Cliente.Documento <> "" Then
        If Not oValida.ValidarRUT(txtCliente.Cliente.Documento) Then
            MsgBox "RUT INCORRECTO!!!, por favor valide con el cliente el número de RUT ya que no cumple con la validación.", vbExclamation, "RUT INCORRECTO"
            ValidoRUT = False
        End If
    ElseIf txtCliente.Cliente.Tipo = TC_Persona And txtCliente.Cliente.RutPersona <> "" Then
        If Not oValida.ValidarRUT(txtCliente.Cliente.RutPersona) Then
            MsgBox "RUT INCORRECTO!!!, por favor valide con el cliente el número de RUT ya que no cumple con la validación.", vbExclamation, "RUT INCORRECTO"
            ValidoRUT = False
        End If
    End If
    Set oValida = Nothing
    
errVR:
End Function

Private Sub CorrijoStockCAGADAEliminarEnvio()
    
    Dim rsC As rdoResultset
    Cons = "SELECT VTeCodigo, RVTArticulo, RVTCantidad  FROM VentaTelefonica INNER JOIN RenglonVtaTelefonica on VTeCodigo = RVTVentaTelefonica " & _
        "WHERE VTeAnulado between '20160701' and '20161121' and VTeTipo = 44 AND VTeCodigo > 590539 and VTeCodigo IN (SELECT MSEDocumento FROM MovimientoStockEstado GROUP BY MSEDocumento having SUM(MSECantidad)<>0) " & _
        " AND RVTArticulo not in (select artid from Articulo where ArtTipo = 151) Order by vtecodigo"
    cBase.QueryTimeout = 90
    Set rsC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsC.EOF
        MarcoStockVenta 400, rsC("RVTArticulo"), rsC("RVTCantidad"), rsC("RVTCantidad"), 0, 7, rsC("VTeCodigo"), 5
        Debug.Print rsC("VTeCodigo")
        rsC.MoveNext
    Loop
    rsC.Close
End Sub

Private Sub CargarFletesEnvio()
On Error GoTo errCFE
Dim RsF As rdoResultset
    
    lblFlete.Caption = 0
    lblTotalCflete.Caption = 0
    
    Cons = "SELECT SUM(IsNull(EnvValorFlete, 0)) + SUM(IsNull(EnvValorPiso, 0)) FROM Envio WHERE EnvCodigo IN (" & strCodigoEnvio & ")"
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    lblFlete.Caption = Format(RsF(0), "#,##0.00")
    RsF.Close
    
    lblTotalCflete.Caption = Format(CCur(lblFlete.Caption) + CCur(labTotal.Caption), "#,##0.00")
errCFE:
End Sub

Private Sub CargarFletesEnvioPorVenta()
On Error GoTo errCFE
Dim RsF As rdoResultset
    
    lblFlete.Caption = 0
    lblTotalCflete.Caption = 0
    
    Cons = "SELECT SUM(IsNull(EnvValorFlete, 0)) + SUM(IsNull(EnvValorPiso, 0)) FROM Envio WHERE EnvTipo = 3 AND EnvDocumento " & tCodigo.Text
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    lblFlete.Caption = Format(RsF(0), "#,##0.00")
    RsF.Close
    
    lblTotalCflete.Caption = Format(CCur(lblFlete.Caption) + CCur(labTotal.Caption), "#,##0.00")
errCFE:
End Sub


