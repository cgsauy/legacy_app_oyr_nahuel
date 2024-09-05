VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmNotaCtdo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Devolución"
   ClientHeight    =   4815
   ClientLeft      =   2445
   ClientTop       =   2445
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNotaCtdo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8190
   Begin VB.TextBox tEnvio 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3540
      MaxLength       =   7
      TabIndex        =   10
      Text            =   "8888888"
      Top             =   3900
      Width           =   735
   End
   Begin VB.TextBox tFicha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1020
      MaxLength       =   7
      TabIndex        =   8
      Text            =   "8888888"
      ToolTipText     =   "F1 muestra las devoluciones ingresadas."
      Top             =   3900
      Width           =   735
   End
   Begin AACombo99.AACombo cLocal 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   780
      Width           =   1875
      _ExtentX        =   3307
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imprimir"
            Object.ToolTipText     =   "Emitir Nota"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "envio"
            Object.ToolTipText     =   "Formulario de Envíos"
            ImageKey        =   "Envio"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "remito"
            ImageKey        =   "Remito"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   6350
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1020
      MaxLength       =   70
      TabIndex        =   12
      Top             =   4200
      Width           =   7155
   End
   Begin MSComctlLib.ListView lvArticulo 
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   2220
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cant"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Artículo"
         Object.Width           =   4535
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Compro"
         Object.Width           =   1303
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "A Retirar"
         Object.Width           =   1481
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Envío"
         Object.Width           =   1005
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Remito"
         Object.Width           =   1183
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Nota"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Ficha"
         Object.Width           =   1007
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Unitario"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "IVA"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   4560
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.TextBox tNumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   317
      Left            =   3780
      MaxLength       =   6
      TabIndex        =   4
      Top             =   780
      Width           =   855
   End
   Begin VB.TextBox tSerie 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   317
      Left            =   3540
      MaxLength       =   1
      TabIndex        =   3
      Top             =   780
      Width           =   242
   End
   Begin vsViewLib.vsPrinter vsFicha 
      Height          =   1515
      Left            =   660
      TabIndex        =   31
      Top             =   4680
      Visible         =   0   'False
      Width           =   7035
      _Version        =   196608
      _ExtentX        =   12409
      _ExtentY        =   2672
      _StockProps     =   229
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      PageBorder      =   0
      BackColor       =   -2147483633
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   5820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotaCtdo.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotaCtdo.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotaCtdo.frx":086E
            Key             =   "Remito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotaCtdo.frx":0B88
            Key             =   "Envio"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotaCtdo.frx":0EA2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCFE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "eFactura"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4800
      TabIndex        =   36
      Top             =   780
      Width           =   735
   End
   Begin VB.Label lPCn 
      BackStyle       =   0  'Transparent
      Caption         =   "Salida de Caja:"
      Height          =   375
      Left            =   5760
      TabIndex        =   35
      Top             =   420
      Width           =   2415
   End
   Begin VB.Label lPNC 
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora:"
      Height          =   255
      Left            =   3240
      TabIndex        =   34
      Top             =   420
      Width           =   2415
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Envío:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2820
      TabIndex        =   9
      Top             =   3900
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Ficha:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   3900
      Width           =   795
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0E0FF&
      Caption         =   "  Devolución"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   1920
      Width           =   8235
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0E0FF&
      Caption         =   "  Contado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   0
      TabIndex        =   32
      Top             =   420
      Width           =   8235
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   255
      Left            =   60
      TabIndex        =   30
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label labDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   900
      TabIndex        =   29
      Top             =   1380
      Width           =   6975
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S&ucursal:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   795
   End
   Begin VB.Label Label10 
      Caption         =   "&Lista"
      Height          =   255
      Left            =   5700
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentario:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label labNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   900
      TabIndex        =   28
      Top             =   1140
      Width           =   4695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   60
      TabIndex        =   27
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label lblRUT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "21.025996.0012"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6540
      TabIndex        =   26
      Top             =   1140
      Width           =   1395
   End
   Begin VB.Label labDato1 
      BackStyle       =   0  'Transparent
      Caption         =   "R.U.C.:"
      Height          =   255
      Left            =   5760
      TabIndex        =   25
      Top             =   1140
      Width           =   675
   End
   Begin VB.Label labIVA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1,252,200.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6180
      TabIndex        =   23
      Top             =   4020
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "I.V.A.:"
      Height          =   255
      Left            =   5460
      TabIndex        =   22
      Top             =   4020
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label labImporteNota 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1,252,252.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6840
      TabIndex        =   21
      Top             =   3780
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Importe:"
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   3780
      Width           =   735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Importe Descontado:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4740
      TabIndex        =   19
      Top             =   1620
      Width           =   1695
   End
   Begin VB.Label labImporteDescontado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6660
      TabIndex        =   18
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label labFechaDocumento 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "10-Dic-1998"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6840
      TabIndex        =   16
      Top             =   780
      Width           =   1095
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   15
      Top             =   780
      Width           =   735
   End
   Begin VB.Label labImporteTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   900
      TabIndex        =   14
      Top             =   1620
      Width           =   1695
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Importe:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   1620
      Width           =   735
   End
   Begin VB.Label lblNroDoc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Número:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2820
      TabIndex        =   2
      Top             =   780
      Width           =   735
   End
   Begin VB.Menu MnuMenu 
      Caption         =   "&Menú"
      Begin VB.Menu MnuEmitir 
         Caption         =   "&Imprimir Nota"
         Enabled         =   0   'False
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEnvio 
         Caption         =   "&Envío"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuRemito 
         Caption         =   "&Remito"
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVolver 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuPrinter 
      Caption         =   "Impresora"
      Begin VB.Menu MnuPrintConfig 
         Caption         =   "Configurar"
      End
      Begin VB.Menu MnuPrintLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrintOpt 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmNotaCtdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tag Utilizados.-----*******--------------------------------------
'tnumero                            Guardo el código del documento.
'labFechaDocumento                  Guardo ID de cliente.
'labImporteTotal                    Guardo Fecha de Modificación del documento.
'labImporteDescontado               Guardo el código de moneda utilizada.
'------------------------*******-------------------------------------------
'Modificaciones
'31-10-2003     Agregue anulación de Instalaciones.
'28/10/2005     Al dar menos en la lista y llegar a cero vacío el id de dev, x lo tanto si incrementa
'               vuelvo a buscar el mismo.
'........................................................................................................
Option Explicit
Private Const FormatoFH = "mm/dd/yyyy hh:mm:ss"
Dim oCnfgPrintSalidaCaja As New clsImpresoraTicketsCnfg
Private oFichas As New Collection

Private oClienteCtdo As clsClienteCFE
Private EmpresaEmisora As clsClienteCFE

'Private Type tDevolucion
'    idArt As Long
'    idDev As Long
'    Cant As Integer
'End Type

Private itmx As ListItem
Dim CodDocumentoEnvio As Long        'Esta la utilizo sólo cuando la factura paga artículos de fletes.
Private Rs As rdoResultset
'Private arrDevolucion() As tDevolucion

Private Sub cLocal_GotFocus()
    With cLocal
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Seleccione el Local de emisión del documento."
End Sub
Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cLocal.ListIndex > -1 Then
            tSerie.SetFocus
        Else
            MsgBox "El ingreso del local es obligatorio.", vbExclamation, "ATENCIÓN"
            cLocal.SetFocus
        End If
    End If
End Sub
Private Sub cLocal_LostFocus()
    cLocal.SelStart = 0
    Status.SimpleText = ""
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    DoEvents
End Sub
Private Sub Form_Load()

    On Error GoTo ErrLoad
    
    If Not ValidarVersionEFactura Then
        MsgBox "La versión del componente CGSAEFactura está desactualizado, debe distribuir software." _
                    & vbCrLf & vbCrLf & "Se cancelará la ejecución.", vbCritical, "EFactura"
        End
    End If
   
    Set EmpresaEmisora = New clsClienteCFE
    EmpresaEmisora.CargoInformacionCliente cBase, 1, False
        
    oCnfgPrintSalidaCaja.CargarConfiguracion "MovimientosDeCaja", "TickeadoraMovimientosDeCaja"
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = 5475
    
    s_LoadMenuOpcionPrint
    lPNC.Caption = "Imp. Nota:" & paINContadoN
    If Not paPrintEsXDefNC Then lPNC.ForeColor = &HC0&
    
    lPCn.Caption = "Imp. Salida Caja: " & IIf(oCnfgPrintSalidaCaja.Opcion = 0, paIConformeN, oCnfgPrintSalidaCaja.ImpresoraTickets)
    If Not paPrintEsXDefCn Then lPCn.ForeColor = &HC0&
    
    'ReDim arrDevolucion(0)
    CodDocumentoEnvio = 0
    CargoLocales
    BuscoCodigoEnCombo cLocal, paCodigoDeSucursal
    
    LimpioCampos
    DeshabilitoCampos
    crAbroEngine
    
    On Error Resume Next
    With vsFicha
        .Device = paIConformeN
        .PaperBin = paIConformeB
        .paperSize = paIConformeP
    End With
    Exit Sub
ErrLoad:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio el sgte. error: " & Trim(Err.Description)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.SimpleText = vbNullString
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set clsGeneral = Nothing
    Set miconexion = Nothing
    CierroConexion
    crCierroEngine
    End
End Sub
Private Sub Label2_Click()
    Foco tFicha
End Sub

Private Sub Label6_Click()
    Foco tEnvio
End Sub

Private Sub Label8_Click()
    Foco tComentario
End Sub

Private Sub lPCn_DblClick()
    ImprimoRetirosPorDevolucion 124595, 7313938, 7785437
End Sub

Private Sub lPNC_DblClick()
    ImprimoSalidaCajaTicket 7764105, 6, oCnfgPrintSalidaCaja.ImpresoraTickets
End Sub

Private Sub MnuEmitir_Click()
    AccionImprimir
End Sub

Private Sub MnuEnvio_Click()
    FormEnvio
End Sub

Private Sub MnuPrintConfig_Click()
On Error Resume Next
    
    prj_LoadConfigPrint True
    
    Dim iQ As Integer
    For iQ = MnuPrintOpt.LBound To MnuPrintOpt.UBound
        MnuPrintOpt(iQ).Checked = False
        MnuPrintOpt(iQ).Checked = (MnuPrintOpt(iQ).Caption = paOptPrintSel)
    Next
    
    lPNC.Caption = "Imp. Nota:" & paINContadoN
    If Not paPrintEsXDefNC Then lPNC.ForeColor = &HC0& Else lPNC.ForeColor = vbBlack
    
    lPCn.Caption = "Imp. Salida Caja: " & IIf(oCnfgPrintSalidaCaja.Opcion = 0, paIConformeN, oCnfgPrintSalidaCaja.ImpresoraTickets)
    If Not paPrintEsXDefCn Then lPCn.ForeColor = &HC0& Else lPNC.ForeColor = vbBlack
    
End Sub

Private Sub MnuPrintOpt_Click(Index As Integer)
On Error GoTo errLCP
Dim objPrint As New clsCnfgPrintDocument
Dim iQ As Integer
    
    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If .ChangeConfigPorOpcion(MnuPrintOpt(Index).Caption) Then
            For iQ = MnuPrintOpt.LBound To MnuPrintOpt.UBound
                MnuPrintOpt(iQ).Checked = False
            Next
            MnuPrintOpt(Index).Checked = True
        End If
    End With
    Set objPrint = Nothing
    
    On Error Resume Next
    prj_LoadConfigPrint False
    
    lPNC.Caption = "Imp. Nota:" & paINContadoN
    If Not paPrintEsXDefNC Then lPNC.ForeColor = &HC0& Else lPNC.ForeColor = vbBlack
    
    lPCn.Caption = "Imp. Salida Caja: " & IIf(oCnfgPrintSalidaCaja.Opcion = 0, paIConformeN, oCnfgPrintSalidaCaja.ImpresoraTickets)
    If Not paPrintEsXDefCn Then lPCn.ForeColor = &HC0& Else lPNC.ForeColor = vbBlack
    
    Exit Sub
errLCP:
    MsgBox "Error al setear los datos de configuración: " & Err.Description, vbExclamation, "ATENCIÓN"
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub
Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0
    tComentario.SelLength = Len(tComentario.Text)
    Status.SimpleText = " Ingrese un comentario para la nota."
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionImprimir
End Sub
Private Sub tComentario_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub tEnvio_GotFocus()
    With tFicha
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese el código de envío asociado a la ficha de alta de stock."
End Sub

Private Sub tEnvio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tEnvio.Text) = "" Then
            Foco tComentario
        Else
            If Not IsNumeric(tEnvio.Text) Then
                MsgBox "Debe ingresar un valor numérico.", vbExclamation, "ATENCIÓN"
            Else
                Foco tComentario
            End If
        End If
    End If
End Sub

Private Sub tFicha_GotFocus()
    With tFicha
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese un código de ficha de alta de stock."
End Sub

Private Sub tFicha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        'Presento las fichas ingresadas.
        Dim msgShow As String
        Dim cteShow As String
        cteShow = "Devolución: {idDev} cantidad: {CantArts}"
        Dim oFicha As clsFichasDevolucion
        Dim oItem As clsArtEnFicha
        For Each oFicha In oFichas
            For Each oItem In oFicha.Fichas
                msgShow = msgShow & IIf(msgShow <> "", vbCrLf, "") & "Devolución: " & oItem.IDFicha & " cantidad: " & oItem.Cantidad
            Next
        Next
        If msgShow <> "" Then MsgBox msgShow, vbInformation, "Devoluciones ingresadas"
    End If
End Sub

Private Sub tFicha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tFicha.Text) = "" Then
            If SaltoAEnvio Then Foco tEnvio Else Foco tComentario
        Else
            If IsNumeric(tFicha.Text) Then
                BuscoDevolucion tFicha.Text
            Else
                MsgBox "Ingrese un número.", vbExclamation, "ATENCIÓN"
            End If
        End If
    End If
End Sub

Private Sub tNumero_Change()
    LimpioCampos
    DeshabilitoCampos
End Sub

Private Sub tNumero_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub tNumero_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.SimpleText = " Ingrese el número del documento."
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    loc_BuscarCambiosEnFicha
    
    Select Case Button.Key
        Case "imprimir": AccionImprimir
        Case "salir": Unload Me
        Case "envio": FormEnvio
    End Select

End Sub

Private Sub tSerie_Change()
On Error Resume Next

    If IsNumeric(tSerie.Text) Then
        tSerie.Text = ""
    Else
        If Trim(tSerie.Text) <> "" Then
            If (Asc(UCase(tSerie.Text)) > 64 And Asc(UCase(tSerie.Text)) < 91) Or Asc(UCase(tSerie.Text)) = 209 Then
                tSerie.Text = UCase(tSerie.Text)
                tNumero.SetFocus
            Else
                tSerie.Text = ""
            End If
        End If
    End If

End Sub

Private Sub tSerie_GotFocus()

    tSerie.SelStart = 0
    tSerie.SelLength = Len(tSerie.Text)
    Status.SimpleText = " Ingrese la serie del documento."

End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Trim(tSerie.Text) <> "" Then tNumero.SetFocus

End Sub

Private Sub tSerie_LostFocus()

    If Trim(tSerie.Text) <> "" Then
        If (Asc(UCase(tSerie.Text)) > 64 And Asc(UCase(tSerie.Text)) < 91) Or Asc(UCase(tSerie.Text)) = 209 Then
                tSerie.Text = UCase(tSerie.Text)
        End If
    End If
    Status.SimpleText = vbNullString

End Sub

Private Sub tNumero_GotFocus()

    tNumero.SelStart = 0
    tNumero.SelLength = Len(tNumero.Text)
    Status.SimpleText = " Ingrese el número del documento."

End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        MnuEmitir.Enabled = False
        Toolbar1.Buttons("imprimir").Enabled = False
        If IsNumeric(tNumero.Text) And Trim(tSerie.Text) <> vbNullString And cLocal.ListIndex > -1 Then
            If CLng(tNumero.Text) < 1 Then
                MsgBox "Ingrese un número mayor que cero.", vbExclamation, "ATENCIÓN"
            Else
                BuscoFactura
                If lvArticulo.ListItems.Count > 0 And lvArticulo.Enabled Then
                    lvArticulo.ListItems(1).Selected = True: lvArticulo.SetFocus
                End If
            End If
        Else
            If cLocal.ListIndex = -1 Then
                MsgBox "Seleccione el local de emisión del documento.", vbExclamation, "ATENCIÓN": Foco cLocal
            ElseIf Trim(tSerie.Text) = "" Then
                MsgBox "Ingrese un nro. de serie.", vbExclamation, "ATENCIÓN": Foco tSerie
            Else
                If Trim(tNumero.Text) <> "" Then
                    MsgBox "El formato no es numérico.", vbExclamation, "ATENCIÓN": Foco tNumero
                End If
            End If
        End If
    End If

End Sub

Private Function BuscoIDFactura() As Long

    Cons = "Select DocCodigo, DocFecha Fecha, DocSerie + '-' + CAST(DocNumero as varchar(10)) Número , DocTotal Total, " _
        & " CASE WHEN CliTipo = 1 THEN RTRIM(CPeNombre1) + ' ' + RTRIM(CPeApellido1) ELSE CEmNombre END Cliente  " _
        & " From Documento INNER JOIN Cliente ON DocCliente = CliCodigo " _
        & " LEFT OUTER JOIN CPersona ON CPeCliente = CliCodigo " _
        & " LEFT OUTER JOIN CEmpresa ON CEmCliente = CliCodigo " _
        & " Where DocTipo = " & TipoDocumento.Contado _
        & " And DocSerie = '" & tSerie.Text & "' And DocNumero = " & tNumero.Text _
        & " And DocSucursal = " & cLocal.ItemData(cLocal.ListIndex) & " AND DocAnulado = 0"
    
    Dim RsF As rdoResultset
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsF.EOF Then
        RsF.Close
        Exit Function
    Else
        Dim idAux As Long
        idAux = RsF("DocCodigo")
        RsF.MoveNext
        If RsF.EOF Then
            BuscoIDFactura = idAux
            RsF.Close
            Exit Function
        Else
            RsF.Close
            'Voy a lista de ayuda.
        End If
    End If
    
    Dim oAyuda As New clsListadeAyuda
    oAyuda.CerrarSiEsUnico = True
    If oAyuda.ActivarAyuda(cBase, Cons, 7000, 1, "Buscar contados") > 0 Then
        BuscoIDFactura = oAyuda.RetornoDatoSeleccionado(0)
    End If
    Set oAyuda = Nothing
    
End Function

Private Sub BuscoFactura()
'Para buscar un documento se considera la propiedad iDocumento, la cual indica el tipo de documento.
On Error GoTo ErrBF


    Screen.MousePointer = vbHourglass
    Dim idDoc As Long
    idDoc = BuscoIDFactura()
    If idDoc = 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "No existe un documento con esas características.", vbExclamation, "ATENCIÓN"
        tNumero.SetFocus
        Exit Sub
    End If
            
    Cons = "SELECT DocAnulado, DocCliente, DocFecha,  DocFModificacion, DocMoneda, DocTotal, DocCodigo, Renglon.*, MonSigno, ArtNombre, ArtTipo, IsNull(IVAPorcentaje, 0) IVAPorcentaje" _
        & " FROM Documento INNER JOIN Renglon ON DocCodigo = RenDocumento " _
        & " INNER JOIN Articulo ON RenArticulo = ArtID INNER JOIN Moneda ON DocMoneda = MonCodigo" _
        & " LEFT OUTER JOIN ArticuloFacturacion ON AFaArticulo = ArtID LEFT OUTER JOIN TipoIVA ON IvaCodigo = AFaIva " _
        & " Where DocCodigo = " & idDoc
       
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    LimpioCampos
    DeshabilitoCampos
    
    If RsAux.EOF Then
        Screen.MousePointer = vbDefault
        RsAux.Close
        MsgBox "No existe un documento con esas características.", vbExclamation, "ATENCIÓN"
        tNumero.SetFocus
    Else
        If RsAux!DocAnulado Then
            Screen.MousePointer = vbDefault
            RsAux.Close
            MsgBox "El documento seleccionado fue anulado.", vbExclamation, "ATENCIÓN"
            tNumero.SetFocus
        Else
            'Fecha de emisión.
            tNumero.Tag = RsAux!DocCodigo
            labFechaDocumento.Caption = Format(RsAux!DocFecha, "d-Mmm-yyyy")
            labFechaDocumento.Tag = RsAux!DocCliente
            CargoCliente Val(tNumero.Tag), Val(labFechaDocumento.Tag)
            'Importe total del documento.
            labImporteTotal.Caption = Trim(RsAux!MonSigno) & " " & Format(RsAux!DocTotal, "#,##0.00")
            labImporteTotal.Tag = RsAux!DocFModificacion
            labImporteDescontado.Caption = Trim(RsAux!MonSigno)
            labImporteDescontado.Tag = RsAux!DocMoneda
            CargoArticulos
            RsAux.Close
            tFicha.Tag = EsServicio
            'Busco si tiene nota de devolución.------------
            BuscoOtrasNotas
            RecalculoTotales
            If tFicha.Enabled Then BuscoDevolucionesDocumento
            If lvArticulo.Enabled Then lvArticulo.SetFocus Else If tComentario.Enabled Then tComentario.SetFocus
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrBF:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al buscar el documento."
        
End Sub
Private Sub CargoArticulos()
Dim RsLocal As rdoResultset
    
    MnuEnvio.Enabled = False: Toolbar1.Buttons("envio").Enabled = False
    MnuRemito.Enabled = False: Toolbar1.Buttons("remito").Enabled = False
    CodDocumentoEnvio = 0
    
    'RsAux resultset con todos los artículos que tiene el documento.
    Do While Not RsAux.EOF
    
        'Levanto los datos del artículo.
'        Cons = "Select ArtNombre, IvaPorcentaje, ArtTipo" _
            & " From Articulo, ArticuloFacturacion, TipoIva" _
            & " Where ArtID = " & RsAux!RenArticulo & " And ArtID = AFaArticulo And AFaIVA = IVaCodigo"
            
        'Veo si este artículo es de flete.
        Cons = "Select TFlArticulo From TipoFlete Where TFlArticulo = " & RsAux!RenArticulo
        Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            
        If RsLocal.EOF And RsAux!RenArticulo <> paArticuloPisoAgencia Then
            RsLocal.Close
            
            'Veo si es del tipo servicio.
            If RsAux!ArtTipo = paTipoArticuloServicio Then
                Set itmx = lvArticulo.ListItems.Add(, "S" & RsAux!RenArticulo, "")
            Else
                Set itmx = lvArticulo.ListItems.Add(, "A" & RsAux!RenArticulo, "")
            End If
            itmx.SubItems(7) = 0
            itmx.Tag = RsAux("RenIVA")       'guardo el iva x artículo.
            itmx.SubItems(1) = Trim(RsAux("ArtNombre"))
            itmx.SubItems(2) = RsAux!RenCantidad        'Cantidad total en la factura.
            itmx.SubItems(3) = RsAux!RenARetirar
            itmx.SubItems(8) = Format(RsAux!RenPrecio, "#,##0.00")
            itmx.SubItems(9) = Format(RsAux!IVAPorcentaje, "#,##0.00")
            
            'Cons = "Select SUM(RReAEntregar) From Remito, RenglonRemito Where RemDocumento = " & RsAux!DocCodigo _
                & " And RReArticulo = " & RsAux!RenArticulo _
                & " And RemCodigo = RReRemito"
            'remito
            
            Cons = "SELECT SUM(RenARetirar)" & _
                " FROM Renglon INNER JOIN Documento ON RenDocumento = DocCodigo And DocAnulado = 0 AND DocTipo = 6" & _
                " INNER JOIN RemitoDocumento ON RenDocumento = RDoRemito" & _
                " WHERE RenArticulo = " & RsAux!RenArticulo & _
                " AND RDoDocumento = " & RsAux("DocCodigo")
            Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            
            If Not IsNull(RsLocal(0)) Then
                itmx.SubItems(5) = RsLocal(0)
                MnuRemito.Enabled = True: Toolbar1.Buttons("remito").Enabled = True
            Else
                itmx.SubItems(5) = 0
            End If
            RsLocal.Close
            
            Cons = "Select SUM(RReAEntregar) From Remito, RenglonRemito Where RemDocumento = " & RsAux!DocCodigo _
                & " And RReArticulo = " & RsAux!RenArticulo _
                & " And RemCodigo = RReRemito"
            Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            
            If Not IsNull(RsLocal(0)) Then
                itmx.SubItems(5) = Val(itmx.SubItems(5)) + RsLocal(0)
                MnuRemito.Enabled = True: Toolbar1.Buttons("remito").Enabled = True
            Else
                itmx.SubItems(5) = 0
            End If
            RsLocal.Close
            
            
            'Sumo la cantidad de artículos que están para envío.
            Cons = "Select SUM(REvAEntregar)" & _
                " FROM Envio INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio" & _
                " WHERE (EnvDocumento = " & RsAux!DocCodigo & _
                " OR EnvDocumento IN (SELECT RDoRemito FROM RemitoDocumento INNER JOIN Documento ON RDoRemito = DocCodigo AND DocTipo = 6 " & _
                                    " WHERE RDoDocumento = " & RsAux("DocCodigo") & "))" & _
                " AND REvArticulo = " & RsAux!RenArticulo
            Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not IsNull(RsLocal(0)) Then
                itmx.SubItems(4) = RsLocal(0)
                MnuEnvio.Enabled = True: Toolbar1.Buttons("envio").Enabled = True
                CodDocumentoEnvio = 0       'Tiene envío si va a este va con el documento.
            Else
                itmx.SubItems(4) = 0
            End If
            RsLocal.Close

            Cons = "Select Sum(RenCantidad) From Nota, Documento, Renglon " _
                & " Where NotFactura = " & tNumero.Tag _
                & " And DocTipo IN (" & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ") And RenArticulo = " & RsAux!RenArticulo _
                & " And NotNota = DocCodigo And DocCodigo = RenDocumento And DocAnulado = 0"
            Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not IsNull(RsLocal(0)) Then itmx.SubItems(6) = RsLocal(0) Else itmx.SubItems(6) = 0
            RsLocal.Close
            
            itmx.Text = CInt(itmx.SubItems(2)) - (CInt(itmx.SubItems(4)) + CInt(itmx.SubItems(5)) + CInt(itmx.SubItems(6)))
            
        Else
            
            'El artículo es de Flete.
            RsLocal.Close
            
            Set itmx = lvArticulo.ListItems.Add(, "F" & RsAux!RenArticulo, "")
            itmx.Tag = RsAux("RenIVA")      'guardo el iva x artículo.Rs!IVaPorcentaje
            itmx.SubItems(1) = Trim(RsAux("ArtNombre"))
            itmx.SubItems(2) = RsAux!RenCantidad
            itmx.SubItems(3) = 0    'A retirar
            itmx.SubItems(5) = 0    'Remito
            itmx.SubItems(7) = 0
            itmx.SubItems(8) = Format(RsAux!RenPrecio, "#,##0.00")
            itmx.SubItems(4) = 0
            itmx.SubItems(9) = Format(RsAux!IVAPorcentaje, "#,##0.00")
            
            'Veo si el documento es el que paga el envío.
            'Si el envío no fue entregado entonces tiene que cambiar la forma de pago en el envío y la nota la
            'hace el envío.
            Cons = "Select EnvDocumento From Envio Where EnvDocumentoFactura = " & RsAux!DocCodigo _
                & " And EnvEstado <> " & EstadoEnvio.Entregado
            
            Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            
            If Not RsLocal.EOF Then
                CodDocumentoEnvio = RsLocal(0)
                If Not MnuEnvio.Enabled Then MnuEnvio.Enabled = True: Toolbar1.Buttons("envio").Enabled = True
            End If
            RsLocal.Close
            
            'Veo si hay notas para el artículo.
            Cons = "Select Sum(RenCantidad) From Nota, Documento, Renglon " _
                & " Where NotFactura = " & tNumero.Tag _
                & " And DocTipo = " & TipoDocumento.NotaDevolucion & " And RenArticulo = " & RsAux!RenArticulo _
                & " And NotNota = DocCodigo And DocCodigo = RenDocumento And DocAnulado = 0"
        
            Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not IsNull(RsLocal(0)) Then itmx.SubItems(6) = RsLocal(0) Else itmx.SubItems(6) = 0
            RsLocal.Close
            
            itmx.Text = 0
        End If
        RsAux.MoveNext
    Loop

End Sub

Private Sub BuscoOtrasNotas()

    Cons = "Select Sum(DocTotal) From Nota, Documento " _
        & " Where NotFactura = " & tNumero.Tag _
        & " And DocTipo IN (" & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")" _
        & " And DocAnulado = 0 And NotNota = DocCodigo"
        
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not Rs.EOF Then
        If Not IsNull(Rs(0)) Then
            labImporteDescontado.Caption = labImporteDescontado.Caption & " " & Format(Rs(0), "#,##0.00")
        Else
            labImporteDescontado.Caption = labImporteDescontado.Caption & " 0.00"
        End If
    End If
    Rs.Close

End Sub
Private Sub lvArticulo_AfterLabelEdit(Cancel As Integer, NewString As String)

    If Not IsNumeric(NewString) Then
        MsgBox "No se ingreso un numéro.", vbExclamation, "ATENCIÓN"
        Cancel = 1
    Else
        If CLng(NewString) > CLng(lvArticulo.SelectedItem.SubItems(2)) Or _
            CLng(NewString) < 0 Then
            MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            Cancel = 1
        Else
            If lvArticulo.SelectedItem.SubItems(7) = 0 And Val(lvArticulo.SelectedItem.Text) = 0 Then
                lvArticulo.Tag = "1"
            End If
        End If
    End If

End Sub

Private Sub lvArticulo_GotFocus()
    Status.SimpleText = " Seleccione un artículo e indique si lo devuelve ('S', 'N'), modifique la cantidad ('+', '-')."
End Sub

Private Sub lvArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    If lvArticulo.ListItems.Count > 0 Then
        Select Case KeyCode
            Case vbKeyReturn
                If SaltoAFicha Then
                    tFicha.SetFocus
                Else
                    If SaltoAEnvio Then
                        tEnvio.SetFocus
                    Else
                        tComentario.SetFocus
                    End If
                End If
                
            Case vbKeyAdd
                If CLng(lvArticulo.SelectedItem.Text) < CInt(lvArticulo.SelectedItem.SubItems(2)) - (CInt(lvArticulo.SelectedItem.SubItems(4)) + CInt(lvArticulo.SelectedItem.SubItems(5)) + CInt(lvArticulo.SelectedItem.SubItems(6))) Then
                    '28/10 si llego a cero borre el id de dev. x lo tanto lo busco para el documento nuevamente
                    lvArticulo.SelectedItem.Text = CLng(lvArticulo.SelectedItem.Text) + 1
                    If Val(lvArticulo.SelectedItem.Text) = 1 Then
                        'el valor anterior era cero
                        lvArticulo.Tag = "1"
                    End If
                    RecalculoTotales
                End If
                
            
            Case vbKeySubtract
                If CLng(lvArticulo.SelectedItem.Text) > 0 Then
                    If CantidadArticulosEnFicha(Mid(lvArticulo.SelectedItem.Key, 2)) > 0 Then
                        If lvArticulo.SelectedItem.Text <= CantidadArticulosEnFicha(Mid(lvArticulo.SelectedItem.Key, 2)) Then
                            Dim result As Integer
                            result = vbCancel
                            Do While result = vbCancel
                                result = MsgBox("Si resta un artículo se eliminarán las fichas ingresadas para el artículo." & vbCrLf & vbCrLf & "¿Desea restar el artículo?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Eliminar fichas del artículo")
                            Loop
                            If result = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                    lvArticulo.SelectedItem.Text = CLng(lvArticulo.SelectedItem.Text) - 1
                    RecalculoTotales
                    If CLng(lvArticulo.SelectedItem.Text) = 0 Then
                        lvArticulo.SelectedItem.SubItems(7) = 0
                        EliminarFichas Mid(lvArticulo.SelectedItem.Key, 2)
                    End If
                End If
                
            Case vbKeyS
                lvArticulo.SelectedItem.Text = CInt(lvArticulo.SelectedItem.SubItems(2)) - (CInt(lvArticulo.SelectedItem.SubItems(4)) + CInt(lvArticulo.SelectedItem.SubItems(5)) + CInt(lvArticulo.SelectedItem.SubItems(6)))
                RecalculoTotales
            
            Case vbKeyN
                lvArticulo.SelectedItem.Text = 0
                lvArticulo.SelectedItem.SubItems(7) = 0
                EliminarFichas Mid(lvArticulo.SelectedItem.Key, 2)
                RecalculoTotales
                
        End Select
        
    End If

End Sub

Private Sub lvArticulo_LostFocus()

    If lvArticulo.Tag = "1" Then
        loc_BuscarCambiosEnFicha
    End If
    lvArticulo.Tag = ""
    Status.SimpleText = vbNullString

End Sub

Private Sub AccionImprimir()
Dim Msg As String
Dim lnDocumento As Long
Dim sPiso As Boolean
Dim aUsuario As Long, strDefensa As String, sImprimoRetiro As Boolean
'Dim cCofis As Currency ', cNeto As Currency
Dim iPosArr As Integer, iResto As Integer
Dim bInstalacion As Boolean ', bInsRealizada As Boolean

    bInstalacion = False
    If Trim(labImporteNota.Caption) = vbNullString Then
        MsgBox "No hay artículos seleccionados para devolver.", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else
        iPosArr = 0
        For Each itmx In lvArticulo.ListItems
            If Val(itmx.Text) > 0 Then
                iPosArr = 1
                'No salgo hasta que iposarr > 0 y binstalacion = true
                If Not bInstalacion Then
                    bInstalacion = Val(itmx.SubItems(2)) - ExisteInstalacion(Mid(itmx.Key, 2)) < Val(itmx.Text)      ', bInsRealizada)
                End If
            End If
            'Si esta en true x obvio iposarr = 1
            If bInstalacion Then Exit For
        Next
        If iPosArr = 0 Then
            MsgBox "No hay artículos seleccionados para devolver.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
    End If
    iPosArr = 0
    
    If bInstalacion Then
        MsgBox "Existen instalaciones asociadas al documento, elimine la instalación para poder emitir la nota.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    If Val(tFicha.Tag) = 0 Then
        If Val(tEnvio.Text) <> 0 Then
            If Not SaltoAEnvio Then
                MsgBox "No es necesario ingresar un código de envío.", vbExclamation, "ATENCIÓN"
            Else
                If ValidoEnvioDevolucion(tEnvio.Text) = 0 Then
                    Foco tEnvio
                    Exit Sub
                End If
            End If
        Else
            If SaltoAEnvio Then
                MsgBox "Se van a imprimir fichas de alta de stock." & vbCrLf & "Recuerde que debe cumplirlas o asignarle un envío a la brevedad.", vbInformation, "ATENCIÓN"
            End If
        End If
        
        'Válido las devoluciones.
        For Each itmx In lvArticulo.ListItems
            If CLng(itmx.Text) > 0 And Mid(itmx.Key, 1, 1) = "A" And Val(itmx.SubItems(7)) > 0 Then
                If CantidadArticulosEnFicha(Mid(itmx.Key, 2)) > Val(itmx.Text) Then
                    MsgBox "La ficha asignada para el artículo '" & itmx.SubItems(1) & "' tiene más artículos de los que se quiere devolver, no podrá emitir la nota.", vbCritical, "ATENCIÓN"
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        Next
                    
    End If
    
    If MsgBox("¿Desea emitir la nota?", vbQuestion + vbYesNo, "EMITIR") = vbNo Then Exit Sub
    
    Dim bNegativo As Boolean
    bNegativo = False
    For Each itmx In lvArticulo.ListItems
        If CCur(itmx.SubItems(8)) < 0 Then bNegativo = True
    Next
    
    If bNegativo Then
        For Each itmx In lvArticulo.ListItems
            If CInt(itmx.SubItems(2)) <> CInt(itmx.Text) Then
                If CInt(itmx.SubItems(2)) = CInt(itmx.SubItems(6)) Then
                    'Tiene Nota, puede ser la nota del envío.
                    'Consulto si el artículo es el que paga el flete.
                    Cons = "Select TFlArticulo From TipoFlete Where TFlArticulo = " & Mid(itmx.Key, 2, Len(itmx.Key))
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If RsAux.EOF Then MsgBox "Esta factura posee valores negativos no podrá emitir la nota.", vbExclamation, "ATENCIÓN": RsAux.Close: Exit Sub
                    RsAux.Close
                Else
                    MsgBox "Esta factura posee valores negativos no podrá emitir la nota.", vbExclamation, "ATENCIÓN": Exit Sub
                End If
            End If
        Next
    End If
    
    Dim objSuceso As New clsSuceso
    objSuceso.ActivoFormulario paCodigoDeUsuario, "Emisión de Nota de Devolución", cBase
    Me.Refresh
    aUsuario = objSuceso.RetornoValor(True)
    strDefensa = objSuceso.RetornoValor(False, True)
    Set objSuceso = Nothing
    If aUsuario = 0 Then Exit Sub

    FechaDelServidor
    If prmEFacturaProductivo = "" Then CargareFacturaONOFF
    
    Dim tipoCAE As Byte
    If Val(lblNroDoc.Tag) > 0 Then
        tipoCAE = IIf(Val(lblNroDoc.Tag) = CFE_eFactura, CFE_eFacturaNotaDebito, CFE_eTicketNotaDebito)
    Else
        tipoCAE = IIf(oClienteCtdo.RUT <> "", CFE_eFacturaNotaDebito, CFE_eTicketNotaDebito)
    End If

    sImprimoRetiro = False

    On Error GoTo ErrAI
    cBase.BeginTrans
    On Error GoTo ErrResumo
    
    Screen.MousePointer = vbHourglass
    
    'Si la factura pago un servicio, no hago movimiento de stock físico.
    If Val(tFicha.Tag) > 0 Then
        Cons = "Select * From Servicio Where SerCodigo = " & Val(tFicha.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Edit
            RsAux!SerModificacion = Format(gFechaServidor, FormatoFH)
            RsAux!SerDocumento = Null
            RsAux.Update
        End If
        RsAux.Close
        
        'Veo si la factura esta en Pendientes y aún no fue liquidada.
        Cons = "Select * From DocumentoPendiente Where DPeDocumento = " & Val(tNumero.Tag) & " AND DPeFLiquidacion Is Null"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Delete
        End If
        RsAux.Close
    End If

    Cons = "Select * From Documento Where DocCodigo = " & tNumero.Tag
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close: Screen.MousePointer = vbDefault
        Msg = "No se encontro la relación al documento, reintente."
        GoTo ErrResumo
    Else
        If RsAux!DocFModificacion = CDate(labImporteTotal.Tag) Then
            Msg = "Error al intentar almacenar la información."
            'Updateo el documento, le cambio la fecha de modif.
            RsAux.Edit
            RsAux!DocFModificacion = Format(gFechaServidor, FormatoFH)
            RsAux.Update
            RsAux.Close
            
            Dim CAE As New clsCAEDocumento
            If Val(prmEFacturaProductivo) = 0 Then
                Dim nroDoc As String
                nroDoc = NumeroDocumento(paDNDevolucion)
                With CAE
                    .Desde = 1
                    .Hasta = 9999999
                    .Serie = Mid(nroDoc, 1, 1)
                    .Numero = Mid(nroDoc, 2)
                    .IdDGI = "9014113"
                    .TipoCFE = tipoCAE
                    .Vencimiento = "31/12/" & CStr(Year(Date))
                End With
            Else
                Dim caeG As New clsCAEGenerador
                Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, tipoCAE, paCodigoDeSucursal)
                Set caeG = Nothing
            End If
            Dim oDoc As New clsDocumentoCGSA
            With oDoc
                Set .Cliente = oClienteCtdo
                .Comentario = Trim(tComentario.Text)
                .Digitador = aUsuario
                .Emision = gFechaServidor
                .IVA = CCur(labIVA.Caption)
                .Moneda.Codigo = Val(labImporteDescontado.Tag)
                .Numero = CAE.Numero
                .Serie = CAE.Serie
                .Sucursal = paCodigoDeSucursal
                .Tipo = TD_NotaDevolucion
                .Total = CCur(labImporteNota.Caption)
                .Vendedor = aUsuario
                .NotaDevuelve = .Total
                .NotaSalidaCaja = .Total
            End With
            
            
'            Cons = "INSERT INTO Documento " _
'                & " (DocFecha, DocTipo, DocSerie, DocNumero, DocCliente, DocMoneda, DocTotal, DocIva, DocAnulado, DocSucursal, DocUsuario, DocFModificacion, DocComentario)" _
'                & " Values ('" & Format(gFechaServidor, FormatoFH) & "'" _
'                & ", " & TipoDocumento.NotaDevolucion _
'                & ", '" & CAE.Serie & "', " & CAE.Numero _
'                & ", " & labFechaDocumento.Tag & ", " & labImporteDescontado.Tag _
'                & ", " & CCur(labImporteNota.Caption) & ", " & CCur(labIVA.Caption) _
'                & ", 0," & paCodigoDeSucursal & ", " & aUsuario _
'                & ", '" & Format(gFechaServidor, FormatoFH) & "'"
'
'            If Trim(tComentario.Text) = vbNullString Then
'                Cons = Cons & ", Null)"
'            Else
'                Cons = Cons & ", '" & tComentario.Text & "')"
'            End If
'            cBase.Execute (Cons)
'
'            Cons = "SELECT MAX(DocCodigo) From Documento" _
'                & " WHERE DocTipo = " & TipoDocumento.NotaDevolucion _
'                & " AND DocSerie = '" & CAE.Serie & "' AND DocNumero = " & CAE.Numero
'            Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'            lnDocumento = Rs(0)
'            Rs.Close
                                
            Dim oRen As New clsDocumentoRenglon
            For Each itmx In lvArticulo.ListItems
            If CLng(itmx.Text) > 0 Then
                Set oRen = New clsDocumentoRenglon
                With oRen
                    .Articulo.ID = Mid(itmx.Key, 2, Len(itmx.Key))
                    .Articulo.Nombre = itmx.SubItems(1)
                    .Articulo.TipoIVA.Porcentaje = itmx.SubItems(9)
                    .IVA = Format(CCur(itmx.Tag), "###0.000")
                    .Precio = CCur(itmx.SubItems(8))
                    .Cantidad = itmx.Text
                    If CCur(itmx.Text) <= CCur(itmx.SubItems(3)) Then
                        .CantidadARetirar = CCur(itmx.Text)
                    Else
                        .CantidadARetirar = CCur(itmx.SubItems(3))
                    End If
                End With
                oDoc.Renglones.Add oRen
            End If
            Next
            
            Set oDoc.Conexion = cBase
            oDoc.Codigo = oDoc.InsertoDocumentoBD(Val(tNumero.Tag))
            lnDocumento = oDoc.Codigo
            
            For Each itmx In lvArticulo.ListItems
                iResto = 0
                If CLng(itmx.Text) > 0 Then
                    iResto = Val(itmx.Text)
                    'Válido cantidades
                    If Val(tFicha.Tag) = 0 And Mid(itmx.Key, 1, 1) = "A" Then
                        
                        'Veo si tengo ficha de devolución asignada, sino hago una nueva si corresponde.
                        Dim oFichasArt As clsFichasDevolucion
                        Set oFichasArt = ArticuloEnFicha(Mid(itmx.Key, 2))
                        If Not oFichasArt Is Nothing Then
                            'Voy dando de baja las fichas.
                            Dim iDto As Integer
                            iDto = 0
                            Dim oArtFicha As clsArtEnFicha
                            For Each oArtFicha In oFichasArt.Fichas
                                Cons = "Select * From Devolucion Where DevID = " & oArtFicha.IDFicha & " And DevNota Is Null"
                                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                                RsAux.Edit
                                RsAux!DevNota = lnDocumento
                                RsAux.Update
                                RsAux.Close
                                iResto = iResto - oArtFicha.Cantidad
                                iDto = iDto + oArtFicha.Cantidad
                                If iDto = Val(itmx.Text) Then
                                    Exit For
                                End If
                            Next
                        End If
                        
                        If iResto <> 0 Then
                            'Lo que queda por devolver es mayor a lo que tiene para retirar.
                            If iResto > Val(itmx.SubItems(3)) Then
                                'Creo las fichas de dev. para la diferencia.
                                Cons = "Select * From Devolucion Where DevID = 0"
                                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                                RsAux.AddNew
                                RsAux!DevFactura = Val(tNumero.Tag)
                                RsAux!DevCliente = labFechaDocumento.Tag
                                RsAux!DevNota = lnDocumento
                                RsAux!DevArticulo = Mid(itmx.Key, 2, Len(itmx.Key))
                                RsAux!DevCantidad = iResto - Val(itmx.SubItems(3))
                                If Val(tEnvio.Text) > 0 Then RsAux!DevEnvio = Val(tEnvio.Text)
                                RsAux.Update
                                RsAux.Close
                                iResto = Val(itmx.SubItems(3))
                                sImprimoRetiro = True
                            End If
                        End If  'Se dev. todo.
                        
                        If iResto > 0 Then
                            MarcoStockXDevolucion CLng(Mid(itmx.Key, 2, Len(itmx.Key))), CCur(iResto), CCur(iResto), TipoLocal.Deposito, paCodigoDeSucursal, aUsuario, TipoDocumento.NotaDevolucion, lnDocumento
                        
                            Cons = "Select * From Renglon Where RenDocumento = " & tNumero.Tag _
                                & " And RenArticulo = " & Mid(itmx.Key, 2, Len(itmx.Key))
                            
                            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                            RsAux.Edit
                            RsAux!RenARetirar = RsAux!RenARetirar - iResto
                            RsAux.Update
                            RsAux.Close
                        End If
                    End If
                                        
                    
'                    '-----------------------------------------------------------------------------------------
'                    'Inserto los renglones de la NOTA
'                    Cons = "INSERT INTO Renglon (RenDocumento, RenArticulo, RenCantidad, RenPrecio, RenIVA, RenARetirar)" _
'                        & " VALUES (" & lnDocumento & ", " & Mid(itmx.Key, 2, Len(itmx.Key)) _
'                        & ", " & itmx.Text & ", " & CCur(itmx.SubItems(8)) _
'                        & ", " & Format(CCur(itmx.Tag), "###0.000")
'
'                    If CCur(itmx.Text) <= CCur(itmx.SubItems(3)) Then
'                        Cons = Cons & ", " & CCur(itmx.Text)
'                    Else
'                        Cons = Cons & ", " & CCur(itmx.SubItems(3))
'                    End If
'                    Cons = Cons & ")"
'                    cBase.Execute (Cons)
'                    '-----------------------------------------------------------------------------------------
                    
                    'Si el artículo facturo una diferencia de Envío voy a eliminar la misma y corrijo el valor del envío
                    If Mid(itmx.Key, 2, Len(itmx.Key)) = paArticuloDiferenciaEnvio Then
                        Cons = "Select * From DiferenciaEnvio Where DEvDocumento = " & tNumero.Tag
                        Set Rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        If Not Rs.EOF Then
                            If Not IsNull(Rs!DevValorFlete) And Not IsNull(Rs!DevValorPiso) Then
                                Cons = "Update Envio Set EnvValorFlete = EnvValorFlete - " & Rs!DevValorFlete _
                                        & " , EnvIvaFlete = EnvIvaFlete - " & Rs!DevIvaFlete _
                                        & " , EnvValorPiso = EnvValorPiso - " & Rs!DevValorPiso _
                                        & " , EnvIvaPiso = EnvIvaPiso - " & Rs!DevIvaPiso
                            ElseIf Not IsNull(Rs!DevValorFlete) Then
                                Cons = "Update Envio Set EnvValorFlete = EnvValorFlete - " & Rs!DevValorFlete _
                                        & " , EnvIvaFlete = EnvIvaFlete - " & Rs!DevIvaFlete
                            Else
                                Cons = "Update Envio Set  EnvValorPiso = EnvValorPiso - " & Rs!DevValorPiso _
                                        & " , EnvIvaPiso = EnvIvaPiso - " & Rs!DevIvaPiso
                            End If
                            Cons = Cons & " Where EnvCodigo = " & Rs!DevEnvio
                            
                            cBase.Execute (Cons)
                            Rs.Delete
                        End If
                        Rs.Close
                    End If
                End If
            Next
            
'            'INSERTO RELACION NOTA
'            Cons = "INSERT INTO Nota (NotFactura, NotNota, NotDevuelve, NotSalidaCaja)" _
'                & " Values (" & tNumero.Tag & "," & lnDocumento _
'                & ", " & CCur(labImporteNota.Caption) & ", " & CCur(labImporteNota.Caption) & ")"
'            cBase.Execute (Cons)
            
            Dim oDocRel As New clsDocumentoAsociado
            With oDocRel
                .ID = Val(tNumero.Tag)
                .Devuelve = CCur(labImporteNota.Caption)
                .Fecha = labFechaDocumento.Caption
                .Numero = tNumero.Text
                .Serie = tSerie.Text
                .Tipo = TD_Contado
                If Val(lblNroDoc.Tag) > 0 Then .TipoEFactura = Val(lblNroDoc.Tag)
            End With
            oDoc.DocumentosAsociados.Add oDocRel
            
            '------------------------------------------------------------
            clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.Notas, paCodigoDeTerminal, aUsuario, lnDocumento, _
                                   Descripcion:="Nota de Devolución: " & CAE.Serie & " " & CAE.Numero, Defensa:=Trim(strDefensa)
            
            If EmitirCFE(oDoc, CAE) <> "" Then RsAux.Edit
                        
            Dim idDocError As Long
            idDocError = MayorIDDocumento(oDoc.Serie & oDoc.Numero)
            If idDocError <> oDoc.Codigo Then
                Msg = "AVISAR A ADRIAN: No encontré documento: " & Err.Description
                RsAux.Close
                RsAux.Edit
            End If
            
            Set oDoc = Nothing
            
            cBase.CommitTrans
            On Error GoTo ErrAIF
            
            ImprimoNota lnDocumento
            If oCnfgPrintSalidaCaja.Opcion = 0 Then
                ImprimoSalidaCaja lnDocumento, CInt(aUsuario)
            Else
                ImprimoSalidaCajaTicket lnDocumento, CInt(aUsuario), oCnfgPrintSalidaCaja.ImpresoraTickets
            End If
            If sImprimoRetiro Then ImprimoRetirosPorDevolucion CLng(labFechaDocumento.Tag), tNumero.Tag, lnDocumento
        Else
            RsAux.Close
            Msg = "Otra terminal modificó el documento, no podrá realizar la nota."
            GoTo ErrResumo
        End If
    End If
    On Error Resume Next
    LimpioCampos
    tSerie.Text = vbNullString
    tNumero.Text = vbNullString
    DeshabilitoCampos
    cLocal.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
            
ErrAI:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al iniciar la transacción."
    Exit Sub
    
ErrAIF:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error después de almacenar los datos."
    Exit Sub
    
ErrResumo:
    Resume Relajo

Relajo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError Msg
    Exit Sub

End Sub

Private Function MayorIDDocumento(ByVal sNroDoc As String) As Long
Dim RsDoc As rdoResultset

    Cons = "SELECT MAX(DocCodigo) From Documento" _
        & " WHERE DocTipo = 3" _
        & " AND DocSerie = '" & Mid(sNroDoc, 1, 1) & "'" _
        & " AND DocNumero = " & Mid(sNroDoc, 2, Len(sNroDoc))
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    MayorIDDocumento = RsDoc(0)
    RsDoc.Close

End Function

Private Sub RecalculoTotales()
Dim bHabilito As Boolean, bNegativo As Boolean
          
    labImporteNota.Caption = "0.00"
    labIVA.Caption = "0.00"
    bHabilito = False
    For Each itmx In lvArticulo.ListItems
        If CLng(itmx.Text) > 0 Then
            labImporteNota.Caption = CCur(labImporteNota.Caption) + (CLng(itmx.Text) * CCur(itmx.SubItems(8)))
            'labIVA.Caption = CCur(labIVA.Caption) + (CLng(itmx.Text) * CCur(itmx.SubItems(8))) - ((CLng(itmx.Text) * CCur(itmx.SubItems(8))) / ((CCur(itmx.Tag) / 100) + 1)))
            labIVA.Caption = CCur(labIVA.Caption) + (CLng(itmx.Text) * CCur(itmx.Tag))
            bHabilito = True
        End If
        If CCur(itmx.SubItems(8)) < 0 Then bNegativo = True
    Next
    
    labImporteNota.Caption = Format(labImporteNota.Caption, FormatoMonedaP)
    labIVA.Caption = Format(labIVA.Caption, FormatoMonedaP)
    
    If bNegativo Then MsgBox "Existe un importe negativo, no se podrá emitir nota parcial.", vbInformation, "ATENCIÓN"

    If CCur(labImporteNota.Caption) >= 0 Then
        MnuEmitir.Enabled = bHabilito
        Toolbar1.Buttons("imprimir").Enabled = bHabilito
        If CCur(labImporteNota.Caption) > 0 Then
            HabilitoCampos
        ElseIf bHabilito Then
            tComentario.Enabled = True: tComentario.BackColor = vbWindowBackground
        End If
    Else
        MnuEmitir.Enabled = False
        Toolbar1.Buttons("imprimir").Enabled = False
    End If

End Sub

Private Sub CargoCliente(ByVal idDoc As Long, ByVal idCliente As Long)

    lblNroDoc.Tag = ""
    'Intento levantar la información del ecomprobante.
    Dim rsCliente As rdoResultset
    Set oClienteCtdo = New clsClienteCFE
    Cons = "SET QUOTED_IDENTIFIER ON SET CONCAT_NULL_YIELDS_NULL ON SET ANSI_PADDING ON SET ANSI_WARNINGS ON SET ANSI_NULLS ON SET ARITHABORT ON SELECT EComTipo, EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:TipoDocRecep)[1]', 'tinyint') TipoDoc, " & _
            "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:DocRecep)[1]', 'nvarchar(20)') Documento, " & _
            "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:RznSocRecep)[1]', 'nvarchar(100)') Nombre, " & _
            "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:DirRecep)[1]', 'nvarchar(100)') Direccion, " & _
            "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:CiudadRecep)[1]', 'nvarchar(20)') Localidad, " & _
            "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:DeptoRecep)[1]', 'nvarchar(20)') Departamento " & _
            "FROM eComprobantes WHERE EComID = " & idDoc
    Set rsCliente = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsCliente.EOF Then
        With oClienteCtdo
            .Codigo = idCliente
            If Not IsNull(rsCliente("Nombre")) Then .NombreCliente = Trim(rsCliente!Nombre)
            If Not IsNull("EComTipo") Then lblNroDoc.Tag = rsCliente("EComTipo")
            If Not IsNull(rsCliente("Direccion")) Then .Direccion.Domicilio = Trim(rsCliente("Direccion"))
            If Not IsNull(rsCliente("Departamento")) Then .Direccion.Departamento = Trim(rsCliente("Departamento"))
            If Not IsNull(rsCliente("Localidad")) Then .Direccion.Localidad = Trim(rsCliente("Localidad"))
            If Not IsNull(rsCliente("TipoDoc")) Then
                .CodigoDGICI = rsCliente("TipoDoc")
                If .CodigoDGICI = TD_RUT Then
                    .RUT = rsCliente("Documento")
                Else
                    .CI = rsCliente("Documento")
                End If
            End If
        End With
    End If
    rsCliente.Close
    
    If oClienteCtdo.Codigo = 0 Then
    
        Cons = "SELECT CliCiRuc, CliTipo, PDDTipoDocIdentidad, DepNombre, LocNombre, CalNombre, DirPuerta, IsNull(DirApartamento, '') Apartamento, Nombre = CASE WHEN CliTipo = 1 THEN (RTrim(CPeApellido1) + RTrim(' ' + ISNULL(CPeApellido2, ''))+', ' + RTrim(CPeNombre1)) + RTrim(' ' + ISNULL(CPeNombre2,'')) ELSE RTRIM(CEmNombre) END, Ruc = CPeRuc " _
            & " FROM Cliente INNER JOIN PaisDelDocumento ON CliPaisDelDocumento = PDDId " _
            & " LEFT OUTER JOIN CPersona ON CliCodigo = CPeCliente " _
            & " LEFT OUTER JOIN CEmpresa ON CliCodigo = CEmCliente " _
            & " LEFT OUTER JOIN Direccion ON DirCodigo = CliDireccion " _
            & " LEFT OUTER JOIN Calle ON DirCalle = CalCodigo " _
            & " LEFT OUTER JOIN Localidad ON CalLocalidad = LocCodigo LEFT OUTER JOIN Departamento ON LocDepartamento = DepCodigo " _
            & " WHERE CliCodigo = " & idCliente
    
        Set rsCliente = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        With oClienteCtdo
            .Codigo = idCliente
            If Not IsNull(rsCliente("Nombre")) Then .NombreCliente = Trim(rsCliente!Nombre)
            If Not IsNull(rsCliente("DepNombre")) Then
                .Direccion.Domicilio = Trim(rsCliente("CalNombre")) & " " & Trim(rsCliente("DirPuerta")) & " " & Trim(rsCliente("Apartamento"))
                .Direccion.Departamento = Trim(rsCliente("DepNombre"))
                .Direccion.Localidad = Trim(rsCliente("LocNombre"))
            End If
            
            .TipoCliente = rsCliente("CliTipo")
            .CodigoDGICI = rsCliente("PDDTipoDocIdentidad")
            If Not IsNull(rsCliente("CliCiRUC")) Then
                If .TipoCliente = TC_Empresa Then
                    .RUT = rsCliente("CliCiRUC")
                Else
                    .CI = rsCliente("CliCiRUC")
                End If
            End If
            If Not IsNull(rsCliente("Ruc")) Then .RUT = rsCliente("Ruc")
        End With
        rsCliente.Close
    End If
    If oClienteCtdo.RUT <> "" Then
        labDato1.Caption = "R.U.C.:"
        lblRUT.Caption = clsGeneral.RetornoFormatoRuc(oClienteCtdo.RUT)
    Else
        labDato1.Caption = "Cédula"
        lblRUT.Caption = clsGeneral.RetornoFormatoCedula(Trim(oClienteCtdo.CI))
    End If
        
    labNombre.Caption = " " & oClienteCtdo.NombreCliente
    labDireccion.Caption = oClienteCtdo.Direccion.Domicilio
    
End Sub

Private Sub FormEnvio()
On Error GoTo ErrFE
Dim aIdEnvio As Long

    aIdEnvio = 0
    If CodDocumentoEnvio = 0 Then
        Cons = "Select EnvCodigo From Envio Where EnvDocumento = " & tNumero.Tag
    Else
        Cons = "Select EnvCodigo From Envio Where EnvDocumento = " & CodDocumentoEnvio
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then aIdEnvio = RsAux!EnvCodigo
    RsAux.Close
    
    If aIdEnvio <> 0 Then
        Dim objEnvio As New clsEnvio
        objEnvio.InvocoEnvio aIdEnvio, gPathListados
        Set objEnvio = Nothing
        Me.Refresh
        
        tNumero_KeyPress (vbKeyReturn)
    Else
        
    End If
    Exit Sub
    
ErrFE:
    clsGeneral.OcurrioError " Ocurrio un error inesperado." & Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargoLocales()
On Error GoTo ErrCL
    cLocal.Clear
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cLocal, ""
    Exit Sub
ErrCL:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los locales."
    Screen.MousePointer = vbDefault
End Sub

Private Sub ImprimoNota(Documento As Long)
On Error GoTo ErrCrystal
Dim result As Integer, JobSRep1 As Integer, JobSRep2 As Integer, jobnum As Integer
Dim NombreFormula As String, CantForm As Integer, aTexto As String

    Screen.MousePointer = 11
    'Inicializo el Reporte y SubReportes
    jobnum = crAbroReporte(gPathListados & "NotaDevolucion.RPT")
    If jobnum = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    If ChangeCnfgPrint Then
        prj_LoadConfigPrint False
    
        lPNC.Caption = "Imp. Nota:" & paINContadoN
        If Not paPrintEsXDefNC Then lPNC.ForeColor = &HC0& Else lPNC.ForeColor = vbBlack
    
        lPCn.Caption = "Imp. Salida Caja: " & IIf(oCnfgPrintSalidaCaja.Opcion = 0, paIConformeN, oCnfgPrintSalidaCaja.ImpresoraTickets)
        If Not paPrintEsXDefCn Then lPCn.ForeColor = &HC0& Else lPNC.ForeColor = vbBlack
    End If
    
    If Trim(Printer.DeviceName) <> Trim(paINContadoN) Then SeteoImpresoraPorDefecto paINContadoN
    If Not crSeteoImpresora(jobnum, Printer, paINContadoB) Then GoTo ErrCrystal
    
    'Obtengo la cantidad de formulas que tiene el reporte.
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte Contado --------------------------------------------------------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento": result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDNDevolucion & "'")
            Case "cliente"
                If labDato1.Caption <> "R.U.C.:" And Trim(lblRUT.Caption) <> vbNullString Then
                    aTexto = "(" & Trim(lblRUT.Caption) & ")"
                Else
                    aTexto = ""
                End If
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labNombre.Caption) & aTexto & "'")
            Case "direccion": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labDireccion.Caption) & "'")
            Case "ruc"
                If labDato1.Caption = "R.U.C.:" And Trim(lblRUT.Caption) <> vbNullString Then
                    aTexto = clsGeneral.RetornoFormatoRuc(lblRUT.Caption)
                Else
                    aTexto = vbNullString
                End If
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(aTexto) & "'")
            
            Case "codigobarras":
                    result = crSeteoFormula(jobnum%, NombreFormula, "''")
                    'Result = crSeteoFormula(JobNum%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.NotaDevolucion, Documento) & "'")
                    
            Case "signomoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'" & BuscoSignoMoneda(labImporteDescontado.Tag) & "'")
            Case "nombremoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'(" & BuscoNombreMoneda(labImporteDescontado.Tag) & ")'")
            Case "textoretira"
                'Detallamos el documento al cual se le hace la nota.
                aTexto = "'" & Trim(cLocal.Text) & " " & Trim(tSerie.Text) & " " & Trim(tNumero.Text) & "'"
                result = crSeteoFormula(jobnum%, NombreFormula, aTexto)
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where DocCodigo = " & Documento
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
        
    'Subreporte srContado.rpt  y srContado.rpt - 01-----------------------------------------------------------------------------
    JobSRep1 = crAbroSubreporte(jobnum, "srContado.rpt")
    If JobSRep1 = 0 Then GoTo ErrCrystal
    
    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"
    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
    
    JobSRep2 = crAbroSubreporte(jobnum, "srContado.rpt - 01")
    If JobSRep2 = 0 Then GoTo ErrCrystal
    If crSeteoSqlQuery(JobSRep2, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------
    
    'If crMandoAPantalla(JobNum, "Factura Contado") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
    If Not crCierroSubReporte(JobSRep2) Then GoTo ErrCrystal
    
    'crEsperoCierreReportePantalla
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroSubReporte JobSRep1
    crCierroSubReporte JobSRep2
    Screen.MousePointer = 0
    Exit Sub
End Sub
Private Sub ImprimoSalidaCaja(Nota As Long, Usuario As Integer)
Dim aTexto As String
Dim NombreFormula As String, result As Integer
Dim JobNumMC As Integer, CantFormMC As Integer

    'Inicializa el Engine del Crystal y setea la impresora para el JOB
    On Error GoTo ErrCrystal
    
    'Inicializo el Reporte y SubReportes
    JobNumMC = crAbroReporte(gPathListados & "MovimientoNota.RPT")
    If JobNumMC = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    If ChangeCnfgPrint Then
        prj_LoadConfigPrint False
    
        lPNC.Caption = "Imp. Nota:" & paINContadoN
        If Not paPrintEsXDefNC Then lPNC.ForeColor = &HC0& Else lPNC.ForeColor = vbBlack
    
        lPCn.Caption = "Imp. Salida Caja: " & IIf(oCnfgPrintSalidaCaja.Opcion = 0, paIConformeN, oCnfgPrintSalidaCaja.ImpresoraTickets)
        If Not paPrintEsXDefCn Then lPCn.ForeColor = &HC0& Else lPNC.ForeColor = vbBlack
    End If
    
    If Trim(Printer.DeviceName) <> Trim(paIConformeN) Then SeteoImpresoraPorDefecto paIConformeN
    If Not crSeteoImpresora(JobNumMC, Printer, paIConformeB) Then GoTo ErrCrystal

    'Obtengo la cantidad de formulas que tiene el reporte.
    CantFormMC = crObtengoCantidadFormulasEnReporte(JobNumMC)
    If CantFormMC = -1 Then GoTo ErrCrystal
    
    Cons = "Select * from Documento Where DocCodigo = " & Nota
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    For I = 0 To CantFormMC - 1
        NombreFormula = crObtengoNombreFormula(JobNumMC, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "entradasalida": result = crSeteoFormula(JobNumMC%, NombreFormula, "'SALIDA DE CAJA'")
            Case "sucursal": result = crSeteoFormula(JobNumMC%, NombreFormula, "'Sucursal: " & BuscoNombreSucursal(paCodigoDeSucursal) & "'")
            Case "comentario": result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & Trim(tComentario.Text) & "'")
            Case "importe": result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & Format(RsAux!DocTotal, FormatoMonedaP) & "'")
            Case "tipo"
                aTexto = "FACTURA " & Trim(tSerie.Text) & " " & Trim(tNumero.Text)
                If Not RsAux.EOF Then aTexto = "N. DEVOLUCIÓN " & RsAux!DocSerie & RsAux!Docnumero & " sobre " & aTexto
                result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & aTexto & "'")
                
            Case "moneda": result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & BuscoSignoMoneda(labImporteDescontado.Tag) & "'")
            Case "usuario": result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & BuscoInicialUsuario(Usuario) & "'")
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    RsAux.Close
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT * " _
            & " From " & paBD & ".dbo.MovimientoDisponibilidad MovimientoDisponibilidad, " _
                            & paBD & ".dbo.MovimientoDisponibilidadRenglon MovimientoDisponibilidadRenglon, " _
                            & paBD & ".dbo.Disponibilidad Disponibilidad " _
            & " Where MDiID = 0" _
            & " And MDiID = MDRIdMovimiento And MDRIdDisponibilidad = DisID"
            
    If crSeteoSqlQuery(JobNumMC%, Cons) = 0 Then GoTo ErrCrystal
            
    'If crMandoAPantalla(JobNumMC, "Movimiento de Caja") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(JobNumMC, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(JobNumMC, True, False) Then GoTo ErrCrystal
    
    'crEsperoCierreReportePantalla

    crCierroTrabajo JobNumMC
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    On Error Resume Next
    crCierroTrabajo JobNumMC
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
End Sub

Private Sub ImprimoSalidaCajaTicket(Nota As Long, Usuario As Integer, Impresora As String)
Dim aTexto As String
    
    Dim oCol As New Collection
    Dim oDocAimprimir As New clsDocAImprimir
    
    oDocAimprimir.TipoDocumento = Imp_MovimientoCaja
    oDocAimprimir.IDDocumento = 0
    oCol.Add oDocAimprimir
    
    Dim oCnfg As New clsConfigImpresora

    Dim oDocs As clsDocAImprimir
    Dim oPrint As New clsImpresionDeDocumentos
    
    Set oPrint.Conexion = cBase
    oPrint.PathReportes = gPathListados
    oPrint.NombreBaseDatos = miconexion.RetornoPropiedad(False, False, False, True)
    
    oCnfg.Impresora = Impresora
    Set oPrint.DondeImprimo = oCnfg
    
    Cons = "Select * from Documento Where DocCodigo = " & Nota
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
    oPrint.ImprimoSalidaCajaTicketPorPrm 0, BuscoNombreSucursal(paCodigoDeSucursal) _
                , BuscoInicialUsuario(Usuario), _
                "N. DEVOLUCIÓN " & RsAux!DocSerie & RsAux!Docnumero _
                , Now, " ", BuscoSignoMoneda(labImporteDescontado.Tag) & " " & Format(RsAux!DocTotal, FormatoMonedaP), Trim(tComentario.Text), "Salida de caja"
    
    RsAux.Close
    
End Sub

Private Sub ImprimoRetirosPorDevolucion(Cliente As Long, Factura As Long, Nota As Long)
Dim aTexto As String

    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    SeteoImpresoraPorDefecto paIConformeN
    
    With vsFicha
        .paperSize = paIConformeP
        .PaperBin = paIConformeB
        .Orientation = orLandscape
        
        .StartDoc: .EndDoc
        
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión para los retiros." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
        
        .FileName = "Retiros por Devolucion"
        .FontSize = 8.25
        .TableBorder = tbNone
        
        .TextAlign = taRightBaseline: .FontBold = True
        .AddTable ">2000|<3000", "RETIRO:| Contado " & Trim(tSerie.Text) & " " & Trim(tNumero.Text), ""
        .TextAlign = taLeftBaseline
        
        .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = ""
        
        .FontBold = False
        .FontSize = 24: .FontName = "3 of 9 Barcode"
        .Paragraph = CodigoDeBarras(TipoDocumento.NotaDevolucion, Nota)
        .FontName = "Tahoma": .FontSize = 8.25

        .FontBold = True
        .AddTable "^2400|<3800", CodigoDeBarras(TipoDocumento.NotaDevolucion, Nota) & "|RETIRO POR DEVOLUCION DE MERCADERIA", ""
         .Paragraph = "": .FontBold = False
         
        .AddTable "<900|<1800", "Fecha:|" & Format(gFechaServidor, "d-Mmm yyyy hh:mm"), ""
        .AddTable "<900|<5100", "Cliente:|" & Trim(labNombre.Caption), ""
        
        .Paragraph = "": .Paragraph = ""
        
        .AddTable "<6000|<1000|<1200", "Artículo|Cantidad|Devolución", ""
        
        Cons = "Select * From Devolucion, Articulo " & _
                " Where DevCliente = " & Cliente & _
                " And DevFactura = " & Factura & _
                " And DevNota = " & Nota & _
                " And DevLocal is Null And DevArticulo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            .AddTable "<6000|<1000|<1200", Format(RsAux!ArtCodigo, "(#,000,000) ") & Trim(RsAux!ArtNombre) & "|" & RsAux!DevCantidad & "|" & RsAux!DevID, ""
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        .EndDoc
        .PrintDoc   'Cliente
        .PrintDoc
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al realizar la impresión de los retiros.", Err.Description
End Sub

Private Function ValidoEnvioDevolucion(ByVal lEnvioDev As Long) As Long
Dim rsEnv As rdoResultset
    ValidoEnvioDevolucion = 0
        Cons = "Select * From Envio Where EnvCodigo = " & lEnvioDev _
        & " And EnvEstado in (" & EstadoEnvio.AConfirmar & ", " & EstadoEnvio.AImprimir & ")" _
        & " And EnvCodigo Not In (Select DevEnvio From Devolucion Where DevEnvio = " & lEnvioDev & ")"
    Set rsEnv = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsEnv.EOF Then
        ValidoEnvioDevolucion = lEnvioDev
    End If
    rsEnv.Close
    If ValidoEnvioDevolucion = 0 Then MsgBox "Envío incorrecto, no podrá emitir la nota.", vbInformation, "ATENCIÓN"
End Function

Private Function BuscoSignoMoneda(Codigo As Variant) As String
On Error GoTo ErrBU
    
Dim Rs As rdoResultset

    BuscoSignoMoneda = ""

    Cons = "SELECT * FROM Moneda WHERE MonCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not Rs.EOF Then BuscoSignoMoneda = Trim(Rs!MonSigno)
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Function BuscoNombreMoneda(Codigo As Long) As String

    On Error GoTo ErrBU
    Dim Rs As rdoResultset
    BuscoNombreMoneda = ""

    Cons = "SELECT * FROM Moneda WHERE MonCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then BuscoNombreMoneda = Trim(Rs!MonNombre)
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Function BuscoNombreSucursal(Codigo As Long) As String
On Error GoTo ErrBU
    
Dim Rs As rdoResultset

    BuscoNombreSucursal = ""

    Cons = "SELECT * FROM Sucursal WHERE SucCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not Rs.EOF Then BuscoNombreSucursal = Trim(Rs!SucAbreviacion)
    
    Rs.Close
    Exit Function
ErrBU:
End Function

Function BuscoInicialUsuario(Codigo As Integer) As String
On Error GoTo ErrBU
Dim Rs As rdoResultset
    BuscoInicialUsuario = ""
    Cons = "SELECT * FROM USUARIO WHERE UsuCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then BuscoInicialUsuario = Trim(Rs!UsuInicial)
    Rs.Close
    Exit Function
ErrBU:
End Function

Private Sub LimpioCampos()
    CodDocumentoEnvio = 0
    labNombre.Caption = vbNullString
    labDireccion.Caption = vbNullString
    lblRUT.Caption = ""
    tNumero.Tag = vbNullString
    lvArticulo.ListItems.Clear
    labFechaDocumento.Caption = vbNullString
    lblCFE.Caption = vbNullString
    labFechaDocumento.Tag = vbNullString
    labImporteTotal.Caption = vbNullString
    labImporteDescontado.Caption = vbNullString
    labImporteNota.Caption = vbNullString
    labIVA.Caption = vbNullString
    tComentario.Text = vbNullString
    tEnvio.Text = vbNullString
    tFicha.Text = vbNullString: tFicha.Tag = ""
    Set oFichas = New Collection
    Set oClienteCtdo = Nothing
    lblNroDoc.Tag = ""
End Sub
Private Sub DeshabilitoCampos()
    MnuEmitir.Enabled = False: Toolbar1.Buttons("imprimir").Enabled = False
    MnuEnvio.Enabled = False: Toolbar1.Buttons("envio").Enabled = False
    MnuRemito.Enabled = False: Toolbar1.Buttons("remito").Enabled = False
    lvArticulo.Enabled = False
    tComentario.Enabled = False: tComentario.BackColor = vbButtonFace
    tFicha.Enabled = False: tFicha.BackColor = vbButtonFace
    tEnvio.Enabled = False: tEnvio.BackColor = vbButtonFace
End Sub
Private Sub HabilitoCampos()
    lvArticulo.Enabled = True
    tComentario.Enabled = True: tComentario.BackColor = vbWindowBackground
    If Val(tFicha.Tag) = 0 Then
        tFicha.Enabled = True: tFicha.BackColor = vbWindowBackground
        tEnvio.Enabled = True: tEnvio.BackColor = vbWindowBackground
    End If
End Sub

Private Function EsServicio() As Long
Dim rsSer As rdoResultset
    EsServicio = 0
    Cons = "Select * From Servicio Where SerDocumento = " & Val(tNumero.Tag)
    Set rsSer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsSer.EOF Then EsServicio = rsSer!SerCodigo
    rsSer.Close
End Function

Private Sub BuscoDevolucionesDocumento(Optional iIDArt As Long = 0)
'Busco si el producto tiene ficha de devolución ingresada.
On Error GoTo ErrVN
Dim lDev As Long

    Screen.MousePointer = 11
    'Veo si existe una devolución para el documento.
    Cons = "Select * From Devolucion Where DevFactura = " & Val(tNumero.Tag) _
        & " And DevNota Is Null And DevLocal Is Not Null And DevFAltaLocal Is Not Null" _
        & " And DevAnulada Is Null"
        
    If iIDArt > 0 Then
        Cons = Cons & " And DevArticulo = " & iIDArt
    End If

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF
            For Each itmx In lvArticulo.ListItems
                If Val(Mid(itmx.Key, 2, Len(itmx.Key))) = RsAux!DevArticulo Then
                    If RsAux!DevCantidad > Val(itmx.Text) Then
                        Screen.MousePointer = 0
                        If iIDArt = 0 Then
                            MsgBox "Existe una ficha de devolución para el artículo: " & Trim(itmx.SubItems(1)) & " pero la cantidad supera lo que está marcado para devolución.", vbExclamation, "ATENCIÓN"
                        End If
                        Exit For
                    Else
                        itmx.SubItems(7) = RsAux!DevID
                        AgregarFicha RsAux("DevArticulo"), RsAux("DevCantidad"), RsAux("DevID")
                    End If
                End If
            Next
            RsAux.MoveNext
        Loop
    End If
    RsAux.Close
    
    Dim iCant As Integer
    
    If iIDArt = 0 Then
        'Busco x sgdo intento a aquellos artículos que no lo encontre por documento.
        For Each itmx In lvArticulo.ListItems
            If Mid(itmx.Key, 1, 1) = "A" And Val(itmx.Text) > 0 _
                And Val(itmx.SubItems(3)) <> Val(itmx.SubItems(2)) _
                And Val(itmx.SubItems(3)) < Val(itmx.Text) And Val(itmx.SubItems(7)) = 0 Then
                iCant = itmx.Text
                lDev = BuscoDevolucionArticuloCliente(Mid(itmx.Key, 2, Len(itmx.Key)), iCant)
                If lDev > 0 Then
                    itmx.SubItems(7) = lDev
                    AgregarFicha Mid(itmx.Key, 2), iCant, lDev
                End If
                
            End If
        Next
    Else
        lDev = BuscoDevolucionArticuloCliente(iIDArt, iCant)
        If lDev > 0 Then
            itmx.SubItems(7) = lDev
            AgregarFicha Mid(itmx.Key, 2), iCant, lDev
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrVN:
    clsGeneral.OcurrioError "Error al buscar fichas de devolución para el documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function BuscoDevolucionArticuloCliente(ByVal lArticulo As Long, ByRef iCantidad As Integer) As Long

    BuscoDevolucionArticuloCliente = 0
    'Busco para el artículo y el cliente.
    Cons = "Select * From Devolucion Where DevArticulo = " & lArticulo _
        & " And DevCliente = " & Val(labFechaDocumento.Tag) _
        & " And DevCantidad <= " & iCantidad _
        & " And DevNota IS Null And DevLocal Is Not Null And DevFAltaLocal Is Not Null" _
        & " And DevAnulada Is Null And DevFactura Is Null"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        'Encontre ficha.
        'Válido que sea única.
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            BuscoDevolucionArticuloCliente = RsAux!DevID
            iCantidad = RsAux!DevCantidad
        End If
    End If
    RsAux.Close
    
End Function

Private Function SaltoAFicha() As Boolean
    SaltoAFicha = False
    If Not tFicha.Enabled Then Exit Function
    For Each itmx In lvArticulo.ListItems
        If Mid(itmx.Key, 1, 1) = "A" And Val(itmx.Text) > 0 And Val(itmx.SubItems(7)) = 0 And Val(itmx.Text) > Val(itmx.SubItems(3)) Then
            SaltoAFicha = True
        End If
    Next
End Function

Private Sub BuscoDevolucion(ByVal lDev As Long)
Screen.MousePointer = 11
Dim lAux As Long
    
    'Veo si existe una devolución para el documento.
    Cons = "Select * From Devolucion Where DevID = " & lDev
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If Not RsAux.EOF Then
        If IsNull(RsAux!DevAnulada) And IsNull(RsAux!DevNota) And Not IsNull(RsAux!DevLocal) And Not IsNull(RsAux!DevFAltaLocal) Then
            
            If Not IsNull(RsAux!DevFactura) Then
                If RsAux!DevFactura <> Val(tNumero.Tag) Then
                    Screen.MousePointer = 0
                    RsAux.Close
                    MsgBox "La devolución está asignada a otro documento.", vbCritical, "ATENCIÓN"
                    Exit Sub
                End If
            Else
                'Válido que no exista una devolución para este artículo con la factura.
                lAux = DevolucionParaDocumentoArticulo(RsAux!DevArticulo)
                If lAux > 0 Then
                    RsAux.Close
                    lDev = lAux
                    Cons = "Select * From Devolucion Where DevID = " & lDev
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
                End If
            End If
            
            For Each itmx In lvArticulo.ListItems
                If Val(Mid(itmx.Key, 2)) = RsAux!DevArticulo Then
                    
                    If RsAux!DevCantidad > Val(itmx.Text) Then
                        Screen.MousePointer = 0
                        RsAux.Close
                        MsgBox "La ficha de devolución para el artículo: " & Trim(itmx.SubItems(1)) & " supera la cantidad de lo que es posible devolver.", vbExclamation, "ATENCIÓN"
                        Exit Sub
                    Else
                        'Si todo lo que compro esta en la factura no la tomo.
                        If Val(itmx.SubItems(2)) <> Val(itmx.SubItems(3)) Then
                            If RsAux!DevCliente <> Val(labFechaDocumento.Tag) Then
                                If MsgBox("El cliente de la ficha no es el del documento." & vbCrLf & "¿Desea ingresar la ficha de todas formas?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible error") = vbYes Then
                                    itmx.SubItems(7) = RsAux!DevID
                                    AgregarFicha RsAux!DevArticulo, RsAux!DevCantidad, RsAux!DevID
                                End If
                            Else
                                itmx.SubItems(7) = RsAux!DevID
                                AgregarFicha RsAux!DevArticulo, RsAux!DevCantidad, RsAux!DevID
                            End If
                        End If
                    End If
                End If
            Next
        Else
            MsgBox "La devolución ingresada no es válida.", vbCritical, "ATENCIÓN"
        End If
    Else
        MsgBox "La devolución ingresada no es válida.", vbCritical, "ATENCIÓN"
    End If
    RsAux.Close
    tFicha.Text = ""
    Screen.MousePointer = 0
    Exit Sub
    
ErrVN:
    clsGeneral.OcurrioError "Ocurrio un error al validar la nota.", Err.Description
    Screen.MousePointer = 0

End Sub

Private Function SaltoAEnvio() As Boolean
Dim iPos As Integer

    SaltoAEnvio = False
    If Not tEnvio.Enabled Then Exit Function
    For Each itmx In lvArticulo.ListItems
        If Mid(itmx.Key, 1, 1) = "A" And Val(itmx.Text) > 0 And Val(itmx.SubItems(3)) <> Val(itmx.Text) Then
            If CantidadArticulosEnFicha(Mid(itmx.Key, 2)) + Val(itmx.SubItems(3)) < Val(itmx.Text) Then
                SaltoAEnvio = True
                Exit Function
            ElseIf Val(itmx.SubItems(3)) > 0 And Val(itmx.SubItems(3)) < Val(itmx.Text) Then
                SaltoAEnvio = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function DevolucionParaDocumentoArticulo(ByVal lArt As Long) As Long
Dim rsDPD As rdoResultset

    DevolucionParaDocumentoArticulo = 0
    Cons = "Select * From Devolucion Where DevFactura = " & Val(tNumero.Tag) _
        & " And DevArticulo = " & lArt _
        & " And DevNota IS Null And DevLocal Is Not Null And DevFAltaLocal Is Not Null" _
        & " And DevAnulada Is Null"
        
    Set rsDPD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If Not rsDPD.EOF Then
        MsgBox "Existe una Ficha para el documento y el artículo que no es la que Ud. ingresó." & vbCrLf & "Se cargará la Ficha encontrada.", vbExclamation, "ATENCIÓN"
        DevolucionParaDocumentoArticulo = rsDPD!DevID
    End If
    rsDPD.Close
    
End Function
', ByRef bInstalada As Boolean
Private Function ExisteInstalacion(ByVal lArtID As Long) As Long
Dim rsIns As rdoResultset
On Error GoTo errEI
    
    ExisteInstalacion = False
'    bInstalada = False
    Cons = "Select IsNull(Sum(RInCantidad), 0) From Instalacion, RenglonInstalacion " _
        & " Where InsDocumento = " & Val(tNumero.Tag) _
        & " And InsID = RInInstalacion AND InsFechaRealizada IS Null And InsAnulada Is Null AND RInArticulo = " & lArtID
    Set rsIns = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsIns.EOF Then
 '       bInstalada = Not IsNull(rsIns!InsFechaRealizada)
        ExisteInstalacion = rsIns(0)
    End If
    Exit Function
errEI:
    clsGeneral.OcurrioError "Error al buscar si existen instalaciones para el documento.", Err.Description
End Function

Private Sub s_LoadMenuOpcionPrint()
Dim vOpt() As String
Dim iQ As Integer
    
    MnuPrintLine1.Visible = (paOptPrintList <> "")
    MnuPrintOpt(0).Visible = (paOptPrintList <> "")
    
    If paOptPrintList = "" Then
        Exit Sub
    ElseIf InStr(1, paOptPrintList, "|", vbTextCompare) = 0 Then
        MnuPrintOpt(0).Caption = paOptPrintList
    Else
        vOpt = Split(paOptPrintList, "|")
        For iQ = 0 To UBound(vOpt)
            If iQ > 0 Then Load MnuPrintOpt(iQ)
            With MnuPrintOpt(iQ)
                .Caption = Trim(vOpt(iQ))
                .Checked = (LCase(.Caption) = LCase(paOptPrintSel))
                .Visible = True
            End With
        Next
    End If
    
End Sub

Private Sub loc_BuscarCambiosEnFicha()
On Error Resume Next
    If lvArticulo.Tag = "" Then Exit Sub
    lvArticulo.Tag = ""
    For Each itmx In lvArticulo.ListItems
        If Val(itmx.SubItems(7)) = 0 And Val(itmx.Text) > 0 Then
            BuscoDevolucionesDocumento Mid(itmx.Key, 2, Len(itmx.Key))
        End If
    Next
End Sub

Private Function EmitirCFE(ByVal Documento As clsDocumentoCGSA, ByVal CAE As clsCAEDocumento) As String
On Error GoTo errEC
    With New clsCGSAEFactura
        .URLAFirmar = prmURLFirmaEFactura
        .TasaBasica = TasaBasica
        .TasaMinima = TasaMinima
        .ImporteConInfoDeCliente = prmImporteConInfoCliente
        Set .Connect = cBase
        If Not .GenerarEComprobante(CAE, Documento, EmpresaEmisora, paCodigoDeSucursal) Then
            EmitirCFE = .XMLRespuesta
        End If
    End With
    Exit Function
errEC:
    EmitirCFE = "Error en firma: " & Err.Description
End Function


Private Sub EnvioALog(ByVal Texto As String)
On Error GoTo errEAL
    Open "\\ibm3200\oyr\efactura\logEFactura.txt" For Append As #1
    Print #1, Now & Space(5) & "Terminal: " & miconexion.NombreTerminal & Space(5); "NOTA CONTADO" & Space(5) & Texto
    Close #1
    Exit Sub
errEAL:
End Sub

Private Function ArticuloEnFicha(ByVal idArticulo As Long) As clsFichasDevolucion
    Set ArticuloEnFicha = Nothing
    Dim oFicha As clsFichasDevolucion
    For Each oFicha In oFichas
        If oFicha.Articulo = idArticulo Then
            Set ArticuloEnFicha = oFicha
            Exit Function
        End If
    Next
End Function

Private Function CantidadArticulosEnFicha(ByVal idArticulo As Long) As Integer
    Dim oFicha As clsFichasDevolucion
    Set oFicha = ArticuloEnFicha(idArticulo)
    If Not oFicha Is Nothing Then
        CantidadArticulosEnFicha = oFicha.CantidadTotalEnFichas()
    End If
End Function

Private Sub EliminarFichas(ByVal idArticulo As Long)
    Dim oFicha As clsFichasDevolucion
    Dim ixF As Integer
    For ixF = 1 To oFichas.Count
        Set oFicha = oFichas(ixF)
        If oFicha.Articulo = idArticulo Then
            oFichas.Remove ixF
            Exit Sub
        End If
    Next
End Sub


Private Sub AgregarFicha(ByVal idArticulo As Long, ByVal Cantidad As Integer, ByVal IDFicha As Long)
    Dim oFicha As clsFichasDevolucion
    Set oFicha = ArticuloEnFicha(idArticulo)
    If oFicha Is Nothing Then
        Set oFicha = New clsFichasDevolucion
        oFicha.Articulo = idArticulo
        oFichas.Add oFicha
    Else
        'Veo si ya fue ingresada.
        Dim ficha As clsArtEnFicha
        For Each ficha In oFicha.Fichas
            If ficha.IDFicha = IDFicha Then
                MsgBox "La ficha " & IDFicha & " ya fue ingresada.", vbExclamation, "Ficha duplicada"
            End If
        Next
    End If
    oFicha.AddFicha IDFicha, Cantidad
    
End Sub

Private Function ValidarVersionEFactura() As Boolean
On Error GoTo errEC
    With New clsCGSAEFactura
        ValidarVersionEFactura = .ValidarVersion()
    End With
    Exit Function
errEC:
End Function

Private Sub CargareFacturaONOFF()
Dim rsP As rdoResultset
Dim sQy As String
    sQy = "SELECT IsNull(ParValor, 0) FROM Parametro WHERE ParNombre = 'eFacturaActiva'"
    Set rsP = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not rsP.EOF Then
        prmEFacturaProductivo = rsP(0)
    End If
    rsP.Close
End Sub
