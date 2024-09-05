VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form Notas 
   BackColor       =   &H00C2B000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota"
   ClientHeight    =   5490
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
   Icon            =   "Notas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8190
   Begin AACombo99.AACombo cLocal 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
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
      Height          =   270
      Left            =   1080
      MaxLength       =   70
      TabIndex        =   8
      Top             =   4860
      Width           =   6855
   End
   Begin MSComctlLib.ListView lvArticulo 
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cant."
         Object.Width           =   671
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Artículo"
         Object.Width           =   3212
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Compro"
         Object.Width           =   899
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "A Retirar"
         Object.Width           =   987
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "A Enviar"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Remito"
         Object.Width           =   865
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Devueltos"
         Object.Width           =   1093
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Unitario"
         Object.Width           =   1147
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   5235
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
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox tSerie 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   317
      Left            =   1690
      MaxLength       =   1
      TabIndex        =   3
      Top             =   840
      Width           =   242
   End
   Begin vsViewLib.vsPrinter vsFicha 
      Height          =   1515
      Left            =   1140
      TabIndex        =   32
      Top             =   3480
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
      Left            =   0
      Top             =   0
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
            Picture         =   "Notas.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Notas.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Notas.frx":086E
            Key             =   "Remito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Notas.frx":0B88
            Key             =   "Envio"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Notas.frx":0EA2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label labDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   30
      Top             =   2040
      Width           =   6015
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Sucursal"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "&Lista"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentario:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DETALLE"
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
      TabIndex        =   29
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label labNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label labDocumento 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "21.025996.0012"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label labDato1 
      BackStyle       =   0  'Transparent
      Caption         =   "R.U.C.:"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   7935
   End
   Begin VB.Label labIVA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1,252,200.00"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6600
      TabIndex        =   23
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "I.V.A.:"
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label labImporteNota 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1,252,252.00"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Importe:"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label labArticulo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "200"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Devoluciones:"
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe Descontado"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label labImporteDescontado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5280
      TabIndex        =   16
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label labFechaDocumento 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10-Dic-1998"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2760
      TabIndex        =   14
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Emisión"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label labImporteTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3840
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe Factura"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Número"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label labFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10-Dic-1998"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6960
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   7935
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
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "S&alir del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Notas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tag Utilizados.-----*******--------------------------------------
'tnumero                             Guardo el código del documento.
'labFechaDocumento        Guardo ID de cliente.
'labImporteTotal                  Guardo Fecha de Modificación del documento.
'labImporteDescontado      Guardo el código de moneda utilizada.
'------------------------*******-------------------------------------------

'Modificación
Option Explicit

Private Type tDevolucion
    idArt As Long
    idDev As Long
    Cant As Integer
End Type

Dim CodDocumentoEnvio As Long        'Esta la utilizo solo cuando la factura solo paga artículos de fletes.
Private Rs As rdoResultset
Private iDocumento As Integer       'Propiedad que indica el tipo de documento.
Private arrDevolucion() As tDevolucion

Public Property Get pDocumento() As Integer
    pDocumento = iDocumento
End Property
Public Property Let pDocumento(iParametro As Integer)
    iDocumento = iParametro
End Property
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
Private Sub LimpioCampos()
    CodDocumentoEnvio = 0
    labNombre.Caption = vbNullString
    labDireccion.Caption = vbNullString
    labDocumento.Caption = vbNullString
    tNumero.Tag = vbNullString
    lvArticulo.ListItems.Clear
    labFechaDocumento.Caption = vbNullString
    labFechaDocumento.Tag = vbNullString
    labImporteTotal.Caption = vbNullString
    labImporteDescontado.Caption = vbNullString
    labArticulo.Caption = vbNullString
    labImporteNota.Caption = vbNullString
    labIVA.Caption = vbNullString
    tComentario.Text = vbNullString
End Sub
Private Sub DeshabilitoCampos()
    
    MnuEmitir.Enabled = False: Toolbar1.Buttons("imprimir").Enabled = False
    MnuEnvio.Enabled = False: Toolbar1.Buttons("envio").Enabled = False
    MnuRemito.Enabled = False: Toolbar1.Buttons("remito").Enabled = False
    lvArticulo.Enabled = False
    tComentario.Enabled = False: tComentario.BackColor = vbButtonFace

End Sub
Private Sub HabilitoCampos()

    lvArticulo.Enabled = True
    tComentario.Enabled = True: tComentario.BackColor = vbWindowBackground
    
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = 6150
    
    CodDocumentoEnvio = 0
    CargoLocales
    BuscoCodigoEnCombo cLocal, paCodigoDeSucursal
    SetearLView lvValores.Grilla Or lvValores.FullRow, lvArticulo
    Me.Caption = "Nota de Devolución"
    Me.BackColor = &HC2B000
    
    LimpioCampos
    DeshabilitoCampos
    labFecha.Caption = Format(Date, "d-Mmm-yyyy")
    
    vsFicha.Device = paIConformeN
    vsFicha.PaperBin = paIConformeB
    vsFicha.PaperSize = 1
    vsFicha.PaperHeight = vsFicha.PageHeight / 2
    
    Exit Sub
    
ErrLoad:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio el sgte. error: " & Trim(Err.Description)
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.SimpleText = vbNullString
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Forms(Forms.Count - 2).SetFocus
End Sub
Private Sub Label1_Click()
    Foco tSerie
End Sub
Private Sub Label8_Click()
    Foco tComentario
End Sub
Private Sub lvArticulo_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.SimpleText = " Seleccione un artículo e indique si lo devuelve ('S', 'N'), modifique la cantidad ('+', '-')."
End Sub
Private Sub MnuEmitir_Click()
    AccionImprimir
End Sub

Private Sub MnuEnvio_Click()
    FormEnvio
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
Private Sub tComentario_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.SimpleText = " Ingrese un comentario para la nota."
End Sub
Private Sub tNumero_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)

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
                MsgBox "El formato no es numérico.", vbExclamation, "ATENCIÓN": Foco tNumero
            End If
        End If
    End If

End Sub

Private Sub tNumero_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.SimpleText = " Ingrese el número del documento."
End Sub

Private Sub BuscoFactura()
'Para buscar un documento se considera la propiedad iDocumento, la cual indica el tipo de documento.

On Error GoTo ErrBF

    Screen.MousePointer = vbHourglass
    
    Cons = "Select DocAnulado, DocCliente, DocFecha,  DocFModificacion, DocMoneda, DocTotal, DocCodigo, Renglon.*, MonSigno" _
        & " From Documento, Renglon, Moneda" _
        & " Where DocTipo = " & iDocumento _
        & " And DocSerie = '" & tSerie.Text & "' And DocNumero = " & tNumero.Text _
        & " And DocSucursal = " & cLocal.ItemData(cLocal.ListIndex) _
        & " And DocCodigo = RenDocumento And DocMoneda = MonCodigo"
        
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
            CargoCliente
            'Importe total del documento.
            labImporteTotal.Caption = Trim(RsAux!MonSigno) & " " & Format(RsAux!DocTotal, "#,##0.00")
            labImporteTotal.Tag = RsAux!DocFModificacion
            labImporteDescontado.Tag = RsAux!DocMoneda
            CargoArticulos
            RsAux.Close
            'Busco si tiene nota de devolución.------------
            BuscoOtrasNotas
'            For Each itmX In lvArticulo.ListItems
'                If CLng(itmX.Text) > 0 Then
'                    MnuEmitir.Enabled = True: Toolbar1.Buttons("imprimir").Enabled = True
'                    HabilitoCampos
'                    Exit For
'                End If
'            Next
            For Each itmX In lvArticulo.ListItems
                If CCur(itmX.SubItems(7)) < 0 Then
                    MsgBox "Existe un importe negativo, no se podrá emitir nota parcial.", vbInformation, "ATENCIÓN"
                    Exit For
                End If
            Next
            RecalculoTotales
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
    
    'RSAUX resultset con todos los artículos que tiene el documento.
    Do While Not RsAux.EOF
    
        'Levanto los datos del artículo.
        Cons = "Select ArtNombre, IvaPorcentaje, ArtTipo" _
            & " From Articulo, ArticuloFacturacion, TipoIva" _
            & " Where ArtID = " & RsAux!RenArticulo & " And ArtID = AFaArticulo And AFaIVA = IVaCodigo"
            
        Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        If Not Rs.EOF Then
            
            'Veo si este artículo es de flete.
            Cons = "Select TFlArticulo From TipoFlete Where TFlArticulo = " & RsAux!RenArticulo
            Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                
            If RsLocal.EOF And RsAux!RenArticulo <> paArticuloPisoAgencia Then
                RsLocal.Close
                
                'Veo si es del tipo servicio.
                
                If Rs!ArtTipo = paTipoArticuloServicio Then
                    Set itmX = lvArticulo.ListItems.Add(, "S" & RsAux!RenArticulo, "")
                Else
                    Set itmX = lvArticulo.ListItems.Add(, "A" & RsAux!RenArticulo, "")
                End If
                itmX.Tag = Rs!IVaPorcentaje
                itmX.SubItems(1) = Trim(Rs!ArtNombre)
                itmX.SubItems(2) = RsAux!RenCantidad        'Cantidad total en la factura.
                itmX.SubItems(3) = RsAux!RenARetirar
                itmX.SubItems(7) = Format(RsAux!RenPrecio, "#,##0.00")
                
                Cons = "Select SUM(RReAEntregar) From Remito, RenglonRemito Where RemDocumento = " & RsAux!DocCodigo _
                    & " And RReArticulo = " & RsAux!RenArticulo _
                    & " And RemCodigo = RReRemito"
                    
                Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                
                If Not IsNull(RsLocal(0)) Then
                    itmX.SubItems(5) = RsLocal(0)
                    MnuRemito.Enabled = True: Toolbar1.Buttons("remito").Enabled = True
                Else
                    itmX.SubItems(5) = 0
                End If
                RsLocal.Close
                
                'Sumo la cantidad de artículos que están para envío.
                Cons = "Select SUM(REvAEntregar) From Envio, RenglonEnvio Where EnvDocumento = " & RsAux!DocCodigo _
                    & " And REvArticulo = " & RsAux!RenArticulo & " And EnvCodigo = REvEnvio"
                
                Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not IsNull(RsLocal(0)) Then
                    itmX.SubItems(4) = RsLocal(0)
                    MnuEnvio.Enabled = True: Toolbar1.Buttons("envio").Enabled = True
                    CodDocumentoEnvio = 0       'Tiene envío si va a este va con el documento.
                Else
                    itmX.SubItems(4) = 0
                End If
                RsLocal.Close
    
                Cons = "Select Sum(RenCantidad) From Nota, Documento, Renglon " _
                    & " Where NotFactura = " & tNumero.Tag _
                    & " And DocTipo IN (" & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ") And RenArticulo = " & RsAux!RenArticulo _
                    & " And NotNota = DocCodigo And DocCodigo = RenDocumento And DocAnulado = 0"
            
                Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                
                If Not IsNull(RsLocal(0)) Then itmX.SubItems(6) = RsLocal(0) Else itmX.SubItems(6) = 0
                RsLocal.Close
                
                itmX.Text = CInt(itmX.SubItems(2)) - (CInt(itmX.SubItems(4)) + CInt(itmX.SubItems(5)) + CInt(itmX.SubItems(6)))
                
            Else
                
                'El artículo es de Flete.
                RsLocal.Close
                
                Set itmX = lvArticulo.ListItems.Add(, "F" & RsAux!RenArticulo, "")
                itmX.Tag = Rs!IVaPorcentaje
                itmX.SubItems(1) = Trim(Rs!ArtNombre)
                itmX.SubItems(2) = RsAux!RenCantidad
                itmX.SubItems(3) = 0    'A retirar
                itmX.SubItems(5) = 0    'Remito
                itmX.SubItems(7) = Format(RsAux!RenPrecio, "#,##0.00")
                itmX.SubItems(4) = 0
                
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
                If Not IsNull(RsLocal(0)) Then itmX.SubItems(6) = RsLocal(0) Else itmX.SubItems(6) = 0
                RsLocal.Close
                
                itmX.Text = 0
            End If
        End If
        Rs.Close
        
        RsAux.MoveNext
    Loop

End Sub

Private Sub BuscoOtrasNotas()

    Cons = "Select Sum(DocTotal) From Nota, Documento " _
        & " Where NotFactura = " & tNumero.Tag _
        & " And DocTipo IN (" & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")" _
        & " And DocAnulado = 0 And NotNota = DocCodigo"
        
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    labImporteDescontado.Caption = "0.00"
    
    If Not Rs.EOF Then
        If Not IsNull(Rs(0)) Then
            labImporteDescontado.Caption = Format(CCur(labImporteDescontado.Caption) + Rs(0), "#,##0.00")
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
        End If
    End If

End Sub

Private Sub lvArticulo_BeforeLabelEdit(Cancel As Integer)
    
    If UCase(lvArticulo.SelectedItem.SubItems(4)) = "NO" Then MsgBox "No se puede modificar si no se marca para hacer nota."

End Sub

Private Sub lvArticulo_GotFocus()

    Status.SimpleText = " Seleccione un artículo e indique si lo devuelve ('S', 'N'), modifique la cantidad ('+', '-')."
    
End Sub

Private Sub lvArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    If lvArticulo.ListItems.Count > 0 Then
        Select Case KeyCode
            Case vbKeyReturn
                tComentario.SetFocus
                
            Case vbKeyAdd
                If CLng(lvArticulo.SelectedItem.Text) < CInt(lvArticulo.SelectedItem.SubItems(2)) - (CInt(lvArticulo.SelectedItem.SubItems(4)) + CInt(lvArticulo.SelectedItem.SubItems(5)) + CInt(lvArticulo.SelectedItem.SubItems(6))) Then
                    lvArticulo.SelectedItem.Text = CLng(lvArticulo.SelectedItem.Text) + 1
                    RecalculoTotales
                End If
                
            
            Case vbKeySubtract
                If CLng(lvArticulo.SelectedItem.Text) > 0 Then
                    lvArticulo.SelectedItem.Text = CLng(lvArticulo.SelectedItem.Text) - 1
                    RecalculoTotales
                End If
        End Select
        
    End If

End Sub

Private Sub lvArticulo_LostFocus()

    Status.SimpleText = vbNullString

End Sub

Private Sub AccionImprimir()
Dim Msg As String
Dim NroDoc As String        'Nro. de nota de devolución.
Dim lnDocumento As Long, lEnvioDev As Long
Dim sPiso As Boolean
Dim aUsuario As Long, strDefensa As String, sImprimoRetiro As Boolean
Dim rsSer As rdoResultset, IdServicio As Long
Dim cCofis As Currency, cNeto As Currency
Dim iPosArr As Integer, iResto As Integer
    
    If Trim(labImporteNota.Caption) = vbNullString Then
        MsgBox "No hay artículos seleccionados para devolver.", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else
        iPosArr = 0
        For Each itmX In lvArticulo.ListItems
            If Val(itmX.Text) > 0 Then
                iPosArr = 1
                Exit For
            End If
        Next
        If iPosArr = 0 Then
            MsgBox "No hay artículos seleccionados para devolver.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
    End If
    iPosArr = 0
      
    'Consulto antes si es de servicio si es no controlo las devoluciones
    IdServicio = 0
    Cons = "Select * From Servicio Where SerDocumento = " & tNumero.Tag
    Set rsSer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsSer.EOF Then IdServicio = rsSer!SerCodigo
    rsSer.Close
    
    lEnvioDev = 0
    If IdServicio = 0 Then
        BuscoIDDev
        'Consulto si hay devoluciones a imprimir --> Pregunto id de envío
        
        For Each itmX In lvArticulo.ListItems
            
            If CLng(itmX.Text) > 0 And Mid(itmX.Key, 1, 1) = "A" _
                And (Not (Val(itmX.Text) = Val(itmX.SubItems(2)) And Val(itmX.SubItems(2)) = Val(itmX.SubItems(3)))) Then
                
                iPosArr = 0
                'Busco en el array de devoluciones el id del artículo.
                iPosArr = BuscoPosArray(Mid(itmX.Key, 2, Len(itmX.Key)))
                If iPosArr > 0 Then
                    'Si no carga todo pido aca.
                    If arrDevolucion(iPosArr).Cant < Val(itmX.Text) Then
                        'Pido ID de envío.
                        lEnvioDev = Val(InputBox("Se van a emitir fichas de devolución." & vbCrLf & "Ingrese el código del envío que se asociara a la devolución.", "Envío asignado a la Devolución"))
                        If lEnvioDev = 0 Then
                            MsgBox "Recuerde que debe cumplir la devolución o asignarle un envío a la brevedad.", vbInformation, "ATENCIÓN"
                        Else
                            'Válido el envío.
                            lEnvioDev = ValidoEnvioDevolucion(lEnvioDev)
                            If lEnvioDev = 0 Then Screen.MousePointer = 0: Exit Sub
                        End If
                        Exit For
                    End If
                Else
                    'Pido ID de envío.
                    lEnvioDev = Val(InputBox("Se van a emitir fichas de devolución." & vbCrLf & "Ingrese el código del envío que se asociara a la devolución.", "Envío asignado a la Devolución"))
                    If lEnvioDev = 0 Then
                        MsgBox "Recuerde que debe cumplir la devolución o asignarle un envío a la brevedad.", vbInformation, "ATENCIÓN"
                    Else
                        'Válido el envío.
                        lEnvioDev = ValidoEnvioDevolucion(lEnvioDev)
                        If lEnvioDev = 0 Then Screen.MousePointer = 0: Exit Sub
                    End If
                    Exit For
                End If
            End If
        Next
    End If
    
    
    If MsgBox("¿Desea emitir la nota?", vbQuestion + vbYesNo, "EMITIR") = vbNo Then Exit Sub
    
    Dim sNegativo As Boolean
    sNegativo = False
    For Each itmX In lvArticulo.ListItems
        If CCur(itmX.SubItems(7)) < 0 Then sNegativo = True
    Next
    
    If sNegativo = True Then
        For Each itmX In lvArticulo.ListItems
            If CInt(itmX.SubItems(2)) <> CInt(itmX.Text) Then
                If CInt(itmX.SubItems(2)) = CInt(itmX.SubItems(6)) Then
                    'Tiene Nota, puede ser la nota del envío.
                    'Consulto si el artículo es el que paga el flete.
                    Cons = "Select TFlArticulo From TipoFlete Where TFlArticulo = " & Mid(itmX.Key, 2, Len(itmX.Key))
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If RsAux.EOF Then MsgBox "Esta factura posee valores negativos no podrá emitir la nota parcial.", vbExclamation, "ATENCIÓN": RsAux.Close: Exit Sub
                    RsAux.Close
                Else
                    MsgBox "Esta factura posee valores negativos no podrá emitir la nota parcial.", vbExclamation, "ATENCIÓN": Exit Sub
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
    
    On Error GoTo ErrAI
    cBase.BeginTrans
    On Error GoTo ErrResumo
    sImprimoRetiro = False
    Screen.MousePointer = vbHourglass
    
    'Si la factura pago un servicio, no hago movimiento de stock físico.
    If IdServicio > 0 Then
        Cons = "Select * From Servicio Where SerCodigo = " & IdServicio
        Set rsSer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsSer.EOF Then
            rsSer.Edit
            rsSer!SerModificacion = Format(gFechaServidor, FormatoFH)
            rsSer!SerDocumento = Null
            rsSer.Update
        End If
        rsSer.Close
        
        'Veo si la factura esta en Pendientes.
        Cons = "Select * From DocumentoPendiente Where DPeDocumento = " & Val(tNumero.Tag)
        Set rsSer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsSer.EOF Then
            rsSer.Delete
        End If
        rsSer.Close
        
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
            
            'Veo si tiene aplicado el cofis.
            'No me importa si el cliente cumple con la condición si la factura tenía cofis.
            If Not IsNull(RsAux!DocCofis) Then labDocumento.Tag = "1"

            'Updateo el documento, le cambio la fecha de modif.
            RsAux.Edit
            RsAux!DocFModificacion = Format(gFechaServidor, FormatoFH)
            RsAux.Update
            RsAux.Close
            
            NroDoc = NumeroDocumento(paDNDevolucion)
            
            Cons = "INSERT INTO Documento " _
                & " (DocFecha, DocTipo, DocSerie, DocNumero, DocCliente, DocMoneda, DocTotal, DocIva, DocAnulado, DocSucursal, DocUsuario, DocFModificacion, DocComentario)" _
                & " Values ('" & Format(gFechaServidor, FormatoFH) & "'" _
                & ", " & TipoDocumento.NotaDevolucion _
                & ", '" & Mid(NroDoc, 1, 1) & "', " & Mid(NroDoc, 2, Len(NroDoc)) _
                & ", " & labFechaDocumento.Tag & ", " & labImporteDescontado.Tag _
                & ", " & CCur(labImporteNota.Caption) & ", " & CCur(labIVA.Caption) _
                & ", 0," & paCodigoDeSucursal & ", " & aUsuario _
                & ", '" & Format(gFechaServidor, FormatoFH) & "'"
                
            If Trim(tComentario.Text) = vbNullString Then
                Cons = Cons & ", Null)"
            Else
                Cons = Cons & ", '" & tComentario.Text & "')"
            End If
            cBase.Execute (Cons)

            Cons = "SELECT MAX(DocCodigo) From Documento" _
                & " WHERE DocTipo = " & TipoDocumento.NotaDevolucion _
                & " AND DocSerie = '" & Mid(NroDoc, 1, 1) & "' AND DocNumero = " & Mid(NroDoc, 2, Len(NroDoc))
        
            Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            lnDocumento = Rs(0)
            Rs.Close
            
            For Each itmX In lvArticulo.ListItems
            
                If CLng(itmX.Text) > 0 Then
                    'Válido cantidades
                    If IdServicio = 0 Then
                        
                        'Tengo artículos en la factura.
                        If Mid(itmX.Key, 1, 1) = "A" Then
                        
                            iPosArr = 0
                            'Veo si tengo ficha de devolución asignada, sino hago una nueva si corresponde.
                            'Si la cantidad = a la inicial no hago nada.
                            If Not (Val(itmX.Text) = Val(itmX.SubItems(2)) And Val(itmX.SubItems(2)) = Val(itmX.SubItems(3))) Then
                                                        
                                'Busco en el array de devoluciones el id del artículo.
                                iPosArr = BuscoPosArray(Mid(itmX.Key, 2, Len(itmX.Key)))
                            
                                If iPosArr > 0 Then
                                    'Le asigno todo lo posible a la devolución.
                                    
                                    If arrDevolucion(iPosArr).Cant <= Val(itmX.Text) Then
                                        Cons = "Select * From Devolucion Where DevID = " & arrDevolucion(iPosArr).idDev
                                        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                                        RsAux.Edit
                                        RsAux!DevNota = lnDocumento
                                        RsAux.Update
                                        RsAux.Close
                                    End If
                                    iResto = Val(itmX.Text) - arrDevolucion(iPosArr).Cant
                                    
                                    If iResto <> 0 Then
                                        'Lo que queda por devolver es mayor a lo que tiene para retirar.
                                        If iResto > Val(itmX.SubItems(3)) Then
                                            'Creo las fichas de dev. para la diferencia.
                                            Cons = "Select * From Devolucion Where DevID = 0"
                                            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                                            RsAux.AddNew
                                            RsAux!DevFactura = Val(tNumero.Tag)
                                            RsAux!DevCliente = labFechaDocumento.Tag
                                            RsAux!DevNota = lnDocumento
                                            RsAux!DevArticulo = Mid(itmX.Key, 2, Len(itmX.Key))
                                            RsAux!DevCantidad = iResto - Val(itmX.SubItems(3))
                                            If lEnvioDev > 0 Then RsAux!DevEnvio = lEnvioDev
                                            RsAux.Update
                                            RsAux.Close
                                            iResto = Val(itmX.SubItems(3))
                                            sImprimoRetiro = True
                                        End If
                                    End If  'Se dev. todo.
                                Else    'No hay Dev.
                                    If Val(itmX.Text) > Val(itmX.SubItems(3)) Then
                                        'Lo que devuelve es mayor a lo que tiene para retirar.
                                        Cons = "Select * From Devolucion Where DevID = 0"
                                        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                                        RsAux.AddNew
                                        RsAux!DevFactura = Val(tNumero.Tag)
                                        RsAux!DevCliente = labFechaDocumento.Tag
                                        RsAux!DevNota = lnDocumento
                                        RsAux!DevArticulo = Mid(itmX.Key, 2, Len(itmX.Key))
                                        RsAux!DevCantidad = Val(itmX.Text) - Val(itmX.SubItems(3))
                                        If lEnvioDev > 0 Then RsAux!DevEnvio = lEnvioDev
                                        RsAux.Update
                                        RsAux.Close
                                        sImprimoRetiro = True
                                        iResto = Val(itmX.SubItems(3))
                                    ElseIf Val(itmX.Text) <= Val(itmX.SubItems(3)) Then
                                        iResto = Val(itmX.Text)
                                    End If
                                End If
                            Else
                                'No hay necesidad de hacer devolución.
                                iResto = Val(itmX.Text)
                            End If  'No hay dev.
                            
                            If iResto > 0 Then
                                MarcoStockXDevolucion CLng(Mid(itmX.Key, 2, Len(itmX.Key))), CCur(iResto), CCur(iResto), TipoLocal.Deposito, paCodigoDeSucursal, aUsuario, TipoDocumento.NotaDevolucion, lnDocumento
                            
                                Cons = "Select * From Renglon Where RenDocumento = " & tNumero.Tag _
                                    & " And RenArticulo = " & Mid(itmX.Key, 2, Len(itmX.Key))
                                
                                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                                RsAux.Edit
                                RsAux!RenARetirar = RsAux!RenARetirar - iResto
                                RsAux.Update
                                RsAux.Close
                            End If
                        End If
                    End If
                    
                    '-----------------------------------------------------------------------------------------
                    'Inserto los renglones de la NOTA
                    Cons = "INSERT INTO Renglon (RenDocumento, RenArticulo, RenCantidad, RenPrecio, RenIVA, RenARetirar, RenCofis)" _
                        & " VALUES (" & lnDocumento & ", " & Mid(itmX.Key, 2, Len(itmX.Key)) _
                        & ", " & itmX.Text & ", " & CCur(itmX.SubItems(7)) _
                        & ", " & Format(CCur(itmX.SubItems(7)) - (CCur(itmX.SubItems(7)) / CCur(1 + (CCur(itmX.Tag) / 100))), "###0.000")
                        
                    If CCur(itmX.Text) <= CCur(itmX.SubItems(3)) Then
                        Cons = Cons & ", " & CCur(itmX.Text)
                    Else
                        Cons = Cons & ", " & CCur(itmX.SubItems(3))
                    End If
                    
                    'NETO DEL COFIS----------------------------------------------------------------
                    If Val(labDocumento.Tag) = 1 Then
                        cNeto = CCur(itmX.SubItems(7)) / CCur(1 + (CCur(itmX.Tag) / 100))
                        'Tengo el neto tengo que sacarle el cofis.
                        cNeto = cNeto - (cNeto / (1 + (paCofis / 100)))
                        cNeto = Format(cNeto, "###0.00")
                        cCofis = cCofis + (cNeto * Val(itmX.Text))
                        Cons = Cons & ", " & Format(cNeto, "###0.00") & ")"
                    Else
                        Cons = Cons & ", Null)"
                    End If
                    cBase.Execute (Cons)
                    '-----------------------------------------------------------------------------------------
                    
                    
                    'Si el artículo facturo una diferencia de Envío voy a eliminar la misma y corrijo el valor del envío
                    If Mid(itmX.Key, 2, Len(itmX.Key)) = paArticuloDiferenciaEnvio Then
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
            
            If labDocumento.Tag = "1" Then
                Cons = "Update Documento Set DocCofis = " & Format(cCofis, "###0.00") & " Where DocCodigo = " & lnDocumento
                cBase.Execute (Cons)
            End If
            
            'INSERTO RELACION NOTA
            Cons = "INSERT INTO Nota (NotFactura, NotNota, NotDevuelve, NotSalidaCaja)" _
                & " Values (" & tNumero.Tag & "," & lnDocumento _
                & ", " & CCur(labImporteNota.Caption) _
                & ", " & CCur(labImporteNota.Caption) _
                & ")"
                
            cBase.Execute (Cons)
            '------------------------------------------------------------
            
            clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.Notas, paCodigoDeTerminal, aUsuario, lnDocumento, _
                                   Descripcion:="Nota de Devolución: " & Mid(NroDoc, 1, 1) & " " & Mid(NroDoc, 2, Len(NroDoc)), Defensa:=Trim(strDefensa)
                                   
            cBase.CommitTrans
            On Error GoTo ErrAIF
            ImprimoNota (lnDocumento)
            ImprimoSalidaCaja lnDocumento, CInt(aUsuario)
            If sImprimoRetiro Then ImprimoRetirosPorDevolucion CLng(labFechaDocumento.Tag), tNumero.Tag, lnDocumento
        Else
            On Error GoTo ErrAIF
            RsAux.Close
            Msg = "Otra terminal modificó el documento, no podrá realizar la nota."
            GoTo ErrResumo
        End If
    End If
    
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

End Sub
Private Sub RecalculoTotales()
Dim bHabilito As Boolean
    labImporteNota.Caption = "0.00"
    labIVA.Caption = "0.00"
    labArticulo.Caption = "0"
    bHabilito = False
    For Each itmX In lvArticulo.ListItems
        If CLng(itmX.Text) > 0 Then
            labImporteNota.Caption = CCur(labImporteNota.Caption) + (CLng(itmX.Text) * CCur(itmX.SubItems(7)))
            labIVA.Caption = CCur(labIVA.Caption) + ((CLng(itmX.Text) * CCur(itmX.SubItems(7))) - ((CLng(itmX.Text) * CCur(itmX.SubItems(7))) / ((CCur(itmX.Tag) / 100) + 1)))
            labArticulo.Caption = CLng(labArticulo.Caption) + CLng(itmX.Text)
            bHabilito = True
        End If
    Next
    labImporteNota.Caption = Format(labImporteNota.Caption, FormatoMonedaP)
    labIVA.Caption = Format(labIVA.Caption, FormatoMonedaP)
    
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

Private Sub CargoCliente()

    Cons = "Select CliCiRuc, CliTipo, CliDireccion, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2), Ruc = (CPeRuc), Estatal = 0 " _
       & " From Cliente, CPersona " _
       & " Where CliCodigo = " & RsAux!DocCliente _
       & " And CliCodigo = CPeCliente " _
                                            & " UNION " _
       & " Select CliCiRuc, CliTipo, CliDireccion, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')') , Ruc = (Null), Estatal = CEmEstatal " _
       & " From Cliente, CEmpresa " _
       & " Where CliCodigo = " & RsAux!DocCliente _
       & " And CliCodigo = CEmCliente"
    
'    Cons = "Select CliTipo, CliCIRuc, CliDireccion From Cliente Where CliCodigo = " & RsAux!DocCliente
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Rs.EOF Then
        Rs.Close
        Screen.MousePointer = vbDefault
        MsgBox "No se encontro la información del cliente.", vbExclamation, "ATENCIÓN"
        labDocumento.Tag = ""
    Else
        labDocumento.Tag = 0
        labNombre.Caption = " " & Trim(Rs!Nombre)
        If Not IsNull(Rs!CliDireccion) Then labDireccion.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, Rs!CliDireccion) Else labDireccion.Caption = vbNullString
        If Rs!CliTipo = TipoCliente.Empresa Then
            labDato1.Caption = "R.U.C.:"
            If Not IsNull(Rs!CliCIRuc) Then labDocumento = Trim(Rs!CliCIRuc)
            labDocumento.Tag = "1"
        Else
            'labDato1.Caption = "C.I.:"
            If Not IsNull(Rs!CliCIRuc) Then labNombre.Caption = Trim(labNombre.Caption) & " (" & clsGeneral.RetornoFormatoCedula(Rs!CliCIRuc) & ")"
            If Not IsNull(Rs!RUC) Then
                labDocumento.Caption = Trim(Rs!RUC)
                labDocumento.Tag = 1
            End If
        End If
        Rs.Close
    End If
    
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
Dim Result As Integer, JobSRep1 As Integer, JobSRep2 As Integer, jobnum As Integer
Dim NombreFormula As String, CantForm As Integer, aTexto As String

    Screen.MousePointer = 11
    'Inicializo el Reporte y SubReportes
    jobnum = crAbroReporte(gPathListados & "NotaDevolucion.RPT")
    If jobnum = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
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
            Case "nombredocumento": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDNDevolucion & "'")
            Case "cliente"
                'If labDato1.Caption = "C.I.:" And Trim(labDocumento.Caption) <> vbNullString Then aTexto = "(" & Trim(labDocumento.Caption) & ")"
                Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labNombre.Caption) & "'")
            Case "direccion": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labDireccion.Caption) & "'")
            Case "ruc"
                If labDato1.Caption = "R.U.C.:" And Trim(labDocumento.Caption) <> vbNullString Then
                    aTexto = clsGeneral.RetornoFormatoRuc(labDocumento.Caption)
                Else
                    aTexto = vbNullString
                End If
                Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(aTexto) & "'")
            
            Case "codigobarras":
                    Result = crSeteoFormula(jobnum%, NombreFormula, "''")
                    'Result = crSeteoFormula(JobNum%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.NotaDevolucion, Documento) & "'")
                    
            Case "signomoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & BuscoSignoMoneda(labImporteDescontado.Tag) & "'")
            Case "nombremoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'(" & BuscoNombreMoneda(labImporteDescontado.Tag) & ")'")
            Case "textoretira"
                'Detallamos el documento al cual se le hace la nota.
                aTexto = "'" & Trim(cLocal.Text) & " " & Trim(tSerie.Text) & " " & Trim(tNumero.Text) & "'"
                Result = crSeteoFormula(jobnum%, NombreFormula, aTexto)
            Case Else: Result = 1
        End Select
        If Result = 0 Then GoTo ErrCrystal
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

Dim NombreFormula As String, Result As Integer
Dim JobNumMC As Integer, CantFormMC As Integer

    'Inicializa el Engine del Crystal y setea la impresora para el JOB
    On Error GoTo ErrCrystal
    
    'Inicializo el Reporte y SubReportes
    JobNumMC = crAbroReporte(gPathListados & "MovimientoNota.RPT")
    If JobNumMC = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
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
            Case "entradasalida": Result = crSeteoFormula(JobNumMC%, NombreFormula, "'SALIDA DE CAJA'")
            Case "sucursal": Result = crSeteoFormula(JobNumMC%, NombreFormula, "'Sucursal: " & BuscoNombreSucursal(paCodigoDeSucursal) & "'")
            Case "comentario": Result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & Trim(tComentario.Text) & "'")
            Case "importe": Result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & Format(RsAux!DocTotal, FormatoMonedaP) & "'")
            Case "tipo"
                aTexto = "FACTURA " & Trim(tSerie.Text) & " " & Trim(tNumero.Text)
                If Not RsAux.EOF Then aTexto = "N. DEVOLUCIÓN " & RsAux!DocSerie & RsAux!Docnumero & " sobre " & aTexto
                Result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & aTexto & "'")
                
            Case "moneda": Result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & BuscoSignoMoneda(labImporteDescontado.Tag) & "'")
            Case "usuario": Result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & BuscoInicialUsuario(Usuario) & "'")
            Case Else: Result = 1
        End Select
        If Result = 0 Then GoTo ErrCrystal
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

Private Sub BuscoIDDev()
'Busco si el producto tiene ficha de devolución ingresada.
On Error GoTo ErrVN
Dim bXBusqueda As Boolean, idDev As Long, iCant As Long

    ReDim arrDevolucion(0)      'Inicializo array de devolución.
    For Each itmX In lvArticulo.ListItems
    
        If Mid(itmX.Key, 1, 1) = "A" Then
            
            'Si devuelve.
            If Val(itmX.Text) > 0 Then
            
                'Si lo que compro es igual a lo que tiene para retirar y es lo que devuelve.
                If Not (Val(itmX.Text) = Val(itmX.SubItems(2)) And Val(itmX.SubItems(2)) = Val(itmX.SubItems(3))) Then
            
                    idDev = 0
                    Cons = "Select * From Devolucion Where DevFactura = " & Val(tNumero.Tag) _
                        & " And DevArticulo = " & Mid(itmX.Key, 2, Len(itmX.Key)) _
                        & " And DevNota IS Null And DevLocal Is Not Null And DevFAltaLocal Is Not Null" _
                        & " And DevAnulada Is Null"
            
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                    If Not RsAux.EOF Then
                        'Válido la cantidad de la dev. con lo que devuelve.
                        If RsAux!DevCantidad > Val(itmX.Text) Then
                            MsgBox "Existe un devolución ingresada pero posee más articulos de los que se quiere devolver.", vbCritical, "ATENCIÓN"
                            bXBusqueda = True
                        Else
                            ReDim Preserve arrDevolucion(UBound(arrDevolucion) + 1)
                            arrDevolucion(UBound(arrDevolucion)).idArt = Mid(itmX.Key, 2, Len(itmX.Key))
                            arrDevolucion(UBound(arrDevolucion)).idDev = RsAux!DevID
                            arrDevolucion(UBound(arrDevolucion)).Cant = RsAux!DevCantidad
                            idDev = RsAux!DevID
                            bXBusqueda = False
                        End If
                    Else
                        bXBusqueda = True
                    End If
                    RsAux.Close
                    
                    If bXBusqueda Then
                    
                        If MsgBox("¿El cliente tiene una Ficha de Devolución para el artículo " & UCase(Trim(itmX.SubItems(1))) & " ?", vbQuestion + vbYesNo, "CLIENTE TIENE FICHA?") = vbYes Then
                        
                            idDev = 0: iCant = 0
                            '2da busqueda.
                            'Busco para el artículo y el cliente.
                            Cons = "Select * From Devolucion Where DevArticulo = " & Mid(itmX.Key, 2, Len(itmX.Key)) _
                                & " And DevCliente = " & Val(labFechaDocumento.Tag) _
                                & " And DevCantidad <= " & Val(itmX.Text) _
                                & " And DevNota IS Null And DevLocal Is Not Null And DevFAltaLocal Is Not Null" _
                                & " And DevAnulada Is Null"
                        
                            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                            
                            If Not RsAux.EOF Then
                                
                                'Encontre ficha.
                                'Válido que sea única.
                                RsAux.MoveNext
                                If RsAux.EOF Then
                                    
                                    RsAux.MoveFirst
                                    idDev = RsAux!DevID
                                    iCant = RsAux!DevCantidad
                                    RsAux.Close
                                    
                                Else
                                    'Hay + de 1 ficha.
                                    'Abro lista de ayuda para que seleccione.
                                    RsAux.Close
                                    
                                    Cons = "Select DevID, DevID as 'Código', IsNull(DocSerie, 'No') + ' ' + IsNull(Convert(char,DocNumero), 'Hay') as 'Documento', DevFAltaLocal as 'Ingresó', LocNombre as 'Local', DevCantidad as 'Cantidad' " _
                                        & " From Devolucion" _
                                            & " Left Outer Join Documento ON DocCodigo = DevFactura " _
                                        & " , Local " _
                                        & " Where DevArticulo = " & Mid(itmX.Key, 2, Len(itmX.Key)) _
                                        & " And DevCliente = " & Val(labFechaDocumento.Tag) _
                                        & " And DevCantidad <= " & Val(itmX.Text) _
                                        & " And DevNota IS Null And DevLocal = LocCodigo And DevFAltaLocal Is Not Null" _
                                        & " And DevAnulada Is Null"
                                    
                                    Dim objLista As New clsListadeAyuda, mIDSel As Long
                                    mIDSel = objLista.ActivarAyuda(cBase, Cons, 5000, 1)
                                    If mIDSel > 0 Then mIDSel = objLista.RetornoDatoSeleccionado(0)
                                    Set objLista = Nothing
                                    Me.Refresh
                                    
                                    If mIDSel > 0 Then
                                                        
                                        idDev = mIDSel
                                                        
                                        'Válido que si tengo documento no me haya elegido otra devolución que tenga otro documento.
                                        Cons = "Select * From Devolucion Where DevID = " & idDev _
                                            & " And DevArticulo = " & Mid(itmX.Key, 2, Len(itmX.Key)) _
                                            & " And DevNota IS Null And DevLocal Is Not Null And DevFAltaLocal Is Not Null" _
                                            & " And DevAnulada Is Null"
                                        
                                        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                                                        
                                        If Not IsNull(RsAux!DevFactura) Then
                                            If Val(tNumero.Tag) <> RsAux!DevFactura Then
                                                If MsgBox("La ficha que seleccionó esta asociada a otro documento." & vbCrLf _
                                                    & "¿Confirma que es está la devolución para la factura de compra?", vbQuestion + vbYesNo + vbDefaultButton2, "OTRA FACTURA") = vbNo Then
                                                    idDev = 0
                                                End If
                                            End If
                                        ElseIf RsAux!DevCantidad > Val(itmX.Text) Then
                                            MsgBox "La ficha que seleccionó tiene más artículos de los que se quiere devolver, esto no es posible.", vbInformation, "ATENCIÓN"
                                            idDev = 0
                                        Else
                                            iCant = RsAux!DevCantidad
                                        End If
                                        RsAux.Close
                                    End If
                                End If
                            Else
                                RsAux.Close
                            End If
                            
                            If idDev = 0 Then
                                
                                If MsgBox("Aparentemente no existe ingreso." & vbCr & "¿Desea emitir la ficha de devolución para retirar en domicilio?" _
                                    & vbCrLf & "Presione <NO> para ingresar el ID de devolución.", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                                    
                                    Dim sID As String
                                    'Pido el Id de devolucion.
                                    sID = InputBox("Ingrese el número de devolución que desea asignar.", "Número de Devolución")
                                    If IsNumeric(sID) Then
                                        Cons = "Select * From Devolucion Where DevID = " & Val(sID) _
                                            & " And DevArticulo = " & Mid(itmX.Key, 2, Len(itmX.Key)) _
                                            & " And DevNota IS Null And DevLocal Is Not Null And DevFAltaLocal Is Not Null" _
                                            & " And DevAnulada Is Null"
                                        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                                        If RsAux.EOF Then
                                            RsAux.Close
                                            MsgBox "El código de devolución no existe ó no cumple con las condiciones para ser asignado a esta nota." & vbCrLf & vbCrLf & "SE EMITIRÁ LA FICHA DE DEVOLUCIÓN.", vbInformation, "ATENCIÓN"
                                        Else
                                            If RsAux!DevCantidad > Val(itmX.Text) Then
                                                MsgBox "La ficha que seleccionó tiene más artículos de los que se quiere devolver, esto no es posible.", vbInformation, "ATENCIÓN"
                                                idDev = 0
                                            Else
                                                idDev = RsAux!DevID
                                                iCant = RsAux!DevCantidad
                                            End If
                                            RsAux.Close
                                        End If
                                    Else
                                        MsgBox "Se emitirá la Ficha de Devolución.", vbInformation, "ATENCIÓN"
                                    End If
                                End If
                            End If
                            
                            If idDev > 0 Then
                                ReDim Preserve arrDevolucion(UBound(arrDevolucion) + 1)
                                arrDevolucion(UBound(arrDevolucion)).idArt = Mid(itmX.Key, 2, Len(itmX.Key))
                                arrDevolucion(UBound(arrDevolucion)).idDev = idDev
                                arrDevolucion(UBound(arrDevolucion)).Cant = iCant
                            End If
                            
                        End If
                    
                    End If      'Si es por busqueda
                
                End If      'Lo que devuelve es todo lo a retirar.
                
            End If  'Cantidad
            
        End If
        
    Next
    Exit Sub
    
ErrVN:
    clsGeneral.OcurrioError "Ocurrio un error al validar la nota.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValidoNotaOLD() As Boolean
On Error GoTo ErrVN

    ValidoNotaOLD = False
    For Each itmX In lvArticulo.ListItems
        
        If Mid(itmX.Key, 1, 1) = "A" Then
        
            Cons = "Select * From Devolucion Where DevFactura = " & tNumero.Tag _
                & " And DevArticulo = " & Mid(itmX.Key, 2, Len(itmX.Key)) & " And DevNota = Null"
            
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If RsAux.EOF Then
                RsAux.Close
                'Si lo que retira es mayor que lo que tiene para devolver en la factura.
                If Val(itmX.SubItems(3)) < Val(itmX.Text) Then
                    If MsgBox("Para el artículo " & Trim(itmX.SubItems(1)) & " no existe un ingreso de mercadería en depósito." & Chr(vbKeyReturn) _
                    & "Se emitirán fichas de Retiro de Artículos en el domicilio del cliente." & Chr(13) & Chr(10) _
                    & "¿Confirma emitir estas fichas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Function
                End If
            Else
                If Val(labFechaDocumento.Tag) <> RsAux!DevCliente And CLng(itmX.Text) > 0 Then
                'Si la devolución está hecha para otro cliente y quiere devolver este artículo lo hecho.
                    RsAux.Close
                    MsgBox "Existe un alta en depósito para el artículo '" & Trim(itmX.SubItems(1)) & "' pero dicho ingreso se hizo para otro cliente." & Chr(13) & "Debe emitir nota especial.", vbExclamation, "ATENCIÓN"
                    Exit Function
                ElseIf Val(labFechaDocumento.Tag) = RsAux!DevCliente And CLng(itmX.Text) = 0 Then
                'Si hay devolución para el cliente y no devuelve.
                    RsAux.Close
                    MsgBox "Existe un alta en depósito para el artículo '" & Trim(itmX.SubItems(1)) & "' y no se devuelve ningún artículo." & Chr(13) & "Procedimiento Erróneo", vbExclamation, "ATENCIÓN"
                    Exit Function
                'si lo que devuelve aún está en la factura.
                ElseIf Val(itmX.SubItems(3)) >= Val(itmX.Text) And Val(itmX.SubItems(3)) <> 0 Then
                    
                    MsgBox "Existe un alta en depósito para el artículo '" & Trim(itmX.SubItems(1)) & "' x " & RsAux!DevCantidad & ", y se desean devolver artículos que aún están en la factura (" & Val(itmX.SubItems(3)) & ")." & Chr(13) & "Procedimiento Erróneo", vbExclamation, "ATENCIÓN"
                    RsAux.Close
                    Exit Function
                ElseIf RsAux!DevCantidad > Val(itmX.Text) And Val(itmX.Text) > 0 And Val(labFechaDocumento.Tag) <> RsAux!DevCliente Then
                'Si lo que esta devuelto excede lo que quiere devolver lo rajo.
                    RsAux.Close
                    MsgBox "Existe un alta en depósito para el artículo '" & Trim(itmX.SubItems(1)) & "' que excede lo que se quiere devolver." & Chr(13) & "Procedimiento Erróneo", vbExclamation, "ATENCIÓN"
                    Exit Function
                ElseIf Val(itmX.SubItems(3)) + RsAux!DevCantidad > Val(itmX.Text) And Val(labFechaDocumento.Tag) = RsAux!DevCliente Then
                    'Si lo que devolvio + lo que hay en la factura excede la cantidad que devuelve.
                    MsgBox "Existe un alta en depósito para el artículo '" & Trim(itmX.SubItems(1)) & "' x " & RsAux!DevCantidad & ", y se desean devolver artículos que aún están en la factura (" & Val(itmX.SubItems(3)) & ")." & Chr(13) & "Procedimiento Erróneo", vbExclamation, "ATENCIÓN"
                    RsAux.Close: Exit Function
                ElseIf RsAux!DevCantidad < Val(itmX.Text) - Val(itmX.SubItems(3)) And Val(labFechaDocumento.Tag) = RsAux!DevCliente Then
                    'lo que devolvio en el depósito < cant. que devuelve - cant. en la factura
                    If MsgBox("Para el artículo '" & Trim(itmX.SubItems(1)) & "' existe un ingreso de mercadería en depósito por '" & RsAux!DevCantidad & "'." & Chr(13) _
                        & "Se emitirán fichas de Retiro para los artículos restantes que se devuelven." & Chr(13) & Chr(13) _
                        & "¿Confirma emitir estas fichas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then RsAux.Close: Exit Function Else RsAux.Close
                ElseIf RsAux!DevCantidad > Val(itmX.Text) - Val(itmX.SubItems(3)) And Val(labFechaDocumento.Tag) = RsAux!DevCliente Then
                    RsAux.Close
                    MsgBox "El cliente devolvió en el depósito los artículos que tenía retirados, pero aún hay artículos en la factura.", vbExclamation, "ATENCIÓN"
                    Exit Function
                Else
                    RsAux.Close
                End If
            End If
        End If
    Next
    ValidoNotaOLD = True
    Exit Function
ErrVN:
    clsGeneral.OcurrioError "Ocurrio un error al validar la nota.", Err.Description
    Screen.MousePointer = 0
End Function
Private Sub ImprimoRetirosPorDevolucion(Cliente As Long, Factura As Long, Nota As Long)
Dim aTexto As String

    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    SeteoImpresoraPorDefecto paIConformeN
    
    vsFicha.PaperSize = 1
    vsFicha.PaperHeight = vsFicha.PageHeight / 2
    vsFicha.Orientation = orPortrait
    
    With vsFicha
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
        .AddTable ">2000|<3500", "RETIRO:| Contado " & Trim(tSerie.Text) & " " & Trim(tNumero.Text), ""
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

Private Function BuscoPosArray(ByVal idArt As Long) As Integer
Dim iPos As Integer
    BuscoPosArray = 0
    For iPos = 0 To UBound(arrDevolucion)
        If arrDevolucion(iPos).idArt = idArt Then BuscoPosArray = iPos: Exit For
    Next iPos
End Function

Private Function ValidoEnvioDevolucion(ByVal lEnvioDev As Long) As Long
Dim rsEnv As rdoResultset
    ValidoEnvioDevolucion = 0
    Cons = "Select * From Envio Where EnvCodigo = " & lEnvioDev _
        & " And EnvEstado in (" & EstadoEnvio.AConfirmar & ", " & EstadoEnvio.AImprimir & ")"
    Set rsEnv = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsEnv.EOF Then
        ValidoEnvioDevolucion = lEnvioDev
    End If
    rsEnv.Close
    If ValidoEnvioDevolucion = 0 Then MsgBox "Envío incorrecto, no podrá emitir la nota.", vbInformation, "ATENCIÓN"
End Function
