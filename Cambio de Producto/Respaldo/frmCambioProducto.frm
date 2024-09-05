VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmCambioProducto 
   Caption         =   "Cambio de Producto por Otro Igual"
   ClientHeight    =   4950
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCambioProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   741
      Appearance      =   1
      ImageList       =   "imgIcono"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario [Ctrl+X]"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo [Ctrl+N]"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar [Ctrl+G]"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar [Ctrl+C]"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ingreso a Taller"
      ForeColor       =   &H00000080&
      Height          =   2055
      Left            =   60
      TabIndex        =   27
      Top             =   2760
      Width           =   7635
      Begin VB.TextBox tUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox tAclaracion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Top             =   1380
         Width           =   6375
      End
      Begin AACombo99.AACombo cReparar 
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Top             =   1020
         Width           =   2235
         _ExtentX        =   3942
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsMotivo 
         Height          =   1035
         Left            =   4020
         TabIndex        =   18
         Top             =   300
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   1826
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
      Begin VB.TextBox tMotivo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   300
         Width           =   2715
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario:"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "&Aclaración:"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "&Reparar en:"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Motivos:"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Producto que se lleva"
      ForeColor       =   &H00000080&
      Height          =   795
      Left            =   60
      TabIndex        =   26
      Top             =   1920
      Width           =   7635
      Begin VB.TextBox tSerieS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   15
         Top             =   300
         Width           =   2235
      End
      Begin VB.Label Label6 
         Caption         =   "Nº. S&erie:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Producto que devuelve"
      ForeColor       =   &H00000080&
      Height          =   1395
      Left            =   60
      TabIndex        =   25
      Top             =   480
      Width           =   7635
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   13
         Top             =   960
         Width           =   3255
      End
      Begin VB.CommandButton bFactura 
         Caption         =   "Detalle de Factura"
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   300
         Width           =   1755
      End
      Begin VB.TextBox tNumeroD 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6540
         MaxLength       =   6
         TabIndex        =   9
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox tSerieD 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6180
         MaxLength       =   2
         TabIndex        =   8
         Top             =   600
         Width           =   315
      End
      Begin AACombo99.AACombo cSucursal 
         Height          =   315
         Left            =   3420
         TabIndex        =   6
         Top             =   600
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
      End
      Begin AACombo99.AACombo cTipoDoc 
         Height          =   315
         Left            =   1020
         TabIndex        =   4
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.TextBox tCBarra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   1
         Top             =   300
         Width           =   2235
      End
      Begin VB.TextBox tSerieI 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   11
         Top             =   960
         Width           =   2235
      End
      Begin VB.Label Label11 
         Caption         =   "&Producto:"
         Height          =   195
         Left            =   3420
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "&Número:"
         Height          =   195
         Left            =   5460
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Sucursal:"
         Height          =   195
         Left            =   2640
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Factura:"
         Height          =   195
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº. S&erie:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
   End
   Begin vsViewLib.vsPrinter vsFicha 
      Height          =   555
      Left            =   0
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   1635
      _Version        =   196608
      _ExtentX        =   2884
      _ExtentY        =   979
      _StockProps     =   229
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
   End
   Begin ComctlLib.ImageList imgIcono 
      Left            =   7500
      Top             =   2580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCambioProducto.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCambioProducto.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCambioProducto.frx":0736
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCambioProducto.frx":0848
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuOpNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuOpLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuOpCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuOpLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpSalir 
         Caption         =   "&Salir del Formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "?"
      Begin VB.Menu MnuHelp 
         Caption         =   "Ayuda ..."
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmCambioProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const cteAmarillo = &HC0FFFF
Private gDocumento As Long, gCliente As Long
Private gFechaDocumento As String

Private Sub bFactura_Click()
    If gDocumento > 0 Then EjecutarApp App.Path & "\Detalle de Factura.exe", CStr(gDocumento)
End Sub

Private Sub cReparar_GotFocus()
On Error Resume Next
    With cReparar
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cReparar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cReparar.ListIndex > -1 Then tAclaracion.SetFocus
End Sub

Private Sub cSucursal_Change()
    gDocumento = 0
    tSerieD.Text = "": tNumeroD.Text = ""
End Sub

Private Sub cSucursal_Click()
    gDocumento = 0
    tSerieD.Text = "": tNumeroD.Text = ""
End Sub

Private Sub cSucursal_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tSerieD.SetFocus
End Sub

Private Sub cSucursal_LostFocus()
    cTipoDoc.SelStart = 0
End Sub

Private Sub cTipoDoc_Change()
    gDocumento = 0
    tArticulo.Text = ""
End Sub

Private Sub cTipoDoc_Click()
    gDocumento = 0
End Sub

Private Sub cTipoDoc_GotFocus()
On Error Resume Next
    With cTipoDoc
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cTipoDoc_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then If cTipoDoc.ListIndex > -1 Then cSucursal.SetFocus Else tSerieI.SetFocus
End Sub

Private Sub cTipoDoc_LostFocus()
On Error Resume Next
    cTipoDoc.SelStart = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    ObtengoSeteoForm Me, 500, 500
    MiBotones True
    CargoCombos
    With vsMotivo
        .Rows = 1
        .Cols = 1
        .FormatString = "Motivo"
    End With
    PrueboBandejaImpresora
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    Set clsGeneral = Nothing
    Set miConexion = Nothing
End Sub
Private Sub CargoCombos()
On Error GoTo ErrCC
    cTipoDoc.Clear
    cTipoDoc.AddItem "Contado": cTipoDoc.ItemData(cTipoDoc.NewIndex) = TipoDocumento.Contado
    cTipoDoc.AddItem "Crédito": cTipoDoc.ItemData(cTipoDoc.NewIndex) = TipoDocumento.Credito
    cTipoDoc.AddItem "Remito": cTipoDoc.ItemData(cTipoDoc.NewIndex) = TipoDocumento.Remito
    
    'Cargo Sucursales---------------------------------------------------------------------------
    Cons = "Select SucCodigo, SucAbreviacion from Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cSucursal
    CargoCombo Cons, cReparar
    '-----------------------------------------------------------------------------------------------
    Exit Sub
ErrCC:
    clsGeneral.OcurrioError "Error al cargar combos", Err.Description
End Sub

Private Sub MiBotones(ByVal bNuevo As Boolean)
    With tooMenu
        .Buttons("nuevo").Enabled = bNuevo
        .Buttons("grabar").Enabled = Not bNuevo
        .Buttons("cancelar").Enabled = Not bNuevo
    End With
    MnuOpNuevo.Enabled = bNuevo
    MnuOpGrabar.Enabled = Not bNuevo
    MnuOpCancelar.Enabled = Not bNuevo
    EstadoObjetos Not bNuevo
End Sub

Private Sub EstadoObjetos(ByVal bHabilito As Boolean)
    
    tCBarra.Text = ""
    
    cTipoDoc.Text = "": cSucursal.Text = "": tSerieD.Text = "": tNumeroD.Text = ""
    tSerieI.Text = "": tArticulo.Text = "": tSerieS.Text = "": tMotivo.Text = "": vsMotivo.Rows = 1
    cReparar.Text = "": tAclaracion.Text = "": tUsuario.Text = ""
    
    tCBarra.Enabled = bHabilito
    bFactura.Enabled = bHabilito
    cTipoDoc.Enabled = bHabilito
    cSucursal.Enabled = bHabilito
    tSerieD.Enabled = bHabilito
    tNumeroD.Enabled = bHabilito
    tSerieI.Enabled = bHabilito
    tSerieS.Enabled = bHabilito
    tMotivo.Enabled = bHabilito
    vsMotivo.Enabled = bHabilito
    cReparar.Enabled = bHabilito
    tAclaracion.Enabled = bHabilito
    tUsuario.Enabled = bHabilito
    tArticulo.Enabled = bHabilito
    
    If bHabilito Then
        tCBarra.BackColor = vbWindowBackground
        cTipoDoc.BackColor = cteAmarillo
        cSucursal.BackColor = cteAmarillo
        tSerieD.BackColor = cteAmarillo
        tNumeroD.BackColor = cteAmarillo
        tSerieI.BackColor = vbWindowBackground
        tArticulo.BackColor = cteAmarillo
        tSerieS.BackColor = vbWindowBackground
        tMotivo.BackColor = vbWindowBackground
        vsMotivo.BackColor = vbWindowBackground
        cReparar.BackColor = cteAmarillo
        tAclaracion.BackColor = cteAmarillo
        tUsuario.BackColor = cteAmarillo
    Else
        tCBarra.BackColor = vbButtonFace
        cTipoDoc.BackColor = vbButtonFace
        cSucursal.BackColor = vbButtonFace
        tSerieD.BackColor = vbButtonFace
        tNumeroD.BackColor = vbButtonFace
        tSerieI.BackColor = vbButtonFace
        tSerieS.BackColor = vbButtonFace
        tMotivo.BackColor = vbButtonFace
        vsMotivo.BackColor = vbButtonFace
        cReparar.BackColor = vbButtonFace
        tAclaracion.BackColor = vbButtonFace
        tUsuario.BackColor = vbButtonFace
        tArticulo.BackColor = vbButtonFace
        
        tCBarra.Text = ""
        cTipoDoc.Text = ""
        cSucursal.Text = ""
        tSerieD.Text = ""
        tNumeroD.Text = ""
        tSerieI.Text = ""
        tSerieS.Text = ""
        tMotivo.Text = ""
        vsMotivo.Rows = 1
        cReparar.Text = ""
        tAclaracion.Text = ""
        tUsuario.Text = ""
        tArticulo.Text = ""
        cTipoDoc.Text = "": cSucursal.Text = "": cReparar.Text = ""
    End If
    
End Sub

Private Sub FormatoBarras(Texto As String)
Dim aCodDoc As Long
Dim gTipo As Integer

    On Error GoTo errInt
    
    Texto = UCase(Texto)
    gTipo = CLng(Mid(Texto, 1, InStr(Texto, "D") - 1))
    aCodDoc = CLng(Trim(Mid(Texto, InStr(Texto, "D") + 1, Len(Texto))))
    
    Select Case gTipo
        Case TipoDocumento.Contado, TipoDocumento.Credito: BuscoDocumento Tipo:=gTipo, Codigo:=aCodDoc
        Case Else:  MsgBox "El código de barras ingresado no es correcto. El documento no coincide con los predefinidos (contado ó crédito).", vbCritical, "ATENCIÓN"
    End Select
    Screen.MousePointer = 0
    Exit Sub
    
errInt:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al interpretar el código de barras."
End Sub

Private Sub BuscoDocumento(Optional Tipo As Integer, Optional Sucursal As Long, Optional Serie As String, Optional Numero As Long, Optional Codigo As Long = 0)
On Error GoTo errBD
    
    Screen.MousePointer = 11
    
    If Codigo <> 0 Then
        Cons = "Select * from Documento Where DocCodigo = " & Codigo
        If Tipo > 0 Then Cons = Cons & " And DocTipo = " & Tipo
    Else
        Cons = "Select * from Documento" _
                & " Where DocSucursal = " & Sucursal _
                & " And DocTipo = " & Tipo & " And DocSerie = '" & Trim(Serie) & "'" _
                & " And DocNumero = " & Numero
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    gDocumento = 0
    
    If Not RsAux.EOF Then
        
        gFechaDocumento = RsAux!DocFModificacion
        
        'Cargo los datos del Documento Seleccionado-----------------
        BuscoCodigoEnCombo cTipoDoc, RsAux!DocTipo
        BuscoCodigoEnCombo cSucursal, RsAux!DocSucursal
        tSerieD.Text = Trim(RsAux!DocSerie)
        tNumeroD.Text = RsAux!DocNumero
        tNumeroD.Tag = RsAux!DocFecha
        '-----------------------------------------------------------------------------
        
        If RsAux!DocAnulado Then
            RsAux.Close: Screen.MousePointer = 0
            MsgBox "El documento ingresado ha sido anulado. Verifique", vbCritical, "DOCUMENTO ANULADO"
            gDocumento = 0
            Exit Sub
        Else
            If Not IsNull(RsAux!DocPendiente) Then
                RsAux.Close: Screen.MousePointer = 0
                MsgBox "La mercadería está pendiente de entrega. Verifique", vbInformation, "ATENCIÓN"
                gDocumento = 0
                Exit Sub
            End If
        End If
        gDocumento = RsAux!DocCodigo
        gCliente = RsAux!DocCliente
        RsAux.Close
        'Solo cargo los artìculos que retiro o se le enviaron.
        Screen.MousePointer = 0
    Else
        RsAux.Close
        tNumeroD.Tag = ""
        Screen.MousePointer = 0
        gDocumento = 0: gCliente = 0
        MsgBox "No existe un documento para las características ingresadas.", vbExclamation, "ATENCIÓN"
    End If
    Exit Sub
errBD:
    clsGeneral.OcurrioError "Error al cargar el documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Label1_Click()
    With tSerieI
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label10_Click()
On Error Resume Next
    With tUsuario
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label11_Click()
On Error Resume Next
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label2_Click()
On Error Resume Next
    With tCBarra
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label3_Click()
On Error Resume Next
    With cTipoDoc
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label4_Click()
On Error Resume Next
    With cSucursal
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub Label5_Click()
On Error Resume Next
    With tSerieD
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label6_Click()
On Error Resume Next
    With tSerieS
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label7_Click()
On Error Resume Next
    With tMotivo
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label8_Click()
On Error Resume Next
    With cReparar
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label9_Click()
On Error Resume Next
    With tAclaracion
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub MnuOpCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuOpGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuOpNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuOpSalir_Click()
    Unload Me
End Sub

Private Sub tAclaracion_GotFocus()
On Error Resume Next
    With tAclaracion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tAclaracion_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And Trim(tAclaracion.Text) <> "" Then tUsuario.SetFocus
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = ""
End Sub

Private Sub tArticulo_GotFocus()
On Error Resume Next
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        'Veo si ya tengo ingresado.
        If Val(tArticulo.Tag) > 0 Then
            tSerieS.SetFocus
        Else
            If Trim(tArticulo.Text) <> "" Then
                If IsNumeric(tArticulo.Text) Then
                    BuscoArticuloPorCodigo Val(tArticulo.Text)
                Else
                    BuscoArticuloPorNombre
                End If
                If Val(tArticulo.Tag) > 0 Then
                    'Como ingreso por aca tengo que validar que este en el documento.
                    If gDocumento > 0 Then
                        ArticuloEnDocumento
                    Else
                        cTipoDoc.SetFocus
                    End If
                End If
            End If
        End If
    End If
    
End Sub

Private Sub tCBarra_Change()
    gDocumento = 0
End Sub

Private Sub tCBarra_GotFocus()
On Error Resume Next
    With tCBarra
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCBarra_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Trim(tCBarra.Text) = "" Then
            cTipoDoc.SetFocus
        Else
            FormatoBarras Trim(tCBarra.Text)
            On Error Resume Next
            If tCBarra.Enabled Then Foco tCBarra
        End If
    End If
End Sub

Private Sub tMotivo_GotFocus()
    With tMotivo
        If .Text = "" Then .Text = "%"
        If .Text = "%" Then .SelStart = Len(.Text): Exit Sub
        .SelStart = 0
        .SelLength = Len(tMotivo.Text)
    End With
End Sub

Private Sub tMotivo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        If Val(tArticulo.Tag) = 0 Then
            MsgBox "Debe ingresar un artículo de ingreso.", vbExclamation, "ATENCIÓN"
            tArticulo.SetFocus
            Exit Sub
        End If
                
        If Trim(tMotivo.Text) = "" Then
            On Error Resume Next
            cReparar.SetFocus
        Else
            On Error GoTo ErrBM
            Screen.MousePointer = 11
                
            Cons = "Select MSeID, Nombre = MSeNombre From MotivoServicio " _
                & " Where MSeTipo = (Select ArtTipo From Articulo Where ArtID = " & Val(tArticulo.Tag) & ")" _
                & " And MSeNombre Like '" & clsGeneral.Replace(tMotivo.Text, " ", "%") & "%'"
            
            tMotivo.Tag = ""
            tMotivo.Tag = ListaAyuda(Cons, "Lista de Motivos")
            
            If Val(tMotivo.Tag) > 0 Then InsertoMotivoEnGrilla Val(tMotivo.Tag)
            
            Screen.MousePointer = 0
            tMotivo.Text = "": tMotivo.Tag = ""
        End If
    End If
    Exit Sub
ErrBM:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al buscar los motivos.", Trim(Err.Description)
End Sub
Private Sub tNumeroD_Change()
    gDocumento = 0
End Sub

Private Sub tNumeroD_GotFocus()
On Error Resume Next
    With tNumeroD
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tNumeroD_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If cTipoDoc.ListIndex = -1 Then
            MsgBox "Debe seleccionar el tipo de documento a buscar.", vbExclamation, "ATENCIÓN"
            Foco cTipoDoc: Exit Sub
        End If
        If cTipoDoc.ItemData(cTipoDoc.ListIndex) <> TipoDocumento.Remito Then
            If cSucursal.ListIndex = -1 Or Trim(tSerieD.Text) = "" Or IsNumeric(tSerieD.Text) Then
                MsgBox "Los datos ingresados no son correctos. Verifique.", vbExclamation, "ATENCIÓN"
                Foco cSucursal: Exit Sub
            End If
            If Not IsNumeric(tNumeroD.Text) Then
                MsgBox "Debe ingresar el número del documento.", vbExclamation, "ATENCIÓN"
                Foco tNumeroD: Exit Sub
            End If
        End If
        '--------------------------------------------------------------------------------------------------------------
        BuscoDocumento cTipoDoc.ItemData(cTipoDoc.ListIndex), cSucursal.ItemData(cSucursal.ListIndex), tSerieD.Text, tNumeroD.Text
        If gDocumento > 0 Then
            'Busco si hay en la factura solo un artículo.
            HayUnoSolo
            If Val(tArticulo.Tag) = 0 Then tSerieI.SetFocus Else tSerieS.SetFocus
        End If
    End If
    
End Sub

Private Sub tooMenu_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "salir": Unload Me
        Case "nuevo": AccionNuevo
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
    End Select
End Sub

Private Sub tSerieD_Change()
    gDocumento = 0
End Sub

Private Sub tSerieD_GotFocus()
On Error Resume Next
    With tSerieD
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tSerieD_KeyPress(KeyAscii As Integer)
On Error Resume Next
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then tNumeroD.SetFocus
End Sub

Private Sub tSerieI_Change()
    tSerieI.Tag = ""
End Sub

Private Sub tSerieI_GotFocus()
    With tSerieI
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tSerieI_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tSerieI.Text) <> "" And tSerieI.Tag = "" Then
            'Busco el producto por nro de serie.
            BuscoProductoPorNroSerie tSerieI.Text
        Else
            If gDocumento > 0 Then
                If Val(tArticulo.Tag) = 0 Then
                    'Voy a Artículo
                    tArticulo.SetFocus
                Else
                    tSerieS.SetFocus
                End If
            Else
                tCBarra.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub BuscoProductoPorNroSerie(ByVal sSerie As String)
On Error GoTo errBP
Dim lCant As Long, iPend As Integer
    
    'Busco en la tabla producto si tengo alguno con ese nro. de serie.
    Cons = "Select Count(*) From Producto " _
        & " Where ProNroSerie = '" & Trim(sSerie) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    lCant = RsAux(0)
    RsAux.Close
    
    If lCant > 0 Then
        
        Cons = "Select ProCodigo, ArtNombre, IsNull(DocSerie + Cast(DocNumero as char),'') as Factura , IsNull(CEmFantasia, rTrim(CPeNombre1) + ' '  + rTrim(CPeApellido1)) as 'Cliente' " _
            & " From Producto Left Outer Join Documento On DocCodigo = ProDocumento" _
                    & " Left Outer Join CPersona On ProCliente = CPeCliente " _
                    & " Left Outer Join CEmpresa On ProCliente = CEmCliente " _
            & " , Articulo " _
            & " Where ProNroSerie = '" & Trim(sSerie) & "'" _
            & " And ProArticulo = ArtID"
        'Presento lista de ayuda con los productos que contiene el nro. de serie.
        lCant = ListaAyuda(Cons, "Productos con Número de Serie")
        
        If lCant > 0 Then
            'Seleccionó un producto.
            'Cargo Artículo.
            Screen.MousePointer = 11
            Cons = "Select * From Producto, Articulo Where ProCodigo = " & lCant _
                & " And ProArticulo = ArtID "
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not IsNull(RsAux!ProNroSerie) Then
                tSerieI.Text = Trim(RsAux!ProNroSerie)
            Else
                tSerieI.Text = ""
            End If
            tSerieI.Tag = "P" & lCant
            tArticulo.Text = Trim(RsAux!ArtNombre)
            tArticulo.Tag = RsAux!ArtID
            If RsAux!ArtNroSerie Then tSerieS.Tag = 1 Else tSerieS.Tag = ""
            If Not IsNull(RsAux!ProDocumento) Then
                lCant = RsAux!ProDocumento
            Else
                lCant = 0
            End If
            RsAux.Close
            
            'Váldio que el producto no tenga un servicio abierto.
            If TieneServicioAbierto(Val(Mid(tSerieI.Tag, 2))) Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            If gDocumento = 0 And lCant = 0 Then
                cTipoDoc.SetFocus
                Screen.MousePointer = 0
                MsgBox "Debe ingresar la factura de compra.", vbInformation, "ATENCIÓN"
                Exit Sub
            Else
                If gDocumento > 0 Then
            
                    If lCant = 0 Then
                        ValidoArticuloEntregado
                        If gDocumento > 0 Then tSerieS.SetFocus
                        Exit Sub
                    Else
                        If gDocumento <> lCant Then
                            cTipoDoc.Text = "": cSucursal.Text = ""
                            BuscoDocumento Codigo:=gDocumento
                            'Válido los pendientes.
                            ValidoArticuloEntregado
                            If gDocumento > 0 Then tSerieS.SetFocus
                            Exit Sub
                        Else
                            ValidoArticuloEntregado
                            If gDocumento > 0 Then tSerieS.SetFocus
                            Exit Sub
                        End If
                    End If
                Else
                    cTipoDoc.Text = "": cSucursal.Text = ""
                    BuscoDocumento Codigo:=lCant
                    ValidoArticuloEntregado
                    If gDocumento > 0 Then tSerieS.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '--------------------------------------------------------
    '               No esta en la tabla Producto.
    '--------------------------------------------------------
    
    'Busco en la tabla productos vendidos.
    Cons = "Select  ArtID, ArtNombre, DocSerie + '  ' + Cast (DocNumero as Char) as 'Factura' From ProductosVendidos, Articulo, Documento" _
        & " Where PVeNSerie = '" & Trim(sSerie) & "' And PVeArticulo = ArtID And PVeDocumento = DocCodigo"
    lCant = ListaAyuda(Cons, "Productos con Número de Serie")
    '--------------------------------------------------------
    If lCant > 0 Then
        Cons = "Select * From ProductosVendidos, Articulo" _
            & " Where PVeNSerie = '" & sSerie & "' And PVeArticulo = " & lCant & " And PVeArticulo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        tSerieI.Text = Trim(RsAux!PVeNSerie)
        tSerieI.Tag = "V"       'Me digo que esta en la tabla productos vendidos.
        tArticulo.Text = Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!ArtID
        If RsAux!ArtNroSerie Then tSerieS.Tag = 1 Else tSerieS.Tag = ""
        lCant = RsAux!PVeDocumento
        RsAux.Close
        'Limpio los datos del documento.
        cTipoDoc.Text = "": cSucursal.Text = ""
        BuscoDocumento Codigo:=lCant
        ValidoArticuloEntregado
        If gDocumento > 0 Then tSerieS.SetFocus
    Else
        'No selecciono o no encontró un artículo.
        If gDocumento = 0 Then
            cTipoDoc.SetFocus
        Else
            tArticulo.SetFocus
        End If
    End If
    Exit Sub
    
errBP:
    clsGeneral.OcurrioError "Error al buscar el producto por nro. de serie.", Err.Description, "Error"
End Sub

Private Function ListaAyuda(ByVal Cons As String, ByVal sTitulo As String) As Long
On Error GoTo errLP
    
    Dim objLista As New clsListadeAyuda
    ListaAyuda = 0
    If objLista.ActivarAyuda(cBase, Cons, 4500, 1, sTitulo) > 0 Then
        ListaAyuda = objLista.RetornoDatoSeleccionado(0)
    End If
    Set objLista = Nothing
    Exit Function
    
errLP:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al presentar la lista de ayuda.", Err.Description
End Function

Private Sub ValidoArticuloEntregado()
Dim cCantV As Currency, cPend As Currency

    'Tengo ingresado un Documento pero el producto que selecciono no esta asociado a un documento.
    'Válido que la factura contenga este producto y el mismo fue entregado (ojo puede estar a enviar).
    Cons = "Select * From Renglon Where RenDocumento = " & gDocumento & " And RenArticulo = " & Val(tArticulo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "La factura que seleccionó no tiene asociado el producto, debe ingresar el documento correcto.", vbExclamation, "ATENCIÓN"
        cTipoDoc.Text = "": cSucursal.Text = "": cTipoDoc.SetFocus
        Exit Sub
    Else
        cCantV = RsAux!RenCantidad
        cPend = RsAux!RenARetirar
        RsAux.Close
    End If
    
    'Válido que no tenga pendientes.
    If cCantV = cPend Then
        MsgBox "No hay artículos entregados para esa factura, verifique.", vbExclamation, "ATENCIÓN"
        cTipoDoc.Text = "": cTipoDoc.SetFocus
        Exit Sub
    End If
    
    'Busco mercadería en Envío
    cPend = cPend + CantidadArticuloEnEnvio
    If cCantV = cPend Then
        MsgBox "No hay artículos entregados para esa factura, verifique.", vbExclamation, "ATENCIÓN"
        cTipoDoc.Text = "": cTipoDoc.SetFocus
        Exit Sub
    End If
    
    'Busco Mercadería en Remito.
    cPend = cPend + CantidadArticuloEnRemito
    If cCantV = cPend Then
        MsgBox "No hay artículos entregados para esa factura, verifique.", vbExclamation, "ATENCIÓN"
        cTipoDoc.Text = "": cTipoDoc.SetFocus
        Exit Sub
    End If
    
End Sub

Private Function CantidadArticuloEnEnvio() As Currency
    
    CantidadArticuloEnEnvio = 0
    Cons = "Select * From Envio, RenglonEnvio Where EnvDocumento = " & gDocumento & " And REvArticulo = " & Val(tArticulo.Tag) _
        & " And EnvCodigo = REvEnvio"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        CantidadArticuloEnEnvio = RsAux!REvAEntregar
    End If
    RsAux.Close
    
End Function

Private Function CantidadArticuloEnRemito() As Currency
    
    CantidadArticuloEnRemito = 0
    Cons = "Select * From Remito, RenglonRemito Where RemDocumento = " & gDocumento & " And RReArticulo = " & Val(tArticulo.Tag) _
        & " And RemCodigo = RReRemito"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        CantidadArticuloEnRemito = RsAux!RReAEntregar
    End If
    RsAux.Close
End Function


Private Sub BuscoArticuloPorCodigo(ByVal CodArticulo As Long)
On Error GoTo errBA
    Screen.MousePointer = 11
    Cons = "Select * From Articulo Where ArtCodigo = " & CodArticulo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        tArticulo.Text = Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!ArtID
        If RsAux!ArtNroSerie Then tSerieS.Tag = 1 Else tSerieS.Tag = ""
    Else
        MsgBox "No se encontro un artículo con código: " & CodArticulo & " .", vbInformation, "ATENCIÓN"
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errBA:
    clsGeneral.OcurrioError "Ocurrió un error al bucar el artículo por código.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoArticuloPorNombre()
On Error GoTo errBAN
    Screen.MousePointer = 11
    Cons = "Select * From Articulo Where ArtNombre Like '" & clsGeneral.Replace(tArticulo.Text, " ", "%") & "%'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            tArticulo.Text = Trim(RsAux!ArtNombre)
            tArticulo.Tag = RsAux!ArtID
            If RsAux!ArtNroSerie Then tSerieS.Tag = 1 Else tSerieS.Tag = ""
            RsAux.Close
        Else
            RsAux.Close
            Cons = "Select ArtCodigo, ArtCodigo as 'Código', ArtNombre as 'Producto' From Articulo Where ArtNombre Like '" & clsGeneral.Replace(tArticulo.Text, " ", "%") & "%'"
            tArticulo.Tag = ListaAyuda(Cons, "Lista de Artículos")
            If Val(tArticulo.Tag) > 0 Then
                BuscoArticuloPorCodigo Val(tArticulo.Tag)
            Else
                tArticulo.Tag = ""
            End If
        End If
    Else
        RsAux.Close
        MsgBox "No existe un artículo para los datos ingresados.", vbExclamation, "ATENCIÓN"
    End If
    Screen.MousePointer = 0
    Exit Sub
errBAN:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo por nombre.", Err.Description
End Sub

Private Sub ArticuloEnDocumento()
On Error GoTo errAED
    Screen.MousePointer = 11
    
    Cons = "Select * From Renglon Where RenDocumento = " & gDocumento & " And RenArticulo = " & Val(tArticulo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "El artículo no está en la factura, verifique.", vbExclamation, "ATENCIÓN"
        tArticulo.Text = "": tArticulo.Tag = "": tSerieS.Tag = ""
    End If
    RsAux.Close
    
    If Val(tArticulo.Tag) > 0 Then
        ValidoArticuloEntregado
        If gDocumento > 0 Then
            If Trim(tSerieI.Text) = "" Then
                'Verifico si el producto tuvo nro. de serie y no lo ingreso.
                Cons = "Select * From Producto " _
                    & " Where ProArticulo = " & Val(tArticulo.Tag) _
                    & " And ProDocumento = " & gDocumento
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    RsAux.MoveNext
                    If RsAux.EOF Then
                        If MsgBox("Existe un producto ingresado para el cliente." & vbCrLf & "¿Desea asignarlo al mismo?" & vbCrLf & "VALIDE EL NRO. DE SERIE, SI LO TIENE", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                            RsAux.MoveFirst
                            If Not IsNull(RsAux!ProNroSerie) Then tSerieI.Text = Trim(RsAux!ProNroSerie)
                            tSerieI.Tag = "P" & RsAux!ProCodigo
                        End If
                        RsAux.Close
                    Else
                        RsAux.Close
                        tSerieI.Tag = ""
                        Cons = "Select ProCodigo, ProCodigo as Código, IsNull(ProNroSerie, '') as 'Nro. de Serie' , IsNull(rTrim(DocSerie) + ' ' + Cast(DocNumero as Char), '') as Factura From Producto, Documento " _
                            & " Where ProArticulo = " & Val(tArticulo.Tag) _
                            & " And ProDocumento = " & gDocumento
                        tSerieI.Tag = ListaAyuda(Cons, "Productos del Documento")
                        If Val(tSerieI.Tag) > 0 Then tSerieI.Tag = "P" & Trim(tSerieI.Tag) Else tSerieI.Tag = ""
                    End If
                Else
                    RsAux.Close
                    'Veo si esta en productosvendidos
                    Cons = "Select * From ProductosVendidos Where PVeDocumento = " & gDocumento & " and PVeArticulo = " & Val(tArticulo.Tag)
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not RsAux.EOF Then
                        MsgBox "Existe un producto para el documento con nro. de serie asignado el sistema le cargará este dato, VALIDE QUE SEA CORRECTO.", vbInformation, "ATENCIÓN"
                        tSerieI.Text = Trim(RsAux!PVeNSerie)
                    End If
                    RsAux.Close
                End If
            End If
            tSerieS.SetFocus
        Else
            tArticulo.Text = "": tArticulo.Tag = ""
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub

errAED:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al validar el artículo.", Err.Description
End Sub

Private Sub tSerieS_GotFocus()
On Error Resume Next
    With tSerieS
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tSerieS_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Val(tArticulo.Tag) = 0 Then
            tSerieS.Text = "": tArticulo.SetFocus
            MsgBox "Ingrese el artículo que devuelve.", vbInformation, "ATENCIÓN"
        Else
            If Val(tSerieS.Tag) > 0 Then
                If Trim(tSerieS.Text) = "" Then
                    MsgBox "Debe ingresar el número de serie del artículo, de lo contrario no podrá grabar los datos.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
            End If
            tMotivo.SetFocus
        End If
    End If
    
End Sub

Private Sub tSerieS_LostFocus()
On Error Resume Next
    If Val(tSerieS.Tag) > 0 And tSerieS.Enabled Then
        If Trim(tSerieS.Text) = "" Then
            MsgBox "Debe ingresar el número de serie del artículo, de lo contrario no podrá grabar los datos.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
    End If
End Sub

Private Sub InsertoMotivoEnGrilla(ByVal lID As Long)
On Error GoTo errIM
Dim rsM As rdoResultset
Dim sName As String

    Cons = "Select * From MotivoServicio Where MSeID = " & lID
    Set rsM = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    sName = Trim(rsM("MSeNombre"))
    rsM.Close
    
    'Verifico que no este insertado.
    With vsMotivo
        For I = 1 To .Rows - 1
            If Val(.Cell(flexcpData, I, 0)) = lID Then MsgBox "El motivo ya fue ingresado, verifique.", vbInformation, "ATENCIÓN": Exit Sub
        Next I
        .AddItem sName
        .Cell(flexcpData, .Rows - 1, 0) = lID
    End With
    Exit Sub
    
errIM:
    clsGeneral.OcurrioError "Error al insertar el motivo en la grilla.", Err.Description
End Sub

Private Sub tUsuario_Change()
    tUsuario.Tag = ""
End Sub

Private Sub tUsuario_GotFocus()
    With tUsuario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = 0
            tUsuario.Tag = BuscoUsuarioDigito(Val(tUsuario.Text), True)
            If Val(tUsuario.Tag) > 0 Then AccionGrabar
        Else
            MsgBox "Ingrese su dígito de usuario.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub
Private Sub vsMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
            Case vbKeyDelete: If vsMotivo.Row > 0 Then vsMotivo.RemoveItem vsMotivo.Row
        End Select
End Sub
Private Sub AccionNuevo()
    MiBotones False
    tSerieI.SetFocus
End Sub
Private Sub AccionGrabar()
    If ValidoDatos Then
        If MsgBox("¿Confirma cambiar el producto por otro igual y dar ingreso a Servicio el producto que entrega el cliente?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
            AccionGrabarTaller
        End If
    End If
End Sub
Private Sub AccionCancelar()
    MiBotones True
End Sub
Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    If gDocumento = 0 Then
        MsgBox "Debe ingresar un documento.", vbExclamation, "ATENCIÓN"
        cTipoDoc.SetFocus: Exit Function
    End If
    If Val(tArticulo.Tag) = 0 Then
        MsgBox "Debe ingresar un artículo a devolver.", vbExclamation, "ATENCIÓN"
        tArticulo.SetFocus: Exit Function
    End If
    If Trim(tSerieS.Tag) <> "" Then
        'Es necesario que ingrese el nro. de serie.
        If Trim(tSerieS.Text) = "" Then
            MsgBox "Debe ingresar un número de serie para el artículo que se lleva el cliente.", vbExclamation, "ATENCIÓN"
            tSerieS.SetFocus: Exit Function
        End If
    End If
    If cReparar.ListIndex = -1 Then
        MsgBox "Debe ingresar el local de reparación.", vbExclamation, "ATENCIÓN"
        cReparar.SetFocus: Exit Function
    End If
    If Trim(tAclaracion.Text) = "" Then
        MsgBox "Debe ingresar un comentario de reparación.", vbExclamation, "ATENCIÓN"
        tAclaracion.SetFocus: Exit Function
    End If
    If Val(tUsuario.Tag) = 0 Then
        MsgBox "Ingrese su dígito de usuario.", vbExclamation, "ATENCIÓN"
        tUsuario.SetFocus: Exit Function
    End If
    If Mid(tSerieI.Tag, 1, 1) = "P" Then
        If TieneServicioAbierto(Val(Mid(tSerieI.Tag, 2))) Then
            Exit Function
        End If
    End If
    ValidoDatos = True
End Function
Private Sub AccionGrabarTaller()
On Error GoTo ErrBT
Dim IdServicio As Long, idProCli As Long, idProEmp As Long

    IdServicio = 0
    Screen.MousePointer = 11
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo ErrResumir
    
    '------------------------------------------------------------------------------------------------------------------------
    'Inserto el nuevo producto al cliente.
    Cons = "Select * From Producto Where ProCliente = " & gCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!ProArticulo = Val(tArticulo.Tag)
    RsAux!ProCliente = gCliente
    If Trim(tSerieS.Text) <> "" Then RsAux!ProNroSerie = Trim(tSerieS.Text)
    RsAux!ProFacturaS = Trim(tSerieD.Text)
    RsAux!ProFacturaN = Trim(tNumeroD.Text)
    If IsDate(tNumeroD.Tag) Then RsAux!ProCompra = Format(CDate(tNumeroD.Tag), "mm/dd/yyyy")
    RsAux!ProFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:nn:ss")
    RsAux!ProDocumento = gDocumento
    RsAux.Update
    RsAux.Close
    
    'Saco el nuevo ID Para el producto del cliente.
    Cons = "Select Max(ProCodigo) From Producto Where ProCliente = " & gCliente & " And ProDocumento = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    idProCli = RsAux(0)
    RsAux.Close
    '------------------------------------------------------------------------------------------------------------------------
    
    'Le inserto un comentario al documento.
    Cons = "Select * From Comentario Where ComCliente = " & gCliente & " And ComDocumento = " & gDocumento _
        & " And ComUsuario = " & Val(tUsuario.Tag)
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!ComCliente = gCliente
    RsAux!ComFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux!ComComentario = "Producto: " & Trim(tArticulo.Text) & Space(10) & "Comentario: " & Trim(tAclaracion.Text)
    RsAux!ComTipo = paTipoComentario
    RsAux!ComUsuario = Val(tUsuario.Tag)
    RsAux!ComDocumento = gDocumento
    RsAux.Update
    RsAux.Close
    
    'Inserto o cambio el producto a la empresa.
    If Trim(tSerieI.Tag) <> "" Then
        If Mid(tSerieI.Tag, 1, 1) = "P" Then
            idProEmp = Val(Mid(tSerieI.Tag, 2))
            Cons = "Select * From Producto Where ProCodigo = " & idProEmp
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            RsAux.Edit
            RsAux!ProCliente = paClienteEmpresa
            RsAux!ProFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:nn:ss")
            RsAux!ProFacturaN = Null
            RsAux!ProFacturaS = Null
            RsAux!ProCompra = Null
            RsAux!ProDireccion = Null
            RsAux!ProDocumento = Null
            RsAux.Update
            RsAux.Close
        Else
            'Elimino el producto de la tabla productos vendidos. Ya que inserte un producto
            Cons = "Select * From ProductosVendidos Where PVeDocumento = " & gDocumento _
                & " And PVeArticulo = " & Val(tArticulo.Tag) & " And PVeNSerie = '" & tSerieI.Text & "'"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                RsAux.Delete
            End If
            RsAux.Close
            
            Cons = "Select * From Producto Where ProCodigo = 0"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            RsAux.AddNew
            RsAux!ProArticulo = Val(tArticulo.Tag)
            RsAux!ProCliente = paClienteEmpresa
            If Trim(tSerieI.Text) <> "" Then RsAux!ProNroSerie = tSerieI.Text
            RsAux!ProFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:nn:ss")
            RsAux.Update
            RsAux.Close
            
            'Saco el nuevo ID Para el producto del cliente.
            Cons = "Select Max(ProCodigo) From Producto Where ProCliente = " & paClienteEmpresa & " And ProArticulo = " & Val(tArticulo.Tag)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            idProEmp = RsAux(0)
            RsAux.Close
            
        End If
    Else
        'No esta en la producto ni en productosvendidos.
            Cons = "Select * From Producto Where ProCodigo = 0"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            RsAux.AddNew
            RsAux!ProArticulo = Val(tArticulo.Tag)
            RsAux!ProCliente = paClienteEmpresa
            If Trim(tSerieI.Text) <> "" Then RsAux!ProNroSerie = tSerieI.Text
            RsAux!ProFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:nn:ss")
            RsAux.Update
            RsAux.Close
            
            'Saco el nuevo ID Para el producto del cliente.
            Cons = "Select Max(ProCodigo) From Producto Where ProCliente = " & paClienteEmpresa & " And ProArticulo = " & Val(tArticulo.Tag)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            idProEmp = RsAux(0)
    End If
    '-------------------------------------------------------------------------
        
    IdServicio = InsertoServicio(idProEmp, EstadoP.FueraGarantia, EstadoS.Taller, cReparar.ItemData(cReparar.ListIndex), Trim(tAclaracion.Text), Usuario:=tUsuario.Tag)
    If vsMotivo.Rows > 1 Then InsertoMotivos IdServicio
        
    'Si ingresa directo al local inserto la tabla taller.
    If cReparar.ItemData(cReparar.ListIndex) = paCodigoDeSucursal Then InsertoServicioTaller IdServicio, tUsuario.Tag
    HagoCambioDeEstado Val(tArticulo.Tag), paEstadoARecuperar, IdServicio
        
    cBase.CommitTrans
    'Imprimo fichas.
    ImprimoFichaTaller IdServicio
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción.", Trim(Err.Description)
    Screen.MousePointer = 0
    Exit Sub
ErrResumir:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información de taller.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Function TieneServicioAbierto(ByVal idProducto As Long) As Boolean

    TieneServicioAbierto = False
    Cons = "Select * From Servicio Where SerProducto = " & idProducto _
            & " And SerEstadoServicio Not IN (" & EstadoS.Anulado & ", " & EstadoS.Cumplido & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        MsgBox "El producto seleccionado tiene un servicio abierto (Nro.: " & RsAux!SerCodigo & ")." & vbCrLf & "DEBE CUMPLIR EL SERVICIO ANTES DE CAMBIAR EL PRODUCTO", vbInformation, "ATENCIÓN"
        TieneServicioAbierto = True
    End If
    RsAux.Close

End Function

Private Function InsertoServicio(idProducto As Long, EstadoProducto As Integer, EstadoServicio As Integer, LocalReparacion As Long, Optional Comentario As String = "", Optional LocalRecepcion As Long = -1, Optional Usuario As Long = -1) As Long
    
    If LocalRecepcion = -1 Then LocalRecepcion = paCodigoDeSucursal
    If Usuario = -1 Then Usuario = paCodigoDeUsuario
    
    '---------------------------------------------
    'Inserto
    Cons = "INSERT INTO Servicio (SerProducto, SerFecha, SerEstadoProducto, SerLocalIngreso, " _
        & " SerLocalReparacion, SerEstadoServicio, SerUsuario, SerModificacion, SerComentario) Values (" _
        & idProducto & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & EstadoProducto & ", " & LocalRecepcion
    
    If LocalReparacion = 0 Then Cons = Cons & ", Null " Else Cons = Cons & ", " & LocalReparacion
    
    Cons = Cons & ", " & EstadoServicio & ", " & Usuario & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', "
    If Comentario = "" Then Cons = Cons & "Null)" Else Cons = Cons & "'" & Comentario & "')"
    cBase.Execute (Cons)
    
    '---------------------------------------------
    'Saco el mayor código de servicio.
    Cons = "Select Max(SerCodigo) From Servicio"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    InsertoServicio = RsAux(0)
    RsAux.Close
    '---------------------------------------------
    
End Function

Private Sub InsertoMotivos(IdServicio As Long)
    With vsMotivo
        For I = 1 To .Rows - 1
            Cons = "Insert Into ServicioRenglon (SReServicio, SReTipoRenglon,  " _
                & " SReMotivo, SReCantidad) Values (" & IdServicio & ", " & TipoRenglonS.Llamado & ",  " & Val(.Cell(flexcpData, I, 0)) & ", 1)"
            cBase.Execute (Cons)
        Next I
    End With
End Sub

Private Sub InsertoServicioTaller(IdServicio As Long, Optional Usuario As Integer = -1)

    If Usuario = -1 Then Usuario = paCodigoDeUsuario
    If cReparar.ItemData(cReparar.ListIndex) <> paCodigoDeSucursal Then
        'Inserto también el local para el traslado.
        Cons = "Insert Into Taller(TalServicio, TalFIngresoRealizado, TalFIngresoRecepcion, TalModificacion, TalUsuario, TalLocalAlCliente) Values (" _
            & IdServicio & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
            & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & Usuario & ", " & cReparar.ItemData(cReparar.ListIndex) & ")"
    Else
        Cons = "Insert Into Taller(TalServicio, TalFIngresoRealizado, TalFIngresoRecepcion, TalModificacion, TalUsuario) Values (" _
            & IdServicio & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
            & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & Usuario & ")"
    End If
    cBase.Execute (Cons)
    
End Sub

Private Sub ImprimoFichaTaller(IdServicio As Long)
Dim aTexto As String

    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    'Seteo por defecto la impresora
    
    SeteoImpresoraPorDefecto paIConformeN
    
    With vsFicha
        '.PaperSize = 1
        '.PaperHeight = .PaperHeight / 2
        '.Orientation = orLandscape
        .Orientation = orPortrait
        .PaperSize = 256
        .PaperHeight = 7920 '.PaperHeight / 2

        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
        .FileName = "Ficha de Ingreso a Taller"
        .FontSize = 10
        .TableBorder = tbNone
        
        .TextAlign = taRightBaseline
        .FontBold = True
        .AddTable ">2000|<1500", "Servicio Número:|" & IdServicio, ""
        .Paragraph = "": .AddTable ">2000|<1500", "|STOCK", ""
        .FontBold = False
        .TextAlign = taLeftBaseline
        .FontSize = 8.25
        .Paragraph = "": .Paragraph = "": .Paragraph = ""
        .AddTable "<900|<1800|>1400|<1000", "Fecha:|" & Format(gFechaServidor, "d-Mmm yyyy hh:mm") & "|Recibido por:|" & tUsuario.Text, ""
        
        .Paragraph = ""
        aTexto = aTexto & sNombreEmpresa
            
        .AddTable "<900|<4500|<950|4600", aTexto, ""
        .AddTable "<900|<9000", "Dirección:|" & sDireccion, ""
        
        .Paragraph = ""
        .FontBold = True
        aTexto = Trim(tArticulo.Text)
        .AddTable "<900|<8000", "Artículo:|" & aTexto, ""
        .FontBold = False
        
        .AddTable "<900|<1500|<1500|<1100|<1200|<1800|<900|<500", "Factura:|" & "|Fecha Compra:|" & "|Nro. Serie:|" & Trim(tSerieI.Text) & "|Estado:| Fuera de Garantía", ""
        
        .Paragraph = ""
        .AddTable "<900|3000", "Local:|" & Trim(cReparar.Text), ""
        
        .Paragraph = ""
        aTexto = ""
        For I = 1 To vsMotivo.Rows - 1
            If aTexto = "" Then aTexto = Trim(vsMotivo.Cell(flexcpText, I, 1)) Else aTexto = aTexto & ", " & Trim(vsMotivo.Cell(flexcpText, I, 1))
        Next I
        .AddTable "<900|<10100", "Motivos:|" & aTexto, ""
        If Trim(tAclaracion.Text) <> "" Then .AddTable "<1000|<10000", "Aclaración:|" & Trim(tAclaracion.Text), ""
        .Paragraph = "": .Paragraph = ""
        .FontSize = 7
        aTexto = "1) - Para retirar el aparato es indispensable presentar esta boleta. -"
        .AddTable "900|10100", "Nota:|" & aTexto, ""
        aTexto = "2) - El plazo de retiro del aparato es de 90 días contados a partir de la fecha de esta boleta. Expirado dicho plazo se perderá todo derecho a reclamo " _
            & "sobre el mismo. -"
        .AddTable "900|10100", "|" & aTexto, ""
        .Paragraph = ""
        .FontSize = 9.25
        .Paragraph = "Vía Cliente"
        
        If idSucGallinal <> paCodigoDeSucursal Then
            .EndDoc
            .PrintDoc   'Cliente
            
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
            .FileName = "Ficha de Ingreso a Taller"
        Else
            .Paragraph = "": .Paragraph = ""
            .Paragraph = "----------------------------------------------------------------------------------------------------------------"
            .Paragraph = "": .Paragraph = ""
        End If
        
        .FontSize = 10
        .TableBorder = tbNone
        
        .TextAlign = taRightBaseline
        .FontBold = True
        .AddTable ">2000|<1500", "Servicio Número:|" & IdServicio, ""
        .Paragraph = "": .AddTable ">2000|<1500", "|STOCK", ""
        .FontBold = False
        .TextAlign = taLeftBaseline
        .FontSize = 8.25
        .Paragraph = "": .Paragraph = "": .Paragraph = ""
        .AddTable "<900|<1800|>1400|<1000", "Fecha:|" & Format(gFechaServidor, "d-Mmm yyyy hh:mm") & "|Recibido por:|" & tUsuario.Text, ""
        
        .Paragraph = ""
        aTexto = aTexto & sNombreEmpresa & "|Teléfono:|"
            
        .AddTable "<900|<4500|<950|4600", aTexto, ""
        .AddTable "<900|<9000", "Dirección:|" & sDireccion, ""
        
        .Paragraph = ""
        .FontBold = True
        .AddTable "<900|<8000", "Artículo:|" & Trim(tArticulo.Text), ""
        .FontBold = False
        
        .AddTable "<900|<1500|<1500|<1100|<1200|<1800|<900|<500", "Factura:|" & "|Fecha Compra:|" & "|Nro. Serie:|" & Trim(tSerieI.Text) & "|Estado:| Fuera de Garantía", ""
        
        .Paragraph = ""
        .AddTable "<900|3000", "Local:|" & Trim(cReparar.Text), ""
        
        .Paragraph = ""
        aTexto = ""
        For I = 1 To vsMotivo.Rows - 1
            If aTexto = "" Then aTexto = Trim(vsMotivo.Cell(flexcpText, I, 1)) Else aTexto = aTexto & ", " & Trim(vsMotivo.Cell(flexcpText, I, 1))
        Next I
        .AddTable "<900|<10100", "Motivos:|" & aTexto, ""
        If Trim(tAclaracion.Text) <> "" Then .AddTable "<1000|<10000", "Aclaración:|" & Trim(tAclaracion.Text), ""
        .Paragraph = "": .Paragraph = ""
        .Paragraph = "Recibido:"
        .Paragraph = ""
        .Paragraph = "Reparado:"
        .Paragraph = ""
        .FontSize = 9.25
        .Paragraph = "Vía Archivo"
        .EndDoc
        .PrintDoc   'Archivo
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión para el servicio " & IdServicio, Err.Description
End Sub


'------------------------------------------------------------------------------------------------------------------------------------
'   Setea la impresora pasada como parámetro como: por defecto
'------------------------------------------------------------------------------------------------------------------------------------
Private Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

Private Sub HagoCambioDeEstado(IDArticulo As Long, EstadoNuevo As Integer, IdServicio As Long)
    
    'Cambio el estado del artículo como Sano a Recuperar.
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, EstadoNuevo, 1, TipoDocumento.ServicioCambioEstado, IdServicio
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, -1, TipoDocumento.ServicioCambioEstado, IdServicio
        
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, EstadoNuevo, 1, 1
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, 1, -1
    
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, EstadoNuevo, 1
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, -1
    
End Sub

Private Sub PrueboBandejaImpresora()
On Error GoTo ErrPBI
    
    If idSucGallinal = paCodigoDeSucursal Then
        With vsFicha
            .PageBorder = pbNone
            .Device = paIConformeN
            If .Device <> paIConformeN Then MsgBox "Ud no tiene instalada la impresora para imprimir Conformes. Avise al administrador.", vbExclamation, "ATENCIÒN"
            If .PaperBins(paIConformeB) Then .PaperBin = paIConformeB Else MsgBox "Esta mal definida la bandeja de conformes en su sucursal, comuniquele al administrador.", vbInformation, "ATENCIÓN": paIConformeB = .PaperBin
            .PaperSize = 256 'Hoja carta
            .Orientation = orPortrait
           ' .PaperHeight = .PaperHeight / 2
            .MarginTop = 300
            .MarginLeft = 500
        End With
    Else
        With vsFicha
            .PageBorder = pbNone
            .Orientation = orPortrait
            .MarginTop = 300
            .MarginLeft = 500
        End With
    End If
    Exit Sub
ErrPBI:
    clsGeneral.OcurrioError "Ocurrio un error al setear la impresora, consulte con el administrador de impresión este problema.", Err.Description
End Sub

Private Sub HayUnoSolo()
Dim rsD As rdoResultset
Dim codArt As Long
On Error Resume Next
    Cons = "Select ArtCodigo as 'Código', ArtNombre as 'Artículo' From Renglon, Articulo Where RenDocumento = " & gDocumento _
        & " And RenArticulo = ArtID"
    Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsD.EOF Then
        rsD.MoveNext
        If rsD.EOF Then
            rsD.MoveFirst
            tArticulo.Text = rsD(0)
            rsD.Close
            tArticulo_KeyPress vbKeyReturn
        Else
            rsD.Close
            codArt = ListaAyuda(Cons, "Artículos en la Factura")
            If codArt > 0 Then
                tArticulo.Text = codArt
                tArticulo_KeyPress vbKeyReturn
            End If
        End If
    Else
        rsD.Close
    End If
End Sub
