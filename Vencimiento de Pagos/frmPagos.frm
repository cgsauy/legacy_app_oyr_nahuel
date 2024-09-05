VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plazos"
   ClientHeight    =   3570
   ClientLeft      =   3405
   ClientTop       =   3810
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7725
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   5100
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vencimientos"
      ForeColor       =   &H00000080&
      Height          =   1710
      Left            =   60
      TabIndex        =   10
      Top             =   1560
      Width           =   7575
      Begin VB.TextBox tCofis 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   19
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox tIva 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox tImporte 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   13
         Text            =   "1,000,000.00"
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox tNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   305
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   305
         Left            =   1920
         MaxLength       =   9
         TabIndex        =   11
         Top             =   240
         Width           =   435
      End
      Begin AACombo99.AACombo cComprobante 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
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
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   660
         Width           =   855
         _ExtentX        =   1508
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsCuota 
         Height          =   1395
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2461
         _ConvInfo       =   1
         Appearance      =   0
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
         FocusRect       =   4
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cofis:"
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   1065
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   705
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.:"
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   1380
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   60
      TabIndex        =   9
      Top             =   480
      Width           =   7575
      Begin VB.TextBox tIDCompra 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         MaxLength       =   11
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Total a PAGAR:"
         Height          =   255
         Left            =   4800
         TabIndex        =   22
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lTotalGasto 
         Alignment       =   1  'Right Justify
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
         Left            =   6120
         TabIndex        =   21
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "&ID Compra:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3315
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "bd"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5503
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagos.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagos.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagos.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagos.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagos.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagos.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagos.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagos.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagos.frx":10E2
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
   Begin VB.Menu MnuBases 
      Caption         =   "&Bases"
      Begin VB.Menu MnuBx 
         Caption         =   "MnuBx"
         Index           =   0
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public prmIDGasto As Long

Dim sNuevo As Boolean, sModificar As Boolean
Dim RsCom As rdoResultset
Dim bEsNuevo As Boolean
Dim gFModificacion As Date

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    LoadME
    
    If prmIDGasto <> 0 Then
        CargoCamposDesdeBD prmIDGasto
        If vsCuota.Rows = 1 Then
            If Val(tIDCompra.Tag) <> 0 And MnuNuevo.Enabled Then AccionNuevo: bEsNuevo = True
        End If
    End If
    
End Sub

Private Sub LoadME()
    On Error Resume Next
    
    bEsNuevo = False
    sNuevo = False: sModificar = False
    CargoDatosCombo
    InicializoGrillas
    LimpioFicha
    DeshabilitoIngreso
    Botones False, False, False, False, False, Toolbar1, Me
    
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
    
End Sub
Private Sub CargoDatosCombo()

    On Error Resume Next
    
    'Cargo los valores para los comprobantes de pago
    cComprobante.Clear
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.Compracontado)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.Compracontado
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraCredito)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraCredito
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaCredito)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraNotaCredito
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaDevolucion)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraNotaDevolucion
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraRecibo)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraRecibo
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraReciboDePago)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraReciboDePago
            
    'Cargo las monedas en el combo
    cons = "Select MonCodigo, MonSigno from Moneda"
    CargoCombo cons, cMoneda

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco tProveedor
End Sub

Private Sub Label10_Click()
    Foco tIDCompra
End Sub


Private Sub Label5_Click()
    Foco cMoneda
End Sub

Private Sub Label9_Click()
    Foco tFecha
End Sub

Private Sub MnuBx_Click(Index As Integer)

On Error Resume Next

    If Not AccionCambiarBase(MnuBx(Index).Tag, MnuBx(Index).Caption) Then Exit Sub
    Screen.MousePointer = 11
    
    LoadME
   
    'Cambio el Color del fondo de controles ----------------------------------------------------------------------------------------
    prmColorBase MnuBx(Index).Tag
    '-------------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
        
End Sub

Public Function prmColorBase(keyConn As String)
    
    For I = MnuBx.LBound To MnuBx.UBound
        If LCase(Trim(MnuBx(I).Tag)) = LCase(keyConn) Then
            Dim arrC() As String
            arrC = Split(MnuBases.Tag, "|")
            If arrC(I) <> "" Then Me.BackColor = arrC(I) Else Me.BackColor = vbButtonFace
            Exit For
        End If
    Next
    
    Frame1.BackColor = Me.BackColor
    Frame2.BackColor = Me.BackColor
    
End Function

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
   
Dim aImporte As Currency
    On Error Resume Next
    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso
    
    With vsCuota
        .Rows = 2
        .Cell(flexcpText, 1, 0) = 1
        .Cell(flexcpText, 1, 1) = Format(Now, "dd/mm/yyyy")
        
        aImporte = CCur(tImporte.Text)
        If Trim(tIva.Text) <> "" Then aImporte = aImporte + CCur(tIva.Text)
        If Trim(tCofis.Text) <> "" Then aImporte = aImporte + CCur(tCofis.Text)
        .Cell(flexcpText, 1, 2) = aImporte
        .SetFocus
        .Select .Rows - 1, 1
    End With
    
End Sub

Sub AccionModificar()

    sModificar = True
    
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso
        
End Sub

Private Sub AccionGrabar()

    If vsCuota.EditText <> "" Then Exit Sub     'Si esta editando no hay accion
    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo errGrabar
    FechaDelServidor
    
    'Veo fecha de modificacion
    cons = "Select * from Compra Where ComCodigo = " & CLng(tIDCompra.Text)
    Set RsCom = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If gFModificacion <> RsCom!ComFModificacion Then
        MsgBox "La ficha de compra ha sido modificada desde otra terminal. Para grabar vuelva a cargar los datos.", vbExclamation, "ATENCIÓN"
        RsCom.Close: Screen.MousePointer = 0: Exit Sub
    Else
        RsCom.Edit
        RsCom!ComFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsCom.Update: RsCom.Close
    End If
    
    If sModificar Then
       cons = "Delete CompraVencimiento Where CVeIDCompra = " & CLng(tIDCompra.Text)
       cBase.Execute cons
    End If
    
    cons = "Select * from CompraVencimiento Where CVeIDCompra = " & CLng(tIDCompra.Text)
    Set RsCom = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    With vsCuota
    For I = 1 To .Rows - 1
        RsCom.AddNew
        RsCom!CVeIDCompra = CLng(tIDCompra.Text)
        RsCom!CVeCuota = .Cell(flexcpText, I, 0)
        RsCom!CVeVencimiento = Format(.Cell(flexcpText, I, 1), sqlFormatoF)
        RsCom!CVeImporte = CCur(.Cell(flexcpText, I, 2))
        RsCom.Update
    Next
    
    RsCom.Close
    End With
        
    sNuevo = False: sModificar = False
    DeshabilitoIngreso
    Botones False, True, True, False, False, Toolbar1, Me
    Foco tIDCompra
    gFModificacion = gFechaServidor
    Screen.MousePointer = 0
    
    If bEsNuevo Then
        If MsgBox("Desea volver a la pantalla Ingreso de Gastos.", vbQuestion + vbYesNo, "Salir del formulario") = vbYes Then Unload Me Else bEsNuevo = False
    End If
    Exit Sub
    
errGrabar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
End Sub

Sub AccionEliminar()

    If MsgBox("Confirma eliminar los vencimientos de las cuotas", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Exit Sub
    
    On Error GoTo Error
    cons = "Delete CompraVencimiento Where CVeIDCompra = " & CLng(tIDCompra.Text)
    cBase.Execute cons
    
    DeshabilitoIngreso
    CargoCamposDesdeBD CLng(tIDCompra.Tag)
    Exit Sub
    
Error:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
End Sub

Sub AccionCancelar()

Dim aCompra As Long
    
    On Error Resume Next
    aCompra = CLng(tIDCompra.Tag)
    DeshabilitoIngreso
    LimpioFicha
    
    CargoCamposDesdeBD aCompra
    
    sNuevo = False: sModificar = False
    
End Sub

Private Sub AccionListaDeAyuda()

    On Error GoTo errAyuda
    
    If Not IsDate(tFecha.Text) And Val(tProveedor.Tag) = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    Dim aLista As New clsListadeAyuda
    Dim aSeleccionado As Long: aSeleccionado = 0
    
    cons = " Select ID_Compra = ComCodigo, Fecha = ComFecha, Proveedor = PClFantasia, Comprobante = ComSerie + Convert(char(10), ComNumero), Moneda = MonSigno , Importe = ComImporte, Comentarios = ComComentario" _
            & " from Compra, ProveedorCliente, Moneda" _
            & " Where ComProveedor = PClCodigo" _
            & " And ComMoneda = MonCodigo" & " And ComTipoDocumento = " & TipoDocumento.CompraCredito
            
    If IsDate(tFecha.Text) Then cons = cons & " And ComFecha >= '" & Format(tFecha.Text, sqlFormatoF) & "'"
    If Val(tProveedor.Tag) <> 0 Then cons = cons & " And ComProveedor = " & Val(tProveedor.Tag)
    
    cons = cons & " Order by ComFecha DESC"
    
    aLista.ActivoListaAyudaSQL cBase, cons
    Me.Refresh
    
    If IsNumeric(aLista.ItemSeleccionadoSQL) Then aSeleccionado = CLng(aLista.ItemSeleccionadoSQL)
    Set aLista = Nothing
    
    If aSeleccionado <> 0 Then LimpioFicha: CargoCamposDesdeBD aSeleccionado
    
    Screen.MousePointer = 0
    Exit Sub
        
errAyuda:
    clsGeneral.OcurrioError "Error al activar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tIDCompra_Change()
    If Val(tIDCompra.Tag) <> 0 Then Botones False, False, False, False, False, Toolbar1, Me
    tIDCompra.Tag = 0
End Sub

Private Sub tIDCompra_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Val(tIDCompra.Tag) = 0 And IsNumeric(tIDCompra.Text) Then CargoCamposDesdeBD CLng(tIDCompra.Text) Else Foco tFecha
    End If
    
End Sub

Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 And Not sNuevo And Not sModificar Then AccionListaDeAyuda
    
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tProveedor: Exit Sub
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = 0
End Sub

Private Sub tProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And Val(tProveedor.Tag) <> 0 Then AccionListaDeAyuda
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    
    On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then Foco tIDCompra: Exit Sub
        Screen.MousePointer = 11
        cons = "Select PClCodigo, PClFantasia as 'Nombre Fantasía', PClNombre as 'Razón Social' from ProveedorCliente " _
                & " Where PClNombre like '" & Trim(tProveedor.Text) & "%' Or PClFantasia like '" & Trim(tProveedor.Text) & "%'"
        
        Dim aLista As New clsListadeAyuda, mSel As Long
        mSel = aLista.ActivarAyuda(cBase, cons, 5500, 1, "Lista de Proveedores")
        Me.Refresh: DoEvents
        If mSel <> 0 Then
            tProveedor.Text = Trim(aLista.RetornoDatoSeleccionado(1))
            tProveedor.Tag = aLista.RetornoDatoSeleccionado(0)
        Else
            tProveedor.Text = ""
        End If
        Set aLista = Nothing
        Screen.MousePointer = 0
    End If
    Exit Sub
    Screen.MousePointer = 0

errBuscar:
    clsGeneral.OcurrioError "Error al procesar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select

End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    DoEvents
    If Val(tIDCompra.Tag) = 0 Then
        MsgBox "Ocurrió un error al verificar el ID de compra. Cancele y vuelva a ingresar los datos.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    With vsCuota
    Dim aSuma, aImporte As Currency: aSuma = 0: aImporte = 0
    Dim aFecha As Date
    For I = 1 To .Rows - 1
        If Not IsDate(.Cell(flexcpText, I, 1)) Then
            MsgBox "La fecha de vencimiento de la cuota no es correcta.", vbExclamation, "ATENCIÓN"
            .Select I, 1: Exit Function
        End If
        
        If I = 1 Then
            aFecha = .Cell(flexcpText, I, 1)
        Else
            If aFecha >= CDate(.Cell(flexcpText, I, 1)) Then
                MsgBox "La fecha de vencimiento de las cuotas no son correctas (verifique el orden ascendente de los vencimientos).", vbExclamation, "ATENCIÓN"
                .Select I, 1: Exit Function
            End If
            aFecha = .Cell(flexcpText, I, 1)
        End If
                
        aSuma = aSuma + CCur(.Cell(flexcpText, I, 2))
        
    Next
    aImporte = CCur(tImporte.Text)
    If Trim(tIva.Text) <> "" Then aImporte = aImporte + CCur(tIva.Text)
    If Trim(tCofis.Text) <> "" Then aImporte = aImporte + CCur(tCofis.Text)
    
    If aImporte <> aSuma Then
        MsgBox "El importe del comprobante (" & Format(aImporte, FormatoMonedaP) & ") no coincide con la suma de las cuotas (" & Format(aSuma, FormatoMonedaP) & ")", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    End With
    
    ValidoCampos = True
    
End Function

Private Sub DeshabilitoIngreso()
        
    tProveedor.Enabled = True: tProveedor.BackColor = Blanco
    tFecha.Enabled = True: tFecha.BackColor = Blanco
    tIDCompra.Enabled = True: tIDCompra.BackColor = Blanco
        
    cMoneda.Enabled = False: cMoneda.BackColor = Blanco
    cComprobante.Enabled = False: cComprobante.BackColor = Blanco
    tNumero.Enabled = False: tNumero.BackColor = Blanco
    tSerie.Enabled = False: tSerie.BackColor = Blanco
    tImporte.Enabled = False: tImporte.BackColor = Blanco
    tIva.Enabled = False: tIva.BackColor = Blanco
    tCofis.Enabled = False: tCofis.BackColor = Blanco
            
    vsCuota.BackColor = Inactivo
    vsCuota.Editable = False
    
End Sub

Private Sub HabilitoIngreso()

    tProveedor.Enabled = False: tProveedor.BackColor = Inactivo
    tFecha.Enabled = False: tFecha.BackColor = Inactivo
    tIDCompra.Enabled = False: tIDCompra.BackColor = Inactivo
    
    vsCuota.BackColor = Blanco
    vsCuota.Editable = True
    
End Sub

Private Sub CargoCamposDesdeBD(aCompra As Long)
    
    On Error GoTo errCargo
    If aCompra = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    LimpioFicha
    
    cons = "Select * from Compra, ProveedorCliente " _
               & " Where ComCodigo = " & aCompra _
               & " And ComProveedor = PClCodigo"
    Set RsCom = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        
    If RsCom.EOF Then
        MsgBox "No existe una compra para el código ingresado.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        RsCom.Close: Exit Sub
    End If
    
    tIDCompra.Text = Format(RsCom!ComCodigo, "#,###,##0")
    tIDCompra.Tag = RsCom!ComCodigo
    tFecha.Text = Format(RsCom!ComFecha, FormatoFP)
    tProveedor.Text = Trim(RsCom!PClFantasia)
    tProveedor.Tag = RsCom!ComProveedor
    
    BuscoCodigoEnCombo cComprobante, RsCom!ComTipoDocumento
    If Not IsNull(RsCom!ComSerie) Then tSerie.Text = Trim(RsCom!ComSerie)
    If Not IsNull(RsCom!ComNumero) Then tNumero.Text = Trim(RsCom!ComNumero)
    
    BuscoCodigoEnCombo cMoneda, RsCom!ComMoneda
    Dim aBruto As Currency
    tImporte.Text = Format(RsCom!ComImporte, FormatoMonedaP): aBruto = Format(RsCom!ComImporte, FormatoMonedaP)
    If Not IsNull(RsCom!ComIva) Then tIva.Text = Format(RsCom!ComIva, FormatoMonedaP): aBruto = aBruto + Format(RsCom!ComIva, FormatoMonedaP)
    If Not IsNull(RsCom!ComCofis) Then tCofis.Text = Format(RsCom!ComCofis, FormatoMonedaP): aBruto = aBruto + Format(RsCom!ComCofis, FormatoMonedaP)
    
    
    lTotalGasto.Caption = Format(aBruto, FormatoMonedaP)
    
    gFModificacion = RsCom!ComFModificacion

    RsCom.Close
    
    'Cargo los vencimientos de las cuotas
    cons = "Select * from CompraVencimiento Where CVeIDCompra = " & tIDCompra.Tag
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    With vsCuota
    
    If Not rsAux.EOF Then
        .Rows = 1
        Do While Not rsAux.EOF
            .AddItem rsAux!CVeCuota
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!CVeVencimiento, "dd/mm/yyyy")
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!CVeImporte, FormatoMonedaP)
            rsAux.MoveNext
        Loop
        Botones False, True, True, False, False, Toolbar1, Me
    Else
        If cComprobante.ItemData(cComprobante.ListIndex) = TipoDocumento.CompraCredito Then Botones True, False, False, False, False, Toolbar1, Me Else Botones False, False, False, False, False, Toolbar1, Me
    End If
    
    End With
    
    rsAux.Close
    Screen.MousePointer = 0
    Exit Sub

errCargo:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de la compra.", Err.Description
End Sub

Private Sub LimpioFicha()
    lTotalGasto.Caption = ""
    tIDCompra.Text = ""
    tFecha.Text = ""
    tProveedor.Text = ""
    
    cMoneda.Text = "": tImporte.Text = "": tIva.Text = "": tCofis.Text = ""
    cComprobante.Text = "": tSerie.Text = "": tNumero.Text = ""
    vsCuota.Rows = 1: vsCuota.Rows = 1
    
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsCuota
        .Rows = 1: .Cols = 1
        .Editable = False
        .FormatString = "Nº Cuota|Vencimiento|>Importe|"
            
        .WordWrap = True
        .ColWidth(0) = 1000: .ColWidth(1) = 1200: .ColWidth(2) = 1400
        .ColAlignment(0) = flexAlignLeftCenter
        .ColDataType(1) = flexDTDate: .ColDataType(2) = flexDTCurrency
    End With
    
End Sub

Private Sub vsCuota_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim Suma As Currency: Suma = 0
Dim Diferencia As Currency

    With vsCuota
    If Row = .Rows - 1 And Col = 2 Then
        'Si todavia no se llego al importe agrego otra columna
        For I = 1 To .Rows - 1
            Suma = Suma + .Cell(flexcpValue, I, 2)
        Next
        Diferencia = CCur(tImporte.Text) - Suma
        If Trim(tIva.Text) <> "" Then Diferencia = Diferencia + CCur(tIva.Text)
        If Trim(tCofis.Text) <> "" Then Diferencia = Diferencia + CCur(tCofis.Text)
        If Diferencia > 0 Then
            .AddItem Trim(Str(.Cell(flexcpValue, .Rows - 1, 0) + 1))
            .Cell(flexcpText, .Rows - 1, 2) = Format(Diferencia, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 1) = Format(CDate(.Cell(flexcpText, .Rows - 2, 1)) + 30, "dd/mm/yyyy")
            .Select .Rows - 1, 1
        End If
    End If
    End With
    
End Sub

Private Sub vsCuota_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If Col = 0 Or Col = 3 Then Cancel = True
    
End Sub

Private Sub vsCuota_KeyDown(KeyCode As Integer, Shift As Integer)

    With vsCuota
    If KeyCode = vbKeyDelete Then
        If .Row = .Rows - 1 And .Row > 1 Then
            Dim aValor As Currency
            aValor = CCur(.Cell(flexcpText, .Rows - 2, 2)) + CCur(.Cell(flexcpText, .Rows - 1, 2))
            .Cell(flexcpText, .Rows - 2, 2) = Format(aValor, FormatoMonedaP)
            
            .RemoveItem .Row
            
        End If
    End If
    End With
    
End Sub

Private Sub vsCuota_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With vsCuota
    
    Select Case Col
        Case 1:
            If Not IsDate(.EditText) Then
                MsgBox "Debe ingresar la fecha de vencimiento para la cuota.", vbExclamation, "ATENCIÓN"
                .Select Row, 1: Cancel = True
            Else
                .EditText = Format(.EditText, "dd/mm/yyyy")
            End If
                        
        Case 2:
            If Not IsDate(.Cell(flexcpText, Row, 1)) Then
                MsgBox "Debe ingresar la fecha de vencimiento para la cuota.", vbExclamation, "ATENCIÓN"
                .Select Row, 1: Cancel = True: Exit Sub
            End If
            
            If Not IsNumeric(.EditText) Then
                'MsgBox "El importe ingresado no es correcto.", vbExclamation, "ATENCIÓN"
                .Select Row, 2: Cancel = True: Exit Sub
            End If
            If CCur(.EditText) = 0 Or CCur(.EditText) < 0 Then Cancel = True: Exit Sub
            .EditText = Format(.EditText, FormatoMonedaP)
    End Select
    
    End With
    
End Sub
