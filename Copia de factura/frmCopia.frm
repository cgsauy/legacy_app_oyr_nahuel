VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Begin VB.Form frmCopia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copia de Facturas"
   ClientHeight    =   3930
   ClientLeft      =   3765
   ClientTop       =   5160
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCopia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6135
   Begin VB.CommandButton bCFG 
      Caption         =   "Impresoras"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   1035
   End
   Begin VB.CommandButton bImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4980
      Picture         =   "frmCopia.frx":0442
      TabIndex        =   4
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Documento"
      ForeColor       =   &H00000080&
      Height          =   3315
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   5895
      Begin VB.TextBox tNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   1
         Top             =   540
         Width           =   1335
      End
      Begin MSMask.MaskEdBox tRuc 
         Height          =   285
         Left            =   1020
         TabIndex        =   3
         Top             =   900
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   0
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
      Begin VB.Label lblDoc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2400
         TabIndex        =   23
         Top             =   540
         Width           =   3375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " INFORMACIÓN DEL DOCUMENTO"
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
         Left            =   180
         TabIndex        =   21
         Top             =   240
         Width           =   5595
      End
      Begin VB.Label lComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   180
         TabIndex        =   19
         Top             =   2580
         UseMnemonic     =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentarios:"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Número:"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&R.U.C.:"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label labDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   16
         Top             =   1980
         Width           =   4695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Emisión:"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   1280
         Width           =   735
      End
      Begin VB.Label labEmision 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   3480
         TabIndex        =   14
         Top             =   1260
         Width           =   2295
      End
      Begin VB.Label labCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   13
         Top             =   1620
         Width           =   4695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label111 
         BackStyle       =   0  'Transparent
         Caption         =   "Digitador:"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lUsuario 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Número"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "C.I.:"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label lCI 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   7
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label lVendedor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Emisión"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5040
         TabIndex        =   6
         Top             =   900
         Width           =   615
      End
   End
   Begin VSPrinter8LibCtl.VSPrinter vspPrinter 
      Height          =   2295
      Left            =   2040
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   3495
      _cx             =   6165
      _cy             =   4048
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   8.82352941176471
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VSReport8LibCtl.VSReport vsrReport 
      Left            =   1440
      Top             =   3720
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
   Begin VB.Menu MnuPrinter 
      Caption         =   "MnuPrinter"
      Visible         =   0   'False
      Begin VB.Menu MnuDondeImprimo 
         Caption         =   "¿Dónde imprimo?"
      End
      Begin VB.Menu MnuPrintConfig 
         Caption         =   "Configurar Impresoras"
      End
      Begin VB.Menu MnuPrintLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrintOpt 
         Caption         =   "MnuPrintOpt"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmCopia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TAG
    'TSerie = Nro. Documento
    'TNumero = Codigo de Moneda
    'TRuc = Cód. de Cliente
    'labcliente = CI
    'Label1 = Monto total de un documento. (Se usa para Crédito)

Option Explicit
Public prmIDDocumento As Long
Private idDocSeleccionado As Long
Private tipoDocSeleccionado As TipoDocumento


'Variables para Crystal Engine.---------------------------------
Private result As Integer, JobSRep1 As Integer, JobSRep2 As Integer, jobnum As Integer
Private NombreFormula As String, CantForm As Integer, aTexto As String

Private Sub bImprimir_Click()
Dim aResultado As Integer
    
    If Val(tNumero.Tag) = 0 Then
        MsgBox "Debe seleccionar el documento para realizar la copia.", vbExclamation, "Posible Error "
        Exit Sub
    End If
    
    aResultado = MsgBox("Comfima emitir la copia del documento seleccionado." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
            & "Si- Imprimir el documento con el comentario." & Chr(vbKeyReturn) _
            & "No- Imprimir el documento sin el comentario." & Chr(vbKeyReturn) _
            & "Cancelar- No imprime. ", vbQuestion + vbYesNoCancel, "Copia del Documento")
            
    If aResultado = vbCancel Then Exit Sub
    
    Screen.MousePointer = 11
    ImprimoVSReport (oCnfgPrint.Opcion = 1), (aResultado = vbYes)
    If Not crCierroTrabajo(jobnum) Then MsgBox crMsgErr
    
    DeshabilitoImpresion
    Screen.MousePointer = 0
    
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    Me.Refresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

    'On Error Resume Next
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = 4545
    
    zfn_LoadMenuOpcionPrint
    
    DeshabilitoImpresion
    
'    With cTipoDoc
'        .AddItem RetornoNombreDocumento(TipoDocumento.Contado): .ItemData(.NewIndex) = TipoDocumento.Contado
'        .AddItem RetornoNombreDocumento(TipoDocumento.Credito): .ItemData(.NewIndex) = TipoDocumento.Credito
'        .AddItem RetornoNombreDocumento(TipoDocumento.NotaEspecial): .ItemData(.NewIndex) = TipoDocumento.NotaEspecial
'        .AddItem RetornoNombreDocumento(TipoDocumento.NotaDevolucion): .ItemData(.NewIndex) = TipoDocumento.NotaDevolucion
'        .AddItem RetornoNombreDocumento(TipoDocumento.NotaCredito): .ItemData(.NewIndex) = TipoDocumento.NotaCredito
'    End With
            
'    cSucursal.Text = ""
'    'Cargo Sucursales---------------------------------------------------------------------------
'    Cons = "Select SucCodigo, SucAbreviacion from Sucursal Where SucDContado <> null or SucDCredito <> null Order by SucAbreviacion"
'    CargoCombo Cons, cSucursal, ""
'    BuscoCodigoEnCombo cSucursal, paCodigoDeSucursal
'    '-----------------------------------------------------------------------------------------------
'
    idDocSeleccionado = 0
    
    oCnfgPrint.CargarConfiguracion cnfgAppNombreCopia, cnfgKeyTicketCopiaFactura
    crAbroEngine
    
    If prmIDDocumento <> 0 Then ProcesoActivacion
    
End Sub

Private Sub ProcesoActivacion()
On Error GoTo errAA
    
    Dim bHay As Boolean: bHay = False
    idDocSeleccionado = 0
        
    Cons = "Select DocCodigo from Documento Where DocCodigo = " & prmIDDocumento & _
               " And DocTipo In (1,2,3,4,10)"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        bHay = True
        idDocSeleccionado = RsAux("DocCodigo")
    End If
    RsAux.Close
    If idDocSeleccionado > 0 Then CargoDatosDocumento idDocSeleccionado
    Exit Sub
errAA:
    clsGeneral.OcurrioError "Error al cargar los datos del documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Screen.MousePointer = 11
    crCierroEngine
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    Screen.MousePointer = 0
    End
    
End Sub

Private Sub Label1_Click()
    Foco tNumero
End Sub

Private Sub Label4_Click()
    Foco tRuc
End Sub

Private Sub DeshabilitoImpresion()
    LimpioDocumento
End Sub

Private Sub MnuDondeImprimo_Click()
    frmDondeImprimo.Show vbModal
    oCnfgPrint.CargarConfiguracion cnfgAppNombreCopia, cnfgKeyTicketCopiaFactura
End Sub

Private Sub tNumero_Change()
On Error Resume Next
    If idDocSeleccionado > 0 Then LimpioDocumento
End Sub

Private Sub tNumero_GotFocus()
    tNumero.SelStart = 0
    tNumero.SelLength = Len(tNumero.Text)
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tNumero.Text) = "" Then Exit Sub
        If idDocSeleccionado > 0 Then Exit Sub
        idDocSeleccionado = BuscoDocumento(tNumero.Text, "1,2,3,4,10")
        If idDocSeleccionado > 0 Then CargoDatosDocumento idDocSeleccionado
    End If
    
End Sub

Private Sub tRuc_GotFocus()
    tRuc.SelStart = 0: tRuc.SelLength = Len(tRuc.Mask)
End Sub

Private Sub tRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If bImprimir.Enabled Then bImprimir.SetFocus
End Sub

Private Sub CargoDatosDocumento(ByVal idDoc As Long)
Dim RsDoc As rdoResultset
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    LimpioDocumento
               
    Cons = "Select Documento.*, TDoNombre, SucAbreviacion, CliDireccion, CliCIRUC, CPeApellido1, CPeApellido2, CPeNombre1, CPeNombre2, CPeRUC, CEmNombre, CEmFantasia " _
        & " From Documento " _
        & "INNER JOIN Sucursal ON SucCodigo = DocSucursal INNER JOIN TipoDocumento ON TDoID = DocTipo " _
        & "INNER JOIN Cliente ON DocCliente = CliCodigo " _
                & "Left Outer Join CPersona On CPeCliente = CliCodigo " _
                & "Left Outer Join CEmpresa On CEmCliente = CliCodigo " _
        & " WHERE DocCodigo = " & idDoc
                
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsDoc.EOF Then
        Screen.MousePointer = 0
        MsgBox "No existe un documento para la numeración ingresada o bien no pertenece a esta sucursal.", vbExclamation, "No Hay Datos"
    Else
        'Verifico si el Documento no fue anulado (Papel)--------------------------------------
        If RsDoc!DocAnulado Then
            Screen.MousePointer = 0
            MsgBox "El documento ingresado figura como papel anulado.", vbExclamation, "Documento Anulado"
        Else
            tNumero.Text = Trim(RsDoc("DocSerie")) & "-" & RsDoc("DocNumero")
            idDocSeleccionado = idDoc
            tipoDocSeleccionado = RsDoc("DocTipo")
            lblDoc.Caption = Trim(RsDoc("TDoNombre")) & " " & tNumero.Text & " (" & Trim(RsDoc("SucAbreviacion")) & ")"
            
            labEmision.Caption = " " & Format(RsDoc!DocFecha, "dd/mm/yy hh:mm")
            tNumero.Tag = RsDoc!DocMoneda
            Label1.Tag = RsDoc!DocTotal
            If Not IsNull(RsDoc!CliDireccion) Then labDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, RsDoc!CliDireccion)
            
            lUsuario.Caption = BuscoDigitoUsuario(RsDoc!DocUsuario)
            If Not IsNull((RsDoc!DocVendedor)) Then lVendedor.Caption = BuscoDigitoUsuario(RsDoc!DocVendedor)
            If Not IsNull(RsDoc!DocComentario) Then lComentario.Caption = Trim(RsDoc!DocComentario)
            
            'Cargo el Nombre del Cliente del Documento
            If Not IsNull(RsDoc!CPeApellido1) Then  'Es Persona.
                labCliente.Caption = ArmoNombre(Format(RsDoc!CPeApellido1, "#"), Format(RsDoc!CPeApellido2, "#"), Format(RsDoc!CPeNombre1, "#"), Format(RsDoc!CPeNombre2, "#"))
                
                If Not IsNull(RsDoc!CliCIRuc) Then
                    'Guardo en el Tag con formato ---> para formulas de impresiones
                    lCI.Caption = clsGeneral.RetornoFormatoCedula(RsDoc!CliCIRuc)
                    labCliente.Caption = Trim(labCliente.Caption)
                End If
                
                If Not IsNull(RsDoc!CPERuc) Then tRuc.Text = RsDoc!CPERuc   'RetornoFormatoRuc(RsDoc!CPERuc)
            Else
                If Not IsNull(RsDoc!CEmNombre) Then labCliente.Caption = Trim(RsDoc!CEmNombre) Else labCliente.Caption = Trim(RsDoc!CEmFantasia)
                If Not IsNull(RsDoc!CliCIRuc) Then tRuc.Text = RsDoc!CliCIRuc
            End If
            
            If tRuc.Text = "" Then tRuc.Enabled = True
            
            bImprimir.Enabled = True
        End If
    End If
    
    RsDoc.Close
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub LimpioDocumento()
    
    idDocSeleccionado = 0
    tRuc.Text = "": tRuc.Enabled = False
    lCI.Caption = ""
    lblDoc.Caption = ""
    
    labCliente.Caption = ""
    labDireccion.Caption = ""
    labEmision.Caption = ""
    lComentario.Caption = ""
    
    lUsuario.Caption = ""
    lVendedor.Caption = ""
    
    bImprimir.Enabled = False

End Sub

Private Sub ImprimoDocumento(IDDocumento As Long, Tipo As Integer, Comentario As Integer)
On Error GoTo ErrCrystal
Dim monSigno As String, monNombre As String
Dim pstrNombreDoc As String

    Screen.MousePointer = 11
    BuscoDatosMoneda Val(tNumero.Tag), monSigno, monNombre
    
    Select Case Tipo
        Case TipoDocumento.Contado: pstrNombreDoc = paDContado
        Case TipoDocumento.Credito: pstrNombreDoc = paDCredito
        Case TipoDocumento.NotaDevolucion: pstrNombreDoc = paDNDevolucion
        Case TipoDocumento.NotaCredito: pstrNombreDoc = paDNCredito
        Case TipoDocumento.NotaEspecial: pstrNombreDoc = paDNEspecial
    End Select
    
    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte Contado --------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento":
                    result = crSeteoFormula(jobnum%, NombreFormula, "'" & pstrNombreDoc & "'")
                    
            Case "cliente"
                    aTexto = Trim(labCliente.Caption)
                    If lCI.Caption <> "" Then aTexto = aTexto & " (" & Trim(lCI.Caption) & ")"
                    result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
                    
            Case "direccion": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labDireccion.Caption) & "'")
            Case "ruc":
                        If Trim(tRuc.Text) <> "" Then
                            result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(tRuc.FormattedText) & "'")
                        Else
                            result = crSeteoFormula(jobnum%, NombreFormula, "''")
                        End If
            
            'Case "codigobarras":
            '        If TipoDoc = TipoDocumento.Contado Then aTexto = CodigoDeBarras(TipoDoc, CLng(tSerie.Tag)) Else aTexto = ""
            '        Result = crSeteoFormula(JobNum%, NombreFormula, "'" & aTexto & "'")
            
            Case "signomoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'" & monSigno & "'")
            Case "nombremoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'(" & monNombre & ")'")
            
            Case "comentario": If Comentario = vbYes Then result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lComentario.Caption) & "'")
            
            Case "usuario": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lUsuario.Caption) & "'")
            Case "vendedor": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lVendedor.Caption) & "'")
            
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where DocCodigo = " & IDDocumento
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
        
    'Subreporte srContado.rpt  y srContado.rpt - 01-----------------------------------------------------------------------------
    JobSRep1 = crAbroSubreporte(jobnum, "srContado.rpt")
    If JobSRep1 = 0 Then GoTo ErrCrystal
    
    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"
    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
    
    '-------------------------------------------------------------------------------------------------------------------------------------

    'If crMandoAPantalla(jobnum, "Reimpresion Contado") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
    
    'crEsperoCierreReportePantalla
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroSubReporte JobSRep1
    Screen.MousePointer = 0
    Exit Sub
End Sub


Private Sub ImprimoContado_OLD(Comentario As Integer)
On Error GoTo ErrCrystal
Dim monSigno As String, monNombre As String

    Screen.MousePointer = 11
    BuscoDatosMoneda Val(tNumero.Tag), monSigno, monNombre
    
    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte Contado --------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento": result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDContado & "'")
            Case "cliente"
                    aTexto = Trim(labCliente.Caption)
                    If lCI.Caption <> "" Then aTexto = aTexto & " (" & Trim(lCI.Caption) & ")"
                    result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
                    
            Case "direccion": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labDireccion.Caption) & "'")
            Case "ruc":
                        If Trim(tRuc.Text) <> "" Then
                            result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(tRuc.FormattedText) & "'")
                        Else
                            result = crSeteoFormula(jobnum%, NombreFormula, "''")
                        End If
            
            'Case "codigobarras":
            '        If TipoDoc = TipoDocumento.Contado Then aTexto = CodigoDeBarras(TipoDoc, CLng(tSerie.Tag)) Else aTexto = ""
            '        Result = crSeteoFormula(JobNum%, NombreFormula, "'" & aTexto & "'")
            
            Case "signomoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'" & monSigno & "'")
            Case "nombremoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'(" & monNombre & ")'")
            
            Case "comentario": If Comentario = vbYes Then result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lComentario.Caption) & "'")
            
            Case "usuario": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lUsuario.Caption) & "'")
            Case "vendedor": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lVendedor.Caption) & "'")
            
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where DocCodigo = " & 362085 'CLng(tSerie.Tag)
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
        
    'Subreporte srContado.rpt  y srContado.rpt - 01-----------------------------------------------------------------------------
    JobSRep1 = crAbroSubreporte(jobnum, "srContado.rpt")
    If JobSRep1 = 0 Then GoTo ErrCrystal
    
    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"
    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
    
    '-------------------------------------------------------------------------------------------------------------------------------------

    If crMandoAPantalla(jobnum, "Reimpresion Contado") = 0 Then GoTo ErrCrystal
    'If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
    
    crEsperoCierreReportePantalla
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroSubReporte JobSRep1
    Screen.MousePointer = 0
    Exit Sub
End Sub

Private Function InicializoReporteEImpresora(paNImpresora As String, paBImpresora As Integer, Reporte As String) As Boolean
On Error GoTo ErrCrystal
    
    jobnum = crAbroReporte(prmPathListados & Reporte)
    If jobnum = 0 Then GoTo ErrCrystal
    
    If ChangeCnfgPrint Then prj_LoadConfigPrint bShowFrm:=False
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paNImpresora) Then SeteoImpresoraPorDefecto paNImpresora
    If Not crSeteoImpresora(jobnum, Printer, paBImpresora) Then GoTo ErrCrystal
    
    InicializoReporteEImpresora = False
    Exit Function

ErrCrystal:
    InicializoReporteEImpresora = True
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroTrabajo jobnum
    Screen.MousePointer = 0

End Function

Private Sub ImprimoCredito(Comentario As Integer)

Dim RsAuxC As rdoResultset
Dim sTexto As String, MEnvio As Currency, MEntrega As Currency
Dim sConCheques As Boolean

Dim monSigno As String, monNombre As String

    BuscoDatosMoneda Val(tNumero.Tag), monSigno, monNombre
    
    'Consulta para sacar los datos del credito------------------------------------------------------------
     Cons = "Select * from Credito, TipoCuota " _
             & " Where CreFactura = " & idDocSeleccionado _
             & " And CreTipoCuota = TCuCodigo"
    Set RsAuxC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    'Efectivo = 1    ChequeDiferido = 2
    If RsAuxC!CreFormaPago = 2 Then sConCheques = True Else sConCheques = False
    
    Cons = "Select * From Renglon " _
            & " Where RenDocumento = " & idDocSeleccionado _
            & " And (RenArticulo IN (Select Distinct(TFlArticulo) From TipoFlete)" _
            & " Or RenArticulo = " & paArticuloPisoAgencia & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF
            MEnvio = RsAux!RenCantidad * RsAux!RenPrecio
            RsAux.MoveNext
        Loop
    Else
        MEnvio = 0
    End If
    RsAux.Close
    
    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte Credito --------------------------------------------------------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento": result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDCredito & "'")
            Case "cliente": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labCliente.Caption) & "'")
            Case "cedula": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lCI.Caption) & "'")
            Case "direccion": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labDireccion.Caption) & "'")
            
            Case "ruc":
                        If Trim(tRuc.Text) <> "" Then
                            result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(tRuc.FormattedText) & "'")
                        Else
                            result = crSeteoFormula(jobnum%, NombreFormula, "''")
                        End If
            
            'Case "codigobarras": Result = crSeteoFormula(JobNum%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.Credito, CLng(tSerie.Tag)) & "'")
            Case "signomoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'" & monSigno & "'")
            Case "nombremoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'(" & monNombre & ")'")
            
            Case "garantia"
                sTexto = ""
                If Not IsNull(RsAuxC!CreGarantia) Then
                    'Cargo datos de la garantía.
                    Cons = "Select * From Cliente,  CPersona " _
                        & "Where CliCodigo  = " & RsAuxC!CreGarantia & " And CliCodigo = CPeCliente"
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                    sTexto = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc) & " " & ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
                    RsAux.Close
                End If
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & sTexto & "'")
                
            Case "nombrerecibo": result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDRecibo & "'")
            
            Case "usuario": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lUsuario.Caption) & "'")
            Case "vendedor": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lVendedor.Caption) & "'")
            
            Case "comentario": If Comentario = vbYes Then result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lComentario.Caption) & "'")
                            
            Case "reciboflete": If MEnvio > 0 Then result = crSeteoFormula(jobnum%, NombreFormula, "'" & Format(MEnvio, FormatoMonedaP) & "'")
            
            Case "recibocuota"
                If Trim(RsAuxC!CreVaCuota) <> "" Then
                    sTexto = Trim(RsAuxC!CreVaCuota) & " de " & Trim(RsAuxC!CreDeCuota)
                    result = crSeteoFormula(jobnum%, NombreFormula, "'" & sTexto & "'")
                End If
                
            Case "financiacion"
                sTexto = Trim(RsAuxC!TCuAbreviacion) & " - "
                If Not IsNull(RsAuxC!TCuVencimientoE) Then
                    MEntrega = CCur(Label1.Tag) - (RsAuxC!TCuCantidad * RsAuxC!CreValorCuota)
                    If MEntrega > 0 Then sTexto = sTexto & "Ent.: " & Format(MEntrega, FormatoMonedaP) & " "
                End If
                sTexto = sTexto & RsAuxC!TCuCantidad & " x " & Format(RsAuxC!CreValorCuota, FormatoMonedaP)
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & sTexto & "'")
                    
            Case "proximovto"
                If Not IsNull(RsAuxC!CreProximoVto) Then
                    sTexto = Format(RsAuxC!CreProximoVto, "d Mmm yyyy")
                    result = crSeteoFormula(jobnum%, NombreFormula, "'" & sTexto & "'")
                End If
            
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    RsAuxC.Close
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    'Cons = "SELECT Top 1 Documento.DocCodigo, Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal," _
            & " Documento.DocIVA, Documento.DocVendedor, Recibo.DocSerie , Recibo.DocNumero, Recibo.DocTotal" _
            & " From " _
            & " { oj (" & paBD & ".dbo.Documento Documento " _
                        & " LEFT OUTER JOIN " & paBD & ".dbo.DocumentoPago DocumentoPago ON  Documento.DocCodigo = DocumentoPago.DPaDocASaldar)" _
                        & " LEFT OUTER JOIN " & paBD & ".dbo.Documento Recibo ON  DocumentoPago.DPaDocQSalda = Recibo.DocCodigo}" _
            & " Where Documento.DocCodigo = " & CLng(tSerie.Tag)
    
    'If sConCheques Then
    '    Cons = Cons & " Group by Documento.DocCodigo, Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor, Recibo.DocSerie , Recibo.DocNumero, Recibo.DocTotal, Documento.DocComentario "
    'End If
    
    Cons = "SELECT Documento.DocCodigo, Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal," _
            & " Documento.DocIVA, Documento.DocVendedor" _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where Documento.DocCodigo = " & idDocSeleccionado
    
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
        
    'Subreporte srCredito.rpt  ------------------------------------------------------------------------------------------------------
    JobSRep1 = crAbroSubreporte(jobnum, "srCredito.rpt - 01")
    If JobSRep1 = 0 Then GoTo ErrCrystal
    
    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"
        
    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
    
    '-------------------------------------------------------------------------------------------------------------------------------------
    
    'If crMandoAPantalla(jobnum, "Factura Credito") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
            
    'crEsperoCierreReportePantalla
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroSubReporte JobSRep1
    Screen.MousePointer = 0
End Sub

Function ArmoNombre(Ape1 As String, Ape2 As String, Nom1 As String, Nom2 As String) As String

    ArmoNombre = Trim(Ape1) & " " & Trim(Ape2)
    ArmoNombre = Trim(ArmoNombre) & ", " & Trim(Nom1) & " " & Trim(Nom2)
    
End Function

Function BuscoDigitoUsuario(Codigo As Long) As String
On Error GoTo ErrBU
Dim Rs As rdoResultset

    BuscoDigitoUsuario = ""

    Cons = "Select * from Usuario Where UsuCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then BuscoDigitoUsuario = Trim(Rs!UsuDigito)
    Rs.Close
    Exit Function
    
ErrBU:
End Function


Private Function BuscoDatosMoneda(idMoneda As Long, Optional Signo As String = "", Optional Nombre As String = "") As Boolean
    On Error GoTo ErrBU
    
    Dim Rs As rdoResultset
    BuscoDatosMoneda = True

    Cons = "Select * From Moneda Where MonCodigo = " & idMoneda
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then
        Signo = Trim(Rs!monSigno)
        Nombre = Trim(Rs!monNombre)
    End If
    Rs.Close
    Exit Function
    
ErrBU:
    BuscoDatosMoneda = False
End Function


Private Sub zfn_LoadMenuOpcionPrint()
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

Private Sub bCFG_Click()
    PopupMenu MnuPrinter, , bCFG.Left + 60, bCFG.Top + bCFG.Height + 60
End Sub

Private Sub MnuPrintConfig_Click()
On Error Resume Next
    
    prj_LoadConfigPrint True
    
    Dim iQ As Integer
    For iQ = MnuPrintOpt.LBound To MnuPrintOpt.UBound
        MnuPrintOpt(iQ).Checked = False
        MnuPrintOpt(iQ).Checked = (MnuPrintOpt(iQ).Caption = paOptPrintSel)
    Next
    
End Sub

Private Sub MnuPrintOpt_Click(Index As Integer)
On Error GoTo errLCP
Dim objPrint As New clsCnfgPrintDocument
Dim sPrint As String
Dim vPrint() As String
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
    
    Exit Sub
errLCP:
    MsgBox "Error al setear los datos de configuración: " & Err.Description, vbExclamation, "ATENCIÓN"
End Sub

Private Function DatosDocumentoCredito(ByVal Documento As Long, ByVal esTicket As Boolean, ByVal Comentario As Boolean) As String
Dim sQy As String
Dim rsCr As rdoResultset
Dim sFinancProxVto As String, sAbal As String
    
    sQy = "SELECT TCuAbreviacion, TCuAbreviacion, TCuVencimientoE, TCuCantidad, CreValorCuota, CreProximoVto, CreGarantia " & _
               "FROM Credito INNER JOIN TipoCuota ON CreTipoCuota = TCuCodigo " _
             & "WHERE CreFactura = " & Documento
    Set rsCr = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurValues)
    If Not rsCr.EOF Then
            
        sFinancProxVto = Trim(rsCr("TCuAbreviacion")) & " - "
        If Not IsNull(rsCr!TCuVencimientoE) Then
            If CCur(Label1.Tag) - (rsCr!TCuCantidad * rsCr!CreValorCuota) > 0 Then
                sFinancProxVto = sFinancProxVto & "Ent.: " & Format(CCur(Label1.Tag) - (rsCr!TCuCantidad * rsCr!CreValorCuota), FormatoMonedaP) & " "
            End If
        End If
        sFinancProxVto = sFinancProxVto & rsCr!TCuCantidad & " x " & Format(rsCr!CreValorCuota, FormatoMonedaP)

        If Not IsNull(rsCr!CreProximoVto) Then
            sFinancProxVto = sFinancProxVto & vbCrLf & "Próximo vencimiento: " & Format(rsCr!CreProximoVto, "d Mmm yyyy")
        End If

        If Not IsNull(rsCr!CreGarantia) Then
            'Cargo datos de la garantía.
            sQy = "Select CliCIRuc, CPeApellido1, CPeApellido2, CPeNombre1, CPeNombre2 From Cliente,  CPersona " _
                & "Where CliCodigo  = " & rsCr!CreGarantia & " And CliCodigo = CPeCliente"
            Set RsAux = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurReadOnly)
            If Not RsAux.EOF Then
                sAbal = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc) & " " & ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
            End If
            RsAux.Close
        End If

    End If
    rsCr.Close
    If esTicket Then
        
        DatosDocumentoCredito = QueryTicket(Documento, sAbal, sFinancProxVto, Comentario)
    Else
        DatosDocumentoCredito = QueryA5(Documento, sAbal, sFinancProxVto, Comentario)
    End If

End Function

Function QueryTicket(ByVal docCodigo As Long, ByVal Garantia As String, ByVal DatosCredito As String, ByVal Comentario As Boolean) As String

    QueryTicket = "SELECT dbo.FormatDate(DocFecha, 21) FechaHora " & _
            ", dbo.NombreDeDocumento(DocCodigo) + ' ' + RTRIM(DocSerie) + '-' + CONVERT(varchar(7), DocNumero) NombreDocumento " & _
            ", dbo.FormatNumber(DocTotal,2) Total " & _
            ", dbo.FormatNumber(DocIVA,2) IVA " & _
            ", dbo.FormatNumber(DocTotal - DocIVA, 2) Neto " & _
            ", RTRIM(MonNombre) NombreMoneda " & _
            ", '" & IIf(tRuc.Text = "", IIf(lCI.Caption <> "", lCI.Caption & vbCrLf, ""), tRuc.FormattedText & vbCrLf) & labCliente.Caption & "' Cliente " & _
            ", '" & labDireccion.Caption & "' Domicilio " & _
            ", '" & Garantia & "' Garantia " & _
            ", '" & DatosCredito & "' DatosCredito " & _
            ", '" & IIf(Comentario, lComentario.Caption, "") & "' Comentario " & _
            ", UsuDigito Usuario " & _
            "FROM Documento INNER JOIN Moneda ON DocMoneda = MonCodigo " & _
            "INNER JOIN Usuario ON DocUsuario = UsuCodigo WHERE DocCodigo = " & docCodigo

End Function

Function QueryA5(ByVal docCodigo As Long, ByVal Garantia As String, ByVal DatosCredito As String, ByVal Comentario As Boolean) As String

    QueryA5 = "SELECT dbo.FormatDate(DocFecha, 21) FechaHora " & _
            ", dbo.NombreDeDocumento(DocCodigo) + ' ' + RTRIM(DocSerie) + '-' + CONVERT(varchar(7), DocNumero) NombreDocumento " & _
            ", dbo.FormatNumber(DocTotal,2) Total " & _
            ", dbo.FormatNumber(DocIVA,2) IVA " & _
            ", dbo.FormatNumber(DocTotal - DocIVA, 2) Neto " & _
            ", RTRIM(MonNombre) NombreMoneda " & _
            ", '" & IIf(tRuc.Text = "", "Consumo final", tRuc.FormattedText) & "' RUTConsumoFinal " & _
            ", '" & IIf(lCI.Caption <> "", "(" & lCI.Caption & ") ", "") & labCliente.Caption & "' Cliente " & _
            ", '" & labDireccion.Caption & "' Domicilio " & _
            ", '" & Garantia & "' Garantia " & _
            ", '" & DatosCredito & "' DatosCredito " & _
            ", '" & IIf(Comentario, lComentario.Caption, "") & "' Comentario " & _
            ", UsuDigito Usuario " & _
            "FROM Documento INNER JOIN Moneda ON DocMoneda = MonCodigo " & _
            "INNER JOIN Usuario ON DocUsuario = UsuCodigo WHERE DocCodigo = " & docCodigo

End Function

Private Sub ImprimoVSReport(ByVal ticket As Boolean, ByVal Comentario As Boolean)
On Error GoTo errIT

    If ticket Then
        vspPrinter.Device = oCnfgPrint.ImpresoraTickets
    Else
        vspPrinter.Device = paIConformeN
        vspPrinter.PaperBin = paIConformeB
        vspPrinter.paperSize = paIConformePS
    End If
    
    With vsrReport
        .Clear                  ' clear any existing fields
        .FontName = "Tahoma"    ' set default font for all controls
        .FontSize = 8
        
        .Load prmPathListados & IIf(ticket, "CopiaFacturaticket.xml", "CopiaFactura.xml"), "Factura"
    
        .DataSource.ConnectionString = cBase.Connect
        
        If tipoDocSeleccionado <> Credito Then
            If ticket Then
                .DataSource.RecordSource = QueryTicket(idDocSeleccionado, "", "", Comentario)
            Else
                .DataSource.RecordSource = QueryA5(idDocSeleccionado, "", "", Comentario)
            End If
        Else
            .DataSource.RecordSource = DatosDocumentoCredito(idDocSeleccionado, ticket, Comentario)
        End If

        
        .Fields("Renglones").Subreport.DataSource.ConnectionString = cBase.Connect
        .Fields("Renglones").Subreport.DataSource.RecordSource = "SELECT RTRIM(ArtNombre) NomArticulo, Convert(varchar(4), RenCantidad) + ' x ' + dbo.FormatNumber(RenPrecio, 2) QArticulo, RenPrecio, RenCantidad, dbo.FormatNumber(RenPrecio * RenCantidad, 2) TotalRenglon " _
                                    & " FROM Renglon INNER JOIN Articulo ON RenArticulo = ArtID " _
                                    & " WHERE RenDocumento = " & idDocSeleccionado
        
        .Render vspPrinter
    End With
    
    
    vspPrinter.PrintDoc False
    Exit Sub
errIT:
    clsGeneral.OcurrioError "Error al imprimir.", Err.Description, "Impresión en ticket"
End Sub

Private Function BuscoDocumento(ByVal filtroBuscar As String, ByVal tiposDocumento As String) As Long
On Error GoTo errBD
    Screen.MousePointer = 11
    Dim sSerie As String, iNumero As Long
    Dim iQ As Integer, iCodigo As Long
    Dim sQy As String
    
    If InStr(1, filtroBuscar, "D", vbTextCompare) > 1 Then
        iCodigo = Val(Mid(filtroBuscar, InStr(1, filtroBuscar, "D", vbTextCompare) + 1))
        sQy = " WHERE DocCodigo = " & iCodigo & " AND DocTipo IN (" & tiposDocumento & ")"
    Else
        If InStr(filtroBuscar, "-") <> 0 Then
            sSerie = Mid(filtroBuscar, 1, InStr(filtroBuscar, "-") - 1)
            iNumero = Val(Mid(filtroBuscar, InStr(filtroBuscar, "-") + 1))
        Else
            filtroBuscar = Replace(filtroBuscar, " ", "")
            If IsNumeric(Mid(filtroBuscar, 2, 1)) Then
                sSerie = Mid(filtroBuscar, 1, 1)
                iNumero = Val(Mid(filtroBuscar, 2))
            Else
                sSerie = Mid(filtroBuscar, 1, 2)
                iNumero = Val(Mid(filtroBuscar, 3))
            End If
        End If
        sQy = " WHERE DocSerie = '" & sSerie & "' AND DocNumero = " & iNumero & " AND DocTipo IN (" & tiposDocumento & ")"
    End If
    sQy = "SELECT DocCodigo, DocFecha as Fecha" & _
        ", rtrim(TDoNombre) Documento " & _
        ", rTrim(DocSerie) + '-' + rtrim(Convert(Varchar(6), DocNumero)) as Número" & _
        " FROM Documento INNER JOIN TipoDocumento ON DocTipo = TDoId " & sQy
    sQy = sQy & " Order by DocFecha DESC"
    
    iCodigo = 0
    Dim RsDoc As rdoResultset
    Set RsDoc = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not RsDoc.EOF Then
        iCodigo = RsDoc("DocCodigo")
        iQ = 1
        RsDoc.MoveNext: If Not RsDoc.EOF Then iQ = 2
    End If
    RsDoc.Close
    
    Select Case iQ
        Case 2
            Dim miLDocs As New clsListadeAyuda
            iCodigo = miLDocs.ActivarAyuda(cBase, sQy, 6100, 1)
            Me.Refresh
            If iCodigo > 0 Then iCodigo = miLDocs.RetornoDatoSeleccionado(0)
            Set miLDocs = Nothing
    End Select
    BuscoDocumento = iCodigo
    Screen.MousePointer = 0
    Exit Function
errBD:
    MsgBox "Error al buscar el documento: " & Err.Description, vbCritical, "Buscar documento"
    Screen.MousePointer = 0
End Function
