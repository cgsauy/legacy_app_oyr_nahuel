VERSION 5.00
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReImpresion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reimpresión de Documentos"
   ClientHeight    =   4200
   ClientLeft      =   2685
   ClientTop       =   2265
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReImpresion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5415
   Begin VSPrinter8LibCtl.VSPrinter vspPrinter 
      Height          =   2295
      Left            =   360
      TabIndex        =   33
      Top             =   2520
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
      Zoom            =   9.375
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
   Begin VB.CommandButton bCFG 
      Caption         =   "Impresoras"
      Height          =   375
      Left            =   60
      TabIndex        =   31
      Top             =   3780
      Width           =   1035
   End
   Begin VB.CommandButton bOperacion 
      Caption         =   "O&peración"
      Height          =   375
      Left            =   3180
      TabIndex        =   15
      Top             =   3780
      Width           =   1035
   End
   Begin VB.CommandButton bImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   3780
      Width           =   1035
   End
   Begin VB.TextBox tNumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   1740
      Width           =   975
   End
   Begin VB.TextBox tSerie 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   780
      TabIndex        =   10
      Top             =   1740
      Width           =   435
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipos de Documentos"
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   5295
      Begin VB.OptionButton opDocumento 
         Caption         =   "Nota de Dé&bito"
         Enabled         =   0   'False
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   1395
      End
      Begin VB.OptionButton opDocumento 
         Caption         =   "Con&forme"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   8
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton opDocumento 
         Caption         =   "Remi&to"
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton opDocumento 
         Caption         =   "Reci&bo"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   5
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton opDocumento 
         Caption         =   "Nota &Especial"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton opDocumento 
         Caption         =   "No&ta de Crédito"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   4
         Top             =   540
         Width           =   1575
      End
      Begin VB.OptionButton opDocumento 
         Caption         =   "Not&a de Devolución"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton opDocumento 
         Caption         =   "Cré&dito"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton opDocumento 
         Caption         =   "C&ontado"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSMask.MaskEdBox tRuc 
      Height          =   285
      Left            =   1020
      TabIndex        =   13
      Top             =   2220
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
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
   Begin VB.Label lNCDevuelve 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SE DEVUELVE "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3300
      TabIndex        =   32
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lCI 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1020
      TabIndex        =   30
      Top             =   2580
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "C.I.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   29
      Top             =   2580
      Width           =   855
   End
   Begin VB.Label lUsuario 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Número"
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
      Left            =   3540
      TabIndex        =   28
      Top             =   2220
      Width           =   375
   End
   Begin VB.Label lVendedor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión"
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
      Left            =   4860
      TabIndex        =   27
      Top             =   2220
      Width           =   375
   End
   Begin VB.Label Label111 
      BackStyle       =   0  'Transparent
      Caption         =   "Digitador:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2820
      TabIndex        =   26
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label Label60 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4140
      TabIndex        =   25
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   24
      Top             =   2940
      Width           =   855
   End
   Begin VB.Label labCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1020
      TabIndex        =   23
      Top             =   2940
      Width           =   4215
   End
   Begin VB.Label labEmision 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   3420
      TabIndex        =   22
      Top             =   1740
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión"
      Height          =   255
      Left            =   2700
      TabIndex        =   21
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label labDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1020
      TabIndex        =   20
      Top             =   3300
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   19
      Top             =   3300
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&R.U.C.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   2220
      Width           =   615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   60
      TabIndex        =   18
      Top             =   2100
      Width           =   5295
   End
   Begin VB.Label Label2 
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
      Left            =   60
      TabIndex        =   17
      Top             =   1380
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Número"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.Menu MnuPrinter 
      Caption         =   "MnuPrinter"
      Visible         =   0   'False
      Begin VB.Menu MnuPrintConformes 
         Caption         =   "¿Dónde imprimo conformes?"
      End
      Begin VB.Menu MnuPrintSalidasCaja 
         Caption         =   "¿Dónde imprimo salidas de caja?"
      End
      Begin VB.Menu MnuPrintConfig 
         Caption         =   "Configurar Impresoras"
      End
      Begin VB.Menu MnuPrintLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrintOpt 
         Caption         =   "MnuPrinter"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmReImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'TSerie = Nro. Documento
    'TNumero = Codigo de Moneda
    'TRuc = Cód. de Cliente
    'labcliente = CI
    'Frame1 = Tipo de Documento.
    'Label6 = serie y nro. del documento al cual se le hizo un remito.
    'Label1 = Monto total de un documento. (Se usa para Crédito)
Option Explicit
Private Seleccionado As Integer
Dim gSucesoUsr As Long, gSucesoDef As String

'Variables para Crystal Engine.---------------------------------
Private result As Integer, JobSRep1 As Integer, JobSRep2 As Integer, jobnum As Integer, jobCredito As Integer
Private NombreFormula As String, CantForm As Integer, aTexto As String

Public prmIDDocumento As Long

Dim aMSigno As String, aMNombre As String

Private Function fnc_3Vias() As Boolean
    fnc_3Vias = (InStr(1, "," & prmLocal3Vias & ",", "," & paCodigoDeSucursal & ",") > 0)
End Function

Private Function fnc_PrintDocumento(ByVal iDoc As Long) As Boolean
On Error GoTo errMPD
    Dim oPrint As New clsPrintManager
    With oPrint
        .SetDevice paIContadoN, paIContadoB, paPrintCtdoPaperSize
        If .LoadFileData(prmPathListados & "rptRemitoEnvio.txt") Then
            fnc_PrintDocumento = .PrintDocumento("Exec prg_DistribuirEnvio_PrintRemitoCtdo " & iDoc, vspPrinter)
        End If
    End With
    Set oPrint = Nothing
    Exit Function
errMPD:
    clsGeneral.OcurrioError "Error al imprimir el documento de código: " & iDoc, Err.Description, "Impresión de documentos"
End Function

Private Sub bImprimir_Click()
On Error GoTo errImprimir
    
    FechaDelServidor
    
    If DateDiff("n", CDate(labEmision.Caption), gFechaServidor) > 30 Then
        If MsgBox("Si el documento no es reciente, NO SE DEBERÍA reimprimir." & vbCr & _
                    "Desea hacerlo igualmente ? " & vbCr & vbCr & _
                    "ATENCIÓN: no confundir Reimprimir con Copia de Factura", vbExclamation + vbYesNo + vbDefaultButton2, "NO ES RECIENTE") = vbNo Then Exit Sub
                    
    End If

    Screen.MousePointer = 11
    
    If Seleccionado <> 7 Then       'Suceso
        On Error Resume Next
        Screen.MousePointer = 11
        Dim objSuceso As New clsSuceso
        objSuceso.ActivoFormulario paCodigoDeUsuario, "Reimpresión de Documentos", cBase
        gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
        gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
        Set objSuceso = Nothing
        Me.Refresh
        If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub 'Abortó el ingreso del suceso
        '---------------------------------------------------------------------------------------------
                
        Dim aNDocumento As String
        aNDocumento = opDocumento(Seleccionado).Caption
        aNDocumento = Trim(Mid(aNDocumento, 1, InStr(aNDocumento, "&") - 1) & Mid(aNDocumento, InStr(aNDocumento, "&") + 1, Len(aNDocumento)))
                
        clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.Reimpresiones, paCodigoDeTerminal, gSucesoUsr, CLng(tSerie.Tag), _
                                   Descripcion:=aNDocumento, Defensa:=Trim(gSucesoDef)
    End If
    Me.Refresh
    
    Select Case Seleccionado
        Case 0      'CONTADO
            'Primero abrimos el reporte y seteamos la impresora si da error devuelve true
            If fnc_3Vias Then
                fnc_PrintDocumento CLng(tSerie.Tag)
            Else
                'If InicializoReporteEImpresora(paIContadoN, paIContadoB, "Contado.RPT") Then Exit Sub
                'ImprimoNotasYContado paDContado, TipoDocumento.Contado, ""
                ImprimirContadoYNotas_VSReport paDContado, CLng(tSerie.Tag), TipoDocumento.Contado, "", ""
            End If
            
        Case 1      'CREDITO
            'If InicializoReporteEImpresoraCredito(paICreditoN, paICreditoB, "Credito.RPT") Then Exit Sub
            'ImprimoCredito
            'crCierroTrabajo jobCredito
            ImprimoCredito_VSReport
            
        
        Case 2      'Nota de Devolucion
            'Primero abrimos el reporte y seteamos la impresora si da error devuelve true
            If InicializoReporteEImpresora(paIContadoN, paIContadoB, "NotaDevolucion.RPT") Then Exit Sub
            ImprimoNotasYContado paDNDevolucion, TipoDocumento.NotaDevolucion, ""
            
        
        Case 3      'Nota de Crédito
            'Primero abrimos el reporte y seteamos la impresora si da error devuelve true
            If InicializoReporteEImpresora(paIContadoN, paIContadoB, "NotaDevolucion.RPT") Then Exit Sub
            
            ImprimoNotasYContado paDNCredito, TipoDocumento.NotaCredito, lNCDevuelve.Caption
                
        Case 4      'Nota Especial
            'Primero abrimos el reporte y seteamos la impresora si da error devuelve true
            If InicializoReporteEImpresora(paIContadoN, paIContadoB, "NotaDevolucion.RPT") Then Exit Sub
            ImprimoNotasYContado paDNEspecial, TipoDocumento.NotaEspecial, ""
        
        Case 5      'Recibo
            'If InicializoReporteEImpresora(paIReciboN, paIReciboB, "Recibo.RPT", Orientacion:=2, mPaperSize:=11) Then Exit Sub
            If InicializoReporteEImpresora(paIReciboN, paIReciboB, "Recibo.RPT", Orientacion:=2, mPaperSize:=13) Then Exit Sub
            ImprimoReciboDePago CLng(Trim(tSerie.Tag))
            
        Case 6      'Remito
            If fnc_3Vias Then
                fnc_PrintDocumento CLng(tSerie.Tag)
            Else
                ImprimoRemitoNew CLng(tSerie.Tag)
            End If
        
        Case 7      'Conforme
            If oCnfgPrint.Opcion = 0 Then
                If InicializoReporteEImpresora(paIConformeN, paIConformeB, "Conforme.RPT", 1, paIConformeP) Then Exit Sub
                ImprimoConforme
            Else
                ImprimoConformeTickets
            End If
            
            
        Case 8      'Nota de Débito
            If InicializoReporteEImpresora(paIReciboN, paIReciboB, "Aporte.RPT", Orientacion:=2, mPaperSize:=13) Then Exit Sub
            'If InicializoReporteEImpresora(paIReciboN, paIReciboB, "Aporte.RPT", Orientacion:=2, mPaperSize:=11) Then Exit Sub
            ImprimoNotaDebito Val(tSerie.Tag)
    
    End Select
    
    If Seleccionado <> TipoDocumento.Remito And Seleccionado <> TipoDocumento.Credito Then
        If Not crCierroTrabajo(jobnum) Then MsgBox crMsgErr
    End If
    
    DeshabilitoImpresion
    Screen.MousePointer = 0
    
    If prmIDDocumento <> 0 Then Unload Me
    Exit Sub
    
errImprimir:
    clsGeneral.OcurrioError "Error al realizar la reimpresión.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bOperacion_Click()
    EjecutarApp prmPathApp & "\Detalle de operaciones", CLng(tSerie.Tag)
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    Me.Refresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = 4620

    FechaDelServidor
    DeshabilitoImpresion
    CargoTiposDocumento
    
    Seleccionado = 7
    zfn_LoadMenuOpcionPrint
    
    oCnfgPrint.CargarConfiguracion cnfgAppNombreConformes, cnfgKeyTicketConformes
    crAbroEngine
    
    'prmIDDocumento = 361205
    If prmIDDocumento <> 0 Then ProcesoActivacion
    
    With vspPrinter
        .MarginBottom = 550
        .MarginLeft = 550
        .MarginRight = 550
        .MarginTop = 550
        .PageBorder = pbNone
    End With
    
    InicializoReporteEImpresoraCredito paICreditoN, paICreditoB, "Credito.RPT"
    
End Sub

Private Sub ProcesoActivacion()
On Error GoTo errAA
    
    Dim bHay As Boolean: bHay = False
    Dim aTipo As Integer
    
    Cons = "Select * from Documento Where DocCodigo = " & prmIDDocumento
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    If Not RsAux.EOF Then
        bHay = True
        aTipo = RsAux!DocTipo
        
        For I = opDocumento.LBound To opDocumento.UBound
            If Val(opDocumento(I).Tag) = aTipo Then
                opDocumento(I).Value = True
                Exit For
            End If
        Next
        
        tSerie.Text = Trim(RsAux!DocSerie)
        tNumero.Text = Trim(RsAux!DocNumero)
    End If
    RsAux.Close
    
    If bHay Then
        Dim idDoc As Long
        idDoc = BuscoDocumento(opDocumento(Seleccionado).Tag)
        If idDoc > 0 Then CargoDatosDocumento idDoc
    End If
    Exit Sub
    
errAA:
    clsGeneral.OcurrioError "Error al cargar los datos del documento.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    Screen.MousePointer = 11
    crCierroEngine
    cBase.Close
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    
    Screen.MousePointer = 0
    End
    
End Sub

Private Sub Label1_Click()
    Foco tSerie
End Sub

Private Sub Label4_Click()
    Foco tRuc
End Sub

Private Sub MnuPrintConformes_Click()
    
    frmDondeImprimo.Show vbModal
    oCnfgPrint.CargarConfiguracion cnfgAppNombreConformes, cnfgKeyTicketConformes

End Sub

Private Sub MnuPrintSalidasCaja_Click()
On Error Resume Next
    Dim frmSC As New frmDondeImprimoSC
    frmSC.ImpresorapapelA5 = paIRemitoN
    frmSC.Show vbModal
    Unload frmSC
End Sub

Private Sub opDocumento_Click(Index As Integer)
    
    DeshabilitoImpresion
    Seleccionado = Index
    'If Seleccionado = 6 Then
    '    tSerie.Enabled = False: tSerie.BackColor = Inactivo
    'Else
    '    tSerie.Enabled = True: tSerie.BackColor = Blanco
    'End If

End Sub

Private Sub DeshabilitoImpresion()
    
    tSerie.Text = ""
    tSerie.Tag = ""
    tNumero.Text = ""
    
    LimpioDocumento
    
End Sub

Private Sub opDocumento_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If tSerie.Enabled Then tSerie.SetFocus Else tNumero.SetFocus
    End If

End Sub

Private Sub tNumero_GotFocus()
    tNumero.SelStart = 0: tNumero.SelLength = Len(tNumero.Text)
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tNumero.Text) = "" Then Exit Sub
        If Trim(tSerie.Text) = "" And Seleccionado <> 6 Then Exit Sub
        If IsNumeric(tSerie.Text) Or Not IsNumeric(tNumero.Text) Then
            MsgBox "Los datos ingresados para la búsqueda no son correctos.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        Dim idDoc As Long
        idDoc = BuscoDocumento(opDocumento(Seleccionado).Tag)
        If idDoc > 0 Then CargoDatosDocumento idDoc
    End If
    
End Sub

Private Sub tSerie_GotFocus()
    tSerie.SelStart = 0
    tSerie.SelLength = Len(tSerie.Text)
End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then Foco tNumero
    
End Sub

Private Function BuscoDocumento(ByVal TipoDocumento As Integer) As Long
On Error GoTo errBD
Dim sQy As String

    sQy = "SELECT DocCodigo, DocFecha as Fecha" & _
        ", rtrim(TDoNombre) Documento " & _
        ", rTrim(DocSerie) + '-' + rtrim(Convert(Varchar(6), DocNumero)) as Número" & _
        " FROM Documento INNER JOIN TipoDocumento ON DocTipo = TDoId " _
        & " Where DocTipo = " & TipoDocumento _
        & " And DocSucursal = " & paCodigoDeSucursal _
        & " And DocSerie = '" & Trim(tSerie.Text) & "'" _
        & " And DocNumero = " & Trim(tNumero.Text) _
        & " Order by DocCodigo DESC"
        
'    sQy = "SELECT DocCodigo, DocFecha as Fecha" & _
        ", 'Crédito' Documento " & _
        ", rTrim(DocSerie) + '-' + rtrim(Convert(Varchar(6), DocNumero)) as Número" & _
        " FROM Documento " _
        & " Where DocTipo = 2 " _
        & " And DocSucursal = " & paCodigoDeSucursal _
        & " And DocSerie = '" & Trim(tSerie.Text) & "'" _
        & " And DocNumero = " & Trim(tNumero.Text) _
        & " "
        
    Dim iCodigo As Long, iQ As Integer
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
    Exit Function
errBD:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description, "Busco documento"
End Function

Private Sub CargoDatosDocumento(ByVal idDoc As Long)
Dim RsDoc As rdoResultset
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    LimpioDocumento
    Dim iTipoDoc As Integer
    
    'Saco los datos del Documento, Cliente
    Cons = "Select Documento.*, CliDireccion, CliCIRUC, CPeApellido1, CPeApellido2, CPeNombre1, CPeNombre2, CPeRUC, CEmNombre, CEmFantasia " _
        & " From Documento, Cliente " _
                & "Left Outer Join CPersona On CPeCliente = CliCodigo " _
                & "Left Outer Join CEmpresa On CEmCliente = CliCodigo " _
        & " WHERE DocCodigo = " & idDoc _
        & " And DocCliente = CliCodigo"
    If ObtenerResultSet(cBase, RsDoc, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    If RsDoc.EOF Then
        Screen.MousePointer = 0
        MsgBox "No existe un documento para la numeración ingresada o bien no pertenece a esta sucursal.", vbExclamation, "ATENCIÓN"
    Else
        'Verifico si el Documento no fue anulado (Papel)--------------------------------------
        If RsDoc!DocAnulado Then
            Screen.MousePointer = 0
            MsgBox "El documento ingresado figura como papel anulado.", vbExclamation, "ATENCIÓN"
        Else
            tSerie.Tag = RsDoc!DocCodigo
            labEmision.Caption = " " & Format(RsDoc!DocFecha, "dd/mm/yy hh:mm")
            tNumero.Tag = RsDoc!DocMoneda
            Label1.Tag = RsDoc!DocTotal
            iTipoDoc = RsDoc("DocTipo")
            If Not IsNull(RsDoc!CliDireccion) Then labDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, RsDoc!CliDireccion)
            
            lUsuario.Caption = BuscoDigitoUsuario(RsDoc!DocUsuario)
            If Not IsNull((RsDoc!DocVendedor)) Then lVendedor.Caption = BuscoDigitoUsuario(RsDoc!DocVendedor)
            
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
                If Not IsNull(RsDoc!CliCIRuc) Then tRuc.Text = RsDoc!CliCIRuc 'RetornoFormatoRuc(RsDoc!CliCiRuc)
            End If
            If tRuc.Text = "" Then tRuc.Enabled = True
            
            bImprimir.Enabled = True
            If Seleccionado < 4 Or Seleccionado = 7 Then bOperacion.Enabled = True
            
        End If
    End If
    RsDoc.Close
    
    If iTipoDoc = TipoDocumento.NotaCredito And bImprimir.Enabled Then
        lNCDevuelve.Caption = "SE DEVUELVE: "
        
        Cons = "Select isNull(NotSalidaCaja, 0) as NotSalidaCaja from Nota Where NotNota = " & Val(tSerie.Tag)
        'Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If ObtenerResultSet(cBase, RsDoc, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
        If Not RsDoc.EOF Then
            If Not IsNull(RsDoc!NotSalidaCaja) Then lNCDevuelve.Caption = lNCDevuelve.Caption & Format(RsDoc!NotSalidaCaja, FormatoMonedaP)
        End If
        RsDoc.Close
        
    End If
    
    If Val(tSerie.Tag) > 0 Then
        Cons = "SELECT TAIDocumento FROM TicketsAImprimir " _
            & "WHERE TAIDocumento = " & Val(tSerie.Tag)
        'Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If ObtenerResultSet(cBase, RsDoc, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
        If Not RsDoc.EOF Then
            MsgBox "ATENCIÓN!!!" & vbCrLf & vbCrLf & "El documento que desea reimprimir fue impreso en un TICKET, ud debería reimprimir en el servidor de tickets y no aquí.", vbCritical, "POSIBLE ERROR"
        End If
        RsDoc.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub LimpioDocumento()
    
    tSerie.Tag = ""
    tRuc.Text = "": tRuc.Enabled = False
    lCI.Caption = ""
    Frame1.Tag = ""
    Label6.Tag = ""
    labCliente.Caption = ""
    
    labDireccion.Caption = ""
    labEmision.Caption = ""
    
    lUsuario.Caption = ""
    lVendedor.Caption = ""
    lNCDevuelve.Caption = ""
    
    bImprimir.Enabled = False
    bOperacion.Enabled = False

End Sub

Private Sub CargoTiposDocumento()
    opDocumento(0).Tag = TipoDocumento.Contado
    opDocumento(1).Tag = TipoDocumento.Credito
    opDocumento(2).Tag = TipoDocumento.NotaDevolucion
    opDocumento(3).Tag = TipoDocumento.NotaCredito
    opDocumento(4).Tag = TipoDocumento.NotaEspecial
    opDocumento(5).Tag = TipoDocumento.ReciboDePago
    opDocumento(6).Tag = TipoDocumento.Remito
    opDocumento(7).Tag = TipoDocumento.Credito
    opDocumento(8).Tag = TipoDocumento.NotaDebito
End Sub

Private Sub ImprimoNotasYContado(NombreDoc As String, TipoDoc As Integer, TRetira As String)
On Error GoTo ErrCrystal

    Screen.MousePointer = 11
    
    BuscoDatosMoneda Val(tNumero.Tag), aMSigno, aMNombre
    
    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte Contado --------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento": result = crSeteoFormula(jobnum%, NombreFormula, "'" & NombreDoc & "'")
            Case "cliente"
                    aTexto = Trim(labCliente.Caption)
                    If lCI.Caption <> "" Then aTexto = aTexto & " (" & Trim(lCI.Caption) & ")"
                    result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
                    
            Case "direccion": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labDireccion.Caption) & "'")
            Case "ruc":
                If Trim(tRuc.Text) <> "" Then aTexto = tRuc.FormattedText Else aTexto = ""
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
            
            Case "codigobarras":
                    If TipoDoc = TipoDocumento.Contado Then aTexto = CodigoDeBarras(TipoDoc, CLng(tSerie.Tag)) Else aTexto = ""
                    result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
            
            Case "signomoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'" & aMSigno & "'")
            Case "nombremoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'(" & aMNombre & ")'")
            
            Case "textoretira": If Trim(TRetira) <> "" Then result = crSeteoFormula(jobnum%, NombreFormula, "'" & TRetira & "'")
                
            Case "usuario": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lUsuario.Caption) & "'")
            Case "vendedor": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lVendedor.Caption) & "'")
            
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor, Documento.DocCofis" _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where DocCodigo = " & CLng(tSerie.Tag)
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
        
    'Subreporte srContado.rpt  y srContado.rpt - 01-----------------------------------------------------------------------------
    JobSRep1 = crAbroSubreporte(jobnum, "srContado.rpt")
    If JobSRep1 = 0 Then GoTo ErrCrystal
    
     Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion, ArticuloEspecifico.AEsNombre " _
            & " From ({ oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId} Left Outer Join " _
                           & paBD & ".dbo.ArticuloEspecifico On AEsTipoDocumento = 1 And AEsDocumento = RenDocumento And AEsArticulo = RenArticulo)"
                           
    
    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
    
    JobSRep2 = crAbroSubreporte(jobnum, "srContado.rpt - 01")
    If JobSRep2 = 0 Then GoTo ErrCrystal
    If crSeteoSqlQuery(JobSRep2, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------

    'If crMandoAPantalla(jobnum, "Reimpresion Contado") = 0 Then GoTo ErrCrystal
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

Private Sub ImprimoReciboDePago(xRecibo As Long)
On Error GoTo ErrCrystal

    Screen.MousePointer = 11
    
    BuscoDatosMoneda Val(tNumero.Tag), aMSigno, aMNombre
    
    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte Contado --------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento": result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDRecibo & "'")
            Case "cliente": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labCliente.Caption) & "'")
            Case "cedula": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lCI.Caption) & "'")
            Case "ruc":
                If Trim(tRuc.Text) <> "" Then aTexto = Trim(tRuc.FormattedText) Else aTexto = ""
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
                
            Case "signomoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'" & aMSigno & "'")
            Case "nombremoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'(" & aMNombre & ")'")
            
            Case "usuario": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lUsuario.Caption) & "'")
            
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where DocCodigo = " & xRecibo
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
        
    'Subreporte srContado.rpt  y srContado.rpt - 01-----------------------------------------------------------------------------
    JobSRep1 = crAbroSubreporte(jobnum, "srRecibo.rpt")
    If JobSRep1 = 0 Then GoTo ErrCrystal
    
    'cons = "SELECT  DocumentoPago.DPaDocASaldar, DocumentoPago.DPaDocQSalda, DocumentoPago.DPaCuota, DocumentoPago.DPaDe, " _
                       & " DocumentoPago.DPaAmortizacion, DocumentoPago.DPaMora, Documento.DocSerie, Documento.DocNumero, " _
                       & " CreditoCuota.CCuValor, CreditoCuota.CCuVencimiento, Credito.CreProximoVto " _
            & " From { oj ((" & paBD & ".dbo.DocumentoPago DocumentoPago " _
                                & " INNER JOIN " & paBD & ".dbo.Documento Documento ON DocumentoPago.DPaDocASaldar = Documento.DocCodigo)" _
                                & " INNER JOIN " & paBD & ".dbo.Credito Credito ON Documento.DocCodigo = Credito.CreFactura)" _
                                & " INNER JOIN " & paBD & ".dbo.CreditoCuota CreditoCuota ON Credito.CreCodigo = CreditoCuota.CCuCredito}" _
            & " Where DocumentoPago.DPaDocQSalda = " & xRecibo _
            & " And DocumentoPago.DPaCuota = CreditoCuota.CCuNumero"
            
    Cons = "SELECT  DocumentoPago.DPaDocASaldar, DocumentoPago.DPaDocQSalda, DocumentoPago.DPaCuota, DocumentoPago.DPaDe, " _
                       & " DocumentoPago.DPaAmortizacion, DocumentoPago.DPaMora, Documento.DocSerie, Documento.DocNumero, " _
                       & " CreditoCuota.CCuValor, CreditoCuota.CCuVencimiento, Credito.CreProximoVto " _
            & " From { oj ((" & paBD & ".dbo.DocumentoPago DocumentoPago " _
                                & " INNER JOIN " & paBD & ".dbo.Documento Documento ON DocumentoPago.DPaDocASaldar = Documento.DocCodigo)" _
                                & " LEFT OUTER JOIN " & paBD & ".dbo.Credito Credito ON Documento.DocCodigo = Credito.CreFactura)" _
                                & " LEFT OUTER JOIN " & paBD & ".dbo.CreditoCuota CreditoCuota ON Credito.CreCodigo = CreditoCuota.CCuCredito And DocumentoPago.DPaCuota = CreditoCuota.CCuNumero}" _
            & " Where DocumentoPago.DPaDocQSalda = " & xRecibo

    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
    
    JobSRep2 = crAbroSubreporte(jobnum, "srRecibo.rpt - 01")
    If JobSRep2 = 0 Then GoTo ErrCrystal
    If crSeteoSqlQuery(JobSRep2, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------

    'If crMandoAPantalla(jobnum, "Recibo de Pago") = 0 Then GoTo ErrCrystal
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

Private Function ImprimoNotaDebito(xIDNotaDebito As Long)
On Error GoTo ErrCrystal

    Screen.MousePointer = 11
    
    BuscoDatosMoneda Val(tNumero.Tag), aMSigno, aMNombre
    
    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte Contado --------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento": result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDNDebito & "'")
            Case "cliente": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(labCliente.Caption) & "'")
            Case "cedula": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lCI.Caption) & "'")
            Case "ruc":
                If Trim(tRuc.Text) <> "" Then aTexto = Trim(tRuc.FormattedText) Else aTexto = ""
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
                
            Case "signomoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'" & aMSigno & "'")
            Case "nombremoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'(" & aMNombre & ")'")
            
            Case "usuario": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lUsuario.Caption) & "'")
           
            Case "cuenta":
                aTexto = UCase("Concepto: Intereses por Mora")
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
            
            Case "articulo":
                aTexto = ""
                'If Trim(tArticulo.Text) <> "" Then aTexto = "Destinado a la compra de: " & Trim(tArticulo.Text)
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where DocCodigo = " & xIDNotaDebito
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal

    'If crMandoAPantalla(jobnum, "Nota de Debito") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
    'crEsperoCierreReportePantalla
    
    Screen.MousePointer = 0
    Exit Function
ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    Exit Function
End Function

Private Function InicializoReporteEImpresora(paNImpresora As String, paBImpresora As Integer, Reporte As String, Optional Orientacion As Integer = 1, Optional mPaperSize As Integer = -1) As Boolean
On Error GoTo ErrCrystal
    
    jobnum = crAbroReporte(prmPathListados & Reporte)
    If jobnum = 0 Then GoTo ErrCrystal
    
    If ChangeCnfgPrint Then prj_LoadConfigPrint bShowFrm:=False
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paNImpresora) Then SeteoImpresoraPorDefecto paNImpresora
    If Not crSeteoImpresora(jobnum, Printer, paBImpresora, paperSize:=mPaperSize, mOrientation:=Orientacion) Then GoTo ErrCrystal
    
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
    Exit Function
End Function

Private Function InicializoReporteEImpresoraCredito(paNImpresora As String, paBImpresora As Integer, Reporte As String, Optional Orientacion As Integer = 1, Optional mPaperSize As Integer = -1) As Boolean
On Error GoTo ErrCrystal
    
    jobCredito = crAbroReporte(prmPathListados & "Credito.RPT")

    If ChangeCnfgPrint Then prj_LoadConfigPrint bShowFrm:=False
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paNImpresora) Then SeteoImpresoraPorDefecto paNImpresora
    ', paperSize:=mPaperSize, mOrientation:=Orientacion
    If Not crSeteoImpresora(jobCredito, Printer, paBImpresora) Then GoTo ErrCrystal
    
    InicializoReporteEImpresoraCredito = False
    Exit Function

ErrCrystal:
    InicializoReporteEImpresoraCredito = True
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroTrabajo jobCredito
    Screen.MousePointer = 0
    Exit Function
End Function



Private Function ImprimoRemitoNew(IDDocumento As Long)
On Error GoTo errImprimir

    Dim mTT As New Comercio.clsFunciones
    Set mTT.Connect = cBase
    mTT.ImprimirDocumento paDRemito, IDDocumento, paIRemitoN, paIRemitoB, labCliente.Caption, labDireccion.Caption, tRuc.Text, ""
    Set mTT = Nothing
    Exit Function
errImprimir:
    clsGeneral.OcurrioError "Error al imprimr el documento seleccionado.", Err.Description
End Function

Private Sub ImprimoConformeTickets()
Dim Cont As Integer
Dim sAux As String, Documento As String
Dim RsAuxC As rdoResultset
Dim MEntrega As Currency
Dim aFechaDoc As Date


    'Saco la fecha del documento--------------------------------------------------------------------
'    Cons = "Select DocFecha, DocSerie, DocNumero FROM Documento Where DocCodigo = " & CLng(tSerie.Tag)
'    Set RsAuxC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'    aFechaDoc = RsAuxC!DocFecha
'    Documento = UCase(Trim(RsAuxC("DocSerie"))) & " " & Trim(RsAuxC("DocNumero"))
'    RsAuxC.Close
'    '----------------------------------------------------------------------------------------------------------
'
'    'Consulta para sacar los datos del credito------------------------------------------------------------
'     Cons = "Select * from Credito, TipoCuota " _
'             & " Where CreFactura = " & CLng(tSerie.Tag) _
'             & " And CreTipoCuota = TCuCodigo"
'     Set RsAuxC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'     'VaCuota "" = Diferido, pago envio ...si hay, "E" = Entrega, "1...N" Cuota.
'
'    MEntrega = 0
'    If Not IsNull(RsAuxC!TCuVencimientoE) Then
'        Cons = " Select * from CreditoCuota " _
'                & " Where  CCuCredito = " & RsAuxC!CreCodigo _
'                & " And CCuNumero = 0"
'        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'        If Not RsAux.EOF Then MEntrega = RsAux!CCuValor
'        RsAux.Close
'    End If
'    '---------------------------------------------------------------------------------------------------------------------------
'
'    BuscoDatosMoneda Val(tNumero.Tag), aMSigno, aMNombre
'    aMNombre = LCase(aMNombre)
'
'    Dim sTitular As String
'    Dim sAbal As String, sDomAbal As String, sFirmaAbal As String, sCITit As String, sCIAbal As String
'
'    If Trim(lCI.Caption) <> "" Then
'        sCITit = "Titular  C.I.: " & Trim(lCI.Caption)
'    ElseIf Trim(tRuc.Text) <> "" Then
'        sCITit = "Titular  R.U.T.: " & Trim(tRuc.FormattedText)
'    End If
'    sTitular = Trim(labCliente.Caption)
'
'    sAbal = ""
'    Dim sPieConAbal As String, sPieSinAbal As String
'
'    sPieSinAbal = "Vale Nº : " + Documento
'
'
'    Dim sImpEntrega As String
'
'    If Not IsNull(RsAuxC!CreGarantia) Then
'        Cons = "SELECT CliTipo, CliCIRuc, CPeApellido1, CPeApellido2, CPeNombre1, CPeNombre2, CEmFantasia, CEmNombre from Cliente " & _
'                    " Left Outer Join CPersona ON CliCodigo = CPeCliente " & _
'                    " Left Outer Join CEmpresa ON CliCodigo = CEmCliente " & _
'               " WHERE CliCodigo = " & RsAuxC!CreGarantia
'        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'        If Not RsAux.EOF Then
'            sPieSinAbal = ""
'            sPieConAbal = "Vale Nº : " + Documento
'            If RsAux!CliTipo = 1 Then
'                If Not IsNull(RsAux!CliCIRuc) Then sCIAbal = "Garantía  C.I.: " & clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc)
'                sAbal = ArmoNombre(CStr(RsAux!CPeApellido1), Format(RsAux!CPeApellido2, "#"), CStr(RsAux!CPeNombre1), Format(RsAux!CPeNombre2, "#"))
'            Else
'                If Not IsNull(RsAux!CliCIRuc) Then sCIAbal = "Garantía  R.U.T. " & clsGeneral.RetornoFormatoRuc(RsAux!CliCIRuc) Else sCIAbal = "Garantía"
'                sAbal = Trim(RsAux!CEmFantasia)
'                If Not IsNull(RsAux!CEmNombre) Then sAbal = sAbal & " (" & Trim(RsAux!CEmNombre) & ")"
'            End If
'        End If
'        RsAux.Close
'    End If
'
'    Dim sPtsAbal As String
'    If Trim(sAbal) <> "" Then
'        sDomAbal = "Domicilio"
'        sPtsAbal = String(40, ".")
'        sFirmaAbal = "Firma"
'    Else
'        sDomAbal = Chr(2)
'        sPtsAbal = Chr(2) & vbCrLf
'        sFirmaAbal = Chr(2)
'        sCIAbal = Chr(2)
'        sAbal = Chr(2)
'    End If
'
'    'Armo detalle.
'    Dim sDetalle As String
'    Dim sEntCuota As String
''    sDetalle = "VALE [nroConforme], por [nombremoneda] [total], que pagaremos en forma indivisible y solidariamente a CARLOS GUTIERREZ SA o a su orden [Entrega][QCuotas] cuotas consecutivas de [nombremoneda] [valorCuota] exigibles cada [distanciacuota] días venciendo la 1ª al [fecha1Cuota]. " & _
''            "Los pagos indicados incluyen un interés compensatorio, de 35.00% efectivo anual. " & vbCrLf & _
''            "Se aplicará redondeo al valor de las cuotas a fin de obtener múltiplos de 5. " & vbCrLf & vbCrLf & _
''            "La mora se producirá de pleno derecho, sin necesidad de interpelación judicial o extrajudicial alguna, por el no pago de una cuota a su vto..  " & vbCrLf & _
''            "A partir de ese momento devengará un interés moratorio de X,XX% efectivo anual."
'
'
'    sDetalle = "Los pagos indicados incluyen un interés compensatorio, de 35.00% efectivo anual. " & _
'        "El resultado obtenido al aplicar la tasa de interés se aproximará (redondeo) a la media decena superior o inferior, según cuál sea más cercana a aquel resultado. " & _
'        "La mora se producirá de pleno derecho, sin necesidad de interpelación judicial o extrajudicial alguna, por el no pago de una cuota a su vencimiento. " & _
'        "A partir de ese momento devengará un interés moratorio de 5.87% efectivo anual."
'
'    Dim cImpTotal As Currency
'    cImpTotal = (RsAuxC!CreValorCuota * RsAuxC!TCuCantidad) + MEntrega
'
'    sDetalle = Replace(sDetalle, "[nroConforme]", Documento, , , vbTextCompare)
'    'sDetalle = Replace(sDetalle, "[total]", ImporteATexto((RsAuxC!CreValorCuota * RsAuxC!TCuCantidad) + MEntrega), , , vbTextCompare)
'
''Entrega
'    sEntCuota = ""
'    If Not IsNull(RsAuxC!TCuVencimientoE) Then
'
'        If RsAuxC("TCUVencimientoE") > 0 Then
'            'sEntCuota = "con una entrega de " & aMNombre & " " & UCase(ImporteATexto(MEntrega)) & vbCrLf & " con vencimiento el "
'            sEntCuota = "con una entrega de " & aMSigno & " " & Format(MEntrega, "#,##0.00") & vbCrLf & " con vencimiento el " & Format(CDate(aFechaDoc) + RsAuxC!TCuVencimientoE, "dd/mm/yyyy")
'        Else
'            sEntCuota = "entregando hoy " & aMSigno & " " & Format(MEntrega, "#,##0.00")
'        End If
'    End If
''    sDetalle Microsoft Sans Serif= Replace(sDetalle, "[Entrega]", sEntCuota, , , vbTextCompare)
'
''Cuotas
'    Dim sImpCuotas As String
'    sImpCuotas = IIf(sEntCuota = "", "", "y ") & "en [QCuotas] cuotas consecutivas de [valorCuota]" & vbCrLf & "exigibles cada [distanciacuota] días" & vbCrLf & "venciendo la 1ª el [fecha1Cuota]."
'
'    sImpCuotas = Replace(sImpCuotas, "[QCuotas]", RsAuxC!TCuCantidad, , , vbTextCompare)
'    sImpCuotas = Replace(sImpCuotas, "[valorCuota]", aMSigno & " " & Format(RsAuxC!CreValorCuota, "#,##0.00"), , , vbTextCompare)
'    sImpCuotas = Replace(sImpCuotas, "[distanciacuota]", RsAuxC!TCuDistancia, , , vbTextCompare)
'    sImpCuotas = Replace(sImpCuotas, "[fecha1Cuota]", Format(CDate(aFechaDoc) + RsAuxC!TCuVencimientoC, "dd/mm/yyyy"), , , vbTextCompare)
    

    With vsrReport
        .Clear                  ' clear any existing fields
        .FontName = "Tahoma"    ' set default font for all controls
        .FontSize = 8
        
        .Load prmPathListados & "Conforme.xml", "Conforme"
    
        .DataSource.ConnectionString = cBase.Connect
        .DataSource.RecordSource = "prg_Conformes_Impresion (" & CLng(tSerie.Tag) & ")"
        '.DataSource.RecordSource = "SELECT 'VALE Nº: ' + RTrim(DocSerie + CONVERT(varchar(6), DocNumero)) NroConforme " & _
                ", 'SUCURSAL " & UCase(prmNombreSucursal) & "' Sucursal " & _
                ", '*' + Convert(char(1), DocTipo) + 'C' + CONVERT(varchar(6), DocNumero) + '*' conformeBarCode " & _
                ", 'Montevideo, ' + RTrim(Day(DocFecha)) + ' de ' + dbo.Mes(DocFecha) + ' de ' + RTrim(Convert(char(4), Year(DocFecha))) Fecha " & _
                ", '" & sTitular & "' TitularNombre " & _
                ", '" & sCITit & "' TitularCedula " & _
                ", 'Conforme: ' + RTrim(DocSerie + ' ' + CONVERT(varchar(6), DocNumero)) " & _
                ", '" & sCIAbal & "' AbalCedula " & _
                ", '" & sAbal & "' AbalNombre " & _
                ", '" & sPtsAbal & "' AbalPuntos " & _
                ", '" & sDomAbal & "' AbalDomicilio, '" & sFirmaAbal & "' AbalFirma " & _
                ", '" & sDetalle & "' Detalle " & _
                ", '" & String(40, ".") & vbCrLf & vbCrLf & String(40, ".") & "' puntos " & _
                ", '" & IIf(sCIAbal <> "", String(40, ".") & vbCrLf & vbCrLf & String(40, ".") & vbCrLf & vbCrLf, "") & "' puntosabal " & _
                ", '" & sPieConAbal & "' PieConGarantia " & _
                ", '" & sPieSinAbal & "' PieSinGarantia " & _
                ", '" & sEntCuota & "' ImporteEntrega " & _
                ", '" & sImpCuotas & "' ImporteCuotas " & _
                ", ' Por " & aMNombre & " " & aMSigno & " " & Format(cImpTotal, "#,##0.00") & "' ImporteConforme " & _
                "FROM Documento WHERE DocCodigo = " & CLng(tSerie.Tag)
        
        vspPrinter.Device = oCnfgPrint.ImpresoraTickets
        
        .Render vspPrinter
        
    End With
    
    vspPrinter.PrintDoc False

End Sub

'------------------------------------------------------------------------------------------------------------------------------------
'   Caso de Entregas:   Va la entrega Con Envio (no importa).
'   Caso Sin Entrega:   Va el valor de la cuota (se descarta el envio).
'------------------------------------------------------------------------------------------------------------------------------------
Private Sub ImprimoConforme()

Dim Cont As Integer
Dim sAux As String, Fletes As String
Dim RsAuxC As rdoResultset
Dim MEntrega As Currency
Dim aFechaDoc As Date

    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Saco la fecha del documento--------------------------------------------------------------------
    Cons = "Select * from Documento Where DocCodigo = " & CLng(tSerie.Tag)
    Set RsAuxC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aFechaDoc = RsAuxC!DocFecha
    RsAuxC.Close
    '----------------------------------------------------------------------------------------------------------
    
    'Consulta para sacar los datos del credito------------------------------------------------------------
     Cons = "Select * from Credito, TipoCuota " _
             & " Where CreFactura = " & CLng(tSerie.Tag) _
             & " And CreTipoCuota = TCuCodigo"
     Set RsAuxC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
     'VaCuota "" = Diferido, pago envio ...si hay, "E" = Entrega, "1...N" Cuota.
    
    MEntrega = 0
    If Not IsNull(RsAuxC!TCuVencimientoE) Then
        Cons = " Select * from CreditoCuota " _
                & " Where  CCuCredito = " & RsAuxC!CreCodigo _
                & " And CCuNumero = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then MEntrega = RsAux!CCuValor
        RsAux.Close
    End If
    '---------------------------------------------------------------------------------------------------------------------------
    
    BuscoDatosMoneda Val(tNumero.Tag), aMSigno, aMNombre
    aMNombre = LCase(aMNombre)
    
    'Cargo Propiedades para el Conforme --- --------------------------------------------------------------------------------
    For Cont = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, Cont)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "sucursal": result = crSeteoFormula(jobnum%, NombreFormula, "'SUCURSAL " & UCase(prmNombreSucursal) & "'")
            Case "titular"
                aTexto = ""
                If Trim(lCI.Caption) <> "" Then aTexto = Trim(lCI.Caption) Else If Trim(tRuc.Text) <> "" Then aTexto = Trim(tRuc.FormattedText)
                aTexto = Trim(aTexto & " " & Trim(labCliente.Caption))
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
            
            Case "garantia"
                aTexto = ""
                If Not IsNull(RsAuxC!CreGarantia) Then
                    Cons = "Select * from Cliente " & _
                                " Left Outer Join CPersona ON CliCodigo = CPeCliente " & _
                                " Left Outer Join CEmpresa ON CliCodigo = CEmCliente " & _
                           " Where CliCodigo = " & RsAuxC!CreGarantia
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If Not RsAux.EOF Then
                        If RsAux!CliTipo = 1 Then
                            If Not IsNull(RsAux!CliCIRuc) Then aTexto = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc)
                            aTexto = aTexto & " " & ArmoNombre(CStr(RsAux!CPeApellido1), Format(RsAux!CPeApellido2, "#"), CStr(RsAux!CPeNombre1), Format(RsAux!CPeNombre2, "#"))
                        Else
                            If Not IsNull(RsAux!CliCIRuc) Then aTexto = clsGeneral.RetornoFormatoRuc(RsAux!CliCIRuc)
                            aTexto = aTexto & " " & Trim(RsAux!CEmFantasia)
                            If Not IsNull(RsAux!CEmNombre) Then aTexto = aTexto & " (" & Trim(RsAux!CEmNombre) & ")"
                        End If
                    End If
                    RsAux.Close
                    
                    result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
                    
                End If
            
            Case "texto1"       'Texto del Conforme     Importe General
                sAux = ImporteATexto((RsAuxC!CreValorCuota * RsAuxC!TCuCantidad) + MEntrega)
                aTexto = " por la suma de " & aMNombre & " " & UCase(sAux) & ", "
                aTexto = aTexto & "que pagaremos en forma indivisible y solidariamente a CARLOS GUTIERREZ S.A. o a su orden en "
                
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
                
            Case "texto2"       'Texto del Conforme     Importe de Entrega
                aTexto = ""
                If Not IsNull(RsAuxC!TCuVencimientoE) Then
                    aTexto = aTexto & "una entrega de " & aMNombre & " " & UCase(ImporteATexto(MEntrega)) & " con vencimiento el "
                    'Vencimiento de entrega
                    sAux = gFechaServidor + RsAuxC!TCuVencimientoE
                    sAux = " " & Format(sAux, "d") & " de " & Format(sAux, "Mmmm") & " de " & Format(sAux, "yyyy")
                    aTexto = aTexto & sAux & " y "
                End If
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
                             
            Case "texto3"   'Texto del Conforme     Importe de Cuotas
                sAux = ImporteATexto(RsAuxC!CreValorCuota)
                aTexto = RsAuxC!TCuCantidad & " cuotas consecutivas de " & aMNombre & " " & UCase(sAux) _
                            & " exigibles cada " & RsAuxC!TCuDistancia & " días, venciendo la primera el " '" días a partir del "
                
                'Primer Vencimiento de cuotas
                sAux = gFechaServidor + RsAuxC!TCuVencimientoC
                sAux = Format(sAux, "d") & " de " & Format(sAux, "Mmmm") & " de " & Format(sAux, "yyyy")
                
                aTexto = aTexto & " " & sAux
                
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
                                      
            Case "fecha"
                aTexto = "Montevideo, " & Format(aFechaDoc, "d") & " de " & Format(aFechaDoc, "Mmmm") & " de " & Format(aFechaDoc, "yyyy")
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
            
             
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    RsAuxC.Close
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocSerie, Documento.DocNumero " _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where Documento.DocCodigo = " & CLng(tSerie.Tag)

    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
   
    'If crMandoAPantalla(jobnum, "Conforme") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
        
    'crEsperoCierreReportePantalla
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    Screen.MousePointer = 11
End Sub

Private Sub ImprimoCredito()
Dim RsAuxC As rdoResultset
Dim sTexto As String, MEnvio As Currency, MEntrega As Currency
Dim sConCheques As Boolean

    'Consulta para sacar los datos del credito------------------------------------------------------------
     Cons = "Select * from Credito, TipoCuota " _
             & " Where CreFactura = " & CLng(tSerie.Tag) _
             & " And CreTipoCuota = TCuCodigo"
    Set RsAuxC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAuxC!CreFormaPago = TipoPagoSolicitud.ChequeDiferido Then sConCheques = True Else sConCheques = False
    
    'Saco el valor del flete    ---------------------------------------------------------------------
    Cons = "Select * From Renglon " _
            & " Where RenDocumento = " & CLng(tSerie.Tag) _
            & " And (RenArticulo IN (Select Distinct(TFlArticulo) From TipoFlete)" _
            & " Or RenArticulo = " & paArticuloPisoAgencia & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF
            MEnvio = RsAux!RenCantidad * RsAux!RenPrecio
            RsAux.MoveNext
        Loop
    Else
        MEnvio = 0
    End If
    RsAux.Close
    
    'Si es con Diferidos saco valor de las cuotas que no vencen el mismo dia que la factura ---------------------------
    Dim mTCheques As Currency
    mTCheques = 0
    If sConCheques Then
        Cons = "Select DocFecha, DocTotal, CreditoCuota.* " & _
                    " From CreditoCuota, Credito, Documento" & _
                    " Where CCuCredito = CreCodigo And CreFactura = DocCodigo " & _
                    " And CreCodigo = " & RsAuxC!CreCodigo
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            If Format(RsAux!DocFecha, "yyyy/mm/dd") <> Format(RsAux!CCuVencimiento, "yyyy/mm/dd") Then
                mTCheques = mTCheques + RsAux!CCuValor
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    '--------------------------------------------------------------------------------------------------------------------------------
    
    BuscoDatosMoneda Val(tNumero.Tag), aMSigno, aMNombre
    
    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobCredito)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte Credito --------------------------------------------------------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobCredito, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & paDCredito & "'")
            Case "cliente": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & Trim(labCliente.Caption) & "'")
            Case "cedula": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & Trim(lCI.Caption) & "'")
            Case "direccion": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & Trim(labDireccion.Caption) & "'")
            
            Case "ruc":
                If Trim(tRuc.Text) <> "" Then aTexto = tRuc.FormattedText Else aTexto = ""
                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & aTexto & "'")
            
            Case "codigobarras": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.Credito, CLng(tSerie.Tag)) & "'")
            Case "signomoneda": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & aMSigno & "'")
            Case "nombremoneda": result = crSeteoFormula(jobCredito%, NombreFormula, "'(" & aMNombre & ")'")
            
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
                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & sTexto & "'")
                
            Case "nombrerecibo": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & paDRecibo & "'")
            
            Case "usuario": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & Trim(lUsuario.Caption) & "'")
            Case "vendedor": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & Trim(lVendedor.Caption) & "'")
                                
            Case "financiacion"
                sTexto = Trim(RsAuxC!TCuAbreviacion) & " - "
                If Not IsNull(RsAuxC!TCuVencimientoE) Then
                    MEntrega = CCur(Label1.Tag) - (RsAuxC!TCuCantidad * RsAuxC!CreValorCuota)
                    If MEntrega > 0 Then sTexto = sTexto & "Ent.: " & Format(MEntrega, FormatoMonedaP) & " "
                End If
                sTexto = sTexto & RsAuxC!TCuCantidad & " x " & Format(RsAuxC!CreValorCuota, FormatoMonedaP)
                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & sTexto & "'")
                    
            Case "proximovto"
                If Not IsNull(RsAuxC!CreProximoVto) Then
                    sTexto = Format(RsAuxC!CreProximoVto, "d Mmm yyyy")
                    result = crSeteoFormula(jobCredito%, NombreFormula, "'" & sTexto & "'")
                End If
            
            '------------------------------------------------------------------------------------------------------------
            Case "recibotcuota"
                If Not sConCheques Then aTexto = "Cuotas:" Else aTexto = "Al DIA:"
                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & aTexto & "'")
                
            Case "recibotflete"
                If Not sConCheques Then aTexto = "Flete:" Else aTexto = "Ch. Diferidos Total:"
                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & aTexto & "'")
            
            Case "reciboflete":
                If Not sConCheques Then
                    sTexto = Format(MEnvio, FormatoMonedaP)
                Else
                    sTexto = Format(mTCheques, FormatoMonedaP)
                End If
                
                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & sTexto & "'")
            
            Case "recibocuota"
                If Trim(RsAuxC!CreVaCuota) <> "" Then
                    sTexto = Trim(RsAuxC!CreVaCuota) & " de " & Trim(RsAuxC!CreDeCuota)
                    result = crSeteoFormula(jobCredito%, NombreFormula, "'" & sTexto & "'")
                End If
            
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    RsAuxC.Close
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Top 1 Documento.DocCodigo, Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal," _
            & " Documento.DocIVA, Documento.DocVendedor, Recibo.DocSerie , Recibo.DocNumero, Recibo.DocTotal, Documento.DocCofis" _
            & " From " _
            & " { oj (" & paBD & ".dbo.Documento Documento " _
                        & " LEFT OUTER JOIN " & paBD & ".dbo.DocumentoPago DocumentoPago ON  Documento.DocCodigo = DocumentoPago.DPaDocASaldar)" _
                        & " LEFT OUTER JOIN " & paBD & ".dbo.Documento Recibo ON  DocumentoPago.DPaDocQSalda = Recibo.DocCodigo}" _
            & " Where Documento.DocCodigo = " & CLng(tSerie.Tag)
    
    If sConCheques Then
        Cons = Cons & " Group by Documento.DocCodigo, Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor, Recibo.DocSerie , Recibo.DocNumero, Recibo.DocTotal, Documento.DocComentario, Documento.DocCofis "
    End If
    
    If crSeteoSqlQuery(jobCredito%, Cons) = 0 Then GoTo ErrCrystal
    
    Dim iJobSRep1 As Integer, iJobSRep2 As Integer
    
    'Subreporte srCredito.rpt  y srCredito.rpt - 01-----------------------------------------------------------------------------
    iJobSRep1 = crAbroSubreporte(jobCredito, "srCredito.rpt")
    If iJobSRep1 = 0 Then GoTo ErrCrystal
    
    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"
        
    If crSeteoSqlQuery(iJobSRep1, Cons) = 0 Then GoTo ErrCrystal
    
    iJobSRep2 = crAbroSubreporte(jobCredito, "srCredito.rpt - 01")
    If iJobSRep2 = 0 Then GoTo ErrCrystal
    If crSeteoSqlQuery(iJobSRep2, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------
    
    'If crMandoAPantalla(jobCredito, "Factura Credito") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(jobCredito, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobCredito, True, False) Then GoTo ErrCrystal
    
    If Not crCierroSubReporte(iJobSRep1) Then GoTo ErrCrystal
    If Not crCierroSubReporte(iJobSRep2) Then GoTo ErrCrystal
        
    'crEsperoCierreReportePantalla
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroSubReporte JobSRep1
    crCierroSubReporte JobSRep2
    Screen.MousePointer = 0
End Sub

Private Sub ImprimoCredito_VSReport()
Dim RsAuxC As rdoResultset
Dim sTexto As String, MEnvio As Currency, MEntrega As Currency
Dim sConCheques As Boolean

On Error GoTo ErrCrystal

    'Consulta para sacar los datos del credito------------------------------------------------------------
     Cons = "Select * from Credito, TipoCuota " _
             & " Where CreFactura = " & CLng(tSerie.Tag) _
             & " And CreTipoCuota = TCuCodigo"
    Set RsAuxC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAuxC!CreFormaPago = TipoPagoSolicitud.ChequeDiferido Then sConCheques = True Else sConCheques = False
    
    'Saco el valor del flete    ---------------------------------------------------------------------
    Cons = "Select * From Renglon " _
            & " Where RenDocumento = " & CLng(tSerie.Tag) _
            & " And (RenArticulo IN (Select Distinct(TFlArticulo) From TipoFlete)" _
            & " Or RenArticulo = " & paArticuloPisoAgencia & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF
            MEnvio = RsAux!RenCantidad * RsAux!RenPrecio
            RsAux.MoveNext
        Loop
    Else
        MEnvio = 0
    End If
    RsAux.Close
    
    'Si es con Diferidos saco valor de las cuotas que no vencen el mismo dia que la factura ---------------------------
    Dim mTCheques As Currency
    mTCheques = 0
    If sConCheques Then
        Cons = "Select DocFecha, DocTotal, CreditoCuota.* " & _
                    " From CreditoCuota, Credito, Documento" & _
                    " Where CCuCredito = CreCodigo And CreFactura = DocCodigo " & _
                    " And CreCodigo = " & RsAuxC!CreCodigo
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            If Format(RsAux!DocFecha, "yyyy/mm/dd") <> Format(RsAux!CCuVencimiento, "yyyy/mm/dd") Then
                mTCheques = mTCheques + RsAux!CCuValor
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    '--------------------------------------------------------------------------------------------------------------------------------
    BuscoDatosMoneda Val(tNumero.Tag), aMSigno, aMNombre
   
    Screen.MousePointer = 11
    
    Dim oImprimo As clsImpresionCredito
    Set oImprimo = New clsImpresionCredito
    oImprimo.DondeImprimo.Bandeja = paICreditoB
    oImprimo.DondeImprimo.Impresora = paICreditoN
    oImprimo.DondeImprimo.Papel = 1
    oImprimo.PathReportes = prmPathListados
    oImprimo.StringConnect = miConexion.TextoConexion("Comercio")
    
    With oImprimo
        .field_ClienteCedula = Trim(lCI.Caption)
        .field_ClienteDireccion = Trim(labDireccion.Caption)
        If Not IsNull(RsAuxC!CreGarantia) Then
            'Cargo datos de la garantía.
            Cons = "Select * From Cliente,  CPersona " _
                & "Where CliCodigo  = " & RsAuxC!CreGarantia & " And CliCodigo = CPeCliente"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            sTexto = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc) & " " & ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
            RsAux.Close
            .field_ClienteGarantia = sTexto
        Else
            .field_ClienteGarantia = ""
        End If
        .field_ClienteNombre = Trim(labCliente.Caption)
        
        sTexto = Trim(RsAuxC!TCuAbreviacion) & " - "
        If Not IsNull(RsAuxC!TCuVencimientoE) Then
            MEntrega = CCur(Label1.Tag) - (RsAuxC!TCuCantidad * RsAuxC!CreValorCuota)
            If MEntrega > 0 Then sTexto = sTexto & "Ent.: " & Format(MEntrega, FormatoMonedaP) & " "
        End If
        sTexto = sTexto & RsAuxC!TCuCantidad & " x " & Format(RsAuxC!CreValorCuota, FormatoMonedaP)
        .field_CreditoFinanciacion = sTexto
        If Not IsNull(RsAuxC!CreProximoVto) Then
            .field_CreditoProxVto = Format(RsAuxC!CreProximoVto, "d Mmm yyyy")
        Else
            .field_CreditoProxVto = ""
        End If
        .field_MonedaNombre = aMNombre
        .field_MonedaSimbolo = aMSigno
        .field_NombreDocumento = paDCredito
        .field_NombreRecibo = paDRecibo
        
        If Trim(RsAuxC!CreVaCuota) <> "" Then
            .field_ReciboInfoCuota = "Cuotas: " & Trim(RsAuxC!CreVaCuota) & " de " & Trim(RsAuxC!CreDeCuota)
        Else
            .field_ReciboInfoCuota = ""
        End If
        
        If Not sConCheques Then
            .field_ReciboImporteFlete = Format(MEnvio, FormatoMonedaP)
        Else
            .field_ReciboImporteFlete = Format(mTCheques, FormatoMonedaP)
        End If
        
        If Not sConCheques Then
            .field_ReciboInfoTFlete = "Flete:"
        Else
            .field_ReciboInfoTFlete = "Ch. diferidos total:"
        End If
        
        If Not IsNull(RsAuxC!CreProximoVto) Then
            .field_ReciboInfoVto = "Próximo vencimiento: " & Format(RsAuxC!CreProximoVto, "d Mmm yyyy")
        Else
            .field_ReciboInfoVto = ""
        End If
        If Trim(tRuc.Text) <> "" Then .field_RUT = tRuc.FormattedText Else .field_RUT = ""
        
        .field_Vendedor = lVendedor.Caption
        .field_Digitador = lUsuario.Caption
        
        .ImprimoFacturaContado_VSReport CLng(tSerie.Tag)
    End With
    Screen.MousePointer = 0
    Exit Sub
    
    
'    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
'    CantForm = crObtengoCantidadFormulasEnReporte(jobCredito)
'    If CantForm = -1 Then GoTo ErrCrystal
'
'    'Cargo Propiedades para el reporte Credito --------------------------------------------------------------------------------
'    For I = 0 To CantForm - 1
'        NombreFormula = crObtengoNombreFormula(jobCredito, I)
'
'        Select Case LCase(NombreFormula)
'            Case "": GoTo ErrCrystal
'            Case "nombredocumento": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & paDCredito & "'")
'            Case "cliente": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & Trim(labCliente.Caption) & "'")
'            Case "cedula": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & Trim(lCI.Caption) & "'")
'            Case "direccion": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & Trim(labDireccion.Caption) & "'")
'
'            Case "ruc":
'                If Trim(tRuc.Text) <> "" Then aTexto = tRuc.FormattedText Else aTexto = ""
'                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & aTexto & "'")
'
'            Case "codigobarras": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.Credito, CLng(tSerie.Tag)) & "'")
'            Case "signomoneda": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & aMSigno & "'")
'            Case "nombremoneda": result = crSeteoFormula(jobCredito%, NombreFormula, "'(" & aMNombre & ")'")
'
'            Case "garantia"
'                sTexto = ""
'                If Not IsNull(RsAuxC!CreGarantia) Then
'                    'Cargo datos de la garantía.
'                    Cons = "Select * From Cliente,  CPersona " _
'                        & "Where CliCodigo  = " & RsAuxC!CreGarantia & " And CliCodigo = CPeCliente"
'                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
'                    sTexto = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc) & " " & ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
'                    RsAux.Close
'                End If
'                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & sTexto & "'")
'
'            Case "nombrerecibo": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & paDRecibo & "'")
'
'            Case "usuario": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & Trim(lUsuario.Caption) & "'")
'            Case "vendedor": result = crSeteoFormula(jobCredito%, NombreFormula, "'" & Trim(lVendedor.Caption) & "'")
'
'            Case "financiacion"
'                sTexto = Trim(RsAuxC!TCuAbreviacion) & " - "
'                If Not IsNull(RsAuxC!TCuVencimientoE) Then
'                    MEntrega = CCur(Label1.Tag) - (RsAuxC!TCuCantidad * RsAuxC!CreValorCuota)
'                    If MEntrega > 0 Then sTexto = sTexto & "Ent.: " & Format(MEntrega, FormatoMonedaP) & " "
'                End If
'                sTexto = sTexto & RsAuxC!TCuCantidad & " x " & Format(RsAuxC!CreValorCuota, FormatoMonedaP)
'                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & sTexto & "'")
'
'            Case "proximovto"
'                If Not IsNull(RsAuxC!CreProximoVto) Then
'                    sTexto = Format(RsAuxC!CreProximoVto, "d Mmm yyyy")
'                    result = crSeteoFormula(jobCredito%, NombreFormula, "'" & sTexto & "'")
'                End If
'
'            '------------------------------------------------------------------------------------------------------------
'            Case "recibotcuota"
'                If Not sConCheques Then aTexto = "Cuotas:" Else aTexto = "Al DIA:"
'                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & aTexto & "'")
'
'            Case "recibotflete"
'                If Not sConCheques Then aTexto = "Flete:" Else aTexto = "Ch. Diferidos Total:"
'                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & aTexto & "'")
'
'            Case "reciboflete":
'                If Not sConCheques Then
'                    sTexto = Format(MEnvio, FormatoMonedaP)
'                Else
'                    sTexto = Format(mTCheques, FormatoMonedaP)
'                End If
'                result = crSeteoFormula(jobCredito%, NombreFormula, "'" & sTexto & "'")
'
'            Case "recibocuota"
'                If Trim(RsAuxC!CreVaCuota) <> "" Then
'                    sTexto = Trim(RsAuxC!CreVaCuota) & " de " & Trim(RsAuxC!CreDeCuota)
'                    result = crSeteoFormula(jobCredito%, NombreFormula, "'" & sTexto & "'")
'                End If
'
'            Case Else: result = 1
'        End Select
'        If result = 0 Then GoTo ErrCrystal
'    Next
'    '--------------------------------------------------------------------------------------------------------------------------------------------
'    RsAuxC.Close
'
'    'Seteo la Query del reporte-----------------------------------------------------------------
'    Cons = "SELECT Top 1 Documento.DocCodigo, Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal," _
'            & " Documento.DocIVA, Documento.DocVendedor, Recibo.DocSerie , Recibo.DocNumero, Recibo.DocTotal, Documento.DocCofis" _
'            & " From " _
'            & " { oj (" & paBD & ".dbo.Documento Documento " _
'                        & " LEFT OUTER JOIN " & paBD & ".dbo.DocumentoPago DocumentoPago ON  Documento.DocCodigo = DocumentoPago.DPaDocASaldar)" _
'                        & " LEFT OUTER JOIN " & paBD & ".dbo.Documento Recibo ON  DocumentoPago.DPaDocQSalda = Recibo.DocCodigo}" _
'            & " Where Documento.DocCodigo = " & CLng(tSerie.Tag)
'
'    If sConCheques Then
'        Cons = Cons & " Group by Documento.DocCodigo, Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor, Recibo.DocSerie , Recibo.DocNumero, Recibo.DocTotal, Documento.DocComentario, Documento.DocCofis "
'    End If
'
'    If crSeteoSqlQuery(jobCredito%, Cons) = 0 Then GoTo ErrCrystal
'
'    Dim iJobSRep1 As Integer, iJobSRep2 As Integer
'
'    'Subreporte srCredito.rpt  y srCredito.rpt - 01-----------------------------------------------------------------------------
'    iJobSRep1 = crAbroSubreporte(jobCredito, "srCredito.rpt")
'    If iJobSRep1 = 0 Then GoTo ErrCrystal
'
'    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
'            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
'                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"
'
'    If crSeteoSqlQuery(iJobSRep1, Cons) = 0 Then GoTo ErrCrystal
'
'    iJobSRep2 = crAbroSubreporte(jobCredito, "srCredito.rpt - 01")
'    If iJobSRep2 = 0 Then GoTo ErrCrystal
'    If crSeteoSqlQuery(iJobSRep2, Cons) = 0 Then GoTo ErrCrystal
'    '-------------------------------------------------------------------------------------------------------------------------------------
'
'    'If crMandoAPantalla(jobCredito, "Factura Credito") = 0 Then GoTo ErrCrystal
'    If crMandoAImpresora(jobCredito, 1) = 0 Then GoTo ErrCrystal
'    If Not crInicioImpresion(jobCredito, True, False) Then GoTo ErrCrystal
'
'    If Not crCierroSubReporte(iJobSRep1) Then GoTo ErrCrystal
'    If Not crCierroSubReporte(iJobSRep2) Then GoTo ErrCrystal
'
'    'crEsperoCierreReportePantalla
'    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroSubReporte JobSRep1
    crCierroSubReporte JobSRep2
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
        Signo = Trim(Rs!MonSigno)
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
    'prj_LoadConfigPrint True
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
    
    'lPNC.Caption = "Imp. Nota:" & paINContadoN
    'If Not paPrintEsXDefNC Then lPNC.ForeColor = &HC0& Else lPNC.ForeColor = vbBlack
    
    'lPCn.Caption = "Imp. Salida Caja: " & paIConformeN
    'If Not paPrintEsXDefCn Then lPCn.ForeColor = &HC0& Else lPNC.ForeColor = vbBlack
    
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

Private Sub ImprimoCreditoComponente()
On Error GoTo errICC
Dim oDocs As clsDocAImprimir
Dim oPrint As New clsImpresionDeDocumentos
    
    Set oPrint.Conexion = cBase
    oPrint.PathReportes = prmPathListados
    'oPrint.NombreBaseDatos = miConexion.RetornoPropiedad(False, False, False, True)
    
    Dim oCnfg As New clsConfigImpresora
    oCnfg.Impresora = paICreditoN
    oCnfg.Bandeja = paICreditoB
    oCnfg.Papel = 1
    Set oPrint.DondeImprimo = oCnfg
            
    oPrint.ImprimoCredito CLng(tSerie.Tag), paDCredito, paDRecibo, paArticuloPisoAgencia, Val(lUsuario.Caption), Val(lVendedor.Caption), labCliente.Caption, IIf(tRuc.Text <> "", tRuc.Text, lCI.Caption), labDireccion.Caption
    
    Exit Sub
    
errICC:
    clsGeneral.OcurrioError "Error al imprimir.", Err.Description, "Imprimo crédito"

End Sub

Private Sub ImprimirContadoYNotas_VSReport(ByVal NombreDocumento As String, ByVal Documento As Long, ByVal TipoDoc As Integer, ByVal TextoDoc As String, ByVal TextoRetira As String)
On Error GoTo ErrCrystal
Dim aTexto As String, sErr As String
    
    Screen.MousePointer = 11
    sErr = "1"
    Dim oImprimo As clsImpresionContado
    Set oImprimo = New clsImpresionContado
    sErr = "2"
    oImprimo.DondeImprimo.Bandeja = paIContadoB
    sErr = "3"
    oImprimo.DondeImprimo.Impresora = paIContadoN
    sErr = "4"
    oImprimo.DondeImprimo.Papel = 1
    oImprimo.PathReportes = prmPathListados
    oImprimo.StringConnect = miConexion.TextoConexion("Comercio")
    
    sErr = "5"
    With oImprimo
        .field_NombreDocumento = paDContado

        aTexto = Trim(labCliente.Caption)
        If lCI.Caption <> "" Then aTexto = aTexto & " (" & Trim(lCI.Caption) & ")"
        .field_ClienteNombre = aTexto

        .field_ClienteDireccion = Trim(labDireccion.Caption)

        If Trim(tRuc.Text) <> "" Then
            .field_RUT = tRuc.FormattedText
            .field_CFinal = ""      'x defexto está en X
        End If

        If TipoDoc = TipoDocumento.Contado Then aTexto = CodigoDeBarras(TipoDoc, CLng(tSerie.Tag)) Else aTexto = ""
        .field_CodigoDeBarras = aTexto
        .field_TextoRetira = TextoRetira
        sErr = "6"
        .ImprimoFacturaContado_VSReport Documento
    End With
    Screen.MousePointer = 0
    
    
'    Dim oPrinter As New clsConfigImpresora
'    oPrinter.Bandeja = paIContadoB
'    oPrinter.Impresora = paIContadoN
'    oPrinter.Papel = 1
'    Set oImprimo.DondeImprimo = oPrinter
'
'
'    oImprimo.PathReportes = prmPathListados
'    oImprimo.StringConnect = cBase.Connect
'
'    'Paso campos de consulta
'    With oImprimo
'        .field_NombreDocumento = paDContado
'
'        aTexto = Trim(labCliente.Caption)
'        If lCI.Caption <> "" Then aTexto = aTexto & " (" & Trim(lCI.Caption) & ")"
'        .field_ClienteNombre = aTexto
'
'        .field_ClienteDireccion = Trim(labDireccion.Caption)
'
'        If Trim(tRuc.Text) <> "" Then
'            .field_RUT = tRuc.FormattedText
'            .field_CFinal = ""      'x defexto está en X
'        End If
'
'        If TipoDoc = TipoDocumento.Contado Then aTexto = CodigoDeBarras(TipoDoc, CLng(tSerie.Tag)) Else aTexto = ""
'        .field_CodigoDeBarras = aTexto
'        .field_TextoRetira = TextoRetira
'
'        .ImprimoFacturaContado_VSReport Documento
'    End With
'    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    On Error Resume Next
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al imprimir el documento, debe reimprimirlo.", Err.Description, "Paso " & sErr
    Exit Sub
End Sub


