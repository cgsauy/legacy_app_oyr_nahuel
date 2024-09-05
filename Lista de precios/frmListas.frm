VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmListas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listas de Precios y Etiquetas"
   ClientHeight    =   4965
   ClientLeft      =   3450
   ClientTop       =   4110
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6240
   Begin VB.PictureBox picLista 
      Height          =   2715
      Index           =   1
      Left            =   600
      ScaleHeight     =   2655
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
      Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
         Height          =   975
         Left            =   60
         TabIndex        =   25
         Top             =   120
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1720
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
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
   End
   Begin VB.PictureBox picLista 
      Height          =   3795
      Index           =   2
      Left            =   120
      ScaleHeight     =   3735
      ScaleWidth      =   5955
      TabIndex        =   27
      Top             =   960
      Width           =   6015
      Begin VB.CommandButton bPrintEtiqueta 
         Caption         =   "Imprimir"
         Height          =   315
         Left            =   4800
         TabIndex        =   20
         Top             =   2880
         Width           =   1035
      End
      Begin VB.CommandButton bFiltrarEtiqueta 
         Caption         =   "..."
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   3480
         Width           =   375
      End
      Begin VB.ComboBox cEtiquetaAImprimir 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton bAgregar 
         Caption         =   "A&gregar"
         Height          =   315
         Left            =   4800
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cQueEtiqueta 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.VScrollBar vscCantidad 
         Height          =   285
         Left            =   4080
         Min             =   1
         TabIndex        =   28
         Top             =   480
         Value           =   1
         Width           =   255
      End
      Begin VB.TextBox tCantidad 
         Height          =   285
         Left            =   3600
         TabIndex        =   15
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox tArticulo 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   120
         Width           =   3735
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsEtiquetaArt 
         Height          =   1935
         Left            =   60
         TabIndex        =   17
         Top             =   840
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3413
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
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
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
      Begin VB.Label lFiltro 
         Caption         =   "Agregar por &Filtros:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   5760
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label8 
         Caption         =   "Et&iqueta a imprimir:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblIDEspecifico 
         Caption         =   "¿C&uales?:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "&Cantidad:"
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblArticulo 
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComctlLib.TabStrip tabLista 
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   660
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Público"
            Key             =   "definidas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Distribuidores"
            Key             =   "varias"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Etiquetas"
            Key             =   "etiquetas"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLista 
      Height          =   2055
      Index           =   0
      Left            =   240
      ScaleHeight     =   1995
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   1800
      Width           =   4455
      Begin VB.CommandButton bContado 
         Caption         =   "Imprimir"
         Height          =   315
         Left            =   3540
         Picture         =   "frmListas.frx":0442
         TabIndex        =   6
         Top             =   60
         Width           =   1095
      End
      Begin VB.ComboBox cCategoria 
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CommandButton bContadoDto 
         Caption         =   "Imprimir"
         Height          =   315
         Left            =   3540
         Picture         =   "frmListas.frx":0974
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría de Cliente:"
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Listas para Distribuidores (Precios Contado con Descuentos)"
         Height          =   255
         Left            =   60
         TabIndex        =   23
         Top             =   660
         Width           =   4815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Listas para Distribuidores (Precios Contado)"
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   120
         Width           =   3315
      End
   End
   Begin VB.TextBox tVigencia 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lSep 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      Height          =   15
      Left            =   120
      TabIndex        =   26
      Top             =   480
      Width           =   5895
   End
   Begin VB.Label lVigencia 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de &Vigencia:"
      Height          =   255
      Left            =   2820
      TabIndex        =   2
      Top             =   180
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de &Vigencia:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1755
   End
End
Attribute VB_Name = "frmListas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer

Private Type typCombo
    Articulo As Long
    Q As Integer
    Bonificacion As Currency
    EsBonificacion As Boolean
End Type

Enum TReporte
    Contado = 1
    ContadoConDto = 2
    AlPublico = 3
    EtiquetaNormal = 4
    EtiquetaConArgumento = 5
    EtiquetaSinArgumento = 6
End Enum

Private Type tDatosArt
    IDEsp As Long
    VariacionEsp As Currency
    esCombo As Boolean
End Type
Dim artInfo As tDatosArt

'Variables para Crystal Engine.---------------------------------
Private Result As Integer, JobSRep1 As Integer, JobSRep2 As Integer, jobnum As Integer
Private NombreFormula As String, CantForm As Integer, aTexto As String

Private prmVigencia As String

Private Sub ImprimoEtiquetaConArgumento()
    Dim frmR As New frmReport
    Dim sTotFinanciado As String
    If cEtiquetaAImprimir.ListIndex = 0 Then
        frmR.ReportFile = "rptEtiquetaNormal.xml"
        frmR.ReportName = "EtiquetaNormal"
        sTotFinanciado = "'Esp.: ' + RTRIM(Cast(AEsID as varchar(10))) + char(13) + IsNull(AEsComentario COLLATE Modern_Spanish_CI_AI, '')"
    Else
        frmR.ReportFile = "rptEtiquetaArgumento.xml"
        frmR.ReportName = "EtiquetaArgumento"
        sTotFinanciado = "EImTotalFinanciado"
    End If
    frmR.ReportQuery = Replace("SELECT ArtCodigo CodigoArticulo, '' as CodigoEspecifico, ArticuloFacturacion.AFaArgumLargo, IsNull(ArticuloFacturacion.AFaComentarioA, '') ComentarioArt, Garantia.GarNombre, ISNULL(AFaAlto, 0) AFaAlto, IsNull(AFaFrente, 0) AFaFrente, IsNull(AFaProfundidad,0) AFaProfundidad, IsNull(ArticuloWebPage.AWPNombreArt, Articulo.ArtNombre) ArtNombre, RTRIM(EImImporteCtdo) PrecioCtdo, RTRIM(EImPlan) PlanLetra, EImCuota Cuotas, EImTotalFinanciado TotalFinanciado, '' Descuento " _
            & "FROM Articulo INNER JOIN EtiquetaAImprimir ON EImArticulo = ArtId INNER JOIN ArticuloFacturacion ON ArtId = AFaArticulo INNER JOIN ArticuloWebPage ON ArtId = AWPArticulo " _
            & "LEFT OUTER JOIN CGSA.dbo.Garantia Garantia ON ArticuloFacturacion.AFaGarantia=Garantia.GarCodigo Union All " _
            & "SELECT Articulo.ArtCodigo, 'Esp.: ' + RTRIM(Cast(AEsID as varchar(10))), ArticuloFacturacion.AFaArgumLargo, IsNull(AEsComentario COLLATE Modern_Spanish_CI_AI, ''), Garantia.GarNombre, ISNULL(AFaAlto, 0) AFaAlto, IsNull(AFaFrente, 0) AFaFrente, IsNull(AFaProfundidad,0) AFaProfundidad, IsNull(ArticuloWebPage.AWPNombreArt, Articulo.ArtNombre) ArtNombre, RTRIM(EImImporteCtdo), RTRIM(EImPlan), EImCuota Cuotas, [fieldTotalFinanciado] TotalFinanciado, 'Desc.: $ ' + CAST(ABS(AEsVariacionPrecio) AS VARCHAR(10)) Descuento " _
            & "FROM Articulo INNER JOIN ArticuloEspecifico ON ArtId = AEsArticulo INNER JOIN EtiquetaAImprimir ON RTrim(CAST(EImArticulo as varchar(15))) = RTrim(CAST(ArtId as varchar(10))) + RTrim(CAST(AEsID as varchar(10))) " _
            & "INNER JOIN ArticuloFacturacion ON ArtId = AFaArticulo INNER JOIN ArticuloWebPage ON ArtId = AWPArticulo LEFT OUTER JOIN CGSA.dbo.Garantia Garantia ON ArticuloFacturacion.AFaGarantia=Garantia.GarCodigo ", "[fieldTotalFinanciado]", sTotFinanciado)
    frmR.Show (vbModal)
End Sub

Private Sub AccionImprimir(idReporte As Integer)

    If Not IsDate(tVigencia.Text) Then
        MsgBox "La fecha de vigencia no es correcta.", vbExclamation, "Datos Incorrectos"
        tVigencia.SetFocus: Exit Sub
    End If

    prmVigencia = Format(tVigencia.Text, "mm/dd/yyyy 23:59:59")
    
    Select Case idReporte
        Case TReporte.Contado
                        If InicializoReporteEImpresora("", 1, "lprListaContado.RPT") Then Exit Sub
                        rptListaContado
                        If Not crCierroTrabajo(jobnum) Then MsgBox crMsgErr
        
        Case TReporte.ContadoConDto
                        If InicializoReporteEImpresora("", 1, "lprListaConDescuentos.RPT") Then Exit Sub
                        rptListaContadoCategoria
                        If Not crCierroTrabajo(jobnum) Then MsgBox crMsgErr

        Case TReporte.AlPublico
                        rptListaAlPublico
                        
        Case TReporte.EtiquetaNormal
                        If InicializoReporteEImpresora("", 1, "lprEtiquetaNormal.RPT") Then Exit Sub
                        rptImprimoEtiquetas "Etiquetas Normales"
                        If Not crCierroTrabajo(jobnum) Then MsgBox crMsgErr
                    
        Case TReporte.EtiquetaConArgumento
                        If InicializoReporteEImpresora("", 1, "lprEtiquetaArgumento.RPT", 2) Then Exit Sub
                        rptImprimoEtiquetas " Etiquetas Vidriera con Argumento"
                        If Not crCierroTrabajo(jobnum) Then MsgBox crMsgErr
        
        Case TReporte.EtiquetaSinArgumento
                        If InicializoReporteEImpresora("", 1, "lprEtiquetaSinArgumento.RPT", 2) Then Exit Sub
                        rptImprimoEtiquetas " Etiquetas de Vidriera"
                        If Not crCierroTrabajo(jobnum) Then MsgBox crMsgErr
    End Select
    
    Screen.MousePointer = 0
End Sub

Private Sub bAgregar_Click()
    If Val(tArticulo.Tag) > 0 And cQueEtiqueta.ListIndex <> -1 Then
        
        If tCantidad.Enabled Then
            If Val(tCantidad.Text) = 0 Then
                MsgBox "La cantidad no puede ser cero.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        
        etiqueta_AgregoArticuloALista Val(tArticulo.Tag), Trim(tArticulo.Text), tCantidad.Text, cQueEtiqueta.ListIndex, artInfo.IDEsp, artInfo.VariacionEsp, artInfo.esCombo
        LimpioIngresoArticuloEtiqueta
        tArticulo.SetFocus
        
    Else
        MsgBox "Falta ingresar algún dato.", vbExclamation, "ATENCIÓN"
    End If
End Sub

Private Sub bContado_Click()
    AccionImprimir TReporte.Contado
End Sub

Private Sub bContadoDto_Click()
    
    If cCategoria.ListIndex = -1 Then
        MsgBox "Seleccione la categoría de cliente para sacar la lista de precios.", vbExclamation, "Falta Categoría de Cliente"
        tVigencia.SetFocus: Exit Sub
    End If

    AccionImprimir TReporte.ContadoConDto
    
End Sub

Private Sub bFiltrarEtiqueta_Click()
On Error GoTo errBFE
Dim lEsta As Long
Dim iCant As Integer
Dim sNombre As String, lid As Long, esCombo As Boolean
    frmFiltroEtiqueta.Show vbModal, Me
    If frmFiltroEtiqueta.prmHayDatos Then
        Screen.MousePointer = 11
        For iCant = 1 To frmFiltroEtiqueta.prmCantResultado
            BuscoArticuloPorCodigo frmFiltroEtiqueta.prmIDResultado(iCant), sNombre, lid, esCombo
            If lid > 0 Then
                For lEsta = 1 To vsEtiquetaArt.Rows - 1
                    If Val(vsEtiquetaArt.Cell(flexcpData, lEsta, 0)) = lid Then
                        lid = 0: Exit For
                    End If
                Next
            End If
            If lid > 0 Then
                etiqueta_AgregoArticuloALista lid, sNombre, frmFiltroEtiqueta.prmCantidad, frmFiltroEtiqueta.prmQueEtiqueta, 0, 0, esCombo
            End If
        Next
        Screen.MousePointer = 0
    End If
    Set frmFiltroEtiqueta = Nothing
    Screen.MousePointer = 0
    Exit Sub
errBFE:
    clsGeneral.OcurrioError "Ocurrió un error al filtrar.", Err.Description, "Error (filtraretiqueta)"
    Screen.MousePointer = 0
End Sub

Private Sub bPrintEtiqueta_Click()
    
    If vsEtiquetaArt.Rows = 1 Then
        MsgBox "No hay artículos ingresados.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    If cEtiquetaAImprimir.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de etiqueta que desea imprimir.", vbExclamation, "ATENCIÓN"
        cEtiquetaAImprimir.SetFocus
    Else
        Screen.MousePointer = 11
        etiqueta_MandoAImprimir
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cEtiquetaAImprimir_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then bPrintEtiqueta.SetFocus
End Sub

Private Sub cQueEtiqueta_Change()
On Error Resume Next
    If cQueEtiqueta.ListIndex > -1 Then
        If cQueEtiqueta.ListIndex = 2 Then
            tCantidad.Enabled = False
            tCantidad.BackColor = vbButtonFace
        Else
            tCantidad.Enabled = True
            tCantidad.BackColor = vbWindowBackground
        End If
    Else
        tCantidad.Enabled = False
        tCantidad.BackColor = vbButtonFace
    End If
    vscCantidad.Enabled = tCantidad.Enabled
End Sub

Private Sub cQueEtiqueta_Click()
On Error Resume Next
    If cQueEtiqueta.ListIndex > -1 Then
        If cQueEtiqueta.ListIndex = 2 Then
            tCantidad.Enabled = False
            tCantidad.BackColor = vbButtonFace
        Else
            tCantidad.Enabled = True
            tCantidad.BackColor = vbWindowBackground
        End If
    Else
        tCantidad.Enabled = False
        tCantidad.BackColor = vbButtonFace
    End If
    vscCantidad.Enabled = tCantidad.Enabled
End Sub

Private Sub cQueEtiqueta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tArticulo.Tag) > 0 Then
            If tCantidad.Enabled Then
                tCantidad.SetFocus
            Else
                bAgregar.SetFocus
            End If
        Else
            vsEtiquetaArt.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    Me.Refresh
End Sub

Private Sub Form_Load()

    On Error Resume Next
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    crAbroEngine
    tVigencia.Text = Format(Now, "dd/mm/yyyy")
    lVigencia.Caption = Format(Now, "Ddd d/Mmm/yyyy")
    
    Cons = "Select LDiCodigo, LDiNombre from ListasDistribuidores order by LDiNombre"
    CargoCombo Cons, cCategoria
    
    InicializoGrillas
    InicializoObjetosEtiqueta
    
    picLista(1).ZOrder 0
    tabLista.SetFocus
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With tabLista
        .Left = 60
        .Width = Me.ScaleWidth - (.Left * 2)
        .Top = tVigencia.Top + 600
        .Height = Me.ScaleHeight - .Top - 60
    End With
    
    With lSep
        .Left = tabLista.Left
        .Width = tabLista.Width
        .Top = tabLista.Top - 150
    End With
    
    For I = picLista.LBound To picLista.UBound
        With picLista(I)
            .Left = tabLista.ClientLeft
            .Top = tabLista.ClientTop
            .Width = tabLista.ClientWidth
            .Height = tabLista.ClientHeight
            .BorderStyle = 0
        End With
    Next
    
    With vsLista
        .Top = 60: .Left = 60
        .Width = picLista(0).ScaleWidth - (.Left * 2)
        .Height = picLista(0).ScaleHeight
    End With
    
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

Private Sub rptListaContado()
On Error GoTo ErrCrystal
Dim I As Integer
Dim dRige As Date

    Screen.MousePointer = 11
    
    'Saco la máxima fecha de vigencia para cargar valor Rige
    dRige = Now
    
    Cons = "Select Max(Precios.HPrVigencia) from HistoriaPrecio Precios, Articulo, ArticuloFacturacion " & _
               " Where (Precios.HPrVigencia IN " & _
                    " (Select MAX(H.HPrVigencia)" & _
                            " FROM HistoriaPrecio H" & _
                            " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                            " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                            " And H.HPrMoneda = Precios.HPrMoneda " & _
                            " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                            " )) " & _
                " And Precios.HPrArticulo = ArtID" & _
                " And ArtId = AFaArticulo " & _
                " And ArtEnUso = 1 And AFaInterior = 1 " & _
                " And Precios.HPrTipoCuota = " & paTipoCuotaContado & _
                " And Precios.HPrMoneda = " & paMonedaPesos & _
                " And Precios.HPrHabilitado = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then dRige = RsAux(0)
    RsAux.Close
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte Contado --------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "prmvigencia": Result = crSeteoFormula(jobnum%, NombreFormula, "'Rige desde el: " & Format(dRige, "dd/Mmm/yyyy") & "'")
            
            Case Else: Result = 1
        End Select
        If Result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "Select Articulo.ArtCodigo, Articulo.ArtNombre, Articulo.ArtHabilitado," & _
                        " ArticuloFacturacion.AFaInterior, ArticuloFacturacion.AFaArgumCorto," & _
                        " Precios.HPrPrecio, Especie.EspNombre" & _
               " From " & _
                        paBD & ".dbo.HistoriaPrecio Precios, " & _
                        paBD & ".dbo.Articulo, " & paBD & ".dbo.Tipo, " & paBD & ".dbo.ArticuloFacturacion, " & paBD & ".dbo.Especie" & _
                " WHERE (HPrVigencia IN" & _
                            " (Select MAX(H.HPrVigencia)" & _
                            " FROM HistoriaPrecio H" & _
                            " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                            " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                            " And H.HPrMoneda = Precios.HPrMoneda " & _
                            " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                            " )) " & _
                " And Precios.HPrArticulo = ArtID" & _
                " And ArtId = AFaArticulo " & _
                " And ArtTipo = TipCodigo " & _
                " And TipEspecie = EspCodigo" & _
                " And ArtEnUso = 1 And AFaInterior = 1 " & _
                " And Precios.HPrTipoCuota = " & paTipoCuotaContado & _
                " And Precios.HPrMoneda = " & paMonedaPesos & _
                " And Precios.HPrHabilitado = 1"
                
                Cons = Cons & " Order By Especie.EspNombre Asc, Articulo.ArtNombre Asc"
    
    Cons = Trim(Cons) & Chr$(0)
    
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------

    If crMandoAPantalla(jobnum, "Lista de Contados") = 0 Then GoTo ErrCrystal
    'If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
'    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
    
'    crEsperoCierreReportePantalla
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

Private Sub rptListaContadoCategoria()
On Error GoTo ErrCrystal
Dim dRige As Date
Dim paCategoriaCliente As Long
Dim paTipoCuota As Long

Dim bHay As Boolean

    Screen.MousePointer = 11
    
    Cons = "Select * from ListasDistribuidores Where LDICodigo = " & cCategoria.ItemData(cCategoria.ListIndex)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!LDiCatCliente) Then paCategoriaCliente = RsAux!LDiCatCliente
        If Not IsNull(RsAux!LDiTipoCuota) Then paTipoCuota = RsAux!LDiTipoCuota
    End If
    RsAux.Close
    
    'Saco la máxima fecha de vigencia para cargar valor Rige
    dRige = Now
    
    Cons = "Select Max(Precios.HPrVigencia) from HistoriaPrecio Precios, Articulo, ArticuloFacturacion, CategoriaDescuento " & _
               " Where (Precios.HPrVigencia IN " & _
                    " (Select MAX(H.HPrVigencia)" & _
                            " FROM HistoriaPrecio H" & _
                            " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                            " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                            " And H.HPrMoneda = Precios.HPrMoneda " & _
                            " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                            " )) " & _
                " And Precios.HPrArticulo = ArtID" & _
                " And ArtId = AFaArticulo " & _
                " And ArtEnUso = 1 And AFaInterior = 1 " & _
                " And Precios.HPrTipoCuota = " & paTipoCuotaContado & _
                " And Precios.HPrMoneda = " & paMonedaPesos & _
                " And Precios.HPrHabilitado = 1" & _
                " And AFaCategoriaD = CDtCatArticulo" & _
                " And CDtCatCliente = " & paCategoriaCliente & _
                " And CDtCatPlazo = " & paTipoCuota
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    bHay = True
    If Not RsAux.EOF Then
        If Not IsNull(RsAux(0)) Then dRige = RsAux(0) Else bHay = False
    End If
    RsAux.Close
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    If Not bHay Then
        MsgBox "No hay precios vigentes para la categoría seleccionada.", vbExclamation, "No hay Datos"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte Contado --------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "prmvigencia": Result = crSeteoFormula(jobnum%, NombreFormula, "'Rige desde el: " & Format(dRige, "dd/Mmm/yyyy") & "'")
            Case "prmlista": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(cCategoria.Text) & "'")
            
            Case Else: Result = 1
        End Select
        If Result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "Select Articulo.ArtCodigo, Articulo.ArtNombre, Articulo.ArtHabilitado," & _
                        " ArticuloFacturacion.AFaInterior, ArticuloFacturacion.AFaArgumCorto," & _
                        " Precios.HPrPrecio, Especie.EspNombre" & _
               " From " & _
                        paBD & ".dbo.HistoriaPrecio Precios, " & paBD & ".dbo.Articulo, " & _
                        paBD & ".dbo.Tipo, " & paBD & ".dbo.ArticuloFacturacion, " & paBD & ".dbo.Especie, " & paBD & ".dbo.CategoriaDescuento " & _
                " WHERE (HPrVigencia IN" & _
                            " (Select MAX(H.HPrVigencia)" & _
                            " FROM HistoriaPrecio H" & _
                            " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                            " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                            " And H.HPrMoneda = Precios.HPrMoneda " & _
                            " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                            " )) " & _
                " And Precios.HPrArticulo = ArtID" & _
                " And ArtId = AFaArticulo " & _
                " And ArtTipo = TipCodigo " & _
                " And TipEspecie = EspCodigo" & _
                " And ArtEnUso = 1 And AFaInterior = 1 " & _
                " And Precios.HPrTipoCuota = " & paTipoCuotaContado & _
                " And Precios.HPrMoneda = " & paMonedaPesos & _
                " And Precios.HPrHabilitado = 1" & _
                " And AFaCategoriaD = CDtCatArticulo" & _
                " And CDtCatCliente = " & paCategoriaCliente & _
                " And CDtCatPlazo = " & paTipoCuota
                
                '" Order By Especie.EspNombre Asc, Articulo.ArtNombre Asc"
    
    Cons = Trim(Cons) & Chr$(0)
    
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------

    If crMandoAPantalla(jobnum, "Lista Distribuidores") = 0 Then GoTo ErrCrystal
    'If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
'    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
    
'    crEsperoCierreReportePantalla
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

Private Function InicializoReporteEImpresora(paNImpresora As String, paBImpresora As Integer, Reporte As String, Optional Orientacion As Integer = 1) As Boolean
On Error GoTo ErrCrystal
    
    jobnum = crAbroReporte(prmPathListados & Reporte)
    If jobnum = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    'If Trim(Printer.DeviceName) <> Trim(paNImpresora) Then SeteoImpresoraPorDefecto paNImpresora
    If Not crSeteoImpresora(jobnum, Printer, paBImpresora, Orientacion) Then GoTo ErrCrystal
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

Private Sub tabLista_Click()

    Select Case tabLista.SelectedItem.Key
        Case "definidas": picLista(1).ZOrder 0
        Case "varias": picLista(0).ZOrder 0
        Case "etiquetas": picLista(2).ZOrder 0
    End Select
    
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = ""
    artInfo.IDEsp = 0
    artInfo.VariacionEsp = 0
    artInfo.esCombo = False
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Select Case Val(lblArticulo.Tag)
            Case 0:
                lblArticulo.Tag = "1"
                lblArticulo.Caption = "Específico"
            Case 1:
                lblArticulo.Tag = "0"
                lblArticulo.Caption = "Artículo"
        End Select
    End If
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrAP
Dim lEsta As Long, lid As Long
Dim sNombre As String
    
    If KeyAscii = vbKeyReturn Then
       
        
        If Val(tArticulo.Tag) <> 0 Then
            'cQueEtiqueta.ListIndex = 2
            cQueEtiqueta.SetFocus: Exit Sub
        End If
        
        If Val(lblArticulo.Tag) = 1 Then
            loc_InsertArticuloEspecifico
            Exit Sub
        End If
        
        Screen.MousePointer = 11
        If Trim(tArticulo.Text) <> "" Then
            
            Dim bEsCombo As Boolean
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo tArticulo.Text, sNombre, lid, bEsCombo
                If lid > 0 Then
                    tArticulo.Text = sNombre
                    tArticulo.Tag = lid
                    artInfo.esCombo = bEsCombo
                ElseIf lid = -1 Then
                    'No tiene precio
                    tArticulo.Tag = "0"
                Else
                    tArticulo.Tag = "0"
                    MsgBox "No existe un artículo que posea ese código.", vbExclamation, "ATENCIÓN"
                End If
            Else
                lid = BuscoArticuloPorNombre(tArticulo.Text)
                BuscoArticuloPorCodigo lid, sNombre, lid, bEsCombo
                If lid > 0 Then
                    tArticulo.Text = sNombre
                    tArticulo.Tag = lid
                    artInfo.esCombo = bEsCombo
                Else
                    'No tiene precio
                    tArticulo.Tag = "0"
                End If
            End If
            If Val(tArticulo.Tag) > 0 Then
                For lEsta = 1 To vsEtiquetaArt.Rows - 1
                    If Val(vsEtiquetaArt.Cell(flexcpData, lEsta, 0)) = Val(tArticulo.Tag) And Val(vsEtiquetaArt.Cell(flexcpData, lEsta, 1)) = 0 Then
                        MsgBox "El artículo ya esta ingresado, edite la columna de cantidades si desea modificarlas.", vbInformation, "ATENCIÓN"
                        vsEtiquetaArt.Select lEsta, 0, lEsta, vsEtiquetaArt.Cols - 1
                        vsEtiquetaArt.SetFocus
                        tArticulo.Text = ""
                        tArticulo.Tag = ""
                        Exit Sub
                    End If
                Next
                'cQueEtiqueta.ListIndex = 2
                cQueEtiqueta.SetFocus
            End If
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrAP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCantidad_GotFocus()
    With tCantidad
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    If Val(tCantidad.Text) = 0 Then tCantidad.Text = vscCantidad.Value
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCantidad.Text) Then
            If Val(tCantidad.Text) < 1 Then tCantidad.Text = vscCantidad.Value
            vscCantidad.Value = Val(tCantidad.Text)
            bAgregar_Click
        Else
            MsgBox "Formato incorrecto.", vbExclamation, "ATENCIÓN"
            tCantidad.Text = vscCantidad.Value
        End If
    End If
End Sub

Private Sub tCantidad_LostFocus()
    If Not IsNumeric(tCantidad.Text) Then
        tCantidad.Text = vscCantidad.Value
    Else
        If Val(tCantidad.Text) < 1 Then tCantidad.Text = vscCantidad.Value
    End If
End Sub

Private Sub tVigencia_Change()
    lVigencia.Caption = ""
End Sub

Private Sub tVigencia_GotFocus()
    tVigencia.Appearance = 1
    tVigencia.BackColor = vbWindowBackground
    tVigencia.SelStart = 0: tVigencia.SelLength = Len(tVigencia.Text)
End Sub

Private Sub tVigencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tVigencia.Text) Then
            tVigencia.Text = Format(tVigencia, "dd/mm/yyyy")
            lVigencia.Caption = Format(tVigencia, "Ddd d/Mmm/yyyy")
        Else
            lVigencia.Caption = "#Error"
        End If
    End If
End Sub

Private Sub tVigencia_LostFocus()
    tVigencia.BackColor = vbButtonFace
End Sub



Private Sub ImprimoContadoII()
On Error GoTo ErrCrystal
Dim I As Integer
Dim dRige As Date

    Screen.MousePointer = 11
    
    'Saco la máxima fecha de vigencia para cargar valor Rige
    dRige = Now
    
    Cons = "Select Max(Precios.HPrVigencia) from HistoriaPrecio Precios, Articulo, ArticuloFacturacion " & _
               " Where (Precios.HPrVigencia IN " & _
                    " (Select MAX(H.HPrVigencia)" & _
                            " FROM HistoriaPrecio H" & _
                            " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                            " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                            " And H.HPrMoneda = Precios.HPrMoneda " & _
                            " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                            " )) " & _
                " And Precios.HPrArticulo = ArtID" & _
                " And ArtId = AFaArticulo " & _
                " And ArtEnUso = 1 And AFaInterior = 1 " & _
                " And Precios.HPrMoneda = " & paMonedaPesos & _
                " And Precios.HPrHabilitado = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then dRige = RsAux(0)
    RsAux.Close
    
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte Contado --------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "prmvigencia": Result = crSeteoFormula(jobnum%, NombreFormula, "'Rige desde el: " & Format(dRige, "dd/Mmm/yyyy") & "'")
            
            Case Else: Result = 1
        End Select
        If Result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "Select * " & _
               " From " & _
                        paBD & ".dbo.HistoriaPrecio Precios, " & _
                        paBD & ".dbo.Articulo, " & paBD & ".dbo.Tipo, " & paBD & ".dbo.ArticuloFacturacion, " & paBD & ".dbo.Especie, " & _
                        paBD & ".dbo.ListasDePrecios, " & paBD & ".dbo.TipoCuota, " & paBD & ".dbo.TipoPlan" & _
                " WHERE (HPrVigencia IN" & _
                            " (Select MAX(H.HPrVigencia)" & _
                            " FROM HistoriaPrecio H" & _
                            " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                            " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                            " And H.HPrMoneda = Precios.HPrMoneda " & _
                            " And H.HPrVigencia <= '" & prmVigencia & "'" & _
                            " )) " & _
                " And Precios.HPrArticulo = ArtID" & _
                " And ArtId = AFaArticulo " & _
                " And ArtTipo = TipCodigo " & _
                " And TipEspecie = EspCodigo" & _
                " And ArtEnUso = 1 And AFaInterior = 1 " & _
                " And Precios.HPrMoneda = " & paMonedaPesos & _
                " And Precios.HPrHabilitado = 1" & _
                " And ArticuloFacturacion.AFaLista = ListasDePrecios.LDPCodigo " & _
                " And Precios.HPrTipoCuota = TipoCuota.TCuCodigo" & _
                " And Precios.HPrPlan = TipoPlan.PlaCodigo"
                
    Cons = Cons & _
                 " AND LDPNumero = 1 " & _
                 " And TipoCuota.TCuVencimientoE is null " & _
                 " And TCuEspecial = 0 " & _
                 " and TCuDeshabilitado is null"
                '" Order By Especie.EspNombre Asc, Articulo.ArtNombre Asc"
    
    Cons = Trim(Cons) & Chr$(0)
    
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------

    If crMandoAPantalla(jobnum, "Lista de Precios") = 0 Then GoTo ErrCrystal
    'If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
'    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
    
'    crEsperoCierreReportePantalla
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


Private Sub InicializoGrillas()

    On Error Resume Next
    With vsLista
        .Rows = 1: .Cols = 1
        .Editable = False
        .FormatString = ">Nº|<Listas"
        .WordWrap = False
        .ColWidth(0) = 500: .ColWidth(1) = 2100
        .ExtendLastCol = True
    End With
    
    Dim aValor As Long
    Cons = "Select * from ListasDePrecios order by LDPNumero"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsLista
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = RsAux!LDPNumero
            aValor = RsAux!LDPCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!LDPNombre)
        End With
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

Private Sub vscCantidad_Change()
    tCantidad.Text = vscCantidad.Value
End Sub

Private Sub vsEtiquetaArt_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    AjustoTotalEtiqueta
End Sub

Private Sub vsEtiquetaArt_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    
    If vsEtiquetaArt.IsSubtotal(Row) Or Col = 0 Then Cancel = True: Exit Sub
    If Col = 2 And vsEtiquetaArt.Cell(flexcpText, Row, 4) = "" Then Cancel = True: Exit Sub
    If Col = 3 And vsEtiquetaArt.Cell(flexcpText, Row, 4) <> "" Then Cancel = True: Exit Sub
    
End Sub

Private Sub vsEtiquetaArt_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsEtiquetaArt.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyDelete
            If Not vsEtiquetaArt.IsSubtotal(vsEtiquetaArt.Row) Then
                vsEtiquetaArt.RemoveItem vsEtiquetaArt.Row
                AjustoTotalEtiqueta
            End If
        Case vbKeyReturn: cEtiquetaAImprimir.SetFocus
    End Select
End Sub

Private Sub vsEtiquetaArt_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If vsEtiquetaArt.EditText = "" Then
        vsEtiquetaArt.EditText = "0"
    Else
        If Not IsNumeric(vsEtiquetaArt.EditText) Then
            Cancel = True
            MsgBox "Formato inválido.", vbExclamation, "ATENCIÓN"
        Else
            If Val(vsEtiquetaArt.EditText) < 0 Then
                Cancel = True
                MsgBox "La cantidad tiene que ser mayor o igual a cero.", vbExclamation, "ATENCIÓN"
            End If
        End If
    End If
End Sub

Private Sub vsLista_DblClick()
    If vsLista.Rows = 1 Then Exit Sub
    AccionImprimir TReporte.AlPublico
End Sub

Private Sub rptListaAlPublico()
On Error GoTo errSel

    If vsLista.Rows = 1 Then Exit Sub
    Dim miF As New frmPreview
    
    With miF
        .prmHeaderReport = "Lista Nº " & vsLista.Cell(flexcpText, vsLista.Row, 0) & ": " & vsLista.Cell(flexcpText, vsLista.Row, 1)
        .prmCaption = "Listas de Precios Público"
        .prmIDLista = vsLista.Cell(flexcpData, vsLista.Row, 0)
        .prmMonedaPesos = paMonedaPesos
        .prmVigencia = prmVigencia
        .Show
    End With
    
    Set miF = Nothing
    Exit Sub

errSel:
    clsGeneral.OcurrioError "Error al activar la lista de precios.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoArticuloPorCodigo(ByVal lCodArticulo As Long, ByRef sNombre As String, ByRef lIDArt As Long, ByRef esCombo As Boolean)
On Error GoTo errBA

    Screen.MousePointer = 11
    sNombre = ""
    lIDArt = 0
    esCombo = False
    Cons = "Select * From Articulo Where ArtCodigo = " & lCodArticulo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
    Else
        esCombo = RsAux("ArtEsCombo")
        sNombre = Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
        lIDArt = RsAux!ArtID
        RsAux.Close
        
        If Not esCombo Then
            If Not ArticuloTienePrecio(lIDArt) Then
                MsgBox "El artículo no posee precios.", vbInformation, "ATENCIÓN"
                lIDArt = -1
                sNombre = ""
            End If
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
errBA:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo por código.", Err.Description, "Error (buscoarticuloporcodigo)"
    Screen.MousePointer = 0
End Sub

Private Function BuscoArticuloPorNombre(NomArticulo As String) As Long
On Error GoTo errBA
Dim lResultado As Long
    Screen.MousePointer = 11
    BuscoArticuloPorNombre = 0
    Cons = "Select ArtCodigo, ArtCodigo as 'Código', ArtNombre as 'Nombre' From Articulo" _
        & " Where ArtNombre LIKE '" & Replace(NomArticulo, " ", "%") & "%'" _
        & " Order By ArtNombre"
            
    Dim objAyuda As New clsListadeAyuda
    If objAyuda.ActivarAyuda(cBase, Cons, 5000, 1, "Lista de Artículos") Then
        lResultado = objAyuda.RetornoDatoSeleccionado(1)
    Else
        lResultado = 0
    End If
    Set objAyuda = Nothing       'Destruyo la clase.
    Screen.MousePointer = 11
    If lResultado > 0 Then BuscoArticuloPorNombre = lResultado
    Screen.MousePointer = 0
    Exit Function
errBA:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo por código.", Err.Description, "Error (buscoarticuloporcodigo)"
    Screen.MousePointer = 0
End Function

Private Sub InicializoObjetosEtiqueta()
    LimpioIngresoArticuloEtiqueta
    With vsEtiquetaArt
        .Rows = 1
        .Cols = 1
        .FormatString = "<Artículo|>Q Normal|>Q c/Argum.|>Q s/Argum.|Arg.Largo|Medida"
        .ColWidth(0) = 3000
        .ColHidden(4) = True
        .ColHidden(5) = True
        .Editable = True
    End With
    'Cargo combos
    With cQueEtiqueta
        .Clear
        .AddItem "Ambas"
        .AddItem "Normal (chica)"
        .AddItem "Según tabla"
        .AddItem "Vidriera (grande)"
        .ListIndex = 2
    End With
    With cEtiquetaAImprimir
        .Clear
        .AddItem "Normales"
        .AddItem "Vidriera c/Argumento"
        .AddItem "Vidriera s/Argumento"
    End With
End Sub

Private Sub LimpioIngresoArticuloEtiqueta()
    
    tArticulo.Text = ""
    artInfo.IDEsp = 0
    artInfo.VariacionEsp = 0
    artInfo.esCombo = False
    vscCantidad.Value = 1
    tCantidad.Text = vscCantidad.Value

End Sub

Private Sub etiqueta_AgregoArticuloALista(ByVal lIDArticulo As Long, ByVal sNombre As String, ByVal iCantidad As Integer, ByVal iQueEtiqueta As Integer, ByVal IDEsp As Long, ByVal VarPrecioEsp As Currency, ByVal esCombo As Boolean)
On Error GoTo errAA
Dim sArgumLargo As String, sMedida As String
Dim lQEN As Long, lQEV As Long
        
    'Tengo que buscar en la tabla artículo facturación si posee argumento largo.
    GetArticuloArgumLargoMedidas lIDArticulo, sArgumLargo, sMedida, lQEN, lQEV
    
   
    If iQueEtiqueta <> 2 Then
        lQEN = iCantidad
        lQEV = iCantidad
        If lQEN = 0 And lQEV = 0 Then Exit Sub
    End If
    
    With vsEtiquetaArt
        .AddItem sNombre
        .Cell(flexcpData, .Rows - 1, 0) = lIDArticulo
        .Cell(flexcpData, .Rows - 1, 1) = IDEsp
        .Cell(flexcpData, .Rows - 1, 2) = VarPrecioEsp
        .Cell(flexcpData, .Rows - 1, 3) = CStr(IIf(esCombo, 1, 0))
        
        Select Case iQueEtiqueta
            Case 0, 2
                .Cell(flexcpText, .Rows - 1, 1) = lQEN
                If Trim(sArgumLargo) <> "" Then
                    .Cell(flexcpText, .Rows - 1, 2) = lQEV
                    .Cell(flexcpText, .Rows - 1, 3) = "0"
                Else
                    .Cell(flexcpText, .Rows - 1, 2) = "0"
                    .Cell(flexcpText, .Rows - 1, 3) = lQEV
                End If
            
            Case 1
                .Cell(flexcpText, .Rows - 1, 1) = lQEN
                .Cell(flexcpText, .Rows - 1, 2) = "0"
                .Cell(flexcpText, .Rows - 1, 3) = "0"
                
            Case 3
                .Cell(flexcpText, .Rows - 1, 1) = 0
                If Trim(sArgumLargo) <> "" Then
                    .Cell(flexcpText, .Rows - 1, 2) = lQEV
                    .Cell(flexcpText, .Rows - 1, 3) = "0"
                Else
                    .Cell(flexcpText, .Rows - 1, 2) = "0"
                    .Cell(flexcpText, .Rows - 1, 3) = lQEV
                End If
        End Select
        .Cell(flexcpText, .Rows - 1, 4) = Trim(sArgumLargo)
        .Cell(flexcpText, .Rows - 1, 5) = sMedida
    End With
    AjustoTotalEtiqueta
    Exit Sub
errAA:
    clsGeneral.OcurrioError "Ocurrió un error al intentar agregar el artículo a la lista.", Err.Description, "Error (agregoarticuloalista)"
End Sub

Private Sub GetArticuloArgumLargoMedidas(ByVal lIDArticulo As Long, ByRef sArgum As String, _
                                ByRef sMedida As String, ByRef lQEN As Long, ByRef lQEV As Long)
                                
On Error GoTo errGA
Dim cAlto As Currency, cFrente As Currency, cProf As Currency

    sArgum = ""
    sMedida = ""
    lQEN = 0: lQEV = 0
    
    Cons = "Select * From ArticuloFacturacion Where AFaArticulo = " & lIDArticulo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        
        If Not IsNull(RsAux!AFAEtNormales) Then lQEN = RsAux!AFAEtNormales
        If Not IsNull(RsAux!AFAEtVidriera) Then lQEV = RsAux!AFAEtVidriera
                
        If Not IsNull(RsAux!AFaArgumLargo) Then
            sArgum = Trim(RsAux!AFaArgumLargo)
            If Not sArgum Like "*[0-z]*" Then sArgum = ""
        End If
        
        '-----------------------------------------------------------------------
        cAlto = 0: cFrente = 0: cProf = 0
        If Not IsNull(RsAux!AFaAlto) Then cAlto = Trim(RsAux!AFaAlto)
        If Not IsNull(RsAux!AFaFrente) Then cFrente = Trim(RsAux!AFaFrente)
        If Not IsNull(RsAux!AFaProfundidad) Then cProf = RsAux!AFaProfundidad
        '-----------------------------------------------------------------------
        
        If cAlto <> 0 Then sMedida = "(" & cAlto
        If cFrente <> 0 Then
            If sMedida <> "" Then
                sMedida = sMedida & "x" & cFrente
            Else
                sMedida = "(" & cAlto
            End If
        End If
        If cProf <> 0 Then
            If sMedida <> "" Then
                sMedida = sMedida & "x" & cProf
            Else
                sMedida = "(" & cProf
            End If
        End If
        If cAlto <> 0 And cFrente <> 0 And cProf <> 0 Then
            sMedida = sMedida & ")"
        Else
            If sMedida <> "" Then sMedida = sMedida & " cm)"
        End If
    End If
    RsAux.Close
    Exit Sub
errGA:
    clsGeneral.OcurrioError "Ocurrió un error al intentar obtener el argumento largo del artículo.", Err.Description, "Error (getargumento)"
End Sub

Private Sub AjustoTotalEtiqueta()
On Error Resume Next
    With vsEtiquetaArt
        .Subtotal flexSTClear
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 1, "#,###", &HC0FFFF, vbRed, True, "Total", , True
        .Subtotal flexSTSum, -1, 2, "#,###"
        .Subtotal flexSTSum, -1, 3, "#,###"
    End With
End Sub

Private Sub etiqueta_MandoAImprimir()
    
    If Not IsDate(tVigencia.Text) Then
        MsgBox "La fecha de vigencia no es correcta.", vbExclamation, "Datos Incorrectos"
        tVigencia.SetFocus: Exit Sub
    End If
    
    If etiqueta_BorroTablaAuxiliar Then
        If etiqueta_InsertoTablaAuxiliar Then
            Select Case cEtiquetaAImprimir.ListIndex
                Case 0: ImprimoEtiquetaConArgumento 'AccionImprimir TReporte.EtiquetaNormal
                Case 1: ImprimoEtiquetaConArgumento 'AccionImprimir TReporte.EtiquetaConArgumento
                Case 2: AccionImprimir TReporte.EtiquetaSinArgumento
            End Select
            etiqueta_BorroTablaAuxiliar
        End If
   End If

End Sub

Private Function etiqueta_BorroTablaAuxiliar() As Boolean
On Error GoTo errBTA
    etiqueta_BorroTablaAuxiliar = False
    Cons = "Delete EtiquetaAImprimir"
    cBase.Execute (Cons)
    etiqueta_BorroTablaAuxiliar = True
    Exit Function
errBTA:
    clsGeneral.OcurrioError "Ocurrió un error al vaciar la tabla auxiliar de impresión.", Err.Description, "Error (borrotablaauxiliar)"
End Function

Private Function etiqueta_InsertoTablaAuxiliar() As Boolean
On Error GoTo errITA
Dim lCont As Long
Dim sCtdo As String, sCuota As String, sPlan As String, sTotFin As String
    Dim sIDArt As String
    etiqueta_InsertoTablaAuxiliar = False
    Select Case cEtiquetaAImprimir.ListIndex
        Case 0  'Normales
            For lCont = 1 To vsEtiquetaArt.Rows - 1
                If Val(vsEtiquetaArt.Cell(flexcpText, lCont, 1)) > 0 And vsEtiquetaArt.IsSubtotal(lCont) = False Then
                    
                    'Armo los precios para este artículo.
                    If Val(vsEtiquetaArt.Cell(flexcpData, lCont, 1)) > 0 Then
                        etiqueta_CargoPreciosEspecificos Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin, Val(vsEtiquetaArt.Cell(flexcpData, lCont, 2))
                        sIDArt = Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 0)) & Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 1))
                    ElseIf Val(vsEtiquetaArt.Cell(flexcpData, lCont, 3)) = 1 Then
                        etiqueta_CargoPreciosCombo Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin
                        sIDArt = Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 0))
                    Else
                        etiqueta_CargoPrecios Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin
                        sIDArt = Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 0))
                    End If
                    'Inserto la cantidad de copias que pidio.
                    etiqueta_AddRowTablaAux sIDArt, sCtdo, sCuota, sTotFin, sPlan, Val(vsEtiquetaArt.Cell(flexcpText, lCont, 1))
                End If
            Next lCont
            
        Case 1 'c/ y s/argumento
            For lCont = 1 To vsEtiquetaArt.Rows - 1
                If Val(vsEtiquetaArt.Cell(flexcpText, lCont, 2)) > 0 And vsEtiquetaArt.IsSubtotal(lCont) = False Then
                    'Armo los precios para este artículo.
                    'etiqueta_CargoPrecios Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin
                    If Val(vsEtiquetaArt.Cell(flexcpData, lCont, 1)) > 0 Then
                        etiqueta_CargoPreciosEspecificos Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin, Val(vsEtiquetaArt.Cell(flexcpData, lCont, 2))
                        sIDArt = Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 0)) & Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 1))
                    ElseIf Val(vsEtiquetaArt.Cell(flexcpData, lCont, 3)) = 1 Then
                        etiqueta_CargoPreciosCombo Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin
                        sIDArt = Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 0))
                    Else
                        etiqueta_CargoPrecios Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin
                        sIDArt = Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 0))
                    End If
                    'Inserto la cantidad de copias que pidio.
                    'etiqueta_AddRowTablaAux Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, Trim(sTotFin), sPlan, Val(vsEtiquetaArt.Cell(flexcpText, lCont, 2))
                    etiqueta_AddRowTablaAux sIDArt, sCtdo, sCuota, sTotFin, sPlan, Val(vsEtiquetaArt.Cell(flexcpText, lCont, 2))
                End If
            Next lCont
        Case 2
            For lCont = 1 To vsEtiquetaArt.Rows - 1
                If Val(vsEtiquetaArt.Cell(flexcpText, lCont, 3)) > 0 And vsEtiquetaArt.IsSubtotal(lCont) = False Then
                    'Armo los precios para este artículo.
                    'etiqueta_CargoPrecios Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin
                    If Val(vsEtiquetaArt.Cell(flexcpData, lCont, 1)) > 0 Then
                        etiqueta_CargoPreciosEspecificos Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin, Val(vsEtiquetaArt.Cell(flexcpData, lCont, 2))
                        sIDArt = Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 0)) & Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 1))
                    ElseIf Val(vsEtiquetaArt.Cell(flexcpData, lCont, 3)) = 1 Then
                        etiqueta_CargoPreciosCombo Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin
                        sIDArt = Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 0))
                    Else
                        etiqueta_CargoPrecios Val(vsEtiquetaArt.Cell(flexcpData, lCont, 0)), sCtdo, sCuota, sPlan, sTotFin
                        sIDArt = Trim(vsEtiquetaArt.Cell(flexcpData, lCont, 0))
                    End If
                    'Inserto la cantidad de copias que pidio.
                    etiqueta_AddRowTablaAux sIDArt, sCtdo, sCuota, Trim(sTotFin), sPlan, Val(vsEtiquetaArt.Cell(flexcpText, lCont, 3))
                End If
            Next lCont
    End Select
    etiqueta_InsertoTablaAuxiliar = True
    Exit Function
    
errITA:
    clsGeneral.OcurrioError "Ocurrió un error al insertar los artículos en la tabla auxiliar.", Err.Description, "Error (insertotablaauxiliar)"
    etiqueta_BorroTablaAuxiliar
End Function

Public Function etiqueta_AddRowTablaAux(ByVal lArt As Long, ByVal sCtdo As String, ByVal sCuota As String, ByVal sFinanMedida As String, ByVal sPlan As String, ByVal iCant As Integer)
Dim rsAdd As rdoResultset
Dim iCont As Integer
    Cons = "Select * From EtiquetaAImprimir Where EImArticulo = " & lArt
    Set rsAdd = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    For iCont = 1 To iCant
        rsAdd.AddNew
        rsAdd!EImArticulo = lArt
        rsAdd!EImImporteCtdo = sCtdo
        rsAdd!EImCuota = sCuota
        rsAdd!EImTotalFinanciado = sFinanMedida
        rsAdd!EImPlan = sPlan
        rsAdd.Update
    Next
    rsAdd.Close
End Function

Private Sub etiqueta_CargoPreciosEspecificos(ByVal lArt As Long, ByRef sCtdo As String, ByRef sCuota As String, ByRef sPlan As String, ByRef sTotFin As String, ByVal Variacion As Currency)
Dim cCuotaAnt As Currency
Dim iCountEnter As Integer
    
Dim miPlan As Integer
Dim miContado As Currency

    sCtdo = ""
    sPlan = ""
    sCuota = ""
    sTotFin = ""
    
    'Saco el contado
    Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
        & " Where PVIArticulo = " & lArt _
        & " And PViMoneda = 1" _
        & " And PViTipoCuota = " & paTipoCuotaContado _
        & " And PViHabilitado = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then   'Si Hay contado busco el coeficiente p/tipo de Cuota y plan
        miPlan = RsAux!PViPlan
        miContado = RsAux!PViPrecio
    Else
        RsAux.Close
        Exit Sub
    End If
    RsAux.Close
    
    Dim miCoef As Currency
    
    miContado = miContado + Variacion
    miCoef = 0
    
    Select Case cEtiquetaAImprimir.ListIndex
        Case 0: sCtdo = "Cdo: $ " & Format(miContado, "#,###")
        Case 1, 2: sCtdo = "$ " & Format(miContado, "#,###")
    End Select
    
    'Ahora busco para los tipos de cuotas del plan.
    
    Dim rsTC As rdoResultset
    
    '1ero saco los tipos de cuotas a consultar.
    Cons = "Select TCuCodigo, TCuCantidad, PlaNombre " & _
            " From " & _
                        "HistoriaPrecio Precios, TipoCuota, TipoPlan" & _
                " WHERE (HPrVigencia IN" & _
                            " (Select MAX(H.HPrVigencia)" & _
                            " FROM HistoriaPrecio H" & _
                            " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                            " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                            " And H.HPrMoneda = Precios.HPrMoneda " & _
                            " And H.HPrVigencia <= '" & Format(tVigencia.Text, "mm/dd/yyyy 23:59:59") & "'" & _
                            " )) " & _
                " And Precios.HPrArticulo = " & lArt & _
                " And Precios.HPrMoneda = " & paMonedaPesos & _
                " And Precios.HPrHabilitado = 1 AND TCuCantidad > 1" & _
                " And Precios.HPrTipoCuota = TipoCuota.TCuCodigo" & _
                " And Precios.HPrPlan = TipoPlan.PlaCodigo" & _
                " And TipoCuota.TCuVencimientoE is null " & _
                " And TCuEspecial = 0 And TCuVencimientoC = 0 " & _
                " And TCuDeshabilitado is Null" & _
                " Order By TCuCantidad Asc"

    Set rsTC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsTC.EOF
        If Not IsNull(rsTC!PlaNombre) Then sPlan = "(" & Trim(rsTC!PlaNombre) & ")"
        miCoef = 0
        'Busco el coeficiente p/tipo de Cuota y plan
        Cons = "Select CoeCoeficiente from Coeficiente" & _
                    " Where CoePlan = " & miPlan & _
                    " And CoeTipoCuota = " & rsTC(0) & _
                    " And CoeMoneda = 1"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then miCoef = RsAux!CoeCoeficiente
        RsAux.Close
            
        Dim rCuotaF As Currency
        rCuotaF = 0
        
        If miCoef <> 0 Then
            rCuotaF = Redondeo(miContado * miCoef, 1)                                       'Precio de Unitario Financiaciado
            rCuotaF = (Redondeo(rCuotaF / rsTC("TCuCantidad"), 1)) * rsTC("TCuCantidad")    'Ajusto el Valor de la Cta
            rCuotaF = Redondeo(rCuotaF / rsTC("TCuCantidad"), 1)                                'Valor Cuota Finanaciado
        
            Select Case cEtiquetaAImprimir.ListIndex
                Case 0, 1
                    If sCuota <> "" Then sCuota = sCuota & vbCrLf
                    sCuota = sCuota & Trim(rsTC("TCuCantidad")) & " x $ " & Format(rCuotaF, "#,###")
                    
                    If sTotFin <> "" Then sTotFin = sTotFin & vbCrLf
                    sTotFin = sTotFin & "TF $ " & Format(Redondeo((rCuotaF * rsTC("TCuCantidad")), 1), "#,###")
                    
                Case 2
                    If (cCuotaAnt > rCuotaF _
                        Or cCuotaAnt = -1) And rCuotaF > paCuotaMin Then
                        
                        cCuotaAnt = rCuotaF
                        sCuota = "...o en " & Trim(rsTC!TCuCantidad) & " Cuotas de $ " & Format(rCuotaF, "#,###")
                        sTotFin = "Total financiado = $ " & Format(Redondeo(miContado * miCoef, 1), "#,###")
                    End If
            End Select
        
        End If
        rsTC.MoveNext
    Loop
    rsTC.Close
    
End Sub

Private Sub etiqueta_CargoPreciosCombo(ByVal lArt As Long, ByRef sCtdo As String, ByRef sCuota As String, ByRef sPlan As String, ByRef sTotFin As String)
On Error GoTo errCPC
Dim cCuotaAnt As Currency
    Dim oCuotas As Collection
    Set oCuotas = etiqueta_ProcesoPreciosCombo(lArt)
    Dim oCta As clsCuotas
    cCuotaAnt = -1
    For Each oCta In oCuotas
        If paTipoCuotaContado = oCta.ID Then
            Select Case cEtiquetaAImprimir.ListIndex
                Case 0: sCtdo = "Cdo: $ " & Format(oCta.TotalFinanciado, "#,###")
                Case 1, 2: sCtdo = "$ " & Format(oCta.TotalFinanciado, "#,###")
            End Select
        Else
            Select Case cEtiquetaAImprimir.ListIndex
                Case 0, 1
                    If sCuota <> "" Then sCuota = sCuota & vbCrLf
                    sCuota = sCuota & Trim(oCta.Cuotas) & " x $ " & Format(oCta.ImporteCuota, "#,###")
                    
                    If sTotFin <> "" Then sTotFin = sTotFin & vbCrLf
                    sTotFin = sTotFin & "TF $ " & Format(oCta.TotalFinanciado, "#,###")
                    
                Case 2
                    If (cCuotaAnt > oCta.ImporteCuota _
                        Or cCuotaAnt = -1) And oCta.ImporteCuota > paCuotaMin Then
                        cCuotaAnt = oCta.ImporteCuota
                        sCuota = "...o en " & Trim(oCta.Cuotas) & " Cuotas de $ " & Format(oCta.ImporteCuota, "#,###")
                        sTotFin = "Total financiado = $ " & Format(oCta.TotalFinanciado, "#,###")
                    End If
            End Select
        End If
    Next
    Exit Sub
errCPC:
    clsGeneral.OcurrioError "Error al cargar los precios del combo.", Err.Description, "Error en precios"
End Sub

Private Sub etiqueta_CargoPrecios(ByVal lArt As Long, ByRef sCtdo As String, ByRef sCuota As String, ByRef sPlan As String, ByRef sTotFin As String)
Dim cCuotaAnt As Currency
Dim iCountEnter As Integer
    sCtdo = ""
    sPlan = ""
    sCuota = ""
    sTotFin = ""
    
    Cons = "Select * " & _
               " From " & _
                        "HistoriaPrecio Precios, TipoCuota, TipoPlan" & _
                " WHERE (HPrVigencia IN" & _
                            " (Select MAX(H.HPrVigencia)" & _
                            " FROM HistoriaPrecio H" & _
                            " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                            " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                            " And H.HPrMoneda = Precios.HPrMoneda " & _
                            " And H.HPrVigencia <= '" & Format(tVigencia.Text, "mm/dd/yyyy 23:59:59") & "'" & _
                            " )) " & _
                " And Precios.HPrArticulo = " & lArt & _
                " And Precios.HPrMoneda = " & paMonedaPesos & _
                " And Precios.HPrHabilitado = 1" & _
                " And Precios.HPrTipoCuota = TipoCuota.TCuCodigo" & _
                " And Precios.HPrPlan = TipoPlan.PlaCodigo" & _
                " And TipoCuota.TCuVencimientoE is null " & _
                " And TCuEspecial = 0 And TCuVencimientoC = 0 " & _
                " And TCuDeshabilitado is Null" & _
                " Order By TCuCantidad Asc"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    cCuotaAnt = -1
    Do While Not RsAux.EOF
        If Not IsNull(RsAux!PlaNombre) Then sPlan = "(" & Trim(RsAux!PlaNombre) & ")"
        If paTipoCuotaContado = RsAux!TCuCodigo Then
            Select Case cEtiquetaAImprimir.ListIndex
                Case 0: sCtdo = "Cdo: $ " & Format(RsAux!HPrPrecio, "#,###")
                Case 1, 2: sCtdo = "$ " & Format(RsAux!HPrPrecio, "#,###")
            End Select
        Else
            Select Case cEtiquetaAImprimir.ListIndex
'                Case 0
'                    If sCuota <> "" Then sCuota = sCuota & vbCrLf
'                    sCuota = sCuota & Trim(RsAux!TCuCantidad) & " x $ " & Format(RsAux!HPrPrecio / RsAux!TCuCantidad, "#,###")
                
                Case 0, 1
                    If sCuota <> "" Then sCuota = sCuota & vbCrLf
                    sCuota = sCuota & Trim(RsAux!TCuCantidad) & " x $ " & Format(RsAux!HPrPrecio / RsAux!TCuCantidad, "#,###")
                    
                    If sTotFin <> "" Then sTotFin = sTotFin & vbCrLf
                    sTotFin = sTotFin & "TF $ " & Format(RsAux!HPrPrecio, "#,###")
                    
                Case 2
                    If (cCuotaAnt > RsAux!HPrPrecio / RsAux!TCuCantidad _
                        Or cCuotaAnt = -1) And RsAux!HPrPrecio / RsAux!TCuCantidad > paCuotaMin Then
                        
                        cCuotaAnt = RsAux!HPrPrecio / RsAux!TCuCantidad
                        sCuota = "...o en " & Trim(RsAux!TCuCantidad) & " Cuotas de $ " & Format(RsAux!HPrPrecio / RsAux!TCuCantidad, "#,###")
                        sTotFin = "Total financiado = $ " & Format(RsAux!HPrPrecio, "#,###")
                    End If
            End Select
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

Private Function etiqueta_ProcesoPreciosCombo(IDArticulo As Long) As Collection
On Error GoTo errPCombo

Dim arrCombo() As typCombo
Dim cbTotalF As Currency    'Total financiado del combo
Dim cbValorCuotaF As Currency   'Valor de c/cta del combo

Dim arTotalF As Currency    'Total financiado del articulo
Dim arValorCuotaF As Currency   'Valor de c/cta del articulo

Dim RsCom As rdoResultset
Dim iCount As Integer, miIDArticulo As Long, miQArticulo As Integer
    
    'Cargo las cuotas a mano, si hago la consulta me devuelve 12x pero sería correcto hacer la query pero por la urgencia lo hago así.
    Dim oCuotas As New Collection
    Dim oCta As New clsCuotas
    oCta.Cuotas = 0
    oCta.ID = 4
    oCuotas.Add oCta
    
    Set oCta = New clsCuotas
    oCta.Cuotas = 3
    oCta.ID = 6
    oCuotas.Add oCta
    
    Set oCta = New clsCuotas
    oCta.Cuotas = 5
    oCta.ID = 5
    oCuotas.Add oCta
    
    Set oCta = New clsCuotas
    oCta.Cuotas = 10
    oCta.ID = 3
    oCuotas.Add oCta
    
    Set oCta = New clsCuotas
    oCta.Cuotas = 15
    oCta.ID = 7
    oCuotas.Add oCta
    
    Dim aCostoB As Currency
    ReDim arrCombo(0)
    iCount = 1
    Cons = "Select * from Presupuesto, PresupuestoArticulo  " & _
               " Where PreArtCombo = " & IDArticulo & _
               " And PreID = PArPresupuesto And PreMoneda = 1"
    Set RsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsCom.EOF Then
        miIDArticulo = RsCom!PreArticulo
        aCostoB = RsCom!PreImporte
        
        Do While Not RsCom.EOF
            iCount = UBound(arrCombo) + 1
            ReDim Preserve arrCombo(iCount)
            arrCombo(iCount).Articulo = RsCom!PArArticulo
            arrCombo(iCount).Q = RsCom!PArCantidad
            arrCombo(iCount).EsBonificacion = False
            arrCombo(iCount).Bonificacion = 0
            RsCom.MoveNext
        Loop
        
        iCount = UBound(arrCombo) + 1
        ReDim Preserve arrCombo(iCount)
        arrCombo(iCount).Articulo = miIDArticulo
        arrCombo(iCount).Q = 1
        arrCombo(iCount).EsBonificacion = True
        arrCombo(iCount).Bonificacion = aCostoB
                
    End If
    RsCom.Close
    
    'PreArticulo es el articulo Bonificacion
    
    Dim bAvisoPrecio As Boolean
    
    
    Dim miUnitarioF As Currency, miCuotaF As Currency, bNoHabPlan As Boolean, miPlan As Long, bOK As Boolean
    Dim aCoefPD  As Currency
    'Para cada tipo de cuota .
    For Each oCta In oCuotas
        
        For I = 1 To UBound(arrCombo)
            arTotalF = 0: arValorCuotaF = 0
                    
            miIDArticulo = arrCombo(I).Articulo
            miQArticulo = arrCombo(I).Q
                
            If arrCombo(I).EsBonificacion Then   'Proceso el articulo bonificacion
                aCoefPD = 1
                arTotalF = arrCombo(I).Bonificacion
                If arTotalF <> 0 Then
                    'Busco el coeficiente p/tipo de Cuota y plan por defecto para finaciar la bonificacion
                    Cons = "Select * from Coeficiente" & _
                                " Where CoePlan = " & paPlanPorDefecto & _
                                " And CoeTipoCuota = " & oCta.ID & _
                                " And CoeMoneda = 1"
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If Not RsAux.EOF Then aCoefPD = RsAux!CoeCoeficiente
                    RsAux.Close
                End If
                
                'arTotalF = Format(arTotalF * aCoefPD, "0")
                arTotalF = Redondeo(arTotalF * aCoefPD, "1")
                arTotalF = arTotalF * miQArticulo
                If oCta.Cuotas > 0 Then
                    arTotalF = Redondeo(arTotalF / oCta.Cuotas, "1")
                    arTotalF = arTotalF * oCta.Cuotas
                    arValorCuotaF = Format(arTotalF / oCta.Cuotas, "#,##0.00")        'Valor Cuota Finanaciado
                End If
                oCta.ImporteCuota = oCta.ImporteCuota + arValorCuotaF
                oCta.TotalFinanciado = oCta.TotalFinanciado + arTotalF

            Else
                
                bOK = PrecioArticuloParaCombo(miIDArticulo, oCta.ID, oCta.Cuotas, miCuotaF, miUnitarioF)

                If bOK Then
                    arTotalF = miUnitarioF                          'Precio de Unitario Financiaciado
                    arTotalF = arTotalF * miQArticulo
                    arValorCuotaF = Format(miCuotaF * miQArticulo, "#,##0.00")     'Valor Cuota Finanaciado
                    oCta.ImporteCuota = oCta.ImporteCuota + arValorCuotaF
                    oCta.TotalFinanciado = oCta.TotalFinanciado + arTotalF
                End If
            End If
            
            If Not bOK And Not bAvisoPrecio Then bAvisoPrecio = True
        Next
    Next
    If bAvisoPrecio Then
        MsgBox "Importante en el combo existe(n) artículo(s) de los cuales no se logró determinar alguna financiación.", vbExclamation, "Falta Coeficiente"
    End If
    Set etiqueta_ProcesoPreciosCombo = oCuotas
    Screen.MousePointer = 0
    Exit Function
    
errPCombo:
    clsGeneral.OcurrioError "Error al procesar los artículos del combo.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function PrecioArticuloParaCombo(ByVal lArticulo As Long, ByVal lTCuota As Long, ByVal QCtas As Integer, rCuotaF As Currency, rUnitarioF As Currency) As Currency

    Dim miHayContado As Boolean
    PrecioArticuloParaCombo = True
    
    
    If QCtas = 0 Then QCtas = 1
    
    'Saco el valor de la cuota financiado
    Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
            & " Where PVIArticulo = " & lArticulo _
            & " And PViMoneda = 1 And PViTipoCuota = " & lTCuota _
            & " And PViHabilitado = 1"
    
    Dim bHayPrecio As Boolean
    bHayPrecio = False
    Dim rsArt As rdoResultset
    Set rsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    bHayPrecio = Not rsArt.EOF
    If bHayPrecio Then       'Hay Precios Grabados y están Habilitados
        rUnitarioF = Redondeo(rsArt!PViPrecio, "1")  'Format(RsArt!PViPrecio, "#,##0")                          'Precio de Unitario Financiaciado
        rCuotaF = Redondeo(rsArt!PViPrecio / QCtas, "1") 'Format(RsArt!PViPrecio / QCtas, "#,##0.00")     'Valor Cuota Finanaciado
    Else            'No Hay Precios Grabados O no Están Habilitados
        Dim miPlan As Long, miCoef As Currency, miContado As Currency
        '1) Busco SI Hay precio Contado
        miHayContado = False: miContado = 0
        Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
                & " Where PVIArticulo = " & lArticulo _
                & " And PViMoneda = 1" _
                & " And PViTipoCuota = " & paTipoCuotaContado _
                & " And PViHabilitado = 1"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then   'Si Hay contado busco el coeficiente p/tipo de Cuota y plan
            miHayContado = True
            miPlan = RsAux!PViPlan
            miContado = RsAux!PViPrecio
        End If
        RsAux.Close
        
        If miHayContado Then    'Si hay ctdo busco coeficiente
            miCoef = 0
            'Busco el coeficiente p/tipo de Cuota y plan
            Cons = "Select * from Coeficiente" & _
                        " Where CoePlan = " & miPlan & _
                        " And CoeTipoCuota = " & lTCuota & " And CoeMoneda = 1"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then miCoef = RsAux!CoeCoeficiente
            RsAux.Close
                                        
            If miCoef <> 0 Then
                rUnitarioF = Redondeo(miContado * miCoef, "1")                               'Precio de Unitario Financiaciado
                rUnitarioF = (Redondeo(rUnitarioF / QCtas, "1")) * QCtas                    'Ajusto el Valor de la Cta
                rCuotaF = Redondeo(rUnitarioF / QCtas, "1")                                'Valor Cuota Finanaciado
            End If
        End If
    End If
    rsArt.Close
    
End Function


Private Sub rptImprimoEtiquetas(ByVal sTitulo As String)
On Error GoTo ErrCrystal
    Screen.MousePointer = 11
    
    If crMandoAPantalla(jobnum, sTitulo) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    clsGeneral.OcurrioError crMsgErr
    Screen.MousePointer = 0
End Sub

Private Function ArticuloTienePrecio(ByVal lArt As Long) As Boolean
    
    Cons = "Select * " & _
            " From " & _
                    "HistoriaPrecio Precios, TipoCuota, TipoPlan" & _
            " WHERE (HPrVigencia IN" & _
                        " (Select MAX(H.HPrVigencia)" & _
                        " FROM HistoriaPrecio H" & _
                        " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                        " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                        " And H.HPrMoneda = Precios.HPrMoneda " & _
                        " And H.HPrVigencia <= '" & Format(tVigencia.Text, "mm/dd/yyyy 23:59:59") & "'" & _
                        " )) " & _
            " And Precios.HPrArticulo = " & lArt & _
            " And Precios.HPrMoneda = " & paMonedaPesos & _
            " And Precios.HPrHabilitado = 1" & _
            " And Precios.HPrTipoCuota = TipoCuota.TCuCodigo" & _
            " And Precios.HPrPlan = TipoPlan.PlaCodigo" & _
            " And TipoCuota.TCuVencimientoE Is Null " & _
            " And TCuEspecial = 0 And TCuVencimientoC = 0 " & _
            " And TCuDeshabilitado Is Null"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        ArticuloTienePrecio = True
    Else
        ArticuloTienePrecio = False
    End If
    RsAux.Close
    
End Function

Private Function EspecificoIngresado(ByVal IDEsp As Long) As Boolean
    Dim lEsta As Integer
    For lEsta = 1 To vsEtiquetaArt.Rows - 1
        If Val(vsEtiquetaArt.Cell(flexcpData, lEsta, 1)) = IDEsp Then
            MsgBox "El artículo ya esta ingresado, edite la columna de cantidades si desea modificarlas.", vbInformation, "ATENCIÓN"
            vsEtiquetaArt.Select lEsta, 0, lEsta, vsEtiquetaArt.Cols - 1
            vsEtiquetaArt.SetFocus
            tArticulo.Text = ""
            tArticulo.Tag = ""
            EspecificoIngresado = True
            Exit Function
        End If
    Next
    
End Function

Private Sub loc_InsertArticuloEspecifico()
On Error GoTo errTC
    
    Screen.MousePointer = 11
    Dim iRet As Long, iVarP As Currency, bEnv As Boolean
    Dim objLista As New clsListadeAyuda
    iRet = objLista.ActivarAyuda(cBase, "EXEC prg_BuscarArticuloEspecifico '" + tArticulo.Text + "'", 5200, 3, "Ayuda")
    
    Me.Refresh
    If iRet = 0 Then
    Screen.MousePointer = 0
        Set objLista = Nothing
        Exit Sub
    Else
        
        If EspecificoIngresado(objLista.RetornoDatoSeleccionado(3)) Then
            Set objLista = Nothing
            MsgBox "El artículo seleccionado ya está ingresado.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        tArticulo.Text = "E:" & objLista.RetornoDatoSeleccionado(3) & " " & Trim(objLista.RetornoDatoSeleccionado(4))
        tArticulo.Tag = Val(objLista.RetornoDatoSeleccionado(0))
        
        artInfo.IDEsp = objLista.RetornoDatoSeleccionado(3)
        artInfo.VariacionEsp = objLista.RetornoDatoSeleccionado(5)
        
        tArticulo.SetFocus
    End If
    Set objLista = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errTC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar artículos específicos.", Err.Description, "Artículo específico"
End Sub


