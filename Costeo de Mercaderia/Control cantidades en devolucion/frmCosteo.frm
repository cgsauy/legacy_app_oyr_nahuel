VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Begin VB.Form frmCosteo 
   Caption         =   "Costeo de Mercaderia"
   ClientHeight    =   5580
   ClientLeft      =   2835
   ClientTop       =   2430
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCosteo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7230
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   180
      ScaleHeight     =   315
      ScaleWidth      =   6015
      TabIndex        =   32
      Top             =   7140
      Width           =   6075
      Begin VB.CommandButton bSucesosC 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   60
         Picture         =   "frmCosteo.frx":0442
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Sucesos."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmCosteo.frx":0744
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmCosteo.frx":097E
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmCosteo.frx":0C80
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmCosteo.frx":0FC2
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmCosteo.frx":10C4
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmCosteo.frx":12FE
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmCosteo.frx":13E8
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmCosteo.frx":14D2
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   0
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmCosteo.frx":194C
         Height          =   310
         Left            =   4440
         Picture         =   "frmCosteo.frx":1A4E
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   0
         Width           =   310
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información del Último Costeo"
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6255
      Begin VB.Label lCUsuario 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   285
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label lCFecha 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label lCMes 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diciembre 2000"
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Mes Costeado:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   285
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   3915
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   6255
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   4740
         ScaleHeight     =   405
         ScaleWidth      =   1395
         TabIndex        =   3
         Top             =   180
         Width           =   1395
         Begin VB.CommandButton bSucesosG 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   480
            Picture         =   "frmCosteo.frx":1F80
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "Sucesos."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bCancelar 
            Height          =   310
            Left            =   1020
            Picture         =   "frmCosteo.frx":2282
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Salir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bConsultar 
            Height          =   310
            Left            =   60
            Picture         =   "frmCosteo.frx":2384
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Ejecutar."
            Top             =   50
            Width           =   310
         End
      End
      Begin VB.TextBox tMes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Text            =   "8/1999"
         Top             =   240
         Width           =   1575
      End
      Begin ComctlLib.ProgressBar bProgress 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Generación de Ventas"
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Generación de Compras"
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin :"
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lVentaF 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   26
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio:"
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lVentaI 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/12/1999 00:00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   24
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lCompraF 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lCompraI 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/12/1999 00:00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lAccion 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
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
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label lCosteoI 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/12/1999 00:00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio del Costeo:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lCosteoF 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin del Costeo:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "&Mes A Costear:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   5325
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sucursal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4551
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   6060
      TabIndex        =   30
      Top             =   1920
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7646
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
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
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
      OutlineBar      =   1
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
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   3135
      Left            =   180
      TabIndex        =   31
      Top             =   4020
      Width           =   6015
      _Version        =   196608
      _ExtentX        =   10610
      _ExtentY        =   5530
      _StockProps     =   229
      BorderStyle     =   1
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
      PreviewMode     =   1
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
End
Attribute VB_Name = "frmCosteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum TipoCV
    Compra = 1
    Comercio = 2
    Importacion = 3
End Enum

Dim RsAux As rdoResultset
Dim bCargarImpresion  As Boolean

Private Sub AccionCostear()

Dim aIDCosteo As Long
    
    If Not ValidoCampos Then Exit Sub
    If MsgBox("Confirma realizar el costeo de mercadería para el mes de " & tMes.Text, vbQuestion + vbYesNo, "Costear Mercadería") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    On Error GoTo errCostear
    vsConsulta.Tag = tMes.Text: vsConsulta.Rows = 1
    FechaDelServidor
    
    '-------------------------------------------------------------------------------------------------------------------------------
    Cons = "Select * from CMCabezal"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!CabMesCosteo = Format(tMes.Text, sqlFormatoF)
    RsAux!CabFecha = Format(gFechaServidor, sqlFormatoFH)
    RsAux!CabUsuario = paCodigoDeUsuario
    RsAux.Update: RsAux.Close
    
    Cons = "Select Max(CabID) from CMCabezal"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aIDCosteo = RsAux(0)
    RsAux.Close
    '-------------------------------------------------------------------------------------------------------------------------------
    
    lAccion.Caption = "Procesando compras del mes...": lAccion.Refresh
    lCompraI.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCompraI.Refresh
    CargoTablaCMCompra CDate(tMes.Text)
    lCompraF.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCompraF.Refresh
    
    lAccion.Caption = "Procesando ventas del mes...": lAccion.Refresh
    lVentaI.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lVentaI.Refresh
    CargoTablaCMVenta CDate(tMes.Text)
    lVentaF.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lVentaF.Refresh
    
    lAccion.Caption = "Costeando Mercadería...": lAccion.Refresh
    lCosteoI.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCosteoI.Refresh
    CargoTablaCMCosteo aIDCosteo
    lCosteoF.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCosteoF.Refresh
    lAccion.Caption = "Costeo Finalizado OK.": lAccion.Refresh
    
    MsgBox "El costeo para el mes de " & Trim(tMes.Text) & " se ha finalizado con éxito.", vbExclamation, "Costeo Finalizado"
    CargoInformacionCosteo
    
    Screen.MousePointer = 0
    Exit Sub

errCostear:
    clsGeneral.OcurrioError "Ocurrió un error al costear la mercadería.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoTablaCMCompra(Mes As Date)

    '1) Cargar las compras del mes (Credito y Contado)
    '2) Cargar las importes del mes (con fecha de arribo del costeo en el mes)
    
    'ATENCIÓN: solo van los articulos que tengan el campo ArtAMercaderia = True !!!!!!!
        
Dim aCosto As Currency
Dim QyCom As rdoQuery
    
    Cons = "Insert Into CMCompra (ComFecha, ComArticulo, ComTipo, ComCodigo, ComCantidad, ComCosto, ComQOriginal) Values (?,?,?,?,?,?,?)"
    Set QyCom = cBase.CreateQuery("", Cons)
    
    '1) Compras del Mes (contado y credito)------------------------------------------------------------------------------------------------------------------------

    Cons = "Select * from Compra, CompraRenglon, Articulo" _
          & " Where ComCodigo = CReCompra" _
          & " And ComFecha Between '" & Format(Mes, sqlFormatoFH) & "' And '" & Format(UltimoDia(Mes) & " 23:59:59", sqlFormatoFH) & "'" _
          & " And ComTipoDocumento In (" & TipoDocumento.Compracontado & ", " & TipoDocumento.CompraCredito & ")" _
          & " And CReArticulo = ArtID" _
          & " And ArtAMercaderia = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        QyCom.rdoParameters(0) = RsAux!ComFecha
        QyCom.rdoParameters(1) = RsAux!CReArticulo
        QyCom.rdoParameters(2) = TipoCV.Compra
        QyCom.rdoParameters(3) = RsAux!ComCodigo
        
        QyCom.rdoParameters(4) = RsAux!CReCantidad
        QyCom.rdoParameters(6) = RsAux!CReCantidad
        
        If RsAux!ComMoneda <> paMonedaPesos Then
            aCosto = RsAux!CRePrecioU * RsAux!ComTC
        Else
            aCosto = RsAux!CRePrecioU
        End If
        QyCom.rdoParameters(5) = aCosto
        
        QyCom.Execute
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '2) Importaciones Costeadas en el Mes-------------------------------------------------------------------------------------------------------------------------
    Cons = "Select * from CosteoCarpeta, CosteoArticulo, Articulo" _
           & " Where CCaFArribo Between '" & Format(Mes, sqlFormatoFH) & "' And '" & Format(UltimoDia(Mes) & " 23:59:59", sqlFormatoFH) & "'" _
           & " And CCaID = CArIDCosteo" _
           & " And CArIdArticulo = ArtId " _
           & " And ArtAMercaderia = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        QyCom.rdoParameters(0) = RsAux!CCaFArribo
        QyCom.rdoParameters(1) = RsAux!CArIdArticulo
        QyCom.rdoParameters(2) = TipoCV.Importacion
        QyCom.rdoParameters(3) = RsAux!CCaID
        
        QyCom.rdoParameters(4) = RsAux!CArCantidad
        QyCom.rdoParameters(6) = RsAux!CArCantidad
        QyCom.rdoParameters(5) = RsAux!CArCostoP
        QyCom.Execute
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    QyCom.Close
    
End Sub

Private Sub CargoTablaCMCosteo(IDCosteo As Long)

Dim QyCos As rdoQuery
Dim RsVen As rdoResultset, RsCom As rdoResultset

Dim aFVenta As Date, aArticulo As Long
Dim aQVenta As Long, aQCompra As Long, aQCosteo As Long
Dim aQVentaOriginal As Long
Dim bSalir As Boolean, bBorroVenta As Boolean

    'Preparo las Queries para costear-----------------------------------------------------------------------------------------------------------------------------
    Cons = "Insert Into CMCosteo (CosID, CosArticulo, CosTipoVenta, CosIDVenta, CosCantidad, CosCosto, CosVenta) Values (?,?,?,?,?,?,?)"
    Set QyCos = cBase.CreateQuery("", Cons)
    QyCos.rdoParameters(0) = IDCosteo

    Dim QyCMenores As rdoQuery, QyCMayores As rdoQuery
    
    Cons = "Select * from CMCompra " _
            & " Where ComFecha <= ?" _
            & " And ComArticulo = ? " _
            & " And ComCantidad > 0 " _
            & " Order by ComFecha DESC"
    Set QyCMenores = cBase.CreateQuery("", Cons)
    
    Cons = "Select * from CMCompra " _
            & " Where ComFecha >= ?" _
            & " And ComArticulo = ? " _
            & " And ComCantidad > 0 "
    Set QyCMayores = cBase.CreateQuery("", Cons)
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '-------------------------------------------------------------------------------------------------------------------
    Cons = "Select Count(*) from CMVenta"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) <> 0 Then bProgress.Max = RsAux(0)
    RsAux.Close
    bProgress.Value = 0
    '-------------------------------------------------------------------------------------------------------------------

    Cons = "Select * from CMVenta, Articulo" _
           & " Where VenArticulo = ArtID " _
           & " Order by VenFecha, VenArticulo, VenTipo, VenCodigo"
    Set RsVen = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
       
    Do While Not RsVen.EOF
        bProgress.Value = bProgress.Value + 1
        aFVenta = RsVen!VenFecha
        aArticulo = RsVen!VenArticulo
        aQVenta = RsVen!VenCantidad
        aQVentaOriginal = aQVenta
        
        QyCos.rdoParameters(1) = aArticulo
        bBorroVenta = True
        Do While aQVenta <> 0
        
            'Si el artículo es del tipo Servicio lo costeo contra costo 0
            If RsVen!ArtTipo = paTipoArticuloServicio Then
                aQCosteo = aQVenta
                aQVenta = 0
                
                QyCos.rdoParameters(2) = RsVen!VenTipo
                QyCos.rdoParameters(3) = RsVen!VenCodigo
                QyCos.rdoParameters(4) = aQCosteo
                QyCos.rdoParameters(5) = 0
                QyCos.rdoParameters(6) = RsVen!VenPrecio
                QyCos.Execute
            
            Else
        
                'Voy a la maxima fecha de Compra <= a la fecha de venta ------------------------------------
                QyCMenores.rdoParameters(0) = Format(aFVenta, sqlFormatoF) & " 23:59:59"
                QyCMenores.rdoParameters(1) = aArticulo
                Set RsCom = QyCMenores.OpenResultset(rdOpenDynamic, rdConcurValues)
                
                If Not RsCom.EOF Then               'Hay una FC <= FV
                    If aQVenta > 0 Then                 'VENTA DE MERCADERIA---------------------------------------------------
                        aQCompra = RsCom!ComCantidad
                        If aQVenta > aQCompra Then
                            aQVenta = aQVenta - aQCompra
                            aQCosteo = aQCompra
                        Else
                            aQCosteo = aQVenta
                            aQVenta = 0
                        End If
                        
                        QyCos.rdoParameters(2) = RsVen!VenTipo
                        QyCos.rdoParameters(3) = RsVen!VenCodigo
                        QyCos.rdoParameters(4) = aQCosteo
                        QyCos.rdoParameters(5) = RsCom!ComCosto
                        QyCos.rdoParameters(6) = RsVen!VenPrecio
                        QyCos.Execute
                        
                        RsCom.Edit
                        RsCom!ComCantidad = RsCom!ComCantidad - aQCosteo
                        RsCom.Update
                    
                    Else        'DEVOLUCION DE MERCADERIA---------------------------------------------------
                                  'La cantidad debe ser siempre menor a la original, sino voy al inmediato anterior (x q voy a sumar 1 sino me paso)
                         bSalir = False
                        'Como viene DESC hago move next hasta encontrar uno
                        Do While Not bSalir
                            If RsCom!ComCantidad < RsCom!ComQOriginal Then
                                bSalir = True
                            
                                aQCompra = RsCom!ComQOriginal - RsCom!ComCantidad
                                If Abs(aQVenta) > aQCompra Then
                                    aQVenta = (Abs(aQVenta) - aQCompra) * -1
                                    aQCosteo = -aQCompra * -1
                                Else
                                    aQCosteo = aQVenta
                                    aQVenta = 0
                                End If
                                
                                QyCos.rdoParameters(2) = RsVen!VenTipo
                                QyCos.rdoParameters(3) = RsVen!VenCodigo
                                QyCos.rdoParameters(4) = aQCosteo
                                QyCos.rdoParameters(5) = RsCom!ComCosto
                                QyCos.rdoParameters(6) = RsVen!VenPrecio
                                QyCos.Execute
                            
                                RsCom.Edit
                                RsCom!ComCantidad = RsCom!ComCantidad - aQCosteo
                                RsCom.Update
                                
                            Else
                                RsCom.MoveNext
                                If RsCom.EOF Then
                                    'Como no hay una compra (que ya tenga algo costeado) deberia costear la devolucion contra costo 0
                                    bSalir = True: bBorroVenta = False: aQVenta = 0
                                End If
                                
                            End If
                        Loop
                        
                    End If
                    RsCom.Close
    
                Else                                        'NO Hay una FC <= FV
                    RsCom.Close
                    'Voy a la minima fecha de Compra >= a la fecha de venta------------------------------------
                    QyCMayores.rdoParameters(0) = Format(aFVenta, sqlFormatoF) & " 23:59:59"
                    QyCMayores.rdoParameters(1) = aArticulo
                    Set RsCom = QyCMayores.OpenResultset(rdOpenDynamic, rdConcurValues)
                    
                    If Not RsCom.EOF Then               'Hay una FC >= FV
                    
                        If aQVenta > 0 Then                 'VENTA DE MERCADERIA---------------------------------------------------
                            aQCompra = RsCom!ComCantidad
                            If aQVenta > aQCompra Then
                                aQVenta = aQVenta - aQCompra
                                aQCosteo = aQCompra
                            Else
                                aQCosteo = aQVenta
                                aQVenta = 0
                            End If
                            
                            QyCos.rdoParameters(2) = RsVen!VenTipo
                            QyCos.rdoParameters(3) = RsVen!VenCodigo
                            QyCos.rdoParameters(4) = aQCosteo
                            QyCos.rdoParameters(5) = RsCom!ComCosto
                            QyCos.rdoParameters(6) = RsVen!VenPrecio
                            QyCos.Execute
                            
                            RsCom.Edit
                            RsCom!ComCantidad = RsCom!ComCantidad - aQCosteo
                            RsCom.Update
                        
                        Else        'DEVOLUCION DE MERCADERIA---------------------------------------------------
                                  'La cantidad debe ser siempre menor a la original, sino voy al inmediato siguiente
                            bSalir = False
                            Do While Not bSalir
                                If RsCom!ComCantidad < RsCom!ComQOriginal Then
                                    bSalir = True
                                
                                    aQCompra = RsCom!ComQOriginal - RsCom!ComCantidad
                                    'aQCompra = RsCom!ComCantidad                'OJO !!!!!!!!!!!
                                    If Abs(aQVenta) > aQCompra Then
                                        aQVenta = (Abs(aQVenta) - aQCompra) * -1
                                        aQCosteo = -aQCompra * -1
                                    Else
                                        aQCosteo = aQVenta
                                        aQVenta = 0
                                    End If
                                    
                                    QyCos.rdoParameters(2) = RsVen!VenTipo
                                    QyCos.rdoParameters(3) = RsVen!VenCodigo
                                    QyCos.rdoParameters(4) = aQCosteo
                                    QyCos.rdoParameters(5) = RsCom!ComCosto
                                    QyCos.rdoParameters(6) = RsVen!VenPrecio
                                    QyCos.Execute
                                
                                    RsCom.Edit
                                    RsCom!ComCantidad = RsCom!ComCantidad - aQCosteo
                                    RsCom.Update
                                    
                                Else
                                    RsCom.MoveNext
                                    If RsCom.EOF Then bSalir = True: bBorroVenta = False: aQVenta = 0
                                End If
                            Loop
                            RsCom.Close
                        End If
                        
                    Else
                        RsCom.Close
                        'Si no hay datos queda remanente, Primero updateo con lo que queda remanente en la venta
                        '11 de Mayo de 2000 - 1) Si es una devolucion y queda remanete la costeo contra costo 0 (aQVenta < 0)
                                                       ' 2) Registro un suceso en la grilla y borro la Venta para que no quede remanete (aQVenta = 0 And bBorroVenta)
                        If aQVenta < 0 Then
                            'Costeo contra costo 0
                            QyCos.rdoParameters(2) = RsVen!VenTipo
                            QyCos.rdoParameters(3) = RsVen!VenCodigo
                            QyCos.rdoParameters(4) = aQVenta
                            QyCos.rdoParameters(5) = 0
                            QyCos.rdoParameters(6) = RsVen!VenPrecio
                            QyCos.Execute
                            
                            'Registro suceso y seteo señales para borrar la venta
                            AgegoSuceso "Rebote", RsVen!VenFecha, RsVen!VenArticulo, RsVen!VenTipo, RsVen!VenCodigo, RsVen!VenPrecio, aQVenta
                            
                            aQVenta = 0: bBorroVenta = True
                        Else
                        
                            If aQVenta <> aQVentaOriginal Then
                                Cons = "Update CMVenta Set VenCantidad = " & aQVenta _
                                        & " Where VenFecha = '" & Format(RsVen!VenFecha, sqlFormatoFH) & "'" _
                                        & " And VenArticulo = " & RsVen!VenArticulo _
                                        & " And VenTipo = " & RsVen!VenTipo & " And VenCodigo = " & RsVen!VenCodigo
                                cBase.Execute Cons
                            End If
                        End If
                        Exit Do
                        
                    End If
                End If
            End If
        Loop
        
        'Si la venta quedó en cero elimino el registro de la venta
        If aQVenta = 0 And bBorroVenta Then
            Cons = " Delete CMVenta " _
                    & " Where VenFecha = '" & Format(RsVen!VenFecha, sqlFormatoFH) & "'" _
                    & " And VenArticulo = " & RsVen!VenArticulo _
                    & " And VenTipo = " & RsVen!VenTipo & " And VenCodigo = " & RsVen!VenCodigo
            cBase.Execute Cons
        End If
        RsVen.MoveNext
    Loop
    
    RsVen.Close
    QyCMenores.Close: QyCMayores.Close
    
    'Hay que borrar las compras en que las cantidades son iguales a 0
    Cons = "Delete CMCompra Where ComCantidad = 0"
    cBase.Execute Cons
    '---------------------------------------------------------------------------
    bProgress.Value = 0
    
End Sub

Private Sub AgegoSuceso(Texto As String, fecha As Date, Articulo As Long, Tipo As Integer, Id As Long, Precio As Currency, Cantidad As Long)

Dim rsS As rdoResultset

    'On Error Resume Next
    With vsConsulta
        .AddItem ""     '"<Descripción|<Fecha|Tipo|Documento|Artículo|>Q|>Costo (x1)|>Total"
        .Cell(flexcpText, .Rows - 1, 0) = Trim(Texto)
        .Cell(flexcpText, .Rows - 1, 1) = Format(fecha, "dd/mm/yyyy")
        
        Select Case Tipo
            Case TipoCV.Comercio:
                Cons = "Select * from Documento Where DocCodigo = " & Id
                Set rsS = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsS.EOF Then
                    .Cell(flexcpText, .Rows - 1, 2) = "Comercio"
                    .Cell(flexcpText, .Rows - 1, 3) = RetornoNombreDocumento(rsS!DocTipo, Abreviacion:=True) & " "
                    .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & Trim(rsS!DocSerie) & Format(rsS!DocNumero, "000000")
                End If
                rsS.Close
        
            Case TipoCV.Compra:
                Cons = "Select * from Compra Where ComCodigo = " & Id
                Set rsS = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsS.EOF Then
                    .Cell(flexcpText, .Rows - 1, 2) = "Compras"
                    .Cell(flexcpText, .Rows - 1, 3) = RetornoNombreDocumento(rsS!ComTipoDocumento, Abreviacion:=True) & " "
                    If Not IsNull(rsS!ComSerie) Then .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & Trim(rsS!ComSerie)
                    If Not IsNull(rsS!ComNumero) Then .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & Trim(rsS!ComNumero) & " "
                End If
                rsS.Close
        End Select
        
        Cons = "Select * from Articulo Where ArtId = " & Articulo
        Set rsS = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsS.EOF Then .Cell(flexcpText, .Rows - 1, 4) = Format(rsS!ArtCodigo, "(#,000,000)") & " " & Trim(rsS!ArtNombre)
        rsS.Close
        
        .Cell(flexcpText, .Rows - 1, 5) = Cantidad
        .Cell(flexcpText, .Rows - 1, 6) = Format(Precio, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 7) = Format(Precio * Cantidad, FormatoMonedaP)
        
    End With
    
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    LimpioDatos
    AccionCostear
End Sub

Private Sub bImprimir_Click()
    AccionImprimir True
End Sub

Private Sub bPrimero_Click()
    IrAPagina vsListado, 1
End Sub

Private Sub bSiguiente_Click()
    IrAPagina vsListado, vsListado.PreviewPage + 1
End Sub

Private Sub bUltima_Click()
    IrAPagina vsListado, vsListado.PageCount
End Sub

Private Sub bZMas_Click()
    Zoom vsListado, vsListado.Zoom + 5
End Sub

Private Sub bZMenos_Click()
    Zoom vsListado, vsListado.Zoom - 5
End Sub

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub bSucesosC_Click()
    Frame1.ZOrder 0
    Frame1.Visible = True
End Sub

Private Sub bSucesosG_Click()
    vsConsulta.ZOrder 0
    Picture1.ZOrder 0
    Frame1.Visible = False
End Sub

Private Sub chVista_Click()
    
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
    Else
        AccionImprimir
        vsListado.ZOrder 0
    End If

End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error Resume Next
    Frame1.ZOrder 0
    Frame1.Visible = True
    Picture1.BorderStyle = 0
    LimpioDatos
    
    CargoInformacionCosteo
    
    Picture1.Top = Frame1.Top + Frame1.Height - Picture1.Height - 50
    vsConsulta.Top = Frame1.Top
    vsConsulta.Left = Frame1.Left
    vsConsulta.Width = Frame1.Width
    vsConsulta.Height = Frame1.Height - Picture1.Height - 100
    
    vsListado.Top = vsConsulta.Top:  vsListado.Width = vsConsulta.Width
    vsListado.Left = vsConsulta.Left: vsListado.Height = vsConsulta.Height
    
    InicializoGrillas
    bCargarImpresion = True
    With vsListado
        .PaperSize = 1
        .Orientation = orPortrait
        .Zoom = 100
        .MarginLeft = 900: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
    End With

        
End Sub

Private Sub LimpioDatos()
    lAccion.Caption = ""
    lCosteoI.Caption = "": lCosteoF.Caption = ""
    lCompraI.Caption = "": lCompraF.Caption = ""
    lVentaI.Caption = "": lVentaF.Caption = ""
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    Frame1.Height = Me.ScaleHeight - Frame1.Top - Status.Height - 100
    Frame1.Width = Me.ScaleWidth - (Frame1.Left * 2)
    Frame2.Width = Frame1.Width
    
    lAccion.Width = Frame1.Width - (lAccion.Left * 2)
    bProgress.Width = lAccion.Width
    
    Picture1.Top = Frame1.Top + Frame1.Height - Picture1.Height - 50
    vsConsulta.Top = Frame1.Top
    vsConsulta.Left = Frame1.Left
    vsConsulta.Width = Frame1.Width
    vsConsulta.Height = Frame1.Height - Picture1.Height - 100
    
    vsListado.Top = vsConsulta.Top:  vsListado.Width = vsConsulta.Width
    vsListado.Left = vsConsulta.Left: vsListado.Height = vsConsulta.Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub CargoTablaCMVenta(Mes As Date)
'Parámetro: Recibe el primer día del mes a costear.
On Error GoTo ErrCTCMV
Dim QyVta As rdoQuery
Dim strDocumentos As String
Dim Monedas As String
Dim aCosto As Currency
Dim aMoneda As Long, aTC As Currency

    Screen.MousePointer = 11
    
    Monedas = RetornoTCMonedas(PrimerDia(Mes))
    aMoneda = 0
    'Creo query para insertar los datos
    Cons = "Insert Into CMVenta (VenFecha, VenArticulo, VenTipo, VenCodigo, VenCantidad,VenPrecio) Values (?,?,?,?,?,?)"
    Set QyVta = cBase.CreateQuery("", Cons)
    
    'Primer Paso Copio las Ventas---------------------------------------
    'Traigo los documentos Ctdo y Cred, Nota Esp, Nota de Cred. y  Nota de Dev. que no estén anulados
    strDocumentos = TipoDocumento.Contado & ", " & TipoDocumento.Credito _
        & ", " & TipoDocumento.NotaCredito & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial
    
    Cons = "Select DocFecha, DocMoneda, DocTipo, Renglon.* From Documento, Renglon" _
        & " Where DocTipo IN (" & strDocumentos & ")" _
        & " And DocFecha BetWeen '" & Format(PrimerDia(Mes) & " 00:00:00", sqlFormatoFH) & "'" _
        & " And '" & Format(UltimoDia(Mes) & " 23:59:59", sqlFormatoFH) & "'" _
        & " And DocAnulado = 0 And DocCodigo = RenDocumento"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)

    Do While Not RsAux.EOF
        QyVta.rdoParameters(0) = Format(RsAux!DocFecha, sqlFormatoF)
        QyVta.rdoParameters(1) = RsAux!RenArticulo
        QyVta.rdoParameters(2) = TipoCV.Comercio
        QyVta.rdoParameters(3) = RsAux!RenDocumento
        
        Select Case RsAux!DocTipo
            Case TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial
                            QyVta.rdoParameters(4) = RsAux!RenCantidad * -1
                            
            Case Else: QyVta.rdoParameters(4) = RsAux!RenCantidad
        End Select
        
        'Es el precio neto-----------------------------------------------------------------------------
        If aMoneda <> RsAux!DocMoneda Then
            aMoneda = RsAux!DocMoneda
            aTC = ValorTC(RsAux!DocMoneda, Monedas)
        End If
        
        QyVta.rdoParameters(5) = (RsAux!RenPrecio - RsAux!RenIva) * aTC
        '-------------------------------------------------------------------------------------------------
        QyVta.Execute
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Segundo paso Cargo Notas de Compras.
    Cons = "Select * from Compra, CompraRenglon" _
          & " Where ComCodigo = CReCompra" _
          & " And ComFecha Between '" & Format(Mes, sqlFormatoFH) & "' And '" & Format(UltimoDia(Mes) & " 23:59:59", sqlFormatoFH) & "'" _
          & " And ComTipoDocumento In (" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    Do While Not RsAux.EOF
        QyVta.rdoParameters(0) = Format(RsAux!ComFecha, sqlFormatoF)
        QyVta.rdoParameters(1) = RsAux!CReArticulo
        QyVta.rdoParameters(2) = TipoCV.Compra
        QyVta.rdoParameters(3) = RsAux!CReCompra
        QyVta.rdoParameters(4) = RsAux!CReCantidad
        If RsAux!ComMoneda <> paMonedaPesos Then
            aCosto = RsAux!CRePrecioU * RsAux!ComTC
        Else
            aCosto = RsAux!CRePrecioU
        End If
        QyVta.rdoParameters(5) = aCosto
        QyVta.Execute
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    QyVta.Close 'Cierro query
    Screen.MousePointer = 0
    Exit Sub
    
ErrCTCMV:
    clsGeneral.OcurrioError "Ocurrio un error al cargar la tabla de Ventas.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValorTC(PosMoneda As Integer, ByVal strMonedas As String) As Currency
Dim Cont As Integer
    
    Cont = 1: ValorTC = 1
    Do While strMonedas <> ""
        If PosMoneda = Cont Then
            ValorTC = Mid(strMonedas, 1, InStr(1, strMonedas, ":") - 1)
            Exit Function
        Else
            strMonedas = Mid(strMonedas, InStr(1, strMonedas, ":") + 1, Len(strMonedas))
            Cont = Cont + 1
        End If
    Loop
    
End Function

Private Function RetornoTCMonedas(fecha As Date) As String

Dim aTC As Currency, Contador As Integer

    'Armo vector con las TC de las monedas que existen.
    RetornoTCMonedas = ""
    Cons = "Select * From Moneda"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Contador = 1
    Do While Not RsAux.EOF
        If Contador = RsAux!MonCodigo Then
            If RsAux!MonCodigo = paMonedaPesos Then
                aTC = 1
            Else
                aTC = TasadeCambio(RsAux!MonCodigo, paMonedaPesos, fecha)
            End If
        Else
            aTC = 1
        End If
        If RetornoTCMonedas = "" Then RetornoTCMonedas = aTC Else RetornoTCMonedas = RetornoTCMonedas & ":" & aTC
        Contador = Contador + 1
        RsAux.MoveNext
    Loop
    RsAux.Close
End Function

Private Sub Label1_Click()
    Foco tMes
End Sub

Private Sub CargoInformacionCosteo()

    On Error GoTo errInfo
    Cons = "Select * from CMCabezal Where CabId In (Select Max(CabID) from CMCabezal)"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    lCMes.Caption = ""
    lCFecha.Caption = "": lCUsuario.Caption = ""
    
    If Not RsAux.EOF Then
        lCMes.Caption = Format(RsAux!CabMesCosteo, "Mmmm yyyy")
        lCFecha.Caption = Format(RsAux!CabFecha, "dd/mm/yyyy hh:mm")
        lCUsuario.Caption = miConexion.UsuarioLogueado(Nombre:=True)
    End If
    
    If IsDate(lCMes.Caption) Then tMes.Text = Format(DateAdd("m", 1, CDate(lCMes.Caption)), "Mmmm yyyy") Else tMes.Text = ""
    
    RsAux.Close
    Exit Sub
    
errInfo:
    clsGeneral.OcurrioError "Ocurrió un error al cargar la información del último costeo.", Err.Description
End Sub

Private Sub tMes_GotFocus()
    With tMes: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsDate(tMes.Text) Then Foco bConsultar
End Sub

Private Sub tMes_LostFocus()
    If IsDate(tMes.Text) Then tMes.Text = Format(tMes.Text, "Mmmm yyyy")
End Sub

Private Function ValidoCampos() As Boolean

    On Error GoTo errValido
    ValidoCampos = False
    If Not IsDate(tMes.Text) Then
        MsgBox "El mes ingresado para realizar el costeo no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tMes: Exit Function
    End If
    
    If Trim(lCMes.Caption) <> "" Then
        If CDate(lCMes.Caption) >= CDate(tMes.Text) Then
            MsgBox "El mes ingresado para realizar el costeo debe ser mayor al último mes costeado.", vbExclamation, "ATENCIÓN"
            Foco tMes: Exit Function
        End If
        
        If Abs(DateDiff("m", CDate(lCMes.Caption), CDate(tMes.Text))) <> 1 Then
            If MsgBox("El mes ingresado para realizar el costeo no es el siguiente al último mes costeado." & Chr(vbKeyReturn) _
                    & "Está seguro de costear el mes ingresado.", vbYesNo + vbDefaultButton2 + vbExclamation, "ATENCIÓN") = vbNo Then
                Foco tMes: Exit Function
            End If
        End If
    End If
    
    ValidoCampos = True
    Exit Function

errValido:
    clsGeneral.OcurrioError "Ocurrió un error al validar datos.", Err.Description
End Function

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)

Dim aTexto As String

    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    If bCargarImpresion Then
        If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub
        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        aTexto = "Costeo de Mercadería - " & vsConsulta.Tag
        EncabezadoListado vsListado, aTexto, False
        vsListado.filename = "Costeo de Mercadería"
         
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hWnd: vsConsulta.ExtendLastCol = True
        
        vsListado.EndDoc
    End If
    
    If Imprimir Then
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    Screen.MousePointer = 0
    
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub


Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "<Descripción|<Fecha|Tipo|Documento|Artículo|>Q|>Costo (x1)|>Total"
        
        .ColWidth(1) = 1000: .ColWidth(2) = 1000: .ColWidth(3) = 1200: .ColWidth(4) = 2600: .ColWidth(5) = 400: .ColWidth(6) = 1100: .ColWidth(7) = 1100
        .WordWrap = True
    End With
      
End Sub

