VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmCosteo 
   Caption         =   "Costeo de Mercaderia"
   ClientHeight    =   6900
   ClientLeft      =   1995
   ClientTop       =   2775
   ClientWidth     =   9900
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
   ScaleHeight     =   6900
   ScaleWidth      =   9900
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
      Height          =   1035
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9255
      Begin VB.CommandButton bCosteos 
         Caption         =   "Lista Costeos"
         Height          =   315
         Left            =   7680
         TabIndex        =   51
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton bBorrar 
         Caption         =   "Borrar Costeo"
         Height          =   315
         Left            =   7680
         TabIndex        =   50
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton bService 
         Caption         =   "Verifica Service"
         Height          =   315
         Left            =   6120
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton bRespaldar 
         Caption         =   "Respaldar"
         Height          =   315
         Left            =   6120
         TabIndex        =   48
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton bDepurar 
         Caption         =   "Depurar"
         Height          =   315
         Left            =   4860
         TabIndex        =   47
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lCUsuario 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   285
         Left            =   3780
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   3060
         TabIndex        =   11
         Top             =   300
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
      Top             =   1260
      Width           =   9675
      Begin VB.CommandButton bHelp 
         Caption         =   "Ayuda"
         Height          =   315
         Left            =   6120
         TabIndex        =   55
         Top             =   220
         Width           =   1455
      End
      Begin VB.CommandButton bCopiarE 
         Caption         =   "Copiar Ex y Reb"
         Height          =   315
         Left            =   7680
         TabIndex        =   52
         Top             =   220
         Width           =   1455
      End
      Begin VB.TextBox tSHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   46
         Top             =   540
         Width           =   1575
      End
      Begin VB.TextBox tSDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   44
         Top             =   540
         Width           =   1575
      End
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   4440
         ScaleHeight     =   405
         ScaleWidth      =   1575
         TabIndex        =   3
         Top             =   180
         Width           =   1575
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
            Left            =   600
            Picture         =   "frmCosteo.frx":1F80
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "Sucesos."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bCancelar 
            Height          =   310
            Left            =   1080
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
            Left            =   180
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
         Top             =   1260
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lFilas 
         Height          =   255
         Left            =   4500
         TabIndex        =   54
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Proceso: "
         Height          =   255
         Left            =   3720
         TabIndex        =   53
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "&Servicios Gtía.:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   540
         Width           =   1215
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
         Top             =   900
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
      Top             =   6645
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "sucursal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
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
            Object.Width           =   9260
            Key             =   ""
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
    Compra = 1              'Compra Comun (a proveedores de mercaderia locales)
    Comercio = 2            'Cualquier documento del comercio (ctdo, cred, etc...)
    Importacion = 3        'Compra (que entra por importaciones)
    Servicio = 4              'Documento ralacionado a Servicios (Ventas por servicios no facturados)
End Enum

Dim rsAux As rdoResultset
Dim bCargarImpresion  As Boolean

Private Sub AccionCostear()

Dim aIDCosteo As Long
    
    If Not ValidoCampos Then Exit Sub
    If MsgBox("Confirma realizar el costeo de mercadería para el mes de " & tMes.Text, vbQuestion + vbYesNo, "Costear Mercadería") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    On Error GoTo errCostear
    Dim bOk As Boolean
    
    vsConsulta.Tag = tMes.Text: vsConsulta.Rows = 1
    FechaDelServidor
    
    '-------------------------------------------------------------------------------------------------------------------------------
    cons = "Select * from CMCabezal"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsAux.AddNew
    rsAux!CabMesCosteo = Format(tMes.Text, sqlFormatoF)
    rsAux!CabFecha = Format(gFechaServidor, sqlFormatoFH)
    rsAux!CabUsuario = paCodigoDeUsuario
    rsAux.Update: rsAux.Close
    
    cons = "Select Max(CabID) from CMCabezal"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    aIDCosteo = rsAux(0)
    rsAux.Close
    '-------------------------------------------------------------------------------------------------------------------------------
    
    lAccion.Caption = "Procesando compras del mes...": lAccion.Refresh
    lCompraI.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCompraI.Refresh
    CargoTablaCMCompra CDate(tMes.Text)
    lCompraF.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCompraF.Refresh
    
    lAccion.Caption = "Procesando ventas del mes...": lAccion.Refresh
    lVentaI.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lVentaI.Refresh
    bOk = CargoTablaCMVenta(CDate(tMes.Text))
    lVentaF.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lVentaF.Refresh
    
    If Not bOk Then
        If MsgBox("Hubo errores al copiar los registros de ventas" & vbCrLf & _
                    "Continúa con el consteo ?.", vbQuestion + vbYesNo + vbDefaultButton2, "Error al Copiar las Ventas") = vbNo Then
                    
                CargoInformacionCosteo

                If IsDate(tMes.Text) Then
                    tSDe.Text = Format(PrimerDia(tMes.Text), "dd/mm/yyyy")
                    tSHasta.Text = Format(UltimoDia(tMes.Text), "dd/mm/yyyy")
                End If
                Screen.MousePointer = 0
                Exit Sub
        End If
    End If
                    
    lAccion.Caption = "Costeando Mercadería...": lAccion.Refresh
    lCosteoI.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCosteoI.Refresh
    CargoTablaCMCosteo aIDCosteo
    lCosteoF.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCosteoF.Refresh
    lAccion.Caption = "Costeo Finalizado OK.": lAccion.Refresh
    
    MsgBox "El costeo para el mes de " & Trim(tMes.Text) & " se ha finalizado con éxito.", vbExclamation, "Costeo Finalizado"
    CargoInformacionCosteo
    
    If IsDate(tMes.Text) Then
        tSDe.Text = Format(PrimerDia(tMes.Text), "dd/mm/yyyy")
        tSHasta.Text = Format(UltimoDia(tMes.Text), "dd/mm/yyyy")
    End If
    
    Screen.MousePointer = 0
    Exit Sub

errCostear:
    clsGeneral.OcurrioError "Error al costear la mercadería.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoTablaCMCompra(Mes As Date)

    '1) Cargar las compras del mes (Credito y Contado)
    '2) Cargar las importaciones del mes (con fecha de arribo del costeo en el mes)
    
    'ATENCIÓN: solo van los articulos que tengan el campo ArtAMercaderia = True !!!!!!!
        
Dim aCosto As Currency
Dim QyCom As rdoQuery
    lFilas.Tag = 0
    cons = "Insert Into CMCompra (ComFecha, ComArticulo, ComTipo, ComCodigo, ComCantidad, ComCosto, ComQOriginal) Values (?,?,?,?,?,?,?)"
    Set QyCom = cBase.CreateQuery("", cons)
    
    '1) Compras del Mes (contado y credito)------------------------------------------------------------------------------------------------------------------------

    cons = "Select * from Compra, CompraRenglon, Articulo" _
          & " Where ComCodigo = CReCompra" _
          & " And ComFecha Between '" & Format(Mes, sqlFormatoFH) & "' And '" & Format(UltimoDia(Mes) & " 23:59", sqlFormatoFH) & "'" _
          & " And ComTipoDocumento In (" & TipoDocumento.CompraContado & ", " & TipoDocumento.CompraCredito & ")" _
          & " And CReArticulo = ArtID" _
          & " And ArtAMercaderia = 1"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        QyCom.rdoParameters(0) = rsAux!ComFecha
        QyCom.rdoParameters(1) = rsAux!CReArticulo
        QyCom.rdoParameters(2) = TipoCV.Compra
        QyCom.rdoParameters(3) = rsAux!ComCodigo
        
        QyCom.rdoParameters(4) = rsAux!CReCantidad
        QyCom.rdoParameters(6) = rsAux!CReCantidad
        
        If rsAux!ComMoneda <> paMonedaPesos Then
            aCosto = rsAux!CRePrecioU * rsAux!ComTC
        Else
            aCosto = rsAux!CRePrecioU
        End If
        QyCom.rdoParameters(5) = aCosto
        
        QyCom.Execute
        
        rsAux.MoveNext
        
        lFilas.Tag = Val(lFilas.Tag) + 1: lFilas.Caption = lFilas.Tag
        If (Val(lFilas.Tag) Mod 10) = 0 Then lFilas.Refresh
    Loop
    rsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim rsCom As rdoResultset
    '2) Importaciones Costeadas en el Mes-------------------------------------------------------------------------------------------------------------------------
            'Primero las carpetas, Segundo los componentes para cada carpeta
    cons = "Select * from CosteoCarpeta, CosteoArticulo, Articulo" _
           & " Where CCaFArribo Between '" & Format(Mes, sqlFormatoFH) & "' And '" & Format(UltimoDia(Mes) & " 23:59", sqlFormatoFH) & "'" _
           & " And CCaID = CArIDCosteo" _
           & " And CArIdArticulo = ArtId " _
           & " And ArtAMercaderia = 1"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not rsAux.EOF
        QyCom.rdoParameters(0) = rsAux!CCaFArribo
        QyCom.rdoParameters(1) = rsAux!CArIdArticulo
        QyCom.rdoParameters(2) = TipoCV.Importacion
        QyCom.rdoParameters(3) = rsAux!CCaID
        
        QyCom.rdoParameters(4) = rsAux!CArCantidad
        QyCom.rdoParameters(6) = rsAux!CArCantidad
        QyCom.rdoParameters(5) = rsAux!CArCostoP
        QyCom.Execute
                
        rsAux.MoveNext
        
        lFilas.Tag = Val(lFilas.Tag) + 1: lFilas.Caption = lFilas.Tag
        If (Val(lFilas.Tag) Mod 10) = 0 Then lFilas.Refresh
        
    Loop
    rsAux.Close
    
    'Busca los componetes: Articulos que estan an el remito y no en el embarque----------------------------------------------
    cons = "Select * from CosteoCarpeta " _
           & " Where CCaFArribo Between '" & Format(Mes, sqlFormatoFH) & "' And '" & Format(UltimoDia(Mes) & " 23:59", sqlFormatoFH) & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsAux.EOF
        
        cons = "Select * from RemitoCompra, RemitoCompraRenglon, Articulo" & _
                    " Where RCoCodigo = RCRRemito" & _
                    " And RCoTipoFolder = " & rsAux!CCaNivelFolder & _
                    " And RCoIDFolder = " & rsAux!CCaFolder & _
                    " And RCRArticulo Not in (Select AFoArticulo from ArticuloFolder Where AFoTipo = RCoTipoFolder And AFoCodigo = RCoIdFolder)" & _
                    " And RCRArticulo = ArtID" & _
                    " And ArtAMercaderia = 1"
        Set rsCom = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not rsCom.EOF
            QyCom.rdoParameters(0) = rsAux!CCaFArribo
            QyCom.rdoParameters(1) = rsCom!RCRArticulo
            QyCom.rdoParameters(2) = TipoCV.Importacion
            QyCom.rdoParameters(3) = rsAux!CCaID
            
            QyCom.rdoParameters(4) = rsCom!RCRCantidad
            QyCom.rdoParameters(6) = rsCom!RCRCantidad
            QyCom.rdoParameters(5) = 0
            QyCom.Execute
            rsCom.MoveNext
        Loop
        rsCom.Close
        '----------------------------------------------------------------------------------------------------------------------------------------
        
        lFilas.Tag = Val(lFilas.Tag) + 1: lFilas.Caption = lFilas.Tag
        If (Val(lFilas.Tag) Mod 10) = 0 Then lFilas.Refresh
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    QyCom.Close
    
    lFilas.Refresh
    
End Sub

Private Sub CargoTablaCMCosteo(IDCosteo As Long)

Dim QyCos As rdoQuery
Dim RsVen As rdoResultset, rsCom As rdoResultset

Dim aFVenta As Date, aArticulo As Long
Dim aQVenta As Long, aQCompra As Long, aQCosteo As Long
Dim aQVentaOriginal As Long
Dim bBorroVenta As Boolean
    
    lFilas.Tag = 0
    'Preparo las Queries para costear-----------------------------------------------------------------------------------------------------------------------------
    cons = "Insert Into CMCosteo (CosID, CosArticulo, CosTipoVenta, CosIDVenta, CosCantidad, CosCosto, CosVenta) Values (?,?,?,?,?,?,?)"
    Set QyCos = cBase.CreateQuery("", cons)
    QyCos.rdoParameters(0) = IDCosteo

    Dim QyCMenores As rdoQuery, QyCMayores As rdoQuery
    
    cons = "Select * from CMCompra " _
            & " Where ComFecha <= ?" _
            & " And ComArticulo = ? " _
            & " And ComCantidad > 0 " _
            & " Order by ComFecha DESC"
    Set QyCMenores = cBase.CreateQuery("", cons)
    
    cons = "Select * from CMCompra " _
            & " Where ComFecha >= ?" _
            & " And ComArticulo = ? " _
            & " And ComCantidad > 0 "
    Set QyCMayores = cBase.CreateQuery("", cons)
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '-------------------------------------------------------------------------------------------------------------------
    cons = "Select Count(*) from CMVenta"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux(0) <> 0 Then bProgress.Max = rsAux(0)
    rsAux.Close
    bProgress.Value = 0
    '-------------------------------------------------------------------------------------------------------------------

    cons = "Select * from CMVenta, Articulo" _
           & " Where VenArticulo = ArtID " _
           & " Order by VenFecha, VenArticulo, VenTipo, VenCodigo"
    Set RsVen = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
       
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
                QyCMenores.rdoParameters(0) = Format(aFVenta, sqlFormatoF) & " 23:59"
                QyCMenores.rdoParameters(1) = aArticulo
                Set rsCom = QyCMenores.OpenResultset(rdOpenDynamic, rdConcurValues)
                
                If Not rsCom.EOF Then               'Hay una FC <= FV
                    If aQVenta > 0 Then                 'VENTA DE MERCADERIA---------------------------------------------------
                        aQCompra = rsCom!ComCantidad
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
                        QyCos.rdoParameters(5) = rsCom!ComCosto
                        QyCos.rdoParameters(6) = RsVen!VenPrecio
                        QyCos.Execute
                        
                        rsCom.Edit
                        rsCom!ComCantidad = rsCom!ComCantidad - aQCosteo
                        rsCom.Update
                    
                    Else        'DEVOLUCION DE MERCADERIA---------------------------------------------------
                                  'La cantidad debe ser siempre menor a la original, sino voy al inmediato anterior (x q voy a sumar 1 sino me paso)
                                  'IRMA: la sumamos igual, no importa si nos pasamos
                        aQCompra = rsCom!ComCantidad
                        aQCosteo = aQVenta      'QVenta es negativa --> devolucion
                        aQVenta = 0
                                                        
                        QyCos.rdoParameters(2) = RsVen!VenTipo
                        QyCos.rdoParameters(3) = RsVen!VenCodigo
                        QyCos.rdoParameters(4) = aQCosteo
                        QyCos.rdoParameters(5) = rsCom!ComCosto
                        QyCos.rdoParameters(6) = RsVen!VenPrecio
                        QyCos.Execute
                        
                        rsCom.Edit
                        rsCom!ComCantidad = rsCom!ComCantidad - aQCosteo        '- * - = +
                        rsCom.Update
                    End If
                    rsCom.Close
    
                Else                                        'NO Hay una FC <= FV
                    rsCom.Close
                    'Voy a la minima fecha de Compra >= a la fecha de venta------------------------------------
                    QyCMayores.rdoParameters(0) = Format(aFVenta, sqlFormatoF) & " 23:59"
                    QyCMayores.rdoParameters(1) = aArticulo
                    Set rsCom = QyCMayores.OpenResultset(rdOpenDynamic, rdConcurValues)
                    
                    If Not rsCom.EOF Then               'Hay una FC >= FV
                    
                        If aQVenta > 0 Then                 'VENTA DE MERCADERIA---------------------------------------------------
                            aQCompra = rsCom!ComCantidad
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
                            QyCos.rdoParameters(5) = rsCom!ComCosto
                            QyCos.rdoParameters(6) = RsVen!VenPrecio
                            QyCos.Execute
                            
                            rsCom.Edit
                            rsCom!ComCantidad = rsCom!ComCantidad - aQCosteo
                            rsCom.Update
                        
                        Else        'DEVOLUCION DE MERCADERIA---------------------------------------------------
                                  'La cantidad debe ser siempre menor a la original, sino voy al inmediato siguiente
                                  'Cambiamos, siempre le sumamos  no importa si me paso en la QdeCompra !!!! 22/5/00
                            aQCompra = rsCom!ComCantidad
                            aQCosteo = aQVenta
                            aQVenta = 0
                            
                            QyCos.rdoParameters(2) = RsVen!VenTipo
                            QyCos.rdoParameters(3) = RsVen!VenCodigo
                            QyCos.rdoParameters(4) = aQCosteo
                            QyCos.rdoParameters(5) = rsCom!ComCosto
                            QyCos.rdoParameters(6) = RsVen!VenPrecio
                            QyCos.Execute
                        
                            rsCom.Edit
                            rsCom!ComCantidad = rsCom!ComCantidad - aQCosteo
                            rsCom.Update
                            rsCom.Close
                        End If
                        
                    Else
                        rsCom.Close
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
                            
                            '15/12/00 Agrego una compra inventada a la existencia (con la Q que se devuelve)
                            AgregoRegistroCMCompra RsVen!VenFecha, RsVen!VenArticulo, Abs(aQVenta)
                            
                            aQVenta = 0: bBorroVenta = True
                        Else
                        
                            If aQVenta <> aQVentaOriginal Then
                                cons = "Update CMVenta Set VenCantidad = " & aQVenta _
                                        & " Where VenFecha = '" & Format(RsVen!VenFecha, sqlFormatoFH) & "'" _
                                        & " And VenArticulo = " & RsVen!VenArticulo _
                                        & " And VenTipo = " & RsVen!VenTipo & " And VenCodigo = " & RsVen!VenCodigo
                                cBase.Execute cons
                            End If
                        End If
                        Exit Do
                        
                    End If
                End If
            End If
        Loop
        
        'Si la venta quedó en cero elimino el registro de la venta
        If aQVenta = 0 And bBorroVenta Then
            cons = " Delete CMVenta " _
                    & " Where VenFecha = '" & Format(RsVen!VenFecha, sqlFormatoFH) & "'" _
                    & " And VenArticulo = " & RsVen!VenArticulo _
                    & " And VenTipo = " & RsVen!VenTipo & " And VenCodigo = " & RsVen!VenCodigo
            cBase.Execute cons
        End If
        RsVen.MoveNext
        
        lFilas.Tag = Val(lFilas.Tag) + 1: lFilas.Caption = lFilas.Tag
        If (Val(lFilas.Tag) Mod 10) = 0 Then lFilas.Refresh
    Loop
    
    RsVen.Close
    QyCMenores.Close: QyCMayores.Close
    
    'Hay que borrar las compras en que las cantidades son iguales a 0
    cons = "Delete CMCompra Where ComCantidad = 0"
    cBase.Execute cons
    '---------------------------------------------------------------------------
    bProgress.Value = 0
    
End Sub

Private Sub AgregoRegistroCMCompra(Fecha As Date, Articulo As Long, Q As Long)
    
    Dim rsReg As rdoResultset
    Dim aMin As Long
    
    On Error GoTo errGrabar
    aMin = 0
    cons = "Select Min(ComCodigo) From CMCompra"
    Set rsReg = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsReg.EOF Then If Not IsNull(rsReg(0)) Then aMin = rsReg(0)
    rsReg.Close
    
    aMin = aMin - 1
    
    cons = "Select * from CMCompra " _
            & " Where ComFecha = '" & Format(Fecha, sqlFormatoF) & "'" _
            & " And ComArticulo = " & Articulo _
            & " And ComCodigo = " & aMin _
            & " And ComTipo = " & TipoCV.Comercio
    Set rsReg = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsReg.EOF Then
        rsReg.AddNew
        
        rsReg!ComFecha = Format(Fecha, sqlFormatoF)
        rsReg!ComArticulo = Articulo
        rsReg!ComCantidad = Q
        rsReg!ComCodigo = aMin
        rsReg!ComTipo = TipoCV.Comercio
        rsReg!ComCosto = 0
        rsReg!ComQOriginal = Q
        rsReg.Update
    Else
        AgegoSuceso "Error (Agregar Compra)", Fecha, Articulo, TipoCV.Compra, aMin, 0, Q
    End If
    rsReg.Close
    Exit Sub

errGrabar:
    AgegoSuceso "Error (Agregar Compra)", Fecha, Articulo, TipoCV.Compra, aMin, 0, Q
End Sub


Private Sub AgegoSuceso(Texto As String, Fecha As Date, Articulo As Long, Tipo As Integer, Id As Long, Precio As Currency, Cantidad As Long)

Dim rsS As rdoResultset

    'On Error Resume Next
    With vsConsulta
        .AddItem ""     '"<Descripción|<Fecha|Tipo|Documento|Artículo|>Q|>Costo (x1)|>Total"
        .Cell(flexcpText, .Rows - 1, 0) = Trim(Texto)
        .Cell(flexcpText, .Rows - 1, 1) = Format(Fecha, "dd/mm/yyyy")
        
        Select Case Tipo
            Case TipoCV.Comercio:
                cons = "Select * from Documento Where DocCodigo = " & Id
                Set rsS = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                If Not rsS.EOF Then
                    .Cell(flexcpText, .Rows - 1, 2) = "Comercio"
                    .Cell(flexcpText, .Rows - 1, 3) = RetornoNombreDocumento(rsS!DocTipo, Abreviacion:=True) & " "
                    .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & Trim(rsS!DocSerie) & Format(rsS!DocNumero, "000000")
                End If
                rsS.Close
        
            Case TipoCV.Compra:
                cons = "Select * from Compra Where ComCodigo = " & Id
                Set rsS = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                If Not rsS.EOF Then
                    .Cell(flexcpText, .Rows - 1, 2) = "Compras"
                    .Cell(flexcpText, .Rows - 1, 3) = RetornoNombreDocumento(rsS!ComTipoDocumento, Abreviacion:=True) & " "
                    If Not IsNull(rsS!ComSerie) Then .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & Trim(rsS!ComSerie)
                    If Not IsNull(rsS!ComNumero) Then .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & Trim(rsS!ComNumero) & " "
                End If
                rsS.Close
        End Select
        
        cons = "Select * from Articulo Where ArtId = " & Articulo
        Set rsS = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsS.EOF Then .Cell(flexcpText, .Rows - 1, 4) = Format(rsS!ArtCodigo, "(#,000,000)") & " " & Trim(rsS!ArtNombre)
        rsS.Close
        
        .Cell(flexcpText, .Rows - 1, 5) = Cantidad
        .Cell(flexcpText, .Rows - 1, 6) = Format(Precio, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 7) = Format(Precio * Cantidad, FormatoMonedaP)
        
    End With
    
End Sub

Private Sub bBorrar_Click()
    If MsgBox("Esta acción borrará el último costeo y se reestablecerán los datos al costeo anterior." & vbCrLf & _
                    "Está seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Costeo") = vbNo Then Exit Sub
    
    rspEliminarUltimoCosteo
    
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    
    If Not rspVerificoRespaldoUltimoCosteo Then
        If MsgBox("El último Costeo no se respaldó." & vbCrLf & _
                        "Está seguro de continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Falta Respaldar el Costeo") = vbNo Then Exit Sub
                    
        Exit Sub
    End If
    
    LimpioDatos
    AccionCostear
End Sub

Private Sub bCopiarE_Click()
    
    If MsgBox("Esta acción copia las existencias y rebotes del último costeo respaldado." & vbCrLf & _
                    "Está seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2, "Copiar Existencia y Rebotes del Respaldo") = vbNo Then Exit Sub
    
    rspCopiarExistencia

End Sub

Private Sub bCosteos_Click()
On Error GoTo errLista
        
        cons = " Select CabID, CabID as 'Código', CabMesCosteo as 'Mes Costeado', CabFecha as 'Costeado el ...' " & _
                   " from CMCabezal " & _
                   " Order by CabID Desc"

        Dim myLista As New clsListadeAyuda
        myLista.ActivarAyuda cBase, cons, 5600, 1
        Set myLista = Nothing

        Screen.MousePointer = 0
        Exit Sub
errLista:
    clsGeneral.OcurrioError "Error al ver la lista de Costeos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bDepurar_Click()
    
    If Val(lCMes.Tag) = 0 Then Exit Sub
    If MsgBox("Confirma depurar el costeo. " & vbCrLf & "Esto tiene efecto si ud. corrigió la existencia (para eliminar los rebotes).", vbQuestion + vbYesNo, "Depurar ? ") = vbNo Then Exit Sub
    
    lAccion.Caption = "Depurando Mercadería...": lAccion.Refresh
    lCosteoI.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCosteoI.Refresh
    CargoTablaCMCosteo Val(lCMes.Tag)
    lCosteoF.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCosteoF.Refresh
    lAccion.Caption = "Costeo Finalizado OK.": lAccion.Refresh
    
    MsgBox "Tarea finalizada con éxito.", vbExclamation, "Costeo Finalizado"
    
End Sub

Private Sub bHelp_Click()
    AccionMenuHelp
End Sub

Private Sub bImprimir_Click()
    AccionImprimir True
End Sub

Private Sub bPrimero_Click()
    IrAPagina vsListado, 1
End Sub

Private Sub bRespaldar_Click()
    rspRespaldarCosteo
End Sub

Private Sub bService_Click()
    rspVerificaService
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
    
    If IsDate(tMes.Text) Then
        tSDe.Text = Format(PrimerDia(tMes.Text), "dd/mm/yyyy")
        tSHasta.Text = Format(UltimoDia(tMes.Text), "dd/mm/yyyy")
    End If
    
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
    lFilas.Caption = ""
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

Private Function CargoTablaCMVenta(Mes As Date) As Boolean
'Parámetro: Recibe el primer día del mes a costear.
On Error GoTo errCTCMV

Dim QyVta As rdoQuery
Dim strDocumentos As String
Dim Monedas As String
Dim aCosto As Currency
Dim aMoneda As Long, aTC As Currency, aNeto As Currency

    CargoTablaCMVenta = False
    Screen.MousePointer = 11
    lFilas.Tag = 0
    
    Monedas = RetornoTCMonedas(PrimerDia(Mes))
    aMoneda = 0
    'Creo query para insertar los datos
    cons = "Insert Into CMVenta (VenFecha, VenArticulo, VenTipo, VenCodigo, VenCantidad,VenPrecio) Values (?,?,?,?,?,?)"
    Set QyVta = cBase.CreateQuery("", cons)
    
    'Primer Paso Copio las Ventas---------------------------------------
    'Traigo los documentos Ctdo y Cred, Nota Esp, Nota de Cred. y  Nota de Dev. que no estén anulados
    strDocumentos = TipoDocumento.Contado & ", " & TipoDocumento.Credito _
        & ", " & TipoDocumento.NotaCredito & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial
    
    
    cons = "Select DocFecha, DocMoneda, DocTipo, Renglon.* " _
        & " From Documento, Renglon" _
        & " Where DocTipo IN (" & strDocumentos & ")" _
        & " And DocFecha BetWeen '" & Format(PrimerDia(Mes) & " 00:00:00", sqlFormatoFH) & "'" _
                                    & " And '" & Format(UltimoDia(Mes) & " 23:59:59", sqlFormatoFH) & "'" _
        & " And DocAnulado = 0 And DocCodigo = RenDocumento"

qyVentas:
    On Error GoTo errQTOut
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)

    On Error GoTo errCTCMV
    
    Do While Not rsAux.EOF
        QyVta.rdoParameters(0) = Format(rsAux!DocFecha, sqlFormatoF)
        QyVta.rdoParameters(1) = rsAux!RenArticulo
        QyVta.rdoParameters(2) = TipoCV.Comercio
        QyVta.rdoParameters(3) = rsAux!RenDocumento
        
        Select Case rsAux!DocTipo
            Case TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial
                            QyVta.rdoParameters(4) = rsAux!RenCantidad * -1
                            
            Case Else: QyVta.rdoParameters(4) = rsAux!RenCantidad
        End Select
        
        'Es el precio neto-----------------------------------------------------------------------------
        If aMoneda <> rsAux!DocMoneda Then
            aMoneda = rsAux!DocMoneda
            aTC = ValorTC(rsAux!DocMoneda, Monedas)
        End If
        
        aNeto = rsAux!RenPrecio - rsAux!RenIva
        If Not IsNull(rsAux!RenCofis) Then aNeto = aNeto - rsAux!RenCofis
        QyVta.rdoParameters(5) = aNeto * aTC
        'QyVta.rdoParameters(5) = (rsAux!RenPrecio - rsAux!RenIva) * aTC
        
        '-------------------------------------------------------------------------------------------------
        QyVta.Execute
        
        rsAux.MoveNext
        
        lFilas.Tag = Val(lFilas.Tag) + 1: lFilas.Caption = lFilas.Tag
        If (Val(lFilas.Tag) Mod 10) = 0 Then lFilas.Refresh
        
    Loop
    rsAux.Close
    
    
    'Segundo paso Cargo Notas de Compras.
    cons = "Select * from Compra, CompraRenglon" _
          & " Where ComCodigo = CReCompra" _
          & " And ComFecha Between '" & Format(Mes, sqlFormatoFH) & "' And '" & Format(UltimoDia(Mes) & " 23:59", sqlFormatoFH) & "'" _
          & " And ComTipoDocumento In (" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ")"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)

    Do While Not rsAux.EOF
        QyVta.rdoParameters(0) = Format(rsAux!ComFecha, sqlFormatoF)
        QyVta.rdoParameters(1) = rsAux!CReArticulo
        QyVta.rdoParameters(2) = TipoCV.Compra
        QyVta.rdoParameters(3) = rsAux!CReCompra
        QyVta.rdoParameters(4) = rsAux!CReCantidad
        If rsAux!ComMoneda <> paMonedaPesos Then
            aCosto = rsAux!CRePrecioU * rsAux!ComTC
        Else
            aCosto = rsAux!CRePrecioU
        End If
        QyVta.rdoParameters(5) = aCosto
        QyVta.Execute
        rsAux.MoveNext
        
        lFilas.Tag = Val(lFilas.Tag) + 1: lFilas.Caption = lFilas.Tag
        If (Val(lFilas.Tag) Mod 10) = 0 Then lFilas.Refresh
        
    Loop
    rsAux.Close
    
'    '21/5/2001 - Cargo los servicios con costo que no fueron facturados y estan cumplidos--------------------------------------------------------
'    '               Como no fueron facturados, los servicios van a entrar con costo de venta 0
'    cons = "Select * From Servicio, ServicioRenglon" & _
'               " Where SerCodigo = SReServicio " & _
'               " And SerEstadoServicio = " & EstadoS.Cumplido & _
'               " And SerDocumento Is Null " & _
'               " And SerFCumplido Between '" & Format(tSDe.Text, sqlFormatoFH) & "' And '" & Format(tSHasta.Text & " 23:59", sqlFormatoFH) & "'" & _
'               " And SReTipoRenglon = " & TipoRenglonS.Cumplido
    
'   [2017-03-27] Cambio consulta para que no entren al costeo los servicio que no fueron "Arreglados"
    cons = "Select * From Servicio " & _
                " INNER JOIN ServicioRenglon ON SerCodigo = SReServicio " & _
                " INNER JOIN Taller ON Talservicio = SerCodigo AND TalAceptado = 1 AND TalSinArreglo = 0 " & _
            " Where SerEstadoServicio = " & EstadoS.Cumplido & _
            " And SerDocumento Is Null " & _
            " And SerFCumplido Between '" & Format(tSDe.Text, sqlFormatoFH) & "' And '" & Format(tSHasta.Text & " 23:59", sqlFormatoFH) & "'" & _
            " And SReTipoRenglon = " & TipoRenglonS.Cumplido
   
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        QyVta.rdoParameters(0) = Format(rsAux!SerFCumplido, sqlFormatoF)
        QyVta.rdoParameters(1) = rsAux!SReMotivo
        QyVta.rdoParameters(2) = TipoCV.Servicio
        QyVta.rdoParameters(3) = rsAux!SerCodigo
        QyVta.rdoParameters(4) = rsAux!SReCantidad
        QyVta.rdoParameters(5) = 0
        QyVta.Execute
        
        rsAux.MoveNext
        
        lFilas.Tag = Val(lFilas.Tag) + 1: lFilas.Caption = lFilas.Tag
        If (Val(lFilas.Tag) Mod 10) = 0 Then lFilas.Refresh
    Loop
    rsAux.Close
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    QyVta.Close 'Cierro query
    
    lFilas.Refresh
    Screen.MousePointer = 0
    CargoTablaCMVenta = True
    Exit Function
    
errCTCMV:
    clsGeneral.OcurrioError "Error al cargar la tabla de Ventas.", Err.Description
    Screen.MousePointer = 0
    Exit Function

errQTOut:
    If MsgBox("Error al ejecutar la consulta de datos." & vbCrLf & _
                Err.Number & "- " & Err.Description & vbCrLf & vbCrLf & _
                "Reintenta hacer la consulta ?", vbQuestion + vbYesNo, "Error al Consultar") = vbYes Then
        Resume qyVentas
    End If
End Function

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

Private Function RetornoTCMonedas(Fecha As Date) As String

Dim aTC As Currency, Contador As Integer

    'Armo vector con las TC de las monedas que existen.
    RetornoTCMonedas = ""
    cons = "Select * From Moneda"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
    Contador = 1
    Do While Not rsAux.EOF
        If Contador = rsAux!MonCodigo Then
            If rsAux!MonCodigo = paMonedaPesos Then
                aTC = 1
            Else
                aTC = TasadeCambio(rsAux!MonCodigo, paMonedaPesos, Fecha)
            End If
        Else
            aTC = 1
        End If
        If RetornoTCMonedas = "" Then RetornoTCMonedas = aTC Else RetornoTCMonedas = RetornoTCMonedas & ":" & aTC
        Contador = Contador + 1
        rsAux.MoveNext
    Loop
    rsAux.Close
End Function

Private Sub Label1_Click()
    Foco tMes
End Sub

Private Sub CargoInformacionCosteo()

    On Error GoTo errInfo
    cons = "Select * from CMCabezal Where CabId In (Select Max(CabID) from CMCabezal)"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    lCMes.Caption = ""
    lCFecha.Caption = "": lCUsuario.Caption = ""
    
    If Not rsAux.EOF Then
        lCMes.Caption = Format(rsAux!CabMesCosteo, "Mmmm yyyy")
        lCMes.Tag = rsAux!CabID
        lCFecha.Caption = Format(rsAux!CabFecha, "dd/mm/yyyy hh:mm")
        lCUsuario.Caption = miConexion.UsuarioLogueado(Nombre:=True)
    End If
    
    If IsDate(lCMes.Caption) Then tMes.Text = Format(DateAdd("m", 1, CDate(lCMes.Caption)), "Mmmm yyyy") Else tMes.Text = ""
    
    rsAux.Close
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
    
    If Not IsDate(tSDe.Text) Then
        MsgBox "El mes ingresado para los servicios en garantía no es correcto.", vbExclamation, "Posible Error"
        Foco tSDe: Exit Function
    End If
    If Not IsDate(tSHasta.Text) Then
        MsgBox "El mes ingresado para los servicios en garantía no es correcto.", vbExclamation, "Posible Error"
        Foco tSHasta: Exit Function
    End If
    If CDate(tSDe.Text) > CDate(tSHasta.Text) Then
        MsgBox "El período para los servicios en garantía no es correcto.", vbExclamation, "Posible Error"
        Foco tSDe: Exit Function
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
        vsListado.FileName = "Costeo de Mercadería"
         
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        
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

Private Sub tSDe_GotFocus()
    With tSDe: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tSDe_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If IsDate(tSDe.Text) Then tSDe.Text = Format(PrimerDia(tSDe.Text), "dd/mm/yyyy")
        tSHasta.SetFocus
    End If
    
End Sub

Private Sub tSHasta_GotFocus()
    With tSHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tSHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tSHasta.Text) Then tSHasta.Text = Format(UltimoDia(tSHasta.Text), "dd/mm/yyyy")
        bConsultar.SetFocus
    End If
End Sub

Private Sub rspRespaldarCosteo()
On Error GoTo errRespaldo

Dim aIDCosteo As Long, aMes As Date

    Screen.MousePointer = 11
    
    rspUltimoCosteo aIDCosteo, aMes
    
    If aIDCosteo = -1 Then
        Screen.MousePointer = 0
        MsgBox "No hay Costeos para respaldar.", vbExclamation, "No Hay Datos"
        Exit Sub
    End If
    
    Dim bHay As Boolean: bHay = False
    'Valido si Hay Ventas y Existencias a Respaldar     -----------------------------------------------------------------------
    cons = "Select Top 1 * from CMCompra"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True
    rsAux.Close
        
    If Not bHay Then
        cons = "Select Top 1 * from CMVenta"
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then bHay = True
        rsAux.Close
    End If
    
    If Not bHay Then
        Screen.MousePointer = 0
        MsgBox "No hay Existencia ni Ventas Rebotadas para respaldar (en el útlimo costeo)." & vbCrLf & _
                    "ID Costeo: " & aIDCosteo & vbTab & Format(aMes, "Mmmm yyyy"), vbExclamation, "No Hay Datos a Respaldar"
        Exit Sub
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------------
    
    'Valido si ya fue respaldado        ------------------------------------------------------------------------------------------
    bHay = False
    bHay = rspVerificoRespaldoUltimoCosteo
    If bHay Then
        If MsgBox("El costeo de " & Format(aMes, "Mmmm yyyy") & " (id:" & aIDCosteo & "), ya fue respaldado." & vbCrLf & _
                    "Quiere volver a respaldarlo ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Costeo Respaldado") = vbNo Then Screen.MousePointer = 0: Exit Sub
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If MsgBox("Confirma respaldar el costeo de " & Format(aMes, "Mmmm yyyy") & " (id:" & aIDCosteo & ")", vbQuestion + vbYesNo + vbDefaultButton2, "Respaldar Costeo ?") = vbNo Then Screen.MousePointer = 0: Exit Sub
    
    Me.Refresh
    If bHay Then
        cons = "Delete rspCMCompra Where ComCosteo = " & aIDCosteo
        cBase.Execute cons
        
        cons = "Delete rspCMVenta Where VenCosteo = " & aIDCosteo
        cBase.Execute cons
    End If
    
    Dim rsIns As rdoResultset
    Dim mQ As Long
    
    bProgress.Value = 0
    mQ = 0
    
    cons = "Select Count(*) from CMCompra"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux(0)) Then mQ = rsAux(0)
    End If
    rsAux.Close

    If mQ > 0 Then
        bProgress.Max = mQ
        lAccion.Caption = "Respaldando " & mQ & " registros de existencia"
        Me.Refresh
            
        cons = "Select * from CMCompra"
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            cons = "Select * from rspCMCompra Where ComCosteo =" & aIDCosteo
            Set rsIns = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        
            Do While Not rsAux.EOF
                
                rsIns.AddNew
                rsIns!ComCosteo = aIDCosteo
                rsIns!ComFecha = rsAux!ComFecha
                rsIns!ComArticulo = rsAux!ComArticulo
                rsIns!ComTipo = rsAux!ComTipo
                rsIns!ComCodigo = rsAux!ComCodigo
                rsIns!ComCantidad = rsAux!ComCantidad
                rsIns!ComCosto = rsAux!ComCosto
                rsIns!ComQOriginal = rsAux!ComQOriginal
                rsIns.Update
                
                bProgress.Value = bProgress.Value + 1
                rsAux.MoveNext
            Loop
            rsIns.Close
        End If
        rsAux.Close
    End If
    
    bProgress.Value = 0
    mQ = 0
    
    cons = "Select Count(*) from CMVenta"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux(0)) Then mQ = rsAux(0)
    End If
    rsAux.Close
    
    If mQ > 0 Then
        bProgress.Max = mQ
        lAccion.Caption = "Respaldando " & mQ & " registros de Ventas Rebotadas"
        Me.Refresh

        cons = "Select * from CMVenta"
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            cons = "Select * from rspCMVenta Where VenCosteo =" & aIDCosteo
            Set rsIns = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        
            Do While Not rsAux.EOF
                
                rsIns.AddNew
                rsIns!VenCosteo = aIDCosteo
                rsIns!VenFecha = rsAux!VenFecha
                rsIns!VenArticulo = rsAux!VenArticulo
                rsIns!VenTipo = rsAux!VenTipo
                rsIns!VenCodigo = rsAux!VenCodigo
                rsIns!VenCantidad = rsAux!VenCantidad
                rsIns!VenPrecio = rsAux!VenPrecio
                rsIns.Update
                
                bProgress.Value = bProgress.Value + 1
                rsAux.MoveNext
            Loop
            rsIns.Close
        End If
        rsAux.Close
        
    End If
    
    bProgress.Value = 0
    lAccion.Caption = ""
    
    MsgBox "Costeo respaldado OK.", vbInformation, "Fin Respaldo"
    
    Screen.MousePointer = 0
    Exit Sub
    
errRespaldo:
    clsGeneral.OcurrioError "Error al respaldar el costeo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function rspUltimoCosteo(retIdCosteo As Long, retMesCosteo As Date)
    On Error GoTo errCons
    retIdCosteo = -1
    
    cons = "Select Top 1 * from CMCabezal Order by CabID DESC"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        retIdCosteo = rsAux!CabID
        retMesCosteo = rsAux!CabMesCosteo
    End If
    rsAux.Close
    Exit Function
    
errCons:
    clsGeneral.OcurrioError "Error al consultar el último mes costeado.", Err.Description
End Function

Private Sub rspVerificaService()

On Error GoTo errVerifica

    Dim bHay As Boolean
    bHay = False
    
    cons = "Select * From CMVenta " & _
               " Where VenTipo = " & TipoCV.Servicio & _
               " And VenCodigo IN (" & _
                        "Select SerCodigo From Servicio, ServicioRenglon" & _
                        " Where SerCodigo = SReServicio " & _
                        " And SerEstadoServicio = " & EstadoS.Cumplido & _
                        " And SerDocumento Is Null " & _
                        " And SerFCumplido Between '" & Format(tSDe.Text, sqlFormatoFH) & "' And '" & Format(tSHasta.Text & " 23:59", sqlFormatoFH) & "'" & _
                        " And SReTipoRenglon = " & TipoRenglonS.Cumplido & _
                ")"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True
    rsAux.Close
    
    If Not bHay Then
        MsgBox "Registros de service verificados OK.", vbInformation, "Verificación OK"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If bHay Then
        MsgBox "Hay servicios que rebotaron en el costeo anterior." & vbCrLf & _
                    "La fecha de cumplido de éstos servicios fue modificada, y van a volver a entrar este mes." & vbCrLf & _
                    "Presione Aceptar para ver los datos.", vbInformation, "Control de Servicios Modificados"
                    
        Dim myLista As New clsListadeAyuda
        myLista.ActivarAyuda cBase, cons, 5600, 1
        Set myLista = Nothing
    
    
        If MsgBox("Para que el costeo no falle, se recomienda borrar éstos registros." & vbCrLf & _
                        "Desea borrarlos ahora ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Borrar Registros") = vbYes Then
                        
            Screen.MousePointer = 11
            cons = "Select * From CMVenta " & _
                   " Where VenTipo = " & TipoCV.Servicio & _
                   " And VenCodigo IN (" & _
                            "Select SerCodigo From Servicio, ServicioRenglon" & _
                            " Where SerCodigo = SReServicio " & _
                            " And SerEstadoServicio = " & EstadoS.Cumplido & _
                            " And SerDocumento Is Null " & _
                            " And SerFCumplido Between '" & Format(tSDe.Text, sqlFormatoFH) & "' And '" & Format(tSHasta.Text & " 23:59", sqlFormatoFH) & "'" & _
                            " And SReTipoRenglon = " & TipoRenglonS.Cumplido & _
                    ")"
        
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            Do While Not rsAux.EOF
                rsAux.Delete
                rsAux.MoveNext
            Loop
            rsAux.Close
            
            Screen.MousePointer = 0
            
            MsgBox "Registros eliminados OK.", vbInformation, "Eliminación OK"
        End If
    
    End If
    Exit Sub
    
errVerifica:
    clsGeneral.OcurrioError "Error al verificar service.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub rspEliminarUltimoCosteo()
On Error GoTo errBkUp
Dim bHay As Boolean
Dim aIDCosteo As Long, aMes As Date
Dim aIDAnterior, aMesAnterior As Date

    Screen.MousePointer = 11
    
    rspUltimoCosteo aIDCosteo, aMes
    
    If aIDCosteo = -1 Then
        Screen.MousePointer = 0
        MsgBox "No hay Costeos para eliminar.", vbExclamation, "No Hay Datos"
        Exit Sub
    End If
    
    'Consulto Anteúltimo Costeo     ---------------------------------------------------------------------
    aIDAnterior = -1
    cons = "Select Top 1 * from CMCabezal Where CabID < " & aIDCosteo & " Order by CabID DESC"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        aIDAnterior = rsAux!CabID
        aMesAnterior = rsAux!CabMesCosteo
    End If
    rsAux.Close
    '---------------------------------------------------------------------------------------------------------

    Screen.MousePointer = 0
    
    If aIDAnterior <> -1 Then
        If MsgBox("Se eliminará el costeo de " & Format(aMes, "Mmmm yyyy") & " con el ID: " & aIDCosteo & vbCrLf & _
                    "Y se actualizarán los datos al costeo de " & Format(aMesAnterior, "Mmmm yyyy") & " con el ID: " & aIDAnterior & vbCrLf & vbCrLf & _
                    "Está seguro de continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Continúa ?") = vbNo Then Exit Sub
                    
        'Valido Si hay respaldo del costeo Anterior -------------------------------------------------
        bHay = False
        cons = "Select Top 1 * from rspCMCompra Where ComCosteo = " & aIDAnterior
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then bHay = True
        rsAux.Close
        
        If Not bHay Then
            cons = "Select Top 1 * from rspCMVenta Where VenCosteo = " & aIDAnterior
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then bHay = True
            rsAux.Close
        End If
    
        If Not bHay Then
            If MsgBox("No hay datos respaldados para dejar vigente un costeo anterior" & vbCrLf & _
                           "Faltan los respaldos de existencia y ventas rebotadas del costeo de " & Format(aMesAnterior, "Mmmm yyyy") & vbCrLf & vbCrLf & _
                           "Está seguro de continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Continúa ?") = vbNo Then Exit Sub
        End If
        
    Else
        If MsgBox("Se eliminará el costeo de " & Format(aMes, "Mmmm yyyy") & " con el ID: " & aIDCosteo & vbCrLf & _
                "No hay datos para dejar vigente un costeo anterior" & vbCrLf & vbCrLf & _
                "Está seguro de continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Continúa ?") = vbNo Then Exit Sub
    End If
    
    'Veo si hay respaldo para el costeo vigente     --------------------------------------------------------
    bHay = False
    bHay = rspVerificoRespaldoUltimoCosteo
    
    If bHay Then
        If MsgBox("El último costeo de " & Format(aMes, "Mmmm yyyy") & " (id:" & aIDCosteo & "), ya fue respaldado." & vbCrLf & _
                       "Se aconseja eliminar el respaldo (este costeo se eliminará) " & vbCrLf & _
                       "Elimina el respaldo ? ", vbQuestion + vbYesNo, "Eliminar Respaldado Último Costeo") = vbYes Then
            
            Screen.MousePointer = 11
            Me.Refresh
                
            lAccion.Caption = "Eliminando Registros de Existencia (respaldo)": lAccion.Refresh
            cons = "Delete rspCMCompra Where ComCosteo = " & aIDCosteo
            cBase.Execute cons
            
            lAccion.Caption = "Eliminando Registros de Rebotes (respaldo)": lAccion.Refresh
            cons = "Delete rspCMVenta Where VenCosteo = " & aIDCosteo
            cBase.Execute cons
            
            lAccion.Caption = "": lAccion.Refresh
            Screen.MousePointer = 0
        End If
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '1) Elimino Datos Tabla Compra
    lAccion.Caption = "Eliminando Registros de Existencia ": lAccion.Refresh
    cons = "Delete CMCompra"
    cBase.Execute cons
    
    '2) Elimino Datos Tabla Venta
    lAccion.Caption = "Eliminando Registros de Rebotes ": lAccion.Refresh
    cons = "Delete CMVenta"
    cBase.Execute cons
    
    '3) Elimino Datos Tabla Costeo
    lAccion.Caption = "Eliminando Costeo " & Format(aMes, "Mmmm yyyy"): lAccion.Refresh
    cons = "Delete CMCosteo Where CosID = " & aIDCosteo
    cBase.Execute cons
    
    '4) Elimino Datos Tabla Cabezal
    'lAccion.Caption = "Eliminando Cabezal " & Format(aMes, "Mmmm yyyy"): lAccion.Refresh
    'cons = "Delete CMCabezal Where CabID = " & aIDCosteo
    'cBase.Execute cons
    
    '5) Levanto respaldo Tabla Compra   -------------------------------------------------------------------------------------------
    Dim rsIns As rdoResultset
    
    bProgress.Value = 0
    
    cons = "Select Count(*) from rspCMCompra Where ComCosteo =" & aIDAnterior
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux(0)) Then
            bProgress.Max = rsAux(0)
            lAccion.Caption = "Restaurando " & rsAux(0) & " registros de existencia"
            Me.Refresh
        End If
    End If
    rsAux.Close

    cons = "Select * from rspCMCompra Where ComCosteo =" & aIDAnterior
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        cons = "Select * from CMCompra "
        Set rsIns = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
        Do While Not rsAux.EOF
            
            rsIns.AddNew
            rsIns!ComFecha = rsAux!ComFecha
            rsIns!ComArticulo = rsAux!ComArticulo
            rsIns!ComTipo = rsAux!ComTipo
            rsIns!ComCodigo = rsAux!ComCodigo
            rsIns!ComCantidad = rsAux!ComCantidad
            rsIns!ComCosto = rsAux!ComCosto
            rsIns!ComQOriginal = rsAux!ComQOriginal
            rsIns.Update
            
            bProgress.Value = bProgress.Value + 1
            rsAux.MoveNext
        Loop
        rsIns.Close
    End If
    rsAux.Close
    
    bProgress.Value = 0
    
    '---------------------------------------------------------------------------------------------------------------------------------------
    
    '5) Levanto respaldo Tabla Venta   -------------------------------------------------------------------------------------------
    cons = "Select Count(*) from rspCMVenta Where VenCosteo = " & aIDAnterior
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux(0)) Then
            bProgress.Max = rsAux(0)
            lAccion.Caption = "Restaurando " & rsAux(0) & " registros de Ventas Rebotadas"
            Me.Refresh
        End If
    End If
    rsAux.Close
    
    cons = "Select * from rspCMVenta Where VenCosteo = " & aIDAnterior
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        cons = "Select * from CMVenta "
        Set rsIns = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
        Do While Not rsAux.EOF
            
            rsIns.AddNew
            rsIns!VenFecha = rsAux!VenFecha
            rsIns!VenArticulo = rsAux!VenArticulo
            rsIns!VenTipo = rsAux!VenTipo
            rsIns!VenCodigo = rsAux!VenCodigo
            rsIns!VenCantidad = rsAux!VenCantidad
            rsIns!VenPrecio = rsAux!VenPrecio
            rsIns.Update
            
            bProgress.Value = bProgress.Value + 1
            rsAux.MoveNext
        Loop
        rsIns.Close
    End If
    rsAux.Close
    
    bProgress.Value = 0
    
    '4) Elimino Datos Tabla Cabezal
    lAccion.Caption = "Eliminando Cabezal " & Format(aMes, "Mmmm yyyy"): lAccion.Refresh
    cons = "Delete CMCabezal Where CabID = " & aIDCosteo
    cBase.Execute cons
    
    lAccion.Caption = ""
    MsgBox "Costeo Eliminado OK.", vbInformation, "Fin Eliminación"
    
    CargoInformacionCosteo
    
    If IsDate(tMes.Text) Then
        tSDe.Text = Format(PrimerDia(tMes.Text), "dd/mm/yyyy")
        tSHasta.Text = Format(UltimoDia(tMes.Text), "dd/mm/yyyy")
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBkUp:
    clsGeneral.OcurrioError "Error al borrar el costeo y recuperar datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function rspVerificoRespaldoUltimoCosteo() As Boolean
On Error GoTo errVerifico
Dim bHay As Boolean
Dim myIDCosteo As Long, myMes As Date

        
    rspUltimoCosteo myIDCosteo, myMes
    
    rspVerificoRespaldoUltimoCosteo = False
    
    'Veo si hay respaldo para el costeo vigente     --------------------------------------------------------
    bHay = False
    cons = "Select Top 1 * from rspCMCompra Where ComCosteo = " & myIDCosteo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True
    rsAux.Close
    
    If Not bHay Then
        cons = "Select Top 1 * from rspCMVenta Where VenCosteo = " & myIDCosteo
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then bHay = True
        rsAux.Close
    End If
    
    rspVerificoRespaldoUltimoCosteo = bHay
    Exit Function
    
errVerifico:
    clsGeneral.OcurrioError "Error al verificar el último respaldo.", Err.Description
End Function


Private Sub rspCopiarExistencia()
On Error GoTo errBkUp
Dim bHay As Boolean
Dim aIDCosteo As Long, aMes As Date

    Screen.MousePointer = 11
    
    rspUltimoCosteo aIDCosteo, aMes
    
    If aIDCosteo = -1 Then
        Screen.MousePointer = 0
        MsgBox "No hay Costeos para restaurar.", vbExclamation, "No Hay Datos"
        Exit Sub
    End If
    
    
    'Valido Si hay respaldo del costeo Anterior -------------------------------------------------
    bHay = False
    bHay = rspVerificoRespaldoUltimoCosteo
    
    If Not bHay Then
        MsgBox "No hay datos respaldados para el último costeo." & vbCrLf & _
                       "Faltan los respaldos de existencia y ventas rebotadas del costeo de " & Format(aMes, "Mmmm yyyy"), vbExclamation, "Faltan Respaldos"
        Screen.MousePointer = 0: Exit Sub
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '1) Elimino Datos Tabla Compra
    lAccion.Caption = "Eliminando Registros de Existencia ": lAccion.Refresh
    cons = "Delete CMCompra"
    cBase.Execute cons
    
    '2) Elimino Datos Tabla Venta
    lAccion.Caption = "Eliminando Registros de Rebotes ": lAccion.Refresh
    cons = "Delete CMVenta"
    cBase.Execute cons
    
    '5) Levanto respaldo Tabla Compra   -------------------------------------------------------------------------------------------
    Dim rsIns As rdoResultset
    
    bProgress.Value = 0
    
    cons = "Select Count(*) from rspCMCompra Where ComCosteo =" & aIDCosteo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux(0)) Then
            bProgress.Max = rsAux(0)
            lAccion.Caption = "Restaurando " & rsAux(0) & " registros de existencia"
            Me.Refresh
        End If
    End If
    rsAux.Close

    cons = "Select * from rspCMCompra Where ComCosteo =" & aIDCosteo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        cons = "Select * from CMCompra "
        Set rsIns = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
        Do While Not rsAux.EOF
            
            rsIns.AddNew
            rsIns!ComFecha = rsAux!ComFecha
            rsIns!ComArticulo = rsAux!ComArticulo
            rsIns!ComTipo = rsAux!ComTipo
            rsIns!ComCodigo = rsAux!ComCodigo
            rsIns!ComCantidad = rsAux!ComCantidad
            rsIns!ComCosto = rsAux!ComCosto
            rsIns!ComQOriginal = rsAux!ComQOriginal
            rsIns.Update
            
            bProgress.Value = bProgress.Value + 1
            rsAux.MoveNext
        Loop
        rsIns.Close
    End If
    rsAux.Close
    
    bProgress.Value = 0
    
    '---------------------------------------------------------------------------------------------------------------------------------------
    
    '5) Levanto respaldo Tabla Venta   -------------------------------------------------------------------------------------------
    cons = "Select Count(*) from rspCMVenta Where VenCosteo = " & aIDCosteo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux(0)) Then
            bProgress.Max = rsAux(0)
            lAccion.Caption = "Restaurando " & rsAux(0) & " registros de Ventas Rebotadas"
            Me.Refresh
        End If
    End If
    rsAux.Close
    
    cons = "Select * from rspCMVenta Where VenCosteo = " & aIDCosteo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        cons = "Select * from CMVenta "
        Set rsIns = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
        Do While Not rsAux.EOF
            
            rsIns.AddNew
            rsIns!VenFecha = rsAux!VenFecha
            rsIns!VenArticulo = rsAux!VenArticulo
            rsIns!VenTipo = rsAux!VenTipo
            rsIns!VenCodigo = rsAux!VenCodigo
            rsIns!VenCantidad = rsAux!VenCantidad
            rsIns!VenPrecio = rsAux!VenPrecio
            rsIns.Update
            
            bProgress.Value = bProgress.Value + 1
            rsAux.MoveNext
        Loop
        rsIns.Close
    End If
    rsAux.Close
    
    bProgress.Value = 0
    
    lAccion.Caption = ""
    MsgBox "Existencia y Ventas Rebotadas acutalizadas OK.", vbInformation, "Fin Actualización"
    
    CargoInformacionCosteo
    
    If IsDate(tMes.Text) Then
        tSDe.Text = Format(PrimerDia(tMes.Text), "dd/mm/yyyy")
        tSHasta.Text = Format(UltimoDia(tMes.Text), "dd/mm/yyyy")
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBkUp:
    clsGeneral.OcurrioError "Error al recuperar datos del respaldo.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub AccionMenuHelp()
    On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    cons = "Select * from Aplicacion Where AplNombre = '" & Trim(App.Title) & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux!AplHelp) Then aFile = Trim(rsAux!AplHelp)
    rsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

