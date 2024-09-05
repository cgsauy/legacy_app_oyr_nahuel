VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCumplir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cumplir Servicio"
   ClientHeight    =   4800
   ClientLeft      =   3030
   ClientTop       =   2640
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCumplir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8850
   Begin VB.Frame Frame2 
      Caption         =   "Datos Servicio"
      ForeColor       =   &H00800000&
      Height          =   2355
      Left            =   60
      TabIndex        =   39
      Top             =   60
      Width           =   8715
      Begin VB.CommandButton bPDireccion 
         BackColor       =   &H8000000E&
         Caption         =   "Dirección&..."
         Height          =   320
         Left            =   7620
         Picture         =   "frmCumplir.frx":0442
         TabIndex        =   12
         Top             =   1980
         Width           =   975
      End
      Begin VB.TextBox tPDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1995
         Width           =   5055
      End
      Begin VB.TextBox tPArticulo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1395
         Width           =   6075
      End
      Begin VB.TextBox tPNroMaquina 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         MaxLength       =   40
         TabIndex        =   10
         Top             =   1695
         Width           =   1875
      End
      Begin VB.TextBox tPFacturaS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4740
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "AA"
         Top             =   1695
         Width           =   315
      End
      Begin VB.TextBox tPFacturaN 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5100
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1695
         Width           =   675
      End
      Begin VB.TextBox tPFCompra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1695
         Width           =   975
      End
      Begin VB.TextBox tSCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   540
         Width           =   8475
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lSCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Casa 9242557; Celular 099405236"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   900
         TabIndex        =   51
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   7695
      End
      Begin VB.Label lSFecha 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4740
         TabIndex        =   50
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Solicitud:"
         Height          =   255
         Left            =   4020
         TabIndex        =   49
         Top             =   270
         Width           =   675
      End
      Begin VB.Label lSModificado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/09/00 23:55"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7260
         TabIndex        =   48
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Modificado:"
         Height          =   195
         Left            =   6360
         TabIndex        =   47
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lPIdProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   900
         TabIndex        =   46
         Top             =   1395
         Width           =   885
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1395
         Width           =   795
      End
      Begin VB.Label lPTipo 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   2100
         TabIndex        =   2
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Seri&e:"
         Height          =   255
         Left            =   6000
         TabIndex        =   9
         Top             =   1695
         Width           =   735
      End
      Begin VB.Label lDocumento 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº Factura: "
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   1695
         Width           =   915
      End
      Begin VB.Label Label13 
         Caption         =   "F/&Compra:"
         Height          =   255
         Left            =   1740
         TabIndex        =   4
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label lPGarantia 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   900
         TabIndex        =   44
         Top             =   1995
         Width           =   1425
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Garantía:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1995
         Width           =   735
      End
      Begin VB.Label lPEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   900
         TabIndex        =   42
         Top             =   1695
         Width           =   645
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1695
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Nº &Servicio:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lSProceso 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2220
         TabIndex        =   40
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cumplido"
      ForeColor       =   &H00800000&
      Height          =   1995
      Left            =   60
      TabIndex        =   38
      Top             =   2460
      Width           =   8715
      Begin VB.CheckBox chAnulado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "Servicio &Anulado"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2940
         MaskColor       =   &H00000000&
         TabIndex        =   15
         Top             =   240
         Width           =   1515
      End
      Begin AACombo99.AACombo cAsignado 
         Height          =   315
         Left            =   1020
         TabIndex        =   17
         Top             =   555
         Width           =   2175
         _ExtentX        =   3836
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
      Begin VB.TextBox tCantidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8100
         MaxLength       =   4
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox tFCumplido 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3660
         MaxLength       =   10
         TabIndex        =   31
         Top             =   1260
         Width           =   915
      End
      Begin VB.TextBox tComentarioR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         MaxLength       =   75
         TabIndex        =   33
         Top             =   1620
         Width           =   5835
      End
      Begin VB.TextBox tCostoV 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1740
         MaxLength       =   9
         TabIndex        =   20
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox tLiquidar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3645
         MaxLength       =   9
         TabIndex        =   22
         Top             =   900
         Width           =   795
      End
      Begin VB.TextBox tCostoR 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1740
         MaxLength       =   11
         TabIndex        =   29
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox tMotivo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4620
         TabIndex        =   23
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox tUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7620
         MaxLength       =   11
         TabIndex        =   35
         Top             =   1620
         Width           =   975
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsMotivo 
         Height          =   1035
         Left            =   4620
         TabIndex        =   26
         Top             =   540
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1826
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
         SelectionMode   =   1
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
         ExtendLastCol   =   0   'False
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
      Begin AACombo99.AACombo cSEstado 
         Height          =   315
         Left            =   1020
         TabIndex        =   14
         Top             =   220
         Width           =   795
         _ExtentX        =   1402
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
      Begin AACombo99.AACombo cMonedaR 
         Height          =   315
         Left            =   1020
         TabIndex        =   28
         Top             =   1260
         Width           =   675
         _ExtentX        =   1191
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
      Begin AACombo99.AACombo cMonedaV 
         Height          =   315
         Left            =   1020
         TabIndex        =   19
         Top             =   900
         Width           =   675
         _ExtentX        =   1191
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
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Lista"
         Height          =   255
         Left            =   4680
         TabIndex        =   25
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Asi&gnado:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Cu&mplido:"
         Height          =   195
         Left            =   2940
         TabIndex        =   30
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Aclaración:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1650
         Width           =   915
      End
      Begin VB.Label lVisita 
         BackStyle       =   0  'Transparent
         Caption         =   "Visita:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Li&quidar:"
         Height          =   255
         Left            =   2940
         TabIndex        =   21
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "&Reparación:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Esta&do:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario:"
         Height          =   195
         Left            =   6960
         TabIndex        =   34
         Top             =   1650
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   4545
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   15558
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCumplir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gServicio As Long
Dim gProducto As Long, gCliente As Long

Dim aTexto As String

Public Property Get prmServicio() As Long
    prmServicio = gServicio
End Property
Public Property Let prmServicio(Codigo As Long)
    gServicio = Codigo
End Property

Private Sub bPDireccion_Click()
Dim aDirAnterior As Long, aRetorno As Long
    
    On Error GoTo errDirecccion
    If Val(lPIdProducto.Caption) = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    aDirAnterior = Val(tPDireccion.Tag)
    
    Dim objDireccion As New clsDireccion
    objDireccion.ActivoFormularioDireccion cBase, aDirAnterior, gCliente, "Producto", "ProDireccion", "ProCodigo", gProducto
    Me.Refresh
    aRetorno = objDireccion.CodigoDeDireccion
    Set objDireccion = Nothing
    
    If aDirAnterior <> aRetorno Then
        If aRetorno <> 0 Then
            Cons = "Update Producto Set ProDireccion = " & aRetorno & " Where ProCodigo = " & gProducto
        Else
            Cons = "Update Producto Set ProDireccion = Null Where ProCodigo = " & gProducto
        End If
        cBase.Execute Cons
    End If
    
    If aRetorno <> 0 Then
        tPDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, aRetorno, True, True, True)
    Else
        tPDireccion.Text = ""
    End If
    tPDireccion.Tag = aRetorno
    
    Screen.MousePointer = 0
    Exit Sub
    
errDirecccion:
    clsGeneral.OcurrioError "Ocurrió un error al cargar la dirección.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cAsignado_GotFocus()
    Status.Panels(1).Text = "Técnico o camión que realiza el servicio."
End Sub

Private Sub cAsignado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tCostoV
End Sub

Private Sub chAnulado_Click()

    If chAnulado.Value = vbChecked Then
        cMonedaR.Enabled = False: cMonedaR.BackColor = Colores.Gris
        tCostoR.Enabled = False: tCostoR.BackColor = Colores.Gris
        tMotivo.Enabled = False: tMotivo.BackColor = Colores.Gris
        tCantidad.Enabled = False:: tMotivo.BackColor = Colores.Gris
        vsMotivo.Rows = 1
    Else
        cMonedaR.Enabled = True: cMonedaR.BackColor = Colores.Blanco
        tCostoR.Enabled = True: tCostoR.BackColor = Colores.Blanco
        tMotivo.Enabled = True: tMotivo.BackColor = Colores.Blanco
        tCantidad.Enabled = True:: tMotivo.BackColor = Colores.Blanco
    End If
    
End Sub

Private Sub chAnulado_GotFocus()
    Status.Panels(1).Text = "Indica si el servicio va a ser anulado."
End Sub

Private Sub chAnulado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cAsignado
End Sub

Private Sub cMonedaR_GotFocus()
    Status.Panels(1).Text = "Moneda que determina el costo de la reparación."
End Sub

Private Sub cMonedaV_GotFocus()
    Status.Panels(1).Text = "Moneda que determina el costo de la visita."
End Sub

Private Sub cSEstado_GotFocus()
    Status.Panels(1).Text = "Estado con el que se va a cumplir el servicio."
End Sub

Private Sub cSEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chAnulado.SetFocus
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    ObtengoSeteoForm Me
    CargoCombos
    LimpioFicha
    
    If gServicio <> 0 Then CargoDatosServicio gServicio
    
End Sub

Private Sub CargoDatosServicio(IdServicio As Long)
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    gProducto = 0: gCliente = 0
    
    Cons = "Select * from Servicio Where SerCodigo = " & IdServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        tSCodigo.Text = IdServicio
        tSCodigo.Tag = IdServicio
        gProducto = RsAux!SerProducto
        
        lSProceso.Caption = UCase(EstadoServicio(RsAux!SerEstadoServicio))
        lSProceso.Tag = RsAux!SerEstadoServicio
        BuscoCodigoEnCombo cSEstado, RsAux!SerEstadoProducto
                        
        lSFecha.Caption = Format(RsAux!SerFecha, "Ddd d/mm hh:mm"): lSFecha.Tag = RsAux!SerFecha
        lSModificado.Caption = Format(RsAux!SerModificacion, "dd/mm/yy hh:mm")
        
        If Not IsNull(RsAux!SerComentario) Then tComentario.Text = Trim(RsAux!SerComentario) & Chr(vbKeyReturn) & Chr(10)
                
        If Not IsNull(RsAux!SerCostoFinal) Then tCostoR.Text = Format(RsAux!SerCostoFinal, FormatoMonedaP)
        If Not IsNull(RsAux!SerMoneda) Then BuscoCodigoEnCombo cMonedaR, RsAux!SerMoneda Else BuscoCodigoEnCombo cMonedaR, CLng(paMonedaPesos)
        
        If Not IsNull(RsAux!SerFCumplido) Then tFCumplido.Text = Format(RsAux!SerFCumplido, "d/mm/yyyy")
        If Not IsNull(RsAux!SerComentarioR) Then tComentarioR.Text = Trim(RsAux!SerComentarioR)
        
        cSEstado.Tag = 0: tComentario.Tag = 0
        
    Else
        MsgBox "No hay un servicio pendiente con el código: " & IdServicio, vbInformation, "ATENCIÓN"
        IdServicio = 0
    End If
    RsAux.Close
    
    If IdServicio <> 0 Then 'Cargo los motivos de llamado
        aTexto = ""
        Cons = "Select * from ServicioRenglon, MotivoServicio" & _
                   " Where SReServicio = " & IdServicio & _
                   " And SReMotivo = MSeID" & _
                   " And SReTipoRenglon = " & TipoRenglonS.Llamado
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            aTexto = aTexto & Trim(RsAux!MSeNombre) & ", "
            RsAux.MoveNext
        Loop
        RsAux.Close
        If Len(aTexto) > 2 Then aTexto = Mid(aTexto, 1, Len(aTexto) - 2)
        tComentario.Text = Trim(tComentario.Text) & aTexto
    End If
    
    If gProducto <> 0 Then CargoDatosProducto gProducto
    If IdServicio <> 0 Then
        If Val(lSProceso.Tag) <> EstadoS.Anulado And Val(lSProceso.Tag) <> EstadoS.Cumplido And Val(lSProceso.Tag) <> EstadoS.Taller Then
            EstadoControles True, Colores.Blanco
        End If
        CargoDatosVisitas IdServicio
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el servicio.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosProducto(idProducto As Long)
    
    On Error GoTo ErrCE
    Screen.MousePointer = 11
    
    Cons = "Select * from Producto, Articulo " _
            & " Where ProCodigo = " & idProducto _
            & " And ProArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        'gCliente = RsAux!ProCliente
        bPDireccion.Enabled = True
        
        lPIdProducto.Caption = " " & Format(idProducto, "000")
        tPArticulo.Text = Trim(RsAux!ArtNombre)
        tPArticulo.Tag = RsAux!ArtId
        lPTipo.Tag = RsAux!ArtTipo      'Tipo del Articulo para ingreso de motivos
        
        If Not IsNull(RsAux!ProCompra) Then tPFCompra.Text = Format(RsAux!ProCompra, "dd/mm/yyyy")
        If Not IsNull(RsAux!ProFacturaS) Then tPFacturaS.Text = Trim(RsAux!ProFacturaS)
        If Not IsNull(RsAux!ProFacturaN) Then tPFacturaN.Text = RsAux!ProFacturaN
        If Not IsNull(RsAux!ProNroSerie) Then tPNroMaquina.Text = Trim(RsAux!ProNroSerie)
        If Not IsNull(RsAux!ProDireccion) Then
            tPDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!ProDireccion, True, True, True)
            tPDireccion.Tag = RsAux!ProDireccion
        End If
        
        lPGarantia.Caption = " " & RetornoGarantia(RsAux!ArtId)
        lPEstado.Tag = CalculoEstadoProducto(RsAux!ProCodigo)
        lPEstado.Caption = " " & EstadoProducto(Val(lPEstado.Tag))
        
        tPFCompra.Tag = 0: tPFacturaS.Tag = 0: tPFacturaN.Tag = 0: tPNroMaquina.Tag = 0
        
        If Not IsNull(RsAux!ProDocumento) Then lDocumento.Tag = RsAux!ProDocumento
    End If
    RsAux.Close
    
    If gCliente <> 0 Then
        Cons = "Select Cliente.*, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
               & " From Cliente, CPersona " _
               & " Where CliCodigo = " & gCliente _
               & " And CliCodigo = CPeCliente " _
                                                    & " UNION All" _
               & " Select Cliente.*, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
               & " From Cliente, CEmpresa " _
               & " Where CliCodigo = " & gCliente _
               & " And CliCodigo = CEmCliente"
    
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux!CliCIRuc) Then
                If RsAux!CliTipo = TipoCliente.Cliente Then lSCliente.Caption = "  (" & clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc) & ")"
                If RsAux!CliTipo = TipoCliente.Empresa Then lSCliente.Caption = "  (" & clsGeneral.RetornoFormatoRuc(RsAux!CliCIRuc) & ")"
            End If
            lSCliente.Caption = " " & Trim(RsAux!Nombre) & lSCliente.Caption
            lSCliente.Tag = RsAux!CliTipo
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub

ErrCE:
    clsGeneral.OcurrioError "Ocurrió un error al cargar la información del producto.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosVisitas(IdServicio As Long)
    
    'Teoricamente Pueden venir Visitas, Retiros y Entregas ---> No vamos a permitir Taller !!!
    
    Cons = "Select * from ServicioVisita Where VisServicio = " & IdServicio & _
               " and VisSinEfecto = 0 Order by VisCodigo Desc"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Select Case RsAux!VisTipo
            Case TipoServicio.Visita: lVisita.Caption = "&Visita"
            Case TipoServicio.Retiro: lVisita.Caption = "&Retiro"
            Case TipoServicio.Entrega: lVisita.Caption = "&Entrega"
        End Select
        lVisita.Tag = RsAux!VisCodigo       'Id, para dps grabar !!!!
        
        If tFCumplido.Enabled Then tFCumplido.Text = Format(RsAux!VisFecha, "d/mm/yyyy")
        
        BuscoCodigoEnCombo cAsignado, RsAux!VisCamion
        BuscoCodigoEnCombo cMonedaV, RsAux!VisMoneda
        tCostoV.Text = Format(RsAux!VisCosto, FormatoMonedaP)
        
        If cMonedaR.ListIndex = -1 Then BuscoCodigoEnCombo cMonedaR, RsAux!VisMoneda
        If Not IsNull(RsAux!VisLiquidarAlCamion) Then tLiquidar.Text = Format(RsAux!VisLiquidarAlCamion, FormatoMonedaP) Else tLiquidar.Text = "0.00"
    
        'Hay que definir se se puede cumplir o Anular....-----------------------------------------------------------------
        If Val(lSProceso.Tag) <> EstadoS.Anulado And Val(lSProceso.Tag) <> EstadoS.Cumplido And Val(lSProceso.Tag) <> EstadoS.Taller Then
            Select Case RsAux!VisTipo
                Case TipoServicio.Visita:
                    If IsNull(RsAux!VisFImpresion) Then
                        MsgBox "La visita no está impresa, si ud. la cumple va a quedar anulada.", vbExclamation, "Visita sin Imprimir."
                        chAnulado.Value = vbChecked: chAnulado.Enabled = False
                    End If
                Case TipoServicio.Retiro: chAnulado.Value = vbChecked: chAnulado.Enabled = False
                Case TipoServicio.Entrega:
                    If IsNull(RsAux!VisFImpresion) Then
                        MsgBox "La entrega no está impresa, si ud. la cumple, el servicio va a quedar anulado.", vbExclamation, "Entrega sin Imprimir."
                        chAnulado.Value = vbChecked: chAnulado.Enabled = False
                    Else
                         chAnulado.Enabled = False: tMotivo.Enabled = False: tMotivo.BackColor = Colores.Inactivo
                    End If
                    CargoReparacionTaller
                    'Deshabilito los datos del retiro
                    cMonedaR.Enabled = False: cMonedaR.BackColor = Colores.Gris
                    tCostoR.Enabled = False: tCostoR.BackColor = Colores.Gris
                    cSEstado.Enabled = False: cSEstado.BackColor = Colores.Gris

            End Select
        Else
            'Esta Cumplido,  Anulado o en Taller
            CargoReparacionTaller
        End If
        '----------------------------------------------------------------------------------------------------------------------------------
    Else
            'No hay datos de Visitas, puede ser que haya entrado directo a Taller
            CargoReparacionTaller
    End If
    RsAux.Close
    
End Sub

Private Sub CargoReparacionTaller()
    
    On Error GoTo errTaller
    Dim rsTal As rdoResultset
    
    If Val(lSProceso.Tag) = EstadoS.Taller Then
        EstadoControles True, Colores.Blanco
        chAnulado.Value = vbChecked: chAnulado.Enabled = False
        
        'Deshabilito los datos del retiro
        cMonedaR.Enabled = False: cMonedaR.BackColor = Colores.Gris
        tCostoR.Enabled = False: tCostoR.BackColor = Colores.Gris
        cSEstado.Enabled = False: cSEstado.BackColor = Colores.Gris
        cAsignado.Enabled = False: cAsignado.BackColor = Colores.Gris
        tLiquidar.Enabled = False: tLiquidar.BackColor = Colores.Gris
        tCostoV.Enabled = False: tCostoV.BackColor = Colores.Gris
        cMonedaV.Enabled = False: cMonedaV.BackColor = Colores.Gris
        tFCumplido.Text = Format(gFechaServidor, "d/mm/yyyy")
    End If

    'Cargo los renglones de reparación
    vsMotivo.Rows = 1
    Cons = "Select * From ServicioRenglon, Articulo" & _
                " Where SReServicio = " & Val(tSCodigo.Text) & _
                " And SReTipoRenglon = " & TipoRenglonS.Cumplido & _
                " And SReMotivo = ArtID"
    Set rsTal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not rsTal.EOF
        With vsMotivo
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(rsTal!ArtCodigo, "(#,000,000)") & " " & Trim(rsTal!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 1) = rsTal!SReCantidad
        End With
        rsTal.MoveNext
    Loop
    rsTal.Close
    
   
    Exit Sub
    
errTaller:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de taller.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Panels(1).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
End Sub

Private Sub lDocumento_Click()
    Foco tPFacturaS
End Sub

Private Sub Label11_Click()
    Foco tUsuario
End Sub

Private Sub Label12_Click()
    Foco tComentarioR
End Sub

Private Sub Label13_Click()
    Foco tPFCompra
End Sub

Private Sub Label14_Click()
    Foco cSEstado
End Sub

Private Sub Label15_Click()
    Foco tCostoR
End Sub

Private Sub Label16_Click()
    Foco cAsignado
End Sub

Private Sub Label17_Click()
    Foco tCostoR
End Sub

Private Sub Label2_Click()
    Foco tPNroMaquina
End Sub

Private Sub Label26_Click()
    Foco tLiquidar
End Sub

Private Sub Label3_Click()
    Foco tSCodigo
End Sub

Private Sub Label8_Click()
    If vsMotivo.Enabled Then vsMotivo.SetFocus
End Sub

Private Sub lPTipo_Click()
    tPArticulo.SetFocus
End Sub

Private Sub lVisita_Click()
    Foco tCostoV
End Sub


Private Sub AccionGrabar()

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma cumplir el servicio y almacenar la información ingresada", vbQuestion + vbYesNo, "CUMPLIR SERVICIO") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    
    On Error GoTo ErrBegin
    FechaDelServidor
    
    cBase.BeginTrans            'Comienzo la transaccion------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo ErrResumo
    CargoCamposBD
    cBase.CommitTrans         'Finalizo la transaccion------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    LimpioFicha
    Foco tSCodigo
    Screen.MousePointer = 0
    Exit Sub
    
ErrBegin:
    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0: Exit Sub
ErrResumo:
    Resume ErrCommit
ErrCommit:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al intentar realizar la transacción.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(tCantidad.Text) Then
            MsgBox "La cantidad ingresada no es correcta. Verifique.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        If Val(tMotivo.Tag) = 0 Then
            MsgBox "Debe seleccionar el artículo a agregar en la lista.", vbExclamation, "ATENCIÓN"
            Foco tMotivo: Exit Sub
        End If
        
        AgregoMotivo idMotivo:=CLng(tMotivo.Tag), Articulo:=True
        tMotivo.Text = "": tCantidad.Text = "": Foco tMotivo
    End If
    
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub


Private Sub tComentarioR_GotFocus()
    Status.Panels(1).Text = "Aclaración del cumplido."
End Sub

Private Sub tComentarioR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub tCostoR_GotFocus()
    Status.Panels(1).Text = "Costo de la reparación."
End Sub

Private Sub tCostoR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFCumplido
End Sub

Private Sub tCostoR_LostFocus()
    If IsNumeric(tCostoR.Text) Then tCostoR.Text = Format(tCostoR.Text, FormatoMonedaP)
End Sub

Private Sub tCostoV_GotFocus()
    Status.Panels(1).Text = "Costo de la visita."
End Sub

Private Sub tCostoV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tLiquidar
End Sub

Private Sub tCostoV_LostFocus()
    If IsNumeric(tCostoV.Text) Then tCostoV.Text = Format(tCostoV.Text, FormatoMonedaP)
End Sub

Private Sub tFCumplido_GotFocus()
    Status.Panels(1).Text = "Fecha en que se realizó el servicio."
End Sub

Private Sub tFCumplido_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentarioR
End Sub

Private Sub tFCumplido_LostFocus()
    If IsDate(tFCumplido.Text) Then tFCumplido.Text = Format(tFCumplido.Text, "d/mm/yyyy")
End Sub

Private Sub tLiquidar_GotFocus()
    Status.Panels(1).Text = "Importe de la visita a liquidar al técnico/camión."
End Sub

Private Sub tLiquidar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If (chAnulado.Value = vbUnchecked And chAnulado.Enabled) Then Foco tMotivo Else Foco tFCumplido
End Sub

Private Sub tLiquidar_LostFocus()
    If IsNumeric(tLiquidar.Text) Then tLiquidar.Text = Format(tLiquidar.Text, FormatoMonedaP)
End Sub

Private Sub tMotivo_Change()
    tMotivo.Tag = 0
End Sub

Private Sub tMotivo_GotFocus()
    If tCantidad.Visible Then
        Status.Panels(1).Text = "Ingrese el artículo para cumplir el servicio. [F12]- Presupuestos"
    Else
        Status.Panels(1).Text = "Ingrese el presupuesto para cumplir el servicio. [F12]- Artículos"
    End If
End Sub

Private Sub tMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF12
            If tCantidad.Visible Then
                tCantidad.Visible = False: tMotivo.Width = vsMotivo.Width
            Else
                tCantidad.Visible = True: tMotivo.Width = 3435
            End If
            Call tMotivo_GotFocus
            
    End Select
    
End Sub

Private Sub tMotivo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        On Error GoTo errAgregar
        If Trim(tMotivo.Text) = "" Then Foco tCostoR: Exit Sub
        
        Dim aIdMotivo As Long, aMotivo As String
        aIdMotivo = 0
        
        If tCantidad.Visible = False Then           'ES PRESUPUESTO
            Screen.MousePointer = 11
            
            If IsNumeric(tMotivo.Text) Then
                Cons = "Select * from Presupuesto" & _
                          " Where PreCodigo = " & Trim(tMotivo.Text) & _
                          " And PreEsPresupuesto = 1"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    aIdMotivo = RsAux!PreID
                Else
                    MsgBox "No existe un presupuesto para el código ingresado " & Format(tMotivo.Text, "(#,000)"), vbInformation, "Motivo inexistente"
                End If
                RsAux.Close
            
            Else
                Cons = "Select PreID, 'Presupuesto' = PreNombre , 'Código' = PreCodigo from Presupuesto" & _
                          " Where PreNombre like '" & Trim(tMotivo.Text) & "%'" & _
                          " And PreEsPresupuesto = 1" & _
                          " Order by PreNombre"
                Dim objLista As New clsListadeAyuda
                If objLista.ActivarAyuda(cBase, Cons, 5000, 1) > 0 Then
                    aIdMotivo = objLista.RetornoDatoSeleccionado(0)
                End If
                Set objLista = Nothing
                Me.Refresh
            End If
            
            If aIdMotivo <> 0 Then AgregoMotivo idMotivo:=aIdMotivo, Presupuesto:=True: tMotivo.Text = ""
            Screen.MousePointer = 0
        
        Else        'Es articulo
            Screen.MousePointer = 11
            
            If IsNumeric(tMotivo.Text) Then
                Cons = "Select * from Articulo " & _
                          " Where ArtCodigo = " & Trim(tMotivo.Text)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    tMotivo.Text = Trim(RsAux!ArtNombre)
                    aIdMotivo = RsAux!ArtId
                Else
                    MsgBox "No existe un artículo para el código ingresado " & Format(tMotivo.Text, "(#,000)"), vbInformation, "Artículo inexistente"
                End If
                RsAux.Close
            
            Else
                Cons = "Select ArtID, 'Artículo' = ArtNombre , 'Código' = ArtCodigo from Articulo" & _
                          " Where ArtNombre like '" & Trim(tMotivo.Text) & "%'" & _
                          " Order by ArtNombre"
                Dim objListaA As New clsListadeAyuda
                If objListaA.ActivarAyuda(cBase, Cons, 5000, 1) > 0 Then
                    tMotivo.Text = objListaA.RetornoDatoSeleccionado(1)
                    aIdMotivo = objListaA.RetornoDatoSeleccionado(0)
                End If
                Set objListaA = Nothing
                Me.Refresh
            End If
            
            If aIdMotivo <> 0 Then tMotivo.Tag = aIdMotivo: tCantidad.Text = "1": Foco tCantidad
            Screen.MousePointer = 0
        End If
    End If
    Exit Sub

errAgregar:
    clsGeneral.OcurrioError "Ocurrió un error al procesar el presupuesto/artículo", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AgregoMotivo(idMotivo As Long, Optional Presupuesto As Boolean = False, Optional Articulo As Boolean = False)
    On Error GoTo errAgregar
    
    Screen.MousePointer = 11
    Dim aValor As Long, I As Integer
        
    If Presupuesto Then
        Cons = "Select * from PresupuestoArticulo, Articulo " & _
                                 "Left Outer Join PrecioVigente On ArtID = PViArticulo " & _
                                                                          " And PViTipoCuota = " & paTipoCuotaContado & _
                                                                          " And PViMoneda = " & paMonedaPesos & _
                    " Where PArPresupuesto = " & idMotivo & _
                    " And PArArticulo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            With vsMotivo
                If Not ArticuloIngresado(RsAux!ArtId) Then
                    .AddItem Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                    aValor = RsAux!ArtId: .Cell(flexcpData, .Rows - 1, 0) = aValor
                    
                    .Cell(flexcpText, .Rows - 1, 1) = RsAux!PArCantidad
                    If Not IsNull(RsAux!PViPrecio) Then .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PViPrecio, FormatoMonedaP) Else .Cell(flexcpText, .Rows - 1, 2) = "0.00"
                    If Not IsNumeric(tCostoR.Text) Then tCostoR.Text = .Cell(flexcpText, .Rows - 1, 2) Else tCostoR.Text = Format(CCur(tCostoR.Text) + .Cell(flexcpValue, .Rows - 1, 2), FormatoMonedaP)
                End If
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        'Agrego el Articulo del Preosupuesto (Bonificacion)
        Cons = "Select * from Presupuesto, Articulo" & _
                   " Where PreID = " & idMotivo & _
                   " And PreArticulo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            With vsMotivo
                If Not ArticuloIngresado(RsAux!ArtId) Then
                    .AddItem Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                    aValor = RsAux!ArtId: .Cell(flexcpData, .Rows - 1, 0) = aValor
                    
                    .Cell(flexcpText, .Rows - 1, 1) = "1"
                    .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PreImporte, FormatoMonedaP)
                    If Not IsNumeric(tCostoR.Text) Then tCostoR.Text = .Cell(flexcpText, .Rows - 1, 2) Else tCostoR.Text = Format(CCur(tCostoR.Text) + .Cell(flexcpValue, .Rows - 1, 2), FormatoMonedaP)
                End If
            End With
        End If
        RsAux.Close
    End If
    
    If Articulo Then
        Cons = "Select * from Articulo " & _
                                 "Left Outer Join PrecioVigente On ArtID = PViArticulo " & _
                                                                          " And PViTipoCuota = " & paTipoCuotaContado & _
                                                                          " And PViMoneda = " & paMonedaPesos & _
                    " Where ArtId = " & idMotivo
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            With vsMotivo
                If Not ArticuloIngresado(RsAux!ArtId) Then
                    .AddItem Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                    aValor = RsAux!ArtId: .Cell(flexcpData, .Rows - 1, 0) = aValor
                    
                    .Cell(flexcpText, .Rows - 1, 1) = tCantidad.Text
                    If Not IsNull(RsAux!PViPrecio) Then .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PViPrecio, FormatoMonedaP) Else .Cell(flexcpText, .Rows - 1, 2) = "0.00"
                    If Not IsNumeric(tCostoR.Text) Then tCostoR.Text = .Cell(flexcpText, .Rows - 1, 2) Else tCostoR.Text = Format(CCur(tCostoR.Text) + .Cell(flexcpValue, .Rows - 1, 2), FormatoMonedaP)
                End If
            End With
        End If
        RsAux.Close
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errAgregar:
    clsGeneral.OcurrioError "Ocurrió un error al agregar el item a la lista.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tMotivo_LostFocus()
    Status.Panels(1).Text = ""
End Sub


Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    If Val(lSProceso.Tag) <> EstadoS.Taller Then
        If cSEstado.ListIndex = -1 Then
            MsgBox "Debe seleccionar el estado para cumplir el servicio.", vbExclamation, "ATENCIÓN"
            Foco cSEstado: Exit Function
        End If
        If cAsignado.ListIndex = -1 Then
            MsgBox "Debe seleccionar quien realizó el servicio.", vbExclamation, "ATENCIÓN"
            Foco cAsignado: Exit Function
        End If
        
        If cMonedaV.ListIndex = -1 Then
            MsgBox "Debe seleccionar la moneda para registrar el valor de la visita.", vbExclamation, "ATENCIÓN"
            Foco cMonedaV: Exit Function
        End If
        If Trim(tCostoV.Text) = "" Or Not IsNumeric(tCostoV.Text) Then
            MsgBox "Debe ingresar el valor de la visita.", vbExclamation, "ATENCIÓN"
            Foco tCostoV: Exit Function
        End If
        If Trim(tLiquidar.Text) = "" Or Not IsNumeric(tLiquidar.Text) Then
            MsgBox "Debe ingresar el importe a liquidar al técnico/camión que realizó el servicio.", vbExclamation, "ATENCIÓN"
            Foco tLiquidar: Exit Function
        End If
        
        If chAnulado.Value = vbUnchecked Then
            If cMonedaR.ListIndex = -1 Then
                MsgBox "Debe seleccionar la moneda para registrar el valor de la reparación.", vbExclamation, "ATENCIÓN"
                Foco cMonedaR: Exit Function
            End If
            If Trim(tCostoR.Text) = "" Or Not IsNumeric(tCostoR.Text) Then
                MsgBox "Debe ingresar el valor de la reparación.", vbExclamation, "ATENCIÓN"
                Foco tCostoR: Exit Function
            End If
        End If
    End If
    
    If Not IsDate(tFCumplido.Text) Then
        MsgBox "La fecha de cumplido del servicio no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tFCumplido: Exit Function
    Else
        If CDate(tFCumplido.Text & " 23:59:59") < CDate(lSFecha.Tag) Then
            MsgBox "La fecha de cumplido no puede ser menor a la fecha de solicitud del servicio.", vbExclamation, "ATENCIÓN"
            Foco tFCumplido: Exit Function
        End If
    End If
    
    If Not clsGeneral.TextoValido(tComentarioR.Text) Then
        MsgBox "Hay carácteres no válidos en el texto aclaración. Verifique.", vbExclamation, "ATENCIÓN"
        Foco tComentarioR: Exit Function
    End If
    
    If Val(tUsuario.Tag) = 0 Then
        MsgBox "Para grabar debe ingresar el dígito de usuario.", vbExclamation, "ATENCIÓN"
        Foco tUsuario: Exit Function
    End If
    
    If chAnulado.Value = vbChecked And Trim(tComentarioR.Text) = "" Then
        MsgBox "Ingrese un cometario para registrar el motivo de la anulación.", vbExclamation, "Comentario de Anulación"
        Foco tComentarioR: Exit Function
    End If
    
    ValidoCampos = True
    
End Function

Private Sub EstadoControles(Estado As Boolean, ColorFondo As Long)
    
    tPArticulo.Enabled = Estado: tPArticulo.BackColor = ColorFondo
    tPFCompra.Enabled = Estado: tPFCompra.BackColor = ColorFondo
    tPFacturaS.Enabled = Estado: tPFacturaS.BackColor = ColorFondo
    tPFacturaN.Enabled = Estado: tPFacturaN.BackColor = ColorFondo
    tPNroMaquina.Enabled = Estado: tPNroMaquina.BackColor = ColorFondo
    tPDireccion.BackColor = ColorFondo
    
    If Estado = True And Val(lDocumento.Tag) <> 0 Then
        tPArticulo.Enabled = False: tPArticulo.BackColor = Colores.Inactivo
        tPFCompra.Enabled = False: tPFCompra.BackColor = Colores.Inactivo
        tPFacturaS.Enabled = False: tPFacturaS.BackColor = Colores.Inactivo
        tPFacturaN.Enabled = False: tPFacturaN.BackColor = Colores.Inactivo
    End If
    
    chAnulado.Enabled = Estado
    cSEstado.Enabled = Estado: cSEstado.BackColor = ColorFondo
    cAsignado.Enabled = Estado: cAsignado.BackColor = ColorFondo
    cMonedaV.Enabled = Estado: cMonedaV.BackColor = ColorFondo
    tCostoV.Enabled = Estado: tCostoV.BackColor = ColorFondo
    tLiquidar.Enabled = Estado: tLiquidar.BackColor = ColorFondo
    
    cMonedaR.Enabled = Estado: cMonedaR.BackColor = ColorFondo
    tCostoR.Enabled = Estado: tCostoR.BackColor = ColorFondo
    tFCumplido.Enabled = Estado: tFCumplido.BackColor = ColorFondo
    
    tMotivo.Enabled = Estado: tMotivo.BackColor = ColorFondo
    tCantidad.Enabled = Estado: tCantidad.BackColor = ColorFondo
    vsMotivo.Enabled = Estado: vsMotivo.BackColor = ColorFondo
    
    tComentarioR.Enabled = Estado: tComentarioR.BackColor = ColorFondo
    tUsuario.Enabled = Estado: tUsuario.BackColor = ColorFondo
    
End Sub

Private Sub CargoCamposBD()
    
    'Tabla Servicio-----------------------------------------------------------------------------------------------------
    Cons = "Select * from Servicio Where SerCodigo = " & gServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    
    RsAux!SerEstadoProducto = cSEstado.ItemData(cSEstado.ListIndex)
    If chAnulado.Value = vbUnchecked Then
        RsAux!SerMoneda = cMonedaR.ItemData(cMonedaR.ListIndex)
        RsAux!SerCostoFinal = CCur(tCostoR.Text)
    End If
    
    If Trim(tComentarioR.Text) <> "" Then RsAux!SerComentarioR = Trim(tComentarioR.Text) Else RsAux!SerComentarioR = Null
    RsAux!SerFCumplido = Format(tFCumplido.Text, sqlFormatoF)
    
    RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsAux!SerUsuario = Val(tUsuario.Tag)
    If chAnulado.Value = vbUnchecked Then RsAux!SerEstadoServicio = EstadoS.Cumplido Else RsAux!SerEstadoServicio = EstadoS.Anulado
    
    RsAux.Update: RsAux.Close
    
    If Val(lSProceso.Tag) <> EstadoS.Taller Then
        'Tabla Visita-----------------------------------------------------------------------------------------------------
        Cons = "Select * from ServicioVisita Where VisSinEfecto = 0 AND VisCodigo = " & Val(lVisita.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If (Not RsAux.EOF) Then
        RsAux.Edit
        
        RsAux!VisMoneda = cMonedaV.ItemData(cMonedaV.ListIndex)
        RsAux!VisCosto = CCur(tCostoV.Text)
        RsAux!VisLiquidarAlCamion = CCur(tLiquidar.Text)
        RsAux!VisFModificacion = Format(gFechaServidor, sqlFormatoFH)
        
        RsAux.Update: RsAux.Close
        End If
    End If
    
    'Tabla ServicioRenglon------------------------------------------------------------------------------------------------
    'Esto lo grabo si es el cumplido de una visita
    If Val(lSProceso.Tag) = EstadoS.Visita Then
        With vsMotivo
        If .Rows > 1 Then
            Cons = "Select * from ServicioRenglon Where SReServicio = " & gServicio & " And SReTipoRenglon = " & TipoRenglonS.Cumplido
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            For I = 1 To .Rows - 1
                RsAux.AddNew
                RsAux!SReServicio = gServicio
                RsAux!SReTipoRenglon = TipoRenglonS.Cumplido
                RsAux!SReMotivo = .Cell(flexcpData, I, 0)
                RsAux!SReCantidad = .Cell(flexcpValue, I, 1)
                RsAux!SReTotal = .Cell(flexcpValue, I, 2)
                RsAux.Update
            Next
            RsAux.Close
        End If
        End With
    End If
    
End Sub

Private Sub LimpioFicha()

    On Error Resume Next
    tSCodigo.Tag = 0
    lSProceso.Caption = "": lSFecha.Caption = "": lSModificado.Caption = ""
    tComentario.Text = ""
    lSCliente.Caption = ""
    
    
    lPIdProducto.Caption = "": tPArticulo.Text = ""
    lPEstado.Caption = "": tPFCompra.Text = "": tPFacturaS.Text = "": tPFacturaN.Text = "": tPNroMaquina.Text = ""
    lPGarantia.Caption = "": tPDireccion.Text = "": bPDireccion.Enabled = False: bPDireccion.Tag = 0
    tPFCompra.Tag = 0: tPFacturaS.Tag = 0: tPFacturaN.Tag = 0: tPNroMaquina.Tag = 0
    
    cSEstado.Text = ""
    chAnulado.Value = vbUnchecked
    cAsignado.Text = ""
    cMonedaV.Text = "": tCostoV.Text = "": tLiquidar.Text = ""
    cMonedaR.Text = "": tCostoR.Text = "": tFCumplido.Text = ""
    tComentarioR.Text = "": tUsuario.Text = ""
    
    
    vsMotivo.Rows = 1: tMotivo.Text = "": tCantidad.Text = ""
    EstadoControles False, Colores.Gris
     
     
End Sub

Private Sub CargoCombos()
        
    On Error Resume Next
    With vsMotivo
        .Rows = 1: .Cols = 1
        .FormatString = "<Reparaciones / Repuestos|>Q|>Total"
        .ColWidth(0) = 2490: .ColWidth(1) = 400
        .WordWrap = False: .ExtendLastCol = True
    End With

    
    Cons = "Select * from Moneda Where MonFactura = 1 Order by MonSigno"
    CargoCombo Cons, cMonedaV
    CargoCombo Cons, cMonedaR
    
    cSEstado.AddItem EstadoProducto(EstadoP.Abonado): cSEstado.ItemData(cSEstado.NewIndex) = EstadoP.Abonado
    cSEstado.AddItem EstadoProducto(EstadoP.FueraGarantia): cSEstado.ItemData(cSEstado.NewIndex) = EstadoP.FueraGarantia
    cSEstado.AddItem EstadoProducto(EstadoP.SinCargo): cSEstado.ItemData(cSEstado.NewIndex) = EstadoP.SinCargo
    
    Cons = "Select * from Camion order by CamNombre"
    CargoCombo Cons, cAsignado
    
End Sub

Private Sub tPArticulo_GotFocus()
    Status.Panels(1).Text = "Presione [F1] para cambiar el tipo del producto."
End Sub

Private Sub tPArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        If Val(tSCodigo.Tag) = 0 Then Exit Sub
        On Error GoTo errLista
        Screen.MousePointer = 11
        Cons = "Select ArtId, ArtNombre 'Artículo', ArtCodigo 'Código' from Articulo Where ArtTipo = " & Val(lPTipo.Tag) & " Order by ArtNombre"
        Dim miLista As New clsListadeAyuda, aIDSel As Long, aTipoSel As String
        If miLista.ActivarAyuda(cBase, Cons, 5000, 1) > 0 Then
            aIDSel = miLista.RetornoDatoSeleccionado(0)
            aTipoSel = miLista.RetornoDatoSeleccionado(1)
        End If
        Set miLista = Nothing
        If aIDSel <> 0 Then
            tPArticulo.Text = aTipoSel
            tPArticulo.Tag = aIDSel
            ZActualizoCampoProducto CLng(lPIdProducto.Caption), aIDSel, Articulo:=True
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub

errLista:
    clsGeneral.OcurrioError "Ocurrió un error al activar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tPArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tPFCompra
End Sub

Private Sub tPDireccion_GotFocus()
    Status.Panels(1).Text = "Dirección del producto (para realizar servicios)."
End Sub

Private Sub tPFacturaN_Change()
    tPFacturaN.Tag = 1
End Sub

Private Sub tPFacturaN_GotFocus()
    With tPFacturaN: .SelStart = 0: .SelLength = Len(.Text): End With
    Status.Panels(1).Text = "Número de la factura de compra."
End Sub

Private Sub tPFacturaN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tPFacturaN.Tag) <> 0 And Trim(lPIdProducto.Caption) <> "" Then
            ZActualizoCampoProducto CLng(lPIdProducto.Caption), Trim(tPFacturaN.Text), FacturaN:=True
            tPFacturaN.Tag = 0
        End If
        Foco tPNroMaquina
    End If
End Sub

Private Sub tPFacturaS_Change()
    tPFacturaS.Tag = 1
End Sub

Private Sub tPFacturaS_GotFocus()
    With tPFacturaS: .SelStart = 0: .SelLength = Len(.Text): End With
    Status.Panels(1).Text = "Número de serie de la factura de compra."
End Sub

Private Sub tPFacturaS_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        If Val(tPFacturaS.Tag) <> 0 And Trim(lPIdProducto.Caption) <> "" Then
            ZActualizoCampoProducto CLng(lPIdProducto.Caption), Trim(tPFacturaS.Text), FacturaS:=True
            tPFacturaS.Tag = 0
        End If
        Foco tPFacturaN
    End If
End Sub

Private Sub tPFCompra_Change()
    tPFCompra.Tag = 1
End Sub

Private Sub tPFCompra_GotFocus()
    With tPFCompra: .SelStart = 0: .SelLength = Len(.Text): End With
    Status.Panels(1).Text = "Fecha de compra del producto."
End Sub

Private Sub tPFCompra_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Val(tPFCompra.Tag) <> 0 And Trim(lPIdProducto.Caption) <> "" Then
            If Not IsDate(tPFCompra.Text) Then MsgBox "La fecha ingresada no es correcta. Verifique", vbExclamation, "ATENCIÓN": Exit Sub
            ZActualizoCampoProducto CLng(lPIdProducto.Caption), Trim(tPFCompra.Text), FCompra:=True
            tPFCompra.Text = Format(tPFCompra.Text, "dd/mm/yyyy")
            tPFCompra.Tag = 0
        End If
        Foco tPFacturaS
    End If
    
End Sub

Private Sub tPNroMaquina_Change()
    tPNroMaquina.Tag = 1
End Sub

Private Sub tPNroMaquina_GotFocus()
    With tPNroMaquina: .SelStart = 0: .SelLength = Len(.Text): End With
    Status.Panels(1).Text = "Número de máquina del producto (# serie)."
End Sub

Private Sub tPNroMaquina_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tPNroMaquina.Tag) <> 0 And Trim(lPIdProducto.Caption) <> "" Then
            ZActualizoCampoProducto CLng(lPIdProducto.Caption), Trim(tPNroMaquina.Text), NroMaquina:=True
            tPNroMaquina.Tag = 0
        End If
        Foco tMotivo
    End If
End Sub

Private Sub tSCodigo_Change()
    If Val(tSCodigo.Tag) <> 0 Then LimpioFicha
End Sub

Private Sub tSCodigo_GotFocus()
    Status.Panels(1).Text = "Número de servicio a consultar."
End Sub

Private Sub tSCodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(tSCodigo.Text) Then Exit Sub
        If Val(tSCodigo.Tag) <> 0 Then If tMotivo.Enabled Then Foco tMotivo Else If tLiquidar.Enabled Then Foco tLiquidar Else Foco tFCumplido: Exit Sub
        
        CargoDatosServicio Val(tSCodigo.Text)
    End If
    
End Sub

Private Sub tUsuario_Change()
    tUsuario.Tag = 0
End Sub

Private Sub tUsuario_GotFocus()
    With tUsuario: .SelStart = 0: .SelLength = Len(.Text): End With
    Status.Panels(1).Text = "Último usuario que trabajó con la ficha."
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Val(tUsuario.Tag) <> 0 Then AccionGrabar: Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(tUsuario.Text) Then Exit Sub
        Dim aId As Long
        aId = BuscoUsuarioDigito(CLng(tUsuario.Text), Codigo:=True)
        tUsuario.Text = BuscoUsuario(aId, Identificacion:=True)
        tUsuario.Tag = aId
        If Val(tUsuario.Tag) <> 0 Then AccionGrabar
    End If
    
End Sub

Private Function ArticuloIngresado(IDArticulo As Long) As Boolean

    On Error GoTo errFunction
    ArticuloIngresado = True
    With vsMotivo
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 0) = IDArticulo Then
                MsgBox "El artículo " & .Cell(flexcpText, I, 0) & " ya está ingresado en la lista." & Chr(vbKeyReturn) & "Para modifcar la cantidad elimínelo de la lista y vuelva a ingresarlo.", vbInformation, "Item Ingresado"
                Screen.MousePointer = 0: Exit Function
            End If
        Next
    End With
    '-----------------------------------------------------------------------------------------------------
    ArticuloIngresado = False
    Exit Function

errFunction:
End Function

Private Sub ZActualizoCampoProducto(idProducto As Long, Valor As Variant, _
                                Optional FCompra As Boolean = False, Optional FacturaS As Boolean = False, Optional FacturaN As Boolean = False, _
                                Optional NroMaquina As Boolean = False, Optional Articulo As Boolean = False)
    
    On Error GoTo errActualizar
    If idProducto = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    Cons = "Select * from Producto Where ProCodigo = " & idProducto
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        
        If FCompra Then If Trim(Valor) = "" Then RsAux!ProCompra = Null Else RsAux!ProCompra = Format(Valor, sqlFormatoF)
        If FacturaS Then If Trim(Valor) = "" Then RsAux!ProFacturaS = Null Else RsAux!ProFacturaS = Trim(Valor)
        If FacturaN Then If Trim(Valor) = "" Then RsAux!ProFacturaN = Null Else RsAux!ProFacturaN = CLng(Valor)
        If NroMaquina Then If Trim(Valor) = "" Then RsAux!ProNroSerie = Null Else RsAux!ProNroSerie = Trim(Valor)
        If Articulo Then If Valor <> 0 Then RsAux!ProArticulo = CLng(Valor)
        RsAux.Update
    End If
    RsAux.Close
    
    If Articulo Then lPGarantia.Caption = " " & RetornoGarantia(tPArticulo.Tag)
    If FCompra Or Articulo Then
        lPEstado.Tag = CalculoEstadoProducto(gProducto)
        lPEstado.Caption = " " & EstadoProducto(Val(lPEstado.Tag))
    End If
    Screen.MousePointer = 0
    
    Exit Sub
errActualizar:
    clsGeneral.OcurrioError "Ocurrió un error al actualizar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub vsMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not tMotivo.Enabled Then Exit Sub
    Select Case KeyCode
        Case vbKeyDelete
            If vsMotivo.Rows = 1 Then Exit Sub
            vsMotivo.RemoveItem vsMotivo.Row
    End Select
    
End Sub

