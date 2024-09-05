VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F93D243E-5C15-11D5-A90D-000021860458}#10.0#0"; "orFecha.ocx"
Begin VB.Form frmDeOperacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Operaciones"
   ClientHeight    =   6270
   ClientLeft      =   1935
   ClientTop       =   2145
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDeOperacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7905
   Begin VB.CommandButton bCalcular 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7560
      TabIndex        =   28
      Top             =   5940
      Width           =   255
   End
   Begin orctFecha.orFecha tFechaPago 
      Height          =   285
      Left            =   6240
      TabIndex        =   27
      Top             =   5940
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.Width           =   1275
      EnabledMes      =   -1  'True
      EnabledPrimerUltimoDia=   -1  'True
      FechaFormato    =   "Ddd dd/mm/yyyy"
      FechaValor      =   "01/01/2001"
   End
   Begin VSFlex6DAOCtl.vsFlexGrid lArticulo 
      Height          =   1215
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2143
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
   Begin VB.CommandButton bPago 
      Caption         =   "  &Pagos..."
      Height          =   325
      Left            =   6915
      TabIndex        =   16
      Top             =   690
      Width           =   855
   End
   Begin VSFlex6DAOCtl.vsFlexGrid lCuota 
      Height          =   3255
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5741
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
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
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
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
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
   Begin VB.Label lCtasVencidas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe Acumulado (Ctas. Vencidas):"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Top             =   5940
      Width           =   6075
   End
   Begin VB.Label lCliente 
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
      Height          =   270
      Left            =   120
      TabIndex        =   22
      Top             =   30
      Width           =   7695
   End
   Begin VB.Label lAnulada 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "ANULADA !!"
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
      Height          =   225
      Left            =   5520
      TabIndex        =   21
      Top             =   795
      Width           =   1335
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDeOperacion.frx":0442
            Key             =   "Si"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDeOperacion.frx":075C
            Key             =   "No"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDeOperacion.frx":0A76
            Key             =   "Vencida"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDeOperacion.frx":0D90
            Key             =   "Blanco"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDeOperacion.frx":10AA
            Key             =   "Gestor"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDeOperacion.frx":13C4
            Key             =   "Alerta"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Cumplimiento:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   20
      Top             =   2370
      Width           =   1095
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Puntaje:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   19
      Top             =   2370
      Width           =   735
   End
   Begin VB.Label lCumplimiento 
      BackStyle       =   0  'Transparent
      Caption         =   "0223332000"
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
      Left            =   3360
      TabIndex        =   18
      Top             =   2370
      Width           =   2175
   End
   Begin VB.Label lPuntaje 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   7200
      TabIndex        =   17
      Top             =   2370
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelada:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2370
      Width           =   855
   End
   Begin VB.Label lCancelada 
      BackStyle       =   0  'Transparent
      Caption         =   "$"
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
      Left            =   1200
      TabIndex        =   14
      Top             =   2370
      Width           =   375
   End
   Begin VB.Label lComentario 
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label lMoneda 
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7200
      TabIndex        =   12
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Carlos"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lSucursal 
      BackStyle       =   0  'Transparent
      Caption         =   "Depósito Colonia"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "14-Ene-1998"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "A 00000"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lTipo 
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda:"
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Comentarios:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Sucursal:"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión:"
      Height          =   255
      Left            =   4140
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Número:"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   120
      Top             =   330
      Width           =   7695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   2320
      Width           =   7695
   End
End
Attribute VB_Name = "frmDeOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gDocumento As Long
Dim gCredito As Long
Dim sHayNota As Boolean

Dim aTexto As String, pathApp As String

Dim mTipoCredito As Integer

Private Sub bCalcular_Click()
    Call tFechaPago_KeyPress(vbKeyReturn)
End Sub

Private Sub bPago_Click()
    EjecutarApp pathApp & "Detalle de pagos", CStr(gDocumento)
End Sub

Private Sub bPago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    pathApp = App.Path & "\"
    
    FechaDelServidor
    tFechaPago.FechaValor = gFechaServidor
    gCredito = 0
    
    ObtengoSeteoForm Me, Me.Left, Me.Top, Me.Width, Me.Height
    
    'Linea de Comandos    -------------------------------------
    If Trim(Command()) <> "" Then
        aTexto = Trim(Command())
        gDocumento = Val(aTexto)
        'gDocumento = 361229
    End If
    '---------------------------------------------------------------
    InicializoGrillas
    LimpioFicha
    
    sHayNota = False
    If gDocumento <> 0 Then
        bPago.Enabled = True
        CargoArticulo gDocumento
        CargoCredito gDocumento
    End If
    
    If gCredito <> 0 And lCuota.Enabled Then CargoCuotas gCredito
    
    If gCredito <> 0 And mTipoCredito = TipoCredito.Gestor Then
        Me.Show
        MsgBox "El crédito está en gestor." & vbCrLf & vbCrLf & _
                    "Cuando el crédito está en gestor no se le pueden cobrar cuotas ni INFORMARLE ningún saldo." & vbCrLf & _
                    "Sólo se puede informar las fechas e importes que pagó (ver pagos).", vbInformation, "Credito EN GESTOR"
    End If
    
End Sub

Private Sub CargoCliente(Codigo As Long)

    Cons = "Select Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
               & " From CPersona " _
               & " Where CPeCliente = " & Codigo _
                                                    & " UNION ALL" _
               & " Select Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
               & " From CEmpresa " _
               & " Where CEmCliente = " & Codigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Nombre) Then lCliente.Caption = Trim(RsAux!Nombre)
    End If
    RsAux.Close
    
    If mTipoCredito = TipoCredito.Gestor Then lCliente.Caption = Trim(lCliente.Caption) & " (EN GESTOR)"
    
End Sub

Private Sub CargoCredito(Documento As Long)

Dim aCliente As Long: aCliente = 0
    
    On Error GoTo errCargar
    Cons = "Select * From Documento Left Outer Join Credito On DocCodigo = CreFactura" & _
                                    " Left Outer Join TipoCuota On CreTipoCuota = TCuCodigo, " & _
                    " Sucursal, Moneda " & _
            " Where DocCodigo = " & Documento & _
            " And DocSucursal = SucCodigo And DocMoneda = MonCodigo "
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    aCliente = RsAux!DocCliente
    lMoneda.Tag = RsAux!DocMoneda
    If Not IsNull(RsAux!CreCodigo) Then
        'Operacion Credito
        gCredito = RsAux!CreCodigo
        
        'Cancelada
        lCancelada.Caption = "NO"
        If RsAux!CreSaldoFactura = 0 Then lCancelada.Caption = "SI"
        If Not IsNull(RsAux!CreCumplimiento) Then lCumplimiento.Caption = Trim(RsAux!CreCumplimiento) Else lCumplimiento.Caption = "S/D"
        lPuntaje.Caption = "N/D"
        If Not IsNull(RsAux!CrePuntaje) Then lPuntaje.Caption = Trim(RsAux!CrePuntaje)
    
        mTipoCredito = RsAux!CreTipo
        
    Else
        'Cancelada
        lCancelada.Caption = "N/D"
        lCumplimiento.Caption = "N/D"
        lPuntaje.Caption = "N/D"
        If Not sHayNota Then bPago.Enabled = False
        lCuota.Enabled = False
    End If
    
    lTipo.Caption = RetornoNombreDocumento(RsAux!DocTipo)
    If Not IsNull(RsAux!TCuAbreviacion) Then lTipo.Caption = lTipo.Caption & " (" & Trim(RsAux!TCuAbreviacion) & ")"
    
    Select Case RsAux!DocTipo
        Case TipoDocumento.Credito
            If Not IsNull(RsAux!CreFormaPago) Then If RsAux!CreFormaPago = TipoPagoSolicitud.ChequeDiferido Then lTipo.Caption = Trim(lTipo.Caption) & " (Ch. Dif.)"
               
        Case Else
            lArticulo.Height = Me.ScaleHeight - lArticulo.Top - 40
            tFechaPago.Visible = False: bCalcular.Visible = False
            lCtasVencidas.Visible = False
        
    End Select
    
    If RsAux!DocAnulado Then lAnulada.Visible = True Else lAnulada.Visible = False ' lAnulada.Caption = "Anulada" Else lAnulada.Caption = ""
    lNumero.Caption = Trim(RsAux!DocSerie) & " " & RsAux!DocNumero
    lFecha.Caption = Format(RsAux!DocFecha, "d-Mmm-yy hh:mm")
    lUsuario.Caption = z_BuscoUsuario(RsAux!DocUsuario, Identificacion:=True)
    
    lSucursal.Caption = "N/D"
    If Not IsNull(RsAux!DocSucursal) Then lSucursal.Caption = Trim(RsAux!SucAbreviacion)
    
    lComentario.Caption = "N/D"
    
    If Not IsNull(RsAux!DocComentario) Then lComentario.Caption = Trim(RsAux!DocComentario)
    lMoneda.Caption = Trim(RsAux!Monsigno)
        
    RsAux.Close
    
    CargoCliente aCliente
    
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del crédito"
End Sub

Private Sub CargoCuotas(Credito As Long)

Dim aValor As Currency, aCtasVencidas As Currency
Dim aVence As Date

    aCtasVencidas = 0
    
    On Error GoTo errCargar
    Dim mCoefMora As Double, mRedondeo As String
    mCoefMora = dis_arrMonedaProp(Val(lMoneda.Tag), enuMoneda.pCoeficienteMora)
    mRedondeo = dis_arrMonedaProp(Val(lMoneda.Tag), enuMoneda.pRedondeo)
    
    lCuota.Rows = 1
    Cons = "Select * From CreditoCuota" _
           & " Where CCuCredito = " & Credito _
           & " Order by CCuVencimiento"
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        With lCuota
            .AddItem Format(RsAux!CCuVencimiento, "dd/mm/yyyy")
            
            If RsAux!CCuNumero = 0 Then aTexto = "E" Else aTexto = Trim(RsAux!CCuNumero)
            .Cell(flexcpText, .Rows - 1, 1) = aTexto
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!CCuValor, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!CCuSaldo, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!CCuValor - RsAux!CCuSaldo, "#,##0.00")
        
            If Not IsNull(RsAux!CCuMora) Then
                .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!CCuMora, "#,##0.00")
            Else
                .Cell(flexcpText, .Rows - 1, 5) = "0.00"
            End If
        
            aVence = CDate(.Cell(flexcpText, .Rows - 1, 0))
            If Not IsNull(RsAux!CCuUltimoPago) Then
                .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!CCuUltimoPago, "dd/mm/yy")
                If DateDiff("d", aVence, RsAux!CCuUltimoPago) > 0 Then .Cell(flexcpText, .Rows - 1, 6) = DateDiff("d", aVence, RsAux!CCuUltimoPago)
                
            Else
                If DateDiff("d", aVence, gFechaServidor) > 0 Then .Cell(flexcpText, .Rows - 1, 6) = DateDiff("d", aVence, gFechaServidor)
            End If
        
            'Si el la Cuota no está cancelada, Calculamos MORA Y ponemos Icono
            If RsAux!CCuSaldo > 0 And Not lAnulada.Visible Then
                .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages(IconoDeVencimiento(RsAux!CCuVencimiento)).ExtractIcon
            Else
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            End If
            
            If Not lAnulada.Visible Then
                If RsAux!CCuSaldo > 0 And RsAux!CCuVencimiento < gFechaServidor Then
                    If Abs(RsAux!CCuVencimiento - gFechaServidor) > paToleranciaMora Then
                        If IsNull(RsAux!CCuUltimoPago) Then
                            aValor = CalculoMora(RsAux!CCuSaldo, RsAux!CCuVencimiento, RsAux!CCuMoraACuenta, mCoefMora)
                        Else
                            Dim mFDesde As Date
                            mFDesde = IIf(RsAux!CCuUltimoPago > RsAux!CCuVencimiento, RsAux!CCuUltimoPago, RsAux!CCuVencimiento)
                            aValor = CalculoMora(RsAux!CCuSaldo, mFDesde, RsAux!CCuMoraACuenta, mCoefMora)
                        End If
                        aValor = Redondeo(aValor, mRedondeo)
                        'Se lo Sumo al Saldo
                        .Cell(flexcpText, .Rows - 1, 3) = Format(.Cell(flexcpValue, .Rows - 1, 3) + aValor, FormatoMonedaP)
                    End If
                    aCtasVencidas = aCtasVencidas + .Cell(flexcpValue, .Rows - 1, 3)
                End If
            End If
        End With
        RsAux.MoveNext
        
    Loop
    RsAux.Close
    
    'Totales
    With lCuota
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 2, , Colores.Rojo, Colores.Blanco, True, "TOTALES"
        .Subtotal flexSTSum, -1, 3
        .Subtotal flexSTSum, -1, 4
        .Subtotal flexSTSum, -1, 5
    End With
    
    lCtasVencidas.Caption = "Importe Acumulado (Ctas. Vencidas): " & Format(aCtasVencidas, FormatoMonedaP) & "              Pagando &el: "
    
    If aCtasVencidas = 0 Then mTipoCredito = 0
    Exit Sub

errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos de las cuotas.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoArticulo(Documento As Long)
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    
    Cons = "Select Tipo = 'C', ArtCodigo, ArtNombre, RenCantidad, RenPrecio" _
           & " From Renglon, Articulo" _
           & " Where RenDocumento = " & Documento _
           & " And RenArticulo = ArtID" _
                    & " UNION ALL " _
           & "Select Tipo = 'D', ArtCodigo, ArtNombre, RenCantidad, RenPrecio" _
           & " From Nota, Documento, Renglon, Articulo" _
           & " Where NotFactura = " & Documento _
           & " And NotNota = DocCodigo " _
           & " And DocCodigo = RenDocumento" _
           & " And RenArticulo = ArtID " _
           & " And DocAnulado = 0"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        With lArticulo
            .AddItem "(" & Format(RsAux!ArtCodigo, "#,000,000") & ") " & Trim(RsAux!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 1) = RsAux!RenCantidad
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!RenPrecio * RsAux!RenCantidad, FormatoMonedaP)
            
            'Veo si es Devolucion para ponerle el icono
            If Trim(RsAux!Tipo) = "D" Then
                .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.RojoClaro
                .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("Vencida").ExtractIcon
                sHayNota = True
            End If
        End With
        RsAux.MoveNext
    Loop
    
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos comprados.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing

End Sub

Private Sub lArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me
    
End Sub

Private Sub lCuota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub InicializoGrillas()

    With lCuota
        .Rows = 1: .Cols = 1
        .FormatString = ">Vencimiento|<Nº|>Importe|>Saldo|>Amortización|>Mora Paga|>D/A|Pagó el"
        .ColWidth(0) = 1300: .ColWidth(1) = 500: .ColWidth(2) = 1200: .ColWidth(3) = 1200
        '.ColWidth(5) = 1200:
        .ColWidth(6) = 500
        .WordWrap = False: .ExtendLastCol = True
        .ColDataType(0) = flexDTDate
    End With

    With lArticulo
        .Rows = 1: .Cols = 1
        .FormatString = "<Artículo|>Q|>Importe"
        .ColWidth(0) = 5300: .ColWidth(1) = 500
        .WordWrap = False: .ExtendLastCol = True
    End With

End Sub

Private Sub LimpioFicha()
    bPago.Enabled = False
    lTipo.Caption = "N/D": lNumero.Caption = "N/D": lFecha.Caption = "N/D": lMoneda.Caption = "N/D"
    lComentario.Caption = "N/D"
    lUsuario.Caption = "N/D": lSucursal.Caption = "N/D": lAnulada.Visible = False ' lAnulada.Caption = ""
        
    lCancelada.Caption = "N/D"
    lCumplimiento.Caption = "N/D"
    lPuntaje.Caption = "N/D"
    
    lArticulo.Rows = 1: lCuota.Rows = 1
End Sub

Private Sub tFechaPago_KeyPress(KeyAscii As Integer)
    On Error GoTo errCalculo
    
    If KeyAscii = vbKeyReturn Then
        If Not IsDate(tFechaPago.FechaValor) Then Exit Sub
        If CDate(tFechaPago.FechaValor) < Date Then
            tFechaPago.FechaValor = Date
            Exit Sub
        End If
        
        Screen.MousePointer = 11
        gFechaServidor = CDate(tFechaPago.FechaValor & " " & Time)
        CargoCuotas gCredito
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errCalculo:
    clsGeneral.OcurrioError "Error al procesar el cálculo de los vencimientos.", Err.Description
    Screen.MousePointer = 0
End Sub
