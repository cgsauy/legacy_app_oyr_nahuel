VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmListado 
   Caption         =   "Diferencias de Cambio"
   ClientHeight    =   7530
   ClientLeft      =   1575
   ClientTop       =   1830
   ClientWidth     =   10830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10830
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Width           =   3915
      _ExtentX        =   6906
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
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   7858
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
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   11595
      TabIndex        =   6
      Top             =   6720
      Width           =   11655
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   480
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Grabar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0544
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   3120
         Picture         =   "frmListado.frx":0FF2
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2760
         Picture         =   "frmListado.frx":10DC
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   2040
         Picture         =   "frmListado.frx":11C6
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmListado.frx":1400
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "frmListado.frx":1502
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "frmListado.frx":18C8
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":19CA
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1320
         Picture         =   "frmListado.frx":1CCC
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1680
         Picture         =   "frmListado.frx":200E
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   960
         Picture         =   "frmListado.frx":2310
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   265
         Left            =   6000
         TabIndex        =   19
         Top             =   140
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   7275
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10874
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   10335
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "TC Anterior:"
         Height          =   195
         Left            =   6300
         TabIndex        =   23
         Top             =   285
         Width           =   975
      End
      Begin VB.Label lTCAnterior 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Cambio:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7260
         TabIndex        =   22
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label lTC 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Cambio:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3960
         TabIndex        =   21
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TC fin de Mes:"
         Height          =   195
         Left            =   2880
         TabIndex        =   20
         Top             =   285
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Mes a Procesar:"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   285
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

Private Sub AccionLimpiar()
    tFecha.Text = ""
    lTC.Caption = "": lTCAnterior.Caption = ""
    vsConsulta.Rows = 1
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub AccionGrabar()
Dim txtError As String
    If vsConsulta.Rows = 1 Then
        MsgBox "Para grabar las diferencias primero debe procesar la información.", vbExclamation, "GRABAR"
        Exit Sub
    End If
    
    If MsgBox("Confirma grabar las diferencias de cambio para el mes de " & tFecha.Text, vbQuestion + vbYesNo + vbDefaultButton2, "Grabar Diferencias") = vbNo Then Exit Sub
    
    On Error GoTo errorBT
    Dim aCompra As Long, aIdSubrubro As Long
    Dim rsCom As rdoResultset
    
    Dim mNewID As Long
    
    pbProgreso.Max = vsConsulta.Rows - 1
    pbProgreso.Value = 0
    FechaDelServidor
    cBase.BeginTrans    'COMIENZO TRANSACCION----------------------------------------------!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    With vsConsulta
        'ColData --> 0=idEmbarque, 1=idCaso, 2=idProveedorGasto
        For I = 1 To .Rows - 1  'Por el Total
            pbProgreso.Value = pbProgreso.Value + 1
            txtError = .Cell(flexcpText, I, 0)
            
            If .Cell(flexcpBackColor, I, 0) <> Colores.Inactivo Then
            
                Select Case .Cell(flexcpData, I, 1)
                    Case 1, 2: If .Cell(flexcpValue, I, 8) > 0 Then aIdSubrubro = paSubrubroDifCambio Else aIdSubrubro = paSubrubroDifCambioG
                    Case 3, 4: aIdSubrubro = paSubrubroDivisa
                End Select
                '0) Autonumerico en Tabla cceComprobantes      ----------------------------------------------------------------------
                mNewID = -1
                Cons = "Select * from ZureoCGSA.dbo.genAutonumerico Where AutTabla = 'cceComprobantes'"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    mNewID = RsAux!AutContador + 1
                    RsAux.Edit
                    RsAux!AutContador = mNewID
                    RsAux.Update
                End If
                RsAux.Close
                If mNewID = -1 Then Err.Raise 8000, "DBFncs", "Resultado de la función get_TableCounter = -1"

                    
                '1) Cabezal con los datos del Comprobante   ----------------------------------------------------------------------
                Cons = "Select * from ZureoCGSA.dbo.cceComprobantes Where ComID = 0"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                RsAux.AddNew
                RsAux!ComIDEmpresa = 1
                RsAux!ComID = mNewID
                RsAux!ComProveedor = .Cell(flexcpData, I, 2)
                RsAux!ComFecha = Format(UltimoDia(tFecha.Text), sqlFormatoF)
                
                RsAux!ComTipo = TipoDocumento.CompraSalidaCaja
                RsAux!ComMoneda = paMonedaPesos
                RsAux!ComTotal = .Cell(flexcpValue, I, 8)
                RsAux!ComTC = 1
                
                RsAux!ComNumero = "DC" & Format(tFecha.Text, "yyyymm")
                
                aTexto = "DC " & Format(tFecha.Text, "MMM/yy") & " " & vsConsulta.Cell(flexcpText, I, 0) & "; "
                If Trim(vsConsulta.Cell(flexcpText, I, 2)) <> "" Then aTexto = aTexto & " LC: " & vsConsulta.Cell(flexcpText, I, 2) & " "
                aTexto = aTexto & vsConsulta.Cell(flexcpText, I, 1)
                RsAux!ComMemo = Trim(aTexto)
            
                RsAux!ComFechaModificacion = Now
                RsAux!ComSaldoCero = Null
                RsAux.Update
                RsAux.Close
                
                txtError = txtError & " Cta1: " & aIdSubrubro & " Cta2: " & .Cell(flexcpData, I, 5)
                '2) Paso las cuentas asignadas al comprobante (en CGSA estan separadas) ---------------------------------------------------
                Cons = "Select * from  ZureoCGSA.dbo.cceComprobanteCuenta Where CCuIDComprobante = " & mNewID
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                RsAux.AddNew
                RsAux!CCuIDComprobante = mNewID
                RsAux!CCuIDCuenta = aIdSubrubro
                RsAux!CCuIDProyecto = 0: RsAux!CCuIDDepartamento = 0: RsAux!CCuIDReferencia = 0
                RsAux!CCuMoneda = 0
                
                RsAux!CCuImporteCuenta = .Cell(flexcpValue, I, 8)
                RsAux!CCuDebe = .Cell(flexcpValue, I, 8)
                RsAux!CCuHaber = Null
                RsAux.Update
                
                RsAux.AddNew
                RsAux!CCuIDComprobante = mNewID
                RsAux!CCuIDCuenta = Val(.Cell(flexcpData, I, 5))  'Es la cta Acreedor del Banco
                RsAux!CCuIDProyecto = 0: RsAux!CCuIDDepartamento = 0: RsAux!CCuIDReferencia = 0
                RsAux!CCuMoneda = 0
                If (.Cell(flexcpData, I, 6) <> 0) And (.Cell(flexcpData, I, 6) <> paMonedaPesos) Then
                    RsAux!CCuImporteCuenta = 0 '.Cell(flexcpValue, I, 8)
                Else
                    RsAux!CCuImporteCuenta = .Cell(flexcpValue, I, 8)
                End If
                RsAux!CCuDebe = Null
                RsAux!CCuHaber = .Cell(flexcpValue, I, 8)
                RsAux.Update

                
                If aIdSubrubro = paSubrubroDivisa Then
                    'Cargo tabla: GastoImportacion----------------------------------------------------------------
                    Cons = "Select * from GastoImportacion Where GImIDCompra = " & aCompra
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    RsAux.AddNew
                    RsAux!GImIDCompra = mNewID
                    RsAux!GImIDSubrubro = aIdSubrubro
                    RsAux!GImImporte = .Cell(flexcpValue, I, 8)
                    RsAux!GImCostear = .Cell(flexcpValue, I, 8)
                    RsAux!GImNivelFolder = Folder.cFEmbarque
                    RsAux!GImFolder = .Cell(flexcpData, I, 0)
                    RsAux.Update: RsAux.Close
                End If
            End If
        Next
    End With
    
    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
    
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError txtError & vbCrLf & "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError txtError & vbCrLf & "No se ha podido realizar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub AccionGrabar_BKUPViejo()

    If vsConsulta.Rows = 1 Then
        MsgBox "Para grabar las diferencias primero debe procesar la información.", vbExclamation, "GRABAR"
        Exit Sub
    End If
    
    If MsgBox("Confirma grabar las diferencias de cambio para el mes de " & tFecha.Text, vbQuestion + vbYesNo + vbDefaultButton2, "Grabar Diferencias") = vbNo Then Exit Sub
    
    On Error GoTo errorBT
    Dim aCompra As Long, aIdSubrubro As Long
    Dim rsCom As rdoResultset
    
    pbProgreso.Max = vsConsulta.Rows - 2
    pbProgreso.Value = 0
    FechaDelServidor
    cBase.BeginTrans    'COMIENZO TRANSACCION----------------------------------------------!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    Cons = "Select * from Compra Where ComCodigo = 0"
    Set rsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
    With vsConsulta
        'ColData --> 0=idEmbarque, 1=idCaso, 2=idProveedorGasto
        For I = 1 To .Rows - 2  'Por el Total
            pbProgreso.Value = pbProgreso.Value + 1
            
            If .Cell(flexcpBackColor, I, 0) <> Colores.Inactivo Then
            
                Select Case .Cell(flexcpData, I, 1)
                    Case 1, 2: If .Cell(flexcpValue, I, 8) > 0 Then aIdSubrubro = paSubrubroDifCambio Else aIdSubrubro = paSubrubroDifCambioG
                    Case 3, 4: aIdSubrubro = paSubrubroDivisa
                End Select
                
                'Cargo tabla: Compra----------------------------------------------------------------
                rsCom.AddNew
                rsCom!ComTipoDocumento = TipoDocumento.CompraCredito
                rsCom!ComFecha = Format(UltimoDia(tFecha.Text), sqlFormatoF)
                rsCom!ComProveedor = .Cell(flexcpData, I, 2)
                rsCom!ComMoneda = paMonedaPesos
                rsCom!ComImporte = .Cell(flexcpValue, I, 8)
                
                rsCom!ComSerie = "DC"
                rsCom!ComNumero = Format(tFecha.Text, "yyyymm")
                
                rsCom!ComIva = Null
                rsCom!ComTC = CCur(lTC.Tag)
                
                aTexto = "DC " & Format(tFecha.Text, "MMM/yy") & " " & vsConsulta.Cell(flexcpText, I, 0) & "; "
                If Trim(vsConsulta.Cell(flexcpText, I, 2)) <> "" Then aTexto = aTexto & " LC: " & vsConsulta.Cell(flexcpText, I, 2) & " "
                aTexto = aTexto & vsConsulta.Cell(flexcpText, I, 1)
                rsCom!ComComentario = Trim(aTexto)
                
                rsCom!ComFModificacion = Format(gFechaServidor, sqlFormatoFH)
                rsCom!ComSaldo = .Cell(flexcpValue, I, 8)
                rsCom!ComDCDe = .Cell(flexcpData, I, 3)
                rsCom.Update
                
                'Max ID Compra-------------------------------------------------------------------------
                Cons = "Select Max(ComCodigo) from Compra"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                aCompra = RsAux(0)
                RsAux.Close
                
                'Cargo tabla: GastoSubrubro----------------------------------------------------------------
                Cons = "Select * from GastoSubrubro Where GSrIDCompra = " & aCompra
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                RsAux.AddNew
                RsAux!GSrIDCompra = aCompra
                RsAux!GSrIDSubrubro = aIdSubrubro
                RsAux!GSrImporte = .Cell(flexcpValue, I, 8)
                RsAux.Update: RsAux.Close
                
                If aIdSubrubro = paSubrubroDivisa Then
                    'Cargo tabla: GastoImportacion----------------------------------------------------------------
                    Cons = "Select * from GastoImportacion Where GImIDCompra = " & aCompra
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    RsAux.AddNew
                    RsAux!GImIDCompra = aCompra
                    RsAux!GImIDSubrubro = aIdSubrubro
                    RsAux!GImImporte = .Cell(flexcpValue, I, 8)
                    RsAux!GImCostear = .Cell(flexcpValue, I, 8)
                    RsAux!GImNivelFolder = Folder.cFEmbarque
                    RsAux!GImFolder = .Cell(flexcpData, I, 0)
                    RsAux.Update: RsAux.Close
                End If
            End If
        Next
    End With
    
    rsCom.Close
    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
    
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub bImprimir_Click()
    AccionImprimir True
End Sub
Private Sub bNoFiltros_Click()
    AccionLimpiar
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


Private Sub chVista_Click()
    
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
    Else
        AccionImprimir
        vsListado.ZOrder 0
    End If

End Sub

Private Sub Label2_Click()
    Foco tFecha
End Sub

Private Sub tFecha_Change()
    If vsConsulta.Rows > 1 Then vsConsulta.Rows = 1
End Sub

Private Sub tFecha_GotFocus()
    With tFecha: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsDate(tFecha.Text) Then
            Dim aTC As Currency, aFechaTC As String
            
            'Ultimo dia del Mes a Procesar
            aTC = TasadeCambio(paMonedaDolar, paMonedaPesos, UltimoDia(CDate(tFecha.Text)), aFechaTC)
            lTC.Caption = " " & Format(aTC, "#.000") & " del " & Format(aFechaTC, "Ddd d/mm/yy")
            lTC.Tag = CCur(Format(aTC, "#.000"))
            
            'Ultimo dia del Mes Anterior a Procesar
            aTC = TasadeCambio(paMonedaDolar, paMonedaPesos, PrimerDia(CDate(tFecha.Text)) - 1, aFechaTC)
            lTCAnterior.Caption = " " & Format(aTC, "#.000") & " del " & Format(aFechaTC, "Ddd d/mm/yy")
            lTCAnterior.Tag = CCur(Format(aTC, "#.000"))
            
        End If
        bConsultar.SetFocus
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas
    AccionLimpiar
    
    FechaDelServidor
    tFecha.Text = Format(gFechaServidor, "Mmmm yyyy")

    bCargarImpresion = True
    With vsListado
        .PaperSize = 1
        .Orientation = orLandscape
        .Zoom = 100
        .MarginLeft = 900: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
    End With
    
    If paSubrubroDifCambioG = 0 Then MsgBox "El parámetro Diferencias de Cambio Ganadas no está cargado." & Chr(vbKeyReturn) & "Consulte a su administrador de bases de datos.", vbInformation, "Faltan Parámetros"
    If paSubrubroDifCambio = 0 Then MsgBox "El parámetro Diferencias de Cambio Perdidas no está cargado." & Chr(vbKeyReturn) & "Consulte a su administrador de bases de datos.", vbInformation, "Faltan Parámetros"
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Carpeta|<Banco|<Nº LC|<M/Orig.|>Arbitraje|>Divisa en U$S|>Divisa en $|TC ant.|>Dif. TC anterior|>A Importaciones|> A Dif. Cambio|"
        
        .ColWidth(0) = 690: .ColWidth(1) = 1600: .ColWidth(2) = 800
        .ColWidth(5) = 1300: .ColWidth(6) = 1600: .ColWidth(7) = 650
        .ColWidth(10) = 1300
        .WordWrap = False
        
        .ColHidden(8) = True
        '.MergeCells = flexMergeSpill
    End With
      
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            
            Case vbKeyE: AccionConsultar
            
            Case vbKeyP: IrAPagina vsListado, 1
            Case vbKeyA: IrAPagina vsListado, vsListado.PreviewPage - 1
            Case vbKeyS: IrAPagina vsListado, vsListado.PreviewPage + 1
            Case vbKeyU: IrAPagina vsListado, vsListado.PageCount
            
            Case vbKeyAdd: Zoom vsListado, vsListado.Zoom + 5
            Case vbKeySubtract: Zoom vsListado, vsListado.Zoom - 5
            
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir True
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyC: AccionConfigurar
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11

    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    vsListado.Left = fFiltros.Left
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 150
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
Dim aTotalF As Currency

    On Error GoTo errConsultar
    If Not ValidoCampos Then Exit Sub
    Screen.MousePointer = 11
    vsConsulta.Rows = 1: bCargarImpresion = True: vsConsulta.Refresh
    
    Dim aPrimerDia As Date, aUltimoDia As Date      '1=true, 0=false
    aPrimerDia = PrimerDia(CDate(tFecha.Text)) & " 00:00:00"
    aUltimoDia = UltimoDia(CDate(tFecha.Text)) & " 23:59:59"
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    Dim aQTotal As Long: aQTotal = 0
    
    
'    Cons = " Select * from Embarque inner join Carpeta on carid = embcarpeta " & _
'        " Where CarCodigo = 5282"
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    If Not RsAux.EOF Then CargoDatos Caso:=1, ADifCambio:=True
'    RsAux.Close
'    Exit Sub
    
    '1) Divisas Impagas de Embarques Costeados y Arribados (<= Mes a Proceasar)
    Cons = " Select Count(*) from Embarque, Carpeta" & _
                " Where EmbDivisaPaga = 0 " & _
                " And EmbCosteado = 1" & _
                " And EmbFLocal <= '" & Format(aUltimoDia, sqlFormatoFH) & "'" & _
                " And EmbCarpeta = CarID" & _
                " And (CarFAnulada Is Null or CarFAnulada > '" & Format(aUltimoDia, sqlFormatoFH) & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then aQTotal = aQTotal + RsAux(0)
    RsAux.Close
    
    '2) Divisas Impagas de Embarques SIN Costear y arribados en un mes anterior
    Cons = " Select Count(*) from Embarque, Carpeta " & _
                " Where EmbDivisaPaga = 0 " & _
                " And EmbCosteado = 0" & _
                " And EmbFLocal < '" & Format(aPrimerDia, sqlFormatoFH) & "'" & _
                " And EmbCarpeta = CarID" & _
                " And (CarFAnulada Is Null or CarFAnulada > '" & Format(aUltimoDia, sqlFormatoFH) & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then aQTotal = aQTotal + RsAux(0)
    RsAux.Close
    
    '3) Divisas Impagas de Embarques SIN Costear y arribados dentro del mes
    Cons = " Select Count(*) from Embarque, Carpeta left outer join BancoLocal on CarBcoEmisor = BLoCodigo" & _
                " Where EmbDivisaPaga = 0 " & _
                " And EmbCosteado = 0" & _
                " And EmbFLocal between '" & Format(aPrimerDia, sqlFormatoFH) & "' And '" & Format(aUltimoDia, sqlFormatoFH) & "'" & _
                " And EmbCarpeta = CarID" & _
                " And (CarFAnulada Is Null or CarFAnulada > '" & Format(aUltimoDia, sqlFormatoFH) & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then aQTotal = aQTotal + RsAux(0)
    RsAux.Close
    
    '4) Divisas Impagas de Embarques SIN Arribar
    Cons = " Select Count(*) from Embarque, Carpeta" & _
                " Where EmbDivisaPaga = 0 " & _
                " And (EmbFLocal > '" & Format(aUltimoDia, sqlFormatoFH) & "' Or EmbFLocal is null)" & _
                " And EmbCarpeta = CarID " & _
                " And CarFApertura <= '" & Format(aUltimoDia, sqlFormatoFH) & "'" & _
                " And (CarFAnulada Is Null or CarFAnulada > '" & Format(aUltimoDia, sqlFormatoFH) & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then aQTotal = aQTotal + RsAux(0)
    RsAux.Close
    
    If aQTotal = 0 Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        Screen.MousePointer = 0: Exit Sub
    Else
        pbProgreso.Max = aQTotal
    End If
    '-------------------------------------------------------------------------------------------------------------------------------------------
    
    '1) Divisas Impagas de Embarques Costeados y Arribados (<= Mes a Proceasar)
    Cons = " Select * from Embarque, Carpeta " & _
                " Where EmbDivisaPaga = 0 " & _
                " And EmbCosteado = 1" & _
                " And EmbFLocal <= '" & Format(aUltimoDia, sqlFormatoFH) & "'" & _
                " And EmbCarpeta = CarID" & _
                " And (CarFAnulada Is Null or CarFAnulada > '" & Format(aUltimoDia, sqlFormatoFH) & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then CargoDatos Caso:=1, ADifCambio:=True
    RsAux.Close

    '2) Divisas Impagas de Embarques SIN Costear y arribados en un mes anterior
    Cons = " Select * from Embarque, Carpeta" & _
                " Where EmbDivisaPaga = 0 " & _
                " And EmbCosteado = 0" & _
                " And EmbFLocal < '" & Format(aPrimerDia, sqlFormatoFH) & "'" & _
                " And EmbCarpeta = CarID" & _
                " And (CarFAnulada Is Null or CarFAnulada > '" & Format(aUltimoDia, sqlFormatoFH) & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then CargoDatos Caso:=2, ADifCambio:=True
    RsAux.Close
    
    '3) Divisas Impagas de Embarques SIN Costear y arribados dentro del mes
    'Cons = " Select * from Embarque, Carpeta left outer join BancoLocal on CarBcoEmisor = BLoCodigo"
    Cons = " Select * from Embarque, Carpeta" & _
                " Where EmbDivisaPaga = 0 " & _
                " And EmbCosteado = 0" & _
                " And EmbFLocal between '" & Format(aPrimerDia, sqlFormatoFH) & "' And '" & Format(aUltimoDia, sqlFormatoFH) & "'" & _
                " And EmbCarpeta = CarID" & _
                " And (CarFAnulada Is Null or CarFAnulada > '" & Format(aUltimoDia, sqlFormatoFH) & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then CargoDatos Caso:=3, AImportaciones:=True
    RsAux.Close
    
    '4) Divisas Impagas de Embarques SIN Arribar
    Cons = " Select * from Embarque, Carpeta" & _
                " Where EmbDivisaPaga = 0 " & _
                " And (EmbFLocal > '" & Format(aUltimoDia, sqlFormatoFH) & "' Or EmbFLocal is null)" & _
                " And EmbCarpeta = CarID " & _
                " And CarFApertura <= '" & Format(aUltimoDia, sqlFormatoFH) & "'" & _
                " And (CarFAnulada Is Null or CarFAnulada > '" & Format(aUltimoDia, sqlFormatoFH) & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then CargoDatos Caso:=4, AImportaciones:=True
    RsAux.Close
    
    With vsConsulta
        If .Rows > 1 Then
            .Select 1, 0, .Rows - 1
            .Sort = flexSortGenericAscending
            .Select 0, 0, 0, 0
            
            .Cell(flexcpBackColor, 1, 9, .Rows - 1, .Cols - 1) = Colores.Obligatorio
            .Subtotal flexSTSum, -1, 9, , Colores.Rojo, Colores.Blanco, True, "Totales"
            .Subtotal flexSTSum, -1, 10
        End If
    End With
    
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    Screen.MousePointer = 0
    pbProgreso.Value = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub CargoDatos(Caso As Integer, Optional AImportaciones As Boolean = False, Optional ADifCambio As Boolean = False)

Dim aTexto As String, aValor As Long
Dim rsDiv As rdoResultset, aTCDiv As Currency
Dim aIdMoneda As Long, aSignoM As String, pintMonedaCCta As Integer

    aIdMoneda = 0
    With vsConsulta
    ' Nº Carpeta|Banco|Nº LC|M/Original|Arbitraje|Divisa en U$S|Divisa en $|TC anterior|Dif. TC anterior|A Importaciones| A Dif. Cambio
    Do While Not RsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        
        'Busco el/los registros de los gastos de divisa para la carpeta
        'Busco TC Anterior------------------------------------------------------------------------------------------------------
        'Desde el embarque siempre se registra el gasto divisa en Dolares. !!!
        '1) Si la Fecha del G es < al ultimo dia del penultimo mes tomo TC del ultimo dia del penultimo mes
        '2) Si es mayor TC del (G)
        aTCDiv = lTCAnterior.Tag
        'A la vista ProveedorCliente Agregar el mismo campo q a la BancoLocal
        Cons = "Select * from Compra, GastoImportacion, ProveedorCliente " & _
                " Where ComCodigo = GImIDCompra " & _
                " And GImIdSubrubro = " & paSubrubroDivisa & _
                " And GImNivelFolder = " & Folder.cFEmbarque & _
                " And GImFolder = " & RsAux!EmbID & _
                " And ComMoneda = " & paMonedaDolar & _
                " And ComSaldo > 0 " & " And ComFecha <= '" & Format(UltimoDia(CDate(tFecha.Text)), sqlFormatoF) & " 23:59'" & _
                " And ComProveedor = PClCodigo" & _
                " And ComTipoDocumento = " & TipoDocumento.CompraCredito
                
                
        Cons = "Select * from Compra, GastoImportacion, ProveedorCliente " & _
                " Where ComCodigo = GImIDCompra " & _
                " And GImIdSubrubro = " & paSubrubroDivisa & _
                " And GImNivelFolder = " & Folder.cFEmbarque & _
                " And GImFolder = " & RsAux!EmbID & _
                " And ComMoneda = " & paMonedaDolar & _
                " And ComProveedor = PClCodigo" & _
                " And ComTipoDocumento = " & TipoDocumento.CompraCredito
                
        Set rsDiv = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsDiv.EOF
        
            'Agrego los datos de la carpeta------------------------------------------------------
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = RsAux!CarCodigo & "." & Trim(RsAux!EmbCodigo)
            aValor = RsAux!EmbID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            'If Not IsNull(RsAux!BLoNombre) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!BLoNombre)
            If Not IsNull(rsDiv!PClNombre) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(rsDiv!PClNombre) Else .Cell(flexcpText, .Rows - 1, 1) = Trim(rsDiv!PClFantasia)
            aValor = Caso: .Cell(flexcpData, .Rows - 1, 1) = aValor
            
            If Not IsNull(RsAux!CarCartaCredito) Then .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!CarCartaCredito)
            
            If aIdMoneda <> RsAux!EmbMoneda Then
                aIdMoneda = RsAux!EmbMoneda
                aSignoM = BuscoSignoMoneda(RsAux!EmbMoneda)
            End If
            .Cell(flexcpText, .Rows - 1, 3) = aSignoM
            aValor = 0: .Cell(flexcpData, .Rows - 1, 3) = aValor        'Seteo en cero, p/despues va el id de la compra por la DC
            
            If Not IsNull(RsAux!EmbArbitraje) Then .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!EmbArbitraje, "#,##0.000") Else .Cell(flexcpText, .Rows - 1, 4) = "1.000"
            '------------------------------------------------------------------------------------------------------------
            
            'Datos del Gasto registrado como divisa
            '.Cell(flexcpText, .Rows - 1, 5) = Format(rsDiv!ComSaldo, FormatoMonedaP)          'Divisa en Dolares
            .Cell(flexcpText, .Rows - 1, 5) = Format(SaldoCreditoZureo(rsDiv!ComCodigo), FormatoMonedaP)
            
            If rsDiv!ComFecha < CDate(tFecha.Text) Then aTCDiv = lTCAnterior.Tag Else aTCDiv = rsDiv!ComTC
            .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpValue, .Rows - 1, 5) * CCur(lTC.Tag), FormatoMonedaP)  'Div actualizada en $
            
            .Cell(flexcpText, .Rows - 1, 7) = Format(aTCDiv, "#,##0.000")
            
            aValor = rsDiv!ComProveedor: .Cell(flexcpData, .Rows - 1, 2) = aValor
            aValor = rsDiv!ComCodigo: .Cell(flexcpData, .Rows - 1, 3) = aValor
            
            .Cell(flexcpText, .Rows - 1, 8) = Format((.Cell(flexcpValue, .Rows - 1, 6)) - (.Cell(flexcpValue, .Rows - 1, 5) * .Cell(flexcpValue, .Rows - 1, 7)), FormatoMonedaP)
            
            If AImportaciones Then .Cell(flexcpText, .Rows - 1, 9) = .Cell(flexcpText, .Rows - 1, 8)
            If ADifCambio Then .Cell(flexcpText, .Rows - 1, 10) = .Cell(flexcpText, .Rows - 1, 8)
            
            aValor = ContraCuentaZureo(rsDiv!ComCodigo, pintMonedaCCta) 'Para pasar el Gasto
            .Cell(flexcpData, .Rows - 1, 5) = aValor
            .Cell(flexcpData, .Rows - 1, 6) = pintMonedaCCta

            rsDiv.MoveNext
        Loop
        rsDiv.Close
        '----------------------------------------------------------------------------------------------------------------------------
        
        RsAux.MoveNext
    Loop
    End With
    
End Sub

Private Function ContraCuentaZureo(idCredito As Long, retIDMoneda As Integer) As Long
Dim mSQL As String, rsZ As rdoResultset, mValor As Long
    
    retIDMoneda = 0
    
    mSQL = "Select CCuIDCuenta, CueMoneda " & _
          " From ZureoCGSA.dbo.cceComprobanteCuenta" & _
                " Left Outer Join ZureoCGSA.dbo.cceCuentas ON  CCuIDCuenta = CueID" & _
        " Where CCuIDComprobante = " & idCredito & _
        " And CCuHaber IS NOT Null"
    Set rsZ = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not rsZ.EOF Then
        mValor = rsZ!CCuIDCuenta
        If Not IsNull(rsZ!CueMoneda) Then retIDMoneda = rsZ!CueMoneda
    End If
    rsZ.Close

    ContraCuentaZureo = mValor
    
End Function

Private Function SaldoCreditoZureo(idCredito As Long) As Currency

Dim mSQL As String, rsZ As rdoResultset, mValor As Currency

    mValor = 0
    mSQL = "Select ComID, ComTotal, Sum(CPaAsignado) as ComAsignado " & _
            " From ZureoCGSA.dbo.cceComprobantes LEFT OUTER JOIN ZureoCGSA.dbo.cceComprobantePago ON ComID = CPaComprobante " & _
            " Where ComID = " & idCredito & _
            " Group by ComID, ComTotal"
    Set rsZ = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not rsZ.EOF Then
        mValor = rsZ!ComTotal
        If Not IsNull(rsZ!ComAsignado) Then mValor = mValor - rsZ!ComAsignado
    End If
    rsZ.Close
    
    SaldoCreditoZureo = mValor
    
End Function

Private Function BuscoSignoMoneda(IdMoneda As Long) As String

    On Error GoTo errorM
    
    Dim rsMon As rdoResultset, aRetorno As String
    aRetorno = ""
    Cons = "Select * From Moneda Where MonCodigo = " & IdMoneda
    Set rsMon = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsMon.EOF Then
        If Not IsNull(rsMon!MonSigno) Then aRetorno = Trim(rsMon!MonSigno)
    End If
    rsMon.Close
errorM:
    BuscoSignoMoneda = aRetorno
End Function

Private Function ValidoCampos() As Boolean
    On Error Resume Next
    ValidoCampos = False
    
    If Not IsDate(tFecha.Text) Then
        MsgBox "Debe ingresar la fecha para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    
    If Val(lTC.Tag) = 0 Then
        MsgBox "El tipo de cambio no se ha cargado. Presione <Enter> en la fecha a cosnultar.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    
    ValidoCampos = True
    
End Function

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "Mmmm yyyy")
End Sub

Private Sub vsConsulta_DblClick()
    
    With vsConsulta
        If .Row = .Rows - 1 Then Exit Sub
        'Ojo estan en la 9 y la 10 porque la 8 esta oculta
        If .Cell(flexcpBackColor, .Row, 0, , .Cols - 4) <> Colores.Gris Then
            .Cell(flexcpBackColor, .Row, 0, , .Cols - 4) = Colores.Gris
            .Cell(flexcpText, .Rows - 1, 9) = Format(.Cell(flexcpValue, .Rows - 1, 9) - .Cell(flexcpValue, .Row, 9), FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 10) = Format(.Cell(flexcpValue, .Rows - 1, 10) - .Cell(flexcpValue, .Row, 10), FormatoMonedaP)
        Else
            .Cell(flexcpBackColor, .Row, 0, , .Cols - 4) = .BackColor
            .Cell(flexcpText, .Rows - 1, 9) = Format(.Cell(flexcpValue, .Rows - 1, 9) + .Cell(flexcpValue, .Row, 9), FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 10) = Format(.Cell(flexcpValue, .Rows - 1, 10) + .Cell(flexcpValue, .Row, 10), FormatoMonedaP)
        End If
    End With
        
End Sub

Private Sub vsConsulta_GotFocus()
    Status.Panels(4).Text = "[Espacio] o [DblClick] generar/no generar diferencia de cambio."
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then Call vsConsulta_DblClick
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
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
        
        aTexto = "Diferencias de Cambio - " & Trim(tFecha.Text) & "  (TC: " & Trim(lTC.Caption) & ")"
        EncabezadoListado vsListado, aTexto, False
        vsListado.FileName = "Diferencias de Cambio"
         
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        
        vsListado.EndDoc
        'bCargarImpresion = False
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

