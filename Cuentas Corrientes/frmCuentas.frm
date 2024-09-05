VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCuentas 
   Caption         =   "Cuentas Corrientes"
   ClientHeight    =   7905
   ClientLeft      =   945
   ClientTop       =   1800
   ClientWidth     =   12870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCuentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   12870
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   2055
      Left            =   120
      TabIndex        =   18
      Top             =   780
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
      _ExtentY        =   3625
      _StockProps     =   229
      BorderStyle     =   1
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
      Zoom            =   60
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   660
      Left            =   120
      TabIndex        =   21
      Top             =   60
      Width           =   12675
      Begin VB.CheckBox chSaldoI 
         Caption         =   "c/&Saldo Inicial"
         Height          =   255
         Left            =   7560
         TabIndex        =   6
         Top             =   270
         Width           =   1455
      End
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1020
         MaxLength       =   40
         TabIndex        =   1
         Top             =   240
         Width           =   3195
      End
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6300
         MaxLength       =   12
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox tDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4920
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Al:"
         Height          =   255
         Left            =   6060
         TabIndex        =   4
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   4380
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   7275
      TabIndex        =   19
      Top             =   6240
      Width           =   7335
      Begin VB.CommandButton bExportar 
         Height          =   310
         Left            =   5340
         Picture         =   "frmCuentas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Exportar"
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmCuentas.frx":074C
         Height          =   310
         Left            =   4440
         Picture         =   "frmCuentas.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmCuentas.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmCuentas.frx":1242
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmCuentas.frx":132C
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmCuentas.frx":1416
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmCuentas.frx":1650
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "frmCuentas.frx":1752
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5700
         Picture         =   "frmCuentas.frx":1B18
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmCuentas.frx":1C1A
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmCuentas.frx":1F1C
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmCuentas.frx":225E
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmCuentas.frx":2560
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6240
         TabIndex        =   24
         Top             =   120
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
      TabIndex        =   20
      Top             =   7650
      Width           =   12870
      _ExtentX        =   22701
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14499
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3255
      Left            =   1200
      TabIndex        =   22
      Top             =   1920
      Width           =   4215
      _ExtentX        =   7435
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsTotal 
      Height          =   3255
      Left            =   6240
      TabIndex        =   25
      Top             =   1560
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
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
   Begin MSComDlg.CommonDialog ctlDlg 
      Left            =   10740
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MnuCol 
      Caption         =   "Columnas"
      Visible         =   0   'False
      Begin VB.Menu MnuOpComentarios 
         Caption         =   "Ver Comentarios"
      End
      Begin VB.Menu MnuOpPlazos 
         Caption         =   "Ver Plazos"
      End
      Begin VB.Menu MnuFactura 
         Caption         =   "Ir a Factura"
      End
      Begin VB.Menu MnuT2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuT1 
         Caption         =   "Mostrar/Ocultar columnas"
         Begin VB.Menu Col 
            Caption         =   "C1"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsAux As rdoResultset
Private aTexto As String
Dim aDisponibilidades As String

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bExportar_Click()
        
On Error GoTo errCancel
    
    With ctlDlg
        .CancelError = True
        
        .FileName = "CtasCorrientes"
        
        .Filter = "Libro de Microsoft Exel|*.xls|" & _
                     "Texto (delimitado por tabulaciones)|*.txt|" & "Texto (delimitado por comas)|*.txt"
        
        .ShowSave
        
        'Confirma exportar el contenido de la lista al archivo:
        If MsgBox("Confirma exportar el contenido de la lista al archivo: " & .FileName, vbQuestion + vbYesNo) = vbYes Then
        
            On Error GoTo errSaving
            Screen.MousePointer = 11
            Me.Refresh
            DoEvents
            
            Dim mSSetting As SaveLoadSettings
            
            Select Case .FilterIndex
                Case 1, 2: mSSetting = flexFileTabText
                Case 3: mSSetting = flexFileCommaText
            End Select
            
            'vsConsulta.MergeCells = flexMergeNever
            vsConsulta.SaveGrid .FileName, mSSetting ', True

            Screen.MousePointer = 0
        End If
        
    End With
    
errCancel:
    Screen.MousePointer = 0
    Exit Sub
errSaving:
     clsGeneral.OcurrioError "Error al exportar el contenido de la lista.", Err.Description
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


Private Sub chSaldoI_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
    Else
        AccionImprimir
        vsListado.ZOrder 0
    End If
End Sub

Private Sub Col_Click(Index As Integer)
    vsConsulta.ColHidden(Index) = Col(Index).Checked
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
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

Private Sub Form_Load()
    
    On Error Resume Next
    ObtengoSeteoForm Me, Me.Left, Me.Top, Me.Width, Me.Height
    pbProgreso.Value = 0
    InicializoGrilla
    vsConsulta.ZOrder 0
    
    picBotones.BorderStyle = vbBSNone
    PropiedadesImpresion
    
    FechaDelServidor
    tDesde.Text = Format(PrimerDia(gFechaServidor), "dd/mm/yyyy")
    tHasta.Text = Format(UltimoDia(gFechaServidor), "dd/mm/yyyy")
    
    'cTipo.AddItem "Contado": cTipo.ItemData(cTipo.NewIndex) = 0
    'cTipo.AddItem "Crédito": cTipo.ItemData(cTipo.NewIndex) = 1
    'cTipo.ListIndex = 1
    
    chSaldoI.Value = vbChecked
    
    For I = 0 To vsConsulta.Cols - 2
        If I <> 0 Then Load Col(I)
        Col.Item(I).Caption = vsConsulta.Cell(flexcpText, 0, I)
    Next
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 50
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next

    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub Label1_Click()
    Foco tDesde
End Sub

Private Sub Label2_Click()
    Foco tProveedor
End Sub

Private Sub Label3_Click()
    Foco tHasta
End Sub

Private Sub MnuFactura_Click()
On Error GoTo errMF

    If Not IsNumeric(vsConsulta.Cell(flexcpText, vsConsulta.Row, 0)) Then Exit Sub
    Dim aTipo As Integer, aDoc As Long
    
    aTipo = Val(vsConsulta.Cell(flexcpData, vsConsulta.Row, 2))
    aDoc = vsConsulta.Cell(flexcpValue, vsConsulta.Row, 0)
    
    If Trim(aDoc) <> "" Then
        Select Case aTipo
            Case TipoDocumento.CompraCredito, TipoDocumento.CompraNotaCredito, TipoDocumento.CompraContado, TipoDocumento.CompraNotaDevolucion
                    EjecutarApp prmPathApp & "Ingreso de Facturas.exe", CStr(aDoc)
            
            Case TipoDocumento.CompraReciboDePago
                    EjecutarApp prmPathApp & "Recibos de Pago.exe", CStr(aDoc)
        End Select
    
    End If
    
    
errMF:
End Sub

Private Sub MnuOpComentarios_Click()
    On Error GoTo errCom
    Dim rs1 As rdoResultset, aComentarios As String, aTitle As String
    
    With vsConsulta
        If Not IsNumeric(.Cell(flexcpText, .Row, 0)) Then Exit Sub
        
        Screen.MousePointer = 11
        cons = "Select * from Compra Where ComCodigo = " & .Cell(flexcpValue, .Row, 0)
        Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rs1.EOF Then
            If Not IsNull(rs1!ComComentario) Then aComentarios = Trim(rs1!ComComentario)
        End If
        rs1.Close
        
        aTitle = "Comentarios Documento " & Trim(.Cell(flexcpText, .Row, 2))
        If Trim(aComentarios) = "" Then
            MsgBox "No hay comentarios ingresados.", vbExclamation, aTitle
        Else
            MsgBox aComentarios, vbInformation, aTitle
        End If
    End With
        
    Screen.MousePointer = 0
    Exit Sub
errCom:
    clsGeneral.OcurrioError "Error al buscar los comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuOpPlazos_Click()
On Error GoTo errCom
    Dim rs1 As rdoResultset, aCompra As Long, aTitle As String, bHay As Boolean
    
    With vsConsulta
        If Not IsNumeric(.Cell(flexcpText, .Row, 0)) Then Exit Sub
        
        Screen.MousePointer = 11
        
        If Trim(.Cell(flexcpText, .Row, 9)) <> "" Then
            aCompra = .Cell(flexcpData, .Row, 9)
            aTitle = "Plazos Documento " & Trim(.Cell(flexcpText, .Row, 9))
        Else
            aCompra = .Cell(flexcpValue, .Row, 0)
            aTitle = "Plazos Documento " & Trim(.Cell(flexcpText, .Row, 2))
        End If
        
        cons = "Select * from CompraVencimiento Where CVeIDCompra = " & aCompra
        Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rs1.EOF Then bHay = True Else bHay = False
        rs1.Close
        
        If Not bHay Then
            MsgBox "No hay plazos ingresados.", vbExclamation, aTitle
        Else
            EjecutarApp prmPathApp & "Vencimiento de Pagos.exe", CStr(aCompra)
        End If
    End With
        
    Screen.MousePointer = 0
    Exit Sub
errCom:
    clsGeneral.OcurrioError "Error al buscar los plazos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tDesde_GotFocus()
    With tDesde: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    If vsConsulta.Rows > 1 Then
        'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
        Screen.MousePointer = 11
        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        EncabezadoListado vsListado, "Cuentas Corrientes - " & Trim(tProveedor.Text) & " Desde " & Trim(tDesde.Text) & " al " & Trim(tHasta.Text), False
        vsListado.FileName = "Cuentas Corrientes"
        
        vsConsulta.ExtendLastCol = False
        vsListado.RenderControl = vsConsulta.hwnd
        vsConsulta.ExtendLastCol = True
        
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
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function TInicializoListado() As Boolean
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    With vsListado
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Function
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    EncabezadoListado vsListado, "Cuentas Corrientes - Desde " & Trim(tDesde.Text) & " al " & Trim(tHasta.Text), False
    vsListado.FileName = "Cuentas Corrientes"
    vsConsulta.ExtendLastCol = False
    
    TInicializoListado = True
    Exit Function
    
errImprimir:
    TInicializoListado = False
    Screen.MousePointer = 0
End Function

Private Sub TCargoProveedor(Nombre As String)


    vsListado.Paragraph = Trim(Nombre)
    vsListado.RenderControl = vsConsulta.hwnd
    
    vsListado.Paragraph = " "
    
End Sub

Private Function TFinListado(Optional Imprimir As Boolean = False) As Boolean
    
    vsListado.EndDoc
    vsConsulta.ExtendLastCol = True
    
    If Imprimir Then
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    Screen.MousePointer = 0
    TFinListado = True
    Exit Function

errImprimir:
    TFinListado = False
    Screen.MousePointer = 0
End Function

Private Sub PropiedadesImpresion()
  
  With vsListado
        .PaperSize = vbPRPSLetter
        .PhysicalPage = True
        .Orientation = orLandscape
        
        .PreviewMode = pmScreen
        
        .PreviewPage = 1
        .FontName = "Tahoma": .FontSize = 10: .FontBold = False: .FontItalic = False
        .TextAlign = 0: .PageBorder = 3
        .Columns = 1
        .TableBorder = tbBoxRows
        .Zoom = 100
        .MarginBottom = 750: .MarginTop = 750
        .MarginRight = 350
    End With

End Sub


Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tHasta
End Sub

Private Sub tDesde_LostFocus()
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, "dd/mm/yyyy")
End Sub

Private Sub tHasta_GotFocus()
    With tHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco chSaldoI
End Sub

Private Sub tHasta_LostFocus()
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, "dd/mm/yyyy")
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = 0
End Sub

Private Sub tProveedor_GotFocus()
    With tProveedor: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then Foco tDesde: Exit Sub
        Screen.MousePointer = 11
        Dim aQ As Long, aIdProveedor As Long, aTexto As String
        
        aQ = 0
        cons = "Select PClCodigo, PClFantasia as Nombre, PClNombre as 'Razón Social' from ProveedorCliente " _
                & " Where PClNombre like '" & Trim(tProveedor.Text) & "%' Or PClFantasia like '" & Trim(tProveedor.Text) & "%'"
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            aQ = 1: aIdProveedor = rsAux!PClCodigo: aTexto = Trim(rsAux!Nombre)
            rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
        End If
        rsAux.Close
        
        Select Case aQ
            Case 0:
                    MsgBox "No existe una empresa para el con el nombre ingresado.", vbExclamation, "No existe Empresa"
            
            Case 1:
                    tProveedor.Text = aTexto
                    tProveedor.Tag = aIdProveedor
                    Foco tDesde
        
            Case 2:
                    Dim aLista As New clsListadeAyuda
                    Dim mIdSel As Long
                    mIdSel = aLista.ActivarAyuda(cBase, cons, 5500, 1, "Lista de Proveedores")
                    If mIdSel <> 0 Then
                        tProveedor.Text = Trim(aLista.RetornoDatoSeleccionado(1))
                        tProveedor.Tag = aLista.RetornoDatoSeleccionado(0)
                        
                        Foco tDesde
                    Else
                        tProveedor.Text = ""
                    End If
                    Set aLista = Nothing
        End Select
        
        If Val(tProveedor.Tag) > 0 Then BuscoComentarios CLng(tProveedor.Tag)
    End If
    
    Screen.MousePointer = 0
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al procesar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoComentarios(idCliente As Long)
    On Error GoTo errCom
    Dim bHay As Boolean: bHay = False
    
    cons = "Select * from Comentario " & _
               " Where ComCliente = " & idCliente & _
               " And ComTipo = " & paTComentarioCtaCorr
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True
    rsAux.Close
    
    If Not bHay Then Exit Sub
    Dim objComentarios As New clsCliente
    objComentarios.Comentarios idCliente, CInt(paTComentarioCtaCorr)
    Set objComentarios = Nothing
        
    Exit Sub
errCom:
    clsGeneral.OcurrioError "Error al buscar los comentarios de cuentas corrientes.", Err.Description
End Sub
Private Sub vsConsulta_DblClick()
    On Error GoTo errDC
    With vsConsulta
        If .Rows = 1 Then Exit Sub
        If .Cell(flexcpText, .Row, 8) = "" Then Exit Sub
        For I = 1 To .Rows - 1
            If CStr(.Cell(flexcpValue, I, 0)) = CStr(.Cell(flexcpData, .Row, 8)) Then
                .Select I, 0, , .Cols - 1
                I = .CellTop
                Exit For
            End If
        Next
    End With
errDC:
End Sub


Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        For I = 0 To vsConsulta.Cols - 2
            If vsConsulta.ColHidden(I) Then Col.Item(I).Checked = False Else Col.Item(I).Checked = True
        Next
        PopupMenu MnuCol
    End If
    
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionConsultar()

    On Error GoTo ErrCDML
    If Not ValidoDatos Then Exit Sub
    
    Dim bHay As Boolean
    
    Screen.MousePointer = 11
    chVista.Value = vbUnchecked
    vsConsulta.Rows = 1: vsConsulta.Refresh
    
    If Val(tProveedor.Tag) <> 0 Then
        'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
        pbProgreso.Value = 0
        cons = "Select Count(*) from Compra Left Outer Join CompraPago On ComCodigo = CPaDocQSalda" _
               & " Where ComFecha Between '" & Format(tDesde.Text, sqlFormatoF) & "' And '" & Format(tHasta.Text, sqlFormatoF) & "'" _
               & " And ComProveedor = " & Val(tProveedor.Tag)
        
        '3/5/2003 Cargo Facturas y Notas de Créditos  + todos los docs del proveedor q están asignados a Acr.Varios
        '               Por problema de nota de crédito que devuelve mercaderia y plata (se hace entrada de caja x la guita y eso cierra la cta. corr)
        cons = cons & " And ( ComTipoDocumento IN (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraReciboDePago & ")" & _
                                    " Or ComCodigo In (Select GSrIDCompra from GastoSubrubro Where GSrIDCompra = ComCodigo And GSrIDSubrubro = " & paSubRubroAcreedores & ") )"
        
        
        bHay = True
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            If rsAux(0) = 0 Then bHay = False Else pbProgreso.Max = rsAux(0)
        End If
        rsAux.Close
        
        If Not bHay Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            Screen.MousePointer = 0: Exit Sub
        End If
    
        CargoDatosProveedor Val(tProveedor.Tag)
        '-----------------------------------------------------------------------------------------------------------------
    Else
        
        vsTotal.Rows = 1
        pbProgreso.Value = 0
        cons = "Select Count(Distinct(ComProveedor)) from Compra " _
               & " Where ComFecha Between '" & Format(tDesde.Text, sqlFormatoF) & "' And '" & Format(tHasta.Text, sqlFormatoF) & "'"
               
'        If cTipo.ListIndex <> -1 Then
'            Select Case cTipo.ItemData(cTipo.ListIndex)
'                Case 0: cons = cons & " And ComTipoDocumento Not IN (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraReciboDePago & ")"
'                Case 1: cons = cons & " And ComTipoDocumento IN (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraReciboDePago & ")"
'            End Select
'        End If
        cons = cons & " And ( ComTipoDocumento IN (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraReciboDePago & ")" & _
                                " Or ComCodigo In (Select GSrIDCompra from GastoSubrubro Where GSrIDCompra = ComCodigo And GSrIDSubrubro = " & paSubRubroAcreedores & ") )"
    
        bHay = True
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            If rsAux(0) = 0 Then bHay = False Else pbProgreso.Max = rsAux(0)
        End If
        rsAux.Close
        
        If Not bHay Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            Screen.MousePointer = 0: Exit Sub
        End If
    
        Dim rsPro As rdoResultset
        Dim aTProveedor As String
        vsConsulta.Redraw = False
        
        If Not TInicializoListado Then GoTo ErrCDML
        
        cons = "Select Distinct(ComProveedor), PClNombre, PClFantasia from Compra, ProveedorCliente " _
               & " Where ComFecha Between '" & Format(tDesde.Text, sqlFormatoF) & "' And '" & Format(tHasta.Text, sqlFormatoF) & "'" _
               & " And ComProveedor = PClCodigo"
        
    '    If cTipo.ListIndex <> -1 Then
    '        Select Case cTipo.ItemData(cTipo.ListIndex)
    '            Case 0: cons = cons & " And ComTipoDocumento Not IN (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraReciboDePago & ")"
    '            Case 1: cons = cons & " And ComTipoDocumento IN (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraReciboDePago & ")"
    '        End Select
    '    End If
        cons = cons & " And ( ComTipoDocumento IN (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraReciboDePago & ")" & _
                                    " Or ComCodigo In (Select GSrIDCompra from GastoSubrubro Where GSrIDCompra = ComCodigo And GSrIDSubrubro = " & paSubRubroAcreedores & ") )"
    
        cons = cons & " Order by PClNombre"
        
        Set rsPro = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsPro.EOF
            If Not IsNull(rsPro!PClNombre) Then aTProveedor = rsPro!PClNombre Else aTProveedor = rsPro!PClFantasia
            
            CargoDatosProveedor rsPro!ComProveedor, Proveedor:=aTProveedor, bTodos:=True
                
            If vsConsulta.Rows > 1 Then TCargoProveedor aTProveedor
            
            pbProgreso.Value = pbProgreso.Value + 1
            rsPro.MoveNext
        Loop
        rsPro.Close
        
        With vsTotal
         If .Rows > 1 Then
             .SubtotalPosition = flexSTBelow
             .Subtotal flexSTSum, -1, 1, , Colores.Rojo, Colores.Blanco, True, "Totales"
             .Subtotal flexSTSum, -1, 2
             .Subtotal flexSTSum, -1, 3
    
             vsListado.NewPage
             vsListado.Paragraph = "RESUMEN FINAL "
             vsListado.RenderControl = .hwnd
         End If
        End With
        
        TFinListado
        vsConsulta.Rows = 1
        vsConsulta.Redraw = True
        vsListado.ZOrder 0
        pbProgreso.Value = 0
    End If
    
    Screen.MousePointer = 0
Exit Sub
    
ErrCDML:
    vsConsulta.Redraw = True
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosProveedor(idProveedor As Long, Optional Proveedor As String = "", Optional bTodos As Boolean = False)
 
Dim rs1 As rdoResultset, aDocGasto As String
Dim aPagosP As Currency, aPagosD As Currency
Dim aValor As Long
Dim mImporte As Currency, mVaADH As String
Dim varPOR As Long

    On Error GoTo ErrCDML
       
    Screen.MousePointer = 11
    aPagosP = 0: aPagosD = 0
    vsConsulta.Rows = 1
    
    If Not bTodos Then pbProgreso.Value = 0

    cons = "Select * from Compra Left Outer Join CompraPago On ComCodigo = CPaDocQSalda" _
           & " Where ComFecha Between '" & Format(tDesde.Text, sqlFormatoF) & "' And '" & Format(tHasta.Text, sqlFormatoF) & "'" _
           & " And ComProveedor = " & idProveedor
    
    cons = cons & "And ( ComTipoDocumento IN (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraReciboDePago & ")" & _
                                    " Or ComCodigo In (Select GSrIDCompra from GastoSubrubro Where GSrIDCompra = ComCodigo And GSrIDSubrubro = " & paSubRubroAcreedores & ") )"
        
    cons = cons & " Order by ComCodigo"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If rsAux.EOF Then rsAux.Close: Exit Sub
    
    If chSaldoI.Value = vbChecked Then CargoSaldoInicial idProveedor
    
    With vsConsulta
    Do While Not rsAux.EOF
        If Not bTodos Then pbProgreso.Value = pbProgreso.Value + 1
        
        '1) Si el Tipo de documento es un recibo de pago y lo que pague es una DC no Va !!!!
        '2) Si el Tipo de documento es un Credito y es una DC no va !!                                  (Carlos Sab 3 del Jun)
        If rsAux!ComTipoDocumento = TipoDocumento.CompraCredito And Not IsNull(rsAux!ComDCDe) Then GoTo bpContinuar
        
        If rsAux!ComTipoDocumento = TipoDocumento.CompraReciboDePago Then
            Dim rsDC As rdoResultset, bDC As Boolean
            bDC = False
            cons = "Select * from Compra Where ComCodigo = " & rsAux!CPaDocASaldar
            Set rsDC = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rsDC.EOF Then If Not IsNull(rsDC!ComDCDe) Then bDC = True
            rsDC.Close
            If bDC Then GoTo bpContinuar
        End If
        
        'Sin Autorizar ComUsrAutoriza <> Null And ComVerificado = 0     Cambiar 11/2002
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!ComCodigo, "#,##0")
        If Not IsNull(rsAux!ComUsrAutoriza) Then
            If Not IsNull(rsAux!ComVerificado) Then
                If rsAux!ComVerificado = 0 Then .Cell(flexcpForeColor, .Rows - 1, 0, , 7) = Colores.Rojo
            Else
                .Cell(flexcpForeColor, .Rows - 1, 0, , 7) = Colores.Rojo
            End If
        End If
        
        .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!ComFecha, "dd/mm/yy")
                
        aTexto = RetornoNombreDocumento(rsAux!ComTipoDocumento, Abreviacion:=True) & " "
        If Not IsNull(rsAux!ComSerie) Then aTexto = aTexto & Trim(rsAux!ComSerie) & " "
        If Not IsNull(rsAux!ComNumero) Then aTexto = aTexto & rsAux!ComNumero
        .Cell(flexcpText, .Rows - 1, 2) = aTexto
        aValor = rsAux!ComTipoDocumento
        .Cell(flexcpData, .Rows - 1, 2) = aValor
                    
        mVaADH = VaAlDebeOHaber(rsAux!ComTipoDocumento)
        varPOR = IIf(rsAux!ComTipoDocumento = TipoDocumento.CompraReciboDePago, -1, 1)
        
        If rsAux!ComMoneda = paMonedaPesos Then
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!ComImporte * varPOR, FormatoMonedaP)
            If Not IsNull(rsAux!ComCofis) Then .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!ComCofis * varPOR, FormatoMonedaP)
            If Not IsNull(rsAux!ComIva) Then .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!ComIva * varPOR, FormatoMonedaP)
            
            .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpValue, .Rows - 1, 3) + .Cell(flexcpValue, .Rows - 1, 4) + .Cell(flexcpValue, .Rows - 1, 5), FormatoMonedaP)
            If Not IsNull(rsAux!CPaAmortizacion) Then mImporte = rsAux!CPaAmortizacion Else mImporte = .Cell(flexcpValue, .Rows - 1, 6)
            
            If rsAux!ComTipoDocumento <> TipoDocumento.CompraReciboDePago Then mImporte = Abs(mImporte)
            'mImporte = Abs(mImporte)    'el 19/10/04 saco ABS porque CArlos inrgeso un RdeP en Negativo (devolucion de guita)
            
            Select Case mVaADH
                Case "D": .Cell(flexcpText, .Rows - 1, 10) = Format(mImporte, FormatoMonedaP)
                Case "H": .Cell(flexcpText, .Rows - 1, 11) = Format(mImporte, FormatoMonedaP)
            End Select
            
        Else
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!ComImporte * rsAux!ComTC * varPOR, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 7) = Format(rsAux!ComImporte * varPOR, FormatoMonedaP)
            If Not IsNull(rsAux!ComIva) Then
                .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!ComIva * rsAux!ComTC * varPOR, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 7) = Format(.Cell(flexcpValue, .Rows - 1, 7) + rsAux!ComIva * varPOR, FormatoMonedaP)
            End If
            If Not IsNull(rsAux!ComCofis) Then
                .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!ComCofis * rsAux!ComTC, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 7) = Format(.Cell(flexcpValue, .Rows - 1, 7) + rsAux!ComCofis, FormatoMonedaP)
            End If
            .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpValue, .Rows - 1, 3) + .Cell(flexcpValue, .Rows - 1, 4) + .Cell(flexcpValue, .Rows - 1, 5), FormatoMonedaP)
            
            If Not IsNull(rsAux!CPaAmortizacion) Then mImporte = rsAux!CPaAmortizacion Else mImporte = .Cell(flexcpValue, .Rows - 1, 7)
            
            'mImporte = Abs(mImporte)   'el 19/10/04 saco ABS porque CArlos inrgeso un RdeP en Negativo (devolucion de guita)
            If rsAux!ComTipoDocumento <> TipoDocumento.CompraReciboDePago Then mImporte = Abs(mImporte)
            
            Select Case mVaADH
                Case "D": mImporte = mImporte
                Case "H": mImporte = mImporte * -1
            End Select
            
            .Cell(flexcpText, .Rows - 1, 12) = Format(mImporte, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 13) = Format(mImporte * rsAux!ComTC, FormatoMonedaP)
        End If

        If Not IsNull(rsAux!ComSaldo) Then
            If rsAux!ComTipoDocumento = TipoDocumento.CompraCredito Then .Cell(flexcpText, .Rows - 1, 8) = "0.00"
            If rsAux!ComSaldo > 0 Then .Cell(flexcpText, .Rows - 1, 8) = Format(rsAux!ComSaldo, FormatoMonedaP)
        End If
        
        If Not IsNull(rsAux!CPaDocASaldar) Then
            aDocGasto = rsAux!CPaDocASaldar: .Cell(flexcpData, .Rows - 1, 9) = aDocGasto
            aDocGasto = ""
            cons = "Select * from Compra Where ComCodigo = " & rsAux!CPaDocASaldar
            Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rs1.EOF Then
                If Not IsNull(rs1!ComSerie) Then aDocGasto = Trim(rs1!ComSerie) & " "
                If Not IsNull(rs1!ComNumero) Then aDocGasto = aDocGasto & rs1!ComNumero
            End If
            rs1.Close
            .Cell(flexcpText, .Rows - 1, 9) = aDocGasto
        End If
        
        'Si es recibo de pago arreglo el D/H * -1
        If rsAux!ComTipoDocumento = TipoDocumento.CompraReciboDePago Then
            If .Cell(flexcpText, .Rows - 1, 11) <> "" Then
                aPagosP = aPagosP + (.Cell(flexcpValue, .Rows - 1, 11)) 'SACO ABS 19/10/2004
            End If
            If .Cell(flexcpText, .Rows - 1, 12) <> "" Then
                aPagosD = aPagosD + (.Cell(flexcpValue, .Rows - 1, 12)) 'SACO ABS 19/10/2004
            End If
        End If
        
bpContinuar:
        rsAux.MoveNext
    Loop
    rsAux.Close
   
    If .Rows > 1 Then
        .Cell(flexcpBackColor, 1, 6, .Rows - 1, 7) = Colores.Inactivo: .Cell(flexcpBackColor, 1, 10, .Rows - 1, 13) = Colores.Inactivo
        
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 10, , Colores.Rojo, Colores.Blanco, True, "Totales"
        .Subtotal flexSTSum, -1, 11: .Subtotal flexSTSum, -1, 12: .Subtotal flexSTSum, -1, 13
        '.Subtotal flexSTSum, -1, 3: .Subtotal flexSTSum, -1, 6
                
        If bTodos Then       'Agrego en la Lista del Total
             vsTotal.AddItem ""
             vsTotal.Cell(flexcpText, vsTotal.Rows - 1, 0) = Trim(Proveedor)
             vsTotal.Cell(flexcpText, vsTotal.Rows - 1, 1) = Format(.Cell(flexcpValue, .Rows - 1, 10) - .Cell(flexcpValue, .Rows - 1, 11), FormatoMonedaP)
             vsTotal.Cell(flexcpText, vsTotal.Rows - 1, 2) = .Cell(flexcpText, .Rows - 1, 12)
             vsTotal.Cell(flexcpText, vsTotal.Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 13)
        End If
        
        'Saldo en Pesos----------------------------------------------------------------
        .AddItem "Saldo $"
        .Cell(flexcpText, .Rows - 1, 10) = Format(.Cell(flexcpValue, .Rows - 2, 10) - .Cell(flexcpValue, .Rows - 2, 11), FormatoMonedaP)
        .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco: .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
        
         If aPagosP <> 0 Or aPagosD <> 0 Then
             .AddItem "RPA"
             .Cell(flexcpText, .Rows - 1, 11) = Format(aPagosP, FormatoMonedaP)
             .Cell(flexcpText, .Rows - 1, 12) = Format(aPagosD, FormatoMonedaP)
             .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco: .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
             .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
         End If
    End If
    
    End With
    If Not bTodos Then pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
ErrCDML:
    clsGeneral.OcurrioError "Cargar proveedor " & Trim(Proveedor), Err.Description
    Screen.MousePointer = 0
End Sub

Private Function VaAAcreedores(idGasto As Long) As Boolean
On Error GoTo errFunc
    Dim rsAcr As rdoResultset
    
    VaAAcreedores = False
    
    cons = "Select * from GastoSubrubro Where GSrIDCompra =" & idGasto & " And GSrIDSubrubro = " & paSubRubroAcreedores
    Set rsAcr = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAcr.EOF Then VaAAcreedores = True
    rsAcr.Close

errFunc:
End Function

Private Sub CargoSaldoInicial(idProveedor As Long)
    On Error GoTo errSaldo
    'Busco si el proveedor tiene saldo ingresado
    Dim rsSal As rdoResultset
    cons = "Select * From SaldoCCte " & _
               " Where SCCProveedor = " & idProveedor & _
               " And SCCFecha <= '" & Format(tDesde.Text, sqlFormatoF) & "'" & _
               " Order by SCCFecha desc"
    Set rsSal = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsSal.EOF Then
        With vsConsulta
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsSal!SCCFecha, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 2) = "Saldo Inicial"
            If rsSal!SCCSaldoP >= 0 Then
                .Cell(flexcpText, .Rows - 1, 10) = Format(rsSal!SCCSaldoP, FormatoMonedaP)
            Else
                .Cell(flexcpText, .Rows - 1, 11) = Format(rsSal!SCCSaldoP, FormatoMonedaP)
            End If
            
            .Cell(flexcpText, .Rows - 1, 12) = Format(rsSal!SCCSaldoD, FormatoMonedaP)
            
            If rsSal!SCCSaldoD <> 0 Then
                Dim aTC As Currency
                aTC = TasadeCambio(paMonedaDolar, paMonedaPesos, rsSal!SCCFecha)
                .Cell(flexcpText, .Rows - 1, 13) = Format(rsSal!SCCSaldoD * aTC, FormatoMonedaP)
            End If
        End With
    End If
    rsSal.Close

errSaldo:
End Sub

Private Sub AccionLimpiar()
    tDesde.Text = ""
    tProveedor.Text = ""
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Function ValidoDatos() As Boolean
    On Error Resume Next
    ValidoDatos = False
    
    If Val(tProveedor.Tag) = 0 Then
        If MsgBox("Está seguro de consultar las cuentas corrientes de todos los proveedores.", vbQuestion + vbYesNo + vbDefaultButton2, "Todos los Proveedores ? ") = vbNo Then
        'MsgBox "Debe seleccionar un proveedor para realizar la consulta.", vbExclamation, "ATENCIÓN"
            Foco tProveedor: Exit Function
        End If
    End If
    
    If Not IsDate(tDesde.Text) Then
        MsgBox "La fecha desde ingresada para consultar no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    If Not IsDate(tHasta.Text) Then
        MsgBox "La fecha hasta ingresada para consultar no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tHasta: Exit Function
    End If
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then
        MsgBox "El rango de fechas ingresado para consultar no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    ValidoDatos = True
    
End Function

Private Sub InicializoGrilla()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = ">ID Gasto|<Fecha|<Documento|>Importe $|>Cofis $|>I.V.A. $|>Total $|>Total U$S|>Saldo Actual|<Salda Gasto|>Debe $ |>Haber $|>D/H U$S|>D/H $ (U$S)|"
        .ColWidth(1) = 750: .ColWidth(2) = 1300
        .ColWidth(3) = 1200: .ColWidth(4) = 950: .ColWidth(5) = 1000: .ColWidth(6) = 1200
        .ColWidth(7) = 1000 ': .ColWidth(7) = 1200
        .ColWidth(10) = 1300 ': .ColWidth(9) = 1200
        .ColWidth(11) = 1100 ': .ColWidth(11) = 1200
        .ColWidth(12) = 1100
        .ColWidth(13) = 1300
        
        .MergeCells = flexMergeRestrictRows ' flexMergeRestrictColumns
        .MergeCol(0) = True: .ColAlignment(0) = flexAlignLeftTop
        .MergeCol(1) = True: .ColAlignment(1) = flexAlignLeftTop
        .MergeCol(2) = True: .ColAlignment(2) = flexAlignLeftTop
        .MergeCol(3) = True: .ColAlignment(3) = flexAlignRightTop
        .MergeCol(4) = True: .ColAlignment(4) = flexAlignRightTop
        .MergeCol(5) = True: .ColAlignment(4) = flexAlignRightTop
        .MergeCol(6) = True: .ColAlignment(5) = flexAlignRightTop
        .MergeCol(7) = True: .ColAlignment(6) = flexAlignRightTop
        .MergeCol(8) = True: .ColAlignment(7) = flexAlignRightTop
        .WordWrap = False
    End With
      
    With vsTotal
        .Cols = 1: .Rows = 1:
        .FormatString = "<Proveedor|>Debe / Haber $|>D/H U$S|>D/H $ (U$S)|"
        .ColWidth(0) = 10100
        .ColWidth(1) = 2400: .ColWidth(2) = 1100: .ColWidth(3) = 1300
        
        .WordWrap = False
        .ExtendLastCol = False
    End With
End Sub


Private Function VaAlDebeOHaber(mTipoDoc As Integer) As String

        Select Case mTipoDoc
            Case TipoDocumento.CompraReciboDePago, TipoDocumento.CompraRecibo, TipoDocumento.CompraSalidaCaja, _
                       TipoDocumento.CompraNotaCredito, TipoDocumento.CompraNotaDevolucion
                    
                            VaAlDebeOHaber = "H"
        
            Case Else   'TipoDocumento.CompraCredito,TipoDocumento.CompraEntradaCaja
                            VaAlDebeOHaber = "D"
        End Select
        
End Function
