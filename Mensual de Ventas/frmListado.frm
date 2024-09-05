VERSION 5.00
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListado 
   Caption         =   "Mensual de Ventas por Artículo"
   ClientHeight    =   7440
   ClientLeft      =   1770
   ClientTop       =   2130
   ClientWidth     =   12630
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
   ScaleHeight     =   7440
   ScaleWidth      =   12630
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4335
      Left            =   1200
      TabIndex        =   14
      Top             =   1680
      Width           =   6735
      _Version        =   196608
      _ExtentX        =   11880
      _ExtentY        =   7646
      _StockProps     =   229
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Zoom            =   70
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11595
      TabIndex        =   16
      Top             =   6600
      Width           =   11655
      Begin VB.CommandButton bExportar 
         Height          =   310
         Left            =   5280
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Exportar"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmListado.frx":0986
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":0C88
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":0FCA
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5640
         Picture         =   "frmListado.frx":12CC
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "frmListado.frx":13CE
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmListado.frx":1794
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":1896
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmListado.frx":1AD0
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmListado.frx":1BBA
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":1CA4
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":211E
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":2220
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   310
      End
   End
   Begin VB.Frame Shape1 
      Caption         =   "Filtro de Datos"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   10575
      Begin AACombo99.AACombo cGrupo 
         Height          =   315
         Left            =   4920
         TabIndex        =   11
         Top             =   540
         Width           =   2295
         _ExtentX        =   4048
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
         Left            =   4920
         TabIndex        =   5
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   45
         TabIndex        =   9
         Top             =   540
         Width           =   3255
      End
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         MaxLength       =   12
         TabIndex        =   3
         Top             =   210
         Width           =   1095
      End
      Begin VB.TextBox tDesde 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   12
         TabIndex        =   1
         Top             =   210
         Width           =   1095
      End
      Begin AACombo99.AACombo cSucursal 
         Height          =   315
         Left            =   8100
         TabIndex        =   7
         Top             =   180
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
      Begin AACombo99.AACombo cboTipoArticulo 
         Height          =   315
         Left            =   8100
         TabIndex        =   13
         Top             =   540
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   7380
         TabIndex        =   12
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Sucursal:"
         Height          =   255
         Left            =   7380
         TabIndex        =   6
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   570
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Grupo:"
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   210
         Width           =   855
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   5760
      TabIndex        =   29
      Top             =   1320
      Width           =   4215
      _ExtentX        =   7435
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
   Begin MSComDlg.CommonDialog ctlDlg 
      Left            =   9660
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private QArticuloE As rdoQuery, QArticuloN As rdoQuery
Private RsArticuloE As rdoResultset, RsArticuloN As rdoResultset

Private Sub AccionConsultar()
On Error GoTo ErrBC

Dim Meses As Integer, aArticulo As Long

    If Not ValidoDatos Then Exit Sub
    vsConsulta.Rows = 1: vsConsulta.Refresh
    
    Screen.MousePointer = vbHourglass
    Meses = EncabezadoLista(tDesde.Text, tHasta.Text) + 1
    
    Dim promdias As Integer
    Dim iMenorValor As Integer
    
    Cons = "Select Mes = DatePart(mm,AArFecha), Ano = DatePart(yy,AArFecha), ArtCodigo, ArtNombre, Cantidad = (Sum(AArCantidadNCo) + Sum(AArCantidadNCr) + Sum(AArCantidadECo) + Sum(AArCantidadECr)), Count(Distinct(AArFecha)) Qdias " _
            & " From AcumuladoArticulo, Articulo" _
            & " Where AArArticulo = ArtID" _
            & " And AArFEcha Between " & Format(tDesde.Text, "'mm/dd/yyyy'") & " And " & Format(tHasta.Text, "'mm/dd/yyyy'")
            
    If cMoneda.ListIndex <> -1 Then Cons = Cons & " And AArMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    If cSucursal.ListIndex <> -1 Then Cons = Cons & " And AArSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And AArArticulo  = " & Val(tArticulo.Tag)
    
    If cGrupo.ListIndex <> -1 Then Cons = Cons & " And AArArticulo In (Select AGrArticulo From ArticuloGrupo Where AGrGrupo = " & cGrupo.ItemData(cGrupo.ListIndex) & ")"
    If cboTipoArticulo.ListIndex > -1 Then Cons = Cons & " And ArtTipo In (" & cboTipoArticulo.ItemData(cboTipoArticulo.ListIndex) & ")"
    
    Cons = Cons & " Group by ArtNombre, AArArticulo, ArtCodigo, DatePart(mm,AArFecha), DatePart(yy,AArFecha)"
    Cons = Cons & " Order by ArtNombre"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        vsConsulta.Redraw = False
        aArticulo = 0
        Do While Not RsAux.EOF
            If RsAux!ArtCodigo <> aArticulo Then aArticulo = RsAux!ArtCodigo
            
            With vsConsulta
                .AddItem ""
                .Cell(flexcpData, .Rows - 1, 0) = aArticulo
                
                .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!ArtCodigo, "(#,000,000) ") & Trim(RsAux!ArtNombre)
            
                For I = 1 To Meses: .Cell(flexcpText, .Rows - 1, I) = 0: Next
                
                Do While RsAux!ArtCodigo = aArticulo
                    For I = 1 To Meses
                        If .Cell(flexcpText, 0, I) = Format(RsAux!Mes & "/" & RsAux!Ano, "mm/yy") Then
                        'le agrego la fecha en el tag para preguntar abajo.
                            If .Cell(flexcpData, 0, I) = "" Then .Cell(flexcpData, 0, I) = "01/" & Format(RsAux!Mes & "/" & RsAux!Ano, "mm/yyyy")
                            .Cell(flexcpText, .Rows - 1, I) = RsAux!Cantidad
                            .Cell(flexcpData, .Rows - 1, I) = CStr(RsAux("qdias"))
                            Exit For
                        End If
                    Next
                    RsAux.MoveNext
                    If RsAux.EOF Then Exit Do
                Loop
            End With
        Loop
                   
    Else
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbExclamation, "No hay datos"
    End If
    RsAux.Close
    
    Dim iMesesProm As Integer
    
    
    If vsConsulta.Rows > 1 Then
        Dim mTotalM As Currency
        Dim aCol
        Dim aTotal As Long
        For I = 1 To vsConsulta.Rows - 1
            aTotal = 0
            
            'DETERMINO EL PROMEDIO POR FILA. pusimos las siguiente condiciones no se suma el mes actual, tampoco se consideran los meses
            'en que la cantidad sea < 1.
            promdias = 0
            iMesesProm = 0
            iMenorValor = 10000
            
            For aCol = 1 To vsConsulta.Cols - 3
                If Val(vsConsulta.Cell(flexcpText, I, aCol)) > 0 And CDate(vsConsulta.Cell(flexcpData, 0, aCol)) <> CDate("01/" & Format(Date, "mm/yyyy")) Then
                    iMesesProm = iMesesProm + 1
                    promdias = promdias + Val(vsConsulta.Cell(flexcpData, I, aCol))
                    If iMenorValor > Val(vsConsulta.Cell(flexcpData, I, aCol)) Then
                        iMenorValor = Val(vsConsulta.Cell(flexcpData, I, aCol))
                    End If
                End If
            Next
            
            If promdias > 0 Then
                If iMenorValor <> 10000 And iMesesProm > 1 Then
                    promdias = promdias - iMenorValor
                    iMesesProm = iMesesProm - 1
                End If
                promdias = (promdias \ iMesesProm) * 0.75
            End If
            
            For aCol = 1 To vsConsulta.Cols - 3
                aTotal = aTotal + vsConsulta.Cell(flexcpValue, I, aCol)
                If Val(vsConsulta.Cell(flexcpData, I, aCol)) < promdias And Val(vsConsulta.Cell(flexcpText, I, aCol)) > 0 Then
                    vsConsulta.Cell(flexcpForeColor, I, aCol) = &HC0&
                    vsConsulta.Cell(flexcpFontBold, I, aCol) = True
                End If
            Next
            vsConsulta.Cell(flexcpText, I, vsConsulta.Cols - 2) = aTotal
        Next
        vsConsulta.AddItem ""
        For aCol = 1 To vsConsulta.Cols - 2
            aTotal = 0
            For I = vsConsulta.FixedRows To vsConsulta.Rows - 2
                aTotal = aTotal + vsConsulta.Cell(flexcpValue, I, aCol)
            Next
            vsConsulta.Cell(flexcpText, vsConsulta.Rows - 1, aCol) = aTotal
        Next

        vsConsulta.Cell(flexcpBackColor, 1, vsConsulta.Cols - 2, vsConsulta.Rows - 1) = Colores.clNaranja
        vsConsulta.Cell(flexcpBackColor, vsConsulta.Rows - 1, 1, , vsConsulta.Cols - 2) = Colores.clNaranja
        'vsConsulta.Select 1, 1
        'vsConsulta.Sort = flexSortGenericAscending
    End If
    
    vsConsulta.Redraw = True
    Screen.MousePointer = vbDefault
    Exit Sub

ErrBC:
    clsGeneral.OcurrioError "Ocurrió un error al consultar.", Err.Description
    cBase.QueryTimeout = 15
    vsConsulta.Redraw = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bExportar_Click()
On Error GoTo errCancel
    
    With ctlDlg
        .CancelError = True
        
        .FileName = "Mensual_Ventas_Articulo"
        
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
                Case 1: mSSetting = flexFileTabText
                Case 2: mSSetting = flexFileTabText
                Case 3: mSSetting = flexFileCommaText
            End Select
            
            vsConsulta.SaveGrid .FileName, mSSetting, True
                
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

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub cboTipoArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub cGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cboTipoArticulo.SetFocus
End Sub

Private Sub chVista_Click()
    
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
    Else
        AccionImprimir
        vsListado.ZOrder 0
    End If
    Me.Refresh

End Sub

Private Sub cSucursal_GotFocus()
    cSucursal.SelStart = 0: cSucursal.SelLength = Len(cSucursal.Text)
End Sub

Private Sub cSucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tArticulo
End Sub

Private Sub cSucursal_LostFocus()
    cSucursal.SelLength = 0
End Sub

Private Sub Form_Activate()
    Me.Refresh
    Screen.MousePointer = vbDefault
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
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyC: AccionConfigurar
            
            Case vbKeyX: Unload Me
        End Select
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me
    CargoCombos
    
    FechaDelServidor
    tDesde.Text = Format(PrimerDia(gFechaServidor), "d-Mmm yyyy")
    tHasta.Text = Format(gFechaServidor, "d-Mmm yyyy")
    
    InicializoGrilla
    
    With vsListado
        .PreviewPage = 1: .Orientation = orPortrait
        .PaperSize = 1: .Columns = 1
        .MarginTop = 700: .MarginBottom = 700
    End With
    
    vsConsulta.ZOrder 0
    
    Form_Resize
    
    Exit Sub
    
ErrLoad:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Err.Description
End Sub


Private Function ValidoDatos() As Boolean

    ValidoDatos = False
    On Error Resume Next

    If Not IsDate(tDesde.Text) Then
        MsgBox "La fecha ingresada en el campo desde no es válida, verifique.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    If Not IsDate(tHasta.Text) Then
        MsgBox "La fecha ingresada en el campo hasta no es válida, verifique.", vbExclamation, "ATENCIÓN"
        Foco tHasta: Exit Function
    End If
    
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then
        MsgBox "La fecha desde no debe ser mayor a la fecha hasta. Verifique.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 And Trim(cMoneda.Text) <> "" Then
        MsgBox "La moneda seleccionada no es válida.", vbExclamation, "ATENCIÓN"
        cMoneda.SetFocus: Exit Function
    End If
    
    If cSucursal.ListIndex = -1 And Trim(cSucursal.Text) <> "" Then
        MsgBox "La sucursal seleccionada no es válida.", vbExclamation, "ATENCIÓN"
        cSucursal.SetFocus: Exit Function
    End If
    
    If Trim(tArticulo.Text) <> "" And Val(tArticulo.Tag) = 0 Then
        MsgBox "El artículo seleccionado no es válido. Verifique", vbExclamation, "ATENCIÓN"
        Foco tArticulo: Exit Function
    End If
    
    ValidoDatos = True
    
End Function

Private Sub Form_Resize()
    
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11

    vsListado.Top = Shape1.Top + Shape1.Height + 80
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    Shape1.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = Shape1.Width
    vsListado.Left = Shape1.Left
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
    picBotones.Width = vsListado.Width
    'pbProgreso.Width = picBotones.Width - pbProgreso.Left - 150
    picBotones.BorderStyle = 0
    Screen.MousePointer = 0
    
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
    Foco cGrupo
End Sub

Private Sub Label2_Click()
    Foco tDesde
End Sub

Private Sub Label3_Click()
    Foco cMoneda
End Sub

Private Sub Label4_Click()
    Foco tHasta
End Sub

Private Sub Label5_Click()
    Foco tArticulo
End Sub

Private Sub Label6_Click()
    Foco cSucursal
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = 0
End Sub


Private Sub tArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Dim aTexto As String
        If Trim(tArticulo.Text) = "" Or Val(tArticulo.Tag) <> 0 Then Foco cGrupo: Exit Sub
    
        Dim aSeleccionado As Long
        
        If IsNumeric(tArticulo.Text) Then
            Cons = "Select * from Articulo Where ArtCodigo = " & tArticulo.Text
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                tArticulo.Text = Trim(RsAux!ArtNombre)
                tArticulo.Tag = RsAux!ArtId
                Foco cGrupo
            Else
                MsgBox "No existe un artículo para el código ingresado.", vbExclamation, "ATENCIÓN"
            End If
            RsAux.Close
            Exit Sub
        End If
        
        On Error GoTo errBuscar
        Screen.MousePointer = 11
        Cons = "Select ArtId, ArtNombre 'Artículo', ArtCodigo 'Código' from Articulo" _
                & " Where ArtNombre LIKE '" & Trim(tArticulo.Text) & "%'" _
                & " ORDER BY ArtNombre"
    
        Dim objLista As New clsListadeAyuda
        objLista.ActivoListaAyuda Cons, False, txtConexion, 4400
        Me.Refresh
        aSeleccionado = objLista.ValorSeleccionado
        aTexto = objLista.ItemSeleccionado
        Set objLista = Nothing
        
        If aSeleccionado > 0 Then
            tArticulo.Text = Trim(aTexto)
            tArticulo.Tag = aSeleccionado
            Foco cGrupo
        End If
        
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tDesde_GotFocus()
    tDesde.SelStart = 0
    tDesde.SelLength = Len(tDesde.Text)
End Sub

Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tHasta
End Sub

Private Sub tDesde_LostFocus()
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, "d-Mmm yyyy")
End Sub

Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cMoneda
End Sub

Private Sub tHasta_LostFocus()
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, "d-Mmm yyyy")
End Sub

Private Sub vsConsulta_DblClick()
    On Error Resume Next
    If vsConsulta.Rows > 1 Then
        If vsConsulta.Row > 0 And Val(vsConsulta.Cell(flexcpData, vsConsulta.Row, 0)) > 0 Then
            Shell App.Path & "\appExploreMsg.exe " & prmPlantilla & ":" & vsConsulta.Cell(flexcpData, vsConsulta.Row, 0), vbNormalFocus
        End If
    End If
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub cMoneda_GotFocus()
    cMoneda.SelStart = 0: cMoneda.SelLength = Len(cMoneda.Text)
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cSucursal
End Sub

Private Sub cMoneda_LostFocus()
    cMoneda.SelLength = 0
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub
    With vsListado
        .StartDoc
        
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim aStr As String
    aStr = "Acumulado de Ventas por Artículo"
    If cSucursal.ListIndex <> -1 Then aStr = aStr & " (" & Trim(cSucursal.Text) & ")"
    EncabezadoListado vsListado, aStr, False
    vsListado.FileName = "Acumulado de Ventas"
    
    vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
    vsListado.EndDoc

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

Private Function EncabezadoLista(Optional Desde As String = "", Optional Hasta As String = "") As Integer

    On Error Resume Next
    InicializoGrilla
    
    Dim aFormato As String
    If Desde <> "" And Hasta <> "" Then
        Dim Meses As Integer, MActual As Date
        
        Meses = DateDiff("m", CDate(Desde), CDate(Hasta))
        MActual = CDate(Desde)
        For I = 0 To Meses
            With vsConsulta
                aFormato = aFormato & Format(DateAdd("m", I, MActual), "MM/YY") & "|"
            End With
                
        Next
    End If
    
    vsConsulta.FormatString = "<Artículo|" & aFormato & "Totales|"
    vsConsulta.ColWidth(0) = 3200

    EncabezadoLista = Meses
    
End Function

Private Sub InicializoGrilla()

    On Error Resume Next
    With vsConsulta
        .OutlineBar = flexOutlineBarNone ' flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True: .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Artículo"
            
        .WordWrap = False
        .ColWidth(0) = 2700
        .MergeCells = flexMergeSpill: .ExtendLastCol = True
    End With
          
End Sub

Private Sub CargoCombos()

    On Error GoTo errCargo

    Cons = "Select SucCodigo, SucAbreviacion From Sucursal " _
        & " Where SucDcontado <> Null Or SucDCredito <> Null"
    CargoCombo Cons, cSucursal, ""
    
    Cons = "Select MonCodigo, MonSigno From Moneda Where MonFactura = 1"
    CargoCombo Cons, cMoneda, ""
    
    Cons = "Select GruCodigo, GruNombre  From Grupo Order by GruNombre"
    CargoCombo Cons, cGrupo, ""
        
    Cons = "Select TipCodigo, TipNombre From Tipo Order by TipNombre"
    CargoCombo Cons, cboTipoArticulo, ""

    Exit Sub

errCargo:
    clsGeneral.OcurrioError "Ocurrió un error los datos en los combos.", Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Sub AccionLimpiar()
    tDesde.Text = "": tHasta.Text = ""
    tArticulo.Text = "": cMoneda.Text = ""
    cSucursal.Text = ""
    cGrupo.Text = ""
End Sub
