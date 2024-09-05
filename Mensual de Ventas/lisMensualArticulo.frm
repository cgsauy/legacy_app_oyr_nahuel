VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Begin VB.Form LisMensualArticulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensual de Ventas por Artículo"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "lisMensualArticulo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   9195
   Begin ComctlLib.ListView lLista 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Fecha"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descripción"
         Object.Width           =   4234
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Terminal"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Usuario"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Defensa"
         Object.Width           =   3881
      EndProperty
   End
   Begin VB.Frame Shape1 
      Caption         =   "Filtro de Datos"
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton bImprimir 
         Caption         =   "&Imprimir"
         Height          =   325
         Left            =   7800
         TabIndex        =   14
         ToolTipText     =   "Generar Historial de Ventas"
         Top             =   210
         Width           =   975
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   45
         TabIndex        =   7
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   3
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton bConsultar 
         Caption         =   "&Consultar"
         Height          =   325
         Left            =   7800
         TabIndex        =   10
         ToolTipText     =   "Generar Historial de Ventas"
         Top             =   580
         Width           =   975
      End
      Begin VB.ComboBox cSucursal 
         Height          =   315
         Left            =   5160
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   570
         Width           =   2415
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
      Begin VB.ComboBox cMoneda 
         Height          =   315
         Left            =   5160
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   630
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Sucursal:"
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   600
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
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin vsViewLib.vsPrinter vsPrinter 
      Height          =   4335
      Left            =   1200
      TabIndex        =   12
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
      Zoom            =   70
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuOpConsultar 
         Caption         =   "&Consultar"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuOpImprimir 
         Caption         =   "&Imprimr"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalFormulario 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "LisMensualArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private QArticuloE As rdoQuery, QArticuloN As rdoQuery
Private RsArticuloE As rdoResultset, RsArticuloN As rdoResultset

Private strFormato As String
Private strEncabezado As String

Dim aColumnas As Integer

Private strTexto As String

Private Sub AccionConsultar()
On Error GoTo ErrBC

Dim Meses As Integer, aArticulo As Long

    If Not ValidoDatos Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Meses = EncabezadoLista(tDesde.Text, tHasta.Text) + 1
    
    Cons = "Select Mes = DatePart(mm,AArFecha), Ano = DatePart(yy,AArFecha), ArtCodigo, ArtNombre, Cantidad = (Sum(AArCantidadNCo) + Sum(AArCantidadNCr) + Sum(AArCantidadECo) + Sum(AArCantidadECr)) " _
            & " From AcumuladoArticulo, Articulo" _
            & " Where AArArticulo = ArtID" _
            & " And AArFEcha Between " & Format(tDesde.Text, "'mm/dd/yyyy'") & " And " & Format(tHasta.Text, "'mm/dd/yyyy'")
            
    If cMoneda.ListIndex <> -1 Then Cons = Cons & " And AArMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    If cSucursal.ListIndex <> -1 Then Cons = Cons & " And AArSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And AArArticulo  = " & Val(tArticulo.Tag)
    
    Cons = Cons & " Group by ArtNombre, AArArticulo, ArtCodigo, DatePart(mm,AArFecha), DatePart(yy,AArFecha)" & " Order by ArtNombre"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        aArticulo = 0
        Do While Not RsAux.EOF
            If RsAux!ArtCodigo <> aArticulo Then aArticulo = RsAux!ArtCodigo
            Set itmX = lLista.ListItems.Add
            itmX.Text = Format(RsAux!ArtCodigo, "000,000")
            itmX.SubItems(1) = Trim(RsAux!ArtNombre)
            
            For I = 1 To Meses: itmX.SubItems(1 + I) = 0: Next
            
            Do While RsAux!ArtCodigo = aArticulo
                For I = 1 To Meses
                    If lLista.ColumnHeaders(2 + I).Text = Format(RsAux!Mes & "/" & RsAux!Ano, "mm/yy") Then
                        itmX.SubItems(1 + I) = RsAux!Cantidad
                        Exit For
                    End If
                Next
                RsAux.MoveNext
                If RsAux.EOF Then Exit Do
            Loop
        Loop
    End If
    RsAux.Close
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrBC:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al consultar."
    cBase.QueryTimeout = 15

End Sub

Private Sub PropiedadesImpresion()

  With vsPrinter
        .PreviewPage = 1
        .FontName = "Tahoma"
        .FontSize = 8
        .FontBold = False
        .FontItalic = False
        .TextAlign = 0  'Left
        .PageBorder = 3
        .Orientation = orLandscape
        .PenStyle = 0
        .BrushStyle = 0
        .PenWidth = 2
        .PenColor = 0
        .BrushColor = 0
        .TextColor = 0
        .Columns = 1
        .Zoom = 70
        .TableBorder = tbTopBottom
    End With
    
End Sub
Private Sub ArmoEncabezado()

    With vsPrinter
        .FontSize = 8: .FontBold = True
        
        strFormato = "+^900|+< 3750"
        strEncabezado = "Código|Artículo"
        
        For I = 3 To lLista.ColumnHeaders.Count
            strFormato = strFormato & "|+>" & CInt(lLista.ColumnHeaders(I).Width) * 1.5
            strEncabezado = strEncabezado & "|" & lLista.ColumnHeaders(I).Text
        Next
        
        .TableBorder = tbBoxRows
        .AddTable strFormato, strEncabezado, "", Gris
        
        .FontBold = False
    End With

End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bImprimir_Click()
    AccionImprimir
End Sub

Private Sub cSucursal_Change()
    Selecciono cSucursal, cSucursal.Text, gTecla
End Sub
Private Sub cSucursal_GotFocus()
    cSucursal.SelStart = 0
    cSucursal.SelLength = Len(cSucursal.Text)
End Sub

Private Sub cSucursal_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cSucursal.ListIndex
End Sub

Private Sub cSucursal_KeyPress(KeyAscii As Integer)
    cSucursal.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub cSucursal_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cSucursal
End Sub

Private Sub cSucursal_LostFocus()
    gIndice = -1
    cSucursal.SelLength = 0
End Sub

Private Sub Form_Activate()
    Me.Refresh
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    gIndice = -1
    SetearLView lvValores.FullRow, lLista
    CargoSucursal
    CargoMoneda
    
    FechaDelServidor
    tDesde.Text = Format(PrimerDia(gFechaServidor), "d-Mmm yyyy")
    tHasta.Text = Format(gFechaServidor, "d-Mmm yyyy")
    
    EncabezadoLista
    PropiedadesImpresion
    Exit Sub
    
ErrLoad:
    Screen.MousePointer = vbDefault
    ErrorInesperado Err.Description
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
    AjustoObjetos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Forms(Forms.Count - 2).SetFocus
End Sub

Private Sub Label1_Click()
    Foco cSucursal
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

Private Sub MnuOpConsultar_Click()
    AccionConsultar
End Sub

Private Sub MnuOpImprimir_Click()
    AccionImprimir
End Sub

Private Sub MnuSalFormulario_Click()
    Unload Me
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = 0
End Sub


Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
    
        If Trim(tArticulo.Text) = "" Or Val(tArticulo.Tag) <> 0 Then Foco cSucursal: Exit Sub
    
        Dim aSeleccionado As Long
        
        If IsNumeric(tArticulo.Text) Then
            Cons = "Select * from Articulo Where ArtCodigo = " & tArticulo.Text
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                tArticulo.Text = Trim(RsAux!ArtNombre)
                tArticulo.Tag = RsAux!ArtId
                Foco cSucursal
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
            Foco cSucursal
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

Private Sub vsPrinter_EndDoc()
    EnumeroPiedePagina vsPrinter
End Sub

Private Sub CargoMoneda()
    
    On Error GoTo ErrCM
    Cons = "Select MonCodigo, MonSigno From Moneda"
    CargoCombo Cons, cMoneda, ""
    Exit Sub

ErrCM:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al cargar las monedas."
End Sub

Private Sub cMoneda_Change()
    Selecciono cMoneda, cMoneda.Text, gTecla
End Sub

Private Sub cMoneda_GotFocus()
    gIndice = -1
    cMoneda.SelStart = 0
    cMoneda.SelLength = Len(cMoneda.Text)
End Sub

Private Sub cMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cMoneda.ListIndex
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    cMoneda.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then Foco tArticulo

End Sub

Private Sub cMoneda_KeyUp(KeyCode As Integer, Shift As Integer)

    ComboKeyUp cMoneda

End Sub

Private Sub cMoneda_LostFocus()
    gIndice = -1
    cMoneda.SelLength = 0
End Sub


Private Sub AjustoObjetos()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    lLista.Left = Shape1.Left
    lLista.Width = Me.Width - 340
    Shape1.Width = lLista.Width
    lLista.Height = Me.Height - 1950
        
End Sub

Private Sub AccionImprimir()
    
    If lLista.ListItems.Count = 0 Then
        MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo errPrint
    Screen.MousePointer = 11
    With vsPrinter
    
        If Not .PrintDialog(pdPrinterSetup) Then Exit Sub
        
        .Preview = True
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
    
        EncabezadoListado vsPrinter, "Acumulado de Ventas por Artículo", False
        .filename = "Acumulado de Ventas (Mensual)"
        .FontSize = 8: .FontBold = False
        
        'ArmoEncabezado
        For Each itmX In lLista.ListItems
            strTexto = Trim(itmX.Text) & "|" & Trim(itmX.SubItems(1))
            For I = 2 To lLista.ColumnHeaders.Count - 1
                strTexto = strTexto & "|" & Trim(itmX.SubItems(I))
            Next
            .AddTable strFormato, "", strTexto, Gris, , True
        Next
        
        .EndDoc
        .PrintDoc
        
    End With
    Screen.MousePointer = 0
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión. " & Trim(Err.Description)
End Sub

Private Sub vsPrinter_NewPage()
    ArmoEncabezado
End Sub

Private Function EncabezadoLista(Optional Desde As String = "", Optional Hasta As String = "") As Integer

    On Error Resume Next
    lLista.ListItems.Clear
    
    lLista.ColumnHeaders.Clear
    lLista.ColumnHeaders.Add Text:="Código", Width:=650
    lLista.ColumnHeaders.Add Text:="Artículo", Width:=3200
    If Desde <> "" And Hasta <> "" Then
        Dim Meses As Integer, MActual As Date
        
        Meses = DateDiff("m", CDate(Desde), CDate(Hasta))
        MActual = CDate(Desde)
        For I = 0 To Meses
            lLista.ColumnHeaders.Add Text:=Format(DateAdd("m", I, MActual), "MM/YY"), Width:=450, Alignment:=1
        Next
    End If
    lLista.Refresh
    EncabezadoLista = Meses
    
End Function

Private Sub CargoSucursal()

    Cons = "Select SucCodigo, SucAbreviacion From Sucursal " _
        & " Where SucDcontado <> Null Or SucDCredito <> Null"
    CargoCombo Cons, cSucursal, ""
    
End Sub


