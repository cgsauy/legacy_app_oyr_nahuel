VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmCobranzas 
   BackColor       =   &H8000000B&
   Caption         =   "Últimas Cobranzas"
   ClientHeight    =   5145
   ClientLeft      =   3930
   ClientTop       =   1920
   ClientWidth     =   6030
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCobranzas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6030
   Begin VB.PictureBox picPie 
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4395
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4200
      Width           =   4455
      Begin VB.TextBox lDoc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox tDocumento 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   720
         TabIndex        =   2
         Top             =   345
         Width           =   795
      End
      Begin VB.CommandButton bExit 
         Caption         =   "&Salir"
         Height          =   315
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº &Doc.:"
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lTImporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1500
         TabIndex        =   8
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lPromedio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3540
         TabIndex        =   7
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lTotales 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   45
         Width           =   1695
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   3915
      Left            =   420
      TabIndex        =   0
      Top             =   180
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   6906
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   5
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
   Begin VB.Menu MnuAcciones 
      Caption         =   "MnuAcciones"
      Visible         =   0   'False
      Begin VB.Menu MnuTitulo 
         Caption         =   "MnuTitulo"
      End
      Begin VB.Menu MnuAL0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPendiente 
         Caption         =   "Dejar como &Pendiente"
      End
      Begin VB.Menu MnuMExtranjera 
         Caption         =   "Tomar Moneda &Extranjera"
      End
      Begin VB.Menu MnuAL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuChequesDif 
         Caption         =   "Ingresar Cheques Di&feridos"
      End
      Begin VB.Menu MnuChequesAlD 
         Caption         =   "Ingresar Cheques Al &Día"
      End
      Begin VB.Menu MnuAL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAnulaciones 
         Caption         =   "&Anulación de Documentos"
      End
      Begin VB.Menu MnuReimprimir 
         Caption         =   "Reimprimir Documento"
      End
      Begin VB.Menu MnuVOpe 
         Caption         =   "&Visualización de Operaciones"
      End
   End
End
Attribute VB_Name = "frmCobranzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mTexto As String
Dim mValor As Long

'DATOS DE LAS MONEDAS -------------------------------------------------------------------
Private Type typMoneda
    Codigo As Integer
    Signo As String * 3
    TCaPesos As Currency
End Type

Private arrMonedas() As typMoneda

Private prmPrimeraHora As String

Private Sub bExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    
    FechaDelServidor
    Me.BackColor = RGB(222, 184, 135)
    picPie.BackColor = Me.BackColor
    picPie.BorderStyle = 0
    lDoc.BackColor = picPie.BackColor
    
    ObtengoSeteoForm Me, WidthIni:=5040, HeightIni:=3870
    Me.Height = 3870
    Me.Width = 5040
    
    InicializoControles
    
    CargoDatos
    
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Error al iniciar el formulario.", Err.Description
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With picPie
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleHeight - .Height
    End With
       
    With vsLista
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop
        .Height = Me.ScaleHeight - picPie.Height
        .Width = Me.ScaleWidth
    End With
    
    bExit.Left = picPie.ScaleWidth - (bExit.Width + 60)
    lDoc.Width = bExit.Left - lDoc.Left - 50
    
    lTotales.Left = 40: lTotales.Top = 45
    lTotales.Width = picPie.ScaleWidth
    
    lPromedio.Left = 40: lPromedio.Top = lTotales.Top
    lPromedio.Width = picPie.ScaleWidth - (lPromedio.Left + 60)
    
    lTImporte.Top = lTotales.Top
    lTImporte.Width = 1695
    lTImporte.Left = 510
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
    EndMain
End Sub


Private Function CargoDatos() As Boolean
On Error GoTo ErrCD
    
    Screen.MousePointer = 11
    vsLista.Rows = 1
    
    cons = "Select Top 10 DocFecha, DocCodigo, DocSerie, DocNumero, DocTotal, DocCliente, DocMoneda ," & _
                                    "CEmNombre as nEmpresa, " & _
                                    " RTrim(CPeNombre1) + ' ' + CPeApellido1 as nPersona, " & _
                                    " UsuIdentificacion" & _
                " From Documento " & _
                        " Left Outer Join CPersona On DocCliente = CPeCliente " & _
                        " Left Outer Join CEmpresa On DocCliente = CEmCliente " & _
                " , Usuario" & _
                " Where DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")" & _
                " And DocAnulado = 0" & _
                " And DocFecha Between " & Format(gFechaServidor, "'mm/dd/yyyy 00:00'") & _
                                                    " And " & Format(gFechaServidor, "'mm/dd/yyyy 23:59'") & _
                " And DocSucursal = " & paCodigoDeSucursal & _
                " And DocUsuario = UsuCodigo" & _
                " Order by DocCodigo Desc"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not rsAux.EOF
        With vsLista
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!DocFecha, "hh:mm")
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!DocSerie) & "-" & rsAux!DocNumero
            mValor = rsAux!DocCodigo: .Cell(flexcpData, .Rows - 1, 1) = mValor
            
            mTexto = Format(rsAux!DocTotal, "#,##0.00")
            If Right(mTexto, 3) = ".00" Then mTexto = Mid(mTexto, 1, Len(mTexto) - 3)
            mTexto = arrMonedaProp(rsAux!DocMoneda, 1) & " " & mTexto
            
            .Cell(flexcpText, .Rows - 1, 2) = mTexto
            
            If Not IsNull(rsAux!nEmpresa) Then
                .Cell(flexcpText, .Rows - 1, 3) = Trim(rsAux!nEmpresa)
            Else
                .Cell(flexcpText, .Rows - 1, 3) = Trim(rsAux!nPersona)
            End If
            mValor = rsAux!DocCliente: .Cell(flexcpData, .Rows - 1, 3) = mValor
            
            .Cell(flexcpText, .Rows - 1, 4) = Trim(rsAux!UsuIdentificacion)
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    If prmAccesoDatos Then CargoTotales
    
    On Error Resume Next
    If vsLista.Rows > 1 Then vsLista.SetFocus
    Screen.MousePointer = 0
    Exit Function
    
ErrCD:
    clsGeneral.OcurrioError "Error al cargar la lista.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function CargoTotales() As Boolean
On Error GoTo ErrCD
    
    Screen.MousePointer = 11
    Dim mQTotal As Long, mImporte As Currency
    
    cons = "Select DocMoneda, Count(*) as Q, Sum(DocTotal) as Total" & _
                " From Documento " & _
                " Where DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")" & _
                " And DocAnulado = 0" & _
                " And DocFecha Between " & Format(gFechaServidor, "'mm/dd/yyyy 00:00'") & _
                                                    " And " & Format(gFechaServidor, "'mm/dd/yyyy 23:59'") & _
                " And DocSucursal = " & paCodigoDeSucursal & _
                " Group by DocMoneda"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
    Dim mAuxiliar As Currency
    Do While Not rsAux.EOF
        
        mQTotal = mQTotal + rsAux!Q
        
        mAuxiliar = arrMonedaProp(rsAux!DocMoneda, 2)
        mAuxiliar = mAuxiliar * rsAux!Total
        
        mImporte = mImporte + mAuxiliar
    
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    If mImporte > 0 Then
    mTexto = Format(mImporte, "#,##0.00")
    If Right(mTexto, 3) = ".00" Then mTexto = Mid(mTexto, 1, Len(mTexto) - 3)
    lTotales.Caption = "Total   " & mQTotal
    lTImporte.Caption = "$ " & mTexto
    
    'Cargo los datos del promedio por hora ------------------------------------------
    Dim aQHoras As Currency
    aQHoras = DateDiff("n", CDate(prmPrimeraHora), CDate(vsLista.Cell(flexcpText, 1, 0)))
    aQHoras = aQHoras / 60
    If aQHoras = 0 Then aQHoras = 1
    
    mTexto = Format(mImporte / aQHoras, "#,##0")
    lPromedio.Caption = "Prom x Hora: " & Format(mQTotal / aQHoras, "0.#") & _
                                   "   $ " & mTexto
                                   
    '----------------------------------------------------------------------------------------
    End If
    Screen.MousePointer = 0
    Exit Function
    
ErrCD:
    clsGeneral.OcurrioError "Error al cargar los totales.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub InicializoControles()
    
    On Error Resume Next
    bExit.BackColor = Me.BackColor
    lDoc.Text = "": lDoc.ForeColor = Colores.RojoClaro
    
    lPromedio.Caption = "": lTotales.Caption = "": lTImporte.Caption = ""
    
    With vsLista
        .Editable = False
        .Rows = 1: .Cols = 4
        .FormatString = "<Hora|<Nº Doc.|>Importe|<Cliente|Usuario"
        .ColWidth(0) = 525: .ColWidth(1) = 900: .ColWidth(2) = 820: .ColWidth(3) = 1750
        
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .AllowBigSelection = False
        .AllowSelection = False
        
        .BackColorBkg = RGB(222, 184, 135)
        .BackColor = RGB(255, 222, 173)
        .BackColorAlternate = RGB(255, 235, 205)
        
        .BackColorFixed = RGB(222, 184, 135)
        .ForeColorFixed = RGB(250, 240, 230)
        
        .BorderStyle = flexBorderNone
        .HighLight = flexHighlightWithFocus
        .FocusRect = flexFocusNone
        .RowHeight(0) = 300
        .ForeColorSel = Colores.RojoClaro
        .BackColorSel = .BackColorBkg
        .ScrollBars = flexScrollBarNone
        
    End With
    
   tDocumento.BackColor = vsLista.BackColor
   lTotales.ForeColor = vsLista.ForeColorSel
   lPromedio.ForeColor = vsLista.ForeColorSel
   lTImporte.ForeColor = vsLista.ForeColorSel
   
    CargoArrayMonedas
    CargoPrimeraHora
    
End Sub

Private Sub Label5_Click()
    Foco tDocumento
End Sub

Private Sub MnuAnulaciones_Click()
    EjecutarApp prmPathApp & "Anulaciones.exe", MnuTitulo.Tag
End Sub

Private Sub MnuChequesAlD_Click()
    EjecutarApp prmPathApp & "AltaCheques.exe", "T 0|D " & MnuTitulo.Tag
End Sub

Private Sub MnuChequesDif_Click()
    EjecutarApp prmPathApp & "AltaCheques.exe", "T 1|D " & MnuTitulo.Tag
End Sub

Private Sub MnuMExtranjera_Click()
    EjecutarApp prmPathApp & "CompraME.exe", "D " & MnuTitulo.Tag
End Sub

Private Sub MnuPendiente_Click()
    EjecutarApp prmPathApp & "PendientesCaja.exe", "D " & MnuTitulo.Tag
End Sub

Private Sub MnuReimprimir_Click()
    EjecutarApp prmPathApp & "Reimpresion.exe", MnuTitulo.Tag
End Sub

Private Sub MnuVOpe_Click()
    EjecutarApp prmPathApp & "Visualizacion de Operaciones.exe", MnuVOpe.Tag
End Sub

Private Sub tDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 93 Then
        If Val(lDoc.Tag) = 0 Then Exit Sub
        
        Dim mX As Single, mY As Single
        mX = picPie.Left + tDocumento.Left + 300
        mY = picPie.Top + tDocumento.Top + tDocumento.Height
        MnuTitulo.Caption = "Doc. " & tDocumento.Text
        MnuTitulo.Tag = Val(lDoc.Tag)
        MnuVOpe.Tag = Val(tDocumento.Tag)
        
        PopupMenu MnuAcciones, , mX, mY, MnuTitulo
    End If
    
End Sub

Private Sub tDocumento_Change()
    lDoc.Text = "": lDoc.Tag = 0
End Sub

Private Sub tDocumento_GotFocus()
    tDocumento.SelStart = 0: tDocumento.SelLength = (Len(tDocumento.Text))
End Sub

Private Sub tDocumento_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(lDoc.Tag) <> 0 Or Trim(tDocumento.Text) = "" Then
            CargoDatos
            vsLista.SetFocus
            Exit Sub
        End If
        On Error GoTo errDoc
        
        Dim adQ As Integer, adCodigo As Long, adTexto As String, adTextoFmt As String
        
        Screen.MousePointer = 11
        Dim mDSerie As String, mDNumero As String
        mDNumero = tDocumento.Text
        
        If InStr(tDocumento.Text, "-") <> 0 Then
            mDSerie = Mid(tDocumento.Text, 1, InStr(tDocumento.Text, "-") - 1)
            mDNumero = Mid(tDocumento.Text, InStr(tDocumento.Text, "-") + 1)
        ElseIf Not IsNumeric(Left(tDocumento.Text, 1)) Then
            mDSerie = Left(tDocumento.Text, 1)
            mDNumero = Mid(tDocumento.Text, 2)
        End If
        
        adQ = 0
        cons = "Select * from Documento " & _
                   " Where DocNumero = " & Val(mDNumero) & _
                   " And DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")"
        If Trim(mDSerie) <> "" Then cons = cons & " And DocSerie = '" & mDSerie & "'"
        
        cons = cons & " And DocFecha Between " & Format(gFechaServidor - 2, "'mm/dd/yyyy 00:00'") & _
                                                    " And " & Format(gFechaServidor, "'mm/dd/yyyy 23:59'")
                            '" And DocSucursal = " & paCodigoDeSucursal
        
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            adCodigo = rsAux!DocCodigo
            adTexto = zDocumento(rsAux!DocTipo, rsAux!DocSerie, rsAux!DocNumero, adTextoFmt)
            adQ = 1
            rsAux.MoveNext: If Not rsAux.EOF Then adQ = 2
        End If
        rsAux.Close
        
        Select Case adQ
            Case 2
                Dim miLDocs As New clsListadeAyuda
                cons = "Select DocCodigo, DocFecha as Fecha, DocSerie + Convert(char(7),DocNumero) as Numero " & _
                           " from Documento " & _
                           " Where DocNumero = " & Val(mDNumero) & _
                           " And DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")"
                If Trim(mDSerie) <> "" Then cons = cons & " And DocSerie = '" & mDSerie & "'"
                adCodigo = miLDocs.ActivarAyuda(cBase, cons, 4100, 1)
                Me.Refresh
                If adCodigo <> 0 Then adCodigo = miLDocs.RetornoDatoSeleccionado(0)
                Set miLDocs = Nothing
                
                If adCodigo > 0 Then
                    cons = "Select * from Documento Where DocCodigo = " & adCodigo
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    If Not rsAux.EOF Then
                        adTexto = zDocumento(rsAux!DocTipo, rsAux!DocSerie, rsAux!DocNumero, adTextoFmt)
                    End If
                    rsAux.Close
                End If
        End Select
        
        If adCodigo > 0 Then
            tDocumento.Text = adTextoFmt
            lDoc.Tag = adCodigo: lDoc.Text = adTexto
            
            'Saco el Nombre del Cliente ----------------------------------
            cons = "Select  DocCliente, CEmNombre as nEmpresa, " & _
                                  " RTrim(CPeNombre1) + ' ' + CPeApellido1 as nPersona " & _
                        " From Documento " & _
                                " Left Outer Join CPersona On DocCliente = CPeCliente " & _
                                " Left Outer Join CEmpresa On DocCliente = CEmCliente " & _
                        " Where DocCodigo = " & Val(lDoc.Tag)
            Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
            If Not rsAux.EOF Then
                If Not IsNull(rsAux!nEmpresa) Then adTexto = Trim(rsAux!nEmpresa) Else adTexto = Trim(rsAux!nPersona)
                tDocumento.Tag = rsAux!DocCliente
            End If
            rsAux.Close
             lDoc.Text = adTexto
        Else
            lDoc.Text = " No Existe !!"
        End If
        
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errDoc:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function zDocumento(Tipo As Integer, Serie As String, Numero As Long, retSerieNumero As String) As String

    Select Case Tipo
        Case 1: zDocumento = "Ctdo. "
        Case 2: zDocumento = "Créd. "
        Case 3: zDocumento = "N/Dev. "
        Case 4: zDocumento = "N/Créd. "
        Case 5: zDocumento = "Recibo "
        Case 10: zDocumento = "N/Esp. "
    End Select
    
    zDocumento = zDocumento & Trim(Serie) & "-" & Numero
    retSerieNumero = Trim(Serie) & "-" & Numero

End Function

Private Sub vsLista_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyReturn: CargoDatos
        
        Case vbKeyF4: Foco tDocumento
        
        Case 93:
            If vsLista.Rows = 1 Then Exit Sub
            Dim mX As Single, mY As Single
            mX = 800
            mY = 300 + ((vsLista.Row - 1) * vsLista.RowHeight(1))
            MnuTitulo.Caption = "Doc. " & vsLista.Cell(flexcpText, vsLista.Row, 1)
            MnuTitulo.Tag = vsLista.Cell(flexcpData, vsLista.Row, 1)
            MnuVOpe.Tag = vsLista.Cell(flexcpData, vsLista.Row, 3)
            PopupMenu MnuAcciones, , mX, mY, MnuTitulo
    End Select
    
End Sub


Private Function CargoArrayMonedas() As Boolean

On Error GoTo errMonedas
    ReDim Preserve arrMonedas(0)
    
    cons = "Select * from Moneda Order by MonCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        If arrMonedas(0).Codigo <> 0 Then ReDim Preserve arrMonedas(UBound(arrMonedas) + 1)
        With arrMonedas(UBound(arrMonedas))
            .Codigo = rsAux!MonCodigo
            .Signo = Trim(rsAux!MonSigno)
            If .Codigo = prmMonedaPesos Then
                .TCaPesos = 1
            Else
                .TCaPesos = TasadeCambio(.Codigo, CInt(prmMonedaPesos), gFechaServidor)
            End If
        End With
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    Exit Function

errMonedas:
    clsGeneral.OcurrioError "Error al cargar el array de monedas. Los cálculos de intereses y redondeos pueden dar ERROR.", Err.Description, "Error Con Parámetros !!!"
End Function

Private Function arrMonedaProp(mIDMoneda As Long, mDato As Integer) As Variant
    On Error GoTo errArray
    Dim idx As Integer
    
    arrMonedaProp = -1
    
    For idx = LBound(arrMonedas) To UBound(arrMonedas)
        If arrMonedas(idx).Codigo = mIDMoneda Then
            Select Case mDato
                Case 1: arrMonedaProp = Trim(arrMonedas(idx).Signo)
                Case 2: arrMonedaProp = arrMonedas(idx).TCaPesos
            End Select
            Exit For
        End If
    Next
    
errArray:
End Function

Private Function CargoPrimeraHora()
On Error GoTo errPH

    prmPrimeraHora = "09:00"
    
    cons = "Select Min(DocFecha) as DocFecha" & _
                " From Documento " & _
                " Where DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")" & _
                " And DocAnulado = 0" & _
                " And DocFecha Between " & Format(gFechaServidor, "'mm/dd/yyyy 00:00'") & _
                                                    " And " & Format(gFechaServidor, "'mm/dd/yyyy 23:59'") & _
                " And DocSucursal = " & paCodigoDeSucursal
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!DocFecha) Then
            prmPrimeraHora = Format(rsAux!DocFecha, "hh:mm")
        End If
    End If
    rsAux.Close
    
errPH:
End Function
