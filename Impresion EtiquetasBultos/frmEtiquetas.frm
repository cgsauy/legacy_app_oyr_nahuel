VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmEtiquetas 
   Caption         =   "Etiquetas"
   ClientHeight    =   5790
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   Icon            =   "frmEtiquetas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1545
      ScaleWidth      =   9285
      TabIndex        =   0
      Top             =   0
      Width           =   9315
      Begin VB.CommandButton butImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txtCodImp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmEtiquetas.frx":08CA
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Artículo:"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código de impresión:"
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Impresión de etiquetas de bultos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   160
         Width           =   4815
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsEnvio 
      Align           =   1  'Align Top
      Height          =   3555
      Left            =   0
      TabIndex        =   6
      Top             =   1575
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   6271
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
      BackColorFixed  =   1986860
      ForeColorFixed  =   16777215
      BackColorSel    =   5273691
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      BackColorAlternate=   15133671
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
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
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   2715
      _Version        =   196608
      _ExtentX        =   4789
      _ExtentY        =   2143
      _StockProps     =   229
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      PhysicalPage    =   -1  'True
   End
   Begin VB.Menu MnuImpresion 
      Caption         =   "MnuImpresion"
      Visible         =   0   'False
      Begin VB.Menu MnuPrintTodos 
         Caption         =   "Marcar todos los envíos"
      End
      Begin VB.Menu MnuPrintTodosACero 
         Caption         =   "Todos a cero"
      End
   End
End
Attribute VB_Name = "frmEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCnfgPrint As New clsImpTicketsCnfg

Private Sub InicializoGrilla()
    With vsEnvio
        .Redraw = False
        .Rows = 1: .ExtendLastCol = True: .Cols = 1
        .FormatString = " Cant. etiquetas |<Envío | Cantidad | Fecha a entregar|"
        .ColWidth(1) = 1000
        .Redraw = True
    End With
End Sub

Private Sub butImprimir_Click()
    Dim IQ As Integer
    For IQ = vsEnvio.FixedRows To vsEnvio.Rows - vsEnvio.FixedRows
        If Val(vsEnvio.Cell(flexcpText, IQ, 0)) > 0 Then
            ImprimoEtiquetaEnTicket vsEnvio.Cell(flexcpText, IQ, 1), vsEnvio.Cell(flexcpText, IQ, 0)
        End If
    Next
End Sub

Private Sub Form_Load()
    InicializoGrilla
    oCnfgPrint.CargarConfiguracion "FichasAgencia", "FichaBulto"
End Sub

Private Sub Form_Resize()
On Error Resume Next
    vsEnvio.Height = Me.ScaleHeight - vsEnvio.Top
End Sub

Private Sub MnuPrintTodos_Click()
Dim iFila As Integer
    With vsEnvio
        For iFila = 1 To .Rows - 1
            vsEnvio.Cell(flexcpText, iFila, 0) = vsEnvio.Cell(flexcpText, iFila, 2)
        Next
    End With
End Sub

Private Sub MnuPrintTodosACero_Click()
    vsEnvio.Cell(flexcpText, 1, 0, vsEnvio.Rows - 1) = "0"
End Sub

Private Sub txtArticulo_Change()
    If Val(txtArticulo.Tag) > 0 Then
        vsEnvio.Rows = 1
        txtArticulo.Tag = ""
    End If
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(txtArticulo.Tag) > 0 Then
            BuscoCodigoImpresion
        Else
            BuscoArticulo
            If Val(txtArticulo.Tag) > 0 Then BuscoCodigoImpresion
        End If
    End If
End Sub

Private Sub txtCodImp_Change()
    If Val(txtCodImp.Tag) > 0 Then
        vsEnvio.Rows = 1
        txtCodImp.Tag = ""
    End If
End Sub

Private Sub txtCodImp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsNumeric(txtCodImp.Text) Then
        txtArticulo.SetFocus
    End If
End Sub

Private Sub BuscoArticulo()
On Error GoTo ErrBAN
Dim aCodigo As Long: aCodigo = 0

    Screen.MousePointer = vbHourglass
    Cons = "Select ArtID, Código = ArtCodigo, Artículo = rTRIM(ArtNombre) FROM Articulo" _
        & " WHERE (ArtNombre LIKE '" & Replace(Replace(txtArticulo.Text, " ", "%"), "'", "''") & "%'" _
        & " OR ArtCodigo = " & Val(txtArticulo.Text) & ")" _
        & " Order by ArtNombre"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un artículo con esas características.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            txtArticulo.Text = RsAux(2)
            txtArticulo.Tag = RsAux(0)
            RsAux.Close
        Else
            RsAux.Close
            Dim objLista As New clsListadeAyuda
            If objLista.ActivarAyuda(cBase, Cons, 5000, 1, "Lista Artículos") Then
                txtArticulo.Text = objLista.RetornoDatoSeleccionado(2)
                txtArticulo.Tag = objLista.RetornoDatoSeleccionado(0)
            End If
            Set objLista = Nothing       'Destruyo la clase.
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrBAN:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoCodigoImpresion()
On Error GoTo errBCI
    
    If Not (IsNumeric(txtCodImp.Text) And Val(txtArticulo.Tag) > 0) Then
        MsgBox "Debe ingresar el código de impresión y luego seleccionar el artículo.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    vsEnvio.Rows = 1
    Screen.MousePointer = 11
    Cons = "SELECT REvCantidad, REvEnvio, EnvFechaPrometida " & _
        " FROM Envio INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio AND REvArticulo = " & Val(txtArticulo.Tag) & _
        " WHERE EnvCodImpresion = " & Val(txtCodImp.Text)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsEnvio
            .AddItem RsAux("REvCantidad")
            .Cell(flexcpText, .Rows - 1, 1) = RsAux("REvEnvio")
            .Cell(flexcpText, .Rows - 1, 2) = RsAux("REvCantidad")
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux("EnvFechaPrometida"), "dd/MM/yy")
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errBCI:
    clsGeneral.OcurrioError "Ocurrió un error al buscar la información.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub vsEnvio_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If vsEnvio.Rows <= 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyAdd
            vsEnvio.Cell(flexcpText, vsEnvio.Row, 0) = Val(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)) + 1
        Case vbKeyDelete, vbKeySubtract
            vsEnvio.Cell(flexcpText, vsEnvio.Row, 0) = Val(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)) - 1
            If Val(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)) < 0 Then vsEnvio.Cell(flexcpText, vsEnvio.Row, 0) = 0
    End Select
End Sub

Private Sub vsEnvio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsEnvio.Rows > 1 And Button = 2 Then
        PopupMenu MnuImpresion
    End If
End Sub

Private Sub ImprimoEtiquetaEnTicket(ByVal idEnvio As Long, ByVal cantidad As Integer)

Dim rsE As rdoResultset

    Cons = "SELECT EnvCodigo, EnvCodImpresion, EnvDireccion, IsNull(EnvComentario, '') EnvComentario, IsNull(EnvTelefono, '') EnvTelefono, CPersona.*, CEmpresa.*, Cliente.*, ArtNombre " & _
        " FROM Envio " & _
        " INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio AND RevArticulo = " & Val(txtArticulo.Tag) & _
        " INNER JOIN Articulo ON ArtID = REvArticulo " & _
        " INNER JOIN Cliente ON EnvCliente = CliCodigo " & _
        " LEFT Outer Join CPersona ON CliCodigo = CPeCliente " & _
        " LEFT Outer Join CEmpresa ON CliCodigo = CEmCliente " & _
        " WHERE EnvCodigo = " & idEnvio
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsE.EOF Then rsE.Close: Exit Sub
'
'ByVal sNCliente As String, ByVal Direccion As String, ByVal strTelef As String, _
'    ByVal strComentario As String, ByVal factura As String, ByRef arrArt() As String)
    
    With vsListado
        .FileName = "Etiqueta de Agencia"
        .AbortWindow = False
        .MarginTop = 190
        .MarginLeft = 100
        
        .Device = oCnfgPrint.ImpresoraTickets
        .PaperSize = oCnfgPrint.PapelTicket
        .Orientation = orPortrait
        
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
        
        .PageBorder = pbNone
        
        .TextAlign = taLeftTop
        .TableBorder = tbNone
        .FontBold = True
        .FontName = "Tahoma"
        .FontSize = 9
        .AddTable "6000", "Etiqueta de bulto  - Envío: " & rsE("EnvCodigo") & "   Código Imp.: " & rsE("EnvCodImpresion") & " -", ""
        
        .Paragraph = ""
        
        .FontBold = False
        .TextAlign = taCenterTop
        .FontName = "3 of 9 Barcode"
        .FontSize = 24
        .Paragraph = "*EB" & rsE("EnvCodigo") & "*"
        
        .FontName = "Tahoma"
        .FontSize = 9
        .TextAlign = taLeftBaseline
        .Paragraph = ""
        .Paragraph = ""
        
        .TableBorder = tbNone
        .FontBold = True
        Dim sNCliente As String
        Select Case rsE!CliTipo
            Case 1
                sNCliente = Trim(Trim(Format(rsE!CPeNombre1, "#")) & " " & Trim(Format(rsE!CPeNombre2, "#"))) & " " & Trim(Trim(Format(rsE!CPeApellido1, "#")) & " " & Trim(Format(rsE!CPeApellido2, "#")))
            Case 2
                If Not IsNull(rsE!CEmNombre) Then sNCliente = Trim(rsE!CEmFantasia)
                If Not IsNull(rsE!CEmFantasia) Then sNCliente = sNCliente & " (" & Trim(rsE!CEmFantasia) & ")"
        End Select
        
        .AddTable "300|^5500", "|" & sNCliente, ""
        
        .AddTable "5800", clsGeneral.ArmoDireccionEnTexto(cBase, rsE("EnvDireccion"), True, True, False, True, True), ""
        
        .FontBold = False
        
        .Paragraph = ""
        .FontBold = False
        '.Paragraph = "Agencia:  " & Trim(vsEnvio.Cell(flexcpText, vsEnvio.Row, 2))

        .AddTable "950|2500", "Teléfono:|" & rsE("EnvTelefono"), ""
    
        If rsE("EnvComentario") <> "" Then .AddTable "1200|6000", "Comentario:|" & rsE("EnvComentario"), ""
        
        .Paragraph = ""
        .FontSize = 8
        .TableBorder = tbAll
        .AddTable "^500|^5000", "Cant|Artículo", ""
        .AddTable ">500|<5000", "1|" & Trim(rsE("ArtNombre")), ""
        .TableBorder = tbNone
        .FontSize = 9
        .EndDoc
        'Manda a impresora ya que tengo el preview = false
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    rsE.Close
    
    Dim IQ As Integer
    For IQ = 1 To cantidad
        vsListado.PrintDoc
    Next
    
    
End Sub


