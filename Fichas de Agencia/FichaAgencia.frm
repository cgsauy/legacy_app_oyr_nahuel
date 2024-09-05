VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACOMBO.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form FichaAgencia 
   Caption         =   "Fichas de Agencia"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FichaAgencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex6DAOCtl.vsFlexGrid vsArticulo 
      Height          =   1515
      Left            =   120
      TabIndex        =   6
      Top             =   3180
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2672
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
      GridLines       =   1
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsEnvio 
      Height          =   2115
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3731
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
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
   Begin AACombo99.AACombo cCamion 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   180
      Width           =   2835
      _ExtentX        =   5001
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
   End
   Begin VB.CommandButton cmdRemito 
      Caption         =   "&Remitos"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "El&iminar"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5280
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox tFecha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   3
      Text            =   "28-Oct-1999"
      Top             =   180
      Width           =   1215
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   1215
      Left            =   0
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   5955
      _Version        =   196608
      _ExtentX        =   10504
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
   End
   Begin VB.Label labBultosAsignados 
      BackStyle       =   0  'Transparent
      Caption         =   "Bultos Asignados:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label labBultos 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label labArticulos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Artículos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   7215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Camión:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha:"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   180
      Width           =   615
   End
   Begin VB.Shape Marco 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   120
      Top             =   60
      Width           =   7215
   End
End
Attribute VB_Name = "FichaAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strArtFlete As String
Private Sub cmdEliminar_Click()
    
    If cmdEliminar.Caption = "El&iminar" Then
        If MsgBox("¿Confirma eliminar la información de bultos del envío?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
            AccionEliminar
        End If
    Else
        cmdNuevo.Caption = "&Nuevo"
        cmdEliminar.Caption = "El&iminar"
        cmdImprimir.Enabled = False
        If CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 1)) > 0 Then
            cmdEliminar.Enabled = True
            cmdNuevo.Enabled = False
        Else
            cmdEliminar.Enabled = False
            cmdNuevo.Enabled = True
        End If
        vsEnvio.Enabled = True
        vsArticulo.BackColor = Inactivo
        labBultos.Caption = vbNullString
    End If
End Sub
Private Sub cmdImprimir_Click()
    AccionImprimir
End Sub
Private Sub cmdNuevo_Click()
    If cmdNuevo.Caption = "&Grabar" Then
        If AccionGrabar Then
            cmdNuevo.Caption = "&Nuevo"
            cmdEliminar.Caption = "El&iminar"
            cmdImprimir.Enabled = False
            If CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 1)) > 0 Then
                cmdEliminar.Enabled = True
                cmdNuevo.Enabled = False
            Else
                cmdEliminar.Enabled = False
                cmdNuevo.Enabled = True
            End If
            vsEnvio.Enabled = True
            labBultos.Caption = vbNullString
        End If
    Else
        labBultos.Caption = "0"
        cmdNuevo.Caption = "&Grabar"
        cmdEliminar.Caption = "&Cancelar"
        cmdEliminar.Enabled = True
        vsEnvio.Enabled = False
        vsArticulo.BackColor = vbWhite
        cmdImprimir.Enabled = True
        vsArticulo.SetFocus
    End If
End Sub
Private Sub cmdRemito_Click()
Dim respuesta As Integer
    respuesta = MsgBox("Si desea imprimir solo los remitos del envío " & vsEnvio.Cell(flexcpText, vsEnvio.Row, 0) & _
             " presione SI." & Chr(13) & "Si desea imprimir los remitos para todos los envíos de la lista presione NO.", vbYesNoCancel, "ATENCIÓN")
    If respuesta = vbYes Then
        ImprimoRemitos False
    ElseIf respuesta = vbNo Then
        ImprimoRemitos True
    End If
End Sub
Private Sub ImprimoRemitos(sTodos As Boolean)
Dim strDocumentos As String
On Error GoTo ErrIR
    Screen.MousePointer = vbHourglass
    strDocumentos = vbNullString
    'Saco los artículos de flete.---------------------------------------------------------
    'ATENCION: Tengo que controlar el piso para indicar si paga o no piso.
    If sTodos Then
        For I = 1 To vsEnvio.Rows - 1
            'Saco el código de documento y verifico si no lo imprimi.
            If InStr(strDocumentos, vsEnvio.Cell(flexcpData, I, 0) & ",") = 0 Then
                strDocumentos = strDocumentos & vsEnvio.Cell(flexcpData, I, 0) & ","
                ImprimoRemitosAgencia CLng(vsEnvio.Cell(flexcpData, I, 2)), CLng(vsEnvio.Cell(flexcpData, I, 0)), CLng(vsEnvio.Cell(flexcpText, I, 0)), CLng(vsEnvio.Cell(flexcpData, I, 1)), vsEnvio.Cell(flexcpText, I, 3)
            End If
        Next
    Else
        'Saco el código de documento.------------------------------------
        ImprimoRemitosAgencia CLng(vsEnvio.Cell(flexcpData, vsEnvio.Row, 2)), CLng(vsEnvio.Cell(flexcpData, vsEnvio.Row, 0)), CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)), CLng(vsEnvio.Cell(flexcpData, vsEnvio.Row, 1)), vsEnvio.Cell(flexcpText, vsEnvio.Row, 3)
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrIR:
    clsGeneral.OcurrioError "Ocurrio un error al intentar imprimir los remitos."
    Screen.MousePointer = vbDefault
End Sub
Private Sub ImprimoRemitosAgencia(Agencia As Long, lnDocumento As Long, lnEnvio As Long, lnCliente As Long, _
                                                            Direccion As String)
Dim iCant As Integer, cMonto As Currency
Dim Rs As rdoResultset
Dim sNoPagaPiso As Boolean, strTelef As String, strComentario As String
    
    On Error GoTo ErrIRA
    cMonto = 0
    sNoPagaPiso = False
    Cons = "Select * from Agencia Where AgeCodigo = " & Agencia
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!AgeCategoriaPiso) Then
            If RsAux!AgeCategoriaPiso <> paNoPagaPiso Then
                'Busco el artículo. Si da EOF NO PAGA
                Cons = "Select * From Renglon" _
                    & " Where RenDocumento = " & lnDocumento & " And RenArticulo = " & paArticuloPisoAgencia
                Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Rs.EOF Then
                    sNoPagaPiso = True
                End If
                Rs.Close
            Else
                sNoPagaPiso = True
            End If
        End If
    End If
    RsAux.Close
    
    Cons = "Select Cantidad = Sum(RevAEntregar), ArtID, ArtCodigo, ArtNombre, RenPrecio" _
        & " From Envio, RenglonEnvio, Renglon, Articulo" _
        & " Where EnvTipo = " & TipoEnvio.Entrega & "  And EnvDocumento = " & lnDocumento & " And RenDocumento = " & lnDocumento _
        & " And RevAEntregar > 0 And EnvCodigo = RevEnvio And RevArticulo = RenArticulo And RenArticulo = ArtID " _
        & " Group By ArtID, ArtCodigo, ArtNombre, RenPrecio"
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Rs.EOF Then
        Rs.Close
        MsgBox "No se encontró la información de los artículos.", vbExclamation, "ATENCIÓN": Screen.MousePointer = 0
        Exit Sub
    Else
        If EncabezadoImpresion("Remito de Agencia", Direccion, lnCliente) Then
            With vsListado
                .TableBorder = tbAll
                .AddTable "1100|5000", "Cantidad|Artículo", ""
                cMonto = 0
                Do While Not Rs.EOF
                    cMonto = cMonto + Rs!Cantidad * Rs!RenPrecio
                    .AddTable "1100|5000", "", Rs!Cantidad & "|" & Format(Rs!ArtCodigo, "(#,000,000 )") & Trim(Rs!ArtNombre)
                    Rs.MoveNext
                Loop
                Rs.Close
                
                'Imprimo datos de la factura.....................................
                .Paragraph = "": .Paragraph = "Factura: " & BuscoDocumento(vsEnvio.Cell(flexcpData, vsEnvio.Row, 0))
                .Paragraph = "Valor: " & Format(cMonto, FormatoMonedaP)
                
                'Agrego el telefono y el comentario de envío si tiene.
                strTelef = "": strComentario = ""
                Cons = "Select * From Envio Where EnvTipo = " & TipoEnvio.Entrega _
                    & " And EnvDocumento = " & lnDocumento & " And (EnvTelefono Is Not Null or EnvComentario Is Not Null)"
                Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                Do While Not Rs.EOF
                    If Not IsNull(Rs!EnvTelefono) And strTelef = "" Then strTelef = Rs!EnvTelefono
                    If Not IsNull(Rs!EnvComentario) And strComentario = "" Then strComentario = Trim(Rs!EnvComentario)
                    Rs.MoveNext
                Loop
                Rs.Close
                If Trim(strTelef) <> "" Then .Paragraph = "Telefono: " & strTelef
                If Trim(strComentario) <> "" Then .Paragraph = "Comentario: " & strComentario
                
                'Detallo si paga piso.
                If sNoPagaPiso Then .FontBold = True: .Paragraph = "NO PAGA PISO": .FontBold = False
                .EndDoc
                .Device = paIConformeN
                .PaperBin = paIConformeB
                .PrintDoc
                .PrintDoc
                .PrintDoc
            End With
        Else
            Rs.Close
        End If
    End If
    Exit Sub
ErrIRA:
    clsGeneral.OcurrioError "Ocurrio un error al intentar imprimir los remitos para el envío." & lnEnvio
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    Unload Me
End Sub
Private Sub Form_Activate()
    Me.Refresh
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyX: Unload Me
        End Select
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo ErrLoad
    ObtengoSeteoForm Me, Me.Left, , Me.Width
    With vsEnvio
        .Redraw = False
        .Rows = 1: .ExtendLastCol = True: .Cols = 1
        .FormatString = "<Envío|>Bultos|<Agencia|<Localidad - Dirección"
        .ColWidth(2) = 1500
        .Redraw = True
    End With
    With vsArticulo
        .Redraw = False
        .Rows = 1: .ExtendLastCol = True: .Cols = 1
        .FormatString = ">Cantidad|>A Enviar|>Total|<Articulo"
        .Redraw = True
    End With
    AccionLimpiar
    FechaDelServidor
    tFecha.Text = Format(gFechaServidor, "d-Mmm-yyyy")
    
    Cons = "Select CamCodigo, CamNombre From Camion"
    CargoCombo Cons, cCamion
    
    'Saco los artículos de flete.
    strArtFlete = CargoArticulosDeFlete
    
    PrueboBandejaImpresora
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar el formulario. " & Trim(Err.Description)
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
    
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    cmdSalir.Top = Me.ScaleHeight - (Status.Height + 70 + cmdSalir.Height)
    cmdImprimir.Top = cmdSalir.Top: cmdNuevo.Top = cmdSalir.Top: cmdEliminar.Top = cmdSalir.Top: cmdRemito.Top = cmdSalir.Top
    
    vsEnvio.Width = Me.ScaleWidth - (vsEnvio.Left * 2)
    vsArticulo.Width = vsEnvio.Width: Marco.Width = vsEnvio.Width
    labArticulos.Width = vsEnvio.Width
    
    'Ajusto ubicación de botones
    cmdSalir.Left = Me.ScaleWidth - (vsEnvio.Left + cmdSalir.Width)
    cmdImprimir.Left = cmdSalir.Left - (cmdImprimir.Width + 105)
    cmdRemito.Left = cmdImprimir.Left - (cmdRemito.Width + 105)
    
    'Grillas y otros objetos.
    vsEnvio.Height = Me.ScaleHeight / 2
    labArticulos.Top = vsEnvio.Top + vsEnvio.Height + 40
    labBultos.Top = labArticulos.Top: labBultosAsignados.Top = labBultos.Top
    vsArticulo.Top = labArticulos.Top + labArticulos.Height + 10
    vsArticulo.Height = cmdSalir.Top - (vsArticulo.Top + 80)

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
    Foco tFecha
End Sub
Private Sub Label2_Click()
    Foco cCamion
End Sub
Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0
    tFecha.SelLength = Len(tFecha.Text)
    Status.SimpleText = " Ingrese la fecha de envío. [Enter] Consulta"
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cCamion.ListIndex = -1 Then
            MsgBox "Debe seleccionar un camión.", vbExclamation, "ATENCIÓN"
            cCamion.SetFocus: Exit Sub
        End If
        If Not IsDate(tFecha.Text) Then MsgBox "La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN": Exit Sub
        tFecha.Text = Format(tFecha.Text, FormatoFP)
        AccionConsultar
    End If
End Sub
Private Sub tFecha_LostFocus()
    Status.SimpleText = ""
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "d-Mmm-yyyy")
End Sub
Private Sub AccionConsultar()
On Error GoTo ErrAC
Dim aValor As Long
    Screen.MousePointer = vbHourglass
    AccionLimpiar
    Cons = "Select EnvCodigo, EnvDocumento, EnvBulto, EnvCliente,  EnvDireccion, EnvAgencia,  AgeNombre" _
        & " From Envio, Agencia " _
        & " Where EnvFechaPrometida = '" & Format(tFecha.Text, "mm/dd/yyyy") & "'" _
        & " And EnvEstado = " & EstadoEnvio.AImprimir _
        & " And EnvHabilitado = 1 And EnvCamion = " & cCamion.ItemData(cCamion.ListIndex) _
        & " And EnvAgencia = AgeCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
        With vsEnvio
            .AddItem RsAux!EnvCodigo
            If Not IsNull(RsAux!EnvBulto) Then .Cell(flexcpText, .Rows - 1, 1) = RsAux!EnvBulto Else .Cell(flexcpText, .Rows - 1, 1) = 0
            If Not IsNull(RsAux!AgeNombre) Then .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!AgeNombre)
            If Not IsNull(RsAux!EnvDireccion) Then
                .Cell(flexcpText, .Rows - 1, 3) = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!EnvDireccion, True, True, False, True, True)
            End If
            aValor = RsAux!EnvDocumento: .Cell(flexcpData, .Rows - 1, 0) = aValor
            aValor = RsAux!EnvCliente: .Cell(flexcpData, .Rows - 1, 1) = aValor
            aValor = RsAux!EnvAgencia: .Cell(flexcpData, .Rows - 1, 2) = aValor
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    If vsEnvio.Rows > 1 Then
        vsEnvio.SetFocus
        vsEnvio.Select 1, 0, 1, vsEnvio.Cols - 1
        If CInt(vsEnvio.Cell(flexcpText, vsEnvio.Row, 1)) > 0 Then
            cmdEliminar.Enabled = True
        Else
            cmdNuevo.Enabled = True
        End If
        cmdRemito.Enabled = True
        AccionConsultarArticulos CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0))
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrAC:
    clsGeneral.OcurrioError "Ocurrio un error al consultar."
    Screen.MousePointer = vbDefault
End Sub
Private Sub AccionLimpiar()
    vsEnvio.Rows = 1
    vsArticulo.Rows = 1: vsArticulo.BackColor = Inactivo
    cmdRemito.Enabled = False
    cmdImprimir.Enabled = False
    cmdNuevo.Enabled = False
    cmdEliminar.Enabled = False
    labBultos.Caption = vbNullString
End Sub
Private Sub cCamion_GotFocus()
    cCamion.SelStart = 0: cCamion.SelLength = Len(cCamion.Text)
    Status.SimpleText = " Seleccione el camión que contiene los envíos."
End Sub
Private Sub cCamion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cCamion.ListIndex > -1 Then Foco tFecha
End Sub
Private Sub cCamion_LostFocus()
    cCamion.SelStart = 0: Status.SimpleText = ""
End Sub
Private Sub AccionConsultarArticulos(lnEnvio As Long)
On Error GoTo ErrACA
Dim aValor As Long
    Screen.MousePointer = vbHourglass
    vsArticulo.Rows = 1
    Cons = "Select REvAEntregar, REvCantidad, ArtID, ArtCodigo, ArtNombre From RenglonEnvio, Articulo" _
        & " Where REvEnvio = " & lnEnvio _
        & " And REvArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        'Solo cargo los que tengan a enviar.
        If RsAux!RevAEntregar > 0 Then
            With vsArticulo
                .AddItem "0"
                .Cell(flexcpText, .Rows - 1, 1) = RsAux!RevAEntregar
                .Cell(flexcpText, .Rows - 1, 2) = RsAux!REvCantidad
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ArtCodigo, "( #,000,000) ") & Trim(RsAux!ArtNombre)
                aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            End With
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    Screen.MousePointer = vbDefault
    Exit Sub
ErrACA:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los artículos del envío."
    Screen.MousePointer = vbDefault
End Sub
Private Sub AccionImprimir()
On Error GoTo ErrAI
Dim sImprimio As Boolean

    sImprimio = False
    Screen.MousePointer = vbHourglass
    'Mando primero el encabezado y luego cargo los artículos.
    If EncabezadoImpresion("Etiqueta de Bulto", vsEnvio.Cell(flexcpText, vsEnvio.Row, 3), vsEnvio.Cell(flexcpData, vsEnvio.Row, 1)) Then
        With vsListado
            .TableBorder = tbAll
            .AddTable "1100|5000", "Cantidad|Artículo", ""
            For I = 1 To vsArticulo.Rows - 1
                If CInt(vsArticulo.Cell(flexcpText, I, 0)) > 0 Then
                    sImprimio = True
                    .AddTable "1100|5000", "", CInt(vsArticulo.Cell(flexcpText, I, 0)) & "|" & Trim(vsArticulo.Cell(flexcpText, I, 3))
                End If
            Next
            If sImprimio Then .Paragraph = "": .Paragraph = "Factura: " & BuscoDocumento(vsEnvio.Cell(flexcpData, vsEnvio.Row, 0))
            .EndDoc
            If sImprimio Then
                .Device = paIConformeN
                .PaperBin = paIConformeB
                .PrintDoc
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        If sImprimio Then
            labBultos.Caption = CInt(labBultos.Caption) + 1
            'Recorro la lista y quito los artículos impresos.
            UpdateoLista
        End If
        If vsArticulo.Rows = 1 And CLng(labBultos.Caption) > 0 Then AccionGrabar
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrAI:
    clsGeneral.OcurrioError "Ocurrio un error al intentar imprimir."
    Screen.MousePointer = vbDefault
End Sub

Private Function EncabezadoImpresion(Titulo As String, Direccion As String, lnCliente As Long) As Boolean
On Error GoTo errEI
    
    SeteoImpresoraPorDefecto paIConformeN
    With vsListado
        '.Device = paIConformeN
        '.PaperBin = paIConformeB
        .Orientation = orPortrait
        .PaperSize = 256
        .PaperHeight = 7920 '.PaperHeight / 2
        '.MarginTop = 300
        '.MarginLeft = 500
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Function
        End If
        
        .filename = "Etiqueta de Agencia"
        
        .TableBorder = tbNone
        .FontBold = True
        .TextAlign = taRightBaseline
        'Le pongo cinco espacios en blanco.
        .FontSize = 9.5
        .Paragraph = ""
        .AddTable ">4000", Titulo & Space(5), ""
        .FontSize = 8.25
        .TextAlign = taLeftBaseline
    
        .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = ""
        .Paragraph = "Dirección: " & Trim(Direccion)
        .Paragraph = "Nombre:  " & BuscoCliente(lnCliente)
        .Paragraph = "Agencia:  " & Trim(vsEnvio.Cell(flexcpText, vsEnvio.Row, 2))
        .Paragraph = ""
        .FontBold = False
    End With
    EncabezadoImpresion = True
    Exit Function
errEI:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar el objeto de impresión.", Err.Description
    EncabezadoImpresion = False
End Function
Private Sub UpdateoLista()
    'Saco los que no tienen mas artículos para dar a un bulto.
    With vsArticulo
        I = 1
        Do While I <= .Rows - 1
            If CInt(.Cell(flexcpText, I, 0)) > 0 Then
                If CInt(.Cell(flexcpText, I, 1)) - CInt(.Cell(flexcpText, I, 0)) > 0 Then
                    .Cell(flexcpText, I, 1) = CInt(.Cell(flexcpText, I, 1)) - CInt(.Cell(flexcpText, I, 0))
                    .Cell(flexcpText, I, 0) = 0
                Else
                    .RemoveItem I: I = I - 1
                End If
            End If
            I = I + 1
        Loop
    End With
End Sub
Private Sub AccionEliminar()
On Error GoTo ErrAE
    Screen.MousePointer = vbHourglass
    Cons = "Select * From Envio Where EnvCodigo = " & CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0))
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "El envío seleccionado fue eliminado, verifique.", vbExclamation, "ATENCIÓN"
        'Quito el envío de la lista.
        vsEnvio.RemoveItem vsEnvio.Row
        vsArticulo.Rows = 1
        cmdNuevo.Enabled = False
        cmdEliminar.Enabled = False
        cmdImprimir.Enabled = False
    Else
        If RsAux!EnvBulto = CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 1)) Then
            RsAux.Edit
            RsAux!EnvBulto = Null
            RsAux.Update
            RsAux.Close
            cmdEliminar.Enabled = False
            cmdNuevo.Enabled = True
            vsEnvio.Cell(flexcpText, vsEnvio.Row, 1) = 0
        Else
            MsgBox "Otra terminal pudo modificar la cantidad de bultos.", vbInformation, "ATENCIÓN"
            vsEnvio.Cell(flexcpText, vsEnvio.Row, 1) = RsAux!EnvBulto
            RsAux.Close
            If CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 1)) > 0 Then
                cmdEliminar.Enabled = True
                cmdNuevo.Enabled = False
            Else
                cmdEliminar.Enabled = False
                cmdNuevo.Enabled = True
            End If
            AccionConsultarArticulos CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0))
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrAE:
    clsGeneral.OcurrioError "Ocurrio un error al intentar eliminar la información."
    Screen.MousePointer = vbDefault
End Sub
Private Function AccionGrabar() As Boolean
On Error GoTo ErrAG
    Screen.MousePointer = vbHourglass
    If CLng(labBultos.Caption) = 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe imprimir por lo menos un bulto para poder grabar.", vbExclamation, "ATENCIÓN"
        AccionGrabar = False
        Exit Function
    End If
    If MsgBox("¿Confirma almacenar la información de bultos?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
        Cons = "Select * From Envio Where EnvCodigo = " & CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0))
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            RsAux.Close
            MsgBox "El envío seleccionado fue eliminado, verifique." & Chr(13) & "No se almacenará la información impresa.", vbExclamation, "ATENCIÓN"
            'Quito el envío de la lista.
            vsEnvio.RemoveItem vsEnvio.Row: vsArticulo.Rows = 1
            cmdNuevo.Enabled = False
            cmdEliminar.Enabled = False
            cmdImprimir.Enabled = False
        Else
            If RsAux!EnvBulto = CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 1)) Or IsNull(RsAux!EnvBulto) Then
                RsAux.Edit
                RsAux!EnvBulto = CInt(labBultos.Caption)
                RsAux.Update
                RsAux.Close
                cmdEliminar.Enabled = True
                cmdEliminar.Caption = "El&iminar"
                cmdNuevo.Enabled = False
                cmdNuevo.Caption = "&Nuevo"
                vsEnvio.Cell(flexcpText, vsEnvio.Row, 1) = labBultos.Caption
                labBultos.Caption = vbNullString
                cmdImprimir.Enabled = False
                vsEnvio.Enabled = True
                vsArticulo.BackColor = Inactivo
            Else
                MsgBox "Otra terminal pudo modificar la cantidad de bultos del envío, verifique.", vbInformation, "ATENCIÓN"
                vsEnvio.Cell(flexcpText, vsEnvio.Row, 1) = RsAux!EnvBulto
                RsAux.Close
                cmdNuevo.Caption = "&Nuevo"
                cmdEliminar.Cancel = "El&iminar"
                If CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 1)) > 0 Then
                    cmdEliminar.Enabled = True
                    cmdNuevo.Enabled = False
                Else
                    cmdEliminar.Enabled = False
                    cmdNuevo.Enabled = True
                End If
                labBultos.Caption = vbNullString
                cmdImprimir.Enabled = False
                vsEnvio.Enabled = True
                vsArticulo.BackColor = Inactivo
                AccionConsultarArticulos CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0))
            End If
        End If
        AccionGrabar = True
    End If
    Screen.MousePointer = vbDefault
    Exit Function
ErrAG:
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información."
    Screen.MousePointer = vbDefault
End Function
Private Function ObtengoNombreCliente(lnCliente As Long) As String

    Cons = "Select * From Cliente" _
            & " Left Outer Join CPersona ON CliCodigo = CPeCliente" _
            & " Left Outer Join CEmpresa ON CliCodigo = CEmCliente" _
        & " Where CliCodigo = " & lnCliente
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAux!CliTipo = TipoCliente.Cliente Then
        ObtengoNombreCliente = Trim(Trim(Format(RsAux!CPeApellido1, "#")) & " " & Trim(Format(RsAux!CPeApellido2, "#"))) & ", " & Trim(Trim(Format(RsAux!CPeNombre1, "#")) & " " & Trim(Format(RsAux!CPeNombre2, "#")))
    Else
        If Not IsNull(RsAux!CEmNombre) Then
            ObtengoNombreCliente = RsAux!CEmNombre
        Else
            ObtengoNombreCliente = RsAux!CEmFantasia
        End If
    End If
    RsAux.Close
    
    ObtengoNombreCliente = Trim(ObtengoNombreCliente)
    
End Function


'-------------------------------------------------------------------------------------------------------
'   Carga un string con todos los articulos que corresponden a los fletes.
'   Se utiliza en aquellos formularios que no filtren los fletes
'-------------------------------------------------------------------------------------------------------
Private Function CargoArticulosDeFlete() As String
Dim Fletes As String
    On Error GoTo errCargar
    Fletes = ""
    
    'Cargo los articulos a descartar-----------------------------------------------------------
    Cons = "Select Distinct(TFlArticulo) from TipoFlete Where TFlArticulo <> Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        Fletes = Fletes & RsAux!TFlArticulo & ","
        RsAux.MoveNext
    Loop
    RsAux.Close
    Fletes = Fletes & paArticuloPisoAgencia & "," & paArticuloDiferenciaEnvio & ","
    '----------------------------------------------------------------------------------------------
    CargoArticulosDeFlete = Fletes
    Exit Function
    
errCargar:
    CargoArticulosDeFlete = Fletes
End Function

Private Sub vsArticulo_GotFocus()
    Status.SimpleText = " Ingrese la cantidad de artículos que formarán el bulto."
End Sub

Private Sub vsArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrLA

    If vsArticulo.Rows = 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyReturn: If cmdImprimir.Enabled Then AccionImprimir
        Case vbKeyAdd
            If vsArticulo.BackColor = vbWhite Then
                If CInt(vsArticulo.Cell(flexcpText, vsArticulo.Row, 0)) + 1 <= CInt(vsArticulo.Cell(flexcpText, vsArticulo.Row, 1)) Then
                    vsArticulo.Cell(flexcpText, vsArticulo.Row, 0) = CInt(vsArticulo.Cell(flexcpText, vsArticulo.Row, 0)) + 1
                End If
            End If
        Case vbKeySubtract
            If vsArticulo.BackColor = vbWhite Then
                If CInt(vsArticulo.Cell(flexcpText, vsArticulo.Row, 0)) > 0 Then vsArticulo.Cell(flexcpText, vsArticulo.Row, 0) = CInt(vsArticulo.Cell(flexcpText, vsArticulo.Row, 0)) - 1
            End If
    End Select
    Exit Sub
ErrLA:
    clsGeneral.OcurrioError "Ocurrio un error inesperado. ", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub vsEnvio_RowColChange()
    If vsEnvio.Rows = 1 Then Exit Sub
    cmdImprimir.Enabled = False: cmdNuevo.Enabled = False: cmdEliminar.Enabled = False
    If CInt(vsEnvio.Cell(flexcpText, vsEnvio.Row, 1)) > 0 Then
        cmdEliminar.Enabled = True
    Else
        cmdNuevo.Enabled = True
    End If
    AccionConsultarArticulos CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0))
End Sub
Private Sub PrueboBandejaImpresora()
On Error GoTo ErrPBI
    With vsListado
        .PageBorder = pbNone
'        .Device = paIConformeN
'        If .Device <> paIConformeN Then MsgBox "Ud no tiene instalada la impresora para imprimir Conformes. Avise al administrador.", vbExclamation, "ATENCIÒN"
'        If .PaperBins(paIConformeB) Then .PaperBin = paIConformeB Else MsgBox "Esta mal definida la bandeja de conformes en su sucursal, comuniquele al administrador.", vbInformation, "ATENCIÓN": paIConformeB = .PaperBin
'        .PaperSize = 256 'Hoja carta
        .Orientation = orPortrait
       ' .PaperHeight = .PaperHeight / 2
        .MarginTop = 300
        .MarginLeft = 500
    End With
    Exit Sub
ErrPBI:
    clsGeneral.OcurrioError "Ocurrio un error al setear la impresora, consulte con el administrador de impresión este problema.", Err.Description
End Sub

Private Function BuscoCliente(IdCliente As Long) As String
On Error GoTo ErrBC
    BuscoCliente = ""
    Cons = "Select * from Cliente " _
            & " Left Outer Join CPersona ON CliCodigo = CPeCliente " _
            & " Left Outer Join CEmpresa ON CliCodigo = CEmCliente " _
       & " Where CliCodigo = " & IdCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        Select Case RsAux!CliTipo
            Case TipoCliente.Cliente
                BuscoCliente = Trim(Trim(Format(RsAux!CPeNombre1, "#")) & " " & Trim(Format(RsAux!CPeNombre2, "#"))) & ", " & Trim(Trim(Format(RsAux!CPeApellido1, "#")) & " " & Trim(Format(RsAux!CPeApellido2, "#")))
            Case TipoCliente.Empresa
                If Not IsNull(RsAux!CEmNombre) Then BuscoCliente = Trim(RsAux!CEmFantasia)
                If Not IsNull(RsAux!CEmFantasia) Then BuscoCliente = BuscoCliente & " (" & Trim(RsAux!CEmFantasia) & ")"
        End Select
    End If
    RsAux.Close
    Exit Function
ErrBC:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el cliente para imprimir en la ficha.", Err.Description
End Function
Private Function BuscoDocumento(idDocumento As Long) As String
 On Error GoTo ErrBD
    BuscoDocumento = ""
    Cons = "Select * From Documento, Sucursal" _
        & " Where DocCodigo = " & idDocumento _
        & " And DocSucursal = SucCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        BuscoDocumento = Trim(RsAux!SucAbreviacion) & " " & RetornoNombreDocumento(RsAux!DocTipo, True) & " " & Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
    End If
    RsAux.Close
    Exit Function
ErrBD:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el documento para imprimir en la ficha.", Err.Description
End Function

