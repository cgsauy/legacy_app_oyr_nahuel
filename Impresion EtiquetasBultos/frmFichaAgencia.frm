VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
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
   Icon            =   "frmFichaAgencia.frx":0000
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   1986860
      ForeColorFixed  =   16777215
      BackColorSel    =   5273691
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
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
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10001
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Picture         =   "frmFichaAgencia.frx":0742
            Key             =   "printer"
            Object.Tag             =   ""
         EndProperty
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
      Preview         =   0   'False
      PhysicalPage    =   -1  'True
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
      BackColor       =   &H00E0E0E0&
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
   Begin VB.Menu MnuEtiquetas 
      Caption         =   "QEtiquetas"
      Visible         =   0   'False
      Begin VB.Menu MnuEtiTitulo 
         Caption         =   "Etiquetas"
      End
      Begin VB.Menu MnuEtiLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEtiOcasional 
         Caption         =   "Editar cantidad ocasional"
      End
      Begin VB.Menu MnuEtiPermanente 
         Caption         =   "Editar cantidad permanente"
      End
   End
End
Attribute VB_Name = "FichaAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TipoCliente
    Cliente = 1
    Empresa = 2
End Enum

'Definicion de Tipos de Envios------------------------------------------------------------------------------------
Private Enum TipoEnvio
    Entrega = 1
    Service = 2
    Cobranza = 3
End Enum

Private Enum TipoPagoEnvio
    PagaAhora = 1
    PagaDomicilio = 2
    FacturaCamión = 3
End Enum

Private Enum EstadoEnvio
    AImprimir = 0
    AConfirmar = 1
    Rebotado = 2
    Impreso = 3
    Entregado = 4
    Anulado = 5
End Enum
'-----------------------------------------------------------------------------------------------------------------------

Private strArtFlete As String
Public sCodBarra As String, sMoneda As String

Private Sub cCamion_Click()
    AccionLimpiar
End Sub

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
        AccionConsultarArticulos CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)), Val(vsEnvio.Cell(flexcpData, vsEnvio.Row, 3))
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
Dim strDocumentos As String, sDir As String
On Error GoTo ErrIR
    Screen.MousePointer = vbHourglass
    strDocumentos = vbNullString
    'Saco los artículos de flete.---------------------------------------------------------
    'ATENCION: Tengo que controlar el piso para indicar si paga o no piso.
    If sTodos Then
        For I = 1 To vsEnvio.Rows - 1
            'Saco el código de documento y verifico si no lo imprimi.
'            If InStr(strDocumentos, vsEnvio.Cell(flexcpData, I, 0) & ",") = 0 Then
'                strDocumentos = strDocumentos & vsEnvio.Cell(flexcpData, I, 0) & ","
                
            '24/04/2007 modifique pasar el texto de la direccion y armo string con enter.
                If CLng(vsEnvio.Cell(flexcpData, I, 4)) > 0 Then
                    sDir = clsGeneral.ArmoDireccionEnTexto(cBase, CLng(vsEnvio.Cell(flexcpData, vsEnvio.Row, 4)), True, True, False, True, True, False, False)
                    'sDir = UCase(Mid(sDir, 1, InStr(1, sDir, vbCr) - 1)) & Mid(sDir, InStr(1, sDir, vbCrLf))
                Else
                    sDir = vsEnvio.Cell(flexcpText, I, 3)
                End If
                ImprimoRemitosAgencia CLng(vsEnvio.Cell(flexcpData, I, 2)), CLng(vsEnvio.Cell(flexcpData, I, 0)), CLng(vsEnvio.Cell(flexcpText, I, 0)), CLng(vsEnvio.Cell(flexcpData, I, 1)), sDir
'            End If
        Next
    Else
        'Saco el código de documento.------------------------------------
        If CLng(vsEnvio.Cell(flexcpData, vsEnvio.Row, 4)) > 0 Then
            sDir = clsGeneral.ArmoDireccionEnTexto(cBase, CLng(vsEnvio.Cell(flexcpData, vsEnvio.Row, 4)), True, True, False, True, True, False, False)
            'sDir = UCase(Mid(sDir, 1, InStr(1, sDir, vbCr) - 1)) & Mid(sDir, InStr(1, sDir, vbCrLf))
        Else
            sDir = vsEnvio.Cell(flexcpText, vsEnvio.Row, 3)
        End If
        ImprimoRemitosAgencia CLng(vsEnvio.Cell(flexcpData, vsEnvio.Row, 2)), CLng(vsEnvio.Cell(flexcpData, vsEnvio.Row, 0)), CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)), CLng(vsEnvio.Cell(flexcpData, vsEnvio.Row, 1)), sDir
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrIR:
    clsGeneral.OcurrioError "Ocurrió un error al intentar imprimir los remitos."
    Screen.MousePointer = vbDefault
End Sub
Private Sub ImprimoRemitosAgencia(Agencia As Long, lnDocumento As Long, lnEnvio As Long, lnCliente As Long, _
                                                            Direccion As String)
Dim iCant As Integer, cMonto As Currency
Dim rs As rdoResultset
Dim sNoPagaPiso As Boolean, strTelef As String, strComentario As String
    On Error GoTo ErrIRA
    cMonto = 0
    Cons = "Select  RevAEntregar as Cantidad, ArtID, ArtCodigo, ArtNombre, RenPrecio As Pre" _
        & " From Envio, RenglonEnvio, Renglon, Articulo" _
        & " Where EnvTipo = " & TipoEnvio.Entrega & "  And EnvCodigo = " & lnEnvio & " And RenDocumento = " & lnDocumento _
        & " And RevAEntregar > 0 And EnvCodigo = RevEnvio And RevArticulo = RenArticulo And RenArticulo = ArtID "
    
    Set rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If rs.EOF Then
        rs.Close
        
        Cons = "Select  RevAEntregar as Cantidad, ArtID, ArtCodigo, ArtNombre, RVTPrecio as Pre" _
            & " From Envio, RenglonEnvio, RenglonVtaTelefonica, Articulo " _
            & " Where EnvCodigo = " & lnEnvio & " And EnvDocumento = RVTVentaTelefonica  " _
            & " And RevAEntregar > 0 And EnvCodigo = RevEnvio And RevArticulo = RVTArticulo And RVTArticulo = ArtID "
        
        Set rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
        If rs.EOF Then
            rs.Close
            MsgBox "No se encontró la información de los artículos, o el mismo es un flete de servicio.", vbExclamation, "ATENCIÓN": Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    
    
    Dim sNCli As String
    Dim arrArt() As String
    sNCli = BuscoCliente(lnCliente)
    ReDim arrArt(0)
    Do While Not rs.EOF
        ReDim Preserve arrArt(UBound(arrArt) + 1)
        cMonto = cMonto + (rs!Cantidad * rs!Pre)
        arrArt(UBound(arrArt)) = rs!Cantidad & "|" & Format(rs!ArtCodigo, "(#,000,000 )") & Trim(rs!ArtNombre)
        rs.MoveNext
    Loop
    rs.Close
    
    
    Dim cMontoEnvio As Currency
    'Agrego el telefono y el comentario de envío si tiene.
    strTelef = "": strComentario = ""
    Cons = "Select * From Envio Left outer join Moneda On EnvMoneda = MonCodigo Where EnvCodigo = " & lnEnvio
    Set rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rs.EOF Then
        If Not IsNull(rs!EnvTelefono) And strTelef = "" Then strTelef = rs!EnvTelefono
        If Not IsNull(rs!EnvComentario) Then strComentario = Trim(rs!EnvComentario)
        If rs!EnvFormaPago = TipoPagoEnvio.FacturaCamión Then
            'Tengo que sumar este monto ya que lo paga en domicilio.
            If Not IsNull(rs!EnvValorFlete) Then cMontoEnvio = rs!EnvValorFlete
        End If
        If Not IsNull(rs!MonSigno) Then sMoneda = Trim(rs!MonSigno)
    End If
    rs.Close
            
    Dim iQ As Byte, iQ2 As Integer
    For iQ = 1 To 3
        
        If EncabezadoImpresion("Remito de Agencia", sNCli, Direccion) Then
            With vsListado
                
                .TableBorder = tbAll
                .AddTable "1100|7000", "Cantidad|Artículo", "", &HF0F0F0
                
                For iQ2 = 1 To UBound(arrArt)
                    .AddTable "1100|7000", "", arrArt(iQ2)
                Next
                
                'Imprimo datos de la factura.....................................
                .Paragraph = ""
                .TableBorder = tbNone
                If Val(vsEnvio.Cell(flexcpData, vsEnvio.Row, 3)) = 2 Then
                    .Paragraph = "Servicio:     " & vsEnvio.Cell(flexcpData, vsEnvio.Row, 0) & Space(10) & "Valor: " & sMoneda & " " & Format(cMonto, FormatoMonedaP)
                Else
                    .Paragraph = "Factura:      " & BuscoDocumento(vsEnvio.Cell(flexcpData, vsEnvio.Row, 0)) & Space(10) & "Valor: " & sMoneda & " " & Format(cMonto, FormatoMonedaP)
                End If
                .Paragraph = ""
                .AddTable "1100|9000", "Envío:|" & lnEnvio, ""
                '.Paragraph = "Envío: " & lnEnvio
                
                sCodBarra = "": sMoneda = ""
                .Font = "tahoma": .FontSize = 8.25
                
                If cMontoEnvio > 0 Then
                    .FontBold = True
                    .FontSize = 10
                    If sNoPagaPiso Then
                        .Paragraph = "COBRAR: " & sMoneda & " " & Format(cMonto, "#,###.00") & Space(25) & "NO PAGA PISO"
                    Else
                        .Paragraph = "COBRAR: " & sMoneda & " " & Format(cMonto, "#,###.00")
                    End If
                    .FontBold = False
                    .FontSize = 8.25
                End If
                '.Paragraph = ""
                If Trim(strTelef) <> "" Then .AddTable "1100|9000", "Teléfono:|" & strTelef, ""
                .Paragraph = ""
                If Trim(strComentario) <> "" Then .AddTable "1100|9000", "Comentario:|" & strComentario, ""
                .EndDoc
                'COMO TENGO SETEADO el Preview a false manda a impresora directamente.
'                .PrintDoc False
            End With
        End If
    Next
    Exit Sub
ErrIRA:
    clsGeneral.OcurrioError "Error al intentar imprimir los remitos para el envío." & lnEnvio
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
    Status.Panels("printer").Text = paPrintConfD
    ObtengoSeteoForm Me, Me.Left, , Me.Width
    With vsEnvio
        .Redraw = False
        .Rows = 1: .ExtendLastCol = True: .Cols = 1
        .FormatString = "<Envío|>Bultos|<Agencia|<Localidad - Dirección|"
        .ColWidth(0) = 900
        .ColWidth(2) = 1500
        .ColHidden(4) = True
        .Redraw = True
    End With
    With vsArticulo
        .Redraw = False
        .Rows = 1: .ExtendLastCol = True: .Cols = 1
        .FormatString = ">Cantidad|>A Enviar|>Total|>Etiquetas|<Articulo"
        .Redraw = True
    End With
    AccionLimpiar
    FechaDelServidor
    tFecha.Text = Format(gFechaServidor, FormatoFP)
    
    Cons = "Select CamCodigo, CamNombre From Camion order by CamNombre"
    CargoCombo Cons, cCamion
    
    'Saco los artículos de flete.
    strArtFlete = CargoArticulosDeFlete
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al iniciar el formulario. " & Trim(Err.Description)
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

Private Sub MnuEtiOcasional_Click()
    loc_EditQEtiquetas False
End Sub

Private Sub MnuEtiPermanente_Click()
    loc_EditQEtiquetas True
End Sub

Private Sub Status_PanelClick(ByVal Panel As ComctlLib.Panel)
    If "printer" = Panel.Key Then
        prj_GetPrinter True
        Panel.Text = paPrintConfD
    End If
End Sub

Private Sub tFecha_Change()
    If tFecha.Tag <> "" Then
        AccionLimpiar
        tFecha.Tag = ""
    End If
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0
    tFecha.SelLength = Len(tFecha.Text)
    Status.Panels(1) = " Ingrese la fecha de envío. [Enter] Consulta"
End Sub

Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            With tFecha
                If IsDate(.Text) Then .Text = Format(CDate(.Text) + 1, FormatoFP): .Tag = .Text
                .SelStart = 0: .SelLength = Len(.Text)
                KeyCode = 0
            End With
        
        Case vbKeyDown
            With tFecha
                If IsDate(.Text) Then .Text = Format(CDate(.Text) - 1, FormatoFP): .Tag = .Text
                .SelStart = 0: .SelLength = Len(.Text)
            End With
            KeyCode = 0
    End Select
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cCamion.ListIndex = -1 Then
            MsgBox "Debe seleccionar un camión.", vbExclamation, "ATENCIÓN"
            cCamion.SetFocus: Exit Sub
        End If
        If Not IsDate(tFecha.Text) Then
            MsgBox "La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN": Exit Sub
        Else
            tFecha.Text = Format(tFecha.Text, FormatoFP)
            tFecha.Tag = tFecha.Text
        End If
        AccionConsultar
    End If
End Sub
Private Sub tFecha_LostFocus()
    Status.Panels(1) = ""
'    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "d-Mmm-yyyy")
End Sub
Private Sub AccionConsultar()
On Error GoTo ErrAC
Dim aValor As Long
    
    Screen.MousePointer = vbHourglass
    AccionLimpiar
    Cons = "Select EnvCodigo, EnvDocumento, EnvBulto, EnvCliente,  EnvDireccion, EnvAgencia,  AgeNombre, EnvTipo" _
        & " From Envio, Agencia " _
        & " Where EnvFechaPrometida = '" & Format(tFecha.Text, "mm/dd/yyyy") & "'" _
        & " And EnvEstado = " & EstadoEnvio.AImprimir _
        & " And EnvCamion = " & cCamion.ItemData(cCamion.ListIndex) _
        & " And EnvAgencia = AgeCodigo And EnvDocumento Is Not Null Order by EnvCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
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
            aValor = RsAux!EnvTipo: .Cell(flexcpData, .Rows - 1, 3) = aValor
            If Not IsNull(RsAux!EnvDireccion) Then
                aValor = RsAux!EnvDireccion: .Cell(flexcpData, .Rows - 1, 4) = aValor
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    If vsEnvio.Rows > 1 Then
        If Not vsEnvio.Enabled Then vsEnvio.Enabled = True
        vsEnvio.SetFocus
        vsEnvio.Select 1, 0, 1, vsEnvio.Cols - 1
        If CInt(vsEnvio.Cell(flexcpText, vsEnvio.Row, 1)) > 0 Then
            cmdEliminar.Enabled = True
        Else
            cmdNuevo.Enabled = True
        End If
        cmdRemito.Enabled = True
        AccionConsultarArticulos CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)), Val(vsEnvio.Cell(flexcpData, vsEnvio.Row, 3))
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrAC:
    clsGeneral.OcurrioError "Ocurrió un error al consultar."
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
    Status.Panels(1) = " Seleccione el camión que contiene los envíos."
End Sub
Private Sub cCamion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cCamion.ListIndex > -1 Then Foco tFecha
End Sub
Private Sub cCamion_LostFocus()
    cCamion.SelStart = 0: Status.Panels(1) = ""
End Sub
Private Sub AccionConsultarArticulos(lnEnvio As Long, iTipoEnvio As Integer)
On Error GoTo ErrACA
Dim aValor As Long
    Screen.MousePointer = vbHourglass
    vsArticulo.Rows = 1
    'Tengo que determinar el tipo de documento asignado en el envío
    'Si es remito --> de ahí me voy al contado si no saco del contado.
    
    If iTipoEnvio = 3 Then
        Cons = "Select  REvAEntregar, REvCantidad, ArtID, ArtCodigo, ArtNombre, IsNull(AFaEtiqAgencia, 0) as QEti, RVTPrecio as Pre" _
            & " From Envio, RenglonEnvio, RenglonVtaTelefonica, Articulo " _
            & " LEFT OUTER JOIN ArticuloFacturacion On ArtID = AFaArticulo " _
            & " Where EnvCodigo = " & lnEnvio & " And EnvDocumento = RVTVentaTelefonica  " _
            & " And RevAEntregar > 0 And EnvCodigo = RevEnvio And RevArticulo = RVTArticulo And RVTArticulo = ArtID "
    Else
        Cons = "SELECT DocTipo From Documento WHERE DocCodigo = (SELECT EnvDocumento FROM Envio Where EnvCodigo = " & lnEnvio & ")"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux(0) = 6 Then
            Cons = "Select REvAEntregar, REvCantidad, ArtID, ArtCodigo, ArtNombre, IsNull(AFaEtiqAgencia, 0) as QEti, RenPrecio As Pre " & _
                " FROM Envio, RenglonEnvio, Renglon, RemitoDocumento, Articulo " & _
                " LEFT OUTER JOIN ArticuloFacturacion On ArtID = AFaArticulo " & _
                " WHERE EnvCodigo = " & lnEnvio & " And EnvCodigo = RevEnvio And REvArticulo = ArtID " & _
                " And RevArticulo = RenArticulo And RenDocumento = RDODocumento AND EnvDocumento = RDORemito"
        Else
            'saco del renglón del contado .
            Cons = "Select REvAEntregar, REvCantidad, ArtID, ArtCodigo, ArtNombre, IsNull(AFaEtiqAgencia, 0) as QEti, RenPrecio As Pre " & _
                " FROM Envio INNER JOIN RenglonEnvio ON EnvCodigo = RevEnvio " & _
                "INNER JOIN Articulo ON REvArticulo = ArtID " & _
                "LEFT OUTER JOIN Renglon ON EnvDocumento = RenDocumento AND RenArticulo = ArtID " & _
                "LEFT OUTER JOIN ArticuloFacturacion On ArtID = AFaArticulo " & _
                "WHERE EnvCodigo = " & lnEnvio
        End If
        RsAux.Close
    End If
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    
    Do While Not RsAux.EOF
        'Solo cargo los que tengan a enviar.
        If RsAux!RevAEntregar > 0 Then
            With vsArticulo
                .AddItem "0"
                .Cell(flexcpText, .Rows - 1, 1) = RsAux!RevAEntregar
                .Cell(flexcpText, .Rows - 1, 2) = RsAux!REvCantidad
                .Cell(flexcpText, .Rows - 1, 3) = RsAux!QEti
                .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!ArtCodigo, "( #,000,000) ") & Trim(RsAux!ArtNombre)
                aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                aValor = 0
                If Not IsNull(RsAux("Pre")) Then
                    aValor = RsAux("Pre")
                End If
                .Cell(flexcpData, .Rows - 1, 4) = aValor
            End With
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    Screen.MousePointer = vbDefault
    Exit Sub
ErrACA:
    clsGeneral.OcurrioError "Error al cargar los artículos del envío."
    Screen.MousePointer = vbDefault
End Sub

Private Sub ImprimoEtiquetaEnTicket(ByVal sNCliente As String, ByVal Direccion As String, ByVal strTelef As String, _
    ByVal strComentario As String, ByVal factura As String, ByRef arrArt() As String)
    
    SeteoImpresoraPorDefecto paPrintConfD
    With vsListado
        .FileName = "Etiqueta de Agencia"
        .AbortWindow = False
        .MarginTop = 190
        .MarginLeft = 100
        
        .Device = paPrintConfD
        .PaperSize = 258 ' paPrintConfPaperSize
        
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
        .FontSize = 10
        .AddTable "4000", "Etiqueta de bulto  - Envío: " & Val(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)) & " -", ""
        
        .Paragraph = ""
        
        .FontBold = False
        .TextAlign = taCenterTop
        .FontName = "3 of 9 Barcode"
        .FontSize = 24
        .Paragraph = "*EB" & Val(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)) & "*"
        
        .FontName = "Tahoma"
        .FontSize = 9
        .TextAlign = taLeftBaseline
        .Paragraph = ""
        .Paragraph = ""
        
        .TableBorder = tbNone
        .FontBold = True
        .AddTable "700|5000", "|" & sNCliente, ""
        
        
        .AddTable "6000", Trim(Direccion), ""
        
        .FontBold = False
        
        .Paragraph = ""
        .FontBold = False
        '.Paragraph = "Agencia:  " & Trim(vsEnvio.Cell(flexcpText, vsEnvio.Row, 2))

        .AddTable "820|2500|900|2500", "Agencia|" & Trim(vsEnvio.Cell(flexcpText, vsEnvio.Row, 2)) & "|Teléfono:|" & strTelef, ""
        If Trim(strComentario) <> "" Then .AddTable "1200|6000", "Comentario:|" & strComentario, ""
        
        .Paragraph = ""
        .FontSize = 8
        .TableBorder = tbAll
        .AddTable "^500|^3500|^1400", "Cant|Artículo|Unitario", ""
        
        For I = 1 To UBound(arrArt)
            .AddTable ">500|3500|>1400", "", arrArt(I)
        Next
        
        .TableBorder = tbNone
        .FontSize = 9
        If vsEnvio.Cell(flexcpData, vsEnvio.Row, 3) <> TipoEnvio.Service Then
            .Paragraph = "" ': .Paragraph = "Factura: " & sDoc
            .TableBorder = tbNone
            .AddTable "1100|9000", "Factura:|" & factura, ""
            .Paragraph = ""
            .Font = "tahoma"
            .FontSize = 8.25
        End If
        .EndDoc
        'Manda a impresora ya que tengo el preview = false
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
End Sub


Private Sub AccionImprimir()
On Error GoTo ErrAI
Dim sImprimio As Boolean

    sImprimio = False
    Screen.MousePointer = vbHourglass
    
    Dim strTelef As String, strComentario As String
    
    Dim arrArt() As String
    ReDim arrArt(0)
    
    For I = 1 To vsArticulo.Rows - 1
        If CInt(vsArticulo.Cell(flexcpText, I, 0)) > 0 Then
            ReDim Preserve arrArt(UBound(arrArt) + 1)
            arrArt(UBound(arrArt)) = CInt(vsArticulo.Cell(flexcpText, I, 0)) & "|" & Trim(vsArticulo.Cell(flexcpText, I, 4)) & "|" & Format(vsArticulo.Cell(flexcpData, I, 4), "#,##0.00")
        End If
    Next
    
    If UBound(arrArt) > 0 Then
        Dim sNCli As String, sDoc As String
        sNCli = BuscoCliente(Val(vsEnvio.Cell(flexcpData, vsEnvio.Row, 1)))
        sDoc = BuscoDocumento(vsEnvio.Cell(flexcpData, vsEnvio.Row, 0))
        
        strTelef = "": strComentario = ""
        Dim rs As rdoResultset
        Cons = "Select rtrim(EnvTelefono) EnvTelefono, rtrim(EnvComentario) EnvComentario From Envio Left outer join Moneda On EnvMoneda = MonCodigo Where EnvCodigo = " & Val(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0))
        Set rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not rs.EOF Then
            If Not IsNull(rs!EnvTelefono) Then strTelef = rs!EnvTelefono
            If Not IsNull(rs!EnvComentario) Then strComentario = Trim(rs!EnvComentario)
        End If
        rs.Close
        
        Dim iQ As Byte
        For iQ = 1 To vsArticulo.Cell(flexcpValue, vsArticulo.Row, 3)
        
            ImprimoEtiquetaEnTicket sNCli, vsEnvio.Cell(flexcpText, vsEnvio.Row, 3), strTelef, strComentario, sDoc, arrArt
        
        
        'Mando primero el encabezado y luego cargo los artículos.
'            If EncabezadoImpresion("Etiqueta de Bulto", sNCli, vsEnvio.Cell(flexcpText, vsEnvio.Row, 3)) Then
'
'                With vsListado
'                    If Trim(strTelef) <> "" Then
'                        .AddTable "1100|9000", "Envío:|" & Val(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)), ""
'                        .AddTable "1100|9000", "Teléfono:|" & strTelef, ""
'                        .Paragraph = ""
'                    End If
'
'                    If Trim(strComentario) <> "" Then
'                        .Paragraph = "Comentario: " & strComentario
'                        .Paragraph = ""
'                    End If
'
'                    .TableBorder = tbAll
'                    .AddTable "^1100|^5000|^1500", "Cantidad|Artículo|Precio unitario", "", &HF0F0F0
'                    For I = 1 To UBound(arrArt)
'                        .AddTable "1100|5000|>1500", "", arrArt(I)
'                    Next
'                    If vsEnvio.Cell(flexcpData, vsEnvio.Row, 3) <> TipoEnvio.Service Then
'                        .Paragraph = "" ': .Paragraph = "Factura: " & sDoc
'                        .TableBorder = tbNone
'                        .AddTable "1100|9000", "Factura:|" & sDoc, ""
'                        .Paragraph = ""
'                        .Font = "tahoma"
'                        .FontSize = 8.25
'                    End If
'                    .EndDoc
'                    'Manda a impresora ya que tengo el preview = false
'                End With        '----------------------------------------------------------------------------------------------------------------------------------------------
'            End If
        Next
        labBultos.Caption = CInt(labBultos.Caption) + 1
        'Recorro la lista y quito los artículos impresos.
        UpdateoLista
        If vsArticulo.Rows = 1 And CLng(labBultos.Caption) > 0 Then AccionGrabar
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrAI:
    clsGeneral.OcurrioError "Ocurrió un error al intentar imprimir.", Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Function EncabezadoImpresion(Titulo As String, ByVal sNomCliente As String, Direccion As String) As Boolean
On Error GoTo errEI
    
    SeteoImpresoraPorDefecto paPrintConfD
    With vsListado
        .FileName = "Etiqueta de Agencia"
        .AbortWindow = False
        
        .MarginTop = 500
        .MarginLeft = 500
        
        .Orientation = orLandscape
        .Device = paPrintConfD
        .PaperBin = paPrintConfB
        .PaperSize = paPrintConfPaperSize
        
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Function
        End If
        
        .TableBorder = tbNone
        .FontBold = True
        .TextAlign = taRightBaseline
        'Le pongo cinco espacios en blanco.  24/4/2007 cambie x 2 Juliana me dijo que lo corra 1 cm.
        .FontSize = 9.5
        .Paragraph = ""
        .Paragraph = ""
        .AddTable "^4500", Titulo, ""
        .FontSize = 9
        .TextAlign = taLeftBaseline
    
        .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = ""
        
        .TableBorder = tbNone
        .AddTable "1100|9500", "|" & sNomCliente, ""
        .AddTable "1100|9500", "Dirección:|" & Trim(Direccion), ""
        
'        .Paragraph = "           " & sNomCliente
'        .Paragraph = "Dirección: " & Trim(Direccion)
        .Paragraph = ""
        .FontBold = False
        .AddTable "1100|9000", "Agencia:|" & Trim(vsEnvio.Cell(flexcpText, vsEnvio.Row, 2)), ""
        '.Paragraph = "Agencia:  " & Trim(vsEnvio.Cell(flexcpText, vsEnvio.Row, 2))
        .Paragraph = ""
        
    End With
    EncabezadoImpresion = True
    Exit Function
errEI:
    clsGeneral.OcurrioError "Ocurrió un error al iniciar el objeto de impresión.", Err.Description
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
            AccionConsultarArticulos CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)), Val(vsEnvio.Cell(flexcpData, vsEnvio.Row, 3))
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrAE:
    clsGeneral.OcurrioError "Ocurrió un error al intentar eliminar la información."
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
                AccionConsultarArticulos CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)), Val(vsEnvio.Cell(flexcpData, vsEnvio.Row, 3))
            End If
        End If
        AccionGrabar = True
    End If
    Screen.MousePointer = vbDefault
    Exit Function
ErrAG:
    clsGeneral.OcurrioError "Ocurrió un error al intentar almacenar la información."
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
    Status.Panels(1) = " Ingrese la cantidad de artículos que formarán el bulto."
End Sub

Private Sub vsArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrLA

    If vsArticulo.Rows = 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyReturn: If cmdImprimir.Enabled Then AccionImprimir
        
        Case vbKeyMultiply
            If vsArticulo.BackColor = vbWhite Then
                'If Val(vsArticulo.Cell(flexcpText, vsArticulo.Row, 0)) > 0 Then
                    vsArticulo.Cell(flexcpText, vsArticulo.Row, 3) = Val(vsArticulo.Cell(flexcpText, vsArticulo.Row, 3)) + 1
                'End If
            End If
        Case vbKeyDivide
            If vsArticulo.BackColor = vbWhite Then
                'If Val(vsArticulo.Cell(flexcpText, vsArticulo.Row, 0)) > 0 Then
                    vsArticulo.Cell(flexcpText, vsArticulo.Row, 3) = Val(vsArticulo.Cell(flexcpText, vsArticulo.Row, 3)) - 1
                'End If
            End If
        
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
        
        Case 93
            If vsArticulo.BackColor = vbWhite Then PopupMenu MnuEtiquetas, , vsArticulo.ColPos(3) + vsArticulo.Left, vsArticulo.Top + vsArticulo.RowPos(vsArticulo.Row) + vsArticulo.RowHeight(vsArticulo.Row), MnuEtiTitulo
                
    End Select
    Exit Sub
ErrLA:
    clsGeneral.OcurrioError "Error inesperado. ", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub vsArticulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errMD
    If vsArticulo.Rows = 1 Or vsArticulo.BackColor <> vbWhite Or Shift <> 0 Or Button <> 2 Then Exit Sub
    PopupMenu MnuEtiquetas, , , , MnuEtiTitulo
errMD:
End Sub

Private Sub vsEnvio_RowColChange()
    If vsEnvio.Rows = 1 Then Exit Sub
    cmdImprimir.Enabled = False: cmdNuevo.Enabled = False: cmdEliminar.Enabled = False
    If CInt(vsEnvio.Cell(flexcpText, vsEnvio.Row, 1)) > 0 Then
        cmdEliminar.Enabled = True
    Else
        cmdNuevo.Enabled = True
    End If
    AccionConsultarArticulos CLng(vsEnvio.Cell(flexcpText, vsEnvio.Row, 0)), Val(vsEnvio.Cell(flexcpData, vsEnvio.Row, 3))
End Sub
'Private Sub PrueboBandejaImpresora()
'On Error GoTo ErrPBI
'    With vsListado
'        .PageBorder = pbNone
'        .Orientation = orLandscape
'        .MarginTop = 300
'        .MarginLeft = 500
'    End With
'    Exit Sub
'ErrPBI:
'    clsGeneral.OcurrioError "Ocurrió un error al setear la impresora, consulte con el administrador de impresión este problema.", Err.Description
'End Sub

Private Function BuscoCliente(IdCliente As Long) As String
On Error GoTo ErrBC
    BuscoCliente = ""
    Cons = "Select * from Cliente " _
            & " Left Outer Join CPersona ON CliCodigo = CPeCliente " _
            & " Left Outer Join CEmpresa ON CliCodigo = CEmCliente " _
       & " Where CliCodigo = " & IdCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Select Case RsAux!CliTipo
            Case TipoCliente.Cliente
                BuscoCliente = Trim(Trim(Format(RsAux!CPeNombre1, "#")) & " " & Trim(Format(RsAux!CPeNombre2, "#"))) & " " & Trim(Trim(Format(RsAux!CPeApellido1, "#")) & " " & Trim(Format(RsAux!CPeApellido2, "#")))
            Case TipoCliente.Empresa
                If Not IsNull(RsAux!CEmNombre) Then BuscoCliente = Trim(RsAux!CEmFantasia)
                If Not IsNull(RsAux!CEmFantasia) Then BuscoCliente = BuscoCliente & " (" & Trim(RsAux!CEmFantasia) & ")"
        End Select
    End If
    RsAux.Close
    Exit Function
ErrBC:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el cliente para imprimir en la ficha.", Err.Description
End Function
Private Function BuscoDocumento(idDocumento As Long) As String
 On Error GoTo ErrBD
    BuscoDocumento = ""
    sMoneda = ""
    rdoErrors.Clear
    Cons = "Select * From Documento, Sucursal, Moneda" _
        & " Where DocCodigo = " & idDocumento _
        & " And DocSucursal = SucCodigo And DocMoneda = MonCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        BuscoDocumento = Trim(RsAux!SucAbreviacion) & " " & fnc_NombreDocumento(RsAux!DocTipo) & " " & Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
'        sCodBarra = CodigoDeBarras(RsAux!DocTipo, RsAux!DocCodigo)
        On Error Resume Next
        If Not IsNull(RsAux!MonSigno) Then sMoneda = Trim(RsAux!MonSigno)
    End If
    RsAux.Close
    Exit Function
ErrBD:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el documento para imprimir en la ficha.", Err.Description
End Function

Function CodigoDeBarras(TipoDoc As Integer, CodigoDoc As Long) As String

    If Len(CodigoDoc) < 6 Then
        CodigoDeBarras = TipoDoc & "D" & Format(CodigoDoc, "000000")
    Else
        CodigoDeBarras = TipoDoc & "D" & CodigoDoc
    End If
    CodigoDeBarras = "*" & CodigoDeBarras & "*"
    
End Function

Private Sub loc_EditQEtiquetas(ByVal bABD As Boolean)
Dim sQ As String
On Error GoTo errEE
    sQ = InputBox("Ingrese la Cantidad de etiquetas a imprimir", "Cantidad " & IIf(bABD, "Permanente", "Ocacional"), "1")
    If Not IsNumeric(sQ) Then Exit Sub
    If CInt(sQ) < 1 Then MsgBox "Cantidad incorrecta.", vbExclamation, "Atención": Exit Sub
    
    vsArticulo.Cell(flexcpText, vsArticulo.Row, 3) = CInt(sQ)
    
    If bABD Then
        On Error GoTo errSave
        Cons = "Select * From ArticuloFacturacion Where AFaArticulo = " & vsArticulo.Cell(flexcpData, vsArticulo.Row, 0)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Edit
            RsAux("AFaEtiqAgencia") = CInt(sQ)
            RsAux.Update
        Else
            MsgBox "No existe un registro para el artículo seleccionado, no podrá modificar el dato.", vbExclamation, "Atención"
        End If
        RsAux.Close
    End If
Exit Sub
errEE:
    clsGeneral.OcurrioError "Error al pedir la cantidad.", Err.Description
    Exit Sub
errSave:
    clsGeneral.OcurrioError "Error al grabar la cantidad.", Err.Description
End Sub
'------------------------------------------------------------------------------------------------------------------------------------
'   Setea la impresora pasada como parámetro como: por defecto
'------------------------------------------------------------------------------------------------------------------------------------
Private Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        Debug.Print X.DeviceName
        If UCase(Trim(X.DeviceName)) = UCase(Trim(DeviceName)) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

Private Function fnc_NombreDocumento(Codigo As Integer) As String
    Dim aRet As String
    
    aRet = ""
    Select Case Codigo
        Case 7, 8: aRet = "VTE"
        Case 1: aRet = "CON"
        Case 2: aRet = "CRE"
        Case 4: aRet = "NCR"
        Case 3: aRet = "NDE"
        Case 10: aRet = "NES"
        Case 6:  aRet = "REM"
    End Select
    
    fnc_NombreDocumento = aRet
    
End Function

