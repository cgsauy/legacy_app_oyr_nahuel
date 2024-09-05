VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form EntMercaderia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrega de Mercadería a Camión"
   ClientHeight    =   5940
   ClientLeft      =   30
   ClientTop       =   615
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EntMercaderia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprimir entrega"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar impresión"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "detalle"
            Object.ToolTipText     =   "Imprimir Detalle de Entrega"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   4400
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   2940
      Visible         =   0   'False
      Width           =   6015
      _Version        =   196608
      _ExtentX        =   10610
      _ExtentY        =   3625
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   1020
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   7435
      _ConvInfo       =   1
      Appearance      =   1
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
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
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
   Begin AACombo99.AACombo cCamionero 
      Height          =   315
      Left            =   3780
      TabIndex        =   3
      Top             =   600
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
   End
   Begin VB.TextBox tUsuario 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   6
      Top             =   5340
      Width           =   615
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5685
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
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
   Begin VB.TextBox tCodigo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   5340
      Width           =   855
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercaderia.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercaderia.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercaderia.frx":0736
            Key             =   "Total"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercaderia.frx":0A50
            Key             =   "Parcial"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercaderia.frx":0D6A
            Key             =   "Pasado"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercaderia.frx":1084
            Key             =   "NoDoy"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercaderia.frx":139E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercaderia.frx":14B0
            Key             =   "No"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercaderia.frx":17CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ca&mionero:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2820
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Código de Entrega:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1452
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuDetalle 
         Caption         =   "&Detalle de Entrega"
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "EntMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strLocal As String

'Colores de stock
Private Const cTotal = &HC0C000   '&HFF00&
Private Const cParcial = &HC0FFFF
Private Const cNo = &H80C0FF   '&HC0C0FF
Private Const cPasado = &H80C0FF

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault: Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    'InicializoGrilla
    With vsConsulta
        .Redraw = False
        .Editable = False: .Rows = 1: .Cols = 1: .ExtendLastCol = True
        .FormatString = "Entregar|>Necesita|<Artículo|Stock"
        .ColWidth(2) = 3500
        .ColHidden(3) = True
        .Redraw = True
    End With
    CargoCamiones
    MnuImprimir.Enabled = False: Toolbar1.Buttons("imprimir").Enabled = False
    MnuDetalle.Enabled = False: Toolbar1.Buttons("detalle").Enabled = False
    Exit Sub
ErrLoad:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al cargar el formulario.", Err.Description
End Sub
Private Sub cCamionero_Change()
    vsConsulta.Rows = 1
    MnuImprimir.Enabled = False: Toolbar1.Buttons("imprimir").Enabled = False
    MnuDetalle.Enabled = False: Toolbar1.Buttons("detalle").Enabled = False
End Sub
Private Sub cCamionero_Click()
    vsConsulta.Rows = 1
    MnuImprimir.Enabled = False: Toolbar1.Buttons("imprimir").Enabled = False
    MnuDetalle.Enabled = False: Toolbar1.Buttons("detalle").Enabled = False
End Sub
Private Sub cCamionero_GotFocus()
    cCamionero.SelStart = 0
    cCamionero.SelLength = Len(cCamionero.Text)
    Status.SimpleText = " Seleccione un camionero, lista de impresiones pendientes [ F1 ]."
End Sub
Private Sub cCamionero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And cCamionero.ListIndex > -1 Then ListaAyuda
End Sub
Private Sub cCamionero_LostFocus()
    cCamionero.SelLength = 0
    Status.SimpleText = vbNullString
End Sub

Private Sub CargoCamiones()
On Error GoTo ErrCO
    Cons = "Select CamCodigo, CamNombre From Camion Order by CamNombre"
    CargoCombo Cons, cCamionero
    Exit Sub
ErrCO:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los Camioneros."
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.SimpleText = vbNullString
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Sub

Private Sub Label2_Click()
    Foco tCodigo
End Sub
Private Sub Label3_Click()
    Foco cCamionero
End Sub
Private Sub Label8_Click()
    Foco tUsuario
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuDetalle_Click()
    AccionImprimeDetalle
End Sub

Private Sub MnuImprimir_Click()
    AccionImprimir
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Function StockLocalArticuloyEstado(lnArticulo As Long, iEstado As Integer) As Integer
On Error GoTo errSTL
Dim Rs As rdoResultset
        Screen.MousePointer = vbHourglass
        StockLocalArticuloyEstado = 0
        Cons = "Select Sum(StLCantidad) From StockLocal " _
            & " Where StLArticulo = " & lnArticulo & " And StlTipoLocal = " & TipoLocal.Deposito _
            & " And StLLocal = " & paCodigoDeSucursal & " And StLEstado = " & iEstado
        Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not Rs.EOF Then
            If Not IsNull(Rs(0)) Then StockLocalArticuloyEstado = Rs(0)
        End If
        Rs.Close
        Screen.MousePointer = vbDefault
        Exit Function

errSTL:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error inesperado al buscar el stock del local."

End Function

Private Sub tCodigo_Change()
    AccionCancelar
End Sub

Private Sub tCodigo_GotFocus()
    tCodigo.SelStart = 0
    tCodigo.SelLength = Len(tCodigo.Text)
    Status.SimpleText = " Ingrese el código de impresión de entrega."
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tCodigo.Text) = vbNullString Then Exit Sub
        If IsNumeric(tCodigo.Text) Then
            BuscoCodigoImpresion
            If vsConsulta.Rows > 1 Then vsConsulta.SetFocus
        Else
            MsgBox "El formato ingresado no es numérico.", vbExclamation, "ATENCIÓN"
            tCodigo.SetFocus
        End If
    End If
End Sub

Private Sub tCodigo_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

    Select Case Button.Key
        Case "salir": Unload Me
        Case "imprimir": AccionImprimir
        Case "cancelar": AccionCancelar
        Case "detalle": AccionImprimeDetalle
    End Select

End Sub

Private Sub AccionImprimir()
Dim Msg As String, sPasado As Boolean

    Msg = vbNullString: sPasado = False
    For I = 1 To vsConsulta.Rows - 1
        If CInt(vsConsulta.Cell(flexcpText, I, 0)) <> 0 Then Msg = "hay"
        If CInt(vsConsulta.Cell(flexcpText, I, 0)) > CInt(vsConsulta.Cell(flexcpText, I, 3)) Then sPasado = True
    Next
    
    If sPasado Then
        If MsgBox("Hay artículos que quedaran con stock negativo.¿Desea continuar?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    End If
    If Msg = vbNullString Then MsgBox "No hay datos a imprimir.", vbExclamation, "ATENCIÓN": Exit Sub
    
    If Trim(tUsuario.Tag) <> vbNullString Then
        If Not CInt(tUsuario.Tag) > 0 Then
            MsgBox "Ingrese su digito de usuario.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
    Else
        MsgBox "Ingrese su digito de usuario.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("¿Confirma imprimir la entrega de mercadería seleccionada.", vbQuestion + vbYesNo, "IMPRIMIR") = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    FechaDelServidor
    Msg = vbNullString
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRelajo
    With vsConsulta
    
        For I = 1 To .Rows - 1
        
            If CInt(.Cell(flexcpText, I, 0)) > 0 Then
            
                Cons = "Select * From RenglonEntrega Where ReECodImpresion = " & tCodigo.Text _
                    & " And ReEArticulo = " & CLng(.Cell(flexcpData, I, 0))
                
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                If RsAux.EOF Then
                    RsAux.Close
                    cBase.RollbackTrans
                    Screen.MousePointer = 0
                    MsgBox "Alguna terminal pudo modificar el código de impresión seleccionado, verifique.", vbInformation, "ATENCIÓN"
                    Exit Sub
                Else
                    If CDate(RsAux!ReEFModificacion) <> CDate(tCodigo.Tag) Then
                        Msg = RsAux!ReEFModificacion
                        RsAux.Close: cBase.RollbackTrans
                        Screen.MousePointer = 0
                        MsgBox "Alguna terminal pudo modificar los datos del código de impresión seleccionado." & vbCrLf _
                            & "Dato actual: " & Msg & " , dato leído con: " & CDate(tCodigo.Tag), vbInformation, "ATENCIÓN"
                        Exit Sub
                    Else
                        RsAux.Edit
                        RsAux!ReECantidadEntregada = RsAux!ReECantidadEntregada + CInt(.Cell(flexcpText, I, 0))
                        RsAux!ReEUsuario = tUsuario.Tag
                        RsAux.Update
                    End If
                End If
                RsAux.Close
                
                'Primero doy el alta para el camión y luego la baja para el local.
                Cons = "Select * From StockLocal " _
                    & " Where StLArticulo = " & CLng(.Cell(flexcpData, I, 0)) & " And StlTipoLocal = " & TipoLocal.Camion _
                    & " And StLLocal = " & cCamionero.ItemData(cCamionero.ListIndex) & " And StLEstado = " & paEstadoArticuloEntrega
                
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
                If RsAux.EOF Then
                    RsAux.AddNew
                    RsAux!StLArticulo = CLng(.Cell(flexcpData, I, 0))
                    RsAux!StLTipoLocal = TipoLocal.Camion
                    RsAux!StlLocal = cCamionero.ItemData(cCamionero.ListIndex)
                    RsAux!StLEstado = paEstadoArticuloEntrega
                    RsAux!StLCantidad = CInt(.Cell(flexcpText, I, 0))
                    RsAux.Update
                Else
                    RsAux.Edit
                    RsAux!StLCantidad = RsAux!StLCantidad + CInt(.Cell(flexcpText, I, 0))
                    RsAux.Update
                End If
                RsAux.Close
                MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Camion, cCamionero.ItemData(cCamionero.ListIndex), CLng(.Cell(flexcpData, I, 0)), CInt(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, 1, TipoDocumento.Envios, CLng(tCodigo.Text)
                    
                Cons = "Select * From StockLocal " _
                    & " Where StLArticulo = " & CLng(.Cell(flexcpData, I, 0)) & " And StlTipoLocal = " & TipoLocal.Deposito _
                    & " And StLLocal = " & paCodigoDeSucursal & " And StLEstado = " & paEstadoArticuloEntrega
                
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
                If RsAux.EOF Then
                    RsAux.AddNew
                    RsAux!StLArticulo = CLng(.Cell(flexcpData, I, 0))
                    RsAux!StLTipoLocal = TipoLocal.Deposito
                    RsAux!StlLocal = paCodigoDeSucursal
                    RsAux!StLEstado = paEstadoArticuloEntrega
                    RsAux!StLCantidad = CInt(.Cell(flexcpText, I, 0)) * -1
                    RsAux.Update
                    RsAux.Close
                    'Registro suceso silencioso.
                    clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.DiferenciaDeArticulos, paCodigoDeTerminal, CLng(tUsuario.Tag), 0, CLng(.Cell(flexcpData, I, 0)), _
                          Descripcion:="Entrega de Mercadería al Camión, código: " & tCodigo.Text, Defensa:="Se ingresaron " & CInt(.Cell(flexcpText, I, 0)) & " artículos " & Trim(.Cell(flexcpText, I, 2)) & " sin haber en el local."
                Else
                    If RsAux!StLCantidad - CInt(.Cell(flexcpText, I, 0)) < 0 Then
                        'Registro suceso silencioso.
                        If RsAux!StLCantidad > 0 Then
                            clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.DiferenciaDeArticulos, paCodigoDeTerminal, CLng(tUsuario.Tag), 0, CLng(.Cell(flexcpData, I, 0)), _
                                  Descripcion:="Entrega de Mercadería al Camión, código: " & tCodigo.Text, Defensa:="Se entregaron  " & (RsAux!StLCantidad - CInt(.Cell(flexcpText, I, 0))) * -1 & " del artículo " & Trim(.Cell(flexcpText, I, 2)) & " de más y quedo negativo el stock."
                        Else
                            clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.DiferenciaDeArticulos, paCodigoDeTerminal, CLng(tUsuario.Tag), 0, CLng(.Cell(flexcpData, I, 0)), _
                                Descripcion:="Entrega de Mercadería al Camión, código: " & tCodigo.Text, Defensa:="Se entregaron  " & CInt(.Cell(flexcpText, I, 0)) & " del artículo " & Trim(.Cell(flexcpText, I, 2)) & " de más y quedo negativo el stock."
                        End If
                    End If
                    RsAux.Edit
                    RsAux!StLCantidad = RsAux!StLCantidad - CInt(.Cell(flexcpText, I, 0))
                    RsAux.Update
                    RsAux.Close
                End If
                MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), CInt(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, -1, TipoDocumento.Envios, CLng(tCodigo.Text)
            End If
        Next I
    End With
    Cons = "Update RenglonEntrega Set ReEFModificacion = '" & Format(gFechaServidor, sqlFormatoFH) & "' Where ReECodImpresion = " & tCodigo.Text
    cBase.Execute (Cons)
    cBase.CommitTrans
    
    'IMPRIMO.----------------
    ImprimoHoja "Firma:......................................................."
    
    On Error Resume Next
    AccionCancelar
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrBT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al intentar iniciar la transaccion.", Err.Description
    Exit Sub
    
ErrRelajo:
    Resume Resumo
    
Resumo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError Msg
    Exit Sub
End Sub
Private Sub BuscoCodigoImpresion()
On Error GoTo ErrBusco
Dim IdArticulo As Long

    Screen.MousePointer = vbHourglass
    vsConsulta.Rows = 1
    cCamionero.ListIndex = -1
    
    Cons = "Select RenglonEntrega.*, ArtNombre, ArtCodigo, ArtID From RenglonEntrega, Articulo" _
        & " Where ReECodImpresion = " & tCodigo.Text & " And ArtTipo <> " & paTipoArticuloServicio _
        & " And ReEArticulo = ArtID"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAux.EOF Then
        Screen.MousePointer = vbDefault
        Toolbar1.Buttons("imprimir").Enabled = False: MnuImprimir.Enabled = False
        MnuDetalle.Enabled = False: Toolbar1.Buttons("detalle").Enabled = False
        MsgBox "No se encontro una impresión con ese código verifique.", vbExclamation, "ATENCIÓN"
    Else
        tCodigo.Tag = RsAux!ReEFModificacion
        BuscoCodigoEnCombo cCamionero, RsAux!ReECamion
        With vsConsulta
            Do While Not RsAux.EOF
                If RsAux!ReECantidadtotal > RsAux!ReECantidadEntregada Then
                    .AddItem ""
                    
                    IdArticulo = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = IdArticulo
                    .Cell(flexcpText, .Rows - 1, 1) = RsAux!ReECantidadtotal - RsAux!ReECantidadEntregada
                    .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
                    .Cell(flexcpText, .Rows - 1, 3) = StockLocalArticuloyEstado(RsAux!ReEArticulo, paEstadoArticuloEntrega)
                    
                    'Existe stock en el local.----------------------------
                    If CInt(.Cell(flexcpText, .Rows - 1, 3)) > 0 Then
                        If RsAux!ReECantidadtotal - RsAux!ReECantidadEntregada > CInt(.Cell(flexcpText, .Rows - 1, 3)) Then
                            'La cantidad que tengo que mayor a lo que hay en Stock
                            .Cell(flexcpText, .Rows - 1, 0) = CInt(.Cell(flexcpText, .Rows - 1, 3))
                            '.Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("Parcial").ExtractIcon
                            .Cell(flexcpBackColor, .Rows - 1, 0) = cParcial
                        Else
                            .Cell(flexcpText, .Rows - 1, 0) = RsAux!ReECantidadtotal - RsAux!ReECantidadEntregada
                            '.Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("Total").ExtractIcon
                            .Cell(flexcpBackColor, .Rows - 1, 0) = cTotal
                        End If
                    Else
                        .Cell(flexcpText, .Rows - 1, 0) = 0
                        '.Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("No").ExtractIcon
                        .Cell(flexcpBackColor, .Rows - 1, 0) = cNo
                    End If
                End If
                RsAux.MoveNext
            Loop
            
            If .Rows > 1 Then
                Toolbar1.Buttons("imprimir").Enabled = True: MnuImprimir.Enabled = True
                MnuDetalle.Enabled = True: Toolbar1.Buttons("detalle").Enabled = True
            End If
        End With
    End If
    RsAux.Close
    Screen.MousePointer = vbDefault
    Exit Sub

ErrBusco:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al buscar la información."
End Sub

Private Sub AccionCancelar()
On Error Resume Next
    vsConsulta.Rows = 1
    cCamionero.ListIndex = -1
    Toolbar1.Buttons("imprimir").Enabled = False: MnuImprimir.Enabled = False
    MnuDetalle.Enabled = False: Toolbar1.Buttons("detalle").Enabled = False
    tCodigo.SetFocus
End Sub

Private Sub ListaAyuda()

    Cons = "Select Distinct(ReECodImpresion), Código = ReECodImpresion " _
        & " From RenglonEntrega " _
        & " Where ReECamion = " & cCamionero.ItemData(cCamionero.ListIndex) _
        & " And ReECantidadTotal > ReECantidadEntregada"
    Dim objLista As New clsListadeAyuda
    objLista.ActivoListaAyuda Cons, False, txtConexion, 2500
    Me.Refresh
    tCodigo.Text = ""
    tCodigo.Text = objLista.ValorSeleccionado
    Set objLista = Nothing
    If Val(tCodigo.Text) > 0 Then tCodigo_KeyPress (vbKeyReturn) Else tCodigo.Text = ""

End Sub

Private Sub tUsuario_GotFocus()
    tUsuario.SelStart = 0: tUsuario.SelLength = Len(tUsuario.Text)
    Status.SimpleText = " Ingrese su código de Usuario."
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    tUsuario.Tag = vbNullString
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = BuscoUsuarioDigito(CInt(tUsuario.Text), True)
            If CInt(tUsuario.Tag) > 0 Then
                AccionImprimir
            Else
                tUsuario.Tag = vbNullString
            End If
        Else
            MsgBox "El formato del código de usuario no es numérico.", vbExclamation, "ATENCIÓN"
            tUsuario.SetFocus
        End If
    End If
    
End Sub

Private Sub AccionImprimeDetalle()
    ImprimoHoja "Detalle posible de entrega."
End Sub

Private Sub vsConsulta_GotFocus()
    Status.SimpleText = " Seleccione e indique si retira ('S', 'N'), modifique la cantidad ('+', '-')."
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    
    With vsConsulta
        If .Rows > 1 Then
            Select Case KeyCode
                Case vbKeyN
                    .Cell(flexcpText, .Row, 0) = "0"
                    '.Cell(flexcpPicture, .Row, 0) = ImageList1.ListImages("NoDoy").ExtractIcon
                    .Cell(flexcpBackColor, .Row, 0) = cNo
                    
                Case vbKeyS
                    .Cell(flexcpText, .Row, 0) = .Cell(flexcpText, .Row, 1)
                    If CLng(.Cell(flexcpText, .Row, 0)) > CInt(.Cell(flexcpText, .Row, 3)) Then
                            '.Cell(flexcpPicture, .Row, 0) = ImageList1.ListImages("Pasado").ExtractIcon
                            .Cell(flexcpBackColor, .Rows - 1, 0) = cPasado
                        ElseIf CLng(.Cell(flexcpText, .Row, 0)) < CInt(.Cell(flexcpText, .Row, 1)) Then
                            '.Cell(flexcpPicture, .Row, 0) = ImageList1.ListImages("Parcial").ExtractIcon
                            .Cell(flexcpBackColor, .Row, 0) = cParcial
                        Else
                            '.Cell(flexcpPicture, .Row, 0) = ImageList1.ListImages("Total").ExtractIcon
                            .Cell(flexcpBackColor, .Row, 0) = cTotal
                        End If
                    
                Case vbKeyAdd
                    If CLng(.Cell(flexcpText, .Row, 0)) < CInt(.Cell(flexcpText, .Row, 1)) Then
                        .Cell(flexcpText, .Row, 0) = CInt(.Cell(flexcpText, .Row, 0)) + 1
                        If CLng(.Cell(flexcpText, .Row, 0)) > CInt(.Cell(flexcpText, .Row, 3)) Then
                            '.Cell(flexcpPicture, .Row, 0) = ImageList1.ListImages("Pasado").ExtractIcon
                            .Cell(flexcpBackColor, .Row, 0) = cPasado
                        ElseIf CLng(.Cell(flexcpText, .Row, 0)) < CInt(.Cell(flexcpText, .Row, 1)) Then
                            '.Cell(flexcpPicture, .Row, 0) = ImageList1.ListImages("Parcial").ExtractIcon
                            .Cell(flexcpBackColor, .Row, 0) = cParcial
                        Else
                            '.Cell(flexcpPicture, .Row, 0) = ImageList1.ListImages("Total").ExtractIcon
                            .Cell(flexcpBackColor, .Row, 0) = cTotal
                        End If
                    End If
                
                Case vbKeySubtract
                    If CInt(.Cell(flexcpText, .Row, 0)) > 0 Then
                        .Cell(flexcpText, .Row, 0) = CInt(.Cell(flexcpText, .Row, 0)) - 1
                        If CLng(.Cell(flexcpText, .Row, 0)) = 0 Then
                            '.Cell(flexcpPicture, .Row, 0) = ImageList1.ListImages("NoDoy").ExtractIcon
                            .Cell(flexcpBackColor, .Row, 0) = cNo
                        ElseIf CLng(.Cell(flexcpText, .Row, 0)) > CInt(.Cell(flexcpText, .Row, 3)) Then
                            '.Cell(flexcpPicture, .Row, 0) = ImageList1.ListImages("Pasado").ExtractIcon
                            .Cell(flexcpBackColor, .Row, 0) = cPasado
                        ElseIf CLng(.Cell(flexcpText, .Row, 0)) < CInt(.Cell(flexcpText, .Row, 1)) Then
                            '.Cell(flexcpPicture, .Row, 0) = ImageList1.ListImages("Parcial").ExtractIcon
                            .Cell(flexcpBackColor, .Row, 0) = cParcial
                        End If
                    End If
                
                Case vbKeyReturn: tUsuario.SetFocus
            End Select
        End If
    End With

End Sub

Private Sub vsConsulta_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub ImprimoHoja(Pie As String)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub

    With vsListado
        .Device = paICartaN
        .Orientation = orPortrait
        .PaperSize = 1  'Hoja carta
        .PaperBin = paICartaB         'Bandeja por defecto.
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    vsListado.FileName = "Entrega de Mercaderia"
    
    EncabezadoListado vsListado, "Entrega de Mercadería al Camión", False
    vsListado.FontBold = True
    vsListado.Paragraph = "Código de Impresión = " & tCodigo.Text
    vsListado.Paragraph = "Camión: " & Trim(cCamionero.Text)
    vsListado.Paragraph = "Sucursal = " & aSucursal & Space(10) & "Terminal = " & miConexion.NombreTerminal & Space(20) & Pie
    vsListado.Paragraph = ""
    vsListado.FontBold = False
    
    vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
    
    With vsListado
        .EndDoc
        .Device = paICartaN
        .PaperBin = paICartaB
        .PrintDoc
    End With

    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description

End Sub

