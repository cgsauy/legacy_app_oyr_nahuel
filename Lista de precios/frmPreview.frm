VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmPreview 
   Caption         =   "Preview"
   ClientHeight    =   3450
   ClientLeft      =   1230
   ClientTop       =   2010
   ClientWidth     =   7380
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   7380
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picCopia 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   5400
      ScaleHeight     =   315
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
      Begin VB.VScrollBar vsCopias 
         Height          =   285
         Left            =   960
         Max             =   -1
         Min             =   -1111
         TabIndex        =   6
         Top             =   0
         Value           =   -1
         Width           =   255
      End
      Begin VB.TextBox tCopias 
         Appearance      =   0  'Flat
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
         Left            =   600
         MaxLength       =   5
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Copias:"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   40
         Width           =   615
      End
   End
   Begin VB.HScrollBar fsbZoom 
      Height          =   255
      Left            =   6480
      TabIndex        =   3
      Top             =   1620
      Width           =   1095
   End
   Begin VB.TextBox tPage 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6480
      MaxLength       =   5
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin MSComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Cerrar [Ctrl+X]"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "play"
            Object.ToolTipText     =   "Refrescar consulta. [Ctrl+E]"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stop"
            Object.ToolTipText     =   "Cancelar carga. [Ctrl+C]"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "separator1"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir. [Ctrl+P]"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "printconfig"
            Object.ToolTipText     =   "Configurar página."
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "printcopies"
            Style           =   4
            Object.Width           =   1330
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "firstpage"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previouspage"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pagenumber"
            Style           =   4
            Object.Width           =   815
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nextpage"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lastpage"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "separator4"
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "zoom"
            Style           =   4
            Object.Width           =   1500
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   4920
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":0442
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":075C
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":08B6
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":0D08
            Key             =   "print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":0E1A
            Key             =   "printcfg"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":0F2C
            Key             =   "first"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":137E
            Key             =   "previous"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":17D0
            Key             =   "next"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreview.frx":1C22
            Key             =   "last"
         EndProperty
      EndProperty
   End
   Begin vsViewLib.vsPrinter vspReporte 
      Height          =   2055
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   4215
      _Version        =   196608
      _ExtentX        =   7435
      _ExtentY        =   3625
      _StockProps     =   229
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
      Zoom            =   80
      ZoomMax         =   160
      ZoomStep        =   10
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Properties.
Private m_Lista As Long
Private m_Header As String
Private m_Vigencia As String
Private m_MonedaPesos As Long
Private m_Caption As String
'------------------------------------------------

Private Type tRepProperties
    Device As String
    MarginL As Long
    MarginR As Long
    MarginB As Long
    MarginT As Long
    Orientation As Integer
    PaperSize As Integer
End Type

Private bHeader As Boolean

Private dTitRige As Date
Private sTitTableTC As String
Private sTitTableTCD As String      'Diferidas.
Private lHeightCol As Long

'Guardo dimensión de tabla.
Private sDimTableTC As String, sDimTableTCD As String

Private Const sDimTableFecha = "|>770"
Private Const sDimTableArt = "<300|>780|<3150"
Private Const lDimTable = 5000
'------------------------------------------------

Private arrCuota() As Long
Private arrDiferidos() As Long
Private arrPrecios() As String, arrDif() As String

Private bLoad As Boolean
Private bCancelQuery As Boolean

Public Property Let prmVigencia(ByVal dRige As String)
    m_Vigencia = dRige
End Property

Public Property Let prmMonedaPesos(ByVal lMon As Long)
    m_MonedaPesos = lMon
End Property

Public Property Let prmIDLista(ByVal lid As Long)
    m_Lista = lid
End Property

Public Property Let prmCaption(ByVal sCaption As String)
    m_Caption = sCaption
End Property

Public Property Let prmHeaderReport(ByVal sHeader As String)
    m_Header = sHeader
End Property

Private Sub Form_Load()
On Error GoTo errLoad
    
    bLoad = True
    bHeader = False
    If m_Vigencia = "0:00:00" Then prmVigencia = Format(Now, "mm/dd/yyyy hh:nn:ss")
    With tooMenu
        .ImageList = imgIcon
        .Buttons("salir").Image = "salir"
        .Buttons("play").Image = "refresh"
        .Buttons("stop").Image = "stop"
        .Buttons("print").Image = "print"
        .Buttons("printconfig").Image = "printcfg"
        .Buttons("firstpage").Image = "first"
        .Buttons("previouspage").Image = "previous"
        .Buttons("nextpage").Image = "next"
        .Buttons("lastpage").Image = "last"
    End With
    
    With vspReporte
        .Zoom = 100
        fsbZoom.LargeChange = .ZoomStep
        fsbZoom.SmallChange = .ZoomStep / 2
        fsbZoom.Min = .ZoomMin
        fsbZoom.Max = .ZoomMax
        fsbZoom.Value = .Zoom

        .PaperSize = 1
        .Orientation = orPortrait
        'OJO EN WIN2000
        'si ponemos printer se cuelga el equipo al hacer render control
        .PreviewMode = pmScreen
        .MarginLeft = 576 '649
        .MarginRight = 576 '609
        .MarginTop = 1200
        .PhysicalPage = True
        .PageBorder = pbBottom
        
    End With
    
    StartReport
    
    If m_Caption <> "" Then Me.Caption = m_Caption
    
    With vspReporte
        .Zoom = 100
        fsbZoom.LargeChange = .ZoomStep
        fsbZoom.SmallChange = .ZoomStep / 2
        fsbZoom.Min = .ZoomMin
        fsbZoom.Max = .ZoomMax
        fsbZoom.Value = .Zoom
    End With
    Exit Sub
    
errLoad:
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    With vspReporte
        .Top = tooMenu.Top + tooMenu.Height
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Height = ScaleHeight - .Top
    End With
    With fsbZoom
        .Move tooMenu.Buttons("zoom").Left, tooMenu.Buttons("zoom").Top + ((tooMenu.Height - .Height) / 1.5), tooMenu.Buttons("zoom").Width
    End With
    With picCopia
        .Move tooMenu.Buttons("printcopies").Left + 100, tooMenu.Buttons("printcopies").Top + 50  '((tooMenu.Height - .Height) / 1.5)   ', tooMenu.Buttons("printcopies").Width
    End With
    
    With tPage
        .Move tooMenu.Buttons("pagenumber").Left + 150, tooMenu.Buttons("pagenumber").Top + ((tooMenu.Height - .Height) / 1.5)  ', tooMenu.Buttons("pagenumber").Width
    End With
End Sub

Private Sub fsbZoom_Change()
On Error Resume Next
    If vspReporte Is Nothing Then Exit Sub
    vspReporte.Zoom = fsbZoom.Value
End Sub

Private Sub tCopias_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCopias.Text) Then
            vsCopias.Value = Val(tCopias.Text) * -1
        Else
            MsgBox "Formato incorrecto.", vbExclamation, "ATENCIÓN"
            tCopias.Text = vsCopias.Value * -1
        End If
    End If
End Sub

Private Sub tCopias_LostFocus()
On Error Resume Next
    If IsNumeric(tCopias.Text) Then
        vsCopias.Value = Val(tCopias.Text) * -1
    Else
        tCopias.Text = vsCopias.Value * -1
    End If
End Sub

Private Sub tooMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "salir": Unload Me
        Case "play": StartReport
        Case "stop": ActionStop
        Case "print": ActionPrint
        Case "printconfig": ActionConfigPage
        
        'Botones de reporte.
        Case "firstpage"
            vspReporte.PreviewPage = 1
            SetButtonReport
        Case "previouspage"
            vspReporte.PreviewPage = vspReporte.PreviewPage - 1
            SetButtonReport
        Case "nextpage"
            vspReporte.PreviewPage = vspReporte.PreviewPage + 1
            SetButtonReport
        Case "lastpage"
            vspReporte.PreviewPage = vspReporte.PageCount
            SetButtonReport
    End Select
End Sub

Private Sub ActionStop()
    bCancelQuery = True
End Sub

Private Sub ActionConfigPage()
On Error GoTo errCancel
Dim vProperties As tRepProperties
    
    With vProperties
        .Device = vspReporte.Device
        .MarginB = vspReporte.MarginBottom
        .MarginL = vspReporte.MarginLeft
        .MarginR = vspReporte.MarginRight
        .MarginT = vspReporte.MarginTop
        .Orientation = vspReporte.Orientation
        .PaperSize = vspReporte.PaperSize
    End With
     
    If vspReporte.PrintDialog(pdPageSetup) Then
        
        If vProperties.Device <> vspReporte.Device Or _
            vProperties.MarginB <> vspReporte.MarginBottom Or _
            vProperties.MarginL <> vspReporte.MarginLeft Or _
            vProperties.MarginR <> vspReporte.MarginRight Or _
            vProperties.MarginT <> vspReporte.MarginTop Or _
            vProperties.Orientation <> vspReporte.Orientation Or _
            vProperties.PaperSize <> vspReporte.PaperSize Then
        
            StartReport
            
        End If
    End If
    Exit Sub
errCancel:
    Screen.MousePointer = 0
    Exit Sub
End Sub
Private Sub ActionPrint()
Dim vProperties As tRepProperties
Dim lCopies As Long
    
    With vProperties
        .Device = vspReporte.Device
        .MarginB = vspReporte.MarginBottom
        .MarginL = vspReporte.MarginLeft
        .MarginR = vspReporte.MarginRight
        .MarginT = vspReporte.MarginTop
        .Orientation = vspReporte.Orientation
        .PaperSize = vspReporte.PaperSize
    End With
    
    If Val(tCopias.Text) < 1 Then tCopias.Text = 1
    lCopies = Val(tCopias.Text)
    
    If vspReporte.PrintDialog(pdPrinterSetup) Then
        
        If vProperties.Device <> vspReporte.Device Or _
            vProperties.MarginB <> vspReporte.MarginBottom Or _
            vProperties.MarginL <> vspReporte.MarginLeft Or _
            vProperties.MarginR <> vspReporte.MarginRight Or _
            vProperties.MarginT <> vspReporte.MarginTop Or _
            vProperties.Orientation <> vspReporte.Orientation Or _
            vProperties.PaperSize <> vspReporte.PaperSize Then
            
            StartReport
            
        End If
        vspReporte.AbortWindow = False
        vspReporte.FileName = Me.Caption
        
        tCopias.Text = lCopies
        'Como no se si la impresora acepta cant. de copias hago un loop.
        For lCopies = 1 To Val(tCopias.Text)
            vspReporte.PrintDoc False
        Next
        
    End If
    
End Sub

Private Sub ButtonMenu(ByVal bPlay As Boolean, ByVal bClean As Boolean, ByVal bCancel As Boolean)
    
    With tooMenu
        .Buttons("play").Enabled = bPlay
        .Buttons("stop").Enabled = bCancel
        .Buttons("print").Enabled = bClean
    End With
    
End Sub

Private Sub tPage_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tPage.Text) Then
            If CInt(tPage.Text) > 0 And CInt(tPage.Text) <= vspReporte.PageCount Then
                vspReporte.PreviewPage = CInt(tPage.Text)
            Else
                If CInt(tPage.Text) > vspReporte.PageCount Then
                    vspReporte.PreviewPage = vspReporte.PageCount
                Else
                    vspReporte.PreviewPage = 1
                End If
            End If
            vspReporte.SetFocus
        End If
        SetButtonReport
    End If
End Sub

Private Sub vsCopias_Change()
    tCopias.Text = vsCopias.Value * -1
End Sub

Private Sub vspReporte_EndDoc()
    SetHeader
End Sub

Private Sub SetHeader()
Dim lCantPage As Long
Dim lWidth As Long
Dim sAncho As String
Dim lHeader As Long

    With vspReporte
        bHeader = True
        For lCantPage = 1 To .PageCount
            .StartOverlay lCantPage
            
            lWidth = (.PageWidth - .MarginLeft - .MarginRight) / 3
            sAncho = Format(lWidth) * 2
            'sAncho = 10 & "|^" & ((lWidth * 2) - 750)
            sAncho = 10 & "|^" & sAncho - 10
            lHeader = .MarginTop / 3
            
            .CurrentY = lHeader * 1.8 ' + (lHeader / 2)
            .TextAlign = taLeftBottom
            .Font = "Tahoma": .Font.Size = 14: .Font.Bold = True: .Font.Italic = True
            .TableBorder = tbNone
            .Table = sAncho + ";" + "|" + m_Header
        
            .CurrentY = lHeader * 1.8 ' + (lHeader / 2)
            .TextAlign = taRightBaseline ' taRightBottom
            .Font = "tahoma"
            .Font.Size = 9: .Font.Bold = True: .Font.Italic = False: .FontUnderline = True
            .Table = ">" + CStr(lWidth - 50) + ";" + "Rige desde el: " & Format(dTitRige, "dd/mmm/yy")
            
            
            .TextAlign = taLeftBottom
            ' print second header line
            .Font = "Tahoma": .Font.Size = 8: .Font.Bold = True: .Font.Italic = False: .FontUnderline = False
            .CurrentY = .MarginTop - .TextHeight("H")
            .TableBorder = tbBottom
            .Table = sDimTableArt + sDimTableTC + sDimTableTCD + sDimTableFecha + ";" + "|Código|  Artículo" & sTitTableTC & sTitTableTCD & "|Fecha"
            
            .TableBorder = tbNone
            .TextAlign = taLeftTop
            .EndOverlay
        Next
        .TextAlign = 0 'taLeft
    End With
    bHeader = False
    Exit Sub

End Sub

Private Sub vspReporte_EndPage()
'    vspReporte.TableBorder = tbNone
End Sub

Private Sub vspReporte_MousePage(NewPage As Integer)
    If vspReporte Is Nothing Then Exit Sub
    SetButtonReport
End Sub

Private Sub vspReporte_MouseZoom(NewZoom As Integer)
    If vspReporte Is Nothing Then Exit Sub
    fsbZoom.Value = vspReporte.Zoom
End Sub

Private Sub vspReporte_NewPage()
On Error Resume Next
    'Cdo. carga el report voy enumerando en el textbox.
    tPage.Text = Val(tPage.Text) + 1
    tPage.Refresh
End Sub

Private Sub SetButtonReport()
On Error Resume Next
Dim iCantPag As Integer, iPrePag As Integer
    
    With vspReporte
        iCantPag = .PageCount
        iPrePag = .PreviewPage
    End With
    With tooMenu
        .Buttons("firstpage").Enabled = (iPrePag > 1)
        .Buttons("previouspage").Enabled = (iPrePag > 1)
        .Buttons("nextpage").Enabled = (iPrePag < iCantPag)
        .Buttons("lastpage").Enabled = (iPrePag < iCantPag)
    End With
    
    picCopia.Enabled = (iCantPag > 0)
    If iCantPag = 0 Then
        tCopias.Text = ""
    Else
        tCopias.Text = 1
    End If
    tPage.Text = iPrePag
    If iCantPag > 1 Then
        tPage.Enabled = True
    Else
        tPage.Enabled = False
    End If

End Sub

Private Sub StartReport()
On Error GoTo errSR
    bCancelQuery = False
    Screen.MousePointer = 11
    ButtonMenu False, False, True
    vspReporte.StartDoc
    SetButtonReport
    DoEvents
    If vspReporte.Error <> 0 Then
        MsgBox "Ocurrio un error al iniciar el reporte.", vbCritical, "ATENCIÓN"
        vspReporte.EndDoc
        GoTo evFin
    End If
    If SetVarGlobalReport Then
        If LoadQuery Then
            Screen.MousePointer = 0 'Saco el puntero para que tome el evento cancelar
            DoReport
            vspReporte.EndDoc
        End If
    End If
    
evFin:
    cBase.QueryTimeout = 15
    Screen.MousePointer = 0
    ButtonMenu True, True, False
    SetButtonReport
    Exit Sub
    
errSR:
    cBase.QueryTimeout = 15
    MsgBox Err.Description
End Sub

Private Function SetVarGlobalReport() As Boolean
On Error GoTo errSV
Dim lWidthR As Long, lCont As Long
Dim lWidthCtdo As Long
    SetVarGlobalReport = False
    dTitRige = GetDateRige      'Fecha rige
    If LoadPlan Then
        'Dada la cantidad de cuotas que cargue doy un largo a la tabla.
        lWidthR = vspReporte.PageWidth - vspReporte.MarginLeft - vspReporte.MarginRight - lDimTable
        
        If UBound(arrCuota) >= 0 Then lCont = UBound(arrCuota) + 1
        If UBound(arrDiferidos) >= 0 Then lCont = lCont + UBound(arrDiferidos) + 1
        If lCont = 0 Then SetVarGlobalReport = True: Exit Function
        
        lWidthCtdo = 920
        sDimTableTC = "|>" & lWidthCtdo
        
        lWidthR = lWidthR - lWidthCtdo
        lCont = lCont - 1
        If lCont = 0 Then SetVarGlobalReport = True: Exit Function
        lWidthR = lWidthR / lCont
        
        'Siempre al ctdo le doy + que a los otros.
        If (lWidthR * lCont) + lDimTable > vspReporte.PageWidth - vspReporte.MarginLeft - vspReporte.MarginRight Then
            sDimTableTC = sDimTableTC & "|>" & lWidthR - (((lWidthR * lCont) + lDimTable) - (vspReporte.PageWidth - vspReporte.MarginLeft - vspReporte.MarginRight))
        Else
            sDimTableTC = sDimTableTC & "|>" & lWidthR
        End If
        For lCont = 2 To UBound(arrCuota)
            sDimTableTC = sDimTableTC & "|>" & lWidthR
        Next lCont
        For lCont = 0 To UBound(arrDiferidos)
            sDimTableTCD = sDimTableTCD & "|>" & lWidthR
        Next
        SetVarGlobalReport = True
    End If
    Exit Function
errSV:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los parámetros globales.", Err.Description, "ATENCIÓN"
End Function

Private Function LoadQuery() As Boolean
On Error GoTo errLQ
Dim sCharArt As String
Dim lEspecie As Long
    
    LoadQuery = False
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "Select Articulo.*, Especie.*, Precios.HPrVigencia, Precios.HPrPrecio, TCuCodigo, TCuCantidad, TCuVencimientoC, PlaNombre " & _
               " From " & _
                        "HistoriaPrecio Precios, " & _
                        "Articulo, Tipo, ArticuloFacturacion, Especie, " & _
                        "TipoCuota, TipoPlan" & _
                " WHERE (HPrVigencia IN" & _
                            " (Select MAX(H.HPrVigencia)" & _
                            " FROM HistoriaPrecio H" & _
                            " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                            " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                            " And H.HPrMoneda = Precios.HPrMoneda " & _
                            " And H.HPrVigencia <= '" & m_Vigencia & "'" & _
                            " )) " & _
                " And Precios.HPrArticulo = ArtID" & _
                " And ArtId = AFaArticulo And ArtTipo = TipCodigo " & _
                " And TipEspecie = EspCodigo And ArtEnUso = 1 " & _
                " And Precios.HPrMoneda = " & m_MonedaPesos & _
                " And Precios.HPrHabilitado = 1" & _
                " And ArticuloFacturacion.AFaLista = " & m_Lista & _
                " And Precios.HPrTipoCuota = TipoCuota.TCuCodigo" & _
                " And Precios.HPrPlan = TipoPlan.PlaCodigo"
                
    Cons = Cons & _
                " And TipoCuota.TCuVencimientoE is null " & _
                " And TCuEspecial = 0 " & _
                " and TCuDeshabilitado is null"
    
    'Hago Union con combos
    Cons = Cons & _
            " Union All " & _
                " Select Articulo.*, Especie.*, '' as HPrVigencia, 0 as HPrPrecio, 0 as TCuCodigo, 0 as TCuCantidad, 0 as TCuVencimientoC, '' as PlaNombre " & _
                " From Articulo, Tipo, ArticuloFacturacion, Especie, Presupuesto " & _
                " Where ArtEsCombo = 1 And PreHabilitado = 1 And PreEsPresupuesto = 0 And PreArtCombo = ArtID " & _
                " And ArticuloFacturacion.AFaLista = " & m_Lista & _
                " And ArtId = AFaArticulo And ArtTipo = TipCodigo " & _
                " And TipEspecie = EspCodigo And ArtEnUso = 1 "
                
    Cons = Cons & _
                " Order By Especie.EspNombre Asc, Articulo.ArtNombre Asc"

    cBase.QueryTimeout = 50
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    LoadQuery = True
    Exit Function
    
errLQ:
    cBase.QueryTimeout = 15
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la consulta.", Err.Description, "Load Query"
End Function

Private Sub DoReport()
On Error GoTo errDR
Dim sCharArt As String
Dim lEspecie As Long, lArticulo As Long

Dim sPrecio As String, sDiferido As String
Dim lY As Long


Dim mEspecie As Long, mArticulo As Long
Dim mNameArticulo As String, mPlan As String
Dim mLetra As String

    ReDim arrPrecios(UBound(arrCuota))
    ReDim arrDif(UBound(arrDiferidos))
    
    Do While Not RsAux.EOF
        
        
        mEspecie = RsAux!EspCodigo
        mArticulo = RsAux!ArtCodigo
        
        If lArticulo = 0 Then
            lArticulo = mArticulo
            mNameArticulo = "|" & Format(RsAux!ArtCodigo, "#,000,000") & "|" & Trim(RsAux!ArtNombre)
            mLetra = Mid(RsAux!ArtNombre, 1, 1)
            mPlan = "|" & Format(RsAux!HPrVigencia, "dd/mm") & " " & Trim(RsAux!PlaNombre)
            NuevaEspecie
            lEspecie = mEspecie
        End If
        
        If lArticulo <> mArticulo Then
            
            lArticulo = mArticulo
            With vspReporte
                .Font = "Tahoma"
                .FontSize = 8
                .FontBold = False
                .FontItalic = False
                .FontUnderline = False
                
                lHeightCol = .TextHeight("HOLA")
                sPrecio = Join(arrPrecios, "|")
                sDiferido = Join(arrDif, "|")
                .TableBorder = tbBottom
                .SpaceAfter = 50
                
                If sCharArt <> mLetra Then
                    sCharArt = mLetra
                    .AddTable sDimTableArt + sDimTableTC + sDimTableTCD + sDimTableFecha, "", sCharArt & mNameArticulo & "|$ " & sPrecio & "|" & Join(arrDif, "|") & mPlan
                Else
                    .AddTable sDimTableArt + sDimTableTC + sDimTableTCD + sDimTableFecha, "", mNameArticulo & "|$ " & sPrecio & "|" & Join(arrDif, "|") & mPlan
                End If
            End With
            
            
            mNameArticulo = "|" & Format(RsAux!ArtCodigo, "#,000,000") & "|" & Trim(RsAux!ArtNombre)
            mLetra = Mid(RsAux!ArtNombre, 1, 1)
            mPlan = "|" & Format(RsAux!HPrVigencia, "dd/mm") & " " & Trim(RsAux!PlaNombre)
            
            ReDim arrPrecios(UBound(arrCuota))
            ReDim arrDif(UBound(arrDiferidos))
            
            If lEspecie <> mEspecie Then
                lEspecie = mEspecie
                NuevaEspecie
            End If
            
        End If
        Dim iIndex As Integer
        If RsAux!TCuVencimientoC = 0 Then
            If RsAux!ArtEsCombo Then
                'Cargo todos los precios para el combo
                s_LoadPrecioCombo
            Else
            
                iIndex = GetColPlan(arrCuota, RsAux!TCuCodigo)
                If iIndex >= 0 Then arrPrecios(iIndex) = RsAux!HPrPrecio / RsAux!TCuCantidad
            End If
        Else
            iIndex = GetColPlan(arrDiferidos, RsAux!TCuCodigo)
            If iIndex >= 0 Then arrDif(iIndex) = RsAux!HPrPrecio / RsAux!TCuCantidad
        End If
        
        RsAux.MoveNext
        
        If RsAux.EOF Then
            lArticulo = mArticulo
            With vspReporte
                .Font = "Tahoma"
                .FontSize = 8
                .FontBold = False
                .FontItalic = False
                .FontUnderline = False
                lHeightCol = .TextHeight("HOLA")
                sPrecio = Join(arrPrecios, "|")
                sDiferido = Join(arrDif, "|")
                .TableBorder = tbBottom
                If sCharArt <> mLetra Then
                    sCharArt = mLetra
                    .AddTable sDimTableArt + sDimTableTC + sDimTableTCD + sDimTableFecha, "", sCharArt & mNameArticulo & "|$ " & sPrecio & "|" & Join(arrDif, "|") & mPlan
                Else
                    .AddTable sDimTableArt + sDimTableTC + sDimTableTCD + sDimTableFecha, "", mNameArticulo & "|$ " & sPrecio & "|" & Join(arrDif, "|") & mPlan
                End If
            End With
        End If
        DoEvents
        If bCancelQuery Then Exit Do
        
    Loop
    RsAux.Close
    
    Erase arrPrecios
    Erase arrDif
    Exit Sub
errDR:
    clsGeneral.OcurrioError "Ocurrió un error al cargar el reporte.", Err.Description, "Error (doreport)"
End Sub
Private Function GetColPlan(ByVal arrCol As Variant, ByVal lPlan As Long) As Integer
Dim iIndex As Integer
    GetColPlan = -1
    For iIndex = 0 To UBound(arrCol)
        If arrCol(iIndex) = lPlan Then GetColPlan = iIndex: Exit For
    Next
End Function


Private Sub NuevaEspecie()
Dim p As Long
    
    With vspReporte
        .TableBorder = tbNone
        .FontBold = True
        .Font = "Wingdings"
        .FontSize = 9
        .CurrentY = .CurrentY + 100
        p = .CurrentY
        .Text = "|"
        .FontUnderline = True
        .FontSize = 8
        .Font = "tahoma"
        .CurrentY = p
        .TextAlign = taRightTop
        .AddTable CStr(.PageWidth - .MarginLeft - .MarginRight - 400), "", Trim(RsAux!EspNombre)
        .FontBold = False
        .FontUnderline = False
        .TextAlign = 0
    End With
    

End Sub

Private Function LoadPlan() As Boolean
On Error GoTo errLP
Dim iIndex As Integer, iIndexD As Integer
    
    LoadPlan = False
    ReDim arrCuota(0)
    ReDim arrDiferidos(0)
    
    iIndex = 0: iIndexD = 0
    sTitTableTC = "": sTitTableTCD = ""
    sDimTableTC = "": sDimTableTCD = ""
    
    Cons = "Select * From TipoCuota" _
         & " Where TCuVencimientoE Is Null And TCuEspecial = 0 " _
         & " And TCuDeshabilitado Is Null Order By TCuOrden"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If ExistePlanEnPrecio(RsAux!TCuCodigo) Then
            If RsAux!TCuVencimientoC = 0 Then
                ReDim Preserve arrCuota(iIndex)
                arrCuota(iIndex) = RsAux!TCuCodigo
                iIndex = iIndex + 1
                sTitTableTC = sTitTableTC & "|" & Trim(RsAux!TCuAbreviacion)
            Else
                sTitTableTCD = sTitTableTCD & "|" & Trim(RsAux!TCuAbreviacion)
                ReDim Preserve arrDiferidos(iIndexD)
                arrDiferidos(iIndexD) = RsAux!TCuCodigo
                iIndexD = iIndexD + 1
            End If
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    LoadPlan = True
    Exit Function
errLP:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los tipos de cuotas.", Err.Description, "ATENCIÓN"
End Function

Private Function GetDateRige() As Date

    GetDateRige = Now
    Cons = "Select Max(Precios.HPrVigencia) from HistoriaPrecio Precios, Articulo, ArticuloFacturacion " & _
               " Where (Precios.HPrVigencia IN " & _
                    " (Select MAX(H.HPrVigencia)" & _
                            " FROM HistoriaPrecio H" & _
                            " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                            " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                            " And H.HPrMoneda = Precios.HPrMoneda " & _
                            " And H.HPrVigencia <= '" & m_Vigencia & "'" & _
                            " )) " & _
                " And Precios.HPrArticulo = ArtID" & _
                " And ArtId = AFaArticulo " & _
                " And ArtEnUso = 1 " & _
                " And Precios.HPrMoneda = " & m_MonedaPesos & _
                " And Precios.HPrHabilitado = 1" & _
                " And AFaLista = " & m_Lista
                
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(RsAux(0)) Then GetDateRige = RsAux(0)
    RsAux.Close

End Function

Private Sub vspReporte_NewTableCell(Row As Integer, Column As Integer, Cell As String)
On Error Resume Next
Dim lColor
    
    If bHeader Then Exit Sub
    vspReporte.FontItalic = False
    If Column = 4 Or Column = 1 Then vspReporte.FontBold = True Else vspReporte.FontBold = False
    
    If Column = UBound(arrCuota) + UBound(arrDiferidos) + 6 Or Column = UBound(arrCuota) + 5 Then
        vspReporte.DrawLine vspReporte.MarginLeft, vspReporte.CurrentY, vspReporte.MarginLeft, vspReporte.CurrentY + lHeightCol + vspReporte.SpaceAfter
    End If
    
    If Column > 4 And Column < 10 Then
        If IsNumeric(Cell) Then
            If CCur(Cell) < paCuotaMin Then
                vspReporte.FontItalic = True
            End If
        End If
    End If
    
End Sub

Private Function ExistePlanEnPrecio(ByVal idPlan As Long) As Boolean
Dim rsTC As rdoResultset
    ExistePlanEnPrecio = False
    Cons = "Select Top 1 * " & _
               " From " & _
                        "HistoriaPrecio Precios, " & _
                        "Articulo, ArticuloFacturacion " & _
                " WHERE (HPrVigencia IN" & _
                            " (Select MAX(H.HPrVigencia)" & _
                            " FROM HistoriaPrecio H" & _
                            " WHERE H.HPrArticulo = Precios.HPrArticulo " & _
                            " AND H.HPrTipoCuota = Precios.HPrTipoCuota " & _
                            " And H.HPrMoneda = Precios.HPrMoneda " & _
                            " And H.HPrVigencia <= '" & m_Vigencia & "'" & _
                            " )) " & _
                " And Precios.HPrArticulo = ArtID" & _
                " And ArtId = AFaArticulo " & _
                " And ArtEnUso = 1 " & _
                " And Precios.HPrMoneda = " & m_MonedaPesos & _
                " And Precios.HPrHabilitado = 1" & _
                " And ArticuloFacturacion.AFaLista = " & m_Lista & _
                " And Precios.HPrTipoCuota = " & idPlan
    Set rsTC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsTC.EOF Then
        ExistePlanEnPrecio = True
    End If
    rsTC.Close
End Function

Private Sub s_LoadPrecioCombo()
Dim rsPC As rdoResultset
Dim iIndex As Integer
        
        Cons = "Select PViTipoCuota, MAx(PlaNombre) as PlaNombre, TCuCodigo, TCuCantidad, TCuAbreviacion, TCuVencimientoC, sum((PViPrecio * PArCantidad)) as Precio, Count(*) as Cant " & _
                " From Presupuesto, PresupuestoArticulo, PrecioVigente, TipoCuota, TipoPlan " & _
                " Where PreArtCombo = " & RsAux!ArtID & " And PreID = PArPresupuesto And PViArticulo = PArArticulo " & _
                " And TCuVencimientoE Is Null And TCuEspecial = 0 And TCuDeshabilitado Is Null " & _
                " And PViHabilitado <> 0  And PViMoneda = 1 And PViTipoCuota = TCuCodigo And PViPlan = PlaCodigo " & _
                " Group By PViTipoCuota, TCuCodigo, TCuCAntidad, TcuAbreviacion, TCuVencimientoC " & _
                " Order by PViTipoCuota"

        Set rsPC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsPC.EOF
            If rsPC!TCuVencimientoC = 0 Then
                iIndex = GetColPlan(arrCuota, rsPC!TCuCodigo)
                If iIndex >= 0 Then arrPrecios(iIndex) = Format(rsPC("Precio") / rsPC("TCuCantidad"), "###0")
            Else
                iIndex = GetColPlan(arrDiferidos, rsPC!TCuCodigo)
                If iIndex >= 0 Then arrDif(iIndex) = Format(rsPC("Precio") / rsPC("TCuCantidad"), "###0")
            End If
            rsPC.MoveNext
        Loop
        rsPC.Close

End Sub
