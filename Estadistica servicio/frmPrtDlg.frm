VERSION 5.00
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración Impresora/Página"
   ClientHeight    =   5670
   ClientLeft      =   2055
   ClientTop       =   1545
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5670
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vsViewLib.vsPrinter VP 
      Height          =   4760
      Left            =   3720
      TabIndex        =   36
      Top             =   300
      Width           =   4095
      _Version        =   196608
      _ExtentX        =   7223
      _ExtentY        =   8396
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
      PageBorder      =   7
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "Impresora..."
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   5280
      Width           =   1125
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "Página..."
      Height          =   300
      Index           =   1
      Left            =   1320
      TabIndex        =   34
      Top             =   5280
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "Intervalo de páginas"
      Height          =   975
      Left            =   60
      TabIndex        =   28
      Top             =   4080
      Width           =   3555
      Begin VB.TextBox tPaHasta 
         Height          =   285
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   32
         Top             =   570
         Width           =   495
      End
      Begin VB.TextBox tPaDesde 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   31
         Top             =   570
         Width           =   495
      End
      Begin VB.OptionButton oPaDesde 
         Caption         =   "Páginas:"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton oPaTodo 
         Caption         =   "Todo"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "a"
         Height          =   255
         Left            =   1920
         TabIndex        =   33
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   6600
      TabIndex        =   25
      Top             =   5280
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   5265
      TabIndex        =   24
      Top             =   5280
      Width           =   1200
   End
   Begin VB.ComboBox cmbDevice 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   300
      Width           =   3540
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tamaño de Página"
      Height          =   1875
      Index           =   2
      Left            =   60
      TabIndex        =   16
      Top             =   2100
      Width           =   3525
      Begin VB.OptionButton opOrient 
         Caption         =   "&Vertical"
         Height          =   315
         Index           =   0
         Left            =   900
         TabIndex        =   27
         Top             =   1260
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton opOrient 
         Caption         =   "&Horizontal"
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   26
         Top             =   1260
         Width           =   1095
      End
      Begin VB.VScrollBar scrlPaperSize 
         Height          =   300
         Index           =   1
         Left            =   2760
         Min             =   -32767
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   727
         Width           =   195
      End
      Begin VB.VScrollBar scrlPaperSize 
         Height          =   300
         Index           =   0
         Left            =   1320
         Min             =   -32767
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   727
         Width           =   195
      End
      Begin VB.TextBox txtPaperSize 
         Height          =   285
         Index           =   1
         Left            =   2220
         TabIndex        =   22
         Text            =   "11"""
         Top             =   735
         Width           =   555
      End
      Begin VB.TextBox txtPaperSize 
         Height          =   300
         Index           =   0
         Left            =   780
         TabIndex        =   19
         Text            =   "8.5"""
         Top             =   727
         Width           =   555
      End
      Begin VB.ComboBox cmbPaperSizes 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   300
         Width           =   2820
      End
      Begin VB.Image imgOrient 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "frmPrtDlg.frx":0000
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image imgOrient 
         Height          =   480
         Index           =   1
         Left            =   300
         Picture         =   "frmPrtDlg.frx":030A
         Top             =   1200
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Alto"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Top             =   780
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ancho"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   780
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formato"
      Height          =   1395
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   3525
      Begin VB.VScrollBar scrlMargin 
         Height          =   300
         Index           =   2
         Left            =   3180
         Min             =   -32767
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   900
         Width           =   195
      End
      Begin VB.TextBox txtMargin 
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   14
         Text            =   "1"""
         Top             =   915
         Width           =   555
      End
      Begin VB.VScrollBar scrlMargin 
         Height          =   300
         Index           =   3
         Left            =   3120
         Min             =   -32767
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   480
         Width           =   195
      End
      Begin VB.TextBox txtMargin 
         Height          =   285
         Index           =   3
         Left            =   2640
         TabIndex        =   8
         Text            =   "1"""
         Top             =   495
         Width           =   555
      End
      Begin VB.VScrollBar scrlMargin 
         Height          =   300
         Index           =   0
         Left            =   1560
         Min             =   -32767
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   900
         Width           =   195
      End
      Begin VB.TextBox txtMargin 
         Height          =   285
         Index           =   0
         Left            =   1020
         TabIndex        =   11
         Text            =   "1"""
         Top             =   915
         Width           =   555
      End
      Begin VB.VScrollBar scrlMargin 
         Height          =   300
         Index           =   1
         Left            =   1560
         Min             =   -32767
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   195
      End
      Begin VB.TextBox txtMargin 
         Height          =   285
         Index           =   1
         Left            =   1020
         TabIndex        =   5
         Text            =   "1"""
         Top             =   495
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Márgenes"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Izquierdo"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   4
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Derecho"
         Height          =   195
         Index           =   7
         Left            =   1920
         TabIndex        =   7
         Top             =   540
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Arriba"
         Height          =   195
         Index           =   6
         Left            =   495
         TabIndex        =   10
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Abajo"
         Height          =   195
         Index           =   9
         Left            =   2040
         TabIndex        =   13
         Top             =   960
         Width           =   420
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Impresoras"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   810
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PaperName(68) As String
Dim aControl As vsPrinter
Dim aPaginaD As Integer, aPaginaH As Integer
Dim bOK As Boolean

Public Property Get pOK() As Boolean
    pOK = bOK
End Property
Public Property Let pOK(OK As Boolean)
    pOK = OK
End Property

Public Property Get pControl() As vsPrinter: End Property
Public Property Let pControl(vsVontrol As vsPrinter)
    Set aControl = vsVontrol
End Property

Public Property Get pPaginaD() As Integer
    pPaginaD = aPaginaD
End Property
Public Property Let pPaginaD(Desde As Integer)
    aPaginaD = 1
End Property
Public Property Get pPaginaH() As Integer
    pPaginaH = aPaginaH
End Property
Public Property Let pPaginaH(Hasta As Integer)
    aPaginaH = aControl.PageCount
End Property

Function ToInches(ByVal v As Variant) As String
    
    v = Format(v / 1440, "#0.##")       ' convert from twips
    If Right(v, 1) = "." Then v = Left(v, Len(v) - 1)   ' trim string
    ToInches = v & """"     ' return measurement
    
End Function

Private Sub UpdatePreview()
    
    Dim I%, S$
    
    ' redraw
    DoEvents
    MousePointer = 11
    
    With VP
        ' load device list
        If cmbDevice.ListCount = 0 Then
            For I = 0 To .NDevices - 1
                cmbDevice.AddItem .Devices(I)
            Next
        End If
        
        ' select current device
        If .Device <> cmbDevice Then
            For I = 0 To .NDevices - 1
                If .Devices(I) = cmbDevice Then
                    cmbDevice.ListIndex = I
                    Exit For
                End If
            Next
            
            'Lista de hojas para la impresora seleccionada
            cmbPaperSizes.Clear
            For I = 1 To 68
                If PaperName(I) <> "" And .PaperSizes(I) = True Then
                    cmbPaperSizes.AddItem PaperName(I)
                    cmbPaperSizes.ItemData(cmbPaperSizes.NewIndex) = I
                End If
            Next
            If .PaperSizes(256) = True Then
                cmbPaperSizes.AddItem "Custom"
                cmbPaperSizes.ItemData(cmbPaperSizes.NewIndex) = 256
            End If
            
        End If
        
        For I = 0 To cmbPaperSizes.ListCount - 1
            If .PaperSize = cmbPaperSizes.ItemData(I) Then
                cmbPaperSizes.ListIndex = I
                Exit For
            End If
        Next
        
        ' show orientations
        I = .Orientation
        opOrient(I).Value = True
        imgOrient(I).Visible = True
        imgOrient(1 - I).Visible = False
        
        ' show margins
        If txtMargin(0) <> ToInches(.MarginTop) Then txtMargin(0) = ToInches(.MarginTop)
        If txtMargin(1) <> ToInches(.MarginLeft) Then txtMargin(1) = ToInches(.MarginLeft)
        If txtMargin(2) <> ToInches(.MarginBottom) Then txtMargin(2) = ToInches(.MarginBottom)
        If txtMargin(3) <> ToInches(.MarginRight) Then txtMargin(3) = ToInches(.MarginRight)
        
        ' select paper size
        If cmbPaperSizes.ListIndex > -1 Then
            If .PaperSize <> cmbPaperSizes.ItemData(cmbPaperSizes.ListIndex) Then
                .PaperSize = cmbPaperSizes.ItemData(cmbPaperSizes.ListIndex)
            End If
            I = .PaperSize
            txtPaperSize(0).Enabled = (I = 256)
            txtPaperSize(1).Enabled = (I = 256)
            scrlPaperSize(0).Enabled = (I = 256)
            scrlPaperSize(1).Enabled = (I = 256)
        End If
        
        ' show paper sizes
        .PhysicalPage = True
        If txtPaperSize(0) <> ToInches(.PageWidth) Then txtPaperSize(0) = ToInches(.PageWidth)
        If txtPaperSize(1) <> ToInches(.PageHeight) Then txtPaperSize(1) = ToInches(.PageHeight)
        .PhysicalPage = False
        
        .StartDoc
        S = "Impresora '" & .Device & "'"
        S = S & ", Driver '" & .Driver & "'"
        S = S & ", Puerto '" & .Port & "'"
        S = S & ", Papel #" & .PaperSize
        .StartDoc
        .FontSize = 12: .FontName = "Tahoma"
        'For I = 1 To 50: VP = S: Next
        .DrawRectangle .MarginLeft - 100, .MarginTop, .MarginLeft - 144, .MarginTop + 1440
        .EndDoc
    End With
    
    MousePointer = 0
    
End Sub

Private Sub cmbDevice_Click()
    UpdatePreview
End Sub

Private Sub cmbPaperSizes_Click()
    Dim I%
    With cmbPaperSizes
        I = .ItemData(.ListIndex)
        If VP.PaperSize = I Then Exit Sub
        'ClearPreview
        VP.PaperSize = I
        txtPaperSize(0).Enabled = (I = 256)
        txtPaperSize(1).Enabled = (I = 256)
        scrlPaperSize(0).Enabled = (I = 256)
        scrlPaperSize(1).Enabled = (I = 256)
        UpdatePreview
    End With
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDialog_Click(Index As Integer)
    
    ' show printer dialog
    If Index = 0 Then
        If Not VP.PrintDialog(pdPrinterSetup) Then Exit Sub
    
    ' show page setup dialog
    Else
        If Not VP.PrintDialog(pdPageSetup) Then Exit Sub
    End If
    
    ' update the preview
    UpdatePreview
End Sub

Private Sub cmdOK_Click()
    
    On Error Resume Next
    If oPaDesde.Value Then
        If Not IsNumeric(tPaDesde.Text) Then MsgBox "Error en la selección de páginas de impresión.", vbExclamation, "ATENCIÓN": tPaDesde.SetFocus: Exit Sub
        If Not IsNumeric(tPaHasta.Text) Then MsgBox "Error en la selección de páginas de impresión.", vbExclamation, "ATENCIÓN": tPaHasta.SetFocus: Exit Sub
        
        If Val(tPaDesde.Text) > Val(tPaHasta.Text) Then MsgBox "Error en la selección de páginas de impresión.", vbExclamation, "ATENCIÓN": tPaDesde.SetFocus: Exit Sub
        If Val(tPaHasta.Text) > aControl.PageCount Then MsgBox "Error en la selección de páginas de impresión.", vbExclamation, "ATENCIÓN": tPaDesde.SetFocus: Exit Sub
        aPaginaD = Val(tPaDesde.Text): aPaginaH = Val(tPaHasta.Text)
    End If
    
    aControl.Device = cmbDevice.Text
    aControl.PaperSize = VP.PaperSize
    
    aControl.Orientation = VP.Orientation
    aControl.MarginBottom = VP.MarginBottom
    aControl.MarginLeft = VP.MarginLeft
    aControl.MarginTop = VP.MarginTop
    aControl.MarginRight = VP.MarginRight
    
    bOK = True
    Unload Me
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    'define paper names
    PaperName(0) = ""
    PaperName(1) = "Letter 8 1/2 x 11 in"
    PaperName(2) = "Letter Small 8 1/2 x 11 in"
    PaperName(3) = "Tabloid 11 x 17 in"
    PaperName(4) = "Ledger 17 x 11 in"
    PaperName(5) = "Legal 8 1/2 x 14 in"
    PaperName(6) = "Statement 5 1/2 x 8 1/2 in"
    PaperName(7) = "Executive 7 1/4 x 10 1/2 in"
    PaperName(8) = "A3 297 x 420 mm"
    PaperName(9) = "A4 210 x 297 mm"
    PaperName(10) = "A4 Small 210 x 297 mm"
    PaperName(11) = "A5 148 x 210 mm"
    PaperName(12) = "B4 (JIS) 250 x 354"
    PaperName(13) = "B5 (JIS) 182 x 257 mm"
    PaperName(14) = "Folio 8 1/2 x 13 in"
    PaperName(15) = "Quarto 215 x 275 mm"
    PaperName(16) = "10x14 in"
    PaperName(17) = "11x17 in"
    PaperName(18) = "Note 8 1/2 x 11 in"
    PaperName(19) = "Envelope #9 3 7/8 x 8 7/8"
    PaperName(20) = "Envelope #10 4 1/8 x 9 1/2"
    PaperName(21) = "Envelope #11 4 1/2 x 10 3/8"
    PaperName(22) = "Envelope #12 4 \276 x 11"
    PaperName(23) = "Envelope #14 5 x 11 1/2"
    PaperName(24) = "C size sheet"
    PaperName(25) = "D size sheet"
    PaperName(26) = "E size sheet"
    PaperName(27) = "Envelope DL 110 x 220mm"
    PaperName(28) = "Envelope C5 162 x 229 mm"
    PaperName(29) = "Envelope C3  324 x 458 mm"
    PaperName(30) = "Envelope C4  229 x 324 mm"
    PaperName(31) = "Envelope C6  114 x 162 mm"
    PaperName(32) = "Envelope C65 114 x 229 mm"
    PaperName(33) = "Envelope B4  250 x 353 mm"
    PaperName(34) = "Envelope B5  176 x 250 mm"
    PaperName(35) = "Envelope B6  176 x 125 mm"
    PaperName(36) = "Envelope 110 x 230 mm"
    PaperName(37) = "Envelope Monarch 3.875 x 7.5 in"
    PaperName(38) = "6 3/4 Envelope 3 5/8 x 6 1/2 in"
    PaperName(39) = "US Std Fanfold 14 7/8 x 11 in"
    PaperName(40) = "German Std Fanfold 8 1/2 x 12 in"
    PaperName(41) = "German Legal Fanfold 8 1/2 x 13 in"
    PaperName(42) = "B4 (ISO) 250 x 353 mm"
    PaperName(43) = "Japanese Postcard 100 x 148 mm"
    PaperName(44) = "9 x 11 in"
    PaperName(45) = "10 x 11 in"
    PaperName(46) = "15 x 11 in"
    PaperName(47) = "Envelope Invite 220 x 220 mm"
    PaperName(48) = "" ' RESERVED--DO NOT USE
    PaperName(49) = "" ' RESERVED--DO NOT USE
    PaperName(50) = "Letter Extra 9 \275 x 12 in"
    PaperName(51) = "Legal Extra 9 \275 x 15 in"
    PaperName(52) = "Tabloid Extra 11.69 x 18 in"
    PaperName(53) = "A4 Extra 9.27 x 12.69 in"
    PaperName(54) = "Letter Transverse 8 \275 x 11 in"
    PaperName(55) = "A4 Transverse 210 x 297 mm"
    PaperName(56) = "Letter Extra Transverse 9\275 x 12 in"
    PaperName(57) = "SuperA/SuperA/A4 227 x 356 mm"
    PaperName(58) = "SuperB/SuperB/A3 305 x 487 mm"
    PaperName(59) = "Letter Plus 8.5 x 12.69 in"
    PaperName(60) = "A4 Plus 210 x 330 mm"
    PaperName(61) = "A5 Transverse 148 x 210 mm"
    PaperName(62) = "B5 (JIS) Transverse 182 x 257 mm"
    PaperName(63) = "A3 Extra 322 x 445 mm"
    PaperName(64) = "A5 Extra 174 x 235 mm"
    PaperName(65) = "B5 (ISO) Extra 201 x 276 mm"
    PaperName(66) = "A2 420 x 594 mm"
    PaperName(67) = "A3 Transverse 297 x 420 mm"
    PaperName(68) = "A3 Extra Transverse 322 x 445 mm"
    
    aPaginaD = 1
    aPaginaH = aControl.PageCount
    bOK = False
    
    Dim I%
    For I = 0 To 3
        scrlMargin(I).Tag = scrlMargin(I).Value
    Next
    For I = 0 To 1
        scrlPaperSize(I).Tag = scrlPaperSize(I).Value
    Next
    
    On Error Resume Next
    UpdatePreview
    'Cargo Valores del control
    With aControl
        cmbDevice.Text = .Device
        For I = 0 To cmbPaperSizes.ListCount - 1
            If .PaperSize = cmbPaperSizes.ItemData(I) Then
                cmbPaperSizes.ListIndex = I
                Exit For
            End If
        Next
        
        ' show orientations
        I = .Orientation
        opOrient(I).Value = True
        imgOrient(I).Visible = True
        imgOrient(1 - I).Visible = False
        
        ' show margins
        If txtMargin(0) <> ToInches(.MarginTop) Then txtMargin(0) = ToInches(.MarginTop)
        If txtMargin(1) <> ToInches(.MarginLeft) Then txtMargin(1) = ToInches(.MarginLeft)
        If txtMargin(2) <> ToInches(.MarginBottom) Then txtMargin(2) = ToInches(.MarginBottom)
        If txtMargin(3) <> ToInches(.MarginRight) Then txtMargin(3) = ToInches(.MarginRight)
    
    If .PageCount <> 0 Then
        tPaDesde.Text = aPaginaD
        tPaHasta.Text = aPaginaH
    End If
    
    End With
    
End Sub

Private Sub opOrient_Click(Index As Integer)
    
    If VP.Orientation = Index Then Exit Sub
    VP.Orientation = Index
    
End Sub

Private Sub scrlMargin_Change(Index As Integer)
    
    With scrlMargin(Index)
        If Val(.Value) < Val(.Tag) Then
            txtMargin(Index) = Val(txtMargin(Index)) + 0.1
        Else
            txtMargin(Index) = Val(txtMargin(Index)) - 0.1
        End If
        .Tag = .Value
        txtMargin_LostFocus (Index)
    End With
End Sub


Private Sub scrlPaperSize_Change(Index As Integer)
    
    With scrlPaperSize(Index)
        If Val(.Value) < Val(.Tag) Then
            txtPaperSize(Index) = Val(txtPaperSize(Index)) + 0.1
        Else
            txtPaperSize(Index) = Val(txtPaperSize(Index)) - 0.1
        End If
        .Tag = .Value
        txtPaperSize_LostFocus (Index)
    End With
    
End Sub


Private Sub tPaDesde_GotFocus()
    tPaDesde.SelStart = 0: tPaDesde.SelLength = Len(tPaDesde.Text)
End Sub

Private Sub tPaDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tPaHasta.SetFocus
End Sub

Private Sub tPaHasta_GotFocus()
    tPaHasta.SelStart = 0: tPaHasta.SelLength = Len(tPaHasta.Text)
End Sub

Private Sub txtMargin_GotFocus(Index As Integer)
    With txtMargin(Index)
        .SelStart = 0
        .SelLength = 32000
    End With
End Sub

Private Sub txtMargin_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtMargin_LostFocus (Index)
    End If
End Sub

Private Sub txtMargin_LostFocus(Index As Integer)
    
    Dim v
    With txtMargin(Index)
        v = Val(.Text) * 1440
        If v < 0 Then v = 0
        Select Case Index
            Case 0:
                If VP.MarginTop = v Then Exit Sub
                VP.MarginTop = v
            Case 1:
                If VP.MarginLeft = v Then Exit Sub
                VP.MarginLeft = v
            Case 2:
                If VP.MarginBottom = v Then Exit Sub
                VP.MarginBottom = v
            Case 3:
                If VP.MarginRight = v Then Exit Sub
                VP.MarginRight = v
        End Select
    End With
    UpdatePreview
    
End Sub

Private Sub txtPaperSize_GotFocus(Index As Integer)
    With txtPaperSize(Index)
        .SelStart = 0
        .SelLength = 32000
    End With
End Sub

Private Sub txtPaperSize_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPaperSize_LostFocus (Index)
    End If
End Sub


Private Sub txtPaperSize_LostFocus(Index As Integer)
    Dim v
    v = Val(txtPaperSize(Index)) * 1440
    If v < 0 Then v = 0
    If Index = 0 Then VP.PaperWidth = v Else VP.PaperHeight = v

End Sub


