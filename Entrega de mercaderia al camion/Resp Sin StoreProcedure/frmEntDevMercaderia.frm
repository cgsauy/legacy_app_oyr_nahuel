VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{191D08B9-4E92-4372-BF17-417911F14390}#1.5#0"; "orGridPreview.ocx"
Begin VB.Form frmEntDevMercaderia 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Entrega de Mercader�a"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEntDevMercaderia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgMini 
      Left            =   1440
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntDevMercaderia.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin orGridPreview.GridPreview gpPrint 
      Left            =   480
      Top             =   4560
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      BeginProperty FooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrid 
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5953
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
      BackColorFixed  =   14606046
      ForeColorFixed  =   -2147483630
      BackColorSel    =   13891065
      ForeColorSel    =   0
      BackColorBkg    =   16449535
      BackColorAlternate=   16448250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   4
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
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
   Begin VB.TextBox tUID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5520
      MaxLength       =   6
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox tCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   6
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   7440
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntDevMercaderia.frx":07B4
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntDevMercaderia.frx":08C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntDevMercaderia.frx":09D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntDevMercaderia.frx":0AEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntDevMercaderia.frx":0F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntDevMercaderia.frx":104E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TooMenu 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   582
      ButtonWidth     =   2196
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpiar"
            Key             =   "undo"
            Object.ToolTipText     =   "Cancelar datos ingresados"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Grabar "
            Key             =   "save"
            Object.ToolTipText     =   "Almacenar informaci�n"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Previa"
            Key             =   "preview"
            Object.ToolTipText     =   "Impresi�n previa"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "aPrint"
                  Text            =   "A Impresora"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "aScreen"
                  Text            =   "A Pantalla"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Buscar"
            Key             =   "find"
            Object.ToolTipText     =   "Buscar c�digos pendientes"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "2do Local"
            Key             =   "cnfgLocal"
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tolero"
                  Text            =   "Tolerancia"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "local2"
                  Text            =   "Local Secundario"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Impresora"
            Key             =   "cnfgprint"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Label lSec 
      BackStyle       =   0  'Transparent
      Caption         =   "Secundario sin tanto stock"
      ForeColor       =   &H00000080&
      Height          =   555
      Left            =   6840
      TabIndex        =   7
      Top             =   420
      Width           =   1290
   End
   Begin VB.Image imgLocal2 
      Height          =   480
      Left            =   6360
      Picture         =   "frmEntDevMercaderia.frx":14A0
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lCamion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cami�n: don pepe el pocho."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&C�digo:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmEntDevMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iTol As Integer
Private iLEnt As Long
Private sLEnt As String

Private Enum TipoSuceso
    DiferenciaDeArticulos = 11
End Enum

Private Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

Public prm_Tipo As Byte             '1 = entrega , 2 = Devoluci�n.
Public prm_Terminal As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errKP
    If Shift = vbAltMask Then
        Select Case KeyCode
            Case vbKeyL: If TooMenu.Buttons("undo").Enabled Then s_CtrlClean
            Case vbKeyG: If TooMenu.Buttons("save").Enabled Then act_Save
            Case vbKeyP: If prm_Tipo = 1 And TooMenu.Buttons("preview").Enabled Then act_Print True, True
            Case vbKeyB: If prm_Tipo = 1 And TooMenu.Buttons("find").Enabled Then s_FindPendiente
            Case vbKeyI: If prm_Tipo = 1 And TooMenu.Buttons("cnfgprint").Enabled Then prj_GetPrinter True
        End Select
    End If
    Exit Sub
errKP:
End Sub

Private Sub Form_Load()

    ObtengoSeteoForm Me, 1000, 500
    If prm_Tipo = 0 Then prm_Tipo = 1
        
    With TooMenu
        .Buttons("preview").Visible = (prm_Tipo = 1)
        .Buttons("find").Visible = (prm_Tipo = 1)
        .Buttons("cnfgLocal").Visible = (prm_Tipo = 1)
        .Buttons("cnfgprint").Visible = (prm_Tipo = 1)
    End With
    
    With vsGrid
        .Editable = False: .Rows = 1: .Cols = 1: .ExtendLastCol = True
        .FormatString = IIf(prm_Tipo = 1, "Entregar", "Devuelve") & "|>Ya Tiene |>Necesita|<Art�culo|Stock"
        .ColHidden(.Cols - 1) = True
    End With
    s_CtrlClean
    
    Dim fHeader As New StdFont, ffooter As New StdFont
    With fHeader
        .Bold = True
        .Name = "Arial"
        .Size = 11
    End With
    With ffooter
        .Bold = True
        .Name = "Tahoma"
        .Size = 10
    End With
    
    With gpPrint
        .Caption = IIf(prm_Tipo = 1, "Entrega de Mercader�a", "Devoluci�n de Mercader�a")
        .FileName = IIf(prm_Tipo = 1, "EntregaMercader�a", "Devoluci�nMercader�a")
        .Font = Font
        Set .HeaderFont = fHeader
        .Orientation = opPortrait
        .PaperSize = 1
        .PageBorder = opTopBottom
        .MarginLeft = 600
        .MarginRight = 600
    End With
    Me.Caption = IIf(prm_Tipo = 1, "Entrega de Mercader�a", "Devoluci�n de Entrega de Mercader�a") & " al Cami�n"
    
    If prm_Tipo = 1 Then
        Dim sD As String
        sD = GetSetting(App.Title, "Settings", "tolerancia", "")
        If IsNumeric(sD) Then iTol = Val(sD)
        sD = GetSetting(App.Title, "Settings", "localSecNombre", "")
        If sD <> "" Then sLEnt = sD
        sD = GetSetting(App.Title, "Settings", "localSecID", "")
        If IsNumeric(sD) Then iLEnt = Val(sD)
        
        If sLEnt <> "" Then TooMenu.Buttons("cnfgLocal").Caption = sLEnt
    End If
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    With vsGrid
        .Move 0, .Top, ScaleWidth, ScaleHeight - .Top
        If prm_Tipo = 1 Then
            .ColWidth(3) = .ClientWidth - .ColWidth(0) - .ColWidth(1) - .ColWidth(2) - 1500
        End If
    End With
End Sub

Private Sub s_CtrlClean()
    imgLocal2.Visible = False
    lSec.Visible = False
    lCamion.Caption = ""
    tCodigo.Tag = ""
    tUID.Text = "": tUID.Enabled = False: tUID.BackColor = vbButtonFace
    vsGrid.Rows = vsGrid.FixedRows
    With TooMenu
        .Buttons("save").Enabled = False
        .Buttons("preview").Enabled = False
    End With
    
End Sub

Private Sub s_FindImpresion()
On Error GoTo errFI
Dim lAux As Long, sFEdit As String
    
    Screen.MousePointer = 11
    s_CtrlClean
    Cons = "Select ArtID, ArtCodigo, rTrim(ArtNombre) As ArtNombre, ReECantidadTotal, IsNull(ReECantidadEntregada, 0) as QCamion, rTrim(CamNombre) as CamNombre,  ReECamion, ReEFModificacion" & _
            " From RenglonEntrega, Articulo, Camion Where ReECodImpresion = " & Val(tCodigo.Text) & _
            " And ReEArticulo = ArtID And ReECamion = CamCodigo"
    
    If prm_Tipo = 1 Then
        Cons = Cons & " And (ReECantidadTotal > ReECantidadEntregada or ReECantidadEntregada Is Null)"
    Else
        Cons = Cons & " And ReECantidadEntregada > 0 "
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "No existen datos para el c�digo ingresado.", vbExclamation, "Atenci�n"
    Else
        tCodigo.Tag = Val(tCodigo.Text)
        With lCamion
            .Caption = "Cami�n: " & RsAux!CamNombre
            .Tag = RsAux!ReECamion
        End With
        Do While Not RsAux.EOF
            With vsGrid
                If prm_Tipo = 1 Then
                    .AddItem RsAux!ReECantidadTotal - RsAux("QCamion")
                    '....................................................GUARDO EL STOCK DEL LOCAL
                    .Cell(flexcpData, .Rows - 1, 4) = f_StockLocalArticuloSano(RsAux!ArtID, paCodigoDeSucursal)
                    '....................................................GUARDO EL STOCK DEL LOCAL
                    
                    
                    If .Cell(flexcpData, .Rows - 1, 4) < .Cell(flexcpValue, .Rows - 1, 0) Then
                        'No hay tanto stock
                        If .Cell(flexcpData, .Rows - 1, 4) > 0 Then
                            .Cell(flexcpText, .Rows - 1, 0) = .Cell(flexcpData, .Rows - 1, 4)
                        Else
                            .Cell(flexcpText, .Rows - 1, 0) = 0
                        End If
                    End If
                Else
                    .AddItem 0
                End If
                .Cell(flexcpText, .Rows - 1, 1) = RsAux("QCamion")
                .Cell(flexcpText, .Rows - 1, 2) = RsAux!ReECantidadTotal
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ArtCodigo, "000,000") & " " & RsAux!ArtNombre
                                                
                'Si tengo local Secundario --> pido el stock
                If prm_Tipo = 1 Then
                    .Cell(flexcpData, .Rows - 1, 3) = f_StockLocalArticuloSano(RsAux!ArtID, iLEnt)
                    'Si la cantidad que necesita cargar es menor a la que tiene el local --> le cambio el color a lo que necesita.
                    If Val(.Cell(flexcpData, .Rows - 1, 3)) < RsAux("ReECantidadTotal") + iTol Then
                        .Cell(flexcpForeColor, .Rows - 1, 2) = &H80&
                        '.Cell(flexcpText, .Rows - 1, 0) = "� " & .Cell(flexcpText, .Rows - 1, 2)
                        .Cell(flexcpPicture, .Rows - 1, 0) = imgMini.ListImages(1).Picture
                        imgLocal2.Visible = True
                        lSec.Visible = True
                    End If
                End If
                                                
                lAux = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = lAux
                
                'V�lido Multiusuario.
                sFEdit = RsAux("ReEFModificacion"): .Cell(flexcpData, .Rows - 1, 1) = sFEdit
            End With
            RsAux.MoveNext
        Loop
        
        If vsGrid.Rows > vsGrid.FixedRows Then
            tUID.Enabled = True: tUID.BackColor = vbWindowBackground
            With TooMenu
                .Buttons("save").Enabled = True
                .Buttons("preview").Enabled = True
            End With
            With vsGrid
                .Cell(flexcpFontBold, .FixedRows, 0, .Rows - 1) = True
                .Cell(flexcpForeColor, .FixedRows, 1, .Rows - 1) = &H800000   '&H808000
'                .Cell(flexcpForeColor, .FixedRows, 2, .Rows - 1) = &H8000000C
                .SetFocus
            End With
        End If
        
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errFI:
    objGral.OcurrioError "Error al buscar la informaci�n para el c�digo ingresado.", Err.Description
    vsGrid.Rows = 1
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    Set objGral = Nothing
    CierroConexion
End Sub

Private Sub Label1_Click()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label2_Click()
    With tUID
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub tCodigo_Change()
    If Val(tCodigo.Tag) > 0 Then s_CtrlClean
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tCodigo.Tag) > 0 And tUID.Enabled Then tUID.SetFocus: Exit Sub
        If IsNumeric(tCodigo.Text) Then
            s_FindImpresion
        Else
            MsgBox "Debe ingresar un n�mero.", vbCritical, "Atenci�n"
        End If
    End If
End Sub

Private Sub TooMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case LCase(Button.Key)
        Case "undo": s_CtrlClean
        Case "save": act_Save
        Case "preview": act_Print True, True
        Case "find": s_FindPendiente
        Case "cnfgprint": prj_GetPrinter True
    End Select
End Sub

Private Sub TooMenu_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case LCase(ButtonMenu.Key)
        Case "aprint": act_Print True, True
        Case "ascreen": act_Print False, True
        Case "local2": act_PidoLocalSecundario
        Case "tolero": act_PidoTolerancia
    End Select
End Sub

Private Sub tUID_Change()
    If Val(tUID.Tag) > 0 Then tUID.Tag = ""
End Sub

Private Sub tUID_GotFocus()
    With tUID
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tUID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With tUID
            If IsNumeric(.Text) Then
                .Tag = BuscoUsuarioDigito(Val(.Text), True)
                If Val(.Tag) > 0 And Val(tCodigo.Tag) > 0 And vsGrid.Rows > vsGrid.FixedRows Then act_Save
            Else
                .Tag = 0
                MsgBox "Ingrese su d�gito de usuario.", vbExclamation, "ATENCI�N"
            End If
        End With
    End If
End Sub

Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    
    With vsGrid
        If .Rows > .FixedRows Then
            Select Case KeyCode
                Case vbKeyN
                    .Cell(flexcpText, .Row, 0) = "0"
                    .Cell(flexcpBackColor, .Row, 0) = vbWindowBackground
                    
                Case vbKeyS
                    'Si es Entrega:
                        'Aca va todo aunque el stock diga otra cosa.
                    'Sino es lo que ya tiene.
                    .Cell(flexcpText, .Row, 0) = .Cell(flexcpText, .Row, IIf(prm_Tipo = 1, 2, 1))
                    
                Case vbKeyAdd
                    If prm_Tipo = 1 Then
                        If CLng(.Cell(flexcpText, .Row, 0)) + CInt(.Cell(flexcpText, .Row, 1)) < CInt(.Cell(flexcpText, .Row, 2)) Then
                            .Cell(flexcpText, .Row, 0) = CInt(.Cell(flexcpText, .Row, 0)) + 1
                        End If
                    Else
                        If CLng(.Cell(flexcpText, .Row, 0)) < CInt(.Cell(flexcpText, .Row, 1)) Then
                            .Cell(flexcpText, .Row, 0) = CInt(.Cell(flexcpText, .Row, 0)) + 1
                        End If
                    End If
                Case vbKeySubtract
                    If CInt(.Cell(flexcpText, .Row, 0)) > 0 Then
                        .Cell(flexcpText, .Row, 0) = CInt(.Cell(flexcpText, .Row, 0)) - 1
                    End If
                
                Case vbKeyReturn: tUID.SetFocus
            End Select
        End If
    End With

End Sub

Private Sub s_FindPendiente()
On Error GoTo errFP
Dim lSelect As Long
    Screen.MousePointer = 11
    lSelect = 0
    Dim objLista As New clsListadeAyuda
    Cons = "Select Distinct(ReECodImpresion) as C�digo, CamNombre as 'Cami�n' " & _
                "From RenglonEntrega, Camion Where ReECamion = CamCodigo " & _
                "And ReECantidadEntregada <> ReECantidadTotal"
    If objLista.ActivarAyuda(cBase, Cons, 4200, titulo:="C�digos Pendientes") > 0 Then
        lSelect = objLista.RetornoDatoSeleccionado(0)
    End If
    Set objLista = Nothing
    If lSelect > 0 Then
        s_CtrlClean
        tCodigo.Text = lSelect
        Call tCodigo_KeyPress(vbKeyReturn)
    End If
    Screen.MousePointer = 0
    Exit Sub
errFP:
    objGral.OcurrioError "Error al buscar los c�digos pendientes.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function BuscoUsuarioDigito(Digito As Long, Optional Codigo As Boolean = False, Optional Identificacion As Boolean = False, Optional Iniciales As Boolean = False) As Variant
Dim RsUsr As rdoResultset
Dim aRetorno As Variant
On Error GoTo ErrBUD
    Screen.MousePointer = 11
    Cons = "Select * from Usuario Where UsuDigito = " & Digito
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Codigo Then aRetorno = RsUsr!UsuCodigo
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    BuscoUsuarioDigito = aRetorno
    Screen.MousePointer = 0
    Exit Function
ErrBUD:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar el usuario.", Err.Description
End Function

Private Sub act_Save()
    
    If Val(tUID.Tag) = 0 Then
        MsgBox "Ingrese su d�gito.", vbExclamation, "Atenci�n"
        tUID.SetFocus
        Exit Sub
    End If
    
    Dim iQ As Integer, bSigo As Boolean
    bSigo = False
    With vsGrid
        For iQ = .FixedRows To .Rows - .FixedRows
            If .Cell(flexcpValue, iQ, 0) > 0 Then
                bSigo = True
                Exit For
            End If
        Next
    End With
    If Not bSigo Then
        MsgBox "No hay datos en la lista, todos los art�culos tienen cantidad cero.", vbExclamation, "Atenci�n"
        Exit Sub
    End If
    
    If MsgBox("�Confirma grabar la informaci�n?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        'V�lido datos.
        If db_Save Then
            s_CtrlClean
        End If
    End If
    
End Sub

Private Function db_Save() As Boolean
Dim iQ As Integer, sSuceso As String, sErr As String
Dim rsRE As rdoResultset
    
    On Error Resume Next
    FechaDelServidor
    db_Save = False
    Screen.MousePointer = 11
    
    On Error GoTo errBT
    cBase.BeginTrans
    On Error GoTo errRB
   
    With vsGrid
        For iQ = .FixedRows To .Rows - 1
            If .Cell(flexcpValue, iQ, 0) > 0 Then
                'Pasos
                    '1) agregar o quitar en tabla rengl�n reparto.
                    '2) dar o quitar mercader�a al local y al cami�n
                    '3) marcar movimiento f�sico.
                    '4) grabar suceso silencioso.
                                       
                Cons = "Select * From RenglonEntrega" & _
                    " Where ReECodImpresion = " & CLng(tCodigo.Tag) & " And ReEArticulo = " & .Cell(flexcpData, iQ, 0)
                Set rsRE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If CDate(.Cell(flexcpData, iQ, 1)) <> rsRE("ReEFModificacion") Then
                    sErr = "C�digo de impresi�n modificado por otra terminal, tiene que volver a cargar toda la informaci�n."
                    rsRE.Close
                End If
                rsRE.Edit
                If IsNull(rsRE("ReECantidadEntregada")) Then
                    rsRE("ReECantidadEntregada") = (.Cell(flexcpValue, iQ, 0) * IIf(prm_Tipo = 1, 1, -1))
                Else
                    rsRE("ReECantidadEntregada") = rsRE("ReECantidadEntregada") + (.Cell(flexcpValue, iQ, 0) * IIf(prm_Tipo = 1, 1, -1))
                End If
                rsRE("ReEFModificacion") = Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss")
                rsRE("ReEUsuario") = Val(tUID.Tag)
                rsRE.Update
                rsRE.Close

                modStock.MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, .Cell(flexcpData, iQ, 0), .Cell(flexcpValue, iQ, 0), paEstadoArticuloEntrega, IIf(prm_Tipo = 1, -1, 1)
                modStock.MarcoMovimientoStockFisicoEnLocal TipoLocal.Camion, CLng(lCamion.Tag), .Cell(flexcpData, iQ, 0), .Cell(flexcpValue, iQ, 0), paEstadoArticuloEntrega, IIf(prm_Tipo = 1, 1, -1)

                'mov. f�sicos.
                modStock.MarcoMovimientoStockFisico CLng(tUID.Tag), TipoLocal.Deposito, paCodigoDeSucursal, .Cell(flexcpData, iQ, 0), .Cell(flexcpValue, iQ, 0), paEstadoArticuloEntrega, IIf(prm_Tipo = 1, -1, 1), 21, CLng(tCodigo.Tag)
                modStock.MarcoMovimientoStockFisico CLng(tUID.Tag), TipoLocal.Camion, CLng(lCamion.Tag), .Cell(flexcpData, iQ, 0), .Cell(flexcpValue, iQ, 0), paEstadoArticuloEntrega, IIf(prm_Tipo = 1, 1, -1), 21, CLng(tCodigo.Tag)

                'Suceso solo en entrega.
                If prm_Tipo = 1 Then
                    If .Cell(flexcpData, iQ, 4) < .Cell(flexcpValue, iQ, 0) Then
                        sSuceso = "sin haber stock (stocklocal = " & .Cell(flexcpData, iQ, 4) & ")."
                    End If
                    
                    If .Cell(flexcpData, iQ, 4) - .Cell(flexcpValue, iQ, 0) < 0 Then
                        If sSuceso <> "" Then
                            sSuceso = sSuceso & " Quedo Stock Negativo."
                        Else
                            sSuceso = " y quedo Stock Negativo."
                        End If
                    End If
                    
                    If sSuceso <> "" Then
                        objGral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.DiferenciaDeArticulos, paCodigoDeTerminal, CLng(tUID.Tag), 0, CLng(.Cell(flexcpData, iQ, 0)), _
                                Descripcion:="Entrega de Mercader�a al Cami�n, c�digo: " & tCodigo.Text, Defensa:="Entrega " & CInt(.Cell(flexcpValue, iQ, 0)) & " art(s). " & Mid(Trim(.Cell(flexcpText, iQ, 3)), 1, 15) & "... " & sSuceso
                    End If
                    
                End If
                '.................Suceso
                
            End If
        Next iQ
    End With
    cBase.CommitTrans

    act_Print True, False
    db_Save = True
    Screen.MousePointer = 0
    Exit Function
    
errBT:
    objGral.OcurrioError "Error al iniciar la transacci�n.", Err.Description
    Screen.MousePointer = 0
    Exit Function
    
errError:
    cBase.RollbackTrans
    objGral.OcurrioError "Error al grabar", Err.Description & IIf(sErr <> "", vbCr & sErr, "")
    Screen.MousePointer = 0
    Exit Function
    
errRB:
    Resume errError
End Function

Private Sub loc_AddRenglonEntrega(ByVal lIDImp As Long, ByVal lArt As Long, ByVal iQ As Integer, ByVal iCamion As Integer)
Dim rsRE As rdoResultset
    Cons = "Select * From RenglonEntrega" & _
                " Where ReECodImpresion = " & lIDImp & _
                " And ReEArticulo = " & lArt
    Set rsRE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rsRE.Edit
    rsRE("ReECantidadTotal") = iQ + rsRE("ReECantidadTotal")
    rsRE("ReEFModificacion") = Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss")
    rsRE("ReEUsuario") = Val(tUID.Tag)
    rsRE.Update
    rsRE.Close
End Sub


Private Function f_StockLocalArticuloSano(ByVal lArticulo As Long, ByVal iLocal As Long) As Currency
On Error GoTo errSTL
Dim Rs As rdoResultset
    Screen.MousePointer = 11
    f_StockLocalArticuloSano = 0
    Cons = "Select Sum(StLCantidad) From StockLocal " _
        & " Where StLArticulo = " & lArticulo & " And StlTipoLocal = " & TipoLocal.Deposito _
        & " And StLLocal = " & iLocal & " And StLEstado = " & paEstadoArticuloEntrega
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then
        If Not IsNull(Rs(0)) Then f_StockLocalArticuloSano = Rs(0)
    End If
    Rs.Close
    Screen.MousePointer = 0
    Exit Function
errSTL:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error inesperado al buscar el stock del local.", Err.Description
End Function

Private Sub act_Print(Optional Imprimir As Boolean = False, Optional bDetalle As Boolean = False)
On Error GoTo errImprimir
  
    Screen.MousePointer = 11
    vsGrid.ExtendLastCol = False
    With gpPrint
        .Device = paPrintConfD
        .PaperBin = paPrintConfB
        
        .Header = IIf(prm_Tipo = 1, "Entrega de Mercader�a al Cami�n", "Devoluci�n de Mercader�a del Cami�n")
        
        .LineBeforeGrid "C�digo de Impresi�n: " & Val(tCodigo.Text) & Space(20) & lCamion.Caption, , , True
        .LineBeforeGrid "Sucursal = " & paNombreSucursal & Space(10) & "Terminal = " & prm_Terminal
        .LineBeforeGrid ""
        If bDetalle Then
            .LineBeforeGrid "Detalle posible de Entrega", bbold:=True, bunderline:=True
            .LineBeforeGrid ""
        End If
        
        .AddGrid vsGrid.hWnd
        
        .LineAfterGrid ""
        If Imprimir Then .LineAfterGrid "D�gito de Usuario: " & tUID.Text: .LineAfterGrid ""
        If Not bDetalle Then .LineAfterGrid "Firma: ..........................................................."
    
        If Imprimir Then
            .ChoosePrint = False
            .GoPrint
        Else
            .ShowPreview
        End If
    End With
    
    vsGrid.ExtendLastCol = True
    
    Screen.MousePointer = 0
    Exit Sub
errImprimir:
    Screen.MousePointer = 0
    objGral.OcurrioError "Ocurri� un error al realizar la impresi�n", Err.Description
    vsGrid.ExtendLastCol = False
End Sub

Private Sub act_PidoTolerancia()
Dim sTol As String
    sTol = InputBox("Ingrese la cantidad de art�culos de tolerancia para el local secundario", "Tolerancia", IIf(iTol > 0, iTol, 1))
    If Not IsNumeric(sTol) Then
        sTol = "0"
        If MsgBox("�Confirma limpiar el par�metro de tolerancia?", vbQuestion + vbYesNo, "Atenci�n") = vbNo Then Exit Sub
    End If
    SaveSetting App.Title, "Settings", "tolerancia", sTol
    iTol = Val(sTol)
    MsgBox "Los cambios surgen efecto al cargar un nuevo c�digo de impresi�n.", vbInformation, "Atenci�n"
End Sub

Private Sub act_PidoLocalSecundario()
Dim sLoc As String
    sLoc = InputBox("Ingrese parte o todo el nombre del local secundario", "Local secundario", "")
    If Trim(sLoc) = "" Then
        sLoc = ""
        If MsgBox("�Confirma limpiar el par�metro 'Local secundario'?", vbQuestion + vbYesNo, "Atenci�n") = vbNo Then Exit Sub
    End If
    
    If sLoc <> "" Then
        'Busco el local
        Dim objH As New clsListadeAyuda
        If objH.ActivarAyuda(cBase, "Select LocCodigo, LocNombre as 'Nombre' From Local Where LocNombre Like '" & Replace(sLoc, " ", "%") & "%'", 3500, 1, "Locales") > 0 Then
            sLoc = "OK"
            iLEnt = objH.RetornoDatoSeleccionado(0)
            sLEnt = objH.RetornoDatoSeleccionado(1)
        Else
            sLoc = ""
        End If
        Set objH = Nothing
        If sLoc = "" Then Exit Sub
    Else
        sLEnt = ""
        iLEnt = 0
    End If
    SaveSetting App.Title, "Settings", "localSecNombre", sLEnt
    SaveSetting App.Title, "Settings", "localSecID", iLEnt
    
    MsgBox "Los cambios surgen efecto al cargar un nuevo c�digo de impresi�n.", vbInformation, "Atenci�n"
    TooMenu.Buttons("cnfgLocal").Caption = IIf(sLEnt <> "", sLEnt, "2do Local")
    
End Sub

