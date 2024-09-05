VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{B443E3A5-0B4D-4B43-B11D-47B68DC130D7}#1.7#0"; "orArticulo.ocx"
Begin VB.Form frmAgenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Fletes"
   ClientHeight    =   5880
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAgendaTF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9360
   Begin VB.TextBox tCobra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   12
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox tTipoPrecio 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin prjFindArticulo.orArticulo tArticulo 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1200
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.TextBox tNCorto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox chAgencia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Por &Agencia"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox cFormaPago 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox cRangoHora 
      Height          =   315
      Left            =   1200
      TabIndex        =   26
      Text            =   "Combo1"
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox tZona 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   25
      TabIndex        =   16
      Top             =   2400
      Width           =   3255
   End
   Begin VB.ComboBox cSubGrupo 
      Height          =   315
      Left            =   5400
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox tZonaGrupoZona 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   25
      TabIndex        =   15
      Top             =   2400
      Width           =   3255
   End
   Begin MSComctlLib.TabStrip tsAgenda 
      Height          =   315
      Left            =   1200
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      MultiRow        =   -1  'True
      Style           =   1
      TabFixedWidth   =   1057
      TabMinWidth     =   1057
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Principal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Grupo Zona"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Zona"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   635
      ButtonWidth     =   2064
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            Key             =   "modificar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "cancelar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar staMensaje 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   5625
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tDescripcion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   25
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAgendaTF.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAgendaTF.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAgendaTF.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAgendaTF.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAgendaTF.frx":0752
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsAgenda 
      Height          =   1935
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3413
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483633
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
      AutoResize      =   0   'False
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
      Editable        =   -1  'True
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsSubGrupo 
      Height          =   2775
      Left            =   4560
      TabIndex        =   21
      Top             =   2760
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4895
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
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
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
      AutoResize      =   0   'False
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
      Editable        =   -1  'True
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsHoraEnvio 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
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
      AutoResize      =   0   'False
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
      Editable        =   -1  'True
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cobra:"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Precio:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ar&tículo:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "A&breviación:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Pago:"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Rango Horas:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5160
      Width           =   1035
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&SubGrupo:"
      Height          =   255
      Left            =   4560
      TabIndex        =   17
      Top             =   2400
      Width           =   795
   End
   Begin VB.Label lEs 
      BackStyle       =   0  'Transparent
      Caption         =   "&Grupo Zona:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Agenda"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label lTFNeedAgencia 
      BackStyle       =   0  'Transparent
      Caption         =   "&Flete:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   795
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Visible         =   0   'False
      Begin VB.Menu MnuOpNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu MnuOpModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuOpEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu MnuOpLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuOpCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Visible         =   0   'False
      Begin VB.Menu MnuSaOut 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuOptSubGrupo 
      Caption         =   "MnuOptSubGrupo"
      Visible         =   0   'False
      Begin VB.Menu MnuOSGNombre 
         Caption         =   "Nombre del subgrupo"
      End
      Begin VB.Menu MnuOSGIndependizar 
         Caption         =   "Independizar zona del grupo"
      End
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tHoraEnvio
    Indice As Integer
    Hora As String
End Type
Private arrHE() As tHoraEnvio

Private paPrimeraHoraEnvio  As Long, paUltimaHoraEnvio As Long
Private sHoraEnvio As String
Private iRangoHS As Integer
Private douAgenda As Double

Private Sub chAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tTipoPrecio.SetFocus
End Sub

Private Sub cRangoHora_GotFocus()
    With cRangoHora
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cRangoHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub cSubGrupo_Change()
    If cSubGrupo.ListIndex = -1 Then loc_HideShowSG
End Sub

Private Sub cSubGrupo_Click()
    loc_HideShowSG
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyM: If Toolbar1.Buttons("modificar").Enabled Then AccionModificar
            Case vbKeyC:  If Toolbar1.Buttons("cancelar").Enabled Then AccionCancelar
            Case vbKeyG:  If Toolbar1.Buttons("grabar").Enabled Then AccionGrabar
            Case vbKeyX: Unload Me
        End Select
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    
    With tArticulo
        Set .Connect = cBase
        .FindNombreEnUso = True
    End With
    
    ObtengoSeteoForm Me
    Me.Height = 6255
    Me.Width = 9450
    frm_OcultoCtrl
    CargoDiasHabilitados
    CargoCombo "Select THoID, THoNombre From TipoHorario Order By THoNombre", cRangoHora
    
    With cFormaPago
        .Clear
        .AddItem "Caja": .ItemData(.NewIndex) = 1
        .AddItem "Domicilio": .ItemData(.NewIndex) = 2
        .AddItem "Fact. Camión": .ItemData(.NewIndex) = 3
    End With
    
    db_LoadPrm
    
    With vsSubGrupo
        .Cols = 1
        .Rows = 1
        .FormatString = "Sub Grupo|Zona|"
        .ColHidden(2) = True
        .ExtendLastCol = True
    End With
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inciar el formulario.", Err.Description
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
End Sub

Private Sub CargoDiasHabilitados()
On Error GoTo errCDH
Dim lngCodigo As Long, intCol As Integer, intDia As Integer
Dim rsHora As rdoResultset
Dim cons As String

    With vsAgenda
        
        'Cargo las horaenvio que están en horarioflete.
        cons = "Select HEnNombre, HEnCodigo, HEnIndice, HEnInicio, HEnFin From HoraEnvio " & _
                    " Where HEnCodigo IN (Select Distinct(HFLCodigo) From HorarioFlete) Order By HEnIndice"
        Set rsHora = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If rsHora.EOF Then
            .Rows = 0
            .Cols = 0
        Else
            .Rows = 1
            .Cols = 1
            ReDim arrHE(0)
            Do While Not rsHora.EOF
                .Cols = .Cols + 1
                .Cell(flexcpText, 0, .Cols - 1) = Trim(rsHora(0))
                
                lngCodigo = rsHora!HEnCodigo: .Cell(flexcpData, 0, .Cols - 1) = lngCodigo
                .ColAlignment(.Cols - 1) = flexAlignCenterCenter
                
                ReDim Preserve arrHE(.Cols - 1)
                With arrHE(.Cols - 1)
                    .Indice = rsHora("HEnIndice")
                    .Hora = Format(rsHora("HEnInicio"), "0000") & "-" & Format(rsHora("HEnFin"), "0000")
                End With
                rsHora.MoveNext
            Loop
            
        End If
        rsHora.Close
    
        If .Rows = 0 Then Exit Sub
        
        cons = "Select * From HorarioFlete Order by HFlDiaSemana"
        Set rsHora = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        intDia = 0
        Do While Not rsHora.EOF
            If intDia <> rsHora!HFlDiaSemana Then
                intDia = rsHora!HFlDiaSemana
                .AddItem DiaSemana(intDia) & Space(5)
                .Cell(flexcpData, .Rows - 1, 0) = intDia
            End If
            intCol = ColumnaHora(rsHora!HFlCodigo)
            lngCodigo = rsHora!HFlIndice: .Cell(flexcpData, .Rows - 1, intCol) = lngCodigo
            .Cell(flexcpChecked, .Rows - 1, intCol) = 2
            .Cell(flexcpBackColor, .Rows - 1, intCol) = vbWindowBackground
            rsHora.MoveNext
        Loop
        rsHora.Close
    End With
    
    With vsHoraEnvio
        .Cols = vsAgenda.Cols
        .Cell(flexcpText, 0, 0) = "Hora Envío"
        If .Cols - 1 > 0 Then
            For intCol = 1 To .Cols - 1
                .ColWidth(intCol) = vsAgenda.ColWidth(intCol)
            Next
        End If
    End With
    Exit Sub
errCDH:
    clsGeneral.OcurrioError "Error al inicializar la agenda.", Err.Description
End Sub

Private Function ColumnaHora(lngHorario As Long) As Integer
On Error Resume Next
    Dim I As Integer
    ColumnaHora = -1
    For I = 1 To vsAgenda.Cols - 1
        If Val(vsAgenda.Cell(flexcpData, 0, I)) = lngHorario Then ColumnaHora = I: Exit Function
    Next I
End Function

Private Function DiaSemana(ByVal intDia As Integer) As String
    
    Select Case intDia
        Case 1: DiaSemana = "Domingo"
        Case 2: DiaSemana = "Lunes"
        Case 3: DiaSemana = "Martes"
        Case 4: DiaSemana = "Miércoles"
        Case 5: DiaSemana = "Jueves"
        Case 6: DiaSemana = "Viernes"
        Case 7: DiaSemana = "Sábado"
    End Select
    
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    cBase.Close
    Set clsGeneral = Nothing
End Sub

Private Sub Label1_Click()
    If cRangoHora.Enabled Then cRangoHora.SetFocus
End Sub

Private Sub Label2_Click()
    If cFormaPago.Enabled Then cFormaPago.SetFocus
End Sub

Private Sub Label5_Click()
    If tNCorto.Enabled Then tNCorto.SetFocus
End Sub

Private Sub Label6_Click()
    If tArticulo.Enabled Then tArticulo.SetFocus
End Sub

Private Sub Label7_Click()
    If tTipoPrecio.Enabled Then tTipoPrecio.SetFocus
End Sub

Private Sub Label8_Click()
     If tCobra.Enabled Then tCobra.SetFocus
End Sub

Private Sub lTFNeedAgencia_Click()
On Error Resume Next
    With tDescripcion
        If .Enabled Then
            .SelStart = 0: .SelLength = Len(.Text): .SetFocus
        End If
    End With
End Sub

Private Sub MnuOpCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuOpGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuOpModificar_Click()
    AccionModificar
End Sub

Private Sub MnuOSGIndependizar_Click()
    On Error GoTo errSG
    Dim sSubGrupo As String
    sSubGrupo = InputBox("Ingrese un nombre de subgrupo para la zona seleccionado.", "Sub Grupos", vsSubGrupo.Cell(flexcpText, vsSubGrupo.RowSel, 0))
    'Le modifico a todo el grupo el nombre
    If sSubGrupo <> vsSubGrupo.Cell(flexcpText, vsSubGrupo.RowSel, 0) Then
        Dim pregunta As String
        If sSubGrupo = "" Then
            pregunta = "¿Confirma dejar sin nombre al grupo " & sSubGrupo & "?"
        Else
            pregunta = "¿Confirma que el grupo ahora se llamará " & sSubGrupo & "?"
        End If
        If MsgBox(pregunta, vbYesNo, "GRABAR") = vbYes Then
            db_AsignoSubGrupo vsSubGrupo.Cell(flexcpData, vsSubGrupo.RowSel, 1), sSubGrupo
            ArmoAgendaFlete
        End If
    End If
    Exit Sub
errSG:
    clsGeneral.OcurrioError "Error al independizar el nombre del subgrupo.", Err.Description, "SubGrupos"
End Sub

Private Sub MnuOSGNombre_Click()
On Error GoTo errSG
    Dim sSubGrupo As String
    sSubGrupo = InputBox("Ingrese el nombre que le desea asignar al subgrupo.", "Sub Grupos", vsSubGrupo.Cell(flexcpText, vsSubGrupo.RowSel, 0))
    'Le modifico a todo el grupo el nombre
    If sSubGrupo <> vsSubGrupo.Cell(flexcpText, vsSubGrupo.RowSel, 0) Then
        Dim pregunta As String
        If sSubGrupo = "" Then
            pregunta = "¿Confirma dejar sin nombre al grupo " & sSubGrupo & "?"
        Else
            pregunta = "¿Confirma que el grupo ahora se llamará " & sSubGrupo & "?"
        End If
        If MsgBox(pregunta, vbYesNo, "GRABAR") = vbYes Then
            CambiarNombreASubGrupo vsSubGrupo.Cell(flexcpText, vsSubGrupo.RowSel, 0), sSubGrupo
            ArmoAgendaFlete
        End If
    End If
    Exit Sub
errSG:
    clsGeneral.OcurrioError "Error al asignar el nombre del subgrupo.", Err.Description, "SubGrupos"
End Sub

Private Sub CambiarNombreASubGrupo(ByVal nomActual As String, ByVal nomNuevo As String)
Dim iQ As Integer, grupo As String
    
    
    On Error GoTo errBT
    cBase.BeginTrans
    On Error GoTo errRB
    
    With vsSubGrupo
        grupo = .Cell(flexcpText, .Row, 0)
        For iQ = .FixedRows To .Rows - 1
            If .Cell(flexcpText, iQ, 0) = nomActual Then
                db_AsignoSubGrupo .Cell(flexcpData, iQ, 1), nomNuevo
            End If
        Next
    End With
    
    cBase.CommitTrans
    
    Exit Sub
    
errBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Err.Description
    Exit Sub
errResumo:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar grabar la información.", Err.Description
    Exit Sub
errRB:
    Resume errRB

End Sub


Private Sub MnuSaOut_Click()
    Unload Me
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chAgencia.SetFocus
End Sub

Private Sub tCobra_GotFocus()
    With tCobra
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCobra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then vsAgenda.SetFocus
End Sub

Private Sub tDescripcion_Change()
On Error Resume Next
    If Val(tDescripcion.Tag) > 0 Then
        tDescripcion.Tag = ""
        frm_OcultoCtrl
    End If
End Sub

Private Sub tDescripcion_GotFocus()
On Error Resume Next
    With tDescripcion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyM Then
        KeyCode = 0
    End If
End Sub

Private Sub tDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tDescripcion.Text) <> "" Then
        If Val(tDescripcion.Tag) = 0 Then
            loc_FindByText tDescripcion, "Select TFLCodigo, TFLDescripcion as 'Nombre' From TipoFlete " & _
                                                    "Where TFlDescripcion like '" & Replace(tDescripcion.Text, " ", "%") & "%' Order by 2"
            If Val(tDescripcion.Tag) > 0 Then
                fnc_GetDatosPcpal
                ArmoAgendaFlete
            End If
        End If
        miBotones (Val(tDescripcion.Tag) > 0), False, False
        tsAgenda.Enabled = (Val(tDescripcion.Tag) > 0)
        If tsAgenda.Enabled Then tsAgenda.SetFocus
    End If
End Sub

Private Sub miBotones(bolModificar As Boolean, bolGrabar As Boolean, bolCancelar As Boolean)
Dim bPermisoEditDatos As Boolean

    With Toolbar1
        .Buttons("modificar").Enabled = bolModificar
        .Buttons("grabar").Enabled = bolGrabar
        .Buttons("cancelar").Enabled = bolCancelar
    End With
    MnuOpModificar.Enabled = bolModificar
    MnuOpGrabar.Enabled = bolGrabar
    MnuOpCancelar.Enabled = bolCancelar
    
    With tDescripcion
        .Enabled = Not bolGrabar
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    vsAgenda.Enabled = bolGrabar
    With vsHoraEnvio
        .Enabled = bolGrabar
    End With
    
    If tsAgenda.SelectedItem.Index = 1 And bolGrabar Then
        Dim objC As New clsConexion
        bolGrabar = objC.AccesoAlMenu("TiposFletesDatos")
        Set objC = Nothing
    Else
        bolGrabar = False
    End If
    With cRangoHora
        .Enabled = bolGrabar
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    chAgencia.Enabled = bolGrabar
    With cFormaPago
        .Enabled = bolGrabar
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    With tArticulo
        .Enabled = bolGrabar
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    With tTipoPrecio
        .Enabled = bolGrabar
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    With tCobra
        .Enabled = bolGrabar
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    With tNCorto
        .Enabled = bolGrabar
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
End Sub

Private Sub tNCorto_GotFocus()
    With tNCorto
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tNCorto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cFormaPago.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    Select Case Button.Key
        Case "modificar": AccionModificar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
    End Select
End Sub

Private Sub LimpioDiasHabilitados()
On Error Resume Next
Dim Fila As Integer, Col As Integer
    
    With vsAgenda
        For Fila = 1 To .Rows - 1
            For Col = 1 To .Cols - 1
                If Val(.Cell(flexcpData, Fila, Col)) > 0 Then .Cell(flexcpChecked, Fila, Col) = flexUnchecked
            Next Col
        Next Fila
    End With
    cRangoHora.Text = "": cRangoHora.ListIndex = -1
    frm_CleanHoraEnvio
End Sub

Private Sub AccionModificar()
    
    If tsAgenda.SelectedItem.Index > 1 Then
        If tsAgenda.SelectedItem.Index = 2 And Val(tZonaGrupoZona.Tag) = 0 Then
            MsgBox "Seleccione un grupo zona.", vbExclamation, "Atención"
            tZonaGrupoZona.SetFocus
            Exit Sub
        ElseIf tsAgenda.SelectedItem.Index = 3 And Val(tZona.Tag) = 0 Then
            MsgBox "Seleccione una zona.", vbExclamation, "Atención"
            tZona.SetFocus
            Exit Sub
        ElseIf tsAgenda.SelectedItem.Index = 2 Then
            If vsSubGrupo.Rows = vsSubGrupo.FixedRows Then
                MsgBox "No hay zonas para este grupo.", vbInformation, "Atención"
                Exit Sub
            End If
        End If
    End If
    tsAgenda.Enabled = False
    cSubGrupo.Enabled = False
    vsSubGrupo.Enabled = False
    tZonaGrupoZona.Enabled = False
    tZona.Enabled = False
    miBotones False, True, True
    chAgencia.Enabled = tsAgenda.SelectedItem.Index = 1
    cFormaPago.Enabled = chAgencia.Enabled
    vsAgenda.SetFocus
    
End Sub

Private Sub AccionCancelar()
On Error Resume Next
    miBotones True, False, False
    tDescripcion.SetFocus
    tsAgenda.Enabled = True
    ArmoAgendaFlete
    
    With tZona
        .Enabled = tsAgenda.SelectedItem.Index = 3
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    With tZonaGrupoZona
        .Enabled = tsAgenda.SelectedItem.Index = 2
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    With cSubGrupo
        .Enabled = tsAgenda.SelectedItem.Index = 2
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    vsSubGrupo.Enabled = (tsAgenda.SelectedItem.Index = 2 And Val(tZonaGrupoZona.Tag) > 0)
    
End Sub

Private Sub ArmoAgendaFlete()
    LimpioDiasHabilitados
    Select Case tsAgenda.SelectedItem.Index
        Case 1
            If iRangoHS > -1 Then BuscoCodigoEnCombo cRangoHora, CLng(iRangoHS)
            If sHoraEnvio <> "" Then loc_SetHoraEnvio sHoraEnvio
            ArmoAgendaFleteEnGrilla douAgenda
            
        Case 2: ArmoAgendaGrupoZona
        Case 3: ArmoAgendaZona
    End Select
End Sub

Private Sub ArmoAgendaFleteEnGrilla(ByVal douFleteAgenda As Double)
Dim strMat As String, strAux As String
On Error GoTo errSalir

    Screen.MousePointer = 11
    If douFleteAgenda = 0 Then douFleteAgenda = douAgenda
    strMat = superp_MatrizSuperposicion(douFleteAgenda)
    If strMat = "" Then GoTo errSalir
    
    Do While strMat <> ""
        If InStr(1, strMat, ",") > 0 Then
            MarcoEnGrilla CInt(Mid(strMat, 1, InStr(1, strMat, ",") - 1))
            strMat = Mid(strMat, InStr(1, strMat, ",") + 1, Len(strMat))
        Else
            MarcoEnGrilla CInt(strMat)
            strMat = ""
        End If
    Loop
    
    
errSalir:
    Screen.MousePointer = 0
End Sub

Private Sub MarcoEnGrilla(intIndice As Integer)
On Error Resume Next
Dim Fila As Integer, Col As Integer
    For Fila = 1 To vsAgenda.Rows - 1
        For Col = 1 To vsAgenda.Cols - 1
            If Val(vsAgenda.Cell(flexcpData, Fila, Col)) = intIndice Then vsAgenda.Cell(flexcpChecked, Fila, Col) = flexChecked
        Next Col
    Next Fila
End Sub

Private Sub AccionGrabar()
Dim sHE As String
Dim douAux As Double
Dim bChange As Boolean

    'Válido campos
    If Not fnc_ValidoGrabar Then Exit Sub
    
    If MsgBox("¿Confirma almacenar la agenda ingresada?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
        bChange = False
        On Error GoTo errGrabar
        Screen.MousePointer = 11
        sHE = fnc_GetHoraEnvioGrid
        douAux = CalculoValorSuperposicion
        If tsAgenda.SelectedItem.Index = 1 Then
            db_SavePrincipal douAux, sHE, bChange
            If bChange Then
                If MsgBox("La agenda principal fué modificada" & vbCrLf & vbCrLf & "¿Desea aplicar el cambio a todas las zonas asigandas al tipo de flete?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
                    ModificoZonasTipoFlete douAux, sHE
                End If
            End If
        ElseIf tsAgenda.SelectedItem.Index = 2 Then
            'Para todas las zonas del grupo armo la agenda.
            If Not db_SaveGrupoZona(douAux, sHE, bChange) Then Exit Sub
        Else
            'Zona
            db_SaveAgendaZona Val(tZona.Tag), douAux, sHE, bChange, False
        End If
        
        If bChange Then
            MsgBox "Al modificar la agenda, también se modificará la agenda habilitada." & vbCr & vbCr & "Verifique si la agenda habilitada quedo correcta.", vbInformation, "Atención"
        End If
        
        AccionCancelar
        Screen.MousePointer = 0
    End If
    Exit Sub
errGrabar:
    clsGeneral.OcurrioError "Ocurrió un error al intentar almacenar la información.", Err.Description
    Screen.MousePointer = 0
End Sub

Sub ModificoZonasTipoFlete(ByVal douAux As Double, ByVal sHE As String)
Dim rsZ As rdoResultset
Dim cons As String
    cons = "Select * From FleteAgendaZona " & _
                " WHERE FAZTipoFlete = " & Val(tDescripcion.Tag)
    Set rsZ = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsZ.EOF
        db_SaveAgendaZona rsZ("FAZZona"), douAux, sHE, True, True
        rsZ.MoveNext
    Loop
    rsZ.Close

End Sub

Private Sub db_SavePrincipal(ByVal douAux As Double, ByVal sHE As String, ByRef bChange As Boolean)
Dim cons As String
Dim rsAux As rdoResultset

    cons = "Select * from TipoFlete Where TFlCodigo = " & Val(tDescripcion.Tag)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsAux.Edit
    rsAux!TFLNombreCorto = tNCorto.Text
    rsAux!TFLFormaPago = cFormaPago.ItemData(cFormaPago.ListIndex)
    rsAux!TFLRequiereAgencia = chAgencia.Value
    rsAux!TFlAgenda = douAux
    If douAgenda <> douAux Then rsAux!TFlAgendaHabilitada = douAux: bChange = True
    If IsNull(rsAux!TFlFechaAgeHab) Then rsAux!TFlFechaAgeHab = Format(Now, "mm/dd/yyyy hh:mm:ss")
    If cRangoHora.ListIndex > -1 Then
        rsAux!TFLRangoHS = cRangoHora.ItemData(cRangoHora.ListIndex)
        iRangoHS = cRangoHora.ItemData(cRangoHora.ListIndex)
    Else
        rsAux!TFLRangoHS = Null
        iRangoHS = -1
    End If
    rsAux!TFLHoraEnvio = IIf(sHE = "", Null, sHE)
    sHoraEnvio = sHE
    rsAux!TFLArticulo = tArticulo.prm_ArtID
    If IsNumeric(tTipoPrecio.Text) Then rsAux("TFLTipoPrecioFlete") = tTipoPrecio.Text
    If IsNumeric(tCobra.Text) Then rsAux("TFLCobra") = tCobra.Text
    rsAux.Update
    rsAux.Close
    douAgenda = douAux
    
End Sub
Private Function db_SaveGrupoZona(ByVal douAge As Double, ByVal sHE As String, ByRef bChange As Boolean) As Boolean
Dim iQ As Integer, iSG As Integer
Dim bCh As Boolean
    db_SaveGrupoZona = False
    On Error GoTo errBT
    cBase.BeginTrans
    On Error GoTo errRB
    
    With vsSubGrupo
        iSG = .Cell(flexcpValue, .Row, 0)
        For iQ = .FixedRows To .Rows - 1
            If .Cell(flexcpValue, iQ, 0) = iSG Then
                db_SaveAgendaZona .Cell(flexcpData, iQ, 1), douAge, sHE, bCh, True
                If bCh And Not bChange Then bChange = True
            End If
        Next
    End With
    cBase.CommitTrans
    db_SaveGrupoZona = True
    Exit Function
    
errBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Err.Description
    Exit Function
errResumo:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar grabar la información.", Err.Description
    Exit Function
errRB:
    Resume errRB
End Function

Private Sub db_AsignoSubGrupo(ByVal lCodZona As Long, ByVal SubGrupo As String)
    Dim rsAux As rdoResultset
    Dim cons As String
    cons = "Select * From FleteAgendaZona " & _
                " Where FAZZona = " & lCodZona & " And FAZTipoFlete = " & Val(tDescripcion.Tag)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        rsAux.Edit
        If (SubGrupo <> "") Then
            If (SubGrupo <> "NOTOCARSUBGRUPO") Then rsAux("FAZSubGrupo") = SubGrupo
        Else
            rsAux("FAZSubGrupo") = Null
        End If
        rsAux.Update
    End If
    rsAux.Close
End Sub

Private Sub db_SaveAgendaZona(ByVal lCodZona As Long, ByVal douAgen As Double, ByVal sHE As String, ByRef bChange As Boolean, ByVal esSubGrupo As Boolean)
Dim cons As String
Dim rsAux As rdoResultset
Dim bNew As Boolean
    
    bChange = False
    cons = "Select * From FleteAgendaZona " & _
        " Where FAZZona = " & lCodZona & " And FAZTipoFlete = " & Val(tDescripcion.Tag)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        rsAux.Edit
        bNew = False
    Else
        bNew = True
        rsAux.AddNew
        rsAux!FAZZona = lCodZona
        rsAux!FAZTipoFlete = Val(tDescripcion.Tag)
        bChange = True
    End If
    If douAgen > 0 Then
        If Not bNew Then
            If rsAux!FAZAgenda <> douAgen Then
                bChange = True
                If Not esSubGrupo Then rsAux("FAZSubGrupo") = Null
            End If
        End If
        rsAux!FAZAgenda = douAgen
        If bChange Then rsAux!FAZAgendaHabilitada = douAgen
        If IsNull(rsAux!FAZFechaAgeHab) Then rsAux!FAZFechaAgeHab = Format(Now, "mm/dd/yyyy hh:mm:ss")
    Else
        rsAux!FAZAgenda = douAgen
        rsAux!FAZAgendaHabilitada = Null
        rsAux!FAZFechaAgeHab = Null
    End If
    If cRangoHora.ListIndex > -1 Then
        rsAux!FAZRangoHS = cRangoHora.ItemData(cRangoHora.ListIndex)
    Else
        rsAux!FAZRangoHS = Null
    End If
    If Trim(sHE) = "" Then
        rsAux!FAZHoraEnvio = Null
    Else
        rsAux!FAZHoraEnvio = sHE
    End If
    rsAux.Update
    rsAux.Close
    
End Sub

Private Sub tsAgenda_Click()
    
    With tZonaGrupoZona
        .Enabled = (tsAgenda.SelectedItem.Index = 2)
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
        .Visible = (tsAgenda.SelectedItem.Index <> 3)
    End With
    With tZona
        .Enabled = (tsAgenda.SelectedItem.Index = 3)
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
        .Visible = (tsAgenda.SelectedItem.Index = 3)
    End With
    With cSubGrupo
        .Enabled = tsAgenda.SelectedItem.Index = 2
        .Tag = "": .Clear
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    vsSubGrupo.Rows = 1
    vsSubGrupo.Enabled = cSubGrupo.Enabled
    
    lEs.Caption = IIf(tsAgenda.SelectedItem.Index = 3, "&Zona:", "&Grupo Zona:")
    
    If Val(tDescripcion.Tag) > 0 And tsAgenda.Enabled Then
        ArmoAgendaFlete
    End If
    
End Sub

Private Sub tTipoPrecio_GotFocus()
    With tTipoPrecio
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tTipoPrecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tCobra.SetFocus
End Sub

Private Sub tZona_Change()
    If Val(tZona.Tag) > 0 Then
        tZona.Tag = ""
        LimpioDiasHabilitados
    End If
End Sub

Private Sub tZona_GotFocus()
    With tZona
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tZona_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And Trim(tZona.Text) <> "" Then
        If Val(tZona.Tag) = 0 Then
            loc_FindByText tZona, "Select ZonCodigo, ZonNombre as 'Nombre' From Zona Where ZonNombre like '" & Replace(tZona.Text, " ", "%") & "%' Order by ZonNombre"
            If Val(tZona.Tag) > 0 Then ArmoAgendaFlete
        End If
    End If
    
End Sub

Private Sub tZonaGrupoZona_Change()
    If Val(tZonaGrupoZona.Tag) > 0 Then
        tZonaGrupoZona.Tag = ""
        LimpioDiasHabilitados
        If cSubGrupo.Enabled Then cSubGrupo.Clear: cSubGrupo.Clear: vsSubGrupo.Rows = 1
    End If
End Sub

Private Sub tZonaGrupoZona_GotFocus()
    With tZonaGrupoZona
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tZonaGrupoZona_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tZonaGrupoZona.Text) <> "" Then
        If Val(tZonaGrupoZona.Tag) > 0 Then
            If cSubGrupo.Enabled Then cSubGrupo.SetFocus
        Else
            loc_FindByText tZonaGrupoZona, "Select GZoCodigo, GZoNombre as 'Nombre' From GrupoZona Where GZoNombre like '" & Replace(tZonaGrupoZona.Text, " ", "%") & "%' Order by GZoNombre"
            If Val(tZonaGrupoZona.Tag) > 0 Then ArmoAgendaFlete
        End If
    End If
End Sub

Private Sub vsAgenda_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    If Val(vsAgenda.Cell(flexcpData, Row, Col)) = 0 Then Cancel = True
End Sub

Private Sub ArmoAgendaGrupoZona()
Dim cons As String
Dim rsAux As rdoResultset
Dim sZonaIn As String
Dim douAgendaAux As Double
Dim lColor As Long, lAux As Long
Dim iSG As Integer
    
    With cSubGrupo
        .Clear
        .AddItem "Sub Grupo 0"
        .ItemData(.NewIndex) = 0
    End With
    vsSubGrupo.Rows = 1
    If Val(tZonaGrupoZona.Tag) = 0 Then Exit Sub
    iSG = -1
    
    'Paso 1 Cargo todos los subgrupos que pueda formar en base a la tabla FleteAgendaZona
    'Paso 2 Cargo el resto de las zonas y las asigno como GrupoPcpal al primero.
    sZonaIn = "0"
    douAgendaAux = -1
    lColor = &HEEFFFF
    
    Dim sSubGrupo As String
    sSubGrupo = String(20, "A")
    
    'Si la agenda es nula se considera la agenda de la pcpal.
    cons = "Select FAZZona, ZonNombre, IsNull(FAZAgenda, 0)  as Agenda, FAZRangoHs as RangoHS, IsNull(FAZHoraEnvio, ' ') as HoraE, ISNULL(FAZSubGrupo, '') as SubGrupo From FleteAgendaZona, Zona " & _
        " Where FAZTipoFlete = " & Val(tDescripcion.Tag) & _
        " And FAZZona IN (Select GZZZona From GrupoZonaZona Where GZZGrupo = " & Val(tZonaGrupoZona.Tag) & ")" & _
        " And FAZZona = ZonCodigo Order By FAZSubGrupo, FAZAgenda"
        
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        
        'If douAgendaAux <> rsAux!Agenda Then
        If douAgendaAux <> rsAux!Agenda Or (rsAux("SubGrupo") <> "" And sSubGrupo <> rsAux("SubGrupo")) Then
            lColor = IIf(lColor = &HEEFFFF, vbWhite, &HEEFFFF)
            If rsAux!Agenda > 0 Then iSG = iSG + 1
            If iSG > 0 Then
                With cSubGrupo
                    .AddItem IIf(rsAux("SubGrupo") <> "", rsAux("SubGrupo"), "Sub Grupo " & iSG)
                    '.AddItem "Sub Grupo " & iSG
                    .ItemData(.NewIndex) = iSG
                End With
            End If
            douAgendaAux = rsAux!Agenda
            If rsAux("SubGrupo") <> "" Then sSubGrupo = rsAux("SubGrupo")
        End If
        With vsSubGrupo
            '.AddItem iSG
            If rsAux("SubGrupo") <> "" Then
                .AddItem sSubGrupo
            Else
                .AddItem iSG
            End If
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!ZonNombre)
            douAgendaAux = rsAux!Agenda
            .Cell(flexcpData, .Rows - 1, 0) = douAgendaAux
            lAux = rsAux!FAZZona
            .Cell(flexcpData, .Rows - 1, 1) = lAux
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = lColor
            .Cell(flexcpText, .Rows - 1, 2) = rsAux("HoraE")
        End With
        sZonaIn = sZonaIn & ", " & rsAux!FAZZona
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    If vsSubGrupo.Rows > vsSubGrupo.FixedRows Then
        If vsSubGrupo.Cell(flexcpValue, 1, 0) = 0 Then
            lColor = vsSubGrupo.Cell(flexcpBackColor, 1, 0)
        Else
            'La primera si ó si fue blanca
            lColor = &HEEFFFF
        End If
    Else
        lColor = vbWhite
    End If
        
    'Primero cargo todos los GrupoZona que tengo para el tipoflete
    cons = "Select Distinct(ZonCodigo), ZonNombre From GrupoZonaZona, Zona" & _
            " Where GZZGrupo = " & Val(tZonaGrupoZona.Tag) & _
            " And GZZZona Not IN(" & sZonaIn & ")" & _
            " And GZZZona = ZonCodigo Order by ZonNombre"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        With vsSubGrupo
            .AddItem 0, .FixedRows
            .Cell(flexcpText, .FixedRows, 1) = Trim(rsAux!ZonNombre)
            .Cell(flexcpData, .FixedRows, 0) = 0
            lAux = rsAux!ZonCodigo
            .Cell(flexcpData, .FixedRows, 1) = lAux
            .Cell(flexcpBackColor, .FixedRows, 0, , .Cols - 1) = lColor
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    With vsSubGrupo
        If .Rows > .FixedRows Then
            .Select .FixedRows, 0, .Rows - 1, .Cols - 1
            .Sort = flexSortStringAscending
            .Select .FixedRows, 0
            ArmoAgendaFleteEnGrilla vsSubGrupo.Cell(flexcpData, vsSubGrupo.Row, 0)
        End If
    End With
    
End Sub

Private Sub loc_FindByText(txtCtrl As TextBox, ByVal cons As String)
On Error GoTo errCTF
Dim rsAux As rdoResultset
Dim sNombre As String, lCodigo As Long
    
    Screen.MousePointer = 11
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then
        rsAux.Close
        MsgBox "No se encontraron datos para el filtro ingresado.", vbExclamation, "Atención"
    Else
        rsAux.MoveNext
        If Not rsAux.EOF Then
            rsAux.Close
            lCodigo = fnc_HelpList(cons, "Grupos de Zona", sNombre)
        Else
            rsAux.MoveFirst
            sNombre = Trim(rsAux(1))
            lCodigo = rsAux(0)
            rsAux.Close
        End If
    End If
    If lCodigo > 0 Then
        With txtCtrl
            .Text = sNombre
            .Tag = lCodigo
        End With
    End If
    Screen.MousePointer = 0
    Exit Sub
errCTF:
    clsGeneral.OcurrioError "Error al buscar.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function fnc_HelpList(ByVal cons As String, sTitulo As String, sRetNombre As String) As Long
    
    'Valores a retornar código y nombre
    fnc_HelpList = 0
    sRetNombre = ""
    
    Dim objLista As New clsListadeAyuda
    If objLista.ActivarAyuda(cBase, cons, 4500, 1, sTitulo) > 0 Then
        fnc_HelpList = objLista.RetornoDatoSeleccionado(0)
        sRetNombre = objLista.RetornoDatoSeleccionado(1)
    End If
    Set objLista = Nothing

End Function

Private Sub fnc_GetDatosPcpal()
Dim cons As String
Dim rsAux As rdoResultset
    On Error GoTo errCDP
    iRangoHS = -1
    douAgenda = 0
    sHoraEnvio = ""
    Screen.MousePointer = 11
    cons = "Select TFLNombreCorto, TFLArticulo, IsNull(TFLAgenda, 0) as TFLAgenda, TFLRangoHS, TFLHoraEnvio, TFLFormaPago, " & _
                    " TFLRequiereAgencia, TFLTipoPrecioFlete, TFLCobra" & _
                " From TipoFlete Where TFLCodigo = " & Val(tDescripcion.Tag)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        tNCorto.Text = Trim(rsAux!TFLNombreCorto)
        douAgenda = rsAux("TFLAgenda")
        If Not IsNull(rsAux!TFLRangoHS) Then iRangoHS = rsAux!TFLRangoHS
        If Not IsNull(rsAux!TFLHoraEnvio) Then sHoraEnvio = rsAux!TFLHoraEnvio
        chAgencia.Value = IIf(rsAux!TFLRequiereAgencia, "1", 0)
        If Not IsNull(rsAux!TFLFormaPago) Then BuscoCodigoEnCombo cFormaPago, rsAux!TFLFormaPago
        If Not IsNull(rsAux!TFLArticulo) Then tArticulo.LoadArticulo rsAux!TFLArticulo
        If Not IsNull(rsAux("TFLTipoPrecioFlete")) Then tTipoPrecio.Text = rsAux("TFLTipoPrecioFlete")
        If Not IsNull(rsAux("TFLCobra")) Then tCobra.Text = rsAux("TFLCobra")
    End If
    rsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errCDP:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los datos del flete.", Err.Description
End Sub

Private Sub frm_OcultoCtrl()
    lTFNeedAgencia.Tag = ""
    LimpioDiasHabilitados
    cRangoHora.ListIndex = -1: cRangoHora.Text = ""
    chAgencia.Value = 0
    cFormaPago.ListIndex = -1
    tCobra.Text = ""
    tNCorto.Text = ""
    tTipoPrecio.Text = ""
    tArticulo.Text = ""
    frm_CleanHoraEnvio
    miBotones False, False, False
    douAgenda = 0
    iRangoHS = -1
    tsAgenda.Tabs(1).Selected = True
    tsAgenda_Click
    tsAgenda.Enabled = False
End Sub

Private Function CalculoValorSuperposicion() As Double
On Error Resume Next
Dim Fila As Integer, Col As Integer
Dim douAux As Double
    douAux = 0
    For Fila = 1 To vsAgenda.Rows - 1
        For Col = 1 To vsAgenda.Cols - 1
            If vsAgenda.Cell(flexcpChecked, Fila, Col) = flexChecked Then
                douAux = douAux + superp_ValSuperposicion(Val(vsAgenda.Cell(flexcpData, Fila, Col)))
            End If
        Next Col
    Next Fila
    CalculoValorSuperposicion = douAux
End Function

Private Sub vsHoraEnvio_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsHoraEnvio_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim sHE As String
    sHE = vsHoraEnvio.EditText
    Cancel = Not fnc_ValidoRangoHorario(sHE)
    If Not Cancel Then vsHoraEnvio.EditText = sHE
End Sub

Private Sub vsSubGrupo_DblClick()
Dim lZona As Long, sNombre As String

    If vsSubGrupo.Row = 0 Or Not tsAgenda.Enabled Then Exit Sub
    
    lZona = vsSubGrupo.Cell(flexcpData, vsSubGrupo.Row, 1)
    sNombre = vsSubGrupo.Cell(flexcpText, vsSubGrupo.Row, 1)
    
    'Considero edición de zona.
    tsAgenda.Tabs(3).Selected = True
    With tZona
        .Text = sNombre: .Tag = lZona
    End With
    ArmoAgendaFlete
    AccionModificar
    
End Sub

Private Sub vsSubGrupo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = vbRightButton Then PopupMenu MnuOptSubGrupo
End Sub

Private Sub vsSubGrupo_RowColChange()
Dim iRH As Integer
    If vsSubGrupo.Row = 0 Then Exit Sub
    LimpioDiasHabilitados
    
    If vsSubGrupo.Cell(flexcpValue, vsSubGrupo.Row, 0) > 0 Then
        'Puedo tener rango hora
        loc_SetRangoHsHoraEnvio vsSubGrupo.Cell(flexcpData, vsSubGrupo.Row, 1)
    Else
        loc_SetHoraEnvio vsSubGrupo.Cell(flexcpText, vsSubGrupo.Row, 2)
        If iRangoHS > -1 Then BuscoCodigoEnCombo cRangoHora, CLng(iRangoHS)
    End If
    ArmoAgendaFleteEnGrilla vsSubGrupo.Cell(flexcpData, vsSubGrupo.Row, 0)
    
End Sub

Private Sub ArmoAgendaZona()
Dim douAux As Double
Dim cons As String
Dim rsAux As rdoResultset

    If Val(tZona.Tag) = 0 Then Exit Sub
    cons = "Select * From FleteAgendaZona " & _
                " Where FAZZona = " & Val(tZona.Tag) & " And FAZTipoFlete = " & Val(tDescripcion.Tag)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!FAZAgenda) Then douAux = rsAux!FAZAgenda
                
        If Not IsNull(rsAux!FAZRangoHS) Then
            BuscoCodigoEnCombo cRangoHora, rsAux!FAZRangoHS
        End If
        If Not IsNull(rsAux!FAZHoraEnvio) Then loc_SetHoraEnvio rsAux!FAZHoraEnvio
    End If
    rsAux.Close
    ArmoAgendaFleteEnGrilla douAux
    
End Sub

Private Sub loc_HideShowSG()
Dim iQ As Integer, iQSel As Integer
Dim iSel As Integer

    LimpioDiasHabilitados
    iQSel = 0
    If cSubGrupo.ListIndex = -1 Then
        iSel = -1
    Else
        If InStr(1, cSubGrupo.Text, "Sub grupo", vbTextCompare) > 0 Then
            iSel = cSubGrupo.ItemData(cSubGrupo.ListIndex)
        Else
            iSel = -2
        End If
    End If
    With vsSubGrupo
        For iQ = .FixedRows To .Rows - 1
            If iSel = -1 Then
                .RowHidden(iQ) = False
            ElseIf iSel = -2 Then
                .RowHidden(iQ) = (cSubGrupo.Text <> .Cell(flexcpText, iQ, 0))
            Else
                .RowHidden(iQ) = Not (.Cell(flexcpValue, iQ, 0) = iSel)
            End If
            If Not .RowHidden(iQ) Then iQSel = iQ
        Next
    End With
    If iQSel > 0 Then vsSubGrupo.Select iQSel, 0
End Sub

Private Sub frm_CleanHoraEnvio()
Dim iQ As Integer
    For iQ = 1 To vsHoraEnvio.Cols - 1
        vsHoraEnvio.Cell(flexcpText, 0, iQ) = ""
    Next
End Sub

Private Sub loc_SetHoraEnvio(ByVal sHora As String)
Dim arrHoraE() As String, arrID() As String
Dim iQ As Integer, iCol As Integer
    
    frm_CleanHoraEnvio
    If Trim(sHora) = "" Then Exit Sub
    arrHoraE = Split(sHora, ",")
    For iQ = 0 To UBound(arrHoraE)
        arrID = Split(arrHoraE(iQ), ":")
        If Trim(arrID(0)) <> "" Then
            For iCol = 1 To UBound(arrHE)
                If CLng(arrID(0)) = arrHE(iCol).Indice Then
                    vsHoraEnvio.Cell(flexcpText, 0, iCol) = arrID(1)
                End If
            Next
        End If
    Next
    
End Sub
Private Function fnc_ValidoGrabar() As Boolean

    fnc_ValidoGrabar = False
    If tsAgenda.SelectedItem.Index = 1 Then
    
        If Trim(tNCorto.Text) = "" Then
            MsgBox "La abreviación es obligatoria.", vbExclamation, "Atención"
            tNCorto.SetFocus
            Exit Function
        End If
        If tArticulo.prm_ArtID = 0 Then
            MsgBox "El artículo que factura el flete es obligatoria.", vbExclamation, "Atención"
            tArticulo.SetFocus
            Exit Function
        End If
        
        If cFormaPago.ListIndex = -1 Then
            MsgBox "Seleccione una forma de pago válida.", vbExclamation, "Atención"
            cFormaPago.SetFocus
            Exit Function
        End If
        If Not IsNumeric(tTipoPrecio.Text) And Trim(tTipoPrecio.Text) <> "" Then
            MsgBox "El tipo de precio a aplicar debe ser un número.", vbExclamation, "Atención"
            tTipoPrecio.SetFocus
            Exit Function
        End If
        If Not IsNumeric(tCobra.Text) And Trim(tCobra.Text) <> "" Then
            MsgBox "Para indicar si este tipo de flete cobra debe ser un número mayor que cero.", vbExclamation, "Atención"
            tCobra.SetFocus
            Exit Function
        End If
    End If
    
    If cRangoHora.ListIndex = -1 And cRangoHora.Text <> "" Then
        MsgBox "El rango de horas seleccionado no es válido.", vbExclamation, "Atención"
        cRangoHora.SetFocus
        Exit Function
    End If
    fnc_ValidoGrabar = True

End Function

Private Function fnc_ValidoRangoHorario(ByRef sRango As String) As Boolean

    fnc_ValidoRangoHorario = False
    
    sRango = Trim(sRango)
    
    If InStr(1, sRango, "-") > 0 Then
        Select Case Len(sRango)
            Case 9
                If CLng(Mid(sRango, 1, InStr(1, sRango, "-") - 1)) > CLng(Mid(sRango, InStr(1, sRango, "-") + 1, Len(sRango))) Then
                    MsgBox "El rango ingresado no es válido.", vbExclamation, "Atención"
                    Exit Function
                End If
                
            Case 5
                If InStr(1, sRango, "-") = 1 Then
                    If CLng(Mid(sRango, InStr(1, sRango, "-") + 1, Len(sRango))) < paPrimeraHoraEnvio Then
                        MsgBox "El horario ingresado es menor a la primera hora de entrega.", vbExclamation, "Atención"
                        Exit Function
                    Else
                        If paPrimeraHoraEnvio < 1000 Then
                            sRango = "0" & paPrimeraHoraEnvio & sRango
                        Else
                            sRango = paPrimeraHoraEnvio & sRango
                        End If
                        Exit Function
                    End If
                Else
                    If InStr(1, sRango, "-") = 5 Then
                        If CLng(Mid(sRango, 1, InStr(1, sRango, "-") - 1)) > paUltimaHoraEnvio Then
                            MsgBox "El horario ingresado es mayor que la última hora de envio.", vbExclamation, "ATENCIÓN"
                            Exit Function
                        Else
                            sRango = sRango & paUltimaHoraEnvio
                        End If
                    Else
                        MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "Atención"
                        Exit Function
                    End If
                End If
            
            Case 8
                If CLng(Mid(sRango, 1, InStr(1, sRango, "-") - 1)) > CLng(Mid(sRango, InStr(1, sRango, "-") + 1, Len(sRango))) Then
                    MsgBox "El rango de horario ingresado no es válido.", vbExclamation, "ATENCIÓN"
                    Exit Function
                End If
                
                If InStr(1, sRango, "-") = 4 Then
                    sRango = "0" & sRango
                End If
            
            Case Else
                    MsgBox "Horario inválido. [####-####]", vbExclamation, "ATENCIÓN"
                    Exit Function
        End Select
    Else
        If Trim(sRango) <> "" Then
            MsgBox "Horario inválido. [####-####]", vbExclamation, "ATENCIÓN"
            Exit Function
        End If
    End If
    
    fnc_ValidoRangoHorario = True
    
    'Ahora válido el rango de horas.
'    Dim lhora As Long
'    lhora = (CLng(Mid(sRango, InStr(1, sRango, "-") + 1, Len(sRango))) - CLng(Mid(sRango, 1, InStr(1, sRango, "-") - 1))) / 100
'    If lhora < arrDatosFlete(iIndex).HorarioRango Then
'        If MsgBox("El rango ingresado es menor al posible para el flete seleccionado." & vbCr & vbCr & _
                    "El flete tiene un rango de " & arrDatosFlete(iIndex).HorarioRango & " hora(s) y se asigno un rango de " & lhora & " hora(s)" & vbCr & vbCr & _
                    "¿Confirma mantener el rango ingresado?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible error en horario") = vbNo Then
'            cHoraEntrega.SetFocus
'            ValidoRangoHorario = False
'        End If
'    End If
    
End Function

Private Sub db_LoadPrm()
Dim cons As String
Dim rsAux As rdoResultset
    cons = "Select * From Parametro Where ParNombre IN('PrimeraHoraEnvio', 'UltimaHoraEnvio')"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            Case "primerahoraenvio": paPrimeraHoraEnvio = rsAux!ParValor
            Case "ultimahoraenvio": paUltimaHoraEnvio = rsAux!ParValor
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
End Sub

Private Function fnc_GetHoraEnvioGrid() As String
Dim iCol As Integer
Dim sSalida As String, bHay As Boolean
    
    fnc_GetHoraEnvioGrid = ""
    With vsHoraEnvio
        For iCol = 1 To .Cols - 1
            If Trim(.Cell(flexcpText, 0, iCol)) <> "" Then
                If sSalida <> "" Then sSalida = sSalida & ","
                bHay = True
                sSalida = sSalida & arrHE(iCol).Indice & ":" & .Cell(flexcpText, 0, iCol)
            Else
'                sSalida = sSalida & arrHE(iCol).Indice & ":" & arrHE(iCol).Hora
            End If
        Next iCol
    End With
    If bHay Then fnc_GetHoraEnvioGrid = sSalida
    
End Function

Private Sub loc_SetRangoHsHoraEnvio(ByVal lZona As Long)
Dim cons As String
Dim rsAux As rdoResultset
    
    cons = "Select FAZRangoHs as RangoHS, FAZHoraEnvio  From FleteAgendaZona " & _
        " Where FAZTipoFlete = " & Val(tDescripcion.Tag) & _
        " And FAZZona = " & lZona
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
           
        If Not IsNull(rsAux!RangoHS) Then
            BuscoCodigoEnCombo cRangoHora, rsAux!RangoHS
        Else
            If iRangoHS > -1 Then BuscoCodigoEnCombo cRangoHora, CLng(iRangoHS)
        End If
        If Not IsNull(rsAux!FAZHoraEnvio) Then
            loc_SetHoraEnvio rsAux!FAZHoraEnvio
        Else
            If sHoraEnvio <> "" Then loc_SetHoraEnvio sHoraEnvio
        End If
    Else
        loc_SetHoraEnvio sHoraEnvio
        If iRangoHS > -1 Then BuscoCodigoEnCombo cRangoHora, CLng(iRangoHS)
    End If
    rsAux.Close
            
End Sub
