VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{191D08B9-4E92-4372-BF17-417911F14390}#1.5#0"; "orGridPreview.ocx"
Begin VB.Form SdoDisponibilidad 
   Caption         =   "Saldo Inicial de Disponibilidades"
   ClientHeight    =   5205
   ClientLeft      =   3660
   ClientTop       =   3420
   ClientWidth     =   7125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SdoDisponibilidad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7125
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin ComctlLib.StatusBar Status1 
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1500
         Visible         =   0   'False
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   450
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   4
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "terminal"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               AutoSize        =   2
               Key             =   "usuario"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               AutoSize        =   2
               Key             =   "bd"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               AutoSize        =   1
               Object.Width           =   8440
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton bImprimir 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   5880
      TabIndex        =   16
      Top             =   1650
      Width           =   1155
   End
   Begin orGridPreview.GridPreview orPrev 
      Left            =   5100
      Top             =   420
      _ExtentX        =   873
      _ExtentY        =   873
      PageBorder      =   3
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
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1140
      MaxLength       =   2
      TabIndex        =   11
      Top             =   1620
      Width           =   615
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrilla 
      Height          =   2835
      Left            =   60
      TabIndex        =   12
      Top             =   1980
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5001
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
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
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
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1140
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1320
      Width           =   5895
   End
   Begin VB.TextBox tSaldo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   5520
      TabIndex        =   7
      Top             =   1020
      Width           =   1455
   End
   Begin VB.TextBox tHora 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   3180
      TabIndex        =   5
      Top             =   1020
      Width           =   1455
   End
   Begin VB.TextBox tFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1140
      TabIndex        =   3
      Top             =   1020
      Width           =   1455
   End
   Begin AACombo99.AACombo cDisponibilidad 
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   540
      Width           =   3615
      _ExtentX        =   6376
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
      Text            =   ""
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   4950
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4366
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   1620
      Width           =   735
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3540
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SdoDisponibilidad.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SdoDisponibilidad.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SdoDisponibilidad.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SdoDisponibilidad.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SdoDisponibilidad.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SdoDisponibilidad.frx":099C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentario:"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Saldo:"
      Height          =   255
      Left            =   4860
      TabIndex        =   6
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Hora:"
      Height          =   255
      Left            =   2700
      TabIndex        =   4
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Disponibilidad:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   1215
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalirDelFormulario 
      Caption         =   "Salir"
      Begin VB.Menu MnuSalir 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "SdoDisponibilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sNuevo As Boolean, sModificar As Boolean

Private Sub bImprimir_Click()
    AccionImprimir
    
End Sub

Private Sub cDisponibilidad_Click()
    LimpioGrilla
    If cDisponibilidad.ListIndex > -1 Then CargoGrilla
End Sub
Private Sub cDisponibilidad_Change()
    LimpioGrilla
    Botones False, False, False, False, False, Toolbar1, Me
End Sub
Private Sub cDisponibilidad_GotFocus()
    With cDisponibilidad
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub cDisponibilidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tFecha
End Sub
Private Sub cDisponibilidad_LostFocus()
    cDisponibilidad.SelStart = 0
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    ObtengoSeteoForm Me, 1000, 650, 7455, 5655
    sNuevo = False: sModificar = False
    cons = "Select DisID, DisNombre From Disponibilidad Order by DisNombre"
    CargoCombo cons, cDisponibilidad
    OcultoCampos
    LimpioGrilla
    Botones False, False, False, False, False, Toolbar1, Me
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrio un error al cargar el formulario.", Err.Description
End Sub

Private Sub Form_Resize()
On Error Resume Next
    vsGrilla.Width = Me.ScaleWidth - (vsGrilla.Left * 2)
    vsGrilla.Height = Me.ScaleHeight - (vsGrilla.Top + Status.Height + 70)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
End Sub

Private Sub Label1_Click()
    Foco cDisponibilidad
End Sub
Private Sub Label2_Click()
    Foco tFecha
End Sub
Private Sub Label3_Click()
    Foco tHora
End Sub
Private Sub Label4_Click()
    Foco tSaldo
End Sub
Private Sub Label5_Click()
    Foco tComentario
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub
Private Sub tComentario_LostFocus()
    tComentario.SelStart = 0
End Sub
Private Sub tFecha_GotFocus()
    With tFecha
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Foco tHora
    ElseIf KeyAscii = vbKeyF1 Then
    End If
End Sub
Private Sub tFecha_LostFocus()
    tFecha.SelStart = 0
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, FormatoFP) Else tFecha.Text = ""
End Sub

Private Sub tHora_GotFocus()
    With tHora
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tSaldo
End Sub
Private Sub tHora_LostFocus()
    tHora.SelStart = 0
    If IsDate(tHora.Text) Then tHora.Text = Format(tHora.Text, "hh:mm:ss") Else tHora.Text = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "nuevo"
            AccionNuevo
        Case "modificar"
            AccionModificar
        Case "eliminar"
            AccionEliminar
        Case "grabar"
            AccionGrabar
        Case "cancelar"
            AccionCancelar
        Case "salir"
            Unload Me
    End Select
End Sub

Private Sub tSaldo_GotFocus()
    With tSaldo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tSaldo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentario
End Sub
Private Sub tSaldo_LostFocus()
    tSaldo.SelStart = 0
    If IsNumeric(tSaldo.Text) Then tSaldo.Text = Format(tSaldo.Text, FormatoMonedaP) Else tSaldo.Text = ""
End Sub
Private Sub LimpioGrilla()

    With vsGrilla
        .Redraw = False
        .ExtendLastCol = True
        .Clear
        .Editable = False
        .Rows = 1
        .FixedCols = 0
        .Cols = 6
        .FormatString = "Fecha|Hora|Saldo|Comentario|Usuario|"
        .ColWidth(0) = 1000
        .ColWidth(1) = 800
        .ColWidth(2) = 1100
        .ColWidth(3) = 3300
        .ColWidth(4) = 650
        .AllowUserResizing = flexResizeColumns
        .HighLight = flexHighlightAlways
        .Redraw = True
    End With
End Sub

Private Sub OcultoCampos()
    tFecha.Text = "": tFecha.Enabled = False: tFecha.BackColor = Inactivo
    tHora.Text = "": tHora.Enabled = False: tHora.BackColor = Inactivo
    tSaldo.Text = "": tSaldo.Enabled = False: tSaldo.BackColor = Inactivo
    tComentario.Text = "": tComentario.Enabled = False: tComentario.BackColor = Inactivo
    tUsuario.Text = "": tUsuario.Tag = "": tUsuario.Enabled = False: tUsuario.BackColor = Inactivo
End Sub
Private Sub HabilitoCampos()
    tFecha.Text = "": tFecha.Enabled = True: tFecha.BackColor = Obligatorio
    tHora.Text = "": tHora.Enabled = True: tHora.BackColor = Obligatorio
    tSaldo.Text = "": tSaldo.Enabled = True: tSaldo.BackColor = Obligatorio
    tComentario.Text = "": tComentario.Enabled = True: tComentario.BackColor = vbWhite
    tUsuario.Text = "": tUsuario.Tag = "": tUsuario.Enabled = True: tUsuario.BackColor = Obligatorio
End Sub

Private Sub AccionNuevo()
    If cDisponibilidad.ListIndex = -1 Then Botones False, False, False, False, False, Toolbar1, Me: Exit Sub
    Screen.MousePointer = 11
    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoCampos
    cDisponibilidad.Enabled = False
    Foco tFecha
    Screen.MousePointer = 0
End Sub
Private Sub AccionModificar()
    If cDisponibilidad.ListIndex = -1 Then Botones False, False, False, False, False, Toolbar1, Me: Exit Sub
    Screen.MousePointer = 11
    sModificar = True
    HabilitoCampos
    cDisponibilidad.Enabled = False
    Botones False, False, False, True, True, Toolbar1, Me
    tFecha.Text = Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0))
    tHora.Text = Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 1))
    tSaldo.Text = Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 2))
    tComentario.Text = Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 3))
    Foco tFecha
    Screen.MousePointer = 0
End Sub
Private Sub AccionEliminar()
On Error GoTo ErrAE
    If MsgBox("¿Confirma eliminar el saldo para la fecha " & Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0)) & " " & Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 1)) & "?", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Exit Sub
    Screen.MousePointer = 0
    cons = "Select * From SaldoDisponibilidad Where SDiCodigo = " & vsGrilla.Cell(flexcpData, vsGrilla.Row, 0)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then
        rsAux.Close
        MsgBox "Otra terminal pudo eliminar el saldo seleccionado, verifique.", vbExclamation, "ATENCIÓN"
    Else
        If CDate(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0)) = Format(rsAux!SDiFecha, FormatoFP) And Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 1)) = Trim(rsAux!SDiHora) _
            And CCur(vsGrilla.Cell(flexcpText, vsGrilla.Row, 2)) = rsAux!SDiSaldo Then
            rsAux.Delete
        Else
            MsgBox "Otra terminal modificó los datos del saldo, verifique.", vbExclamation, "ATENCIÓN"
        End If
        rsAux.Close
    End If
    CargoGrilla
    Screen.MousePointer = 0
    Exit Sub
ErrAE:
    clsGeneral.OcurrioError "Ocurrio un error al intentar eliminar el saldo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub AccionCancelar()
    sNuevo = False: sModificar = False
    CargoGrilla
    OcultoCampos
    cDisponibilidad.Enabled = True
End Sub
Private Sub AccionGrabar()
    If ValidoDatos Then
        If MsgBox("¿Confirma almacenar los datos ingresados?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
            If sNuevo Then GraboNuevo Else GraboModificacion
        End If
    End If
End Sub
Private Sub GraboNuevo()
On Error GoTo ErrGN
    Screen.MousePointer = 11
    cons = "Select * From SaldoDisponibilidad Where SDiDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
        & " And SDiFecha = '" & Format(tFecha.Text, sqlFormatoF) & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsAux.AddNew
    CargoCamposBD
    rsAux.Update
    rsAux.Close
    sNuevo = False
    OcultoCampos
    CargoGrilla
    cDisponibilidad.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
ErrGN:
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar el nuevo saldo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub CargoCamposBD()
    rsAux!SDiDisponibilidad = cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    rsAux!SDiFecha = Format(tFecha.Text & " " & Format(tHora.Text, "hh:mm:ss"), sqlFormatoFH)
    rsAux!SDiHora = Format(tHora.Text, "hh:mm:ss")
    rsAux!SDiSaldo = CCur(tSaldo.Text)
    If Trim(tComentario.Text) = "" Then
        rsAux!SDiComentario = Null
    Else
        rsAux!SDiComentario = Trim(tComentario.Text)
    End If
    rsAux!SDiUsuario = tUsuario.Tag
End Sub
Private Sub GraboModificacion()
On Error GoTo ErrGM
Screen.MousePointer = 0
    
    cons = "Select * From SaldoDisponibilidad Where SDiCodigo = " & vsGrilla.Cell(flexcpData, vsGrilla.Row, 0)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then
        rsAux.Close
        MsgBox "Otra terminal pudo eliminar el saldo seleccionado, verifique.", vbExclamation, "ATENCIÓN"
    Else
        If CDate(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0)) = Format(rsAux!SDiFecha, FormatoFP) And Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 1)) = Trim(rsAux!SDiHora) _
            And CCur(vsGrilla.Cell(flexcpText, vsGrilla.Row, 2)) = rsAux!SDiSaldo Then
            rsAux.Edit
            CargoCamposBD
            rsAux.Update
        Else
            MsgBox "Otra terminal modificó los datos del saldo, verifique.", vbExclamation, "ATENCIÓN"
        End If
        rsAux.Close
    End If
    sModificar = False
    OcultoCampos
    CargoGrilla
    cDisponibilidad.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
ErrGM:
    clsGeneral.OcurrioError "Ocurrio un error al intentar modificar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub CargoGrilla()
On Error GoTo ErrCG
Dim CodAux As Long

    Screen.MousePointer = 11
    LimpioGrilla
    If cDisponibilidad.ListIndex = -1 Then Screen.MousePointer = 0: Botones False, False, False, False, False, Toolbar1, Me: Exit Sub
    
    cons = "Select TOP 60 * From SaldoDisponibilidad, Usuario " _
            & " Where SDiDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
            & " And SDiUsuario = UsuCodigo " _
            & " Order by SDiFecha Desc, SDiHora Desc"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not rsAux.EOF
        CodAux = rsAux!SDiCodigo
        With vsGrilla
            .AddItem ""
            .Cell(flexcpData, .Rows - 1, 0) = CodAux
            .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!SDiFecha, "dd/mm/yyyy")
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!SDiHora)
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!SDiSaldo, FormatoMonedaP)
            If Not IsNull(rsAux!SDiComentario) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(rsAux!SDiComentario)
            .Cell(flexcpText, .Rows - 1, 4) = Trim(rsAux!UsuInicial)
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    If vsGrilla.Rows > 1 Then
        vsGrilla.Select 1, 0, 1, 3
        Botones True, True, True, False, False, Toolbar1, Me
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrCG:
    clsGeneral.OcurrioError "Ocurrio un error al cargar la grilla.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValidoDatos() As Boolean
    If cDisponibilidad.ListIndex = -1 Then
        MsgBox "No hay selecciondada una disponbilidad válida.", vbExclamation, "ATENCIÓN"
        ValidoDatos = False: AccionCancelar: Exit Function
    End If
    If Not IsDate(tFecha.Text) Then
        MsgBox "No se ingresó una fecha válida.", vbExclamation, "ATENCIÓN"
        Foco tFecha: ValidoDatos = False: AccionCancelar: Exit Function
    End If
    If Not IsDate(tHora.Text) Then
        MsgBox "No se ingresó una hora válida.", vbExclamation, "ATENCIÓN"
        Foco tHora: ValidoDatos = False: AccionCancelar: Exit Function
    End If
    If Not IsNumeric(tSaldo.Text) Then
        MsgBox "No se ingresó un saldo válido.", vbExclamation, "ATENCIÓN"
        Foco tSaldo: ValidoDatos = False: AccionCancelar: Exit Function
    End If
    If Not clsGeneral.TextoValido(tComentario.Text) Then
        MsgBox "Se ingresó por lo menos una comilla simple.", vbExclamation, "ATENCIÓN"
        Foco tComentario: ValidoDatos = False: AccionCancelar: Exit Function
    End If
    If Not IsNumeric(tUsuario.Tag) Then
        MsgBox "No ingresó un usuario válido.", vbExclamation, "ATENCIÓN"
        Foco tUsuario: ValidoDatos = False: AccionCancelar: Exit Function
    End If
    ValidoDatos = True
End Function

Private Sub tUsuario_GotFocus()
    With tUsuario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsNumeric(tUsuario.Text) Then
        cons = "Select * From Usuario Where UsuDigito = " & tUsuario.Text
        Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not rsAux.EOF Then
            tUsuario.Tag = rsAux!UsuCodigo
        Else
            tUsuario.Tag = ""
        End If
        rsAux.Close
        If IsNumeric(tUsuario.Tag) Then AccionGrabar
    End If
End Sub

Private Sub tUsuario_LostFocus()
    tUsuario.SelStart = 0
End Sub

Private Sub AccionImprimir()

    On Error GoTo errPrint
    With orPrev
        .Caption = "Saldos de Disponibilidades"
        .Header = "Saldos de Disponibilidades"
        .FileName = "Saldos de Disponibilidades"
        
        .LineBeforeGrid cDisponibilidad.Text
        .AddGrid vsGrilla.hwnd
        
        .ShowPreview
        
    End With
    
    Exit Sub

errPrint:
    clsGeneral.OcurrioError "Error al impirmir. " & Trim(Err.Description)
    Screen.MousePointer = 0

End Sub
