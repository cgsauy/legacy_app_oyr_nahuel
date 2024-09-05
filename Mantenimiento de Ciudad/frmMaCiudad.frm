VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMaCiudad 
   Caption         =   "Ciudades"
   ClientHeight    =   4095
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaCiudad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox tDemora 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox tTelediscado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin AACombo99.AACombo cPais 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrilla 
      Height          =   1815
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3201
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
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
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
   Begin VB.TextBox tNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   4800
      MaxLength       =   30
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3840
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "(en días) para el arribo."
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Demora:"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Telediscado:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&País:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre:"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   240
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
            Picture         =   "frmMaCiudad.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCiudad.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCiudad.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCiudad.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCiudad.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCiudad.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCiudad.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCiudad.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCiudad.frx":10E2
            Key             =   ""
         EndProperty
      EndProperty
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
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuAccesos 
      Caption         =   "Accesos"
      Visible         =   0   'False
      Begin VB.Menu MnuFiltrar 
         Caption         =   "Filtrar"
      End
      Begin VB.Menu MnuQuitoFiltro 
         Caption         =   "Quitar Filtros"
      End
   End
End
Attribute VB_Name = "frmMaCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sNuevo As Boolean, sModificar As Boolean
Private Sub cPais_GotFocus()
    With cPais
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cPais_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If (sNuevo Or sModificar) Then
            Foco tNombre
        Else
            CargoGrilla
        End If
    End If
End Sub

Private Sub cPais_LostFocus()
    With cPais
        .SelStart = 0
    End With
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    DoEvents
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    Screen.MousePointer = 11
    ObtengoSeteoForm Me, 1000, 500, 3840, 4230
    sNuevo = False: sModificar = False
    Cons = "Select PaiCodigo, PaiNombre From Pais Order By PaiNombre"
    CargoCombo Cons, cPais
    CargoGrilla
    If vsGrilla.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me
    OcultoCampos
    Screen.MousePointer = 0
    Exit Sub
ErrLoad:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al iniciar el formulario.", Trim(Err.Description), "Ciudades"
End Sub
Private Sub Form_Resize()
On Error Resume Next
    vsGrilla.Height = Me.ScaleHeight - (vsGrilla.Top + 70 + Status.Height)
    vsGrilla.Width = Me.ScaleWidth - (vsGrilla.Left * 2)
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miconexion = Nothing
    End
End Sub
Private Sub Label1_Click()
    Foco tNombre
End Sub
Private Sub Label2_Click()
    Foco cPais
End Sub

Private Sub Label5_Click()
    Foco tTelediscado
End Sub

Private Sub Label6_Click()
    Foco tDemora
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub
Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuFiltrar_Click()
    CargoGrilla
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

Private Sub MnuQuitoFiltro_Click()
    cPais.ListIndex = -1
    CargoGrilla
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub
Private Sub AccionNuevo()
   
    'Prendo Señal que es uno nuevo.
    sNuevo = True
    'Habilito y Desabilito Botones
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoCampos
    Foco cPais
    
End Sub

Private Sub AccionModificar()
    vsGrilla.Enabled = False
    'Habilito y Desabilito Botones
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoCampos
    sModificar = True
    tNombre.Text = vsGrilla.Cell(flexcpText, vsGrilla.Row, 1)
    BuscoCodigoEnCombo cPais, vsGrilla.Cell(flexcpData, vsGrilla.Row, 0)
    tTelediscado.Text = Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 2))
    tDemora.Text = Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 3))
    Foco cPais
End Sub

Private Sub AccionGrabar()
    If ValidoDatos Then
        If sNuevo Then
            If MsgBox("¿Confirma el alta de datos?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then NuevoRegistro
        Else
            If MsgBox("¿Confirma modificar los datos?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then ModificoRegistro
        End If
    End If
End Sub

Private Sub AccionEliminar()
On Error GoTo ErrAE
    
    If MsgBox("¿Confirma eliminar la ciudad " & Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 1)) & " ?", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        Screen.MousePointer = 11
        Cons = "Delete Ciudad Where CiuCodigo = " & vsGrilla.Cell(flexcpData, vsGrilla.Row, 1)
        cBase.Execute (Cons)
        CargoGrilla
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrAE:
    clsGeneral.OcurrioError "Ocurrió un error al eliminar el registro.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub AccionCancelar()
    vsGrilla.Enabled = True
    sNuevo = False: sModificar = False
    CargoGrilla
End Sub
Private Sub tDemora_GotFocus()
    With tDemora
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tDemora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tDemora_LostFocus()
    With tDemora
        .SelStart = 0
    End With
End Sub

Private Sub tNombre_GotFocus()
    With tNombre
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese ."
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tNombre.Text) <> "" Then Foco tTelediscado
End Sub

Private Sub tNombre_LostFocus()
    Status.SimpleText = ""
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
Private Sub OcultoCampos()
    tNombre.BackColor = Colores.Inactivo: tNombre.Enabled = False: tNombre.Text = ""
    tDemora.BackColor = Colores.Inactivo: tDemora.Enabled = False: tDemora.Text = ""
    tTelediscado.BackColor = Colores.Inactivo: tTelediscado.Enabled = False: tTelediscado.Text = ""
    cPais.BackColor = vbWhite
    vsGrilla.Enabled = True: vsGrilla.BackColor = vbWhite
End Sub
Private Sub HabilitoCampos()
    tNombre.BackColor = Obligatorio: tNombre.Enabled = True
    tDemora.BackColor = vbWhite: tDemora.Enabled = True: tDemora.Text = ""
    tTelediscado.BackColor = vbWhite: tTelediscado.Enabled = True: tTelediscado.Text = ""
    cPais.BackColor = Obligatorio
    vsGrilla.Enabled = False: vsGrilla.BackColor = Inactivo
End Sub
Private Sub CargoGrilla()
On Error GoTo ErrCI
Dim CodAux As Integer
    
    Screen.MousePointer = 11
    LimpioGrilla
    OcultoCampos
    vsGrilla.Redraw = False
    Cons = "Select * From Ciudad, Pais Where CiuPais = PaiCodigo "
    If cPais.ListIndex > -1 Then Cons = Cons & " And CiuPais = " & cPais.ItemData(cPais.ListIndex)
    Cons = Cons & " Order by PaiNombre, CiuNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        With vsGrilla
            .AddItem ""
            CodAux = RsAux!PaiCodigo
            .Cell(flexcpData, vsGrilla.Rows - 1, 0) = CodAux
            .Cell(flexcpText, vsGrilla.Rows - 1, 0) = Trim(RsAux!PaiNombre)
            CodAux = RsAux!CiuCodigo
            .Cell(flexcpData, vsGrilla.Rows - 1, 1) = CodAux
            .Cell(flexcpText, vsGrilla.Rows - 1, 1) = Trim(RsAux!CiuNombre)
            If Not IsNull(RsAux!CiuTelediscado) Then .Cell(flexcpText, vsGrilla.Rows - 1, 2) = Trim(RsAux!CiuTelediscado)
            If Not IsNull(RsAux!CiuDemora) Then .Cell(flexcpText, vsGrilla.Rows - 1, 3) = Trim(RsAux!CiuDemora)
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    vsGrilla.Redraw = True
    If vsGrilla.Rows > 1 Then
        vsGrilla.Select 1, 0, 1, 1
        Botones True, True, True, False, False, Toolbar1, Me
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrCI:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos en la grilla.", Trim(Err.Description)
    Screen.MousePointer = 0
    vsGrilla.Redraw = True
End Sub

Private Sub LimpioGrilla()
    With vsGrilla
        .Redraw = False
        .ExtendLastCol = True
        .Clear
        .Rows = 1
        .Cols = 4
        .FormatString = "Pais|Ciudad|Telediscado|Demora|"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1600
        .ColWidth(2) = 1900
        .ColWidth(3) = 1100
        .AllowUserResizing = flexResizeColumns
        .Redraw = True
    End With
End Sub

Private Sub tTelediscado_GotFocus()
    With tTelediscado
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tTelediscado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDemora
End Sub

Private Sub tTelediscado_LostFocus()
    With tTelediscado
        .SelStart = 0
    End With
End Sub

Private Sub vsgrilla_Click()
    If vsGrilla.MouseRow = 0 Then
        vsGrilla.ColSel = vsGrilla.MouseCol
        If vsGrilla.ColSort(vsGrilla.MouseCol) = flexSortGenericAscending Then
            vsGrilla.ColSort(vsGrilla.MouseCol) = flexSortGenericDescending
        Else
            vsGrilla.ColSort(vsGrilla.MouseCol) = flexSortGenericAscending
        End If
        vsGrilla.Sort = flexSortUseColSort
    End If
End Sub
Private Sub NuevoRegistro()
On Error GoTo ErrNI
    Screen.MousePointer = 11
    
    Cons = "Select * From Ciudad Where CiuNombre = '" & Trim(tNombre.Text) & "'" _
        & " And CiuPais = " & cPais.ItemData(cPais.ListIndex)
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "Ya existe una con esas características, verifique.", vbExclamation, "ATENCIÓN"
    Else
        RsAux.AddNew
        CargoCamposBD
        RsAux.Update
        RsAux.Close
    End If
    CargoGrilla
    vsGrilla.Enabled = True
    sNuevo = False
    Screen.MousePointer = 0
    Exit Sub
ErrNI:
    clsGeneral.OcurrioError "Ocurrió un error al dar el alta.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub CargoCamposBD()
    RsAux!CiuPais = cPais.ItemData(cPais.ListIndex)
    RsAux!CiuNombre = Trim(tNombre.Text)
    If Trim(tTelediscado.Text) <> "" Then RsAux!CiuTelediscado = Trim(tTelediscado.Text) Else RsAux!CiuTelediscado = Null
    If Trim(tDemora.Text) <> "" Then RsAux!CiuDemora = Trim(tDemora.Text) Else RsAux!CiuDemora = Null
End Sub
Private Sub ModificoRegistro()
On Error GoTo ErrNI
    
    Screen.MousePointer = 11
    Cons = "Select * From Ciudad " _
        & " Where CiuCodigo <> " & vsGrilla.Cell(flexcpData, vsGrilla.Row, 1) _
        & " And CiuNombre = '" & Trim(tNombre.Text) & "'" _
        & " And CiuPais = " & cPais.ItemData(cPais.ListIndex)
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "Ya existe una ciudad con esas características, verifique.", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else
        RsAux.Close
    End If
    Cons = "Select * From Ciudad " _
        & " Where CiuCodigo = " & vsGrilla.Cell(flexcpData, vsGrilla.Row, 1)
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No se encontró la ciudad seleccionada, verifique si no fue eliminada.", vbExclamation, "ATENCIÓN"
    Else
        RsAux.Edit
        CargoCamposBD
        RsAux.Update
        RsAux.Close
    End If
    CargoGrilla
    vsGrilla.Enabled = True
    sModificar = False
    Screen.MousePointer = 0
    Exit Sub
    
ErrNI:
    clsGeneral.OcurrioError "Ocurrió un error al modificar el registro.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    If cPais.ListIndex = -1 Then
        MsgBox "Debe seleccionar un pais.", vbExclamation, "ATENCIÓN"
        Foco cPais: Exit Function
    End If
    If Trim(tNombre.Text) = "" Then
        MsgBox "Debe ingresar un nombre.", vbExclamation, "ATENCIÓN"
        Foco tNombre: Exit Function
    End If
    If Trim(tDemora.Text) <> "" Then
        If Not IsNumeric(tDemora.Text) Then
            MsgBox "Se ingresó un formato no válido para la demora.", vbExclamation, "ATENCIÓN"
            Foco tDemora: Exit Function
        End If
    End If
    ValidoDatos = True
End Function

Private Sub vsGrilla_DblClick()
    If Toolbar1.Buttons("modificar").Enabled And vsGrilla.Row >= vsGrilla.FixedRows Then
        AccionModificar
    End If
End Sub

Private Sub vsGrilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace And Toolbar1.Buttons("modificar").Enabled And vsGrilla.Row >= vsGrilla.FixedRows Then
        AccionModificar
    End If
End Sub

Private Sub vsGrilla_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And vsGrilla.Enabled = True Then PopupMenu MnuAccesos
End Sub
