VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmMotivo 
   Caption         =   "Motivos de Servicio"
   ClientHeight    =   5805
   ClientLeft      =   2985
   ClientTop       =   2460
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMotivo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   6735
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   250
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
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   2
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
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton bConsultar 
      Height          =   310
      Left            =   6000
      Picture         =   "frmMotivo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ejecutar."
      Top             =   1680
      Width           =   310
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   3195
      Left            =   120
      TabIndex        =   7
      Top             =   2100
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5636
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
   Begin VB.Frame frmDatos 
      Caption         =   "Datos"
      ForeColor       =   &H00000080&
      Height          =   1155
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   6495
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   300
         Width           =   4935
         _ExtentX        =   8705
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
      Begin VB.TextBox tNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   900
         MaxLength       =   35
         TabIndex        =   3
         Top             =   660
         Width           =   4935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   5550
      Width           =   6735
      _ExtentX        =   11880
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
            Object.Width           =   3651
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cConsulta 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   4395
      _ExtentX        =   7752
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cosulta &por Tipo:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1740
      Width           =   1455
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   0
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
            Picture         =   "frmMotivo.frx":0744
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMotivo.frx":0856
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMotivo.frx":0968
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMotivo.frx":0A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMotivo.frx":0B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMotivo.frx":0EA6
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
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMotivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmIDTipoNew As Long

Dim sNuevo As Boolean, sModificar As Boolean
Dim gIdMotivo As Long


Private Sub bConsultar_Click()
    CargoLista
End Sub

Private Sub cConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub cTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tNombre
End Sub

Private Sub Form_Activate()
    If prmIDTipoNew > 0 Then
        AccionNuevo
        BuscoCodigoEnCombo cTipo, prmIDTipoNew
        If cTipo.ListIndex > -1 Then tNombre.SetFocus
    End If
    prmIDTipoNew = 0
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    ObtengoSeteoForm Me
    sNuevo = False: sModificar = False
    
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
    
    Cons = "Select * from Tipo Order by TipNombre"
    CargoCombo Cons, cTipo
    CargoCombo Cons, cConsulta
    
    LimpioFicha
    InicializoGrilla
    DeshabilitoIngreso
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    frmDatos.Width = Me.ScaleWidth - (frmDatos.Left * 2)
    
    vsLista.Width = frmDatos.Width
    vsLista.Height = Me.ScaleHeight - vsLista.Top - Status.Height - 50
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    
    GuardoSeteoForm Me
    End
    
End Sub

Private Sub Label1_Click()
    Foco tNombre
End Sub

Private Sub Label2_Click()
    Foco cTipo
End Sub

Private Sub Label4_Click()
    Foco cConsulta
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

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Sub AccionNuevo()
   
    sNuevo = True
    gIdMotivo = 0
    Botones False, False, False, True, True, Toolbar1, Me
    
    LimpioFicha
    HabilitoIngreso
    Foco cTipo
  
End Sub

Sub AccionModificar()
    
    On Error Resume Next
    sModificar = True
    
    With vsLista
        gIdMotivo = .Cell(flexcpData, .Row, 0)
        tNombre.Text = Trim(.Cell(flexcpText, .Row, 0))
        BuscoCodigoEnCombo cTipo, .Cell(flexcpData, .Row, 1)
    End With
    
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    Foco cTipo
        
End Sub

Sub AccionGrabar()

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo errGrabar
    
    If sNuevo Then
        Cons = "Select * from MotivoServicio Where MSeId = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.AddNew
        CargoCamposBD
        RsAux.Update: RsAux.Close
        
    Else                                    'Modificar----
    
        On Error GoTo errGrabar
        
        Cons = "Select * from MotivoServicio Where MSeID = " & gIdMotivo
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        CargoCamposBD
        RsAux.Update: RsAux.Close
        
        CargoLista
        
    End If
    
    gIdMotivo = 0
    sNuevo = False: sModificar = False
    DeshabilitoIngreso
    LimpioFicha
    If vsLista.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me Else Botones True, False, False, False, False, Toolbar1, Me
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
    Screen.MousePointer = 0
End Sub

Sub AccionEliminar()

    Screen.MousePointer = 11
    
    With vsLista
        If MsgBox("Confirma eliminar el motivo '" & .Cell(flexcpText, .Row, 0) & "'", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Screen.MousePointer = 0: Exit Sub
        On Error GoTo Error
        
        Cons = "SELECT SReMotivo FROM ServicioRenglon WHERE SReTipoRenglon = 1 AND SReMotivo = " & .Cell(flexcpData, .Row, 0)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not RsAux.EOF Then
            MsgBox "No puede eliminar el motivo ya que existen servicios que lo tienen asociado.", vbExclamation, "ATENCIÓN"
            RsAux.Close
            Screen.MousePointer = 0
            Exit Sub
        End If
        RsAux.Close
        
        
        Cons = "Select * from MotivoServicio Where MSeID = " & .Cell(flexcpData, .Row, 0)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Delete
        RsAux.Close
        
        .RemoveItem .Row
    
        LimpioFicha
        DeshabilitoIngreso
    End With
    Screen.MousePointer = 0
    Exit Sub
    
Error:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
End Sub

Sub AccionCancelar()

    On Error Resume Next
    DeshabilitoIngreso
    LimpioFicha
    
    If vsLista.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me Else Botones True, False, False, False, False, Toolbar1, Me
    
    sNuevo = False: sModificar = False
    
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select

End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    If cTipo.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de producto al que está asociado el motivo.", vbExclamation, "ATENCIÓN"
        Foco cTipo: Exit Function
    End If
    If Trim(tNombre.Text) = "" Then
        MsgBox "Debe ingresar el nombre o descripción del motivo de servicio.", vbExclamation, "ATENCIÓN"
        Foco tNombre: Exit Function
    End If
    
    'Valido el nombre para el tipo  ---------------------------------------------------------------------------
    Cons = "Select * from MotivoServicio " _
           & " Where MSeNombre= '" & tNombre.Text & "'" _
           & " And MSeID <> " & gIdMotivo _
           & " And MSeTipo = " & cTipo.ItemData(cTipo.ListIndex)
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        MsgBox "Hay un motivo para el nombre " & Trim(tNombre.Text) & Chr(vbKeyReturn) & "No se podrán almacenar los datos.", vbExclamation, "ATENCÍON"
        RsAux.Close: Exit Function
    End If
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------------
    
    ValidoCampos = True
    
End Function

Private Sub DeshabilitoIngreso()
    
    tNombre.Enabled = False: tNombre.BackColor = Inactivo
    cTipo.Enabled = False: cTipo.BackColor = Inactivo
    
    vsLista.Enabled = True: vsLista.BackColor = Blanco
        
End Sub

Private Sub HabilitoIngreso()
    
    With tNombre
        .Enabled = True: .BackColor = Obligatorio
    End With
    cTipo.Enabled = True: cTipo.BackColor = Colores.Obligatorio
    vsLista.Enabled = False: vsLista.BackColor = Inactivo
    
End Sub

Private Sub CargoCamposBD()
    
    RsAux!MSeNombre = Trim(tNombre.Text)
    RsAux!MSeTipo = cTipo.ItemData(cTipo.ListIndex)
    
End Sub

Private Sub LimpioFicha()
    tNombre.Text = "": cTipo.Text = ""
End Sub

Private Sub InicializoGrilla()

    On Error Resume Next
    With vsLista
        .Cols = 1: .Rows = 1
        .FormatString = "<Nombre|Tipo"
        .ColHidden(1) = True
        .WordWrap = True
        .ExtendLastCol = True
    End With
      
End Sub

Private Sub CargoLista()

    On Error GoTo errCargar
    
    If cConsulta.ListIndex = -1 Then
        MsgBox "Seleccione un tipo de producto para cargar la lista.", vbInformation, "ATENCIÓN"
        Foco cConsulta: Exit Sub
    End If
    
    Dim aValor As Long
    Screen.MousePointer = 11
    Botones True, False, False, False, False, Toolbar1, Me
    Cons = "Select * from MotivoServicio Where MSeTipo = " & cConsulta.ItemData(cConsulta.ListIndex) & " Order by MSeNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    vsLista.Rows = 1
    
    Do While Not RsAux.EOF
        With vsLista
            .AddItem Trim(RsAux!MSeNombre)
            aValor = RsAux!MSeID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            aValor = RsAux!MSeTipo: .Cell(flexcpData, .Rows - 1, 1) = aValor
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If vsLista.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar la lista.", Err.Description
End Sub


Private Sub vsLista_Click()

    If vsLista.MouseRow = 0 Then
        vsLista.ColSel = vsLista.MouseCol
        If vsLista.ColSort(vsLista.MouseCol) = flexSortGenericAscending Then
            vsLista.ColSort(vsLista.MouseCol) = flexSortGenericDescending
        Else
            vsLista.ColSort(vsLista.MouseCol) = flexSortGenericAscending
        End If
        vsLista.Sort = flexSortUseColSort
    End If
    
End Sub

