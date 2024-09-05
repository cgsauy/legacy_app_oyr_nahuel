VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLista 
   Caption         =   "Ingreso de Valores y Condiciones"
   ClientHeight    =   6330
   ClientLeft      =   2220
   ClientTop       =   3225
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLista.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9390
   Begin VB.TextBox tBuscar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin ComctlLib.TabStrip tabValores 
      Height          =   795
      Left            =   600
      TabIndex        =   6
      Top             =   420
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1402
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Valores Calculados"
            Key             =   "valores"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Condiciones   "
            Key             =   "condiciones"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMarco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      ScaleHeight     =   375
      ScaleWidth      =   6495
      TabIndex        =   3
      Top             =   5280
      Width           =   6555
      Begin VB.TextBox tprmSolicitud 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Text            =   "3484"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton bNuevo 
         Caption         =   "&Nuevo"
         Height          =   315
         Left            =   3900
         TabIndex        =   5
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton bEditar 
         Caption         =   "&Editar"
         Height          =   315
         Left            =   5100
         TabIndex        =   4
         Top             =   60
         Width           =   1095
      End
      Begin VB.Label lblDebug 
         Caption         =   "Debug on"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   105
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "[prmSolicitud]:"
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   105
         Width           =   1035
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsValores 
      Height          =   3195
      Left            =   1320
      TabIndex        =   2
      Top             =   1140
      Width           =   2295
      _ExtentX        =   4048
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsCondiciones 
      Height          =   3195
      Left            =   5100
      TabIndex        =   7
      Top             =   1140
      Width           =   3375
      _ExtentX        =   5953
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Filtrar expresiones con la Descripción:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   2715
   End
   Begin VB.Menu MnuFiltros 
      Caption         =   "Filtrar"
      Visible         =   0   'False
      Begin VB.Menu MnuAcciones 
         Caption         =   "Acciones"
      End
      Begin VB.Menu MnuFL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBUso 
         Caption         =   "Buscar uso de Fórmula"
      End
      Begin VB.Menu MnuActualizar 
         Caption         =   "Actualizar Lista"
      End
   End
End
Attribute VB_Name = "frmLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub VerEstadoDebug()
On Error GoTo errD
    lblDebug.Caption = "Debug off": lblDebug.ForeColor = &H80&
    Cons = "SELECT IsNull(parValor, 0) From Parametro WHERE ParNombre = 'rAutoDebugON'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux(0) <> 0 Then lblDebug.Caption = "Debug on": lblDebug.ForeColor = &H8000&
    End If
    RsAux.Close
    Exit Sub
errD:
    clsGeneral.OcurrioError "Error al validar el estado del debug", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoValores(Optional ConFormula As String = "", Optional ConDesc As String = "", Optional selID As Long = 0)
Dim aValor As Long
Dim mRowSel As Integer

    Screen.MousePointer = 11
    vsValores.Rows = 1
    
    Cons = "Select * from ValoresCalculados Where VCaTipo = 1 "
    If ConFormula <> "" Then Cons = Cons & " And VCaTexto like '%[[]" & Trim(ConFormula) & "]%'"
    If ConDesc <> "" Then Cons = Cons & " And VCaDescripcion like '%" & Trim(ConDesc) & "%'"
    Cons = Cons & " Order by VCaNombre"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsValores
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!VCaNombre)
            aValor = RsAux!VCaCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            If Not IsNull(RsAux!VCaDescripcion) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!VCaDescripcion)
            
            If aValor = selID Then mRowSel = .Rows - 1
        End With
        RsAux.MoveNext
        
    Loop
    RsAux.Close
    
    On Error Resume Next
    vsValores.Select mRowSel, 0, , vsValores.Cols - 1
    
    Screen.MousePointer = 0
    Exit Sub
errEvaluar:
    clsGeneral.OcurrioError "Error al cargar la lista de valores", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoCondiciones(Optional ConFormula As String = "", Optional ConDesc As String = "", Optional selID As Long = 0)
Dim aValor As Long
Dim mRowSel As Integer

    Screen.MousePointer = 11
    vsCondiciones.Rows = 1
    
'    Cons = "Select * from ValoresCalculados " & _
'                    " Left Outer Join CondicionResolucion On VCaResolucionSiSi = ConCodigo" & _
'               " Where VCaTipo = 2 "
'
'    If ConFormula <> "" Then Cons = Cons & " And VCaTexto like '%[[]" & Trim(ConFormula) & "]%'"
'    If ConDesc <> "" Then Cons = Cons & " And VCaDescripcion like '%" & Trim(ConDesc) & "%'"
'    Cons = Cons & " Order by VCaNombre"
    
    Set RsAux = cBase.OpenResultset("EXEC QryCondicionesResAutom", rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsCondiciones
            .AddItem ""
            aValor = RsAux!VCaCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            
            .Cell(flexcpText, .Rows - 1, 0) = RsAux("Orden")
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!VCaNombre)
            If Not IsNull(RsAux!VCaDescripcion) Then .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!VCaDescripcion)
            
            If Not IsNull(RsAux!Caminos) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!Caminos)
            If Not IsNull(RsAux!Resolucion) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!Resolucion)
        
            If aValor = selID Then mRowSel = .Rows - 1
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    On Error Resume Next
    vsCondiciones.Select mRowSel, 0, , vsCondiciones.Cols - 1
    
    Screen.MousePointer = 0
    Exit Sub
errEvaluar:
    clsGeneral.OcurrioError "Error al cargar la lista de condiciones", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub InicializoGrilla()

    On Error Resume Next
    With vsValores
        .Cols = 1: .Rows = 1
        .FormatString = "<Nombre|<Descripción"
            
        .WordWrap = True
        .ColWidth(0) = 1800: .ColWidth(1) = 2800
        
        .ExtendLastCol = True: .FixedCols = 0
    End With
    
    With vsCondiciones
        .Cols = 1: .Rows = 1
        .FormatString = "Orden|<Nombre|<Descripción|Caminos|<Resolución"
            
        .WordWrap = True
        .ColWidth(1) = 1800: .ColWidth(2) = 2800: .ColWidth(3) = 3000: .ColWidth(4) = 3000
        
        .ExtendLastCol = True: .FixedCols = 0
        .AllowUserResizing = flexResizeColumns
    End With
      
End Sub

Private Sub bEditar_Click()

    Select Case tabValores.SelectedItem.Key
        Case "valores"
                If vsValores.Rows = 1 Then Exit Sub
                AvtivoFormularioValores vsValores.Cell(flexcpData, vsValores.Row, 0)
        
        Case "condiciones"
                If vsCondiciones.Rows = 1 Then Exit Sub
                AvtivoFormularioCondiciones vsCondiciones.Cell(flexcpData, vsCondiciones.Row, 0)
    End Select
        
End Sub

Private Sub AvtivoFormularioValores(Optional idValor As Long = 0)

    frmValores.prmIdValor = idValor
    frmValores.tprmSolicitud.Text = Trim(tprmSolicitud.Text)
    frmValores.Show vbModal, Me
    Me.Refresh
    
     If frmValores.prmGrabo Then CargoValores selID:=idValor
     
End Sub

Private Sub AvtivoFormularioCondiciones(Optional idValor As Long = 0)

    frmCondiciones.prmIdValor = idValor
    frmCondiciones.tprmSolicitud.Text = Trim(tprmSolicitud.Text)
    frmCondiciones.Show vbModal, Me
    
    Me.Refresh
    
    If frmCondiciones.prmGrabo Then CargoCondiciones selID:=idValor
    
End Sub

Private Sub bNuevo_Click()
    Select Case tabValores.SelectedItem.Key
        Case "valores": AvtivoFormularioValores
        Case "condiciones": AvtivoFormularioCondiciones
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    ObtengoSeteoForm Me
    
    tprmSolicitud.Text = GetSetting("AutoResolucion", "Configuracion", "Solicitud", "") 'tprmSolicitud.Text
    
    InicializoGrilla
    CargoValores
    CargoCondiciones
    
    picMarco.BorderStyle = 0
    
    tabValores.SelectedItem = tabValores.Tabs("valores")
    
    VerEstadoDebug
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    With picMarco
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleHeight - .Height
        
        bEditar.Left = picMarco.ScaleWidth - bEditar.Width - 100
        bNuevo.Left = bEditar.Left - bEditar.Width - 100
    End With
    
    With tabValores
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleTop + 420: .Height = Me.ScaleHeight - picMarco.Height - .Top
    End With
    
    With vsValores
        .Left = tabValores.ClientLeft: .Width = tabValores.ClientWidth
        .Top = tabValores.ClientTop: .Height = tabValores.ClientHeight
    End With
    
    With vsCondiciones
        .Left = vsValores.Left: .Width = vsValores.Width
        .Top = vsValores.Top: .Height = vsValores.Height
    End With
    
    tBuscar.Width = Me.ScaleWidth - tBuscar.Left - 60
    
    lblDebug.Left = (Me.ScaleWidth - lblDebug.Width) / 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    cBase.Close
    eBase.Close
    End
End Sub


Private Sub Label6_DblClick()
On Error GoTo errS
    If IsNumeric(tprmSolicitud.Text) Then
        VBA.SaveSetting "AutoResolucion", "Configuracion", "Solicitud", tprmSolicitud.Text
        MsgBox "Datos guardados en la registry.", vbInformation, "Solicitud default"
    End If
    Exit Sub
errS:
    clsGeneral.OcurrioError "Error al "
End Sub

Private Sub lblDebug_Click()
    MsgBox "Prender/apagar Tabla Parametro, ParNombre = 'rAutoDebugON', se almacena en la tabla Debug.", vbInformation, "Ayuda"
End Sub

Private Sub MnuActualizar_Click()

    CargoValores
    CargoCondiciones
    
End Sub

Private Sub MnuBUso_Click()

    If vsValores.Rows = 1 Then Exit Sub
    
    CargoValores ConFormula:=Trim(MnuBUso.Tag)
    CargoCondiciones ConFormula:=Trim(MnuBUso.Tag)
    
End Sub

Private Sub tabValores_Click()
    With tabValores
        Select Case .SelectedItem.Key
            Case "valores": vsValores.ZOrder 0
            Case "condiciones": vsCondiciones.ZOrder 0
        End Select
    End With
End Sub

Private Sub tBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tBuscar.Text) = "" Then Exit Sub
        
        Select Case tabValores.SelectedItem.Key
            Case "valores": CargoValores ConDesc:=Trim(tBuscar.Text)
            Case "condiciones": CargoCondiciones ConDesc:=Trim(tBuscar.Text)
        End Select
        
    End If
    
End Sub

Private Sub vsCondiciones_DblClick()
    Call bEditar_Click
End Sub

Private Sub vsCondiciones_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        If vsCondiciones.Rows = 1 Then
            MnuBUso.Enabled = False
        Else
            MnuBUso.Enabled = True
            MnuBUso.Tag = Trim(vsCondiciones.Cell(flexcpText, vsCondiciones.Row, 0))
        End If
        PopupMenu MnuFiltros
    End If
    
End Sub

Private Sub vsValores_DblClick()
    Call bEditar_Click
End Sub

Private Sub vsValores_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        If vsValores.Rows = 1 Then
            MnuBUso.Enabled = False
        Else
            MnuBUso.Enabled = True
            MnuBUso.Tag = Trim(vsValores.Cell(flexcpText, vsValores.Row, 0))
        End If
        PopupMenu MnuFiltros
    End If
    
End Sub
