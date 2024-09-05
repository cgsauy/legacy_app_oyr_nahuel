VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Begin VB.Form frmFuncEmpresa 
   Caption         =   "Empleados de Empresas"
   ClientHeight    =   6420
   ClientLeft      =   2205
   ClientTop       =   2610
   ClientWidth     =   9525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFuncEmpresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9525
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   2880
      TabIndex        =   6
      Top             =   1200
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      OutlineBar      =   1
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   6165
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
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
            Object.Width           =   8678
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   9135
      Begin VB.TextBox tEmpresa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   4695
      End
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   425
         Left            =   5880
         ScaleHeight     =   420
         ScaleWidth      =   2415
         TabIndex        =   10
         Top             =   200
         Width           =   2415
         Begin VB.CommandButton bImprimir 
            Height          =   310
            Left            =   720
            Picture         =   "frmFuncEmpresa.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Imprimir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bNoFiltros 
            Height          =   310
            Left            =   1080
            Picture         =   "frmFuncEmpresa.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Quitar filtros."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bCancelar 
            Height          =   310
            Left            =   1800
            Picture         =   "frmFuncEmpresa.frx":090A
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Salir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bConsultar 
            Height          =   310
            Left            =   120
            Picture         =   "frmFuncEmpresa.frx":0A0C
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Ejecutar."
            Top             =   50
            Width           =   310
         End
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Empresa:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   7858
      _StockProps     =   229
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
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.Menu MnuAccesos 
      Caption         =   "MnuAccesos"
      Visible         =   0   'False
      Begin VB.Menu MnuTitulo 
         Caption         =   "Menú Clientes"
         Checked         =   -1  'True
      End
      Begin VB.Menu Mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOperaciones 
         Caption         =   "&Visualización de Operaciones"
      End
      Begin VB.Menu MnuFicha 
         Caption         =   "Ficha de Cliente"
      End
      Begin VB.Menu MnuEmpleos 
         Caption         =   "Empleos"
      End
      Begin VB.Menu MnuComentarios 
         Caption         =   "Comentarios"
      End
      Begin VB.Menu MnuAcL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "Cancelar menú"
      End
   End
End
Attribute VB_Name = "frmFuncEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsAux As rdoResultset
Private aTexto As String

Private Sub AccionLimpiar()
    tEmpresa.Text = "": tEmpresa.Tag = "0"
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub
Private Sub bImprimir_Click()
    AccionImprimir
End Sub
Private Sub bNoFiltros_Click()
    AccionLimpiar
End Sub

Private Sub Label5_Click()
    Foco tEmpresa
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad
    FechaDelServidor
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    InicializoGrillas
    AccionLimpiar
    
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "

    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "<C.I.|<Nombre|<Ingresó|<Ocupación (Cargo)|<Tipo de Ingreso|<Comentarios"
        .ColWidth(0) = 1000: .ColWidth(1) = 2000: .ColWidth(3) = 2000: .ColWidth(4) = 3000
        '.ColWidth(2) = 1000
        .ColWidth(5) = 4500
    End With
      
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyE: AccionConsultar
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11
    picBotones.BorderStyle = vbFlat
    fFiltros.Width = Me.Width - (fFiltros.Left * 2.5)
    
    vsConsulta.Left = fFiltros.Left
    vsConsulta.Top = fFiltros.Top + fFiltros.Height + 50
    vsConsulta.Height = Me.ScaleHeight - (vsConsulta.Top + Status.Height + 90)
    vsConsulta.Width = fFiltros.Width
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
Dim CodAux As Long

    If tEmpresa.Tag = "0" Then vsConsulta.Rows = 1: Exit Sub
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    
    cons = "Select * From Cliente, CPersona, Ocupacion, Empleo " _
        & " Where EmpCEmpresa = " & Val(tEmpresa.Tag) _
        & " And EmpCliente = CPeCliente " _
        & " And CPeCliente = CliCodigo" _
        & " And EmpOcupacion = OcuCodigo" '_
        '& " And EmpTipoExhibido *= TExCodigo"
           
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        MsgBox "No hay datos clientes que trabajen en esa empresa.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: InicializoGrillas: Exit Sub
    End If
    
    Dim aRsEmp As rdoResultset
    
    With vsConsulta
        .Rows = 1
        Do While Not rsAux.EOF
            .AddItem ""
            
            CodAux = rsAux!CliCodigo: .Cell(flexcpData, .Rows - 1, 0) = CodAux
            If Not IsNull(rsAux!CLiCIRUC) Then .Cell(flexcpText, .Rows - 1, 0) = clsGeneral.RetornoFormatoCedula(rsAux!CLiCIRUC)
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!CPeApellido1) & ", " & Trim(rsAux!CPeNombre1)
            If Not IsNull(rsAux!EmpFechaIngreso) Then .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!EmpFechaIngreso, "mm/yy")
            
            aTexto = Trim(rsAux!OcuNombre)
            If Not IsNull(rsAux!EmpCargo) Then aTexto = aTexto & " (" & Trim(rsAux!EmpCargo) & ")"
            .Cell(flexcpText, .Rows - 1, 3) = aTexto
            
            If Not IsNull(rsAux!EmpComentario) Then .Cell(flexcpText, .Rows - 1, 5) = Trim(rsAux!EmpComentario)
            
            'Saco los datos de tablas auxiliares del empleo---------------------------------------------------------------------
            cons = "Select TInNombre, MonSigno, VEmNombre" _
                    & " From Empleo, TipoIngreso, Moneda, VigenciaEmpleo" _
                    & " Where EmpCodigo = " & rsAux!EmpCodigo _
                    & " And EmpTipoIngreso *= TInCodigo" _
                    & " And EmpMoneda *= MonCodigo" _
                    & " And EmpVigencia *= VEmCodigo"
            Set aRsEmp = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            aTexto = ""
            If Not IsNull(aRsEmp!MonSigno) Then aTexto = Trim(aRsEmp!MonSigno)
            If Not IsNull(rsAux!EmpLiquido) Then aTexto = aTexto & " Líq." & Format(rsAux!EmpLiquido, "#,##0") & " "
            If Not IsNull(rsAux!EmpNominal) Then aTexto = aTexto & " Nom." & Format(rsAux!EmpNominal, "#,##0")
            
            If IsNull(rsAux!EmpLiquido) Or IsNull(rsAux!EmpNominal) Then
                If Not IsNull(aRsEmp!TInNombre) Then aTexto = aTexto & " " & Trim(aRsEmp!TInNombre)
                If Not IsNull(rsAux!EmpIngreso) Then aTexto = aTexto & " " & Format(rsAux!EmpIngreso, "#,##0")
            End If
            If Not IsNull(aRsEmp!VEmNombre) Then aTexto = aTexto & " (" & Trim(aRsEmp!VEmNombre) & ")"
            .Cell(flexcpText, .Rows - 1, 4) = aTexto
            aRsEmp.Close
            '----------------------------------------------------------------------------------------------------------------------------
            
            If Not IsNull(rsAux!EmpTipoExhibido) Then
                'aTexto = Trim(rsAux!TExAbreviacion)
                If Not IsNull(rsAux!EmpExhibido) Then
                    aTexto = aTexto & " " & Format(rsAux!EmpExhibido, "dd/mm/yy")
                    If Abs(DateDiff("d", rsAux!EmpExhibido, gFechaServidor)) < 45 Then  'Es reciente
                        .Cell(flexcpBackColor, .Rows - 1, 2) = Colores.Obligatorio
                        If Not IsNull(rsAux!EmpFModificacion) Then
                            If Format(rsAux!EmpFModificacion, "dd/mm/yyyy") = Format(gFechaServidor, "dd/mm/yyyy") Then .Cell(flexcpBackColor, .Rows - 1, 2) = Colores.clVerde
                        End If
                    End If
                End If
                '.Cell(flexcpText, .Rows - 1, 4) = aTexto
            End If
            If Not IsNull(rsAux!EmpVigencia) Then
                If rsAux!EmpVigencia = paempNoTrabajaMas Then .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
                If rsAux!EmpVigencia = paempSeguroParo Then .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris
            End If
            
            rsAux.MoveNext
        Loop
        rsAux.Close
        
    End With
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionImprimir()
Dim J As Integer

    If vsConsulta.Rows = 1 Then
        MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo errPrint
    Screen.MousePointer = 11
    
    With vsListado
        .Orientation = orLandscape
        .Preview = True
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
    
        EncabezadoListado vsListado, "Funcionarios de Empresas", False
        
        .FileName = "Funcionarios de Empresas"
        
        .FontSize = 9: .FontBold = True
        aTexto = tEmpresa.Text
        .Paragraph = "": .Paragraph = "Empresa: " & aTexto: .Paragraph = ""
        .FontSize = 8: .FontBold = False
        vsConsulta.ExtendLastCol = False: .RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        .EndDoc
        If Not .PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 11: Exit Sub
        .PrintDoc
    End With
    
    Screen.MousePointer = 0
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión. ", Err.Description
End Sub


Private Sub MnuComentarios_Click()
    On Error GoTo errObj
    If vsConsulta.Rows = 1 Then Exit Sub
    Screen.MousePointer = 11
    
    Dim objCliente As New clsCliente
    objCliente.Comentarios idCliente:=vsConsulta.Cell(flexcpData, vsConsulta.Row, 0)
    Me.Refresh
    Set objCliente = Nothing
    Screen.MousePointer = 0
    Exit Sub

errObj:
    clsGeneral.OcurrioError "Ocurrió un error al activar la aplicación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuEmpleos_Click()
On Error GoTo errObj
    If vsConsulta.Rows = 1 Then Exit Sub
    Screen.MousePointer = 11
    
    Dim objCliente As New clsCliente
    objCliente.Empleos idCliente:=vsConsulta.Cell(flexcpData, vsConsulta.Row, 0)
    Me.Refresh
    Set objCliente = Nothing
    Screen.MousePointer = 0
    Exit Sub

errObj:
    clsGeneral.OcurrioError "Ocurrió un error al activar la aplicación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuFicha_Click()
    
    On Error GoTo errObj
    If vsConsulta.Rows = 1 Then Exit Sub
    Screen.MousePointer = 11
    
    Dim objCliente As New clsCliente
    objCliente.Personas idCliente:=vsConsulta.Cell(flexcpData, vsConsulta.Row, 0)
    Me.Refresh
    Set objCliente = Nothing
    Screen.MousePointer = 0
    Exit Sub

errObj:
    clsGeneral.OcurrioError "Ocurrió un error al activar la aplicación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuOperaciones_Click()
On Error GoTo errObj
    If vsConsulta.Rows = 1 Then Exit Sub
    
    Dim idCliente As Long
    idCliente = vsConsulta.Cell(flexcpData, vsConsulta.Row, 0)
    EjecutarApp App.Path & "\Visualizacion de Operaciones", CStr(idCliente)
    Exit Sub

errObj:
    clsGeneral.OcurrioError "Ocurrió un error al activar la aplicación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tEmpresa_GotFocus()
    tEmpresa.SelStart = 0: tEmpresa.SelLength = Len(tEmpresa.Text)
End Sub

Private Sub tEmpresa_KeyPress(KeyAscii As Integer)
On Error GoTo ErrBE

    If KeyAscii = vbKeyReturn Then
        tEmpresa.Tag = "0"
        If Trim(tEmpresa.Text) <> "" Then
            Screen.MousePointer = 11
            Dim aQ As Integer: aQ = 0
            Dim aIDSel As Long
            Dim aTexto As String
            
            aTexto = Replace(RTrim(tEmpresa.Text), " ", "%")
            
            cons = "Select CliCodigo, 'Nombre Fantasia' = CEmFantasia, 'Razón Social' = CEmNombre From Cliente, CEmpresa " _
                    & " Where (CEmFantasia Like '" & aTexto & "%'" _
                    & " Or CEmNombre Like '" & aTexto & "%')" _
                    & " And CliCodigo = CEmCliente Order by CEmFantasia, CEmNombre"
            Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
            If Not rsAux.EOF Then
                aQ = 1
                aIDSel = rsAux!CliCodigo: aTexto = Trim(rsAux(1))
                rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
            End If
            rsAux.Close
            
            Select Case aQ
                Case 0: MsgBox "No hay empresas para el texto ingresado.", vbExclamation, "No hay Datos"
                
                Case 2
                        Dim LiAyuda As New clsListadeAyuda
                        aIDSel = 0
                        aIDSel = LiAyuda.ActivarAyuda(cBase, cons, 5500, 1, "Lista de Empresas")
                        If aIDSel <> 0 Then
                            aIDSel = LiAyuda.RetornoDatoSeleccionado(0)
                            aTexto = LiAyuda.RetornoDatoSeleccionado(1)
                        End If
                        Set LiAyuda = Nothing
            
            End Select
            
            If aIDSel <> 0 Then
                tEmpresa.Text = Trim(aTexto)
                tEmpresa.Tag = aIDSel
                Foco bConsultar
            End If
            
            Screen.MousePointer = 0
        End If
    End If
    Exit Sub
ErrBE:
    clsGeneral.OcurrioError "Error al buscar la empresa.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton And vsConsulta.Rows > 1 Then
        vsConsulta.Select vsConsulta.MouseRow, 0
        PopupMenu MnuAccesos
    End If
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

