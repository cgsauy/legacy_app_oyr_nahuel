VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmMaAcceso 
   BackColor       =   &H8000000B&
   Caption         =   "Mantenimiento de Accesos"
   ClientHeight    =   6990
   ClientLeft      =   2370
   ClientTop       =   3060
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaAcceso.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   10335
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList2"
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
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tBuscar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton bNivel 
      Caption         =   "..."
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox tNivel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6735
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10107
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   1485
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2619
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
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
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
      Caption         =   "&Buscar:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lNivel 
      Caption         =   "&Nivel:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   7320
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
            Picture         =   "frmMaAcceso.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaAcceso.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaAcceso.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaAcceso.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaAcceso.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaAcceso.frx":0BA4
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
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuOpL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "Del formulario "
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuFiltros 
      Caption         =   "MnuFiltros"
      Visible         =   0   'False
      Begin VB.Menu MnuFiTitulo 
         Caption         =   "Filtrar referencia"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuFiL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFiFiltrar 
         Caption         =   "Filtrar"
      End
      Begin VB.Menu MnuFiEliminar 
         Caption         =   "Quitar Filtros"
      End
      Begin VB.Menu MnuUsuarios 
         Caption         =   "Ver Usuarios c/Acceso"
      End
   End
End
Attribute VB_Name = "frmMaAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean
Dim aFiltro As String

Private Sub bNivel_Click()
    ListaDeNiveles
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error Resume Next
    sNuevo = False: sModificar = False
    
    Screen.MousePointer = 11
    Botones True, False, False, False, False, Toolbar1, Me
    
    DeshabilitoIngreso
    InicializoGrilla
    
    ObtengoSeteoForm Me, WidthIni:=10215, HeightIni:=6825
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsLista.Width = Me.Width - 350
    vsLista.Height = Me.Height - vsLista.Top - Status.Height - Toolbar1.Height - 450
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    
     GuardoSeteoForm Me
     
End Sub

Private Sub lNivel_Click()
    Foco tNivel
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuFiEliminar_Click()
On Error Resume Next
    
    With vsLista
    For I = 1 To .Rows - 1
        .RowHidden(I) = False
    Next
    End With

End Sub

Private Sub MnuFiFiltrar_Click()
On Error Resume Next
    
    With vsLista
    For I = 1 To .Rows - 1
        If Trim(.Cell(flexcpText, I, 1)) <> Trim(aFiltro) Then .RowHidden(I) = True
    Next
    End With
    
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

Private Sub MnuUsuarios_Click()
Dim aValor As Long

    On Error GoTo errUsu
    aValor = vsLista.Cell(flexcpData, vsLista.Row, 0)
    Screen.MousePointer = 11
    
    Cons = "Select RTrim(cast(UsuCodigo as char(7))) + Rtrim(cast(NSiNivel as char(7))), UsuIdentificacion 'Usuario', NSiNombre as 'Nivel de Acceso'" & _
                " From NivelPermiso, UsuarioNivel, Usuario, NivelSistema" & _
                " Where NPeNivel = UNiNivel" & _
                " And UNiUsuario = UsuCodigo" & _
                " And NPeAplicacion = " & aValor & _
                " And NPeNivel = NSiNivel" & _
                " Order by UsuIdentificacion"
                
    Dim aLista As New clsListadeAyuda
    aLista.ActivoListaAyuda Cons, False, cBase.Connect, 4100
    Me.Refresh
    Set aLista = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errUsu:
    clsGeneral.OcurrioError "Error al activar la lista. " & Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub tBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tBuscar.Text = "" Then Exit Sub
        
        Dim Idx As Long
        With vsLista
            For Idx = 1 To .Rows - 1
                If InStr(LCase(.Cell(flexcpText, Idx, 2)), LCase(tBuscar.Text)) = 0 And InStr(LCase(.Cell(flexcpText, Idx, 3)), LCase(tBuscar.Text)) = 0 Then
                    .RowHidden(Idx) = True
                End If
            Next
        End With
    End If
    
End Sub

Private Sub tNivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then vsLista.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "grabar": AccionGrabar
        Case "eliminar": AccionEliminar
        Case "cancelar": AccionCancelar
    End Select

End Sub

Private Sub AccionNuevo()

    Screen.MousePointer = 11
    CargoAplicaciones 0
    
    tNivel.Text = ""
    Foco tNivel
    
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    sNuevo = True
    Screen.MousePointer = 0
    
End Sub

Private Sub AccionCancelar()

Dim aNivel As Long

    Screen.MousePointer = 11
    aNivel = Val(tNivel.Tag)
    DeshabilitoIngreso
    
    If sNuevo Then
        Botones True, False, False, False, False, Toolbar1, Me
        vsLista.Rows = 1
        tNivel.Text = ""
    Else
        Botones True, True, True, False, False, Toolbar1, Me
        CargoAplicaciones aNivel
    End If
    
    sNuevo = False: sModificar = False
    Screen.MousePointer = 0
    
End Sub

Private Sub AccionModificar()
    
    Screen.MousePointer = 11
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    sModificar = True
    Foco tNivel
    Screen.MousePointer = 0

End Sub

Private Sub AccionGrabar()
    
    If Trim(tNivel.Text) = "" Then
        MsgBox "Ingrese un nombre para el nivel de acceso.", vbExclamation, "ATENCIÓN"
        Foco tNivel: Exit Sub
    End If
    
    If MsgBox("Confirma grabar los permisos para el nivel.", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    CargoDatosBD
    
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------
    
    Botones True, True, True, False, False, Toolbar1, Me
    sNuevo = False: sModificar = False
    DeshabilitoIngreso
    Screen.MousePointer = 0
    Exit Sub

errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
End Sub

Private Sub CargoDatosBD()
    
Dim aCodigoNivel As Long

    FechaDelServidor
    aCodigoNivel = 0
    
    'Inserto Datos en BD NIVEL-SISTEMA-------------------------------------------------------------------
    If sNuevo Then
        Cons = "Insert into NivelSistema (NSiNombre, NSiFModificado, NSiUsuario) Values (" _
               & "'" & Trim(tNivel.Text) & "', " _
               & "'" & Format(gFechaServidor, sqlFormatoFH) & "', " _
               & paCodigoDeUsuario & ")"
    
    Else
        aCodigoNivel = CLng(tNivel.Tag)
        Cons = "Delete NivelPermiso Where NPeNivel = " & aCodigoNivel
        cBase.Execute Cons
        
        Cons = "Update NivelSistema Set " _
               & " NSiNombre = '" & Trim(tNivel.Text) & "', " _
               & " NSiFModificado  = '" & Format(gFechaServidor, sqlFormatoFH) & "', " _
               & " NSiUsuario = " & paCodigoDeUsuario _
               & " Where NSiNivel = " & aCodigoNivel
    
    End If
    
    cBase.Execute Cons
    '---------------------------------------------------------------------------------------------------------------
    
    If aCodigoNivel = 0 Then    'Saco Codigo del Nivel Recién ingresado
        Cons = "Select Max(NSiNivel) from NivelSistema"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        aCodigoNivel = RsAux(0)
        RsAux.Close
        tNivel.Tag = aCodigoNivel
    End If
    '---------------------------------------------------------------------------------------------------------------
    
    'Inserto los datos en BD NIVEL-PERMISO------------------------------------------------------------------
    With vsLista
    For I = 1 To .Rows - 1
        If .Cell(flexcpChecked, I, 0) = flexChecked Then
            Cons = "Insert Into NivelPermiso (NPeNivel, NPeAplicacion) Values (" & aCodigoNivel & ", " & .Cell(flexcpData, I, 0) & ")"
            cBase.Execute Cons
        End If
    Next I
    End With
    '---------------------------------------------------------------------------------------------------------------
    
End Sub
Private Sub HabilitoIngreso()

    tNivel.Enabled = True
    tNivel.BackColor = Obligatorio
    bNivel.Enabled = False
    
    vsLista.Editable = True
    
End Sub

Private Sub DeshabilitoIngreso()

    bNivel.Enabled = True
    tNivel.Enabled = False
    tNivel.BackColor = Inactivo
    
    vsLista.Editable = False
    
End Sub

Private Sub ListaDeNiveles()
Dim aSeleccionado As Long

    On Error GoTo errBuscar
    Screen.MousePointer = 11
    Dim aLista As New clsListadeAyuda
    
    Cons = "Select NSiNivel 'ID_Nivel', NSiNombre 'Nombre del Nivel', NSiFModificado 'Modificado' From NivelSistema " _
            & " Order by NSiNombre"
    aLista.ActivoListaAyuda Cons, False, cBase.Connect, 4500
    Me.Refresh
    
    aSeleccionado = aLista.ValorSeleccionado
    If aSeleccionado <> 0 Then
        tNivel.Text = aLista.ItemSeleccionado
        tNivel.Tag = aSeleccionado
    Else
        tNivel.Text = ""
    End If
    
    Set aLista = Nothing
    
    If Val(tNivel.Tag) <> 0 Then CargoAplicaciones CLng(tNivel.Tag)
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al acceder a los datos.", Err.Description
    Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub CargoAplicaciones(Nivel As Long)
Dim aValor As Long

    On Error GoTo errNivel
    With vsLista
    
    'Cargo Datos del Nivel ----------------------------------------------------------------------
    Cons = "Select * from Aplicacion Left Outer Join NivelPermiso On NPeAplicacion = AplCodigo And NPeNivel = " & Nivel
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
    .Rows = 1
    Do While Not RsAux.EOF
        .AddItem ""
        If Not IsNull(RsAux!AplReferencia) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!AplReferencia)
        .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!AplNombre)
        .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!AplDescripcion)
        If Not IsNull(RsAux!NPeNivel) Then .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
        
        aValor = RsAux!AplCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
                
        RsAux.MoveNext
    Loop
    RsAux.Close
    '------------------------------------------------------------------------------------------------
    Botones True, True, True, False, False, Toolbar1, Me
    
    '.ColSort(0) = flexSortGenericAscending
    .Select 1, 0, , 1
    .Sort = flexSortUseColSort
    End With
    Exit Sub

errNivel:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del nivel seleccionado.", Err.Description
End Sub

Private Sub AccionEliminar()

    'Valido Si hay usuarios con el nivel asignado-----------------------------------------------
    Cons = "Select * from UsuarioNivel Where UNiNivel = " & Val(tNivel.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If MsgBox("Hay usuarios con el nivel seleccionado." & Chr(vbKeyReturn) & "Si elimina el nivel, los usuarios no podrán ingresar a los sistemas" _
                & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Dese eliminarlo.", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then
            RsAux.Close: Exit Sub
        Else
            Cons = "Delete UsuarioNivel Where UNiNivel = " & Val(tNivel.Tag)
            cBase.Execute Cons
        End If
    Else
        If MsgBox("Confirma eliminar el nivel seleccionado.", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then RsAux.Close: Exit Sub
    End If
    RsAux.Close
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET

    Cons = "Delete NivelPermiso Where NPeNivel = " & Val(tNivel.Tag)
    cBase.Execute Cons
    
    Cons = "Delete NivelSistema Where NSiNivel = " & Val(tNivel.Tag)
    cBase.Execute Cons
    
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------
    
    Botones True, False, False, False, False, Toolbar1, Me
    DeshabilitoIngreso
    vsLista.Rows = 1
    tNivel.Tag = 0: tNivel.Text = ""
    
    Screen.MousePointer = 0
    Exit Sub

errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
End Sub


Private Sub InicializoGrilla()

    On Error Resume Next
    With vsLista
        .Cols = 1: .Rows = 1:
        .FormatString = "Habilitado|Referencia|Aplicación|Descripción"
            
        .WordWrap = True
        .ColWidth(0) = 800: .ColWidth(1) = 2000: .ColWidth(2) = 2600
        .ColDataType(0) = flexDTBoolean
    End With
      
End Sub

Private Sub vsLista_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsLista_Click()
    
    With vsLista
    If .MouseRow = 0 Then
        .ColSel = .MouseCol
        If .ColSort(.MouseCol) = flexSortGenericAscending Then
            .ColSort(.MouseCol) = flexSortGenericDescending
        Else
            .ColSort(.MouseCol) = flexSortGenericAscending
        End If
        .Sort = flexSortUseColSort
    End If
    End With
    
End Sub

Private Sub vsLista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton And vsLista.MouseCol = 1 And vsLista.MouseRow > 0 Then
        aFiltro = vsLista.Cell(flexcpText, vsLista.MouseRow, 1)
        MnuFiFiltrar.Caption = "Filtrar " & aFiltro
        PopupMenu MnuFiltros
    End If
    
End Sub
