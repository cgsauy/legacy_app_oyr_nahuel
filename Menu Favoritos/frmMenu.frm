VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmMenu 
   Caption         =   "Organizar Favoritos"
   ClientHeight    =   6090
   ClientLeft      =   2370
   ClientTop       =   2925
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8340
   Begin VB.PictureBox picDetalle 
      Height          =   1215
      Left            =   3480
      ScaleHeight     =   1155
      ScaleWidth      =   4275
      TabIndex        =   9
      Top             =   4260
      Width           =   4335
      Begin VB.CommandButton bAddOK 
         Caption         =   "Agregar"
         Height          =   285
         Left            =   3360
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox tAddOrden 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   840
         Width           =   675
      End
      Begin VB.TextBox tAddTitulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   540
         Width           =   3195
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "En el Orden:"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   870
         Width           =   1035
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Con el Título:"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   570
         Width           =   1035
      End
      Begin VB.Label lAddObjeto 
         BackStyle       =   0  'Transparent
         Caption         =   "Agregar en:"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   300
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "El Objeto:"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label lAddFolder 
         BackStyle       =   0  'Transparent
         Caption         =   "Agregar en:"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   60
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agregar en:"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.ComboBox cTipo 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin MSComctlLib.TreeView vsMain 
      Height          =   5475
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   9657
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   617
      LabelEdit       =   1
      Style           =   5
      SingleSel       =   -1  'True
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.StatusBar StatusB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   5835
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14296
            MinWidth        =   2469
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":08CA
            Key             =   "mnuopen"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":307E
            Key             =   "mnuclose"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5832
            Key             =   "root"
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   3315
      Left            =   3480
      TabIndex        =   2
      Top             =   420
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5847
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
      GridLines       =   0
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
      AutoSearch      =   1
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
   Begin VB.Menu MnuTree 
      Caption         =   "MnuTree"
      Visible         =   0   'False
      Begin VB.Menu MnuMTitulo 
         Caption         =   ""
      End
      Begin VB.Menu MnuL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCortar 
         Caption         =   "Cortar"
      End
      Begin VB.Menu MnuPegar 
         Caption         =   "Pegar"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuNCarpeta 
         Caption         =   "Nueva Carpeta"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typObjMenu
    IdMenu As Long
    IdSubDe As Long
    MnuTitulo As String
    MnuAccion As String
    MnuOrden As Long
End Type

Dim objMenu As typObjMenu

Dim rsIni As rdoResultset
Dim aValor As Long

Dim gMaxId As Long

Private Const mnuRoot = "Favoritos"
Private Const prmAccionPlantilla = "appExploreMsg.exe "

Private Sub bAddOK_Click()

On Error GoTo errAdd
Dim aKey As String, aNName As String
Dim bOk As Boolean, miIDMenu As Long

    'Valido los datos a agregar     --------------------------------------------------------------------------------
    If Trim(tAddTitulo.Text) = "" Then
        MsgBox "Ingrese el título del menú a agregar.", vbExclamation, "Falta Título del menú"
        Exit Sub
    End If
    '------------------------------------------------------------------------------------------------------------------
    
    miIDMenu = objMenu.IdMenu
    objMenu.MnuTitulo = Trim(tAddTitulo.Text)
    objMenu.MnuOrden = Val(tAddOrden.Text)
    
    bOk = GrabarMenu(miIDMenu, objMenu.MnuTitulo, objMenu.MnuAccion, objMenu.IdSubDe, objMenu.MnuOrden)
    If Not bOk Then GoTo errAdd: Exit Sub
    
    With vsMain
        If vsMain.Enabled Then          'Nuevo Objeto
            aKey = "N" & miIDMenu
            aNName = Format(objMenu.MnuOrden, "00") & ") " & Trim(objMenu.MnuTitulo)
            
            .Nodes.Add Val(lAddFolder.Tag), tvwChild, aKey, aNName
            
            If objMenu.MnuAccion = "" Then
                .Nodes(aKey).Image = ImageList1.ListImages("mnuclose").Index
                .Nodes(aKey).SelectedImage = ImageList1.ListImages("mnuopen").Index
            End If
            .Nodes(aKey).Tag = miIDMenu
            .Nodes(Val(lAddFolder.Tag)).Expanded = True
            If vsLista.Rows > 1 And vsLista.Enabled Then vsLista.SetFocus Else cTipo.SetFocus
        Else                                     'Modifico Objeto
            vsMain.SelectedItem.Text = Format(objMenu.MnuOrden, "00") & ") " & Trim(objMenu.MnuTitulo)
            CanceloMenu
        End If
        LimpioObjeto
    End With
    
    Exit Sub
errAdd:
    clsGeneral.OcurrioError "Error al agregar el objeto a la lista", Err.Description
End Sub

Private Function GrabarMenu(lIdMenu As Long, sNombre As String, sAccion As String, lSubDe As Long, lOrden As Long) As Boolean
    On Error GoTo errGrabar
    GrabarMenu = False
    
    cons = "Select * from Menu Where MenCodigo = " & lIdMenu        '-------------------------------------------------------
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then rsAux.AddNew Else rsAux.Edit
    
    rsAux!MenUsuario = paCodigoDeUsuario
    rsAux!MenNombre = Trim(sNombre)
    
    If lIdMenu = 0 Then
        If Trim(sAccion) <> "" Then rsAux!MenAccion = Trim(sAccion) Else rsAux!MenAccion = Null
    End If
    
    rsAux!MenOrden = lOrden
    rsAux!MenSubMenuDe = lSubDe
    rsAux.Update: rsAux.Close
    '------------------------------------------------------------------------------------------------------------------------------------
    
    If lIdMenu = 0 Then  '-------------------------------------------------------
        cons = "Select * from Menu " & _
                   " Where MenNombre = '" & Trim(sNombre) & "'" & _
                   " And MenSubMenuDe = " & lSubDe & _
                   " And MenCodigo > " & gMaxId
        
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            lIdMenu = rsAux!MenCodigo
            gMaxId = lIdMenu
        End If
        rsAux.Close
        
    End If
    
    GrabarMenu = True

errGrabar:
End Function

Private Sub LimpioObjeto()
    
    objMenu.IdMenu = 0
    objMenu.MnuAccion = ""
    objMenu.MnuOrden = 0
    objMenu.MnuTitulo = ""
    
    lAddObjeto.Caption = ""
    tAddTitulo.Text = ""
    tAddOrden.Text = ""
    
End Sub

Private Sub bAddOK_GotFocus()
    StatusB.Panels(1).Text = "Agregar menú a favoritos."
End Sub

Private Sub bAddOK_LostFocus()
    StatusB.Panels(1).Text = ""
End Sub

Private Sub cTipo_Click()
    vsLista.Rows = 1
        
    If cTipo.ListIndex = -1 Then Exit Sub
    Select Case cTipo.ListIndex
        Case 0: CargoAplicaciones prmPathApp, "*.exe"
        Case 1: CargoAplicaciones prmPathHelps, "*.htm"
        Case 2: CargoAplicaciones prmPathProc, "*.htm"
        Case 3: CargoPlantillas
    End Select
    
    'If cTipo.ListIndex <> 3 Then vsLista.SetFocus
    
End Sub

Private Sub cTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cTipo.ListIndex = -1 Then Exit Sub
        Select Case cTipo.ListIndex
            Case 0: CargoAplicaciones prmPathApp, "*.exe"
            Case 1: CargoAplicaciones prmPathHelps, "*.htm"
            Case 2: CargoAplicaciones prmPathProc, "*.htm"
            Case 3: CargoPlantillas
            Case 4: CargoOtrosArchivos
        End Select
        
        If cTipo.ListIndex <> 3 Then vsLista.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    vsMain.SetFocus
End Sub

Private Sub Form_Load()

    On Error Resume Next
    ObtengoSeteoForm Me
    LimpioObjeto
    
    InicializoGrillas
    vsMain.ImageList = ImageList1
    
    CargoMenu paCodigoDeUsuario, 0, 0
        
    Call vsMain_NodeClick(vsMain.Nodes("N0"))
    
End Sub
  
Private Sub CargoMenu(lUsr As Long, lIdPadre As Long, lIdxPadre As Long)

Dim aIndex As Long, aKey As String, aNName As String
Dim rsMnu As rdoResultset

    On Error GoTo errMenu
    If lUsr = 0 Then Exit Sub
    
    If lIdxPadre = 0 Then
        vsMain.Nodes.Clear
        vsMain.Nodes.Add , , "N0", mnuRoot
        vsMain.Nodes(vsMain.Nodes("N0").Index).Image = ImageList1.ListImages("root").Index
        vsMain.Nodes("N0").Tag = 0
        CargoMenu lUsr, 0, vsMain.Nodes("N0").Index
    
    Else
    
        cons = " Select * from Menu " & _
                   " Where MenUsuario = " & lUsr & _
                   " And MenSubMenuDe =" & lIdPadre & _
                   " Order by MenOrden"
        Set rsMnu = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        
        Do While Not rsMnu.EOF
            With vsMain
                If rsMnu!MenCodigo > gMaxId Then gMaxId = rsMnu!MenCodigo
                
                aKey = "N" & rsMnu!MenCodigo
                aNName = Trim(rsMnu!MenNombre)
                If Not IsNull(rsMnu!MenOrden) Then aNName = Format(rsMnu!MenOrden, "00") & ") " & aNName Else aNName = "00) " & aNName
                .Nodes.Add lIdxPadre, tvwChild, aKey, aNName
                aIndex = .Nodes(aKey).Index
                
                .Nodes(aIndex).Tag = rsMnu!MenCodigo
                                   
                If IsNull(rsMnu!MenAccion) Then
                    .Nodes(aIndex).Image = ImageList1.ListImages("mnuclose").Index
                    .Nodes(aIndex).SelectedImage = ImageList1.ListImages("mnuopen").Index
                    CargoMenu lUsr, rsMnu!MenCodigo, aIndex
                'Else
                    '.Nodes(aIndex).Image = ImageList1.ListImages("app").Index
                End If
                
             End With
            rsMnu.MoveNext
        Loop
        rsMnu.Close
        
    End If
    Screen.MousePointer = 0
    Exit Sub

errMenu:
    clsGeneral.OcurrioError "Error al cargar el menú favoritos.", Err.Description, "Cargar_Menu"
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    vsMain.Left = Me.ScaleLeft
    vsMain.Top = Me.ScaleTop + 20
    vsMain.Height = Me.ScaleHeight - (StatusB.Height + vsMain.Top)
    vsMain.Width = Me.ScaleWidth - picDetalle.Width - 100
              
    With picDetalle
        .Left = vsMain.Left + vsMain.Width + 60
        .Top = Me.ScaleTop + (Me.ScaleHeight - (.Height + StatusB.Height))
        .BorderStyle = 0
    End With
    
    cTipo.Left = picDetalle.Left
    cTipo.Top = vsMain.Top
    cTipo.Width = picDetalle.Width
    
    vsLista.Left = picDetalle.Left
    vsLista.Height = Me.ScaleHeight - picDetalle.Height - StatusB.Height - vsLista.Top
    vsLista.Width = picDetalle.Width
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub


Private Sub InicializoGrillas()

    On Error Resume Next
    With vsLista
        .Cols = 1: .Rows = 1
        .FormatString = "<Codigo|<Accion|<Lista de objetos a agregar"
            
        .ColWidth(0) = 0: .ColWidth(1) = 0: .ColWidth(2) = 1800
        .ColHidden(0) = True: .ColHidden(1) = True
        
        .ExtendLastCol = True: .FixedCols = 0
    End With
      
    cTipo.AddItem "Aplicaciones"
    cTipo.AddItem "Archivos de Ayuda"
    cTipo.AddItem "Ayuda de Procedimientos"
    cTipo.AddItem "Plantillas"
    cTipo.AddItem "Otros Archivos ..."
    
    

End Sub

Private Sub CargoPlantillas()
On Error GoTo errCP
    vsLista.Rows = 1
    Screen.MousePointer = 11
    
    StatusB.Panels(1).Text = "Cargando Plantillas."
    
    cons = "Select PlaCodigo, PlaNombre From Plantilla" & _
               " Where PlaTipo Not In (2, 4, 5) " & _
               " Order by PlaNombre"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        With vsLista
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = rsAux!PlaCodigo
            .Cell(flexcpText, .Rows - 1, 1) = prmAccionPlantilla & rsAux!PlaCodigo & ":"
            .Cell(flexcpText, .Rows - 1, 2) = Trim(rsAux!PlaNombre)
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    StatusB.Panels(1).Text = ""
    Screen.MousePointer = 0
    Exit Sub
    
errCP:
    clsGeneral.OcurrioError "Error al cargar las plantillas de datos.", Err.Description
End Sub

Private Sub CargoAplicaciones(sPath As String, sFiles As String)
    On Error GoTo errCargo
    Screen.MousePointer = 11
    StatusB.Panels(1).Text = "Cargando aplicaciones desde " & sPath
    StatusB.Refresh
    vsLista.Rows = 1
    Dim myFile As String
        
    myFile = Dir(sPath & "\" & sFiles)
    
    Do While myFile <> ""
        With vsLista
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = 0
            .Cell(flexcpText, .Rows - 1, 1) = sPath & "\" & Trim(myFile)
            .Cell(flexcpText, .Rows - 1, 2) = Trim(myFile)
        End With
        myFile = Dir()
    Loop
    
    If vsLista.Rows > 1 Then
        vsLista.Select 1, 2
        vsLista.Sort = flexSortGenericAscending
    End If
    StatusB.Panels(1).Text = ""
    Screen.MousePointer = 0
    
    Exit Sub
errCargo:
    clsGeneral.OcurrioError "Error al cargar los archivos.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub MnuCortar_Click()

    estadoCopiar False
    
End Sub

Private Sub MnuNCarpeta_Click()
    objMenu.MnuAccion = ""
    objMenu.IdMenu = 0
    lAddObjeto.Caption = "Nueva Carpeta"
    tAddTitulo.Text = Trim(lAddObjeto.Caption)
    tAddTitulo.SetFocus
End Sub

Private Sub MnuPegar_Click()
    On Error GoTo errMover
    
    Screen.MousePointer = 11
    Dim aIdxExpand As Integer
    aIdxExpand = vsMain.SelectedItem.Index
    
    If Not MoverMenu(Val(MnuMTitulo.Tag), objMenu.IdSubDe) Then GoTo errMover: Exit Sub
    
    estadoCopiar True
    
    CargoMenu paCodigoDeUsuario, 0, 0
    vsMain.Nodes("N0").Expanded = True
    vsMain.Nodes(aIdxExpand).Expanded = True
    
    Set vsMain.SelectedItem = vsMain.Nodes(aIdxExpand)
    Screen.MousePointer = 0
    Exit Sub
    
errMover:
    clsGeneral.OcurrioError "Error al mover el menú seleccionado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tAddOrden_GotFocus()
    StatusB.Panels(1).Text = "Orden de aparición del menú."
    If Not vsMain.Enabled Then StatusB.Panels(1).Text = "[Esc]- Cancelar.    " & StatusB.Panels(1).Text
End Sub

Private Sub tAddOrden_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn: bAddOK.SetFocus
        Case vbKeyEscape: If Not vsMain.Enabled Then CanceloMenu
    End Select
    
End Sub

Private Sub tAddOrden_LostFocus()
    StatusB.Panels(1).Text = ""
End Sub

Private Sub tAddTitulo_GotFocus()
    StatusB.Panels(1).Text = "Título del menú."
    If Not vsMain.Enabled Then StatusB.Panels(1).Text = "[Esc]- Cancelar.    " & StatusB.Panels(1).Text
    
    tAddTitulo.SelStart = 0: tAddTitulo.SelLength = Len(tAddTitulo)
End Sub

Private Sub tAddTitulo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn: If Trim(tAddTitulo.Text) <> "" Then tAddOrden.SetFocus
        Case vbKeyEscape: If Not vsMain.Enabled Then CanceloMenu
    End Select
    
End Sub

Private Sub tAddTitulo_LostFocus()
    StatusB.Panels(1).Text = ""
End Sub

Private Sub vsLista_DblClick()
    On Error GoTo errDC
    If vsLista.Rows = 1 Then Exit Sub
    
    If cTipo.ListIndex = -1 Then Exit Sub
    Select Case cTipo.ListIndex
        Case 0: EjecutarApp vsLista.Cell(flexcpText, vsLista.Row, 1)
        Case 1: EjecutarApp vsLista.Cell(flexcpText, vsLista.Row, 1)
        Case 2: EjecutarApp vsLista.Cell(flexcpText, vsLista.Row, 1)
        
        Case 3:
                Dim miObj As New clsPlantillaI
                miObj.ProcesoPlantillaInteractiva cBase, vsLista.Cell(flexcpText, vsLista.Row, 0), 0, "", "", bPreview:=True
                Set miObj = Nothing
                
    End Select
    
errDC:
End Sub

Private Sub vsLista_GotFocus()
    StatusB.Panels(1).Text = "Elementos a agregar al menú favoritos.  (doble clik para activar el elemento)."
End Sub

Private Sub vsLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then tAddTitulo.SetFocus
End Sub

Private Sub vsLista_LostFocus()
    StatusB.Panels(1).Text = ""
End Sub

Private Sub vsLista_SelChange()
    
    With vsLista
        lAddObjeto.Caption = Trim(.Cell(flexcpText, .Row, 2))
        lAddObjeto.Tag = Trim(.Cell(flexcpText, .Row, 0))
        
        tAddTitulo.Text = Trim(lAddObjeto.Caption)
        tAddOrden = ""
        
        objMenu.MnuAccion = Trim(.Cell(flexcpText, .Row, 1))
        
    End With
End Sub

Private Sub vsMain_GotFocus()
    StatusB.Panels(1).Text = "[Espacio]- Modificar.      [Supr]- Eliminar."
End Sub

Private Sub vsMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errKD
    
    Select Case KeyCode
        Case vbKeyDelete
                        If MnuPegar.Enabled Then Exit Sub
                        With vsMain.SelectedItem
                            If .Children = 0 Then EliminarMenu Val(.Tag)
                        End With
                
        Case vbKeySpace:
                        If MnuPegar.Enabled Then Exit Sub
                        If Val(vsMain.SelectedItem.Tag) > 0 Then ModificarMenu Val(vsMain.SelectedItem.Tag)
              
        Case vbKeyEscape: If MnuPegar.Enabled Then estadoCopiar True
                        
    End Select
    
errKD:
End Sub

Private Sub vsMain_LostFocus()
    StatusB.Panels(1).Text = ""
End Sub

Private Sub vsMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        DoEvents
        With vsMain.SelectedItem
            If Val(.Tag) > 0 And MnuCortar.Enabled Then
                MnuMTitulo.Caption = Trim(.Text)
                MnuMTitulo.Tag = Val(.Tag)
                PopupMenu MnuTree
            Else
                If MnuPegar.Enabled Then
                    PopupMenu MnuTree
                Else
                    MnuMTitulo.Caption = "Favoritos"
                    MnuCortar.Enabled = False
                    PopupMenu MnuTree
                    MnuCortar.Enabled = True
                End If
            End If
        End With
    End If

End Sub

Private Sub vsMain_NodeClick(ByVal Node As MSComctlLib.Node)

    If Node.Image Then
        lAddFolder.Caption = Node.Text
        lAddFolder.Tag = Node.Index
        objMenu.IdSubDe = Val(Node.Tag)
    Else
        lAddFolder.Caption = Node.Parent.Text
        lAddFolder.Tag = Node.Parent.Index
        objMenu.IdSubDe = Val(Node.Parent.Tag)
    End If
End Sub

Private Sub ModificarMenu(lIdMenu As Long)
On Error Resume Next
    
    estadoModificar False
    lAddObjeto.Caption = vsMain.SelectedItem.Text
    
    Dim aPos As Integer
    aPos = InStr(lAddObjeto.Caption, ")")
    tAddTitulo.Text = Trim(Mid(lAddObjeto.Caption, aPos + 1))
    tAddOrden.Text = Trim(Mid(lAddObjeto.Caption, 1, aPos - 1))
    
    objMenu.IdMenu = lIdMenu
    objMenu.IdSubDe = Val(vsMain.SelectedItem.Parent.Tag)
    tAddTitulo.SetFocus
    
End Sub

Private Sub CanceloMenu()
On Error Resume Next

    estadoModificar True
    
    LimpioObjeto
    vsMain.SetFocus
    
End Sub

Private Sub EliminarMenu(lIdMenu As Long)
    
    If lIdMenu = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    cons = "Select * from Menu Where MenCodigo = " & lIdMenu
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then rsAux.Delete
    rsAux.Close
    
    vsMain.Nodes.Remove (vsMain.SelectedItem.Index)
    Screen.MousePointer = 0
    Exit Sub
    
errEliminar:
    clsGeneral.OcurrioError "Error al eliminar el menú.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function MoverMenu(lIdMenuAMover As Long, lIdMenuPadre As Long) As Boolean
    On Error GoTo errMover
    MoverMenu = False
    
    cons = "Select * from Menu Where MenCodigo = " & lIdMenuAMover
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        rsAux.Edit
        rsAux!MenSubMenuDe = lIdMenuPadre
        rsAux.Update
    End If
    rsAux.Close
    
    MoverMenu = True
errMover:
End Function

Private Sub estadoCopiar(bEstado As Boolean)
Dim bkColor As Long

    MnuCortar.Enabled = bEstado
    MnuPegar.Enabled = Not bEstado
    MnuNCarpeta.Enabled = bEstado
    
    If bEstado Then bkColor = vbWindowBackground Else bkColor = vbInactiveBorder
    
    cTipo.Enabled = bEstado: cTipo.BackColor = bkColor
    vsLista.Enabled = bEstado: vsLista.BackColor = bkColor
    
    picDetalle.Enabled = bEstado
    
    tAddTitulo.Enabled = bEstado: tAddTitulo.BackColor = bkColor
    tAddOrden.Enabled = bEstado: tAddOrden.BackColor = bkColor
    
    If Not bEstado Then
        StatusB.Panels(1).Text = "Moviendo menú ... [Esc]- Cancelar"
    Else
        StatusB.Panels(1).Text = "[Espacio]- Modificar.      [Supr]- Eliminar."
    End If
    
End Sub

Private Sub estadoModificar(bEstado As Boolean)
Dim bkColor As Long

    If bEstado Then bkColor = vbWindowBackground Else bkColor = vbInactiveBorder

    cTipo.Enabled = bEstado: cTipo.BackColor = bkColor
    vsLista.Enabled = bEstado: vsLista.BackColor = bkColor
    
    vsMain.Enabled = bEstado: vsMain.Refresh

End Sub


Private Sub CargoOtrosArchivos()
    On Error GoTo errCargoOtros
    Dim myFile As String
    
    Dim FileDialog As CFileDialog
    Set FileDialog = New CFileDialog
    
    myFile = ""
    
    With FileDialog
        .DefaultExt = "*"
        .DialogTitle = "Selección de archivos para Importar"
        .Filter = "Todos los archivos|*.*"
        .FilterIndex = 0
        .Flags = FleFileMustExist + FleHideReadOnly + FleCreatePrompt + FleExplorer '+ FleAllowMultiSelect + FleLongnames
        .hWndParent = Me.hwnd
        .MaxFileSize = 255
        
        If .Show(True) Then myFile = .FileName
    End With
    Set FileDialog = Nothing
    
    If myFile = "" Then Exit Sub
    
    lAddObjeto.Caption = Trim(myFile)
    lAddObjeto.Tag = 0
    tAddTitulo.Text = Trim(lAddObjeto.Caption)
    tAddOrden = ""
    objMenu.MnuAccion = Trim(myFile)
    
    tAddTitulo.SetFocus
    Exit Sub
    
errCargoOtros:
    clsGeneral.OcurrioError "Error al procesar el archivo a agregar.", Err.Description
    Screen.MousePointer = 0
End Sub
