VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quotation Items"
   ClientHeight    =   5850
   ClientLeft      =   3375
   ClientTop       =   2580
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmItems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6420
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      ForeColor       =   &H00000080&
      Height          =   1695
      Left            =   60
      TabIndex        =   11
      Top             =   480
      Width           =   6315
      Begin VB.TextBox tSRubro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1020
         TabIndex        =   5
         Top             =   975
         Width           =   4155
      End
      Begin VB.TextBox tRubro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4155
      End
      Begin VB.TextBox tNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1020
         MaxLength       =   60
         TabIndex        =   3
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label lID 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   12
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Sub Rubro:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Rubro:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Descripción:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   645
         Width           =   855
      End
   End
   Begin VB.TextBox tBuscar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      MaxLength       =   40
      TabIndex        =   1
      Top             =   2280
      Width           =   4755
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   3195
      Left            =   60
      TabIndex        =   8
      Top             =   2580
      Width           =   6315
      _ExtentX        =   11139
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
      AllowUserResizing=   1
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4740
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":0442
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":0DC8
            Key             =   "modificar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":10E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Buscar:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2295
      Width           =   735
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
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean
Dim gIdItem As Long

Public prmIDItem As Long

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error Resume Next

    sNuevo = False: sModificar = False
    LimpioFicha
    InicializoGrilla
       
    HabilitoIngreso Estado:=False
    
    If prmIDItem <> 0 Then
        CargoDatos prmIDItem
        If gIdItem <> 0 Then AccionModificar
    Else
        Foco tBuscar
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco tNombre
End Sub

Private Sub Label3_Click()
    Foco tSRubro
End Sub

Private Sub Label4_Click()
    Foco tBuscar
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

Private Sub AccionNuevo()
    On Error Resume Next
    
    sNuevo = True
    gIdItem = 0
    Botones False, False, False, True, True, Toolbar1, Me
    
    LimpioFicha
    HabilitoIngreso
    Foco tNombre
  
End Sub

Private Sub AccionModificar()
    
    On Error GoTo errModificar
    Screen.MousePointer = 11
    sModificar = True: sNuevo = False
    
    If Val(lID.Caption) <> 0 Then gIdItem = Val(lID.Caption) Else gIdItem = vsLista.Cell(flexcpData, vsLista.Row, 0)
        
    CargoDatos gIdItem
    
    If gIdItem = 0 Then
        AccionCancelar
        Exit Sub
    End If
    
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    Foco tNombre
    Screen.MousePointer = 0
    Exit Sub
    
errModificar:
    clsGeneral.OcurrioError "Error al cargar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatos(aIdItem As Long)
On Error GoTo errDatos

    Screen.MousePointer = 11
    LimpioFicha
    
    cons = "Select * from Items, SubRubro, Rubro " & _
                " Where IteId = " & aIdItem & _
                " And IteSubRubro = SRuId " & _
                " And SRuRubro = RubId"

    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsAux.EOF Then
        gIdItem = rsAux!IteID
        lID.Caption = gIdItem
        
        tNombre.Text = Trim(rsAux!IteNombre)
        tSRubro.Text = Format(rsAux!SRuCodigo, "(0)") & " " & Trim(rsAux!SRuNombre)
        tSRubro.Tag = rsAux!SRuId
        
        tRubro.Text = Format(rsAux!RubCodigo, "(0)") & " " & Trim(rsAux!RubNombre)
        tRubro.Tag = rsAux!RubID
        
    Else
        gIdItem = 0
    End If
    rsAux.Close
    
    If gIdItem = 0 Then
        MsgBox "Posiblemente el registro ha sido eliminado." & vbCrLf & "Vuelva a consultar.", vbExclamation, "Registro Inexistente"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
errDatos:
    clsGeneral.OcurrioError "Error al cargar los datos del item.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionGrabar()

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "Grabar Datos") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo errGrabar
    
    cons = "Select * from Items Where IteID = " & gIdItem
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If sNuevo Then rsAux.AddNew Else rsAux.Edit

    rsAux!IteNombre = Trim(tNombre.Text)
    rsAux!IteSubRubro = Val(tSRubro.Tag)
    
    rsAux.Update: rsAux.Close
    
    gIdItem = 0
    sNuevo = False: sModificar = False
    HabilitoIngreso False
    LimpioFicha
    If vsLista.Rows > 1 Then
        If vsLista.Enabled Then vsLista.SetFocus
        Botones True, True, True, False, False, Toolbar1, Me
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    clsGeneral.OcurrioError "Error al grabar los datos ingresados.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionEliminar()

    Screen.MousePointer = 11
    On Error GoTo Error
    
    gIdItem = vsLista.Cell(flexcpData, vsLista.Row, 0)
        
    Dim bHay As Boolean: bHay = False
    cons = "Select Top 1 * from QuotationItem Where QItItem = " & gIdItem
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True
    rsAux.Close
    
    If bHay Then
        MsgBox "Hay cotizaciones ingresadas con el ítem que ud. quiere eliminar.", vbExclamation, "Hay Cotizaciones con el ítem"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    If MsgBox("Confirma eliminar el ítem '" & vsLista.Cell(flexcpText, vsLista.Row, 1) & "'", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Screen.MousePointer = 0: Exit Sub
    
    cons = "Select * from Items Where IteID = " & gIdItem
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsAux.Delete: rsAux.Close
    
    Screen.MousePointer = 0
    Exit Sub
    
Error:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al eliminar el item seleccionado.", Err.Description
End Sub

Sub AccionCancelar()

    On Error Resume Next
    HabilitoIngreso Estado:=False
    LimpioFicha
    
    If vsLista.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me Else Botones True, False, False, False, False, Toolbar1, Me
    
    sNuevo = False: sModificar = False
    Foco tBuscar
    
End Sub


Private Sub tBuscar_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = vbKeyReturn Then
        If sNuevo Or sModificar Then Exit Sub
        If Trim(tBuscar.Text) = "" Then Exit Sub
        
        On Error GoTo errAyuda
        Dim aValor As Long, sBuscar As String
        Screen.MousePointer = 11
        
        sBuscar = Replace(Trim(tBuscar.Text), " ", "%")
        
        cons = "Select * from Items, SubRubro, Rubro " & _
                    " Where IteNombre like '" & sBuscar & "%'" & _
                    " And IteSubRubro = SRuId " & _
                    " And SRuRubro = RubId"
        
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        vsLista.Rows = 1
        Do While Not rsAux.EOF
            With vsLista
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Trim(rsAux!IteNombre)
                aValor = rsAux!IteID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                
                .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!SRuNombre) & " " & Format(rsAux!SRuCodigo, "(0)")
                .Cell(flexcpText, .Rows - 1, 2) = Trim(rsAux!RubNombre) & " " & Format(rsAux!RubCodigo, "(0)")
                
            End With
            rsAux.MoveNext
        Loop
        rsAux.Close
        
        If vsLista.Rows > 1 Then
            Botones True, True, True, False, False, Toolbar1, Me
            vsLista.SetFocus
            If vsLista.Rows = 2 Then CargoDatos vsLista.Cell(flexcpData, vsLista.Rows - 1, 0)
        Else
            Botones True, False, False, False, False, Toolbar1, Me
        End If
        
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errAyuda:
    clsGeneral.OcurrioError "Error al realizar la búsqueda.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub tNombre_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tNombre.Text) = "" Then Exit Sub
        Foco tSRubro
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
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
On Error GoTo errValidar
    ValidoCampos = False
    
    If Trim(tNombre.Text) = "" Then
        MsgBox "Ingrese la identificación o descripción para el item.", vbExclamation, "Falta Nombre"
        Foco tNombre: Exit Function
    End If
    
    If Val(tSRubro.Tag) = 0 Then
        MsgBox "Seleccione el subrubro asociado al item.", vbExclamation, "Falta seleccionar Subrubro"
        Foco tSRubro: Exit Function
    End If
    
    'Valido Nombre o --------------------------------------------------------------
    cons = "Select * From Items" & _
               " Where IteNombre = '" & Trim(tNombre.Text) & "'" & _
               " And IteID <> " & gIdItem
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        MsgBox "Ya existe una item con la indentificación '" & Trim(tNombre.Text) & "'.", vbExclamation, "Posible Duplicación"
        rsAux.Close: Exit Function
    End If
    rsAux.Close
    '--------------------------------------------------------------------------------------------
    
    ValidoCampos = True
    Exit Function
    
errValidar:
    clsGeneral.OcurrioError "Error al validar los campos.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub HabilitoIngreso(Optional Estado As Boolean = True)

Dim bkColor As Long
    
    If Estado Then bkColor = Colores.Blanco Else bkColor = Colores.Gris
    
    tNombre.Enabled = Estado: tNombre.BackColor = bkColor
    tRubro.Enabled = Estado: tRubro.BackColor = bkColor
    tSRubro.Enabled = Estado: tSRubro.BackColor = bkColor
    
    tBuscar.Enabled = Not Estado
    vsLista.Enabled = Not Estado
    
End Sub

Private Sub LimpioFicha()

    lID.Caption = ""
    tNombre.Text = ""
    tRubro.Text = ""
    tSRubro.Text = ""
    
End Sub


Private Sub InicializoGrilla()

    On Error Resume Next
    With vsLista
        .Cols = 1: .Rows = 1
        .FormatString = "<Item|<Sub Rubro|<Rubro"
            
        .WordWrap = True
        .ColWidth(0) = 2500: .ColWidth(1) = 2000
        
        .ExtendLastCol = True: .FixedCols = 0
    End With
      
End Sub



Private Sub tSRubro_Change()
    If Val(tSRubro.Tag) <> 0 Then tSRubro.Tag = 0
End Sub

Private Sub tSRubro_GotFocus()
    tSRubro.SelStart = 0: tSRubro.SelLength = Len(tSRubro.Text)
End Sub

Private Sub tSRubro_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tSRubro.Text) = "" Then Exit Sub
        
        If Val(tSRubro.Tag) <> 0 Then
            AccionGrabar
            Exit Sub
        End If
        
        ProcesoSubRubro Trim(tSRubro.Text)
        
        If Val(tSRubro.Tag) <> 0 Then AccionGrabar
    End If
    
End Sub

Private Sub ProcesoSubRubro(sTexto As String)

    On Error GoTo errBuscar
    Screen.MousePointer = 11
    
    cons = "Select SRuId, RubId, SRuCodigo as 'Subrubro', SRuNombre as 'Nombre de Subrubro', RubCodigo as 'Rubro', RubNombre as 'Nombre del Rubro'" & _
               " from SubRubro, Rubro Where SRuRubro = RubId "
                   
    If IsNumeric(sTexto) Then
        cons = cons & " And SRuCodigo = " & Val(sTexto)
    Else
        sTexto = Replace(sTexto, " ", "%")
        cons = cons & " And SRuNombre like '" & Trim(sTexto) & "%'"
    End If
    
    Dim aQ As Integer, aIDSub As Long, aIDRubro As Long, aNSub As String, aNRubro As String
    aQ = 0
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        aQ = 1
        aIDSub = rsAux!SRuId: aIDRubro = rsAux!RubID
        aNSub = Format(rsAux(2), "(0)") & " " & Trim(rsAux(3))
        aNRubro = Format(rsAux(4), "(0)") & " " & Trim(rsAux(5))
        
        rsAux.MoveNext
        If Not rsAux.EOF Then aQ = 2: aIDSub = 0
    End If
    rsAux.Close
    
    Select Case aQ
        Case 0: MsgBox "No hay datos para la búsqueda ingresada.", vbExclamation, "No hay Datos"
        
        Case 2:
                    Dim aValor As Long
                    Dim miLista As New clsListadeAyuda
                    aValor = miLista.ActivarAyuda(cBase, cons, 6800, 2, "Lista de Subrubros")
                    If aValor > 0 Then
                        aIDSub = miLista.RetornoDatoSeleccionado(0)
                        aIDRubro = miLista.RetornoDatoSeleccionado(1)
                        aNSub = Format(miLista.RetornoDatoSeleccionado(2), "(0)") & " " & Trim(miLista.RetornoDatoSeleccionado(3))
                        aNRubro = Format(miLista.RetornoDatoSeleccionado(4), "(0)") & " " & Trim(miLista.RetornoDatoSeleccionado(5))
                    End If
                    Set miLista = Nothing
    End Select
    
    If aIDSub > 0 Then
        tSRubro.Text = aNSub: tSRubro.Tag = aIDSub
        tRubro.Text = aNRubro: tRubro.Tag = aIDRubro
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    clsGeneral.OcurrioError "Error al buscar los subrubros.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub vsLista_Click()
    On Error Resume Next
    If vsLista.Rows > 1 Then
        
        With vsLista
            If .MouseRow = 0 Then
                .ColSel = .MouseCol
                If .ColSort(.MouseCol) = flexSortGenericAscending Then .ColSort(.MouseCol) = flexSortGenericDescending Else .ColSort(.MouseCol) = flexSortGenericAscending
                .Sort = flexSortUseColSort
                Exit Sub
            End If
        End With
        If gIdItem <> vsLista.Cell(flexcpData, vsLista.Row, 0) Then CargoDatos vsLista.Cell(flexcpData, vsLista.Row, 0)
        
    End If
End Sub

Private Sub vsLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsLista.Rows = 1 Or tBuscar.Enabled = False Then Exit Sub
    On Error Resume Next
        
    If KeyCode = vbKeyReturn Then
        CargoDatos vsLista.Cell(flexcpData, vsLista.Row, 0)
        Exit Sub
    End If
    
    If (KeyCode >= vbKeyA And KeyCode <= vbKeyZ) And Shift = 0 Then
        tBuscar.Text = LCase(Chr(KeyCode))
        tBuscar.SetFocus: tBuscar.SelStart = Len(tBuscar.Text)
    End If
        
End Sub
