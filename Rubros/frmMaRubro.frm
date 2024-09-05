VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaRubro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rubros Contables"
   ClientHeight    =   5805
   ClientLeft      =   3045
   ClientTop       =   2820
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
   Icon            =   "frmMaRubro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6735
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
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
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   4100
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7011
      _ConvInfo       =   1
      Appearance      =   1
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
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   6495
      Begin VB.CheckBox chExpandir 
         Caption         =   "Expandir subrubros en listados"
         Height          =   255
         Left            =   3000
         MaskColor       =   &H8000000F&
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox tCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox tNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   3
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ID &Rubro:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5550
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "bd"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3757
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
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
            Picture         =   "frmMaRubro.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaRubro.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaRubro.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaRubro.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaRubro.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaRubro.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaRubro.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaRubro.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaRubro.frx":10E2
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
   Begin VB.Menu MnuBases 
      Caption         =   "&Bases"
      Begin VB.Menu MnuBx 
         Caption         =   "BDX"
         Index           =   0
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
Attribute VB_Name = "frmMaRubro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean
Dim gIDRubro As Long


Private Sub chExpandir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tNombre
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error Resume Next

    frmMaRubro.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    frmMaRubro.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    frmMaRubro.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        
    sNuevo = False: sModificar = False
    LimpioFicha
    
    InicializoGrilla
    CargoLista
    
    DeshabilitoIngreso
    
    CentroForm Me
    
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

Private Sub Label10_Click()
    Foco tCodigo
End Sub

Private Sub MnuBx_Click(Index As Integer)

On Error Resume Next

    If Not AccionCambiarBase(MnuBx(Index).Tag, MnuBx(Index).Caption) Then Exit Sub
    Screen.MousePointer = 11
    
    CargoLista
    AccionCancelar
    
    Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
    
    
     'Cambio el Color del fondo de controles ----------------------------------------------------------------------------------------
    Dim arrC() As String
    arrC = Split(MnuBases.Tag, "|")
    If arrC(Index) <> "" Then Me.BackColor = arrC(Index) Else Me.BackColor = vbButtonFace
    
    Frame1.BackColor = Me.BackColor
    chExpandir.BackColor = Me.BackColor
    '-------------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    
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
    gIDRubro = 0
    Botones False, False, False, True, True, Toolbar1, Me
    
    LimpioFicha
    HabilitoIngreso
    Foco tCodigo
  
End Sub

Sub AccionModificar()
    
    On Error Resume Next
    sModificar = True
    
    With vsLista
        gIDRubro = .Cell(flexcpData, .Row, 0)
        tCodigo.Text = .Cell(flexcpText, .Row, 1)
        tNombre.Text = Trim(.Cell(flexcpText, .Row, 2))
        
        If .Cell(flexcpChecked, .Row, 0) = flexChecked Then chExpandir.Value = vbChecked Else chExpandir.Value = vbUnchecked
        
    End With
    
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    Foco tCodigo
        
End Sub

Sub AccionGrabar()

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo errGrabar
    
    If sNuevo Then
        cons = "Select * from Rubro Where RubID = 0"
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        rsAux.AddNew
        CargoCamposBD
        rsAux.Update: rsAux.Close
        
    Else                                    'Modificar----
    
        On Error GoTo errGrabar
        
        cons = "Select * from Rubro Where RubID = " & gIDRubro
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        rsAux.Edit
        CargoCamposBD
        rsAux.Update: rsAux.Close
        
    End If
    
    gIDRubro = 0
    sNuevo = False: sModificar = False
    DeshabilitoIngreso
    LimpioFicha
    CargoLista
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
End Sub

Sub AccionEliminar()

    Screen.MousePointer = 11
    With vsLista
    
    'Valido si hay subrubros
    cons = "Select * from SubRubro Where SRuRubro = " & .Cell(flexcpData, .Row, 0)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        MsgBox "Hay subrubros que pertenecen al rubro seleccionado." & Chr(vbKeyReturn) & "No podrá eliminar el rubro.", vbExclamation, "ATENCIÓN"
        rsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If
    rsAux.Close
    
    If MsgBox("Confirma eliminar el rubro '" & .Cell(flexcpText, .Row, 1) & "'", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Screen.MousePointer = 0: Exit Sub
    On Error GoTo Error
    
    cons = "Select * from Rubro Where RubID = " & .Cell(flexcpData, .Row, 0)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsAux.Delete
    rsAux.Close
    
    LimpioFicha
    DeshabilitoIngreso
    
    CargoLista
    
    End With
    Screen.MousePointer = 0
    Exit Sub
    
Error:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al realizar la operación.", Err.Description
End Sub

Sub AccionCancelar()

    On Error Resume Next
    DeshabilitoIngreso
    LimpioFicha
    
    If vsLista.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me Else Botones True, False, False, False, False, Toolbar1, Me
    
    sNuevo = False: sModificar = False
    
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chExpandir.SetFocus
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
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

    ValidoCampos = False
    
    If Not IsNumeric(tCodigo.Text) Then
        MsgBox "Debe ingresar el código de rubro asignado (campo obligatorio).", vbExclamation, "ATENCIÓN"
        Foco tCodigo: Exit Function
    End If
    
    If Trim(tNombre.Text) = "" Then
        MsgBox "Debe ingresar el nombre del rubro (campo obligatorio).", vbExclamation, "ATENCIÓN"
        Foco tNombre: Exit Function
    End If
    
    'Valido los codigos de rubros ---------------------------------------------------------------------------
    cons = "Select * from Rubro Where RubCodigo = " & tCodigo.Text & " And RubID <> " & gIDRubro
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        MsgBox "Hay un subrubro para el código " & Trim(tCodigo.Text) & Chr(vbKeyReturn) & "No se podrán almacenar los datos.", vbExclamation, "ATENCÍON"
        rsAux.Close: Exit Function
    End If
    rsAux.Close
    
    cons = "Select * from Rubro Where RubNombre = '" & tNombre.Text & "' And RubID <> " & gIDRubro
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        MsgBox "Hay un subrubro para el nombre " & Trim(tNombre.Text) & Chr(vbKeyReturn) & "No se podrán almacenar los datos.", vbExclamation, "ATENCÍON"
        rsAux.Close: Exit Function
    End If
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------------
    
    ValidoCampos = True
    
End Function

Private Sub DeshabilitoIngreso()
        
    tCodigo.Enabled = False: tCodigo.BackColor = Inactivo
    tNombre.Enabled = False: tNombre.BackColor = Inactivo
    chExpandir.Enabled = False
    
    vsLista.Enabled = True: vsLista.BackColor = Blanco
        
End Sub

Private Sub HabilitoIngreso()

    tCodigo.Enabled = True: tCodigo.BackColor = Obligatorio
    tNombre.Enabled = True: tNombre.BackColor = Obligatorio
    chExpandir.Enabled = True
    vsLista.Enabled = False: vsLista.BackColor = Inactivo
    
End Sub

Private Sub CargoCamposBD()
        
    rsAux!RubCodigo = tCodigo.Text
    rsAux!RubNombre = Trim(tNombre.Text)
    If chExpandir.Value = vbChecked Then rsAux!RubExpandir = 1 Else rsAux!RubExpandir = 0
    
End Sub

Private Sub LimpioFicha()

    tCodigo.Text = ""
    tNombre.Text = ""
    chExpandir.Value = vbUnchecked

End Sub


Private Sub InicializoGrilla()

    On Error Resume Next
    With vsLista
        .Cols = 1: .Rows = 1
        .FormatString = "Expandir|<Código Rubro|<Nombre"
            
        .WordWrap = True
        .ColDataType(0) = flexDTBoolean
        .ColWidth(1) = 1300
        .ExtendLastCol = True: .FixedCols = 0
    End With
      
End Sub

Private Sub CargoLista()

    On Error GoTo errCargar
    Dim aValor As Long
    Screen.MousePointer = 11
    Botones True, False, False, False, False, Toolbar1, Me
    cons = "Select * from Rubro Order by RubNombre"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    vsLista.Rows = 1
    
    Do While Not rsAux.EOF
        With vsLista
            .AddItem ""
            If rsAux!RubExpandir Then .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked

            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!RubCodigo)
            aValor = rsAux!RubID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            .Cell(flexcpText, .Rows - 1, 2) = Trim(rsAux!RubNombre)
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    If vsLista.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar la lista.", Err.Description
    Screen.MousePointer = 0
End Sub
