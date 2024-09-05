VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmABM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos de artículos"
   ClientHeight    =   5925
   ClientLeft      =   2895
   ClientTop       =   2235
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
   Icon            =   "frmABM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6420
   Begin VB.Timer tmrInicio 
      Left            =   3240
      Top             =   960
   End
   Begin VB.CheckBox chkAccesorios 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Acc&esorios"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   25
      TabIndex        =   1
      Top             =   600
      Width           =   4575
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5670
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":28B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":29C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":2AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":2BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":2CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":3016
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":3128
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":3442
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":375C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrid 
      Height          =   3735
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6588
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
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
   Begin VB.Label lbTitulo 
      BackColor       =   &H8000000C&
      Caption         =   "Grupos ingresados"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   855
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
Attribute VB_Name = "frmABM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Botones(Nu As Boolean, MoEl As Boolean, Gr As Boolean)

    'Habilito y Desabilito Botones.
    Toolbar1.Buttons("nuevo").Enabled = Nu
    MnuNuevo.Enabled = Nu
    
    Toolbar1.Buttons("modificar").Enabled = MoEl
    MnuModificar.Enabled = MoEl
    
    Toolbar1.Buttons("eliminar").Enabled = MoEl
    MnuEliminar.Enabled = MoEl
    
    Toolbar1.Buttons("grabar").Enabled = Gr
    MnuGrabar.Enabled = Gr
    
    Toolbar1.Buttons("cancelar").Enabled = Gr
    MnuCancelar.Enabled = Gr

End Sub

Private Sub CargarGrupos()
On Error GoTo errCG
    Screen.MousePointer = 11
    Dim sQy As String
    Dim rsG As rdoResultset
    vsGrid.Rows = 1
    vsGrid.Redraw = False
    sQy = "SELECT GruCodigo, rTrim(GruNombre), IsNull(GruAccesorios, 0) From Grupo Order by GruNombre"
    Set rsG = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    Do While Not rsG.EOF
        With vsGrid
            .AddItem rsG(1)
            .Cell(flexcpText, .Rows - 1, 1) = rsG(2)
            .Cell(flexcpData, .Rows - 1, 0) = CStr(rsG(0))
        End With
        rsG.MoveNext
    Loop
    rsG.Close
    vsGrid.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
errCG:
    Screen.MousePointer = 0
    vsGrid.Redraw = True
    clsGeneral.OcurrioError "Error al cargar los grupos en la grilla."
End Sub

Private Sub chkAccesorios_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then loc_Grabar
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad

'obtengo de la registry la última posición del formulario.
    Dim ofncs As New clsFunciones
    ofncs.GetPositionForm Me
    Set ofncs = Nothing
    With vsGrid
        .Rows = 1
        .FixedCols = 0: .Cols = 1: .ExtendLastCol = True
        .BackColor = vbWhite: .SheetBorder = vbWhite
        .BackColorBkg = vbWhite
        .FormatString = "Grupo|^Accesorios"
        .ColWidth(0) = 3500
        '.ColWidth(1) = 1600
        .ColDataType(1) = flexDTBoolean
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByRow
    End With
'Inicializo los ctrls
    loc_SetCtrl False
    Botones True, False, False
    tmrInicio.Interval = 10
    Screen.MousePointer = 0
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al ingresar al formulario."
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    vsGrid.Move vsGrid.Left, vsGrid.Top, ScaleWidth - (vsGrid.Left * 2), ScaleHeight - (vsGrid.Top + Status.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'Guardamos la posición del formulario.
    Dim ofncs As New clsFunciones
    ofncs.SetPositionForm (Me)
    Set ofncs = Nothing
'Cerramos la conexión.
    cBase.Close
'eliminamos la referencia de orcgsa.
    Set clsGeneral = Nothing
    End
    Exit Sub
End Sub

Private Sub MnuCancelar_Click()
    loc_CancelarEdicion
End Sub

Private Sub MnuEliminar_Click()
    loc_Eliminar
End Sub

Private Sub MnuGrabar_Click()
    loc_Grabar
End Sub

Private Sub MnuModificar_Click()
    loc_Edicion
End Sub

Private Sub MnuNuevo_Click()
    loc_Nuevo
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Sub loc_Nuevo()
On Error GoTo ErrAN
    Screen.MousePointer = 11
    
'Prendo Señal que es uno nuevo.
    Toolbar1.Tag = 1
    
'Habilito y Desabilito Botones.
    Botones False, False, True

'Limpiamos los controles para el ingreso.
    loc_CleanCtrl

'Seteamos el estado de c/control para la edición.
    loc_SetCtrl True
    
'Posicionamos en el primer control para el ingreso.
    txtNombre.Tag = 0
    txtNombre.SetFocus
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrAN:
    clsGeneral.OcurrioError "Error en acción nuevo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_Edicion()
    
'Prendo señal que es modificación.
    Toolbar1.Tag = 2
'Habilito y Desabilito Botones y controles.
    loc_SetCtrl True
    Botones False, False, True
    
'Me posiciono en el primer elemento a editar.
    txtNombre.Tag = vsGrid.Cell(flexcpData, vsGrid.Row, 0)
    txtNombre.Text = vsGrid.Cell(flexcpText, vsGrid.Row, 0)
    chkAccesorios.Value = vsGrid.Cell(flexcpValue, vsGrid.Row, 1)
    txtNombre.SetFocus
    Screen.MousePointer = 0

End Sub

Private Sub loc_Grabar()
Dim sRespuesta As String

'Hacemos los controles de datos ingresados y de validación antes de grabar
    If Not fnc_ValidateSave Then Exit Sub
        
    If MsgBox("Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        Screen.MousePointer = 11
        On Error GoTo ErrSave
        Dim Cons As String
        Dim rsaux As rdoResultset
        Cons = "Select * From Grupo Where GruCodigo = " & Val(txtNombre.Tag)
        Set rsaux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        'Si tengo la señal de nuevo
        If Val(Toolbar1.Tag) = 1 Then
            rsaux.AddNew
        Else
            rsaux.Edit
        End If
        rsaux("GruNombre") = Trim(txtNombre.Text)
        rsaux("GruAccesorios") = IIf(chkAccesorios.Value, 1, 0)
        rsaux.Update
        rsaux.Close
        
    'Invocamos a cancelar p/volver a estado de no edición
        loc_CancelarEdicion
    End If
    Exit Sub
    

ErrSave:
    clsGeneral.OcurrioError "No se pudo almacenar la información, reintente.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub loc_Eliminar()
On Error GoTo errDel
    'Verificar si hay datos a validar.
    If MsgBox("Confrima eliminar el grupo seleccionado?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
        Screen.MousePointer = 11
        Dim Cons As String
        Dim rsaux As rdoResultset
        Cons = "Select * From Grupo Where GruCodigo = " & Val(vsGrid.Cell(flexcpData, vsGrid.Row, 0))
        Set rsaux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsaux.EOF Then rsaux.Delete
        rsaux.Close
        'Limpiamos los controles y ponemos el formulario en su nuevo estado.
        loc_CleanCtrl
        Botones True, False, False
        CargarGrupos
        If vsGrid.Rows > 1 Then vsGrid.Select 1, 0
        Screen.MousePointer = 0
    End If
    Exit Sub
errDel:
    clsGeneral.OcurrioError "No se pudo eliminar el registro.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_CancelarEdicion()
    

    Screen.MousePointer = 11
    loc_SetCtrl False
'Si es edición cargamos los valores si no limpiamos la ficha.
    
    Dim sNombre As String
    sNombre = Trim(txtNombre.Text)
    
    loc_CleanCtrl
    CargarGrupos
    Botones True, False, False
    
    Dim intRow As Integer
    For intRow = 1 To vsGrid.Rows - 1
        If vsGrid.Cell(flexcpText, intRow, 0) = sNombre Then
            vsGrid.Select intRow, 0
            Botones True, True, False
            Exit For
        End If
    Next
    
    If vsGrid.Rows > 1 Then vsGrid.SetFocus
    
    'Elimino señal de edición.
    Toolbar1.Tag = 0
    Screen.MousePointer = 0
    
End Sub

Private Sub tmrInicio_Timer()
    tmrInicio.Enabled = False
    CargarGrupos
    If vsGrid.Rows > 1 Then
        vsGrid.Select 1, 0
        Botones True, True, False
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": loc_Nuevo
        Case "modificar": loc_Edicion
        Case "eliminar": loc_Eliminar
        Case "grabar": loc_Grabar
        Case "cancelar": loc_CancelarEdicion
        Case "salir": Unload Me
    End Select

End Sub

Private Function fnc_ValidateSave() As Boolean
'Arrancamos la función en false (valor x defecto) para salir al primer error que encontremos.
    fnc_ValidateSave = False
    If Trim(txtNombre.Text) = "" Then
        MsgBox "El nombre es obligatorio", vbExclamation, "Validación"
        txtNombre.SetFocus
        Exit Function
    End If
    fnc_ValidateSave = True
End Function

Private Sub loc_SetCtrl(ByVal bEdit As Boolean)
'Rutina para habilitar/deshabilitar los controles
   txtNombre.Enabled = bEdit
   txtNombre.BackColor = IIf(bEdit, vbWindowBackground, vbButtonFace)
   chkAccesorios.Enabled = bEdit
   vsGrid.Enabled = Not bEdit
End Sub

Private Sub loc_CleanCtrl()
'limpiamos los controles con sus valores x defecto.
    txtNombre.Text = ""
    chkAccesorios.Value = 0
End Sub

Private Sub txtNombre_GotFocus()
    txtNombre.SelStart = 0
    txtNombre.SelLength = Len(txtNombre.Text)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And Trim(txtNombre.Text) <> "" Then chkAccesorios.SetFocus
End Sub

Private Sub vsGrid_DblClick()
    If MnuModificar.Enabled And vsGrid.Row > 0 Then
        loc_Edicion
    End If
End Sub


Private Sub vsGrid_RowColChange()
On Error Resume Next
    If Not MnuGrabar.Enabled Then
        If MnuNuevo.Enabled And vsGrid.Row > 0 Then
            Botones True, Val(vsGrid.Cell(flexcpData, vsGrid.Row, 0)) > 0, False
        Else
            Botones True, False, False
        End If
    End If
End Sub

