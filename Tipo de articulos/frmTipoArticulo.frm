VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTipoArticulo 
   BackColor       =   &H80000005&
   Caption         =   "Tipo de artículos"
   ClientHeight    =   5925
   ClientLeft      =   2910
   ClientTop       =   2250
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTipoArticulo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   7110
   Begin MSComctlLib.TreeView trTipos 
      CausesValidation=   0   'False
      Height          =   3135
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5530
      _Version        =   393217
      Indentation     =   718
      Style           =   7
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ilImage 
      Left            =   5400
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":0A3A
            Key             =   "tipo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":0D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":10DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":1430
            Key             =   "open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":17E9
            Key             =   "close"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7110
      _ExtentX        =   12541
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
            Object.Visible         =   0   'False
            Style           =   3
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
      TabIndex        =   1
      Top             =   5670
      Width           =   7110
      _ExtentX        =   12541
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
      Left            =   5760
      Top             =   1200
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
            Picture         =   "frmTipoArticulo.frx":1B3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":1C4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":1D5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":1E71
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":1F83
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":2095
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":23AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":24C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":27DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoArticulo.frx":2AF5
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
         Visible         =   0   'False
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Shortcut        =   ^G
         Visible         =   0   'False
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Shortcut        =   ^C
         Visible         =   0   'False
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVolver 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmTipoArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub loc_ShowPopUp()
    MnuSalir.Visible = False
    MnuVolver.Visible = False
    PopupMenu MnuOpciones
    MnuSalir.Visible = True
    MnuVolver.Visible = True
End Sub

Private Sub loc_FillTreeRecursivo(ByVal sHijosDe As String)
Dim iIdx As Integer
    Dim sQy As String, sPadre As String, sKey As String
    
    sQy = "Select TipCodigo, TipNombre, TipHijoDe From Tipo " & _
        "Where TipHijoDe " & IIf(sHijosDe = "", " IS Null", "IN (" & sHijosDe & ")") & _
        " Order by TipHijoDe, TipNombre"
    Set RsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    sHijosDe = ""
    Do While Not RsAux.EOF
        sPadre = ""
        If Not IsNull(RsAux("TipHijoDe")) Then If RsAux("TipHijoDe") > 0 Then sPadre = "T" & RsAux("TipHijoDe")
        With trTipos
            sKey = "T" & RsAux("TipCodigo")
            If sPadre <> "" Then
                iIdx = .Nodes(sPadre).Index
                .Nodes.Add iIdx, tvwChild, sKey, Trim(RsAux("TipNombre"))
                .Nodes(iIdx).Image = ilImage.ListImages("close").Index
            Else
                .Nodes.Add , , sKey, Trim(RsAux("TipNombre"))
            End If
            .Nodes(sKey).Image = ilImage.ListImages("tipo").Index
            .Nodes(sKey).ExpandedImage = ilImage.ListImages("open").Index
            If Not (.Nodes(sKey).Parent Is Nothing) Then .Nodes(sKey).Parent.Expanded = True
        End With
        sHijosDe = sHijosDe & IIf(sHijosDe <> "", ", ", "") & RsAux("TipCodigo")
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If sHijosDe <> "" Then loc_FillTreeRecursivo sHijosDe

End Sub

Private Sub loc_FillTree(Optional sKeySel As String = "")
On Error GoTo errFT
    Screen.MousePointer = 11
    
    trTipos.Visible = False
    Botones True, False, False, False, False, Toolbar1, Me
    trTipos.Nodes.Clear
    
    loc_FillTreeRecursivo ""
    
    Dim ndSelect As Node
    With trTipos
        If .Nodes.Count > 0 Then
            Botones True, True, True, False, False, Toolbar1, Me
            If sKeySel = "" Then
                .Nodes.Item(1).Selected = True
            Else
                For Each ndSelect In .Nodes
                    If ndSelect.Key = sKeySel Then
                        ndSelect.Selected = True
                        ndSelect.EnsureVisible
                        Exit For
                    End If
                Next
            End If
        End If
        .Visible = True
    End With
    
    Screen.MousePointer = 0
    Exit Sub
errFT:
    Screen.MousePointer = 11
    trTipos.Visible = True
    clsGeneral.OcurrioError "Error al cargar el árbol.", Err.Description, "Tipo de artículos."
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    ObtengoSeteoForm Me, 500, 500
    trTipos.Nodes.Clear
    trTipos.ImageList = ilImage
    loc_FillTree
    Screen.MousePointer = 0
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Error al ingresar al formulario."
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
On Error Resume Next
    trTipos.Move 120, Toolbar1.Height, Me.ScaleWidth - 120, Me.ScaleHeight - Toolbar1.Height - Status.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'Guardamos la posición del formulario.
    GuardoSeteoForm Me
'Cerramos la conexión.
    CierroConexion
'eliminamos la referencia de orcgsa.
    Set clsGeneral = Nothing
    End
    Exit Sub
End Sub

Private Sub MnuEliminar_Click()
    loc_EliminarTipo
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
    Dim frmT As New MaTipo
    With frmT
        .prmTipo = 0
        
        If Not (trTipos.SelectedItem Is Nothing) Then .prmHijoDe = Mid(trTipos.SelectedItem.Key, 2)
        
        .Show vbModal, Me
        If .prmTipo > 0 Then loc_FillTree "T" & .prmTipo
    End With
    Set frmT = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
ErrAN:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_Edicion()
    Me.Refresh
    Dim iTipo As Long
    Dim frmT As New MaTipo
    With frmT
        .prmTipo = Mid(trTipos.SelectedItem.Key, 2)
        .Show vbModal, Me
        If .prmTipo > 0 Then iTipo = .prmTipo
    End With
    Set frmT = Nothing
    DoEvents
    If iTipo > 0 Then loc_FillTree "T" & iTipo
    Me.Refresh
End Sub

Private Sub loc_EliminarTipo()
On Error GoTo errDel
    'Verificar si hay datos a validar.
    If MsgBox("Confrima eliminar el tipo seleccionado?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
        Screen.MousePointer = 11
        Dim oTipo As New clsTipo
        If oTipo.DeleteTipo(Mid(trTipos.SelectedItem.Key, 2)) Then
            loc_FillTree
        End If
        Set oTipo = Nothing
        Screen.MousePointer = 0
    End If
    Exit Sub
errDel:
    clsGeneral.OcurrioError "No se pudo eliminar el registro.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": loc_Nuevo
        Case "modificar": loc_Edicion
        Case "eliminar": loc_EliminarTipo
        Case "salir": Unload Me
    End Select

End Sub

Private Sub trTipos_DblClick()
    If Val(trTipos.SelectedItem.Children) = 0 Then loc_Edicion
End Sub

Private Sub trTipos_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 93
            loc_ShowPopUp
        Case vbKeyF2
        Case vbKeyDelete
            loc_EliminarTipo
    End Select
End Sub

Private Sub trTipos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = 0 And Button = vbRightButton Then
        trTipos.SelectedItem = trTipos.HitTest(x, y)
        loc_ShowPopUp
    End If
End Sub

Private Sub trTipos_NodeClick(ByVal Node As MSComCtlLib.Node)
    Botones True, True, True, False, False, Toolbar1, Me
End Sub
