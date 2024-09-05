VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form MaTipo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Productos"
   ClientHeight    =   5550
   ClientLeft      =   6015
   ClientTop       =   4095
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaTipo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7155
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
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
            Object.Width           =   4000
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox tArray 
      Height          =   585
      Left            =   1560
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Text            =   "frmMaTipo.frx":0442
      Top             =   2100
      Width           =   5475
   End
   Begin VB.TextBox tBusqWeb 
      Height          =   285
      Left            =   1560
      MaxLength       =   60
      TabIndex        =   10
      Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Top             =   1740
      Width           =   4515
   End
   Begin AACombo99.AACombo cLocReparacion 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
      _ExtentX        =   3836
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
   Begin AACombo99.AACombo cEspecie 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      BackColor       =   12648447
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
   Begin MSComctlLib.ListView lTipos 
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Local"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox tAbreviacion 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   5460
      MaxLength       =   12
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   19
      Top             =   5280
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   476
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
   Begin VB.TextBox tNombre 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaTipo.frx":04D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaTipo.frx":05EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaTipo.frx":06FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaTipo.frx":080F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaTipo.frx":0921
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaTipo.frx":0A33
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaTipo.frx":0D4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaTipo.frx":0E5F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cArancelMS 
      Height          =   315
      Left            =   1560
      TabIndex        =   14
      Top             =   2760
      Width           =   2715
      _ExtentX        =   4789
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
   Begin AACombo99.AACombo cArancelRM 
      Height          =   315
      Left            =   1560
      TabIndex        =   16
      Top             =   3120
      Width           =   2715
      _ExtentX        =   4789
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
   Begin VB.Label lArancelRM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   4260
      TabIndex        =   21
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Label lArancelMS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   4260
      TabIndex        =   20
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Recargo Resto:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Recargo &Mercosur:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Arra&y Característica:"
      Height          =   435
      Left            =   120
      TabIndex        =   11
      Top             =   2100
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Búsqueda Web:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1740
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Local Reparación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Abreviación:"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos &Ingresados"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   60
      TabIndex        =   17
      Top             =   3480
      Width           =   6675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Especie:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   615
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
         Caption         =   "&Del formulario   Alt+F4"
      End
   End
End
Attribute VB_Name = "MaTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arrArancel() As Currency

'Forms.-------------------------------------
Private frm1Campo As MaUnCampo

'String.---------------------------------------
Private strSeleccionado As String

'Boolean.---------------------------------------
Private sNuevoTipo As Boolean
Private sModificarTipo As Boolean

'RDO.---------------------------------------
Private RsTipo As rdoResultset

'Propiedades.---------------------------------------
Private bTipoLlamado As Byte
Private lSeleccionado As Long

Private Sub cArancelMS_Change()
    If cArancelMS.ListIndex = -1 Then
        lArancelMS.Caption = ""
    Else
        lArancelMS.Caption = Format(arrArancel(cArancelMS.ListIndex), "#,##0.000")
    End If
End Sub

Private Sub cArancelMS_Click()
    If cArancelMS.ListIndex = -1 Then
        lArancelMS.Caption = ""
    Else
        lArancelMS.Caption = Format(arrArancel(cArancelMS.ListIndex), "#,##0.000")
    End If
End Sub

Private Sub cArancelMS_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then cArancelRM.SetFocus
End Sub

Private Sub cArancelRM_Change()
    If cArancelRM.ListIndex = -1 Then
        lArancelRM.Caption = ""
    Else
        lArancelRM.Caption = Format(arrArancel(cArancelRM.ListIndex), "#,##0.000")
    End If
End Sub

Private Sub cArancelRM_Click()
    If cArancelRM.ListIndex = -1 Then
        lArancelRM.Caption = ""
    Else
        lArancelRM.Caption = Format(arrArancel(cArancelRM.ListIndex), "#,##0.000")
    End If
End Sub

Private Sub cArancelRM_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub cEspecie_Click()
    If Not sNuevoTipo And Not sModificarTipo Then
        If cEspecie.ListIndex = -1 Then
            Call Botones(False, False, False, False, False, Toolbar1, Me)
        Else
            Call Botones(True, False, False, False, False, Toolbar1, Me)
        End If
        lTipos.ListItems.Clear
    End If
End Sub

Private Sub cEspecie_Change()
    
    If Not sNuevoTipo And Not sModificarTipo Then
        If cEspecie.ListIndex = -1 Then
            Call Botones(False, False, False, False, False, Toolbar1, Me)
        Else
            Call Botones(True, False, False, False, False, Toolbar1, Me)
        End If
        lTipos.ListItems.Clear
    End If

End Sub

Private Sub cEspecie_GotFocus()

    Foco cEspecie
    Status.SimpleText = "Seleccione una especie."
    
End Sub

Private Sub cEspecie_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn And Not sModificarTipo And Not sNuevoTipo And cEspecie.ListIndex <> -1 Then
        CargoLista
        If lTipos.ListItems.Count > 0 Then
            lTipos.SetFocus
        End If
    ElseIf KeyCode = vbKeyReturn And cEspecie.ListIndex = -1 And cEspecie.Text <> "" Then
        If MsgBox("La especie seleccionada no existe." & Chr(13) & "Desea proceder al ingreso?", vbDefaultButton2 + vbYesNo + vbQuestion, "ATENCION") = vbYes Then
            
            Set frm1Campo = New MaUnCampo
            frm1Campo.pSeleccionado = 0
            frm1Campo.pCampoCodigo = "EspCodigo"
            frm1Campo.pCampoNombre = "EspNombre"
            frm1Campo.pTipoLlamado = TipoLlamado.IngresoNuevo
            frm1Campo.pTabla = "ESPECIE"
            frm1Campo.Caption = "Especies"
            frm1Campo.tDescripcion = cEspecie.Text
            frm1Campo.Show vbModal, Me
            DoEvents
            CargoEspecie
            If frm1Campo.pSeleccionado > 0 Then
                BuscoCodigoEnCombo cEspecie, frm1Campo.pSeleccionado
            End If
            Set frm1Campo = Nothing
        End If
    End If

End Sub

Private Sub cEspecie_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cEspecie.ListIndex > -1 And (sNuevoTipo Or sModificarTipo) Then
        Foco tNombre
    End If

End Sub
Private Sub cEspecie_LostFocus()
    cEspecie.SelLength = 0
    Status.SimpleText = ""
End Sub

Private Sub cLocReparacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tBusqWeb.SetFocus
End Sub

Private Sub Form_Activate()

    Screen.MousePointer = 0
    DoEvents
    cEspecie.SetFocus
    
End Sub

Private Sub Form_Load()
Dim sAux As String
    ObtengoSeteoForm Me, 500, 500
    DeshabilitoIngreso
    sNuevoTipo = False
    sModificarTipo = False
    strSeleccionado = vbNullString
    CargoLista
    CargoEspecie
    
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cLocReparacion
    SacoLargoCampo
    
    ReDim arrArancel(0)
    Cons = "Select * from Arancel"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        cArancelMS.AddItem Trim(RsAux!AraNombre)
        cArancelMS.ItemData(cArancelMS.NewIndex) = RsAux!AraCodigo
        cArancelRM.AddItem Trim(RsAux!AraNombre)
        cArancelRM.ItemData(cArancelRM.NewIndex) = RsAux!AraCodigo
        ReDim Preserve arrArancel(cArancelRM.NewIndex)
        If Not IsNull(RsAux!AraCoeficiente) Then
            arrArancel(cArancelRM.NewIndex) = RsAux!AraCoeficiente
        Else
            arrArancel(cArancelRM.NewIndex) = 0
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = ""

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If sNuevoTipo Or sModificarTipo Then
        If sNuevoTipo And Trim(tNombre.Text) = "" And Trim(tAbreviacion.Text) = "" Then Exit Sub
        
        If MsgBox("Antes de salir desea grabar los datos ingresados?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
            AccionGrabar
            If sNuevoTipo Or sModificarTipo Then Cancel = True
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    GuardoSeteoForm Me
    RsTipo.Close
    Forms(Forms.Count - 2).SetFocus
    
End Sub

Private Sub Label1_Click()

    If cEspecie.Enabled Then
        cEspecie.SetFocus
        cEspecie.SelStart = 0
        cEspecie.SelLength = Len(cEspecie.Text)
    End If

End Sub

Private Sub Label2_Click()

    If sNuevoTipo Or sModificarTipo Then
        tNombre.SetFocus
        tNombre.SelStart = 0
        tNombre.SelLength = Len(tNombre)
    End If

End Sub

Private Sub Label4_Click()

    If tAbreviacion.Enabled Then
        tAbreviacion.SetFocus
    End If
    
End Sub

Private Sub Label8_Click()
    cArancelMS.SetFocus
End Sub

Private Sub Label9_Click()
    cArancelRM.SetFocus
End Sub

Private Sub lTipos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Lista de tipos de productos ingresados para una especie."

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
    
    Screen.MousePointer = 11
    'Prendo Señal que es uno nuevo.
    sNuevoTipo = True
    
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    HabilitoIngreso
    tNombre.Text = ""
    tAbreviacion.Text = ""
    tNombre.SetFocus
    cEspecie.Enabled = False
    Screen.MousePointer = 0
    

End Sub

Sub AccionModificar()

    sModificarTipo = True
    
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    HabilitoIngreso
    
    On Error GoTo Error
    'Cargo el RS con el pais a modificar
    RsTipo.Close
    Cons = "Select * From Tipo Where TipCodigo = " & Right(lTipos.SelectedItem.Key, Len(lTipos.SelectedItem.Key) - 1)
    Set RsTipo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsTipo.EOF Then
        tNombre.Text = Trim(RsTipo!TipNombre)
        If Not IsNull(RsTipo!TipAbreviacion) Then tAbreviacion.Text = RsTipo!TipAbreviacion
        If Not IsNull(RsTipo!TipLocalRep) Then BuscoCodigoEnCombo cLocReparacion, RsTipo!TipLocalRep Else cLocReparacion.Text = ""
        If Not IsNull(RsTipo!TipBusqWeb) Then tBusqWeb.Text = Trim(RsTipo!TipBusqWeb)
        If Not IsNull(RsTipo!TipArrayCaract) Then tArray.Text = Trim(RsTipo!TipArrayCaract): tArray.Tag = Trim(tArray.Text)
        If Not IsNull(RsTipo!TipRecargoMS) Then BuscoCodigoEnCombo cArancelMS, RsTipo!TipRecargoMS
        If Not IsNull(RsTipo!TipRecargoRM) Then BuscoCodigoEnCombo cArancelRM, RsTipo!TipRecargoRM
    Else
        sModificarTipo = False
        MsgBox "El registro seleccionado ha sido eliminado", vbInformation, "ATENCIÓN"
        RsTipo.Close
        CargoLista
        DeshabilitoIngreso
    End If
    Exit Sub
    
Error:
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
    sModificarTipo = False
    CargoLista
    DeshabilitoIngreso

End Sub

Sub AccionGrabar()

    If Not ValidoCampos Then
        MsgBox "Los datos ingresados no son correctos o la ficha está incompleta.", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else
        If Not clsGeneral.TextoValido(tNombre) Then
            MsgBox "Se ha ingresado un caracter no válido, verifique.", vbExclamation, "ATENCION"
            Exit Sub
        End If
    End If
    
    If cEspecie.ListIndex = -1 Then
        MsgBox "Se debe seleccionar una especie.", vbExclamation, "ATENCION"
        Exit Sub
    End If
    
    
    If sNuevoTipo Then                  'Nuevo----------
        If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
            Screen.MousePointer = 11
            On Error GoTo errGrabar
            RsTipo.AddNew
            RsTipo!TipNombre = Trim(tNombre.Text)
            If Trim(tAbreviacion.Text) <> "" Then
                RsTipo!TipAbreviacion = Trim(tAbreviacion.Text)
            Else
                RsTipo!TipAbreviacion = Null
            End If
            If Trim(tBusqWeb.Text) <> "" Then
                RsTipo!TipBusqWeb = Trim(tBusqWeb.Text)
            Else
                RsTipo!TipBusqWeb = Null
            End If
            If Trim(tArray.Text) <> "" Then
                RsTipo!TipArrayCaract = Trim(tArray.Text)
            Else
                RsTipo!TipArrayCaract = Null
            End If
            RsTipo!TipEspecie = cEspecie.ItemData(cEspecie.ListIndex)
            If cLocReparacion.ListIndex > -1 Then RsTipo!TipLocalRep = cLocReparacion.ItemData(cLocReparacion.ListIndex)
            If cArancelMS.ListIndex > -1 Then RsTipo!TipRecargoMS = cArancelMS.ItemData(cArancelMS.ListIndex)
            If cArancelRM.ListIndex > -1 Then RsTipo!TipRecargoRM = cArancelRM.ItemData(cArancelRM.ListIndex)
            RsTipo.Update
            sNuevoTipo = False
            CargoLista
            
            If bTipoLlamado = TipoLlamado.IngresoNuevo Then
                Unload Me
                Exit Sub
            Else
                Screen.MousePointer = 0
                If MsgBox("Desea ingresar un nuevo tipo de artículo.", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
                    AccionNuevo
                    Exit Sub
                End If
            End If
            DeshabilitoIngreso
        End If
    
    Else                                    'Modificar----
    
        If tArray.Tag <> tArray.Text Then
            Cons = "Select * From ArticuloCaracteristica Where ACAArticulo IN (Select ArtID From Articulo Where ArtTipo =" & RsTipo!TipCodigo & ")"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                RsAux.Close
                MsgBox "Las características de comparación para los artículos del tipo ya están ingresadas, al modificar el array puede alterar la información existente", vbExclamation, "CUIDADO"
            Else
                RsAux.Close
            End If
        End If
        
        If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
            Screen.MousePointer = 11
            On Error GoTo errGrabar
            RsTipo.Edit
            RsTipo!TipNombre = Trim(tNombre.Text)
            If Trim(tAbreviacion.Text) <> "" Then
                RsTipo!TipAbreviacion = Trim(tAbreviacion.Text)
            Else
                RsTipo!TipAbreviacion = Null
            End If
            If Trim(tBusqWeb.Text) <> "" Then
                RsTipo!TipBusqWeb = Trim(tBusqWeb.Text)
            Else
                RsTipo!TipBusqWeb = Null
            End If
            If Trim(tArray.Text) <> "" Then
                RsTipo!TipArrayCaract = Trim(tArray.Text)
            Else
                RsTipo!TipArrayCaract = Null
            End If
            RsTipo!TipEspecie = cEspecie.ItemData(cEspecie.ListIndex)
            If cLocReparacion.ListIndex > -1 Then RsTipo!TipLocalRep = cLocReparacion.ItemData(cLocReparacion.ListIndex) Else RsTipo!TipLocalRep = Null
            If cArancelMS.ListIndex > -1 Then RsTipo!TipRecargoMS = cArancelMS.ItemData(cArancelMS.ListIndex) Else RsTipo!TipRecargoMS = Null
            If cArancelRM.ListIndex > -1 Then RsTipo!TipRecargoRM = cArancelRM.ItemData(cArancelRM.ListIndex) Else RsTipo!TipRecargoRM = Null
            RsTipo.Update
            sModificarTipo = False
            
            CargoLista
            DeshabilitoIngreso
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
    If Not RsTipo.EOF Then
        RsTipo.Requery
    End If
End Sub

Sub AccionEliminar()

    If MsgBox("¿Confirma eliminar el tipo: '" & lTipos.SelectedItem & "'?", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        
        On Error GoTo Error
        Screen.MousePointer = 11
        'Verifico que no existan artículos con ese tipo.
        Cons = "Select * from Articulo Where ArtTipo = " & Right(lTipos.SelectedItem.Key, Len(lTipos.SelectedItem.Key) - 1)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            Screen.MousePointer = 0
            MsgBox "Existen dependencias al tipo seleccionado, no podrá eliminarlo.", vbCritical, "ATENCION"
            RsAux.Close
            Exit Sub
        End If
        RsAux.Close
        
        'Cargo el RS con el tipo a eliminar
        RsTipo.Close
        Cons = "Select * From Tipo Where TipCodigo = " & Right(lTipos.SelectedItem.Key, Len(lTipos.SelectedItem.Key) - 1) _
            & " And TipNombre = '" & lTipos.SelectedItem & "'"
        Set RsTipo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Not RsTipo.EOF Then
            RsTipo.Delete
        Else
            Screen.MousePointer = 0
            MsgBox "El registro seleccionado ha sido eliminado o modificado por otra terminal.", vbInformation, "ATENCIÓN"
        End If
        
        RsTipo.Close
        CargoLista
        DeshabilitoIngreso
        Screen.MousePointer = 0
    End If
    Exit Sub
    
Error:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
    RsTipo.Requery
End Sub

Sub AccionCancelar()

    If sNuevoTipo Then
        If lTipos.ListItems.Count > 0 Then
            Call Botones(True, True, True, False, False, Toolbar1, Me)
        Else
            Call Botones(True, False, False, False, False, Toolbar1, Me)
        End If
        
    Else    'Cancelar modificacion
    
        If cEspecie.ListIndex = -1 Then
            lTipos.ListItems.Clear
            Call Botones(False, False, False, False, False, Toolbar1, Me)
        Else
            CargoLista
        End If
    End If
    
    DeshabilitoIngreso
    sNuevoTipo = False
    sModificarTipo = False
    
End Sub

Private Sub tAbreviacion_GotFocus()

    tAbreviacion.SelStart = 0
    tAbreviacion.SelLength = Len(tAbreviacion.Text)
    
End Sub

Private Sub tAbreviacion_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then cLocReparacion.SetFocus
    
End Sub

Private Sub tAbreviacion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Ingrese la abreviación para el tipo de producto."
    
End Sub


Private Sub tArray_GotFocus()
    
    With tArray
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub tArray_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cArancelMS.SetFocus
    End If
End Sub

Private Sub tBusqWeb_GotFocus()
    On Error Resume Next
    tBusqWeb.SelStart = 0
    tBusqWeb.SelLength = Len(tBusqWeb.Text)
End Sub

Private Sub tBusqWeb_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then tArray.SetFocus
End Sub

Private Sub tNombre_GotFocus()

    tNombre.SelStart = 0
    tNombre.SelLength = Len(tNombre)

End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cEspecie.ListIndex <> -1 And Trim(tNombre) <> "" Then
        tAbreviacion.SetFocus
    End If

End Sub

Private Sub tNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Ingrese un nombre para el tipo de producto."

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
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

'-----------------------------------------------------------------------------------------------
'   Carga la lista con los datos de la BD.
'-----------------------------------------------------------------------------------------------
Private Sub CargoLista()
On Error GoTo ErrCargoLista

    Screen.MousePointer = vbHourglass
    lTipos.ListItems.Clear
    
    If cEspecie.ListIndex = -1 Then
        Cons = "Select * From Tipo Where TipEspecie = 0"
    Else
        Cons = "Select * From Tipo Where TipEspecie = " & cEspecie.ItemData(cEspecie.ListIndex)
    End If
    strSeleccionado = tNombre.Text
    Set RsTipo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    I = 0
    Do While Not RsTipo.EOF
        Set itmx = lTipos.ListItems.Add(, "A" + Str(RsTipo!TipCodigo), Trim(RsTipo!TipNombre))
        If Trim(RsTipo!TipNombre) = Trim(strSeleccionado) Then I = lTipos.ListItems(lTipos.ListItems.Count).Index
        If Not IsNull(RsTipo!TipAbreviacion) Then itmx.SubItems(1) = Trim(RsTipo!TipAbreviacion)
        RsTipo.MoveNext
    Loop
    
    If lTipos.ListItems.Count > 0 Then
        Call Botones(True, True, True, False, False, Toolbar1, Me)
        If I <> 0 Then
            lTipos.SelectedItem = lTipos.ListItems(I)
            lSeleccionado = Mid(lTipos.SelectedItem.Key, 2, Len(lTipos.SelectedItem.Key))
        End If
    Else
        lSeleccionado = 0
        If cEspecie.ListIndex = -1 Then
            Call Botones(False, False, False, False, False, Toolbar1, Me)
        Else
            Call Botones(True, False, False, False, False, Toolbar1, Me)
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrCargoLista:
    clsGeneral.OcurrioError "Ocurrió un error al intentar cargar la lista.", Err.Description
End Sub

Private Sub DeshabilitoIngreso()

    tNombre.Text = ""
    tAbreviacion.Text = ""
    tNombre.Enabled = False:    tNombre.BackColor = Inactivo
    
    tAbreviacion.Enabled = False: tAbreviacion.BackColor = Inactivo
    tBusqWeb.Enabled = False: tBusqWeb.BackColor = Inactivo: tBusqWeb.Text = ""
    tArray.Enabled = False: tArray.BackColor = Inactivo: tArray.Text = ""
    
    cEspecie.Enabled = True: cEspecie.BackColor = Obligatorio
    cLocReparacion.Enabled = False: cLocReparacion.BackColor = Inactivo: cLocReparacion.Text = ""
    cArancelMS.Enabled = False: cArancelMS.BackColor = Inactivo: cArancelMS.Text = ""
    cArancelRM.Enabled = False: cArancelRM.BackColor = Inactivo: cArancelRM.Text = ""
    lTipos.Enabled = True
    
End Sub

Private Sub HabilitoIngreso()

    tNombre.Enabled = True: tNombre.BackColor = Obligatorio
    tAbreviacion.Enabled = True: tAbreviacion.BackColor = Blanco
    tBusqWeb.Enabled = True: tBusqWeb.BackColor = Blanco
    tArray.Enabled = True: tArray.BackColor = Blanco
    
    cLocReparacion.Enabled = True: cLocReparacion.BackColor = vbWhite
    cArancelMS.Enabled = True: cArancelMS.BackColor = Blanco
    cArancelRM.Enabled = True: cArancelRM.BackColor = Blanco
    lTipos.Enabled = False
    
End Sub

Private Function ValidoCampos()

    ValidoCampos = True
    
    If Trim(tNombre.Text) = "" Then ValidoCampos = False
    If cEspecie.ListIndex = -1 Then ValidoCampos = False

End Function

Sub CargoEspecie()

On Error GoTo ErrCarga

    Cons = "Select * From Especie Order by EspNombre"
    CargoCombo Cons, cEspecie, ""
    Exit Sub
    
ErrCarga:
    clsGeneral.OcurrioError "Ocurrió un error al cargar las especies.", Err.Description
    
End Sub


Public Property Get pSeleccionado() As Long

    pSeleccionado = lSeleccionado
    
End Property
Public Property Let pSeleccionado(Codigo As Long)

    lSeleccionado = Codigo

End Property

Public Property Get pTipoLlamado() As Byte

    pTipoLlamado = bTipoLlamado

End Property

Public Property Let pTipoLlamado(Codigo As Byte)

    bTipoLlamado = Codigo

End Property
Private Sub SacoLargoCampo()
On Error GoTo errSLC
    Cons = "Select length as Largo from sysColumns Where name = 'TipArrayCaract'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        tArray.MaxLength = RsAux!Largo
    Else
        tArray.MaxLength = 150
    End If
    RsAux.Close
    Exit Sub
errSLC:
    clsGeneral.OcurrioError "Ocurrió el siguiente error al cargar el largo del campo.", Err.Description
End Sub

