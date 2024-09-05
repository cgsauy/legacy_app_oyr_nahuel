VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form EntMercALocal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrega de Mercadería a Local"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EntMercALocal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
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
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   3
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
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   3000
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.OptionButton opAltaBaja 
      Caption         =   "&Baja"
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   19
      Top             =   1680
      Width           =   1095
   End
   Begin VB.OptionButton opAltaBaja 
      Caption         =   "Al&ta"
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   1680
      Width           =   1095
   End
   Begin ComctlLib.ListView lvArticulo 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cant."
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Artículo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Estado "
         Object.Width           =   1481
      EndProperty
   End
   Begin VB.TextBox tUsuario 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   15
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox tComentario 
      Height          =   285
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3840
      Width           =   3975
   End
   Begin VB.ComboBox cEstadoViejo 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3120
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox tCantidad 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox cArticulo 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1080
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   960
      Width           =   3615
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   4575
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin VB.TextBox tFecha 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1080
      MaxLength       =   12
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cLocal 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " A&rtículos ingresados"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentario:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Estado:"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cantidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Local:"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercALocal.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercALocal.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercALocal.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercALocal.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercALocal.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercALocal.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercALocal.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercALocal.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercALocal.frx":10E2
            Key             =   "Alta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EntMercALocal.frx":13FC
            Key             =   "Baja"
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
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Shortcut        =   ^E
         Visible         =   0   'False
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
         Caption         =   "&Volver al Formulario Anterior"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "EntMercALocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modulo de Entrega de Mercadería a local.
'Común a RepCG.
Option Explicit
Dim Msg As String

Private Sub cArticulo_GotFocus()
    cArticulo.SelStart = 0
    cArticulo.SelLength = Len(cArticulo.Text)
    Status.SimpleText = " Ingrese el código o nombre del artículo."
End Sub

Private Sub cArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Trim(cArticulo.Text) <> vbNullString Then
        If Not IsNumeric(cArticulo.Text) Then
            BuscoArtXNombre
        Else
            BuscoArticuloxCodigo cArticulo
        End If
        If cArticulo.ListIndex > -1 Then tCantidad.SetFocus
    ElseIf KeyAscii = vbKeyReturn And lvArticulo.ListItems.Count > 0 Then
        lvArticulo.SetFocus
    End If

End Sub
Private Sub cArticulo_LostFocus()
    Status.SimpleText = vbNullString
End Sub
Private Sub cEstadoViejo_Change()
    Selecciono cEstadoViejo, cEstadoViejo.Text, gTecla
End Sub
Private Sub cEstadoViejo_GotFocus()
    cEstadoViejo.SelStart = 0
    cEstadoViejo.SelLength = Len(cEstadoViejo.Text)
    Status.SimpleText = " Seleccione el estado original del artículo ."
End Sub
Private Sub cEstadoViejo_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cEstadoViejo.ListIndex
End Sub
Private Sub cEstadoViejo_KeyPress(KeyAscii As Integer)
    cEstadoViejo.ListIndex = gIndice
    If KeyAscii = vbKeyReturn And cEstadoViejo.ListIndex > -1 Then
        opAltaBaja(0).SetFocus
    End If
End Sub

Private Sub cEstadoViejo_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cEstadoViejo
End Sub

Private Sub cEstadoViejo_LostFocus()
    gIndice = -1
    cEstadoViejo.SelLength = 0
    Status.SimpleText = vbNullString
End Sub

Private Sub cLocal_Click()
    lvArticulo.ListItems.Clear
    LimpioDatosIngreso
End Sub

Private Sub cLocal_Change()

    Selecciono cLocal, cLocal.Text, gTecla

End Sub

Private Sub cLocal_GotFocus()
    cLocal.SelStart = 0
    cLocal.SelLength = Len(cLocal.Text)
    Status.SimpleText = " Seleccione un Local."
End Sub
Private Sub cLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cLocal.ListIndex
End Sub
Private Sub cLocal_KeyPress(KeyAscii As Integer)
    cLocal.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then
        If cLocal.ListIndex > -1 Then
            cArticulo.SetFocus
        Else
            MsgBox "El ingreso del local es obligatorio.", vbExclamation, "ATENCIÓN"
            cLocal.SetFocus
        End If
    End If
End Sub
Private Sub cLocal_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cLocal
End Sub
Private Sub cLocal_LostFocus()
    gIndice = -1
    cLocal.SelLength = 0
    Status.SimpleText = vbNullString
End Sub

Private Sub Form_Activate()
    DoEvents
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    gIndice = -1
    CargoLocales
    CargoEstado
    DeshabilitoIngreso
    SetearLView lvValores.Grilla Or lvValores.FullRow, lvArticulo
    Exit Sub
ErrLoad:
    clsError.MuestroError Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Forms(Forms.Count - 2).SetFocus
End Sub

Private Sub Label1_Click()
    Foco cLocal
End Sub
Private Sub Label10_Click()
    If lvArticulo.ListItems.Count > 0 And lvArticulo.Enabled Then
        lvArticulo.ListItems(1).Selected = True
        lvArticulo.SetFocus
    End If
End Sub
Private Sub Label2_Click()
    Foco tFecha
End Sub
Private Sub Label3_Click()
    Foco cArticulo
End Sub
Private Sub Label4_Click()
    Foco tCantidad
End Sub
Private Sub Label5_Click()
    Foco cEstadoViejo
End Sub
Private Sub Label7_Click()
    Foco tComentario
End Sub
Private Sub Label8_Click()
    Foco tUsuario
End Sub


Private Sub lvArticulo_GotFocus()
    Status.SimpleText = " Lista de artículos. [Supr] - elimina"
End Sub

Private Sub lvArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            tComentario.SetFocus
        Case vbKeyDelete
            If lvArticulo.ListItems.Count > 0 Then lvArticulo.ListItems.Remove lvArticulo.SelectedItem.Index
    End Select
End Sub

Private Sub lvArticulo_LostFocus()
    Status.SimpleText = vbNullString
End Sub
Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub
Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub
Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub
Private Sub MnuVolver_Click()
    Unload Me
End Sub
Private Sub AccionNuevo()
      
    'Habilito y Desabilito Botones
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso
    tFecha.Text = Format(Date, "d-Mmm-yyyy")
    
End Sub
Private Sub AccionGrabar()

    If ValidoDatos Then
        If MsgBox("¿Confirma almacenar los datos ingresados?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then GraboDatos
    End If
    
End Sub

Private Sub AccionCancelar()
    Screen.MousePointer = vbHourglass
    DeshabilitoIngreso
    Botones True, False, False, False, False, Toolbar1, Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub opAltaBaja_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then InsertoRenglon
End Sub

Private Sub tCantidad_GotFocus()
    tCantidad.SelStart = 0
    tCantidad.SelLength = Len(tCantidad.Text)
    Status.SimpleText = " Ingrese la cantidad de artículos."
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tCantidad.Text) <> vbNullString Then
        If IsNumeric(tCantidad.Text) Then
            If cEstadoViejo.ListIndex = -1 Then BuscoCodigoEnCombo cEstadoViejo, CInt(paEstadoArticuloEntrega)
            cEstadoViejo.SetFocus
        Else
            MsgBox " El formato ingresado no es numérico.", vbExclamation, "ATENCIÓN"
            tCantidad.SetFocus
        End If
    End If
End Sub

Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0
    tComentario.SelLength = Len(tComentario.Text)
    Status.SimpleText = " Ingrese un comentario."
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tUsuario.SetFocus
End Sub
Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0
    tFecha.SelLength = Len(tFecha.Text)
    Status.SimpleText = " Ingrese una fecha."
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tFecha.Text) Then
            cLocal.SetFocus
        Else
            MsgBox " La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            tFecha.SetFocus
        End If
    End If
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "d-Mmm-yyyy")
    Status.SimpleText = vbNullString
End Sub

Private Sub tFecha_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.SimpleText = " Ingrese una fecha."
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        
        Case "nuevo"
            AccionNuevo
        
        Case "grabar"
            AccionGrabar
        
        Case "cancelar"
            AccionCancelar
        
        Case "salir"
            Unload Me
            
    End Select

End Sub
Private Sub DeshabilitoIngreso()
    tFecha.Text = vbNullString
    tFecha.Enabled = False
    tFecha.BackColor = Inactivo
    cLocal.ListIndex = -1
    cLocal.Enabled = False
    cLocal.BackColor = Inactivo
    cArticulo.Enabled = False
    cArticulo.BackColor = Inactivo
    cArticulo.Clear
    tCantidad.Text = vbNullString
    tCantidad.Enabled = False
    tCantidad.BackColor = Inactivo
    cEstadoViejo.Enabled = False
    cEstadoViejo.BackColor = Inactivo
    cEstadoViejo.ListIndex = -1
    tComentario.BackColor = Inactivo
    tComentario.Enabled = False
    tComentario.Text = vbNullString
    lvArticulo.Enabled = False
    lvArticulo.ListItems.Clear
    tUsuario.Enabled = False
    tUsuario.BackColor = Inactivo
    tUsuario.Text = vbNullString
    tUsuario.Tag = vbNullString
    opAltaBaja(0).Enabled = False
    opAltaBaja(0).Value = False
    opAltaBaja(1).Enabled = False
    opAltaBaja(1).Value = False
End Sub
Private Sub HabilitoIngreso()
    tFecha.Enabled = True
    tFecha.BackColor = Obligatorio
    cLocal.Enabled = True
    cLocal.BackColor = Obligatorio
    cArticulo.Enabled = True
    cArticulo.BackColor = Obligatorio
    cArticulo.Clear
    tCantidad.Text = vbNullString
    tCantidad.Enabled = True
    tCantidad.BackColor = Obligatorio
    cEstadoViejo.Enabled = True
    cEstadoViejo.BackColor = Obligatorio
    tComentario.BackColor = Blanco
    tComentario.Enabled = True
    tComentario.Text = vbNullString
    lvArticulo.Enabled = True
    lvArticulo.ListItems.Clear
    tUsuario.Enabled = True
    tUsuario.BackColor = Obligatorio
    tUsuario.Text = vbNullString
    tUsuario.Tag = vbNullString
    opAltaBaja(0).Enabled = True
    opAltaBaja(0).Value = True
    opAltaBaja(1).Enabled = True
End Sub
Private Sub CargoLocales()
On Error GoTo ErrCL

    cLocal.Clear
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cLocal, ""
    Exit Sub
ErrCL:
    clsError.MuestroError "Ocurrio un error al cargar los locales."
    Screen.MousePointer = vbDefault
End Sub
Private Sub CargoEstado()
On Error GoTo ErrCE
    
    Cons = "Select EsMCodigo, EsMAbreviacion From EstadoMercaderia " _
        & " Order by EsMAbreviacion"
    CargoCombo Cons, cEstadoViejo, ""
    Exit Sub
ErrCE:
    Screen.MousePointer = vbDefault
    clsError.MuestroError "Ocurrio un error al cargar los Estados."
End Sub
Private Sub BuscoArtXNombre()
On Error GoTo ErrBAXN
    BuscoArticuloXNombre cArticulo.Text
    cArticulo.Clear
    If LiAyuda.pSeleccionado > 0 Then
        BuscoArticuloID LiAyuda.pSeleccionado, cArticulo
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBAXN:
    clsError.MuestroError "Ocurrio un error inesperado al buscar los artículos."
End Sub
Private Function BuscoStockLocal() As String
On Error GoTo ErrBSL

    Screen.MousePointer = vbHourglass
    Cons = "Select StLCantidad from StockLocal" _
        & " Where StLLocal = " & cLocal.ItemData(cLocal.ListIndex) _
        & " And StLArticulo = " & cArticulo.ItemData(0) _
        & " And StLEstado = " & cEstadoViejo.ItemData(cEstadoViejo.ListIndex) _
        & " And StLTipoLocal = " & TipoLocal.Deposito
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAux.EOF Then
        BuscoStockLocal = "El local no posee el artículo con ese estado."
    Else
        If RsAux!StLCantidad < CInt(tCantidad.Text) Then
            BuscoStockLocal = "El local no posee tantos artículos con ese estado."
        Else
            BuscoStockLocal = vbNullString
        End If
    End If
    RsAux.Close
    
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrBSL:
    clsError.MuestroError "Ocurrio un error al buscar el stock del local."
End Function

Private Sub InsertoRenglon()
Dim Msg As String
On Error GoTo ErrControl

    Screen.MousePointer = vbHourglass
    If cArticulo.ListIndex = -1 Then
        MsgBox "No hay seleccionado un artículo.", vbExclamation, "ATENCIÓN"
        cArticulo.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If Not IsNumeric(tCantidad.Text) Then
        MsgBox "La cantidad ingresada no es correcta.", vbInformation, "ATENCIÓN"
        tCantidad.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If CInt(tCantidad.Text) < 1 Then
        MsgBox "La cantidad ingresada no es correcta.", vbInformation, "ATENCIÓN"
        tCantidad.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If cEstadoViejo.ListIndex = -1 Then
        MsgBox "No hay seleccionado un estado.", vbInformation, "ATENCIÓN"
        cEstadoViejo.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If opAltaBaja(1).Value Then
        Msg = BuscoStockLocal
        If Msg <> vbNullString Then
            MsgBox Msg, vbInformation, "ATENCIÓN"
            Exit Sub
        End If
    End If
    
    'Si llegue aca es porque puedo insertar.
    On Error GoTo ErrInserto
    Msg = " Ya se ingreso ese artículo con las mismas condiciones."
    'Clave Articulo + estado
    Set itmX = lvArticulo.ListItems.Add(, cArticulo.ItemData(0) & "V" & cEstadoViejo.ItemData(cEstadoViejo.ListIndex))
    itmX.Text = tCantidad.Text
    itmX.SubItems(1) = Trim(cArticulo.Text)
    itmX.SubItems(2) = Trim(cEstadoViejo.Text)
    If opAltaBaja(0).Value Then
        itmX.Tag = "Alta"
        itmX.SmallIcon = "Alta"
    Else
        itmX.Tag = "Baja"
        itmX.SmallIcon = "Baja"
    End If
    LimpioDatosIngreso
    cArticulo.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrControl:
    Screen.MousePointer = vbDefault
    clsError.MuestroError "Ocurrio un error al controlar los datos."
    Exit Sub
    
ErrInserto:
    Screen.MousePointer = vbDefault
    If Msg = vbNullString Then
        clsError.MuestroError Err.Description
    Else
        clsError.MuestroError Msg
    End If
End Sub
Private Sub LimpioDatosIngreso()
    cArticulo.Clear
    tCantidad.Text = vbNullString
    cEstadoViejo.ListIndex = -1
   
End Sub
Private Sub tUsuario_GotFocus()
    tUsuario.SelStart = 0
    tUsuario.SelLength = Len(tUsuario.Text)
    Status.SimpleText = " Ingrese su código de Usuario."
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    tUsuario.Tag = vbNullString
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = BuscoUsuario(CInt(tUsuario.Text))
            If CInt(tUsuario.Tag) > 0 Then
                AccionGrabar
            Else
                tUsuario.Tag = vbNullString
            End If
        Else
            MsgBox "El formato del código de usuario no es numérico.", vbExclamation, "ATENCIÓN"
            tUsuario.SetFocus
        End If
    End If
End Sub

Private Function ValidoDatos() As Boolean

    ValidoDatos = True
    
    If lvArticulo.ListItems.Count = 0 Then
        MsgBox "No hay artículos ingresados.", vbExclamation, "ATENCIÓN"
        ValidoDatos = False
        Exit Function
    End If
    
    If tUsuario.Tag = vbNullString Then
        MsgBox "Debe ingresar su código de usuario.", vbExclamation, "ATENCIÓN"
        tUsuario.SetFocus
        ValidoDatos = False
        Exit Function
    End If
    
    If Not TextoValido(tComentario.Text) Then
        MsgBox "Se ingreso un carácter no válido en el comentario.", vbExclamation, "ATENCIÓN"
        tComentario.SetFocus
        ValidoDatos = False
        Exit Function
    End If
    
End Function
Private Sub GraboDatos()
Dim lnCodigoControl As Long

    On Error GoTo ErrGD
    FechaDelServidor
    
    cBase.BeginTrans
    On Error GoTo ErrResumo
    
    Cons = "Select MAX(CMeCodigo) From ControlMercaderia"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If IsNull(RsAux(0)) Then
        lnCodigoControl = 0
    Else
        lnCodigoControl = RsAux(0)
    End If
    RsAux.Close
    
    Cons = "INSERT INTO ControlMercaderia (CMeTipoLocal, CMeLocal, CMeFecha, CMeTipo, CMeComentario, CMeUsuario)" _
        & " Values (" & TipoLocal.Deposito
    
    Cons = Cons & ", " & cLocal.ItemData(cLocal.ListIndex) _
        & ", '" & Format(Now, FormatoFH) & "', " & TipoControlMercaderia.EntregaMercaderia
        
    If Trim(tComentario.Text) = vbNullString Then
        Cons = Cons & ", Null, " & tUsuario.Tag & ")"
    Else
        Cons = Cons & ", '" & tComentario.Text & "', " & tUsuario.Tag & ")"
    End If
    cBase.Execute (Cons)
    
    If cBase.RowsAffected = 0 Then
        Msg = "No se pudo insertar en la tabla ControlMercaderia."
        RsAux.Edit
    End If
    
    Cons = "Select MAX(CMeCodigo) From ControlMercaderia" _
        & " Where CMeCodigo > " & lnCodigoControl _
        & " And CMeTipo = " & TipoControlMercaderia.EntregaMercaderia
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If IsNull(RsAux(0)) Then
        lnCodigoControl = 0
    Else
        lnCodigoControl = RsAux(0)
    End If
    RsAux.Close
    
    For Each itmX In lvArticulo.ListItems
            InsertoRenglonLocal (lnCodigoControl)
    Next
    cBase.CommitTrans
    AccionCancelar
    Exit Sub

ErrGD:
    Screen.MousePointer = vbDefault
    clsError.MuestroError "Ocurrio un error al iniciar la transacción."
    Exit Sub
    
ErrResumo:
    Resume Relajo
    
Relajo:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    If Msg = vbNullString Then
        clsError.MuestroError Err.Description
    Else
        clsError.MuestroError Msg
    End If
    Exit Sub
    
End Sub
Private Sub InsertoRenglonLocal(lnCodControl As Long)

    Cons = "Select * From StockLocal " _
        & " Where StLTipoLocal = " & TipoLocal.Deposito _
        & " And StlLocal = " & cLocal.ItemData(cLocal.ListIndex) _
        & " And StLArticulo = " & Mid(itmX.Key, 1, InStr(itmX.Key, "V") - 1) _
        & " And StLEstado = " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, Len(itmX.Key))
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If itmX.Tag = "Baja" Then
        
        If RsAux.EOF Then
            Msg = "El local no posee más artículos " & Trim(itmX.SubItems(1))
            RsAux.Close
            RsAux.Edit
        Else
            If RsAux!StLCantidad < CInt(itmX.Text) Then
                Msg = "El local no posee tantos artículos " & Trim(itmX.SubItems(1))
                RsAux.Close
                RsAux.Edit
            ElseIf RsAux!StLCantidad = CInt(itmX.Text) Then
                RsAux.Delete
                RsAux.Close
            Else
                RsAux.Edit
                RsAux!StLCantidad = RsAux!StLCantidad - CInt(itmX.Text)
                RsAux.Update
                RsAux.Close
            End If
        End If
        Cons = " INSERT INTO RenglonControlMercaderia Values(" & lnCodControl _
            & ", " & Mid(itmX.Key, 1, InStr(itmX.Key, "V") - 1) _
            & ", " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, Len(itmX.Key)) _
            & ", Null, " & CInt(itmX.Text) * -1 & ")"
            
        cBase.Execute (Cons)
        
        MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), Mid(itmX.Key, 1, InStr(itmX.Key, "V") - 1), CInt(itmX.Text), Mid(itmX.Key, InStr(itmX.Key, "V") + 1, Len(itmX.Key)), -1
    Else
         If RsAux.EOF Then
             RsAux.AddNew
             RsAux!StLTipoLocal = TipoLocal.Deposito
             RsAux!StLLocal = cLocal.ItemData(cLocal.ListIndex)
             RsAux!StLArticulo = Mid(itmX.Key, 1, InStr(itmX.Key, "V") - 1)
             RsAux!StLEstado = Mid(itmX.Key, InStr(itmX.Key, "V") + 1, Len(itmX.Key))
             RsAux!StLCantidad = CInt(itmX.Text)
             RsAux.Update
             RsAux.Close
         Else
             RsAux.Edit
             RsAux!StLCantidad = RsAux!StLCantidad + CInt(itmX.Text)
             RsAux.Update
             RsAux.Close
         End If
        Cons = " INSERT INTO RenglonControlMercaderia Values(" & lnCodControl _
            & ", " & Mid(itmX.Key, 1, InStr(itmX.Key, "V") - 1) _
            & ", " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, Len(itmX.Key)) _
            & ", Null, " & CInt(itmX.Text) * 1 & ")"
            
        cBase.Execute (Cons)

         MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), Mid(itmX.Key, 1, InStr(itmX.Key, "V") - 1), CInt(itmX.Text), Mid(itmX.Key, InStr(itmX.Key, "V") + 1, Len(itmX.Key)), 1
    End If
    
    'Veo si afecta el stock físico.
    Cons = "Select EsMBajaStockTotal From EstadoMercaderia" _
        & " Where EsMCodigo = " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, Len(itmX.Key))
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAux.EOF Then
        Msg = "No se encontró los datos del estado, reintente."
        RsAux.Close
        RsAux.Edit
    Else
        'Si afecta no me caliento no tengo que tocar el total.
'        If RsAux!EsMBajaStockTotal = 0 Then
           'Modifco el stock Total
           RsAux.Close
           Cons = "Select * From StockTotal " _
                & " Where StTArticulo = " & Mid(itmX.Key, 1, InStr(itmX.Key, "V") - 1) _
                & " And StTTipoEstado = " & TipoEstadoMercaderia.Fisico _
                & " And StTEstado = " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, Len(itmX.Key))
            
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If RsAux.EOF Then
                If itmX.Tag = "Baja" Then
                    Msg = "El stock total no coincide con el del local, verifique."
                    RsAux.Close
                    RsAux.Edit
                Else
                    RsAux.AddNew
                    RsAux!StTArticulo = Mid(itmX.Key, 1, InStr(itmX.Key, "V") - 1)
                    RsAux!StTTipoEstado = TipoEstadoMercaderia.Fisico
                    RsAux!StTEstado = Mid(itmX.Key, InStr(itmX.Key, "V") + 1, Len(itmX.Key))
                    RsAux!StTCantidad = CInt(itmX.Text)
                    RsAux.Update
                End If
            Else
                If itmX.Tag = "Baja" Then
                    If RsAux!StTCantidad >= CInt(itmX.Text) Then
                        RsAux.Edit
                        RsAux!StTCantidad = RsAux!StTCantidad - CInt(itmX.Text)
                        RsAux.Update
                    Else
                        Msg = "El stock total no coincide con el del local, verifique."
                        RsAux.Close
                        RsAux.Edit
                    End If
                Else
                    RsAux.Edit
                    RsAux!StTCantidad = RsAux!StTCantidad + CInt(itmX.Text)
                    RsAux.Update
                End If
            End If
            RsAux.Close
'        Else
'            RsAux.Close
'        End If
    End If
    
End Sub
