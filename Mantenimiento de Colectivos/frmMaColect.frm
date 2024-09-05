VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMaColect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colectivos"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaColect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   1200
      MaxLength       =   180
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   4920
      Width           =   5775
   End
   Begin VB.CommandButton bEmail2 
      Caption         =   "Emai&l ..."
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton bEmail1 
      Caption         =   "Ema&il ..."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox tClave 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   17
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CheckBox chCerrado 
      Caption         =   "C&errado"
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton bDireccion 
      Caption         =   "..."
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox tIglesia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox tCivil 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox tNombreColectivo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   60
      TabIndex        =   9
      Top             =   3120
      Width           =   5415
   End
   Begin VB.TextBox tCodigo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   840
      MaxLength       =   6
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin MSMask.MaskEdBox tCiCliente1 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#.###.###-#"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tCICliente2 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#.###.###-#"
      PromptChar      =   "_"
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   5940
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
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
            AutoSize        =   1
            Object.Width           =   7329
            Key             =   "Msg"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentario:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lEmailCli2 
      BackColor       =   &H00800000&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   33
      Top             =   2760
      Width           =   6015
   End
   Begin VB.Label lEmailCli1 
      BackColor       =   &H00800000&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   32
      Top             =   1680
      Width           =   6015
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Contra&seña:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio:"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label labDirCliente1 
      BackColor       =   &H00800000&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   28
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Label labDirCliente2 
      BackColor       =   &H00800000&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   27
      Top             =   2400
      Width           =   6015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label labUsuario 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "AO"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5760
      TabIndex        =   25
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label labDireccion 
      BackColor       =   &H00800000&
      Caption         =   "Ayacucho 1030"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2040
      TabIndex        =   24
      Top             =   3840
      Width           =   4935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Di&rección de Envío:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fech&a C. Religiosa:"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha Civil:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Título del Colectivo:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label labCliente2 
      BackColor       =   &H00800000&
      Caption         =   "Alberta Justiniana"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Label labCliente1 
      BackColor       =   &H00800000&
      Caption         =   "Alberto Mapache"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "C.&I.:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaColect.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaColect.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaColect.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaColect.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaColect.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaColect.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaColect.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaColect.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaColect.frx":10E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&C.I.:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Có&digo:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
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
      Begin VB.Menu MnuLineaOp 
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
      Begin VB.Menu MnuSalirDelForm 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMaColect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Long.---------------------------------------------
Private aDireccion As Long
Private m_IdCol As Long

'BOOLEAN.--------------------------------------
Private sNuevo As Boolean
Private sModificar As Boolean

Public Property Let ColectivoID(ByVal lCod As Long)
    m_IdCol = lCod
End Property


Private Sub bDireccion_Click()
Dim aCopia As Long
    On Error GoTo ErrDirecccion
    Screen.MousePointer = 11
    aCopia = aDireccion
    Dim miDir As New clsDireccion
    If sNuevo Then
        miDir.ActivoFormularioDireccion cBase, aDireccion, 0
    Else
        miDir.ActivoFormularioDireccion cBase, aDireccion, 0, "Colectivo", "ColDirEnvio", "ColCodigo", tCodigo.Text
    End If
    aDireccion = miDir.CodigoDeDireccion
    Set miDir = Nothing
    If aCopia = 0 And aDireccion <> 0 And Not sNuevo Then
        'Updateo la tabla colectivo y le cargo la direccion.
        Cons = "Update Colectivo Set ColDirEnvio = " & aDireccion & " Where ColCodigo = " & tCodigo.Text
        cBase.Execute (Cons)
    End If
    If aDireccion > 0 Then
        labDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, aDireccion, False, True, False, True)
    Else
        labDireccion.Caption = ""
    End If
    If chCerrado.Enabled Then Foco tClave
    Screen.MousePointer = 0
    Exit Sub
ErrDirecccion:
    clsGeneral.OcurrioError "Ocurrió un error al cargar la dirección.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bEmail1_Click()
    EjecutarApp App.Path & "\EMails.exe", Val(tCiCliente1.Tag), True
    lEmailCli1.Caption = CargoDireccionEMail(Val(tCiCliente1.Tag))
End Sub

Private Sub bEmail2_Click()
    EjecutarApp App.Path & "\EMails.exe", Val(tCICliente2.Tag), True
    lEmailCli2.Caption = CargoDireccionEMail(Val(tCICliente2.Tag))
End Sub

Private Sub chCerrado_GotFocus()
    Ayuda "Indique si el colectivo esta cerrado."
End Sub
Private Sub chCerrado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tComentario.SetFocus
End Sub
Private Sub chCerrado_LostFocus()
    Ayuda ""
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Initialize()
    sNuevo = False: sModificar = False
    aDireccion = 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    OcultoCampos
    LimpioCampos
    tCiCliente1.BackColor = Obligatorio
    If m_IdCol > 0 Then
        BuscoColectivo m_IdCol
    End If
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ayuda ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If sNuevo And aDireccion > 0 Then
        'Borro los datos de la direccion copia
        Cons = "Select * from Direccion Where DirCodigo = " & aDireccion
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Delete
        RsAux.Close
    End If
    CierroConexion
    Set clsGeneral = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco tCiCliente1
End Sub

Private Sub Label11_Click()
    Foco tClave
End Sub

Private Sub Label2_Click()
    Foco tCICliente2
End Sub
Private Sub Label3_Click()
    Foco tNombreColectivo
End Sub
Private Sub Label4_Click()
    Foco tCivil
End Sub
Private Sub Label5_Click()
    Foco tIglesia
End Sub
Private Sub Label6_Click()
    Foco tCodigo
End Sub
Private Sub Label7_Click()
    Foco bDireccion
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

Private Sub MnuSalirDelForm_Click()
    Unload Me
End Sub

'-----------------------------------------------------------
'Ingresa un nuevo empleado.
'-----------------------------------------------------------
Sub AccionNuevo()
On Error GoTo ErrAN
    Botones False, False, False, True, True, Toolbar1, Me
    tCodigo.Text = ""
    HabilitoCampos
    LimpioCampos
    Foco tCiCliente1
    sNuevo = True
    Exit Sub
ErrAN:
On Error Resume Next
    Screen.MousePointer = 11
    clsGeneral.OcurrioError "Ocurrio un error al dar Acción Nuevo.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
'-----------------------------------------------------------
'Presiono boton Cancelar.
'-----------------------------------------------------------
Sub AccionCancelar()
On Error GoTo ErrAC
    
    aDireccion = 0
    OcultoCampos
    'Apago las señales
    If sNuevo Then
        sNuevo = False
        LimpioCampos
        Botones True, False, False, False, False, Toolbar1, Me
    Else
        Botones True, True, True, False, False, Toolbar1, Me
        sModificar = False
        BuscoColectivo tCodigo.Text
    End If
    Foco tCodigo
    Exit Sub
ErrAC:
    clsGeneral.OcurrioError "Ocurrio un error al dar Acción Cancelar.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

'-----------------------------------------------------------
'Presiono boton modificar.
'-----------------------------------------------------------
Sub AccionModificar()
On Error GoTo ErrAM
    Screen.MousePointer = 11
    Cons = "SELECT * FROM Colectivo Where ColCodigo = " & tCodigo.Text
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "Otra terminal pudó eliminar el Colectivo seleccionado.", vbInformation, "ATENCIÓN"
        LimpioCampos
        Botones True, False, False, False, False, Toolbar1, Me
        Screen.MousePointer = 0
        Exit Sub
    End If
    If RsAux!ColFModificacion = CDate(tCodigo.Tag) Then
        'Señal de modificación.
        sModificar = True
        Botones False, False, False, True, True, Toolbar1, Me
        'Muestro campos
        HabilitoCampos
        RsAux.Close
    Else
        RsAux.Close
        MsgBox "Otra terminal modificó la ficha, verifique los cambios realizados.", vbInformation, "ATENCIÓN"
        BuscoColectivo tCodigo.Text
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrAM:
    clsGeneral.OcurrioError "Ocurrio un error al dar Acción Modificar.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

'------------------------------------------------------------
'Limpio los objetos de pantalla.
'------------------------------------------------------------
Private Sub LimpioCampos()
On Error Resume Next
    tCiCliente1.Text = "": labCliente1.Caption = ""
    tCICliente2.Text = "": labCliente2.Caption = vbNullString
    tNombreColectivo.Text = ""
    tCivil.Text = ""
    tIglesia.Text = ""
    labDireccion.Caption = ""
    chCerrado.Value = 0
    labUsuario.Caption = ""
    labDirCliente1.Caption = ""
    labDirCliente2.Caption = ""
    lEmailCli1.Caption = "": lEmailCli1.Tag = ""
    lEmailCli2.Caption = "": lEmailCli2.Tag = ""
    aDireccion = 0
    tClave.Text = ""
    tComentario.Text = ""
End Sub
'------------------------------------------------------------
'No dejo acceder a los datos.
'------------------------------------------------------------
Sub OcultoCampos()
    tCodigo.Enabled = True: tCodigo.BackColor = vbWhite
    tCICliente2.Enabled = False
    tNombreColectivo.Enabled = False: tNombreColectivo.BackColor = Inactivo
    tCivil.Enabled = False: tCivil.BackColor = Inactivo
    tIglesia.Enabled = False: tIglesia.BackColor = Inactivo
    bDireccion.Enabled = False
    chCerrado.Enabled = False
    tClave.Enabled = False: tClave.BackColor = Inactivo
    bEmail1.Enabled = False
    bEmail2.Enabled = False
    tComentario.Enabled = False: tComentario.BackColor = Inactivo
End Sub
'------------------------------------------------------------
'Dejo libre los objetos.
'------------------------------------------------------------
Sub HabilitoCampos()
    tCodigo.Enabled = False: tCodigo.BackColor = Inactivo
    'tCiCliente1.Enabled = True: tCiCliente1.BackColor = obligatorio
    tCICliente2.Enabled = True: tCICliente2.BackColor = Obligatorio
    tNombreColectivo.Enabled = True: tNombreColectivo.BackColor = Obligatorio
    tCivil.Enabled = True: tCivil.BackColor = vbWindowBackground
    tIglesia.Enabled = True: tIglesia.BackColor = vbWindowBackground
    bDireccion.Enabled = True
    chCerrado.Enabled = True
    tClave.Enabled = True: tClave.BackColor = vbWindowBackground
    bEmail1.Enabled = True
    bEmail2.Enabled = True
    tComentario.Enabled = True: tComentario.BackColor = vbWindowBackground
End Sub

'------------------------------------------------------------
'Presiono boton grabar, la misma puede ser un alta o una modificación.
'------------------------------------------------------------
Sub AccionGrabar()
    Ayuda "Grabando........."
    If ControlesGrabar Then
        If sNuevo Then
            If MsgBox("¿Confirma el alta del colectivo?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then NuevoColectivo
        Else
            If MsgBox("¿Confirma la modificación de datos?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then ModificoColectivo
        End If
    End If
    Ayuda ""
End Sub
Private Sub NuevoColectivo()
On Error GoTo ErrNC
    
    Screen.MousePointer = 11
    Cons = "Select * From Colectivo Where ColCliente1 = " & CLng(tCiCliente1.Tag) _
        & " Or ColCliente2 = " & CLng(tCICliente2.Tag) & " Or ColCliente1 = " & CLng(tCICliente2.Tag) _
        & " Or ColCliente2 = " & CLng(tCiCliente1.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.AddNew
        CargoCamposEnBD
        RsAux.Update
        RsAux.Close
        Cons = "Select Max(ColCodigo) From Colectivo Where ColCliente1 = " & CLng(tCiCliente1.Tag) _
            & " And ColCliente2 = " & CLng(tCICliente2.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        tCodigo.Text = RsAux(0)
        RsAux.Close
        OcultoCampos
        BuscoColectivo tCodigo.Text
    Else
        If aDireccion > 0 Then
            Cons = "Delete Direccion Where DirCodigo = " & aDireccion
            cBase.Execute (Cons)
        End If
        MsgBox "Ya existe un colectivo ingresado para esos clientes, el código del mismo es " & RsAux!ColCodigo & ".", vbExclamation, "ATENCIÓN"
        RsAux.Close
        AccionCancelar
    End If
    sNuevo = False
    Screen.MousePointer = 0
    Exit Sub
ErrNC:
    clsGeneral.OcurrioError "Ocurrio un error al dar el alta.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub ModificoColectivo()
On Error GoTo ErrMC
    Screen.MousePointer = 11
    Cons = "Select * From Colectivo Where ColCodigo = " & tCodigo.Text
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux!ColFModificacion = CDate(tCodigo.Tag) Then
            RsAux.Edit
            CargoCamposEnBD
            RsAux.Update
            RsAux.Close
            labUsuario.Caption = BuscoUsuario(miconexion.UsuarioLogueado(True), True)
            OcultoCampos
            Botones True, True, True, False, False, Toolbar1, Me
        Else
            RsAux.Close
            MsgBox "Otra terminal modifico la ficha con anterioridad, verifique.", vbExclamation, "ATENCIÓN"
            tCodigo.Text = ""
            AccionCancelar
        End If
    Else
        RsAux.Close
        MsgBox "El colectivo fue eliminado, verifique.", vbExclamation, "ATENCIÓN"
        tCodigo.Text = ""
        AccionCancelar
    End If
    sModificar = False
    Screen.MousePointer = 0
    Exit Sub
ErrMC:
    clsGeneral.OcurrioError "Ocurrio un error al intentar modificar el colectivo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub CargoCamposEnBD()
Dim sComentario As String
    FechaDelServidor
    RsAux!ColCliente1 = tCiCliente1.Tag
    RsAux!ColCliente2 = tCICliente2.Tag
    RsAux!ColNombre = Trim(tNombreColectivo.Text)
    If IsDate(tCivil.Text) Then RsAux!ColFechaCivil = Format(tCivil.Text, sqlFormatoFH) Else RsAux!ColFechaCivil = Null
    If IsDate(tIglesia.Text) Then RsAux!ColFechaIglesia = Format(tIglesia.Text, sqlFormatoFH) Else RsAux!ColFechaIglesia = Null
    If aDireccion > 0 Then RsAux!ColDirEnvio = aDireccion Else RsAux!ColDirEnvio = Null
    If chCerrado.Value = 0 Then RsAux!ColCerrado = False Else RsAux!ColCerrado = True
    RsAux!ColUsuario = miconexion.UsuarioLogueado(True)
    If Trim(tClave.Text) <> "" Then RsAux!ColClave = tClave.Text Else RsAux!ColClave = Null
    RsAux!ColFModificacion = Format(gFechaServidor, sqlFormatoFH)
    sComentario = Replace(tComentario.Text, vbCrLf, " ")
    If Trim(sComentario) <> "" Then RsAux!ColComentario = sComentario Else RsAux!ColComentario = Null
End Sub
Sub AccionEliminar()
On Error GoTo ErrBE
    
    If MsgBox("¿Confirma la baja del Colectivo?", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        Screen.MousePointer = vbHourglass
        'Veo si tiene cuentas ingresadas
        Cons = "Select * From CuentaDocumento Where CDoTipo = " & Cuenta.Colectivo _
            & " And CDoIDTipo = " & tCodigo.Text
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not RsAux.EOF Then
            RsAux.Close: Screen.MousePointer = 0
            MsgBox "El colectivo tiene asociados documentos no podrá eliminarlo.", vbInformation, "ATENCIÓN"
            Exit Sub
        End If
        'Borro al empleado
        Cons = "Delete Colectivo Where ColCodigo = " & tCodigo.Text
        cBase.Execute (Cons)
        'Restauro formulario en espera de uno nuevo.
        Botones True, False, False, False, False, Toolbar1, Me
        LimpioCampos
        OcultoCampos
        Screen.MousePointer = vbDefault
    End If
    Exit Sub

ErrBE:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al intentar eliminar el colectivo.", Err.Description
End Sub

Private Sub tCiCliente1_Change()
    tCiCliente1.Tag = "0": labCliente1.Caption = "": labCliente1.Tag = ""
End Sub
Private Sub tCiCliente1_GotFocus()
    tCiCliente1.SelStart = 0: tCiCliente1.SelLength = 11
    Ayuda "Ingrese la cédula de uno de los integrantes del colectivo."
End Sub
Private Sub tCiCliente1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            tCiCliente1.Tag = "0": labCliente1.Caption = "": labCliente1.Tag = ""
            If clsGeneral.CedulaValida(tCiCliente1.Text) Then
                BuscoClientePorCedula tCiCliente1, 1
                If Not (sNuevo Or sModificar) Then BuscoColectivosPorCliente
            Else
                MsgBox "La cédula ingresada no es válida.", vbExclamation, "ATENCIÓN"
            End If
            If CLng(tCiCliente1.Tag) > 0 Then Foco tCICliente2
        Case vbKeyF4
            AccionBuscarCliente 1
    End Select
End Sub
Private Sub tCiCliente1_LostFocus()
    tCiCliente1.SelStart = 0
    Ayuda ""
End Sub

Private Sub tCICliente2_Change()
    tCICliente2.Tag = "0": labCliente2.Caption = "": labCliente2.Tag = ""
End Sub
Private Sub tCICliente2_GotFocus()
    tCICliente2.SelStart = 0: tCICliente2.SelLength = 11
    Ayuda "Ingrese la cédula del otro integrante del colectivo."
End Sub
Private Sub tCICliente2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            tCICliente2.Tag = "0": labCliente2.Caption = "": labCliente2.Tag = ""
            If clsGeneral.CedulaValida(tCICliente2.Text) Then
                BuscoClientePorCedula tCICliente2, 2
            Else
                MsgBox "La cédula ingresada no es válida.", vbExclamation, "ATENCIÓN"
            End If
            If CLng(tCICliente2.Tag) > 0 Then
                If Trim(tNombreColectivo.Text) = "" Then tNombreColectivo.Text = labCliente1.Tag & " " & labCliente2.Tag
                Foco tNombreColectivo
            End If
        Case vbKeyF4
            AccionBuscarCliente 2
            If CLng(tCICliente2.Tag) > 0 Then
                If Trim(tNombreColectivo.Text) = "" Then tNombreColectivo.Text = labCliente1.Tag & " " & labCliente2.Tag
                Foco tNombreColectivo
            End If
    End Select
End Sub
Private Sub tCICliente2_LostFocus()
    tCICliente2.SelStart = 0
    Ayuda ""
End Sub
Private Sub tCivil_GotFocus()
    With tCivil
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese la fecha del civil."
End Sub
Private Sub tCivil_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tIglesia
End Sub
Private Sub tCivil_LostFocus()
    Ayuda ""
    If IsDate(tCivil.Text) Then tCivil.Text = Format(tCivil.Text, "dd-Mmm yyyy hh:mm") Else tCivil.Text = ""
End Sub

Private Sub tClave_GotFocus()
    With tClave
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese una clave para el colectivo."
End Sub
Private Sub tClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chCerrado.SetFocus
End Sub
Private Sub tClave_LostFocus()
    Ayuda ""
End Sub

Private Sub tCodigo_Change()
    bDireccion.Enabled = False
    LimpioCampos
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese el código del colectivo."
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrBC
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCodigo.Text) Then
            BuscoColectivo tCodigo.Text
        ElseIf Trim(tCodigo.Text) <> "" Then
            bDireccion.Enabled = False
            Botones True, False, False, False, False, Toolbar1, Me
            LimpioCampos
            MsgBox "El código ingresado no es válido.", vbExclamation, "ATENCIÓN"
        End If
    End If
    Exit Sub
ErrBC:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el colectivo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCodigo_LostFocus()
    Ayuda ""
    tCodigo.SelStart = 0
End Sub

Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tIglesia_GotFocus()
    With tIglesia
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese la fecha de la boda por iglesia."
End Sub

Private Sub tIglesia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco bDireccion
End Sub

Private Sub tIglesia_LostFocus()
    tIglesia.SelStart = 0
    Ayuda ""
    If IsDate(tIglesia.Text) Then tIglesia.Text = Format(tIglesia.Text, "dd-Mmm yyyy hh:mm") Else tIglesia.Text = ""
End Sub

Private Sub tNombreColectivo_GotFocus()
    With tNombreColectivo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese un nombre para el colectivo."
End Sub
Private Sub tNombreColectivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tNombreColectivo.Text) <> "" Then Foco tCivil
    End If
End Sub
Private Sub tNombreColectivo_LostFocus()
    Ayuda ""
    tNombreColectivo.SelStart = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select
End Sub
Private Sub BuscoClientePorCedula(strCedula As String, Cliente As Byte)
'Cliente me dice cual de los dos objetos lo invoca : tcicliente1 = 1 y tcicliente2 = 2
On Error GoTo ErrBC
    
    Screen.MousePointer = 11
    Cons = "Select CliCodigo, CliDireccion, CPeApellido1, Nombre = RTRIM(RTRIM(CPeApellido1) + ' ' + RTRIM(CPeApellido2)) + ', ' +  RTRIM(RTRIM(CPeNombre1) + ' ' + RTRIM(CPeNombre2)), CPeSexo " _
        & " From Cliente, CPersona " _
        & " Where CLiCIRUC = '" & strCedula & "'" _
        & " And CliCodigo = CPeCliente "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un cliente ingresado con esa cédula.", vbInformation, "ATENCIÓN"
        If Not sNuevo And Not sModificar Then
            LimpioCampos
            Botones True, False, False, False, False, Toolbar1, Me
        End If
        If Cliente = 1 Then
            tCiCliente1.Tag = "0"
            labCliente1.Caption = " "
            labCliente1.Tag = ""
        Else
            tCICliente2.Tag = "0"
            labCliente2.Caption = " "
            labCliente2.Tag = ""
        End If
    Else
        If Cliente = 1 Then
            tCiCliente1.Tag = RsAux!CliCodigo
            labCliente1.Caption = " " & Trim(RsAux!Nombre)
            labCliente1.Tag = Trim(RsAux!CPeApellido1)
            If Not IsNull(RsAux!CliDireccion) Then labDirCliente1.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, False, True, False, True)
            If Not IsNull(RsAux!CPeSexo) Then lEmailCli1.Tag = Trim(RsAux!CPeSexo)
        Else
            tCICliente2.Tag = RsAux!CliCodigo
            labCliente2.Caption = " " & Trim(RsAux!Nombre)
            labCliente2.Tag = Trim(RsAux!CPeApellido1)
            If Not IsNull(RsAux!CliDireccion) Then labDirCliente1.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, False, True, False, True)
            If Not IsNull(RsAux!CPeSexo) Then lEmailCli2.Tag = Trim(RsAux!CPeSexo)
        End If
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    clsGeneral.OcurrioError "Ocurrio un error al buscar al cliente por cédula.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ControlesGrabar() As Boolean
    ControlesGrabar = False
    If CLng(tCiCliente1.Tag) = 0 Then
        MsgBox "No se selecciono un cliente.", vbExclamation, "ATENCIÓN"
        Foco tCiCliente1: Exit Function
    End If
    If CLng(tCICliente2.Tag) = 0 Then
        MsgBox "No se selecciono un cliente.", vbExclamation, "ATENCIÓN"
        Foco tCICliente2: Exit Function
    End If
    If Trim(tNombreColectivo.Text) = "" Then
        MsgBox "Se debe ingresar un nombre para el colectivo.", vbExclamation, "ATENCIÓN"
        Foco tNombreColectivo: Exit Function
    End If
    If Trim(tCivil.Text) <> "" And Not IsDate(tCivil.Text) Then
        MsgBox "La fecha ingresada no es válida.", vbExclamation, "ATENCIÓN"
        Foco tCivil: Exit Function
    End If
    If Trim(tIglesia.Text) <> "" And Not IsDate(tIglesia.Text) Then
        MsgBox "La fecha ingresada no es válida.", vbExclamation, "ATENCIÓN"
        Foco tIglesia: Exit Function
    End If
    'Válido Dirección EMAIL.
    If Trim(lEmailCli1.Caption) = "" And Trim(lEmailCli2.Caption) = "" Then
        If MsgBox("Ninguno de los integrantes del colectivo tiene Email ingresado." _
            & vbCrLf & "El email es muy importante para mantenerlos informados acerca de los saldos." _
            & vbCrLf & "¿Desea ingresar una dirección de Email?", vbQuestion + vbYesNo, "IMPORTANTE") = vbYes Then
            
            If UCase(lEmailCli1.Tag) = "M" Then
                bEmail1_Click
            Else
                bEmail2_Click
            End If
            Exit Function
        End If
        
    End If
    ControlesGrabar = True
End Function

Private Sub BuscoColectivo(IDColectivo As Long)
On Error GoTo ErrBC
    Screen.MousePointer = 11
    LimpioCampos
    Cons = "Select Colectivo.*,  Dir1 = C1.CliDireccion , Dir2 = C2.CliDireccion, Ced1 = C1.CliCIRUC, Ced2 = C2.CliCIRUC, Nom1 = RTRIM(RTRIM(CP1.CPeApellido1) + ' ' + RTRIM(CP1.CPeApellido2)) + ', ' +  RTRIM(RTRIM(CP1.CPeNombre1) + ' ' + RTRIM(CP1.CPeNombre2)) " _
                        & " , Nom2 = RTRIM(RTRIM(CP2.CPeApellido1) + ' ' + RTRIM(CP2.CPeApellido2)) + ', ' +  RTRIM(RTRIM(CP2.CPeNombre1) + ' ' + RTRIM(CP2.CPeNombre2)) " _
                        & ", Sex1 = CP1.CPeSexo, Sex2 = CP2.CPeSexo" _
        & " From Colectivo, Cliente C1, CPersona CP1, Cliente C2, CPersona CP2 " _
        & " Where ColCodigo = " & IDColectivo _
        & " And ColCliente1 = C1.CliCodigo And C1.CliCodigo = CP1.CPeCliente " _
        & " And ColCliente2 = C2.CliCodigo And C2.CliCodigo = CP2.CPeCliente "
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        Botones True, False, False, False, False, Toolbar1, Me
        MsgBox "No existe un colectivo con ese código.", vbExclamation, "ATENCIÓN"
    Else
        Botones True, True, True, False, False, Toolbar1, Me
        tCodigo.Text = RsAux!ColCodigo
        tCodigo.Tag = RsAux!ColFModificacion
        If Not IsNull(RsAux!Ced1) Then tCiCliente1.Text = RsAux!Ced1
        tCiCliente1.Tag = RsAux!ColCliente1
        labCliente1.Caption = " " & Trim(RsAux!Nom1)
        If Not IsNull(RsAux!Dir1) Then labDirCliente1.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!Dir1, False, True, False, True)
        If Not IsNull(RsAux!Ced2) Then tCICliente2.Text = RsAux!Ced2
        tCICliente2.Tag = RsAux!ColCliente2
        labCliente2.Caption = " " & Trim(RsAux!Nom2)
        If Not IsNull(RsAux!Dir2) Then labDirCliente2.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!Dir2, False, True, False, True)
        tNombreColectivo.Text = Trim(RsAux!ColNombre)
        If Not IsNull(RsAux!ColFechaCivil) Then tCivil.Text = Format(RsAux!ColFechaCivil, "dd-Mmm yyyy hh:mm")
        If Not IsNull(RsAux!ColFechaIglesia) Then tIglesia.Text = Format(RsAux!ColFechaIglesia, "dd-Mmm yyyy hh:mm")
        If Not IsNull(RsAux!ColClave) Then tClave.Text = Trim(RsAux!ColClave)
        If Not IsNull(RsAux!ColComentario) Then tComentario.Text = Trim(RsAux!ColComentario)
        If RsAux!ColCerrado Then chCerrado.Value = 1
        labUsuario.Caption = BuscoUsuario(RsAux!ColUsuario, True)
        If Not IsNull(RsAux!Sex1) Then lEmailCli1.Tag = Trim(RsAux!Sex1)
        If Not IsNull(RsAux!Sex2) Then lEmailCli2.Tag = Trim(RsAux!Sex2)
        'Campos de Direccion
        bDireccion.Enabled = True
        If Not IsNull(RsAux!ColDirEnvio) Then
            aDireccion = RsAux!ColDirEnvio
            labDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, aDireccion, False, True, False, True)
        Else
            aDireccion = 0
        End If
        RsAux.Close
        'Busco email.
        lEmailCli1.Caption = CargoDireccionEMail(Val(tCiCliente1.Tag))
        lEmailCli2.Caption = CargoDireccionEMail(Val(tCICliente2.Tag))
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el colectivo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Ayuda(strTexto As String)
    Status.Panels("Msg").Text = strTexto
End Sub

Private Sub AccionBuscarCliente(Cliente As Byte)
On Error GoTo ErrBC
    Dim frmBusco As New clsBuscarCliente
    Screen.MousePointer = 11
    frmBusco.ActivoFormularioBuscarClientes cBase, True
    Me.Refresh
    Screen.MousePointer = 11
    If frmBusco.BCClienteSeleccionado > 0 And frmBusco.BCTipoClienteSeleccionado = TipoCliente.Cliente Then
        
        Cons = "Select CliCodigo, CliCIRUC, CliDireccion, CPeApellido1, Nombre = RTRIM(RTRIM(CPeApellido1) + ' ' + RTRIM(CPeApellido2)) + ', ' +  RTRIM(RTRIM(CPeNombre1) + ' ' + RTRIM(CPeNombre2)), CPeSexo " _
            & " From Cliente, CPersona " _
            & " Where CLiCodigo = " & frmBusco.BCClienteSeleccionado _
            & " And CliCodigo = CPeCliente "
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        
        If RsAux.EOF Then
            RsAux.Close
            MsgBox "No existe un cliente ingresado con esa cédula.", vbInformation, "ATENCIÓN"
            If Cliente = 1 Then
                tCiCliente1.Text = ""
                tCiCliente1.Tag = "0"
                labCliente1.Caption = " "
                labCliente1.Tag = ""
            Else
                tCICliente2.Text = ""
                tCICliente2.Tag = "0"
                labCliente2.Caption = " "
                labCliente2.Tag = ""
            End If
        Else
            If Cliente = 1 Then
                If Not IsNull(RsAux!CliCIRUC) Then tCiCliente1.Text = RsAux!CliCIRUC Else tCiCliente1.Text = ""
                tCiCliente1.Tag = RsAux!CliCodigo
                labCliente1.Caption = " " & Trim(RsAux!Nombre)
                labCliente1.Tag = Trim(RsAux!CPeApellido1)
                If Not IsNull(RsAux!CliDireccion) Then labDirCliente1.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, False, True, False, True)
                If Not IsNull(RsAux!CPeSexo) Then lEmailCli1.Tag = Trim(RsAux!CPeSexo)
                RsAux.Close
                If sNuevo Or sModificar Then
                    Foco tCICliente2
                Else
                    BuscoColectivosPorCliente
                End If
            Else
                If Not IsNull(RsAux!CliCIRUC) Then tCICliente2.Text = RsAux!CliCIRUC Else tCICliente2.Text = ""
                tCICliente2.Tag = RsAux!CliCodigo
                labCliente2.Caption = " " & Trim(RsAux!Nombre)
                labCliente2.Tag = Trim(RsAux!CPeApellido1)
                If Not IsNull(RsAux!CliDireccion) Then labDirCliente2.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, False, True, False, True)
                If Not IsNull(RsAux!CPeSexo) Then lEmailCli2.Tag = Trim(RsAux!CPeSexo)
                RsAux.Close
            End If
        End If
    End If
    Set frmBusco = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el cliente.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoColectivosPorCliente()
On Error GoTo ErrBCPC

    If Val(tCiCliente1.Tag) > 0 Then
        Screen.MousePointer = 11
        Cons = "Select * From Colectivo Where ColCliente1 = " & CLng(tCiCliente1.Tag) _
            & " Or ColCliente2 = " & CLng(tCiCliente1.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If RsAux.EOF Then
            RsAux.Close
            MsgBox "No hay colectivos para ese cliente.", vbExclamation, "ATENCIÓN"
            LimpioCampos
            tCodigo.Text = "": tCodigo.Tag = "0": Botones True, False, False, False, False, Toolbar1, Me
        Else
            tCodigo.Tag = RsAux!ColCodigo
            RsAux.Close
            BuscoColectivo (tCodigo.Tag)
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrBCPC:
    clsGeneral.OcurrioError "Ocurrio un error al buscar colectivos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function CargoDireccionEMail(ByVal lCliente As Long) As String
    
    CargoDireccionEMail = ""
    Cons = "Select * From EMailDireccion, EMailServer " _
        & " Where EMDIDCliente = " & lCliente _
        & " And EMDServidor = EMSCodigo"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        If CargoDireccionEMail = "" Then
            CargoDireccionEMail = Trim(RsAux!EMDDireccion) & "@" & Trim(RsAux!EMSDireccion)
        Else
            CargoDireccionEMail = CargoDireccionEMail & "; ..."
            Exit Do
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
        
End Function
