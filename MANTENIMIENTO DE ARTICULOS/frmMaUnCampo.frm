VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MaUnCampo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4335
   ClientLeft      =   1245
   ClientTop       =   2295
   ClientWidth     =   4335
   Icon            =   "frmMaUnCampo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4335
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
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
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1850
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView lMotivo 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4471
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   -1
         Key             =   "cDescripcion"
         Text            =   "Descripción"
         Object.Width           =   5733
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4080
      Width           =   4335
      _ExtentX        =   7646
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
   Begin VB.TextBox TDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MaxLength       =   30
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   240
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
            Picture         =   "frmMaUnCampo.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaUnCampo.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaUnCampo.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaUnCampo.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaUnCampo.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaUnCampo.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaUnCampo.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaUnCampo.frx":0DC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Ingresados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Descripción:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Menu MnuOpcion 
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
      Begin VB.Menu linea1 
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
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "MaUnCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mantenimiento UnCampo ----------------------------------------------------------------------
'
'   Este mantenimiento puede ser utilizado para realizar un ABM en una tabla que
'   tenga un solo campo texto y un código autonumérico.
'
'   Se requiere cargar algunas variables locales que indiquen el nombre del campo
'   código y el nombre del campo texto para realizar búsquedas.
'
'--------------------------------------------------------------------------------------------------------
Option Explicit

Dim sNuevo As Boolean
Dim sModificar As Boolean
Dim aSeleccion As String

Private bTipoLlamado As Byte
Private lSeleccionado As Long

' Definición de variables que requieren el mantenimiento. --------------------------------
Private NomTabla As String                              'Nombre de la tabla a grabar.
Private NomCampoCodigo As String                'Nombre del campo código de la tabla.
Private NomCampoTexto As String                   'Nombre del campo texto (nombre, descripción, etc.)

'Generalmente el nombre de la tabla referencia al sustantivo del formulario. PERO...
Private NomSustantivo As String

Private Rs1Campo As rdoResultset

Private Sub AccionModificar()

    On Error GoTo errModificar
    If lMotivo.SelectedItem <> "" Then
        
        Rs1Campo.Close
        Cons = "SELECT * FROM " & NomTabla & " WHERE " & NomCampoCodigo & " = " & Mid(lMotivo.SelectedItem.Key, 2, Len(lMotivo.SelectedItem.Key))
        Set Rs1Campo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
        sModificar = True
        Call Botones(False, False, False, True, True, Toolbar1, Me)
        
        HabilitoIngreso
        TDescripcion.Text = Trim(Rs1Campo(NomCampoTexto))
        TDescripcion.SetFocus

    End If
    Exit Sub

errModificar:
    clsGeneral.OcurrioError "Ha ocurrido un error al cargar los datos para modificarlos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
On Error GoTo ErrActivo
    DoEvents
    Screen.MousePointer = 0
    Rs1Campo.Requery
    If bTipoLlamado = 3 And TDescripcion.Enabled Then TDescripcion.SetFocus
    Exit Sub
ErrActivo:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error inesperado al activar el formulario.", Err.Description
End Sub

Private Sub Form_Load()

    SetearLView lvValores.UnClickIcono, lMotivo
    
    sNuevo = False
    sModificar = False
    aSeleccion = ""
    
    Cons = "SELECT * FROM " & NomTabla & " WHERE " & NomCampoCodigo & " = 0"
    Set Rs1Campo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    CargoLista
    DeshabilitoIngreso
    If pTipoLlamado = 3 Then AccionNuevo
    
End Sub


Private Sub CargoLista()

Dim aIndice As Long

    On Error GoTo ErrCC
    aIndice = 0
    
    Screen.MousePointer = 11
    Rs1Campo.Close
    Cons = "SELECT * FROM " & NomTabla _
            & " Order by " & NomCampoTexto
    
    Set Rs1Campo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    lMotivo.ListItems.Clear
    If Rs1Campo.EOF Then
        TDescripcion.Text = ""
        lSeleccionado = 0
        Call Botones(True, False, False, False, False, Toolbar1, Me)
    Else
        Dim itmx As ListItem
        Do While Not Rs1Campo.EOF
            Set itmx = lMotivo.ListItems.Add(, "A" & Trim(Rs1Campo(NomCampoCodigo)), Trim(Rs1Campo(NomCampoTexto))) 'cod. y detalle
            If Trim(Rs1Campo(NomCampoTexto)) = Trim(aSeleccion) Then
                aIndice = lMotivo.ListItems(lMotivo.ListItems.Count).Index
            End If
            Rs1Campo.MoveNext
        Loop
        Rs1Campo.MoveFirst
        
        Call Botones(True, True, True, False, False, Toolbar1, Me)
        
    End If
    
    If aIndice <> 0 Then
        lMotivo.SelectedItem = lMotivo.ListItems(aIndice)
        lSeleccionado = Mid(lMotivo.SelectedItem.Key, 2, Len(lMotivo.SelectedItem.Key))
    End If
    Screen.MousePointer = 0
    
    Exit Sub
    
ErrCC:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos ingresados.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = ""

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If (sNuevo Or sModificar) And Trim(TDescripcion.Text) <> "" Then
        If MsgBox("¿Desea abandonar sin almacenar los datos ingresados.?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
            AccionGrabar
            If sNuevo Or sModificar Then Cancel = True
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    Rs1Campo.Close
    Forms(Forms.Count - 2).SetFocus

End Sub

Private Sub Label1_Click()

    If TDescripcion.Enabled Then
        TDescripcion.SelStart = 0
        TDescripcion.SelLength = Len(TDescripcion.Text)
        TDescripcion.SetFocus
    End If

End Sub

Private Sub lMotivo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = " Lista de datos ingresados."

End Sub

Private Sub MnuCancelar_Click()

    sNuevo = False
    DeshabilitoIngreso
    ArmoPantalla
    
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
    TDescripcion.SetFocus
    
End Sub

Private Sub MnuVolver_Click()

    Unload Me
    
End Sub

Private Sub TDescripcion_GotFocus()

    TDescripcion.SelStart = 0
    TDescripcion.SelLength = Len(TDescripcion.Text)

End Sub

Private Sub TDescripcion_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(TDescripcion.Text) <> "" Then
            AccionGrabar
        End If
    End If

End Sub

Private Sub TDescripcion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = " Ingrese una descripción."

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As Button)
 
    Select Case Button.Key
    
        Case "nuevo"
            AccionNuevo
            TDescripcion.SetFocus
            
        Case "modificar"
            AccionModificar
            
        Case "grabar"
            AccionGrabar
        
        Case "eliminar"
            AccionEliminar
            
        Case "cancelar"
            AccionCancelar
            
        Case "salir"
            Unload Me
            
    End Select
    
End Sub

Private Sub AccionCancelar()

    sNuevo = False
    sModificar = False
    DeshabilitoIngreso
    ArmoPantalla
            
End Sub
Private Sub AccionNuevo()

    sNuevo = True
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    
    HabilitoIngreso
            
End Sub


Private Sub AccionGrabar()

    If Not ValidoCampos Then
        Exit Sub
    End If
    
    'Ingresa un nuevo dato.
    If sNuevo Then
        If MsgBox("Confirma el alta de datos ingresados.", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
            On Error GoTo errorGrabar
            Rs1Campo.AddNew
            Rs1Campo(NomCampoTexto) = Trim(TDescripcion.Text)
            Rs1Campo.Update
            Rs1Campo.Requery
            sNuevo = False
            
            aSeleccion = Trim(TDescripcion.Text)
            CargoLista                  'Cargo la lista para obtener el seleccionado.
            
            If bTipoLlamado = 3 Then 'TipoLlamado.IngresoNuevo Then
                Unload Me
                Exit Sub
            Else
                If MsgBox("¿Desea continuar ingresando datos?", vbQuestion + vbYesNo + vbDefaultButton2, "NUEVO") = vbYes Then
                    'La cargo para que presente en la lista el último ingresado.
                    AccionNuevo
                    Exit Sub
                End If
            End If
            
        Else
            Exit Sub
        End If
    Else
    'Modifica los datos
        If MsgBox("Confirma la modificación de datos.", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
            On Error GoTo errorGrabar
            Rs1Campo.Edit
            Rs1Campo(NomCampoTexto) = Trim(TDescripcion.Text)
            Rs1Campo.Update
            Rs1Campo.Requery
            sModificar = False
        Else
            Exit Sub
        End If
    End If
    
    aSeleccion = Trim(TDescripcion.Text)
    CargoLista
    ArmoPantalla
    DeshabilitoIngreso
    Exit Sub
    
errorGrabar:
    clsGeneral.OcurrioError "Ocurrió un error al grabar la información. Reintente la operación."
    Rs1Campo.Requery
    Screen.MousePointer = 0
End Sub

Private Function ValidoCampos()

    ValidoCampos = True
    
    'Hay datos ingresados.
    If Len(TDescripcion.Text) = 0 Then
        MsgBox "Se debe ingresar una descripción o nombre.", vbExclamation, "ATENCIÓN"
        TDescripcion.SetFocus
        ValidoCampos = False
        Exit Function
    End If
    
    If Not clsGeneral.TextoValido(TDescripcion.Text) Then
        MsgBox "Se han igresado caracteres no válidos, verifique.", vbInformation, "ATENCIÓN"
        ValidoCampos = False
        Exit Function
    End If
    
End Function

Private Sub ArmoPantalla()

    If lMotivo.ListItems.Count > 0 Then
        Call Botones(True, True, True, False, False, Toolbar1, Me)
    Else
        Call Botones(True, False, False, False, False, Toolbar1, Me)
    End If
    
End Sub

Private Sub DeshabilitoIngreso()

    lMotivo.Enabled = True
    TDescripcion.Text = ""
    TDescripcion.Enabled = False
    TDescripcion.BackColor = vbButtonFace

End Sub

Private Sub HabilitoIngreso()

    lMotivo.Enabled = False
    TDescripcion.Text = ""
    TDescripcion.Enabled = True
    TDescripcion.BackColor = &HC0FFFF

End Sub

Public Sub AccionEliminar()

    If lMotivo.SelectedItem.Index = -1 Then
        MsgBox "Se debe seleccionar un concepto de la lista.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If

    On Error GoTo errorEliminar
    If MsgBox("Confirma la baja del " & NomSustantivo & " '" & Trim(lMotivo.SelectedItem) & "'", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        Screen.MousePointer = vbHourglass
        
        Rs1Campo.Close
        Cons = "SELECT * FROM " & NomTabla & " WHERE " & NomCampoCodigo & "= " & Mid(lMotivo.SelectedItem.Key, 2, Len(lMotivo.SelectedItem.Key))
        Set Rs1Campo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Rs1Campo.Delete
        
        CargoLista
        ArmoPantalla
        DeshabilitoIngreso
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errorEliminar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido eliminar el concepto.", Err.Description
    CargoLista
    
End Sub

'Nombre de la tabla en la que se va a trabajar------------------------------------
Public Property Get pTabla() As String

    pTabla = NomTabla

End Property

Public Property Let pTabla(Texto As String)

    NomTabla = Texto
    
End Property

Public Property Get pTipoLlamado() As Byte

    pTipoLlamado = bTipoLlamado

End Property

Public Property Let pTipoLlamado(Codigo As Byte)

    bTipoLlamado = Codigo

End Property

'-----------------------------------------------------------------------------------------
'Nombre del Campo Código
Public Property Get pCampoCodigo() As String

    pCampoCodigo = NomCampoCodigo

End Property

Public Property Let pCampoCodigo(Texto As String)

    NomCampoCodigo = Texto

End Property

'-----------------------------------------------------------------------------------------
'Nombre del Campo Descripcion
Public Property Get pCampoNombre() As String

    pCampoNombre = NomCampoTexto

End Property

Public Property Let pCampoNombre(Texto As String)

    NomCampoTexto = Texto

End Property
'-----------------------------------------------------------------------------------------
'Nombre del sustantivo (valor a desplegar en el LEBEL)
Public Property Get pSustantivo() As String

    pSustantivo = NomSustantivo

End Property

Public Property Let pSustantivo(Texto As String)

    NomSustantivo = Texto

End Property
'-----------------------------------------------------------------------------------------

Public Property Get pSeleccionado() As Long

    pSeleccionado = lSeleccionado
    
End Property
Public Property Let pSeleccionado(Codigo As Long)

    lSeleccionado = Codigo

End Property

