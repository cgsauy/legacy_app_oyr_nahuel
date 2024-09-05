VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmReglas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5745
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReglasMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin AACombo99.AACombo cAQuien 
      Height          =   315
      Left            =   960
      TabIndex        =   19
      Top             =   3960
      Width           =   3495
      _ExtentX        =   6165
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
   End
   Begin VB.CommandButton bCuenta 
      Caption         =   "Ver cuentas"
      Height          =   315
      Left            =   7200
      TabIndex        =   6
      Top             =   840
      Width           =   1035
   End
   Begin VB.CheckBox chBorrar 
      Alignment       =   1  'Right Justify
      Caption         =   "&Borrar Mail"
      Height          =   195
      Left            =   4980
      TabIndex        =   24
      Top             =   5100
      Width           =   2055
   End
   Begin AACombo99.AACombo cPlantilla 
      Height          =   315
      Left            =   960
      TabIndex        =   23
      Top             =   5040
      Width           =   3255
      _ExtentX        =   5741
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
   End
   Begin VB.CommandButton bAQuien 
      Caption         =   "Vista Previa"
      Height          =   315
      Left            =   7200
      TabIndex        =   21
      Top             =   4680
      Width           =   1035
   End
   Begin VB.TextBox tAQuien 
      Height          =   675
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4320
      Width           =   6135
   End
   Begin VB.TextBox tCondicion 
      Height          =   285
      Left            =   960
      TabIndex        =   17
      Top             =   3660
      Width           =   6135
   End
   Begin VB.CommandButton bCuerpo 
      Caption         =   "Validar"
      Height          =   315
      Left            =   7200
      TabIndex        =   15
      Top             =   3300
      Width           =   1035
   End
   Begin VB.TextBox tCuerpo 
      Height          =   735
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2880
      Width           =   6135
   End
   Begin VB.CommandButton bAsunto 
      Caption         =   "Validar"
      Height          =   315
      Left            =   7200
      TabIndex        =   12
      Top             =   2520
      Width           =   1035
   End
   Begin VB.TextBox tAsunto 
      Height          =   735
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2100
      Width           =   6135
   End
   Begin VB.TextBox tCuenta 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   840
      Width           =   6075
   End
   Begin VB.TextBox tPrioridad 
      Height          =   285
      Left            =   6060
      MaxLength       =   4
      TabIndex        =   3
      Top             =   540
      Width           =   555
   End
   Begin VB.TextBox tNombre 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   540
      Width           =   2955
   End
   Begin VB.CommandButton bDePrevia 
      Caption         =   "Vista Previa"
      Height          =   315
      Left            =   7200
      TabIndex        =   9
      Top             =   1740
      Width           =   1035
   End
   Begin VB.TextBox tDe 
      Height          =   915
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1140
      Width           =   6135
   End
   Begin MSComctlLib.ImageList imgIconos 
      Left            =   6900
      Top             =   -60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReglasMail.frx":0442
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReglasMail.frx":075E
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReglasMail.frx":087E
            Key             =   "modificar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReglasMail.frx":09DA
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReglasMail.frx":0AEE
            Key             =   "grabar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReglasMail.frx":0C02
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReglasMail.frx":0D16
            Key             =   "cuentas"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Abandonar el formulario"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Acción Nuevo"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Acción Modificar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "eliminar"
            Object.ToolTipText     =   "Acción Eliminar"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "grabar"
            Object.ToolTipText     =   "Acción Grabar"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Acción Cancelar"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cuentas"
            Object.ToolTipText     =   "Cuentas de EMail."
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar staMensaje 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   5490
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Plantilla:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "A &Quien:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Condición:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3660
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cuerpo:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Asunto:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2100
      Width           =   675
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cuentas:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Prioridad:"
      Height          =   255
      Left            =   5220
      TabIndex        =   2
      Top             =   540
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&De:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1140
      Width           =   675
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
   Begin VB.Menu MnuPregunta 
      Caption         =   "&?"
      Begin VB.Menu MnuPreAyuda 
         Caption         =   "&Ayuda ..."
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmReglas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sNuevo As Boolean
Private clsGeneral As New clsorCGSA

Private Sub bAQuien_Click()
On Error GoTo errBC
    Dim objLista As New clsListadeAyuda
    If Trim(tAQuien.Text) <> "" Then objLista.ActivoListaAyudaSQL cBase, tAQuien.Text
    Set objLista = Nothing
    Exit Sub
errBC:
    clsGeneral.OcurrioError "Ocurrió el siguiente error al acceder a la lista de ayuda.", Err.Description
End Sub

Private Sub bAsunto_Click()
On Error GoTo errBC
    
    If Trim(tAsunto.Text) = "" Then Exit Sub
    Cons = InputBox("Ingrese un texto que simule el Asunto.", "Prueba de asunto")
    If RetornoCondicionEnArray(tAsunto.Text, Cons) Then
        MsgBox "El asunto cumple con la condición.", vbInformation, "SE CUMPLE"
    Else
        MsgBox "El asunto NO cumple con la condición." & vbCrLf & "Verifique si tiene espacios despues de algún punto y coma.", vbInformation, "NO SE CUMPLE"
    End If
    Exit Sub
errBC:
    clsGeneral.OcurrioError "Ocurrió el siguiente al validar asunto.", Err.Description
End Sub

Private Sub bCuenta_Click()
On Error GoTo errBC
    Cons = "Select CMaCodigo as 'Código', CMaNombre as 'Nombre',  EMSDireccion as 'Servidor' " _
        & " From CuentaMail, EMailServer Where CMaActiva = 1 And CMaServidor = EMSCodigo"
    Dim objLista As New clsListadeAyuda
    If objLista.ActivarAyuda(cBase, Cons, 5000, 0, "Cuentas") Then
        InsertoCuenta objLista.RetornoDatoSeleccionado(0)
    End If
    Set objLista = Nothing
    Exit Sub
errBC:
    clsGeneral.OcurrioError "Ocurrió el siguiente error al acceder a la lista de ayuda.", Err.Description
End Sub

Private Sub bCuerpo_Click()
On Error GoTo errBC
    If Trim(tCuerpo.Text) = "" Then Exit Sub
    Cons = InputBox("Ingrese un texto que simule el cuerpo del mail.", "Prueba de cuerpo")
    If RetornoCondicionEnArray(tCuerpo.Text, Cons) Then
        MsgBox "El texto ingresado cumple con la condición.", vbInformation, "SE CUMPLE"
    Else
        MsgBox "El texto ingresado NO cumple con la condición." & vbCrLf & "Verifique si tiene espacios despues de algún punto y coma.", vbInformation, "NO SE CUMPLE"
    End If
    Exit Sub
errBC:
    clsGeneral.OcurrioError "Ocurrió el siguiente al validar asunto.", Err.Description
End Sub

Private Sub bDePrevia_Click()
On Error GoTo errBC
    Dim objLista As New clsListadeAyuda
    If Trim(tDe.Text) <> "" Then
        If InStr(1, LCase(tDe.Text), "[prmde]") Then
            Cons = InputBox("Ingrese un ejemplo para el parámetro." & vbCrLf & "Ej.: micuenta@servidor.com)", "Prueba")
            Cons = Replace(LCase(tDe.Text), "[prmde]", Cons)
        Else
            Cons = tDe.Text
        End If
        objLista.ActivoListaAyudaSQL cBase, Cons
    End If
    Set objLista = Nothing
    Exit Sub
errBC:
    clsGeneral.OcurrioError "Ocurrió el siguiente error al acceder a la lista de ayuda.", Err.Description
End Sub

Private Sub cAQuien_Change()
    tAQuien.Text = ""
End Sub

Private Sub cAQuien_GotFocus()
On Error Resume Next
    With cAQuien
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione la consulta a quien va el mensaje."
End Sub

Private Sub cAQuien_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If cAQuien.ListIndex > -1 Then
            CargoPlantillaAQuien cAQuien.ItemData(cAQuien.ListIndex)
        Else
            tAQuien.Text = ""
        End If
        cPlantilla.SetFocus
    End If
End Sub

Private Sub cAQuien_LostFocus()
    Ayuda ""
End Sub

Private Sub cPlantilla_GotFocus()
On Error Resume Next
    With cPlantilla
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione una plantilla si desea enviarle una respuesta a quien envía el mail."
End Sub

Private Sub cPlantilla_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then chBorrar.SetFocus
End Sub

Private Sub cPlantilla_LostFocus()
    Ayuda ""
End Sub

Private Sub chBorrar_GotFocus()
    Ayuda "Indique si el mail debe ser eliminado del servidor."
End Sub

Private Sub chBorrar_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub chBorrar_LostFocus()
    Ayuda ""
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad

    Me.Caption = "Reglas de EMail"
    ObtengoSeteoForm Me, 500, 500
    Me.Height = 6405
    miBotones True, False
    sNuevo = False
    EstadoObjetos False
    LimpioObjetos
    With tooMenu
        .ImageList = imgIconos
        .Buttons("salir").Image = imgIconos.ListImages("salir").Index
        .Buttons("nuevo").Image = imgIconos.ListImages("nuevo").Index
        .Buttons("modificar").Image = imgIconos.ListImages("modificar").Index
        .Buttons("eliminar").Image = imgIconos.ListImages("eliminar").Index
        .Buttons("grabar").Image = imgIconos.ListImages("grabar").Index
        .Buttons("cancelar").Image = imgIconos.ListImages("cancelar").Index
        .Buttons("cuentas").Image = imgIconos.ListImages("cuentas").Index
    End With
    Cons = "Select PlaCodigo, PlaNombre From Plantilla Where PlaTipo = 4 Order by PlaNombre"
    CargoCombo Cons, cPlantilla
    Cons = "Select PlaCodigo, PlaNombre From Plantilla Where PlaTipo = 5 Order by PlaNombre"
    CargoCombo Cons, cAQuien
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al iniciar el formulario.", Trim(Err.Description)
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Sub

Private Sub Label1_Click()
On Error Resume Next
    With tNombre
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label2_Click()
On Error Resume Next
    With tDe
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label3_Click()
On Error Resume Next
    With tPrioridad
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label4_Click()
On Error Resume Next
    With tCuenta
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label5_Click()
On Error Resume Next
    With tAsunto
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label6_Click()
On Error Resume Next
    With tCuerpo
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label7_Click()
On Error Resume Next
    With tCondicion
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label8_Click()
On Error Resume Next
    With cAQuien
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label9_Click()
On Error Resume Next
    With cPlantilla
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
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

Private Sub MnuPreAyuda_Click()
On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    Cons = "Select * from Aplicacion Where AplNombre = '" & Trim(App.Title) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux!AplHelp) Then aFile = Trim(RsAux!AplHelp)
    RsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub
Private Sub AccionNuevo()
    'Prendo Señal que es uno nuevo.
    sNuevo = True
    'Habilito y Desabilito Botones
    miBotones False, False
    tNombre.Text = "": tNombre.Tag = ""
    LimpioObjetos
    EstadoObjetos True
    tNombre.SetFocus
End Sub

Private Sub AccionModificar()
On Error GoTo errAM
    'Habilito y Desabilito Botones
    miBotones False, False
    EstadoObjetos True
    tNombre.SetFocus
    Exit Sub
errAM:
    clsGeneral.OcurrioError "Ocurrió el siguiente error al cargar la ficha.", Err.Description
End Sub

Private Sub AccionGrabar()
    If ValidoDatos Then
        If sNuevo Then
            If MsgBox("¿Confirma el alta de datos?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then NuevoRegistro
        Else
            If MsgBox("¿Confirma modificar los datos?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then ModificoRegistro
        End If
    End If
End Sub

Private Sub AccionEliminar()
On Error GoTo ErrAE
    If MsgBox("¿Confirma eliminar el dato seleccionado?", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        Screen.MousePointer = 11
        'Existen foraneas creadas.
        Cons = "Delete ReglaCuentaMail Where RCMCodigo = " & Val(tNombre.Tag)
        cBase.Execute (Cons)
        tNombre.Text = "": tNombre.Tag = ""
        LimpioObjetos
        miBotones True, False
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrAE:
    clsGeneral.OcurrioError "Ocurrio un error al eliminar el registro.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub AccionCancelar()
    sNuevo = False
    miBotones True, False
    EstadoObjetos False
    If Val(tNombre.Tag) = 0 Then
        LimpioObjetos
    Else
        LimpioObjetos
        CargoRegla Val(tNombre.Tag)
    End If
End Sub

Private Sub EstadoObjetos(ByVal bHabilitar As Boolean)
On Error GoTo errEO

    If bHabilitar Then
        tPrioridad.BackColor = vbWindowBackground
        tCuenta.BackColor = &HC0FFFF
        tDe.BackColor = vbWindowBackground
        tAsunto.BackColor = vbWindowBackground
        tCuerpo.BackColor = vbWindowBackground
        tCondicion.BackColor = vbWindowBackground
        cPlantilla.BackColor = vbWindowBackground
        cAQuien.BackColor = vbWindowBackground
    Else
        tPrioridad.BackColor = vbButtonFace
        tCuenta.BackColor = vbButtonFace
        tDe.BackColor = vbButtonFace
        tAsunto.BackColor = vbButtonFace
        tCuerpo.BackColor = vbButtonFace
        tCondicion.BackColor = vbButtonFace
        cPlantilla.BackColor = vbButtonFace
        cAQuien.BackColor = vbButtonFace
    End If
    
    tAQuien.BackColor = vbButtonFace
    tPrioridad.Enabled = bHabilitar
    tCuenta.Enabled = bHabilitar
    tDe.Locked = Not bHabilitar
    tAsunto.Locked = Not bHabilitar
    tCuerpo.Locked = Not bHabilitar
    tCondicion.Enabled = bHabilitar
'    tAQuien.Locked = Not bHabilitar
    cPlantilla.Enabled = bHabilitar
    chBorrar.Enabled = bHabilitar
    cAQuien.Enabled = bHabilitar
    
    'botones
    bCuenta.Enabled = bHabilitar
    bDePrevia.Enabled = bHabilitar
    bAQuien.Enabled = bHabilitar
    bAsunto.Enabled = bHabilitar
    bCuerpo.Enabled = bHabilitar
    Exit Sub
    
errEO:
End Sub

Private Sub tAsunto_GotFocus()
On Error Resume Next
    With tAsunto
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese las condiciones que debe cumplir el asunto para cumplir la condición. (separadas por ;)"
End Sub

Private Sub tAsunto_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        tAsunto.SelStart = Len(tAsunto.Text)
        tCuerpo.SetFocus
    End If
End Sub

Private Sub tAsunto_LostFocus()
    Ayuda ""
End Sub

Private Sub tCondicion_GotFocus()
On Error Resume Next
    With tCondicion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese que condiciones desea validar. ([condDe], [condAsunto], [condCuerpo], [condattach])"
End Sub

Private Sub tCondicion_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then cAQuien.SetFocus
End Sub

Private Sub tCondicion_LostFocus()
    Ayuda ""
End Sub

Private Sub tCuenta_GotFocus()
On Error Resume Next
    With tCuenta
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese los códigos de las cuentas que incluyen esta regla. (separadas por ;) "
End Sub

Private Sub tCuenta_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tDe.SetFocus
End Sub

Private Sub tCuenta_LostFocus()
    Ayuda ""
End Sub

Private Sub tCuerpo_GotFocus()
On Error Resume Next
    With tCuerpo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese las condiciones que debe cumplir el cuerpo del mail para cumplir la condición. (separadas por ;)"
End Sub

Private Sub tCuerpo_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        tCuerpo.SelStart = Len(tCuerpo.Text)
        tCondicion.SetFocus
    End If
End Sub

Private Sub tCuerpo_LostFocus()
    Ayuda ""
End Sub

Private Sub tDe_GotFocus()
On Error Resume Next
    With tDe
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese una consulta SQL que indique una condición para la dirección de quien envía el mail. ('[prmDe]')"
End Sub

Private Sub tDe_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        tDe.SelStart = Len(tDe.Text)
        tAsunto.SetFocus
    End If
End Sub

Private Sub tDe_LostFocus()
    Ayuda ""
End Sub

Private Sub tNombre_Change()
    If Not tPrioridad.Enabled Then
        LimpioObjetos
        tNombre.Tag = "0"
    End If
End Sub

Private Sub tNombre_GotFocus()
On Error Resume Next
    With tNombre
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese el nombre de la regla. [Enter] búsqueda"
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
On Error GoTo errN
    If KeyAscii = vbKeyReturn Then
        If tPrioridad.Enabled Then
            tPrioridad.SetFocus
        Else
            AyudaRegla
        End If
    End If
    Exit Sub
errN:
    clsGeneral.OcurrioError "Ocurrió un error.", Err.Description
End Sub

Private Sub tNombre_LostFocus()
    Ayuda ""
End Sub

Private Sub tooMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
        Case "cuentas": AccionCuentas
    End Select
    
End Sub

Private Sub NuevoRegistro()
On Error GoTo ErrNI
Dim aID As Long
    Screen.MousePointer = 11
    Cons = "Select * From ReglaCuentaMail Where RCMNombre = '" & Trim(tNombre.Text) & "'"
    If Trim(tPrioridad.Text) <> "" Then Cons = Cons & " Or RCMPrioridad = " & Val(tPrioridad.Text)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        MsgBox "La siguiente regla contiene ese nombre o la misma propiedad." & vbCrLf _
            & "Nombre: " & Trim(RsAux!RCMNombre), vbExclamation, "ATENCIÓN"
        RsAux.Close
        Screen.MousePointer = 0
        Exit Sub
    Else
        RsAux.AddNew
        CamposABD
        RsAux.Update
        RsAux.Close
        
        'Cargo el id de la regla.
        Cons = "Select Max(RCMCodigo) From ReglaCuentaMail Where RCMNombre = '" & Trim(tNombre.Text) & "'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        tNombre.Tag = RsAux(0)
        RsAux.Close
        AccionCancelar
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrNI:
    clsGeneral.OcurrioError "Ocurrio un error al dar el alta.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub CamposABD()
    
    RsAux!RCMNombre = Trim(tNombre.Text)
    If IsNumeric(tPrioridad.Text) Then RsAux!RCMPrioridad = Val(tPrioridad.Text) Else RsAux!RCMPrioridad = Null
    If Trim(tCuenta.Text) = "" Then RsAux!RCMCuentas = Null Else RsAux!RCMCuentas = Trim(tCuenta.Text)
    If Trim(tDe.Text) = "" Then RsAux!RCMDe = Null Else RsAux!RCMDe = Trim(tDe.Text)
    If Trim(tAsunto.Text) = "" Then RsAux!RCMAsunto = Null Else RsAux!RCMAsunto = Trim(tAsunto.Text)
    If Trim(tCuerpo.Text) = "" Then RsAux!RCMCuerpo = Null Else RsAux!RCMCuerpo = Trim(tCuerpo.Text)
    If Trim(tCondicion.Text) = "" Then RsAux!RCMResultado = Null Else RsAux!RCMResultado = Trim(tCondicion.Text)
    If Trim(tAQuien.Text) = "" Then RsAux!RCMAQuien = Null Else RsAux!RCMAQuien = Trim(tAQuien.Text)
    If cAQuien.ListIndex > -1 Then RsAux!RCMIdAQuien = cAQuien.ItemData(cAQuien.ListIndex) Else RsAux!RCMIdAQuien = Null
    If cPlantilla.ListIndex > -1 Then RsAux!RCMPlantilla = cPlantilla.ItemData(cPlantilla.ListIndex) Else RsAux!RCMPlantilla = Null
    RsAux!RCMBorrar = chBorrar.Value
    
End Sub
Private Sub ModificoRegistro()
On Error GoTo ErrMR
Dim aIDCod As Long
    Screen.MousePointer = 11
    Cons = "Select * From ReglaCuentaMail Where RCMCodigo <> " & Val(tNombre.Tag)
        
    If Trim(tPrioridad.Text) <> "" Then
        Cons = Cons & "And (RCMNombre = '" & Trim(tNombre.Text) & "' Or RCMPrioridad = " & Val(tPrioridad.Text) & ")"
    Else
        Cons = Cons & "And RCMNombre = '" & Trim(tNombre.Text) & "'"
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        MsgBox "La siguiente regla contiene ese nombre o la misma propiedad." & vbCrLf _
            & "Nombre: " & Trim(RsAux!RCMNombre), vbExclamation, "ATENCIÓN"
        RsAux.Close
    Else
        RsAux.Close
        Cons = "Select * From ReglaCuentaMail Where RCMCodigo = " & Val(tNombre.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        CamposABD
        RsAux.Update
        RsAux.Close
        AccionCancelar
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrMR:
    clsGeneral.OcurrioError "Ocurrio un error al modificar el registro.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub Ayuda(strTexto As String)
    staMensaje.SimpleText = strTexto
End Sub
Private Function ValidoDatos() As Boolean
    
    ValidoDatos = False
    If Trim(tNombre.Text) = "" Then
        MsgBox "Se debe ingresar un nombre.", vbExclamation, "ATENCIÓN"
        tNombre.SetFocus: Exit Function
    End If
    
    If Trim(tPrioridad.Text) <> "" Then
        If Not IsNumeric(tPrioridad.Text) Then
            MsgBox "La prioridad debe ser un número.", vbExclamation, "ATENCIÓN"
            tPrioridad.SetFocus: Exit Function
        Else
            If Val(tPrioridad.Text) < 0 Then
                MsgBox "La prioridad debe ser mayor o igual a cero.", vbExclamation, "ATENCIÓN"
                tPrioridad.SetFocus: Exit Function
            End If
        End If
    End If
    If cAQuien.ListIndex > -1 Then CargoPlantillaAQuien cAQuien.ItemData(cAQuien.ListIndex)
    ValidoDatos = True
End Function

'-----------------------------------------------------------------------------------
'   Habilita y deshabilita los botones y menus del toolbar
'-----------------------------------------------------------------------------------
Private Sub miBotones(ByVal bNuevo As Boolean, ByVal bModEl As Boolean)

    'Habilito y Desabilito Botones.
    tooMenu.Buttons("nuevo").Enabled = bNuevo
    MnuNuevo.Enabled = bNuevo
    
    tooMenu.Buttons("modificar").Enabled = bModEl
    MnuModificar.Enabled = bModEl
    
    tooMenu.Buttons("eliminar").Enabled = bModEl
    MnuEliminar.Enabled = bModEl
    
    tooMenu.Buttons("grabar").Enabled = Not bNuevo
    MnuGrabar.Enabled = Not bNuevo
    
    tooMenu.Buttons("cancelar").Enabled = Not bNuevo
    MnuCancelar.Enabled = Not bNuevo
    

End Sub

Private Sub tPrioridad_GotFocus()
On Error Resume Next
    With tPrioridad
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese el orden de prioridad de la regla (orden ascendente) (nada = inactiva)"
End Sub

Private Sub tPrioridad_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tCuenta.SetFocus
End Sub

Private Sub tPrioridad_LostFocus()
    Ayuda ""
End Sub

Private Sub LimpioObjetos()
    tPrioridad.Text = ""
    tCuenta.Text = ""
    tDe.Text = ""
    tAsunto.Text = ""
    tCuerpo.Text = ""
    tCondicion.Text = ""
    cAQuien.Text = ""
    tAQuien.Text = ""
    cPlantilla.Text = ""
    chBorrar.Value = 0
End Sub

Private Sub AyudaRegla()
On Error GoTo errAR
    
    Cons = "Select * From ReglaCuentaMail Where RCMNombre like '" & Replace(tNombre.Text, " ", "%") & "%'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.MoveNext
        If Not RsAux.EOF Then
            'Lista de ayuda
            RsAux.Close
            Cons = "Select RCMCodigo, RCMNombre as 'Nombre'  From ReglaCuentaMail Where RCMNombre like '" & Replace(tNombre.Text, " ", "%") & "%'"
            Dim objLista As New clsListadeAyuda
            If objLista.ActivarAyuda(cBase, Cons, 5000, 1, "Lista de Reglas") Then
                tNombre.Tag = objLista.RetornoDatoSeleccionado(0)
            End If
            Set objLista = Nothing
        Else
            RsAux.MoveFirst
            tNombre.Tag = RsAux!RCMCodigo
            RsAux.Close
        End If
    Else
        RsAux.Close
        MsgBox "No existe una regla que coincida con esa descripción.", vbInformation, "ATENCIÓN"
    End If
    If Val(tNombre.Tag) > 0 Then
        CargoRegla tNombre.Tag
    Else
        LimpioObjetos
    End If
    Exit Sub
errAR:
    LimpioObjetos
    clsGeneral.OcurrioError "Ocurrió un error al buscar las reglas por nombre.", Err.Description
End Sub

Private Sub CargoRegla(ByVal idRegla As Long)
On Error GoTo errCR
    Screen.MousePointer = 11
    tNombre.Text = "": tNombre.Tag = "": LimpioObjetos
    miBotones True, False
    Cons = "Select * From ReglaCuentaMail Where RCMCodigo = " & idRegla
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        tNombre.Text = Trim(RsAux!RCMNombre)
        tNombre.Tag = RsAux!RCMCodigo
        If Not IsNull(RsAux!RCMPrioridad) Then tPrioridad.Text = RsAux!RCMPrioridad
        If Not IsNull(RsAux!RCMCuentas) Then tCuenta.Text = Trim(RsAux!RCMCuentas)
        tDe.Text = RetornoCampoLongText("RCMDe")
        tAsunto.Text = RetornoCampoLongText("RCMAsunto")
        tCuerpo.Text = RetornoCampoLongText("RCMCuerpo")
        'tAQuien.Text = RetornoCampoLongText("RCMAQuien")
        If Not IsNull(RsAux!RCMIdAQuien) Then
            BuscoCodigoEnCombo cAQuien, RsAux!RCMIdAQuien
            CargoPlantillaAQuien RsAux!RCMIdAQuien
        End If
        
        tCondicion.Text = RetornoCampoLongText("RCMResultado")
        If Not IsNull(RsAux!RCMPlantilla) Then BuscoCodigoEnCombo cPlantilla, RsAux!RCMPlantilla
        If RsAux!RCMBorrar Then chBorrar.Value = 1 Else chBorrar.Value = 0
        miBotones True, True
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errCR:
    clsGeneral.OcurrioError "Ocurrió un error al intentar levantar la información de la regla.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function RetornoCampoLongText(ByVal NombreCampo As String) As String
On Error GoTo errRCLT
    
    RetornoCampoLongText = ""            'Nulo
    RetornoCampoLongText = RsAux(NombreCampo)

errRCLT:
    Resume errFin
errFin:

End Function

Private Sub InsertoCuenta(ByVal idCuenta As String)
On Error Resume Next
Dim arrCuenta As Variant
Dim iPos As Integer
Dim sAux As String
    arrCuenta = Split(tCuenta.Text, ";")
    For iPos = 0 To UBound(arrCuenta)
        If sAux = "" Then
            sAux = arrCuenta(iPos)
        Else
            sAux = sAux & ";" & arrCuenta(iPos)
        End If
        If arrCuenta(iPos) = idCuenta Then Exit Sub
    Next iPos
    
    If sAux = "" Then
        tCuenta.Text = idCuenta
    Else
        tCuenta.Text = sAux & ";" & idCuenta
    End If
    
End Sub
Private Function RetornoCondicionEnArray(ByVal sCondicion As String, ByVal sAValidar As String) As Boolean
On Error GoTo errRCA
Dim arrAux As Variant
Dim iAux As Integer
    RetornoCondicionEnArray = False
    arrAux = Split(sCondicion, ";")
    For iAux = 0 To UBound(arrAux)
        If Chr(34) & LCase(sAValidar) & Chr(34) Like Chr(34) & LCase(arrAux(iAux)) & Chr(34) Then
            RetornoCondicionEnArray = True
            Exit For
        End If
    Next iAux
    Exit Function
errRCA:
    MsgBox "Ocurrió el siguiente error: " & Trim(Err.Description), vbExclamation, "ATENCIÓN"
End Function

Private Sub CargoPlantillaAQuien(ByVal idPla As Long)
On Error GoTo errCP
Dim rsA As rdoResultset
    
    Cons = "Select * From Plantilla Where PlaCodigo = " & idPla
    Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsA.EOF Then
        On Error Resume Next
        tAQuien.Text = rsA("PlaTexto")
    End If
    rsA.Close
    Exit Sub
errCP:
    clsGeneral.OcurrioError "Error al cargar plantilla a quien.", Err.Description
End Sub

Private Sub AccionCuentas()
    EjecutarApp App.Path & "\Cuenta_Mail.exe"
End Sub
