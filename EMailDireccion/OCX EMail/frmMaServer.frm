VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMaServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servidores de Correo"
   ClientHeight    =   3600
   ClientLeft      =   4155
   ClientTop       =   3135
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5400
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
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
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   300
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
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dirección del Servidor"
      ForeColor       =   &H00000080&
      Height          =   915
      Left            =   60
      TabIndex        =   10
      Top             =   1440
      Width           =   5235
      Begin VB.TextBox tHost 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   840
         MaxLength       =   40
         TabIndex        =   12
         Top             =   540
         Width           =   3855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Host:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "(Ej.: bdc.com.uy)"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Ejemplo"
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   60
      TabIndex        =   4
      Top             =   2580
      Width           =   5235
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   840
         MaxLength       =   40
         TabIndex        =   6
         Text            =   " Banco de Crédito"
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   840
         MaxLength       =   40
         TabIndex        =   5
         Text            =   " bdc.com.uy"
         Top             =   540
         Width           =   3855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Host:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   540
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nombre detallado de la Empresa del Servidor"
      ForeColor       =   &H00000080&
      Height          =   915
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   5235
      Begin VB.TextBox tNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   840
         MaxLength       =   40
         TabIndex        =   1
         Top             =   540
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "(Ej.: Banco de Crédito)"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1755
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
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   0
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
            Picture         =   "frmMaServer.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaServer.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaServer.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaServer.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaServer.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaServer.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaServer.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaServer.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaServer.frx":10E2
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
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMaServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean
Dim m_IDServer As Long

Public prmTxtHost As String
Public prmTxtNombre As String
Public prmIdServer As Long

Private cons As String
Private rsAux As rdoResultset
Private rdoCBLocal As rdoConnection

Public Property Set Connect(ByVal objConnect As rdoConnection)
    Set rdoCBLocal = objConnect
End Property

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error Resume Next
    sNuevo = False: sModificar = False
    LimpioFicha
        
    zfn_HabilitoIngreso bOK:=False
    AccionNuevo
    
    If prmIdServer <> 0 Then
        BuscoServidor prmIdServer
        If m_IDServer = 0 Then
            MsgBox "El servidor seleccionado no se encuentra en la base de datos.", vbExclamation, "Posible Error"
        Else
            sNuevo = False
            AccionModificar
        End If
    
    Else
        If prmTxtHost <> "" Then tHost.Text = Trim(LCase(prmTxtHost))
    End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set rdoCBLocal = Nothing
End Sub

Private Sub Label1_Click()
    zfn_Foco tNombre
End Sub

Private Sub Label2_Click()
    zfn_Foco tHost
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
    m_IDServer = 0
    zfn_Botones False, False, False, True, True, Toolbar1, Me
    
    LimpioFicha
    zfn_HabilitoIngreso
    zfn_Foco tNombre
  
End Sub

Private Sub AccionModificar()
    
    On Error Resume Next
    If m_IDServer = 0 Then Exit Sub
    
    sModificar = True
    
    zfn_HabilitoIngreso
    zfn_Botones False, False, False, True, True, Toolbar1, Me
    
    tNombre.SetFocus
        
End Sub

Private Sub BuscoServidor(aId As Long)
On Error GoTo errBuscar

    m_IDServer = 0
    cons = "Select * from EMailServer Where EMSCodigo = " & aId
    Set rsAux = rdoCBLocal.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsAux.EOF Then
        m_IDServer = aId
        tNombre.Text = Trim(rsAux!EMSNombre)
        tHost.Text = Trim(rsAux!EMSDireccion)
    End If
    
    rsAux.Close
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Error al buscar el servidor.", Err.Description
End Sub

Private Sub AccionGrabar()

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "Grabar Datos") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo errGrabar
    
    'Texto = Nombre, Texto2 = Host
    cons = "Select * from EMailServer Where EMSCodigo = " & m_IDServer
    Set rsAux = rdoCBLocal.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        
    If sNuevo Then rsAux.AddNew Else rsAux.Edit
    rsAux!EMSNombre = Trim(tNombre.Text)
    rsAux!EMSDireccion = Trim(LCase(tHost.Text))
    rsAux.Update
    rsAux.Close
    
    If prmTxtHost <> "" Then
        On Error Resume Next            'Cargo variables publicas del datos ingresado   -------------------
        prmTxtNombre = Trim(LCase(tHost.Text))
        
        cons = "Select * from EMailServer " & _
                    " Where EMSNombre = '" & Trim(tNombre.Text) & "'" & _
                    " And EMSDireccion = '" & Trim(LCase(tHost.Text)) & "'"
        Set rsAux = rdoCBLocal.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then prmIdServer = rsAux!EMSCodigo
        rsAux.Close
        
        Unload Me
        Exit Sub                                '-------------------------------------------------------------------------
    End If
    
    If Not sModificar Then m_IDServer = 0
    AccionCancelar
    
    Screen.MousePointer = 0
    Exit Sub
errGrabar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
End Sub

Private Sub AccionEliminar()

    Screen.MousePointer = 11
    On Error GoTo errDelete
    Dim bHay As Boolean
    
    'Valido si hay direcciones con el servidor
    bHay = False
    cons = "Select * from EMailDireccion Where EMDServidor = " & m_IDServer
    Set rsAux = rdoCBLocal.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bHay = True
    rsAux.Close
    
    If bHay Then
        MsgBox "Hay direcciones de correo que pertenecen al servidor seleccionado." & vbCrLf & "No podrá eliminarlo.", vbExclamation, "Hay Direcciones Ingresadas"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    If MsgBox("¿Confirma eliminar el servidor de correo '" & tHost.Text & "'?", vbQuestion + vbYesNo, "Eliminar Servidor") = vbNo Then Screen.MousePointer = 0: Exit Sub
    
    cons = "Select * from EMailServer Where EMSCodigo = " & m_IDServer
    Set rsAux = rdoCBLocal.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsAux.Delete
    rsAux.Close
    
    LimpioFicha
    zfn_HabilitoIngreso bOK:=False
    
    m_IDServer = 0
    AccionCancelar
    
    Screen.MousePointer = 0
    Exit Sub
    
errDelete:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al eliminar el servidor de correo.", Err.Description
End Sub

Private Sub AccionCancelar()

    On Error Resume Next
    zfn_HabilitoIngreso bOK:=False
    
    If m_IDServer <> 0 Then
        zfn_Botones True, True, True, False, False, Toolbar1, Me
    Else
        LimpioFicha
        zfn_Botones True, False, False, False, False, Toolbar1, Me
    End If
    
    sNuevo = False: sModificar = False
    
End Sub

Private Sub tHost_GotFocus()
    If Trim(tHost.Text) <> "" Then tHost.SelStart = Len(tHost.Text)
End Sub

Private Sub tHost_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zfn_Foco tHost
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

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    If Trim(tNombre.Text) = "" Then
        MsgBox "Ingrese la identificación para el servidor de correo.", vbExclamation, "Falta Nombre del Servidor"
        zfn_Foco tNombre: Exit Function
    End If
    
    If Trim(tHost.Text) = "" Then
        MsgBox "Ingrese la ubicación del host.", vbExclamation, "Falta dirección del Servidor"
        zfn_Foco tHost: Exit Function
    End If
    
    If Not IsValidIPHost(tHost.Text) Then
        MsgBox "El nombre del host '" & Trim(tHost.Text) & "' no es correcto.", vbExclamation, "Dirección del Servidor Incorrecta."
        zfn_Foco tHost: Exit Function
    End If
    
    'Valido los nombres ---------------------------------------------------------------------------
    Dim bEOF As Boolean
    cons = "Select * from EMailServer " & _
               " Where ( EMSNombre = '" & tNombre.Text & "' OR EMSDireccion = '" & tHost.Text & "' )" & _
               " And EMSCodigo <> " & m_IDServer
    Set rsAux = rdoCBLocal.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    bEOF = rsAux.EOF
    rsAux.Close
    
    If Not bEOF Then
        MsgBox "Ya existen servidores de correco con el nombre o dirección ingresada." & vbCrLf & _
                    "Verifique en la lista de datos.", vbExclamation, "Coincidencia de datos"
        Exit Function
    End If
    '-----------------------------------------------------------------------------------------------------------
    
    ValidoCampos = True
    
End Function

Private Sub zfn_HabilitoIngreso(Optional bOK As Boolean = True)

    If bOK Then
        tNombre.BackColor = vbWindowBackground
        tHost.BackColor = vbWindowBackground
    Else
        tNombre.BackColor = vbButtonFace
        tHost.BackColor = vbButtonFace
    End If
    
    tNombre.Enabled = bOK
    tHost.Enabled = bOK
    
End Sub

Private Sub LimpioFicha()
    tNombre.Text = ""
    tHost.Text = ""
End Sub

Private Function zfn_Foco(toControl As Control)
    On Error Resume Next

    If Not toControl.Enabled Then Exit Function
    With toControl
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

End Function

Public Sub zfn_Botones(Nu As Boolean, Mo As Boolean, El As Boolean, Gr As Boolean, Ca As Boolean, Toolbar1 As Control, nForm As Form)

    Toolbar1.Buttons("nuevo").Enabled = Nu
    nForm.MnuNuevo.Enabled = Nu
    
    Toolbar1.Buttons("modificar").Enabled = Mo
    nForm.MnuModificar.Enabled = Mo
    
    Toolbar1.Buttons("eliminar").Enabled = El
    nForm.MnuEliminar.Enabled = El
    
    Toolbar1.Buttons("grabar").Enabled = Gr
    nForm.MnuGrabar.Enabled = Gr
    
    Toolbar1.Buttons("cancelar").Enabled = Ca
    nForm.MnuCancelar.Enabled = Ca

End Sub

