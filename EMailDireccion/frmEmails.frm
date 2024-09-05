VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEMails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direcciones E-Mails"
   ClientHeight    =   5940
   ClientLeft      =   3300
   ClientTop       =   2340
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
   Icon            =   "frmEmails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6420
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6420
      _ExtentX        =   11324
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
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      ForeColor       =   &H00000080&
      Height          =   1875
      Left            =   60
      TabIndex        =   17
      Top             =   480
      Width           =   6315
      Begin VB.CommandButton bSModificar 
         Height          =   315
         Left            =   5340
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Modificar Servidor."
         Top             =   1500
         Width           =   315
      End
      Begin VB.CheckBox bNoEnviar 
         Caption         =   "No Enviar &Info."
         Height          =   195
         Left            =   4260
         TabIndex        =   9
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox tServidor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   900
         TabIndex        =   13
         Top             =   1515
         Width           =   3975
      End
      Begin VB.CommandButton bCliente 
         Caption         =   "..."
         Height          =   280
         Left            =   2220
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox tCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   4755
      End
      Begin VB.CommandButton bSAgregar 
         Height          =   315
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Agregar Servidor."
         Top             =   1500
         Width           =   315
      End
      Begin VB.TextBox tDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   900
         MaxLength       =   40
         TabIndex        =   11
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox tNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   900
         MaxLength       =   40
         TabIndex        =   8
         Top             =   900
         Width           =   3255
      End
      Begin MSMask.MaskEdBox tCi 
         Height          =   285
         Left            =   900
         TabIndex        =   4
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label lAlta 
         Alignment       =   1  'Right Justify
         Caption         =   "lIngreso"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2820
         TabIndex        =   18
         Top             =   270
         Width           =   2835
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Servidor:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Dirección:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   945
         Width           =   735
      End
   End
   Begin VB.TextBox tBuscar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   960
      MaxLength       =   40
      TabIndex        =   1
      Top             =   2400
      Width           =   4755
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   3195
      Left            =   60
      TabIndex        =   2
      Top             =   2700
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
      AllowUserResizing=   0
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Buscar:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2415
      Width           =   735
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4740
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
            Picture         =   "frmEmails.frx":0442
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmails.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmails.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmails.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmails.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmails.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmails.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmails.frx":0DC8
            Key             =   "modificar"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEmails.frx":10E2
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
   Begin VB.Menu MnuEMail 
      Caption         =   "MnuEMail"
      Visible         =   0   'False
      Begin VB.Menu MnuEMTitulo 
         Caption         =   "MnuEMTitulo"
      End
      Begin VB.Menu MnuEML1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAvisoLlegada 
         Caption         =   "Agregar Aviso Llegada"
      End
      Begin VB.Menu MnuEML2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEMX 
         Caption         =   "MnuEMX"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmEMails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean
Dim gIdEMail As Long

Public prmIDCliente As Long
Public prmIDEMail As Long

Private Sub bCliente_Click()
    On Error GoTo errBuscar
    If Not sNuevo And Not sModificar Then Exit Sub
    
    Dim idSel As Long
    
    Dim objBuscar As New clsBuscarCliente
    objBuscar.ActivoFormularioBuscarClientes cBase, True
    idSel = objBuscar.BCClienteSeleccionado
    Set objBuscar = Nothing
    
    If idSel <> 0 Then
        Cons = "Select Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2), RTrim(CPeApellido1) + ', ' + RTrim(CPeNombre1) as Mail" _
               & " From CPersona " _
               & " Where CPeCliente =" & idSel _
                                                    & " UNION ALL" _
               & " Select Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), RTrim(CEmFantasia) as Mail" _
               & " From CEmpresa " _
               & " Where CEmCliente = " & idSel
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            tCliente.Text = Trim(RsAux!Nombre)
            tCliente.Tag = idSel
            If Not IsNull(RsAux!Mail) Then tNombre.Text = StrConv(RsAux!Mail, vbProperCase)
        End If
        RsAux.Close
    End If
    
    Screen.MousePointer = 0
    Exit Sub
errBuscar:
    clsGeneral.OcurrioError "Error al buscar el cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bNoEnviar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDireccion
End Sub

Private Sub bSAgregar_Click()
    On Error GoTo errCargar
    frmMaServer.prmIdServer = 0
    frmMaServer.prmTxtNombre = ""
    frmMaServer.prmTxtHost = ""
    frmMaServer.Show vbModal, Me
    tServidor.SetFocus
    
errCargar:
    Screen.MousePointer = 0
End Sub


Private Sub bSModificar_Click()
On Error GoTo errCargar
    
    If Val(tServidor.Tag) = 0 Then
        MsgBox "Debe seleccionar un servidor para modificar sus datos.", vbExclamation, "Posible Error "
        Exit Sub
    End If
    frmMaServer.prmTxtHost = ""
    frmMaServer.prmIdServer = Val(tServidor.Tag)
    
    frmMaServer.Show vbModal, Me
    
    On Error GoTo errBuscar
    Cons = "Select * from EMailServer " & _
                " Where EMSCodigo = " & Val(tServidor.Tag)
    tServidor.Tag = 0
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        tServidor.Text = Trim(RsAux!EMSDireccion)
        tServidor.Tag = RsAux!EMSCodigo
    Else
        tServidor.Text = ""
    End If
    RsAux.Close
    
    If tServidor.Enabled Then tServidor.SetFocus
    
errCargar:
    Screen.MousePointer = 0
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Error al validar el servidor.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub


Private Sub Form_Load()
On Error Resume Next

    bSAgregar.Picture = ImageList1.ListImages("nuevo").ExtractIcon
    bSModificar.Picture = ImageList1.ListImages("modificar").ExtractIcon
    
    sNuevo = False: sModificar = False
    LimpioFicha
    InicializoGrilla
       
    HabilitoIngreso Estado:=False
    
    If prmIDCliente <> 0 Then
        ProcesoActivoForm
    Else
        If prmIDEMail <> 0 Then
            '24/08  -->Cambio modo de activación xJuliana Antes decía:AccionModificar prmIDEMail, True
            Cons = "Select EMDIDCliente from EmailDireccion Where EMDCodigo = " & prmIDEMail
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then If Not IsNull(RsAux!EMDIDCliente) Then prmIDCliente = RsAux!EMDIDCliente
            RsAux.Close
            
            If prmIDCliente <> 0 Then
                ProcesoActivoForm
            Else
                AccionModificar prmIDEMail, True
            End If
        Else
            Foco tBuscar
        End If
    End If
    
End Sub

Private Sub ProcesoActivoForm()
On Error GoTo errAyuda

    Dim aValor As Long
    Screen.MousePointer = 11
    
    Cons = "Select * From EMailDireccion, EMailServer " & _
              " Where EMDIDCliente = " & prmIDCliente & _
              " And EMDServidor = EMSCodigo"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    vsLista.Rows = 1
    Do While Not RsAux.EOF
        With vsLista
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!EMDNombre)
            aValor = RsAux!EMDCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!EMDDireccion) & "@" & Trim(RsAux!EMSDireccion)
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!EMSNombre)
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If vsLista.Rows > 1 Then        'Hay direcciones
        Botones True, True, True, False, False, Toolbar1, Me
    
    Else            'Nuevo automatico
        Botones True, False, False, False, False, Toolbar1, Me
        AccionNuevo
        
        Cons = "Select Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2), " & _
                            " Mail = RTrim(CPeApellido1) + ', ' + RTrim(CPeNombre1)" _
               & " From CPersona " _
               & " Where CPeCliente =" & prmIDCliente _
                                                    & " UNION ALL" _
               & " Select Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), Mail = RTrim(CEmFantasia)" _
               & " From CEmpresa " _
               & " Where CEmCliente = " & prmIDCliente
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            tNombre.Text = StrConv(Trim(RsAux!Mail), vbProperCase)
            tCliente.Text = Trim(RsAux!Nombre)
            tCliente.Tag = prmIDCliente
        End If
        RsAux.Close
        
        Foco tDireccion
    End If
    
    Screen.MousePointer = 0
    Exit Sub
errAyuda:
    clsGeneral.OcurrioError "Error al activar el formulario.", Err.Description
    Screen.MousePointer = 0
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

Private Sub Label4_Click()
    Foco tBuscar
End Sub

Private Sub MnuAvisoLlegada_Click()
    EjecutarApp prmPathApp & "AvisoLlegada.exe ", "M=" & CStr(MnuEMTitulo.Tag)
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuEMTitulo_Click()
 
    If Val(paPlantillasIDMail) = 0 Then Exit Sub
    EjecutarApp prmPathApp & "appExploreMsg.exe ", paPlantillasIDMail & ":" & CStr(MnuEMTitulo.Tag)
    
End Sub

Private Sub MnuEMX_Click(Index As Integer)

    If Val(MnuEMTitulo.Tag) = 0 Then Exit Sub
    On Error GoTo errActualizar
    Screen.MousePointer = 11
    Dim miEMail As Long, miLista As Long
        
    miEMail = 0: miLista = 0
    miEMail = Val(MnuEMTitulo.Tag)
    miLista = Val(MnuEMX(Index).Tag)
    
    If miLista <> 0 And miEMail <> 0 Then
        Cons = "Select * from EMailLista Where EMLLista = " & miLista & " And EMLMail = " & miEMail
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            RsAux.AddNew
            RsAux!EMLLista = miLista
            RsAux!EMLMail = miEMail
            RsAux!EMLFAlta = Format(Now, "mm/dd/yyyy hh:mm:ss")
            RsAux.Update
        Else
            RsAux.Delete
        End If
        RsAux.Close
    End If
    
    Screen.MousePointer = 0
    Exit Sub

errActualizar:
    clsGeneral.OcurrioError "Error al procesar la dirección de correo.", Err.Description
    Screen.MousePointer = 0
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
    gIdEMail = 0
    Botones False, False, False, True, True, Toolbar1, Me
    
    LimpioFicha
    HabilitoIngreso
    Foco tCi
  
End Sub

Private Sub AccionModificar(Optional miIdEmail As Long = 0, Optional bDesdeLoad As Boolean = False)
    
    On Error GoTo errModificar
    Screen.MousePointer = 11
    sModificar = True: sNuevo = False
    If miIdEmail = 0 Then gIdEMail = vsLista.Cell(flexcpData, vsLista.Row, 0) Else gIdEMail = miIdEmail
    
    Cons = "Select * from EMailDireccion Left Outer Join Usuario On EMDUsuAlta = UsuCodigo, EMailServer " & _
                " Where EMDServidor = EMSCodigo " & _
                " And EMDCodigo = " & gIdEMail
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Trim(tNombre.Text) = "" Then tNombre.Text = Trim(RsAux!EMDNombre)
        tDireccion.Text = Trim(RsAux!EMDDireccion)
        
        tServidor.Text = Trim(RsAux!EMSDireccion)
        tServidor.Tag = RsAux!EMSCodigo
        
        If Not IsNull(RsAux!EMDIDCliente) And miIdEmail = 0 Then tCliente.Tag = RsAux!EMDIDCliente
        If Not IsNull(RsAux!EMDIDCliente) And bDesdeLoad Then tCliente.Tag = RsAux!EMDIDCliente
        
        If RsAux!EMDNoEnviarInf Then bNoEnviar.Value = vbChecked
        
        If Not IsNull(RsAux!EMDFAlta) Then
            lAlta.Caption = "Ing. el " & Format(RsAux!EMDFAlta, "dd/mm/yy hh:mm")
            If Not IsNull(RsAux!UsuIdentificacion) Then lAlta.Caption = Trim(lAlta.Caption) & ", por " & Trim(RsAux!UsuIdentificacion) & "."
        Else
            lAlta.Caption = "Ingresado antes de Julio del 2001"
        End If
    Else
        gIdEMail = 0
    End If
    RsAux.Close
    
    If gIdEMail = 0 Then
        MsgBox "Posiblemente el registro ha sido eliminado." & vbCrLf & "Vuelva a consultar.", vbExclamation, "Registro Inexistente"
        sModificar = False: Screen.MousePointer = 0: Exit Sub
    End If
    
    If Val(tCliente.Tag) <> 0 Then
        Dim rsCli As rdoResultset
        Cons = "Select CliCodigo, isNull(CliCiRuc, '') as CI, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2)" _
               & " From CPersona, Cliente Where CliCodigo = CPeCliente And CPeCliente =" & Val(tCliente.Tag) _
                                                    & " UNION ALL" _
               & " Select CliCodigo, CI = '', Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')" _
               & " From CEmpresa, Cliente Where CliCodigo = CEmCliente And CEmCliente = " & Val(tCliente.Tag)
               
        Set rsCli = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsCli.EOF Then
            If Trim(rsCli!Ci) <> "" Then tCi.Text = rsCli!Ci
            tCliente.Text = Trim(rsCli!Nombre)
            tCliente.Tag = rsCli!CliCodigo
        End If
        rsCli.Close
    End If
    
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    Foco tCi
    Screen.MousePointer = 0
    Exit Sub
    
errModificar:
    clsGeneral.OcurrioError "Error al cargar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatos(aIdEmail As Long)
On Error GoTo errDatos
    Screen.MousePointer = 11
    LimpioFicha
    
    Cons = "Select * from EMailDireccion Left Outer Join Usuario On EMDUsuAlta = UsuCodigo, EMailServer " & _
                " Where EMDServidor = EMSCodigo " & _
                " And EMDCodigo = " & aIdEmail
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        gIdEMail = RsAux!EMDCodigo
        
        tNombre.Text = Trim(RsAux!EMDNombre)
        tDireccion.Text = Trim(RsAux!EMDDireccion)
        
        tServidor.Text = Trim(RsAux!EMSDireccion)
        tServidor.Tag = RsAux!EMSCodigo
        bSModificar.Enabled = True
        
        If Not IsNull(RsAux!EMDIDCliente) Then tCliente.Tag = RsAux!EMDIDCliente
        
        If RsAux!EMDNoEnviarInf Then bNoEnviar.Value = vbChecked
        
        If Not IsNull(RsAux!EMDFAlta) Then
            lAlta.Caption = "Ing. el " & Format(RsAux!EMDFAlta, "dd/mm/yy hh:mm")
            If Not IsNull(RsAux!UsuIdentificacion) Then lAlta.Caption = Trim(lAlta.Caption) & ", por " & Trim(RsAux!UsuIdentificacion) & "."
        Else
            lAlta.Caption = "Ingresado antes de Julio del 2001"
        End If
    
    Else
        gIdEMail = 0
    End If
    RsAux.Close
    
    If gIdEMail = 0 Then
        MsgBox "Posiblemente el registro ha sido eliminado." & vbCrLf & "Vuelva a consultar.", vbExclamation, "Registro Inexistente"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    If Val(tCliente.Tag) <> 0 Then
        Dim rsCli As rdoResultset
        Cons = "Select CliCodigo, isNull(CliCiRuc, '') as CI, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2)" _
               & " From CPersona, Cliente Where CliCodigo = CPeCliente And CPeCliente =" & Val(tCliente.Tag) _
                                                    & " UNION ALL" _
               & " Select CliCodigo, CI = '', Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')" _
               & " From CEmpresa, Cliente Where CliCodigo = CEmCliente And CEmCliente = " & Val(tCliente.Tag)
               
        Set rsCli = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsCli.EOF Then
            If Trim(rsCli!Ci) <> "" Then tCi.Text = rsCli!Ci
            tCliente.Tag = rsCli!CliCodigo
            tCliente.Text = Trim(rsCli!Nombre)
        End If
        rsCli.Close
    End If
    Screen.MousePointer = 0
    
    Exit Sub
errDatos:
    clsGeneral.OcurrioError "Error al cargar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Sub AccionGrabar()

    If sNuevo Then
        If Not ValidoDireccionConCliente Then Exit Sub
    End If
    If ValidoDireccionCliente > 0 Then Exit Sub
    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    FechaDelServidor
    On Error GoTo errGrabar
    
    Cons = "Select * from EMailDireccion Where EMDCodigo = " & gIdEMail
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If sNuevo Then
'        gIdEMail = Autonumerico(TAutonumerico.EMailDireccion)
        RsAux.AddNew
        RsAux!EMDFAlta = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
        RsAux!EMDUsuAlta = paCodigoDeUsuario
    Else
        RsAux.Edit
        RsAux("EMDUsuModificado") = paCodigoDeUsuario
    End If

'    rsAux!EMDCodigo = gIdEMail
    RsAux!EMDNombre = Trim(tNombre.Text)
    RsAux!EMDDireccion = Trim(tDireccion.Text)
    RsAux!EMDServidor = Val(tServidor.Tag)
    If Val(tCliente.Tag) <> 0 Then RsAux!EMDIDCliente = Val(tCliente.Tag) Else RsAux!EMDIDCliente = Null
    If bNoEnviar.Value = vbChecked Then RsAux!EMDNoEnviarInf = 1 Else RsAux!EMDNoEnviarInf = 0
    RsAux.Update: RsAux.Close
    
    gIdEMail = 0
    sNuevo = False: sModificar = False
    HabilitoIngreso False
    LimpioFicha
    If vsLista.Rows > 1 Then
        If vsLista.Enabled Then vsLista.SetFocus
        Botones True, True, True, False, False, Toolbar1, Me
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    
    If prmIDCliente <> 0 Then Unload Me
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al grabar los datos ingresados.", Err.Description
End Sub

Private Sub AccionEliminar()

    Screen.MousePointer = 11
    On Error GoTo Error
    
    gIdEMail = vsLista.Cell(flexcpData, vsLista.Row, 0)
        
    If MsgBox("Si ud. elimina la dirección de correo, puede generar incosistencias en la base de datos." & vbCrLf & _
                    "Verifique que no existan mensajes enviados a ésta dirección." & vbCrLf & vbCrLf & _
                    "Está seguro de continuar.", vbQuestion + vbYesNo + vbDefaultButton2, "ELIMINAR") = vbNo Then Screen.MousePointer = 0: Exit Sub
                    
    Dim bHay As Boolean: bHay = False
    Cons = "Select * from logdb.dbo.MensajeUsuario Where MUsIdUsuario = " & gIdEMail & " And MUsTipoUsr = " & 4
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then bHay = True
    RsAux.Close
    
    If bHay Then
        MsgBox "Hay mensajes enviados a la dirección de correo que ud. quiere eliminar.", vbExclamation, "Hay mensajes Enviados"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    If MsgBox("Confirma eliminar la dirección de correo '" & vsLista.Cell(flexcpText, vsLista.Row, 1) & "'", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Screen.MousePointer = 0: Exit Sub
    
    Cons = "Select * from EMailDireccion Where EMDCodigo = " & gIdEMail
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Delete: RsAux.Close
    
    Screen.MousePointer = 0
    Exit Sub
    
Error:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al eliminar la dirección de correo.", Err.Description
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
        Dim aValor As Long, sBuscar As String, sServer As String
        Screen.MousePointer = 11
        
        sBuscar = Trim(tBuscar.Text): sServer = ""
        sBuscar = clsGeneral.Replace(sBuscar, " ", "%")
        If InStr(sBuscar, "@") <> 0 Then
            sServer = Trim(Mid(sBuscar, InStr(sBuscar, "@") + 1))
            sBuscar = Mid(sBuscar, 1, InStr(sBuscar, "@") - 1)
        End If
        
        Cons = "Select * From EMailDireccion, EMailServer " & _
                  " Where (EMDNombre like '" & sBuscar & "%' OR EMDDireccion like '" & sBuscar & "%')" & _
                  " And EMDServidor = EMSCodigo"
        If sServer <> "" Then Cons = Cons & " And EMSDireccion like '" & sServer & "%'"
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        vsLista.Rows = 1
        Do While Not RsAux.EOF
            With vsLista
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!EMDNombre)
                aValor = RsAux!EMDCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
                
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!EMDDireccion) & "@" & Trim(RsAux!EMSDireccion)
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!EMSNombre)
                If Not IsNull(RsAux!EMDIDCliente) Then .Cell(flexcpText, .Rows - 1, 3) = "Si" Else .Cell(flexcpText, .Rows - 1, 3) = "No"
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        
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

Private Sub tCi_Change()
    tCliente.Text = "": tCliente.Tag = 0
End Sub

Private Sub tCi_GotFocus()
    tCi.SelStart = 0: tCi.SelLength = 11
End Sub

Private Sub tCi_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF4: Call bCliente_Click
    End Select
    
End Sub

Private Sub tCi_KeyPress(KeyAscii As Integer)

    On Error GoTo errBC
    If KeyAscii = vbKeyReturn Then
        If Trim(tCi.Text) = "" Then tCliente.SetFocus: Exit Sub
        If Val(tCliente.Tag) <> 0 Then tCliente.SetFocus: Exit Sub
        
        Dim aCi As String
        Screen.MousePointer = 11
        
        If Len(tCi.Text) = 7 Then tCi.Text = clsGeneral.AgregoDigitoControlCI(tCi.Text)
                
        'Valido la Cédula ingresada------------------------------------------------------------------------------------------------------
        If Trim(tCi.Text) <> "" Then
            If Len(tCi.Text) <> 8 Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            If Not clsGeneral.CedulaValida(tCi.Text) Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        
        'Busco el Cliente -------------------------------------------------------------------------------------------------------------------
        If Trim(tCi.Text) <> "" Then
            Cons = "Select CliCodigo, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2), " _
                    & " rTrim(CPeApellido1) + ', ' + Rtrim(CPeNombre1) as ID" _
                    & " From Cliente, CPersona " _
                    & " Where CPeCliente = CliCodigo " _
                    & " And CliCiRuc = '" & Trim(tCi.Text) & "'"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                tCliente.Text = Trim(RsAux!Nombre)
                tCliente.Tag = RsAux!CliCodigo
                tNombre.Text = StrConv(Trim(RsAux!ID), vbProperCase)
            End If
            RsAux.Close
            
            If Val(tCliente.Tag) = 0 Then
                MsgBox "No existe un cliente para la cédula ingresada.", vbExclamation, "ATENCIÓN"
            Else
                tCliente.SetFocus
            End If
        End If
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errBC:
    clsGeneral.OcurrioError "Error al buscar el cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn: Foco tNombre
        Case vbKeyF4: Call bCliente_Click
        
        Case vbKeyDelete: tCliente.Text = "": tCliente.Tag = 0
    End Select
    
End Sub

Private Sub tDireccion_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        ProcesoDireccion tDireccion.Text
        Foco tServidor
    End If
    
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tNombre.Text) = "" Then Exit Sub
        If ProcesoDireccion(tNombre.Text) Then
            MsgBox "En el campo 'Nombre' se debe ingresar el Nombre del Cliente.", vbInformation, "Posible Error"
            tNombre.Text = ""
        Else
            Foco tDireccion
        End If
    End If
    
End Sub

Private Function ProcesoDireccion(sDir As String) As Boolean

    ProcesoDireccion = False
    If InStr(sDir, "@") = 0 Then Exit Function
    ProcesoDireccion = True
    
    Dim sServer As String
    sServer = Trim(Mid(sDir, InStr(sDir, "@") + 1))
    sDir = Mid(sDir, 1, InStr(sDir, "@") - 1)
    
    tDireccion.Text = LCase(sDir)
    tServidor.Text = LCase(sServer)
    
End Function

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
        MsgBox "Ingrese la identificación o descripción para la dirección de correo.", vbExclamation, "Falta Identificación"
        Foco tNombre: Exit Function
    End If
    
    If Trim(tDireccion.Text) = "" Then
        MsgBox "Ingrese la dirección de correo.", vbExclamation, "Falta dirección"
        Foco tDireccion: Exit Function
    End If
    If InStr(tDireccion.Text, "@") <> 0 Then
        MsgBox "La dirección se debe ingresar sin la identificación del host.", vbExclamation, "Posible Error"
        Foco tDireccion: Exit Function
    End If
    
    If Val(tServidor.Tag) = 0 Then
        MsgBox "Seleccione el servidor de correo.", vbExclamation, "Falta Servidor"
        Foco tServidor: Exit Function
    End If
    
    Dim sReason As String
    If Not IsEMailAddress(Trim(tDireccion.Text) & "@" & Trim(tServidor.Text), sReason) Then
        MsgBox "La dirección de correo " & Trim(tDireccion.Text) & "@" & Trim(tServidor.Text) & " no es correcta." & vbCrLf & sReason, vbExclamation, "Dirección Incorrecta."
        Screen.MousePointer = 0: Exit Function
    End If
    
    'Valido Nombre o Direccion--------------------------------------------------------------
    'Anulo el control el 06/07/2007
    'cons = "Select * From EMailDireccion" & _
               " Where EMDNombre = '" & Trim(tNombre.Text) & "'" & _
               " And EMDCodigo <> " & gIdEMail
    'Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    'If Not rsAux.EOF Then
    '    MsgBox "Ya existe una dirección de correo con la indentificación '" & Trim(tNombre.Text) & "'.", vbExclamation, "Posible Duplicación"
    '    rsAux.Close: Exit Function
    'End If
    'rsAux.Close

    '--------------------------------------------------------------------------------------------

    ValidoCampos = True
    
End Function

Private Function ValidoDireccionCliente() As Long

    On Error GoTo errValidoD
    Screen.MousePointer = 11
    ValidoDireccionCliente = 0
    
    Dim bHay As Boolean
    Cons = "Select EMDCodigo, EMDNombre as 'Nombre', EMDDireccion as 'Dirección', EMDIDCliente as Cliente From EMailDireccion" & _
               " Where EMDDireccion = '" & Trim(tDireccion.Text) & "'" & _
               " And EMDCodigo <> " & gIdEMail & _
               " And EMDServidor = " & Val(tServidor.Tag)
    'If Val(tCliente.Tag) <> 0 Then cons = cons & " And EMDIdCliente Is Null"
    Cons = Cons & " And EMDIdCliente Is Null"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then bHay = True Else bHay = False
    RsAux.Close
    
    If bHay Then
        Dim aValor As Long
        Dim miLista As New clsListadeAyuda
        aValor = miLista.ActivarAyuda(cBase, Cons, 4800, 1, "Igual Dirección, Sin Asignar")
        If aValor > 0 Then aValor = miLista.RetornoDatoSeleccionado(0)
        Set miLista = Nothing
        
        If aValor <> 0 Then
            If MsgBox("Ud. seleccionó una dirección de la lista." & vbCrLf & "Quiere editar esta dirección para agregarle el cliente.", vbQuestion + vbYesNo, "Cancelar el Ingreso") = vbYes Then
                ValidoDireccionCliente = aValor
                AccionModificar miIdEmail:=aValor
            End If
        End If
    End If
    Screen.MousePointer = 0
    Exit Function
     
errValidoD:
    clsGeneral.OcurrioError "Error al validar la dirección.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function ValidoDireccionConCliente() As Boolean

    On Error GoTo errValidoD
    Screen.MousePointer = 11
    ValidoDireccionConCliente = False
    
    
    Dim aQ As Integer, aIDSel As Long, aTSel As String
    aQ = 0
    Cons = "Select EMDCodigo, EMDNombre as 'Nombre', EMDDireccion as 'Dirección', EMDIDCliente as Cliente From EMailDireccion" & _
               " Where EMDDireccion = '" & Trim(tDireccion.Text) & "'" & _
               " And EMDCodigo <> " & gIdEMail & _
               " And EMDServidor = " & Val(tServidor.Tag) & _
               " And EMDIdCliente Is Not Null"
    If Val(tCliente.Tag) <> 0 Then Cons = Cons & " And EMDIdCliente <> " & Val(tCliente.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        aQ = 1: aIDSel = RsAux!EMDCodigo
        aTSel = "Ya existe un cliente con la dirección ingresada." & vbCr & _
                    "La identificación de la dirección es: " & RsAux!Nombre
        RsAux.MoveNext
        If Not RsAux.EOF Then aQ = 2
    End If
    RsAux.Close
    
    If aQ = 0 Then
        ValidoDireccionConCliente = True
        Screen.MousePointer = 0
        Exit Function
    End If
    
    Select Case aQ
        Case 1:
                If MsgBox(aTSel & vbCr & vbCr & "Está seguro que quiere continuar con el ingreso.", vbQuestion + vbYesNo + vbDefaultButton2, "Dirección Existente") = vbNo Then
                    Screen.MousePointer = 0
                    Exit Function
                End If
                
        Case 2:
                MsgBox "Ya existen clientes con la dirección de correo ingresada" & vbCrLf & _
                            "Presione Aceptar para ver la lista de clientes.", vbInformation, "Dirección Existente"
                            
                Dim aValor As Long
                Dim miLista As New clsListadeAyuda
                aValor = miLista.ActivarAyuda(cBase, Cons, 4800, 1, "Direcciones con Clientes")
                'If aValor > 0 Then aValor = miLista.RetornoDatoSeleccionado(0)
                Set miLista = Nothing
        
                If MsgBox("Está seguro que quiere continuar con el ingreso.", vbQuestion + vbYesNo + vbDefaultButton2, "Continúa ?") = vbNo Then
                    Screen.MousePointer = 0
                    Exit Function
                End If
    End Select
    
    ValidoDireccionConCliente = True
    Screen.MousePointer = 0
    Exit Function
     
errValidoD:
    clsGeneral.OcurrioError "Error al validar la dirección con clientes.", Err.Description
    Screen.MousePointer = 0
End Function


Private Sub HabilitoIngreso(Optional Estado As Boolean = True)

Dim bkColor As Long
    
    If Estado Then bkColor = Colores.Blanco Else bkColor = Colores.Gris
    
    tNombre.Enabled = Estado: tNombre.BackColor = bkColor
    tDireccion.Enabled = Estado: tDireccion.BackColor = bkColor
    tServidor.Locked = Not Estado: tServidor.BackColor = bkColor
    bSAgregar.Enabled = Estado
    bSModificar.Enabled = Estado
    
    tCi.Enabled = Estado: tCi.BackColor = bkColor
    tCliente.Enabled = Estado: tCliente.BackColor = bkColor
    bCliente.Enabled = Estado
    bNoEnviar.Enabled = Estado
    
    tBuscar.Enabled = Not Estado
    vsLista.Enabled = Not Estado
    
End Sub

Private Sub LimpioFicha()

    tNombre.Text = ""
    tDireccion.Text = ""
    tServidor.Text = ""
    bNoEnviar.Value = vbUnchecked
    
    tCi.Text = ""
    tCliente.Text = "": tCliente.Tag = 0
    lAlta.Caption = ""
    
End Sub


Private Sub InicializoGrilla()

    On Error Resume Next
    With vsLista
        .Cols = 1: .Rows = 1
        .FormatString = "<Identificación|<Dirección|<Servidor|<Cliente"
            
        .WordWrap = True
        .ColWidth(0) = 1500: .ColWidth(1) = 2800: .ColWidth(2) = 1000
        
        .ExtendLastCol = True: .FixedCols = 0
    End With
      
End Sub


Private Sub tServidor_Change()
    tServidor.Tag = 0
End Sub

Private Sub tServidor_GotFocus()
    tServidor.SelStart = 0: tServidor.SelLength = Len(tServidor.Text)
End Sub

Private Sub tServidor_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF3:
            If Not tServidor.Locked Then Call bSAgregar_Click
            
        Case vbKeyF2: Call bSModificar_Click
    End Select
    
End Sub

Private Sub tServidor_KeyPress(KeyAscii As Integer)

    On Error GoTo errBuscar
    If tServidor.Locked Then vsLista.SetFocus: Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        If Val(tServidor.Tag) <> 0 Then AccionGrabar: Exit Sub
        
        If Trim(tServidor.Text) = "" Then Exit Sub
        Screen.MousePointer = 11
        Dim aQ As Long, aId As Long, aTexto As String
        aQ = 0
        
        Cons = "Select EMSCodigo, EMSDireccion as 'Dirección', EMSNombre as Servidor" & _
                " from EMailServer " & _
                " Where EMSNombre like '" & Trim(tServidor.Text) & "%' or EMSDireccion like '" & Trim(tServidor.Text) & "%'" & _
                " Order by EMSNombre"
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aQ = 1: aId = RsAux!EMSCodigo: aTexto = Trim(RsAux(1))
            RsAux.MoveNext: If Not RsAux.EOF Then aQ = 2
        End If
        RsAux.Close
        
        Select Case aQ
            Case 0:
                    If MsgBox("No existe un servidor de correo para la descripción ingresada." & vbCrLf & "Desea ingresarlo ", vbQuestion + vbYesNo, "No hay datos") = vbYes Then
                        frmMaServer.prmTxtNombre = ""
                        frmMaServer.prmTxtHost = Trim(tServidor.Text)
                        frmMaServer.prmIdServer = 0
                        frmMaServer.Show vbModal, Me
                        Me.Refresh
                        If frmMaServer.prmTxtNombre <> "" Then
                            tServidor.Text = Trim(frmMaServer.prmTxtNombre)
                            
                            Cons = "Select * from EmailServer Where EMSDireccion = '" & Trim(tServidor.Text) & "'"
                            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                            If Not RsAux.EOF Then tServidor.Tag = RsAux!EMSCodigo
                            RsAux.Close
                            
                            'If Val(tCliente.Tag) = 0 Then
                            '    If ValidoDireccionCliente = 0 Then tCi.SetFocus
                            'Else
                                AccionGrabar
                            'End If
                        End If
                    End If
            
            Case 1:
                    tServidor.Text = aTexto: tServidor.Tag = aId
                    'If Val(tCliente.Tag) = 0 Then
                    '    If ValidoDireccionCliente = 0 Then tCi.SetFocus
                    'Else
                        AccionGrabar
                    'End If
        
            Case 2:
                    Dim aValor As Long
                    Dim aLista As New clsListadeAyuda
                    
                    aValor = aLista.ActivarAyuda(cBase, Cons, 5500, 1)
                    
                    If aValor > 0 Then
                        tServidor.Text = Trim(aLista.RetornoDatoSeleccionado(1))
                        tServidor.Tag = aLista.RetornoDatoSeleccionado(0)
                    End If
                    Set aLista = Nothing
                    
                    If aValor > 0 Then AccionGrabar
                    
        End Select
        Screen.MousePointer = 0
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    clsGeneral.OcurrioError "Error al procesar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub vsLista_Click()
    On Error Resume Next
    If vsLista.Rows > 1 Then
        If gIdEMail <> vsLista.Cell(flexcpData, vsLista.Row, 0) Then CargoDatos vsLista.Cell(flexcpData, vsLista.Row, 0)
    End If
End Sub

Private Sub vsLista_DblClick()

    If vsLista.Rows = 1 Then Exit Sub
    
    Dim aValor As Long
    aValor = vsLista.Cell(flexcpData, vsLista.Row, 0)
    frmListas.prmDireccion = vsLista.Cell(flexcpText, vsLista.Row, 1)
    frmListas.prmIDEMail = aValor
    frmListas.Show vbModal, Me
    
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

Private Sub vsLista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If vsLista.Rows = 1 Then Exit Sub
    If Button = vbRightButton Then
        Dim aValor As Long
        aValor = vsLista.Cell(flexcpData, vsLista.Row, 0)
        If aValor = 0 Then Exit Sub
        AccionMenuEmail aValor
    End If
    
End Sub

Private Sub AccionMenuEmail(miIDMail As Long)
    On Error GoTo errMnuEmail
    
    Dim miCodigo As Long, aQ As Integer
    Dim I As Integer, aTitulo As String
    
    aTitulo = vsLista.Cell(flexcpText, vsLista.Row, 1)
    MnuEMTitulo.Caption = Trim(aTitulo)
    MnuEMTitulo.Tag = miIDMail
    
    MnuEMX(0).Visible = True
    For I = 1 To MnuEMX.UBound
        Unload MnuEMX(I)
    Next
        
    Cons = "Select * from ListaDistribucion left Outer Join EMailLista On LiDCodigo = EMLLista And EMLMail = " & miIDMail & _
                " Where LiDHabilitado = 1" & _
                " Order by LiDNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        I = MnuEMX.UBound + 1
        Load MnuEMX(I)
        With MnuEMX(I)
            .Caption = Trim(RsAux!LiDNombre)
            If Not IsNull(RsAux!LiDExcluye) Then If RsAux!LiDExcluye = 1 Then .Caption = "NO " & Trim(.Caption)
            .Tag = RsAux!LiDCodigo
            If Not IsNull(RsAux!EMLMail) Then .Checked = True Else .Checked = False
            .Visible = True
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close

    If MnuEMX.UBound > 0 Then
        MnuEMX(0).Visible = False
        PopupMenu MnuEMail, DefaultMenu:=MnuEMTitulo
    End If
    
    Exit Sub
    
errMnuEmail:
End Sub


