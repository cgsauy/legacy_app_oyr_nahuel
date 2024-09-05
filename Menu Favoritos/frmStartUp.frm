VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStartUp 
   Caption         =   "StartUp Menu"
   ClientHeight    =   6855
   ClientLeft      =   1275
   ClientTop       =   2640
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStartUp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   3600
   Begin MSComctlLib.TreeView vsMain 
      Height          =   2475
      Left            =   300
      TabIndex        =   0
      Top             =   1920
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4366
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.PictureBox picMenuPerfil 
      Height          =   435
      Left            =   180
      ScaleHeight     =   375
      ScaleWidth      =   855
      TabIndex        =   5
      Top             =   660
      Width           =   915
      Begin VB.Label lMenuPerfil 
         Alignment       =   2  'Center
         Caption         =   " &Perfiles"
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar barPerfil 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "menuperfil"
            Style           =   4
            Object.Width           =   800
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "perfil"
            Object.ToolTipText     =   "Editar perfil."
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevoperfil"
            Object.ToolTipText     =   "Nuevo perfil."
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "login"
            Object.ToolTipText     =   "Cambiar login."
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "clave"
            Object.ToolTipText     =   "Cambio de clave."
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Object.ToolTipText     =   "Actualizar."
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6600
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5750
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   159
            MinWidth        =   2
            Key             =   "login"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   5340
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
            Picture         =   "frmStartUp.frx":08CA
            Key             =   "perfil"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartUp.frx":0BEE
            Key             =   "nuevoperfil"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartUp.frx":0F0A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartUp.frx":135E
            Key             =   "app"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartUp.frx":17B2
            Key             =   "mnuopen"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartUp.frx":3F66
            Key             =   "mnuclose"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartUp.frx":671A
            Key             =   "login"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartUp.frx":6B6E
            Key             =   "clave"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusA 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6297
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuPerfil 
      Caption         =   "MnuPerfil"
      Visible         =   0   'False
      Begin VB.Menu PerMenu 
         Caption         =   "(Perfiles de Usuario)"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsIni As rdoResultset
Dim aValor As Long

Public prmIdLogin As Long

Private Sub barPerfil_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "refresh": AccionRefresh
        Case "perfil": AccionEditarPerfil
        Case "nuevoperfil": AccionEditarPerfil Nuevo:=True
        Case "clave": AccionUsuarios
        
        Case "login": AccionCambiarLogin
    End Select
    
End Sub

Private Sub AccionRefresh()
    
    On Error GoTo errRefresh
    If prmIdLogin = 0 Then Exit Sub
    
    CargoPerfiles
    If Val(vsMain.Tag) <> 0 Then CargoMenu Val(vsMain.Tag)
    
    Exit Sub
errRefresh:
    clsGeneral.OcurrioError rdoErrors, "Refresh", "Ocurrió un error al recargar el formulario.", Err.Description
End Sub

Private Sub AccionEditarPerfil(Optional Nuevo As Boolean = False)
    
    On Error GoTo errPerfil
    If prmIdLogin = 0 Then Exit Sub
    
    frmPerfil.prmIdLogin = prmIdLogin
    If Not Nuevo Then frmPerfil.prmIDPerfil = Val(vsMain.Tag) Else frmPerfil.prmIDPerfil = 0
    frmPerfil.Show vbModal, Me
    Exit Sub

errPerfil:
    clsGeneral.OcurrioError rdoErrors, "Periles", "Ocurrió un error al invocar al formulario de perfiles.", Err.Description
End Sub

Private Sub AccionUsuarios()
    
    On Error GoTo errPerfil
    If prmIdLogin = 0 Then Exit Sub
    
    frmUsuarios.prmIDUsr = prmIdLogin
    frmUsuarios.prmPassword = True
    frmUsuarios.Show vbModal, Me
    Exit Sub

errPerfil:
    clsGeneral.OcurrioError rdoErrors, "Logins", "Ocurrió un error al invocar al formulario de logins.", Err.Description
End Sub


Private Sub barPerfil_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RefreshMenu picMenuPerfil
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    vsMain.SetFocus
End Sub

Private Sub Form_Load()

    On Error Resume Next
    ObtengoSeteoForm Me
    
     With picMenuPerfil
        .Top = barPerfil.Buttons("menuperfil").Top + 40
        .Height = barPerfil.Buttons("menuperfil").Height - 40
        .Left = barPerfil.Buttons("menuperfil").Left
        .Width = barPerfil.Buttons("menuperfil").Width
        .BorderStyle = 0
        lMenuPerfil.Left = .ScaleLeft + 60: lMenuPerfil.Top = .ScaleTop + 45
        lMenuPerfil.Width = .ScaleWidth - 120: lMenuPerfil.Height = .ScaleHeight
    End With
    
    vsMain.ImageList = ImageList1
    
    Set barPerfil.ImageList = ImageList1
    barPerfil.Buttons("perfil").Image = ImageList1.ListImages("perfil").Index
    barPerfil.Buttons("nuevoperfil").Image = ImageList1.ListImages("nuevoperfil").Index
    barPerfil.Buttons("login").Image = ImageList1.ListImages("login").Index
    barPerfil.Buttons("clave").Image = ImageList1.ListImages("clave").Index
    barPerfil.Buttons("refresh").Image = ImageList1.ListImages("refresh").Index
    
    CargoPerfiles
    CargoDatosLogin
    EstadoBotones
    
End Sub

Private Sub CargoDatosLogin()

    On Error GoTo errLogin
    Dim aPerfilD As Long
    
    StatusB.Panels("login").Text = ""
    
    Cons = "Select * from Login Where LogID = " & prmIdLogin
    Set rsIni = cQuery.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsIni.EOF Then
        aPerfilD = 0
        If Not IsNull(rsIni!LogPerfilD) Then aPerfilD = rsIni!LogPerfilD
        StatusB.Panels("login").Text = Trim(rsIni!LogLogin) & "  "
    
    End If
    rsIni.Close
    
    StatusB.Panels("login").Picture = ImageList1.ListImages("login").ExtractIcon
    
    CargoMenu aPerfilD
    Exit Sub
    
errLogin:
    clsGeneral.OcurrioError rdoErrors, "Datos_Login", "Ocurrió un error al cargar los datos de login.", Err.Description
End Sub

Private Sub CargoPerfiles()

    Dim aKey As String, aMenu As String, I As Integer
    
    PerMenu.Item(0).Visible = True
    Do While PerMenu.Count > 1
        Unload PerMenu.Item(PerMenu.UBound)
    Loop
    
    I = 0
    Cons = "Select * from Perfil Where PerLogin = " & prmIdLogin & " Order by PerNombre"
    Set rsIni = cQuery.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsIni.EOF
        aMenu = Trim(rsIni!PerNombre)
        I = I + 1
        
        Load PerMenu(I)
        PerMenu.Item(I).Caption = aMenu
        PerMenu.Item(I).Tag = rsIni!PerID
        PerMenu.Item(I).Visible = True
        
        rsIni.MoveNext
    Loop
    rsIni.Close
    
    If PerMenu.Count > 1 Then PerMenu.Item(0).Visible = False
    
End Sub

Private Sub CargoMenu(aPerfil As Long)

Dim aIndex As Integer, UltimoMnu As Integer

    On Error GoTo errMenu
    UltimoMnu = 0
    StatusA.Panels(1).Text = ""
    vsMain.Nodes.Clear
    vsMain.Tag = 0
    If aPerfil = 0 Then Exit Sub
    
    Cons = "Select * from PerfilMenu Where PMePerfil = " & aPerfil & " Order by PMeMenu ASC"
    Set rsIni = cQuery.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsIni.EOF
        With vsMain
            If rsIni!PMeNivelMenu = 0 Then
                .Nodes.Add , , LCase(rsIni!PMeNombre), Trim(rsIni!PMeNombre)
            Else
                .Nodes.Add UltimoMnu, tvwChild, LCase(rsIni!PMeNombre), Trim(rsIni!PMeNombre)
            End If
            aIndex = .Nodes(LCase(rsIni!PMeNombre)).Index
            
            If Not IsNull(rsIni!PMeIDApp) Then
                .Nodes(aIndex).Tag = rsIni!PMeIDApp
                .Nodes(aIndex).Image = ImageList1.ListImages("app").Index
            Else
                UltimoMnu = aIndex
                .Nodes(aIndex).Tag = 0
                .Nodes(aIndex).Image = ImageList1.ListImages("mnuclose").Index
                .Nodes(aIndex).ExpandedImage = ImageList1.ListImages("mnuopen").Index
                .Nodes(aIndex).Expanded = True
            End If
         
         End With
        rsIni.MoveNext
    Loop
    rsIni.Close
    
    vsMain.Tag = aPerfil    'Marco en el Tag para trabajar dps. (actualizar, editar, etc)
    
    'Marco el menu como seleccionado------------------------------------------------------------------------------------
    On Error Resume Next
    Dim aMnu As Menu
    For Each aMnu In PerMenu
        If Val(aMnu.Tag) <> aPerfil Then
            aMnu.Checked = False
        Else
            aMnu.Checked = True
            StatusA.Panels(1).Text = aMnu.Caption
        End If
    Next
    '---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    Exit Sub

errMenu:
    clsGeneral.OcurrioError rdoErrors, "Cargar_Perfil", "Ocurrió un error al cargar el perfil seleccionado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    vsMain.Left = Me.ScaleLeft
    vsMain.Top = Me.ScaleTop + StatusA.Height + barPerfil.Height
    vsMain.Height = Me.ScaleHeight - (StatusA.Height + StatusB.Height + barPerfil.Height)
    vsMain.Width = Me.ScaleWidth
          
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion cQuery
    CierroEntorno
    
    Set clsGeneral = Nothing
    End
    
End Sub


Private Sub lMenuPerfil_Click()
    On Error Resume Next
    If picMenuPerfil.Tag = 1 Then
        RefreshMenu picMenuPerfil
        PopupMenu MnuPerfil, , lMenuPerfil.Left, picMenuPerfil.Top + picMenuPerfil.Height
    Else
        RefreshMenu picMenuPerfil
        Call lMenuPerfil_MouseMove(1, 1, 1, 1)
    End If
End Sub

Private Sub lMenuPerfil_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errMarco
    If Val(picMenuPerfil.Tag) <> 0 Then Exit Sub
    Dim a As New clsDiseño
    a.DibujoMarcoAControl picMenuPerfil.hDC, picMenuPerfil.Height, picMenuPerfil.Width
    Set a = Nothing
    picMenuPerfil.Tag = 1
errMarco:
End Sub

Private Sub PerMenu_Click(Index As Integer)
    On Error GoTo errMnu
    
    If Val(PerMenu(Index).Tag) = 0 Then Exit Sub
    CargoMenu Val(PerMenu(Index).Tag)
    
errMnu:
End Sub

Private Sub picMenuPerfil_GotFocus()
    Call lMenuPerfil_Click
End Sub

Private Sub StatusA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RefreshMenu picMenuPerfil
End Sub


Private Sub vsMain_DblClick()
    Call vsMain_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub vsMain_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo errKeyDown
    If vsMain.Nodes.Count = 0 Then Exit Sub
    Dim aIDApp As Long
    aIDApp = vsMain.SelectedItem.Tag
    
    Select Case KeyCode
        Case vbKeyReturn:
                If aIDApp = 0 Then Exit Sub
                AccionEjecutarApp aIDApp
    End Select

errKeyDown:
End Sub

Private Sub AccionEjecutarApp(idApp As Long)

    On Error GoTo errApp
    If Not HayPermisos(idApp, prmIdLogin) Then
        MsgBox "Ud. no tiene permiso de ejecución para la aplicación seleccionada.", vbInformation, "Falta Permiso de Ejecución"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Dim plngRet As Long, pathApp As String
    
    pathApp = App.Path & "\Template " & idApp
    plngRet = Shell(pathApp, 1)
    
    Screen.MousePointer = 0
    Exit Sub
    
errApp:
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al ejecutar la aplicación " & pathApp & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbCritical, "Error de Aplicación"
End Sub

Private Sub RefreshMenu(aMenu As PictureBox)
    On Error Resume Next
    If Val(aMenu.Tag) = 0 Then Exit Sub
    aMenu.Tag = 0
    aMenu.Refresh
End Sub

Private Sub vsMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RefreshMenu picMenuPerfil
End Sub

Private Function HayPermisos(idApp As Long, idLogin As Long) As Boolean

    On Error GoTo errPermisos
    Screen.MousePointer = 11
    HayPermisos = False
    
    Cons = "Select * From LoginNivel" & _
                " Where LNiLogin = " & idLogin & _
                " And LNiNivel in (Select ANiNivel from AppNivel Where ANiApp = " & idApp & ") "
    Set RsAux = cQuery.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then HayPermisos = True
    RsAux.Close
    
    Screen.MousePointer = 0
    Exit Function
    
errPermisos:
    clsGeneral.OcurrioError rdoErrors, "Permisos", "Ocurrió un error al verificar permisos de acceso.", Err.Description
    Screen.MousePointer = 0
End Function


Private Sub AccionCambiarLogin()
    
    On Error GoTo errCL
    frmLogin.Show vbModal, Me
    
    prmIdLogin = 0
    If frmLogin.prmLoginOK Then prmIdLogin = frmLogin.prmIdLogin
    
    CargoPerfiles
    CargoDatosLogin
    
    EstadoBotones
    Exit Sub

errCL:
    clsGeneral.OcurrioError rdoErrors, "CambioLogin", "Ocurrió un error al realizar el cambio del login.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub EstadoBotones()

    On Error Resume Next
    If prmIdLogin = 0 Then
        barPerfil.Buttons("perfil").Enabled = False
        barPerfil.Buttons("nuevoperfil").Enabled = False
        barPerfil.Buttons("clave").Enabled = False
        barPerfil.Buttons("refresh").Enabled = False
    Else
        barPerfil.Buttons("perfil").Enabled = True
        barPerfil.Buttons("nuevoperfil").Enabled = True
        barPerfil.Buttons("clave").Enabled = True
        barPerfil.Buttons("refresh").Enabled = True
    End If
End Sub
