VERSION 5.00
Begin VB.UserControl ctrEMails 
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1455
   ScaleWidth      =   4800
   Begin VB.Timer tmMenus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   360
   End
   Begin VB.ComboBox cDirs 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   540
      Width           =   3195
   End
   Begin VB.Label lEMail 
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   315
   End
   Begin VB.Label lQDirs 
      BackStyle       =   0  'Transparent
      Caption         =   "(13)"
      Height          =   195
      Left            =   2220
      TabIndex        =   1
      Top             =   1020
      Width           =   255
   End
   Begin VB.Menu MnuBDerecho 
      Caption         =   "MnuBDerecho"
      Begin VB.Menu MnuBDTitulo 
         Caption         =   "MnuBDTitulo"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBDL1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBDPlantillaMails 
         Caption         =   "&Ver mails recibidos y enviados"
      End
      Begin VB.Menu MnuBDMail 
         Caption         =   "&Enviar mail"
      End
      Begin VB.Menu MnuBDCopy 
         Caption         =   "&Copiar dirección"
      End
      Begin VB.Menu MnuBDEliminar 
         Caption         =   "Eliminar dirección"
      End
      Begin VB.Menu MnuBDL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBDLista 
         Caption         =   "Listas"
         Index           =   0
      End
      Begin VB.Menu MnuBDL3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBDCancelar 
         Caption         =   "Cancelar"
      End
   End
   Begin VB.Menu MnuAdd 
      Caption         =   "MnuAdd"
      Visible         =   0   'False
      Begin VB.Menu MnuAddTitulo 
         Caption         =   "¿Qué desea hacer?"
      End
      Begin VB.Menu MnuAddL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCambiar 
         Caption         =   "Sustituir"
         Index           =   0
      End
      Begin VB.Menu MnuNewDir 
         Caption         =   "Es una nueva dirección"
      End
      Begin VB.Menu MnuAddL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAddCancelar 
         Caption         =   "&Cancelar"
      End
   End
   Begin VB.Menu MnuServer 
      Caption         =   "MnuServer"
      Begin VB.Menu MnuSerTitulo 
         Caption         =   "Confirme el Servidor de Correo"
      End
      Begin VB.Menu MnuSerL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSerOpcion 
         Caption         =   "MnuSerOpcion"
         Index           =   0
      End
      Begin VB.Menu MnuSerNew 
         Caption         =   "nuevo servidor"
      End
      Begin VB.Menu MnuSerL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSerCancelar 
         Caption         =   "&Cancelar"
      End
   End
   Begin VB.Menu MnuDirs 
      Caption         =   "MnuDirs"
      Begin VB.Menu MnuDiTitulo 
         Caption         =   "Seleccione una dirección de correo"
      End
      Begin VB.Menu MnuDiL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDiItem 
         Caption         =   "MnuDi"
         Index           =   0
      End
      Begin VB.Menu MnuDiL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDiCancelar 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "ctrEMails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Agregué estos 2 props para poder decir donde quiero que se abra
'el popup del botón derecho y no romper lo que tiene ya la ocx.
'Si estos dos están en cero toma la pos del control.
Public Xbd As Integer
Public Ybd As Integer

Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Const m_def_BackColor = &H8000000F
Const m_def_ForeColor = &H80000012

Dim rdoCBase As rdoConnection
Dim m_Modalidad As Integer
Dim m_IdParaDirs As String
Dim m_IDUsuario As Long
Dim m_Enabled  As Boolean

Const m_def_Modalidad = 1

Public Event KeyPress(KeyAscii As Integer)

Private prmIDCliente As Long
Private mSQL As String

Private prmPathApp As String, prmPathSystem As String
Private prmIDPlantillaEMails As Long

Public Function OpenControl(objConnect As rdoConnection) As Boolean
On Error GoTo errFnc

    OpenControl = False
    Set rdoCBase = objConnect
    
    ClearObjects
    
    'Carga de Parametros    ------------------------------------------------------------------------------
    mSQL = "Select * from Parametro Where ParNombre IN ('PlantillasIDMail', 'PathApp', 'PathSystem')"
    Set rsAux = rdoCBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            Case "plantillasidmail": prmIDPlantillaEMails = rsAux!ParValor
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto) & "\"
            Case "pathsystem": prmPathSystem = Trim(rsAux!ParTexto) & "\"
        End Select
        rsAux.MoveNext
    Loop
    '----------------------------------------------------------------------------------------------------------
    OpenControl = True
    Exit Function
    
errFnc:
    clsGeneral.OcurrioError "Error al abrir el control (método OpenControl)."
End Function

Public Function ClearObjects()
    On Error Resume Next
        
    prmIDCliente = 0
    cDirs.Clear
    cDirs.Text = ""
    lQDirs.Caption = ""
    
    lEMail.ForeColor = &H808080
    
    ReDim arrCorreo(0)
    arrCorreo(0).ClaveCompleta = ""

    m_IdParaDirs = ""
    
End Function

Public Function CargarDatos(idCliente As Long) As Boolean

On Error GoTo errFnc
    
    CargarDatos = False
    
    ClearObjects
    prmIDCliente = idCliente

    CargarDatos = fnc_CargoMails

errFnc:
End Function

Public Function UpdatearDirecciones(alIdCliente As Long) As Boolean
' Updatea las nuevas direcciones ingresadas que tengan el id de cliente nulo y que esten en el array
On Error GoTo errFnc

    UpdatearDirecciones = False
    If alIdCliente = 0 Then Exit Function
    
    If arrCorreo(0).ClaveCompleta = "" Then
        UpdatearDirecciones = True
        Exit Function
    End If
        
    Dim mIDs As String, arrI As Integer
    
    For arrI = LBound(arrCorreo) To UBound(arrCorreo)
        mIDs = mIDs & arrCorreo(arrI).DireccionID
        If arrI < UBound(arrCorreo) Then mIDs = mIDs & ","
    Next

    
    mSQL = "Select * from EMailDireccion " & _
                " Where EMDCodigo IN (" & mIDs & ")" & _
                " And (EMDEliminado Is Null OR EMDEliminado = 0)" & _
                " And EMDIdCliente Is Null"
    
    Set rsAux = rdoCBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        rsAux.Edit
        rsAux!EMDIdCliente = alIdCliente
        rsAux.Update
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    UpdatearDirecciones = True
    prmIDCliente = alIdCliente
    
errFnc:
End Function

Public Function GetDirecciones() As String
'   Retorna idDir:Parte1deDir:idServidor    |   ......  |
    On Error GoTo errFnc
    GetDirecciones = ""
    Dim mIDs As String, arrI As Integer
    
    For arrI = LBound(arrCorreo) To UBound(arrCorreo)
        If arrCorreo(arrI).ClaveCompleta <> "" Then
        mIDs = mIDs & arrCorreo(arrI).DireccionID & ":" & _
                               arrCorreo(arrI).Direccion & ":" & _
                               arrCorreo(arrI).ServidorID
        End If
        If arrI < UBound(arrCorreo) Then mIDs = mIDs & "|"
    Next
    
    GetDirecciones = mIDs
errFnc:
End Function

Private Sub cDirs_KeyDown(KeyCode As Integer, Shift As Integer)

    If cDirs.ListIndex = -1 Then Exit Sub
    
    On Error GoTo errKD
    Select Case KeyCode
        Case 93
            mnu_BDerecho
        
        Case vbKeyF2
            Dim modDireccion As String, modNombre As String, mNewV As String
            
            zfn_PartoEmail cDirs.Text, modDireccion, modNombre
            mNewV = InputBox("Ingrese el nombre o identificación para la dirección de correo.", "Cambiar Nombre", modNombre)
            If mNewV = "" Or (mNewV = modNombre) Then Exit Sub
            
            cDirs.Text = modDireccion & " " & zfn_EMNombre(mNewV, 1)
            Call cDirs_KeyPress(vbKeyReturn)
    End Select
    
errKD:
End Sub

Private Sub cDirs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(cDirs.Text) <> "" And cDirs.ListIndex = -1 Then
            fnc_ProcesoDireccion (cDirs.Text)
        Else
            RaiseEvent KeyPress(vbKeyReturn)
        End If
        
    End If
    
End Sub

Private Sub lEMail_Click()
    mnu_Direcciones
End Sub

Private Sub lQDirs_Click()
    mnu_Direcciones
End Sub

Private Sub MnuBDCopy_Click()
    
    On Error Resume Next
    
    Clipboard.Clear
    Dim mDirEmail As String
    
    zfn_PartoEmail cDirs.Text, mDirEmail, ""
    Clipboard.SetText mDirEmail, vbCFText
    
End Sub

Private Sub MnuBDEliminar_Click()
On Error GoTo errElim

    If Val(MnuBDTitulo.Tag) = 0 Then Exit Sub
    If MsgBox("Confirma eliminar la dirección de correo: " & cDirs.Text, vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Dirección") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    
    Dim miEMail As Long, miDir As String
    miEMail = Val(MnuBDTitulo.Tag)
    zfn_PartoEmail Trim(cDirs.Text), miDir, ""
    
    If miEMail <> 0 Then
        '1)  Elimino de la base de Datos
        mSQL = "Select * from EMailDireccion Where EMDCodigo = " & miEMail
        Set rsAux = rdoCBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            rsAux.Edit
            rsAux!EMDEliminado = 1
            rsAux.Update
        End If
        rsAux.Close
    
        '2) Elimino el item del array de datos
        idX = arrIndex(miDir)
        arrDeleteIndex idX
        
        '3) Elimino el item del combo
        cDirs.RemoveItem cDirs.ListIndex
        'If cDirs.ListCount > 0 Then lQDirs.Caption = "(" & cDirs.ListCount & ")" Else lQDirs.Caption = ""
        fnc_CaptionQDirs
    
        Select Case cDirs.ListCount
            Case 0: lEMail.ForeColor = &H808080
            Case Else: lEMail.ForeColor = vbRed
        End Select
    
    End If

    Screen.MousePointer = 0
    Exit Sub
errElim:
    clsGeneral.OcurrioError "Error al eliminar la dirección de correo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuBDLista_Click(Index As Integer)

    If Val(MnuBDTitulo.Tag) = 0 Then Exit Sub
    On Error GoTo errActualizar
    Screen.MousePointer = 11
    Dim miEMail As Long, miLista As Long
        
    miEMail = 0: miLista = 0
    miEMail = Val(MnuBDTitulo.Tag)
    miLista = Val(MnuBDLista(Index).Tag)
    
    If miLista <> 0 And miEMail <> 0 Then
        mSQL = "Select * from EMailLista Where EMLLista = " & miLista & " And EMLMail = " & miEMail
        Set rsAux = rdoCBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
        If rsAux.EOF Then
            rsAux.AddNew
            rsAux!EMLLista = miLista
            rsAux!EMLMail = miEMail
            rsAux!EMLFAlta = Format(Now, "mm/dd/yyyy hh:mm:ss")
            rsAux.Update
        Else
            rsAux.Delete
        End If
        rsAux.Close
    End If
    
    Screen.MousePointer = 0
    Exit Sub

errActualizar:
    clsGeneral.OcurrioError "Error al procesar la dirección de correo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuBDMail_Click()

    EjecutarApp prmPathSystem & "orCreateMsg.exe ", "/P4" & CStr(MnuBDTitulo.Tag)
End Sub

Private Sub MnuBDPlantillaMails_Click()
    
    If Val(prmIDPlantillaEMails) = 0 Then Exit Sub
        
    EjecutarApp prmPathApp & "appExploreMsg.exe ", prmIDPlantillaEMails & ":" & CStr(MnuBDTitulo.Tag)

End Sub

Private Sub MnuCambiar_Click(Index As Integer)
    MnuAddTitulo.Tag = MnuCambiar(Index).Tag
End Sub

Private Sub MnuDiItem_Click(Index As Integer)
    'cDirs.Text = MnuDiItem(Index).Caption
    cDirs.ListIndex = Index
    tmMenus.Tag = "mnu_BDerecho"
    tmMenus.Enabled = True
End Sub

Private Sub MnuNewDir_Click()
    MnuAddTitulo.Tag = MnuNewDir.Tag
End Sub

Private Sub MnuSerNew_Click()
    MnuSerTitulo.Tag = MnuSerNew.Tag
End Sub

Private Sub MnuSerOpcion_Click(Index As Integer)
    MnuSerTitulo.Tag = Trim(MnuSerOpcion(Index).Tag)
End Sub

Private Sub tmMenus_Timer()

    tmMenus.Enabled = False
    Select Case LCase(tmMenus.Tag)
        Case LCase("mnu_BDerecho"): mnu_BDerecho
    End Select
    
End Sub

Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_Modalidad = m_def_Modalidad
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    
    m_Modalidad = PropBag.ReadProperty("Modalidad", m_def_Modalidad)
    
    UserControl.BackColor = m_BackColor
    lQDirs.ForeColor = m_ForeColor
    
    apply_PropModalidad
    
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    
    UserControl.Height = cDirs.Height
    
    lQDirs.Move UserControl.ScaleLeft, UserControl.ScaleTop + ((UserControl.Height - lQDirs.Height) / 2)
    cDirs.Move UserControl.ScaleLeft + lQDirs.Width, UserControl.ScaleTop, (UserControl.ScaleWidth - lQDirs.Width)
    
    lEMail.Move UserControl.ScaleLeft + lQDirs.Width, UserControl.ScaleTop

End Sub

Private Function fnc_ProcesoDireccion(mTextoCompleto As String)

    Dim mEvalNombre As String, mEvalDir As String
    
    zfn_PartoEmail mTextoCompleto, mEvalDir, mEvalNombre
    mEvalDir = LCase(mEvalDir)
    
    'Valido si la direccion ingresada es valida (@, .com, largo, etc)       ---------------------------------------------------------
    Dim sReason As String
    If Not IsEMailAddress(mEvalDir, sReason) Then
        MsgBox "La dirección de correo " & mEvalDir & " no es correcta." & vbCrLf & sReason, vbExclamation, "Dirección Incorrecta."
        Screen.MousePointer = 0: Exit Function
    End If
        
    'Validar si el nombre del servidor ya existe        -------------------------------------------------------------------------------
    Dim mStrServer As String, mStrDir As String, mServerID As Long
    
    mStrDir = Trim(Mid(mEvalDir, 1, InStr(mEvalDir, "@") - 1))
    mStrServer = Trim(Mid(mEvalDir, InStr(mEvalDir, "@") + 1))
    
    If Not fnc_ProcesoServidor(mStrServer, mServerID) Then Exit Function
    If mServerID = 0 Then Exit Function
    
    'Corrigo los valores a procesar con los datos del servidor ya OK    --> Dir Valida con ID de Servidor OK ------
    If mEvalNombre = "" Then
        If Trim(m_IdParaDirs) <> "" Then
            mEvalNombre = m_IdParaDirs
        Else
            If InStr(mStrServer, ".") <> 0 Then mEvalNombre = Mid(mStrServer, 1, InStr(mStrServer, ".") - 1) Else mEvalNombre = mStrServer
        End If
    End If
    
    mEvalDir = mStrDir & "@" & mStrServer
    mTextoCompleto = mEvalDir
    If Trim(mEvalNombre) <> "" Then mTextoCompleto = mTextoCompleto & " " & zfn_EMNombre(mEvalNombre, 1)
    cDirs.Text = mTextoCompleto
    
    'Veo si la direccion esta en el array   ---> Modifico el Nombre de la Dir (solamente)
    Dim mOldValueCombo As String, posArr As Integer
    posArr = arrIndex(mEvalDir)
    If posArr <> -1 Then
        With arrCorreo(posArr)
            mOldValueCombo = .ClaveCompleta & " " & zfn_EMNombre(.DireccionNombre, 1)
            
            .DireccionNombre = mEvalNombre
            .ServidorID = mServerID
            .ServidorNombre = mStrServer
        End With
        fnc_GrabarDireccion posArr
        
    Else
        mOldValueCombo = mTextoCompleto
    End If
    
    'Valido si la Direccion ya está ingresada en la lista
    idX = zfn_PosItemCombo(mOldValueCombo)
    If idX <> -1 Then
        cDirs.List(idX) = mTextoCompleto
        cDirs.ListIndex = idX
        zfn_SelIni
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------------------------------------------
                
    Dim bAddDir As Boolean
    bAddDir = False
    MnuAddTitulo.Tag = ""
    
    If cDirs.ListCount = 0 Then
        bAddDir = True
    Else
        '1) Remuevo los menu para recargar las dirs ...
        For idX = MnuCambiar.LBound To MnuCambiar.UBound
            If idX > MnuCambiar.LBound Then Unload MnuCambiar(idX)
        Next
        '2) Cargo las dirs en c/u de los menu ...       (copio el combo)
        For idX = 0 To cDirs.ListCount - 1
            If idX > MnuCambiar.UBound Then Load MnuCambiar(idX)
            
            MnuCambiar(idX).Caption = (idX + 1) & ") Sustituir a " & "“" & cDirs.List(idX) & "”"
            MnuCambiar(idX).Tag = cDirs.List(idX)
        Next
        
        MnuNewDir.Caption = (idX + 1) & ") ¿“" & mEvalDir & "”" & " es una Nueva Dirección?"
        MnuNewDir.Tag = "-1"
        
        PopupMenu MnuAdd, X:=cDirs.Left + 50, Y:=UserControl.ScaleTop + UserControl.ScaleHeight + 30, defaultmenu:=MnuAddTitulo
    End If
    
    Select Case Trim(MnuAddTitulo.Tag)
        Case "-1": bAddDir = True         'Seleccionó nueva dirección
            
        Case Is <> ""        'Selecciono sustituir una dir q ya existe
            Dim mSelValue As String
            zfn_PartoEmail Trim(MnuAddTitulo.Tag), mSelValue, ""
            posArr = arrIndex(mSelValue)
            If posArr <> -1 Then
                With arrCorreo(posArr)
                    .ClaveCompleta = mEvalDir
                    .DireccionNombre = mEvalNombre
                    .Direccion = mStrDir
                    .ServidorID = mServerID
                    .ServidorNombre = mStrServer
                End With
                
                fnc_GrabarDireccion posArr
                idX = zfn_PosItemCombo(Trim(MnuAddTitulo.Tag))
                cDirs.List(idX) = mTextoCompleto
                cDirs.ListIndex = idX
            End If
    End Select
    
    If bAddDir Then
        Dim exsDirNombre As String, exsDirID As Long
        If fnc_ExisteDireccionSinCliente(mStrDir, mServerID, exsDirNombre, exsDirID) Then
            mTextoCompleto = Replace(mTextoCompleto, "(" & mEvalNombre & ")", "(" & exsDirNombre & ")")
            mEvalNombre = exsDirNombre
        End If
        
        posArr = arrNewItem
        With arrCorreo(posArr)
            .ClaveCompleta = mEvalDir
            .DireccionID = IIf(exsDirID <> 0, exsDirID, -1)
            .Direccion = mStrDir
            .DireccionNombre = mEvalNombre
            
            .ServidorID = mServerID
            .ServidorNombre = mStrServer
            
            .DireccionID = fnc_GrabarDireccion(posArr)
        End With
        
        cDirs.AddItem mTextoCompleto
        cDirs.ItemData(cDirs.NewIndex) = arrCorreo(posArr).DireccionID
        cDirs.Text = mTextoCompleto
        'lQDirs.Caption = "(" & cDirs.ListCount & ")"
        fnc_CaptionQDirs
    End If
    
    cDirs.SelStart = 0
    
End Function

Private Function fnc_ProcesoServidor(mServerName As String, mServerID As Long) As Boolean
'   Si retorna True el servidor es válido.
'   Esta funcion retorna el Nombre y el Código de servidor  (Seleccionado o Agregado)

Dim bOK As Boolean
    
    fnc_ProcesoServidor = False
    
    mServerName = LCase(mServerName)
    '1)     Busco el Nombre del servidor exactamente como lo ingreso    -------------------
    mSQL = "Select * From EMailServer " & _
                " Where EMSDireccion = '" & mServerName & "'"
            
    Set rsAux = rdoCBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        mServerID = rsAux!EMSCodigo
        mServerName = Trim(rsAux!EMSDireccion)
        fnc_ProcesoServidor = True
        bOK = True
    End If
    rsAux.Close
    
    If bOK Then Exit Function
    
    'Remuevo los menu para recargar los servers ...
    MnuSerOpcion(0).Visible = True
    For idX = MnuSerOpcion.LBound To MnuSerOpcion.UBound
        If idX > MnuSerOpcion.LBound Then Unload MnuSerOpcion(idX)
    Next
        
    '2)     Consulto por sugenercias para mostrar menu  --------------------------------------
    '   Para las sugerecias al consultar pelar el punto uy, Orcomodines y Trabuque al nombre del servidor + % al final.
    Dim sLike As String
    Dim nAuxiliar As String, sTrab As String, arrTrab() As String
    
    sTrab = Trabuque(mServerName)
    arrTrab = Split(sTrab, ",")
        
    nAuxiliar = ""
    For idX = LBound(arrTrab) To UBound(arrTrab)
        nAuxiliar = nAuxiliar & "'" & Trim(arrTrab(idX)) & "'"
        If idX < UBound(arrTrab) Then nAuxiliar = nAuxiliar & ","
    Next
    
    sLike = Replace(mServerName, ".uy", "")
    sLike = OrtComodines(sLike, False) & "%"
    
    idX = 0
    
    mSQL = "Select * From EMailServer " & _
                " Where EMSDireccion like '" & sLike & "'" & _
                " OR EMSDireccion IN (" & nAuxiliar & ")"
    
    Set rsAux = rdoCBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        'Cargo los menues con opciones para nombres de servidor
        Do While Not rsAux.EOF
            If idX > 0 Then Load MnuSerOpcion(idX)
            With MnuSerOpcion(idX)
                .Tag = rsAux!EMSCodigo & "|" & Trim(rsAux!EMSDireccion)
                .Caption = idX + 1 & ") Es " & Trim(rsAux!EMSDireccion)
                .Visible = True
            End With
            
            idX = idX + 1
            rsAux.MoveNext
        Loop
    Else
        MnuSerOpcion(0).Visible = False
    End If
    rsAux.Close
    
    MnuSerNew.Caption = idX + 1 & ") ¿“" & mServerName & "”" & " es un Nuevo Servidor?"
    MnuSerNew.Tag = "-1"
    
    MnuSerTitulo.Tag = ""
    
    cDirs.BackColor = &H80&
    cDirs.ForeColor = vbWhite
    
    PopupMenu MnuServer, X:=cDirs.Left + 50, Y:=UserControl.ScaleTop + UserControl.ScaleHeight + 30, defaultmenu:=MnuSerTitulo

    cDirs.BackColor = vbWindowBackground
    cDirs.ForeColor = vbWindowText
    
    Select Case Trim(MnuSerTitulo.Tag)
        Case "-1"            'Ingresa Nuevo
                Set frmMaServer.Connect = rdoCBase
                frmMaServer.prmTxtNombre = ""
                frmMaServer.prmTxtHost = mServerName
                frmMaServer.prmIdServer = 0
                frmMaServer.Show vbModal, Me
                UserControl.Refresh
                If frmMaServer.prmTxtNombre <> "" And frmMaServer.prmIdServer <> 0 Then
                    mServerID = frmMaServer.prmIdServer
                    mServerName = frmMaServer.prmTxtNombre
                    fnc_ProcesoServidor = True
                End If
                        
        Case Is <> ""        'Selecciono uno que ya existe
                arrTrab = Split(MnuSerTitulo.Tag, "|")
                mServerID = arrTrab(0)
                mServerName = Trim(arrTrab(1))
                fnc_ProcesoServidor = True
    End Select
    
                    
End Function

Private Function fnc_ExisteDireccionSinCliente(findDireccion As String, findIdServer As Long, retDirNombre As String, retDirID) As Boolean

    On Error GoTo errValidoD
    Screen.MousePointer = 11
    fnc_ExisteDireccionSinCliente = False
    
    mSQL = "Select * From EMailDireccion" & _
               " Where EMDDireccion = '" & Trim(findDireccion) & "'" & _
               " And EMDServidor = " & findIdServer & _
               " And EMDIdCliente Is Null" & _
               " And (EMDEliminado is Null or EMDEliminado = 0)"
               
    Set rsAux = rdoCBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        retDirNombre = Trim(rsAux!EMDNombre)
        retDirID = Trim(rsAux!EMDCodigo)
        fnc_ExisteDireccionSinCliente = True
    End If
    rsAux.Close
    
    Screen.MousePointer = 0
    Exit Function
     
errValidoD:
    Screen.MousePointer = 0
End Function

Private Function fnc_CargoMails() As Boolean

    On Error GoTo errNT
    fnc_CargoMails = False
    cDirs.Clear
    
    mSQL = "Select EMDCodigo, EMDDireccion, EMSDireccion, EMSCodigo, EMDNombre" & _
                " From EMailDireccion, EMailServer " & _
                " Where EMDServidor = EMSCodigo " & _
                " And EMDIDCliente = " & prmIDCliente & _
                " And (EMDEliminado Is Null or EMDEliminado = 0)"
    
    Set rsAux = rdoCBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)

    Do While Not rsAux.EOF
        
        idX = arrNewItem
        With arrCorreo(idX)
            .ClaveCompleta = Trim(rsAux!EMDDireccion) & "@" & Trim(rsAux!EMSDireccion)
            
            .DireccionID = rsAux!EMDCodigo
            .Direccion = Trim(rsAux!EMDDireccion)
            .DireccionNombre = Trim(rsAux!EMDNombre)
            
            .ServidorID = rsAux!EMSCodigo
            .ServidorNombre = Trim(rsAux!EMSDireccion)
        
            cDirs.AddItem .ClaveCompleta & " " & zfn_EMNombre(.DireccionNombre, 1)
            cDirs.ItemData(cDirs.NewIndex) = .DireccionID
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    If cDirs.ListCount > 0 Then
        'lQDirs.Caption = "(" & cDirs.ListCount & ")"
        fnc_CaptionQDirs
        cDirs.ListIndex = 0
    End If
    
    Select Case cDirs.ListCount
        Case 0: lEMail.ForeColor = &H808080
        Case Else: lEMail.ForeColor = vbRed
    End Select
    fnc_CargoMails = True
    Exit Function
errNT:
    clsGeneral.OcurrioError "Error al cargar las direcciones de correo.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function zfn_PosItemCombo(findTexto) As Integer

    zfn_PosItemCombo = -1
    For idX = 0 To cDirs.ListCount - 1
        If Trim(LCase(findTexto)) = Trim(LCase(cDirs.List(idX))) Then
            zfn_PosItemCombo = idX
            Exit For
        End If
    Next
    
End Function

Private Function fnc_GrabarDireccion(idxArray As Integer) As Long
On Error GoTo errSave
Dim mIDDireccion As Long
Dim retIDDireccion As Long

    fnc_GrabarDireccion = -1
    mIDDireccion = arrCorreo(idxArray).DireccionID
    If mIDDireccion = -1 Then mIDDireccion = 0
    
    mSQL = "Select * from EMailDireccion Where EMDCodigo = " & mIDDireccion
    Set rsAux = rdoCBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    
    If mIDDireccion = 0 Then
        'retIDDireccion = Autonumerico(taba_EMailDireccion, rdoCBase)
        rsAux.AddNew
        rsAux!EMDFAlta = Format(Now, "mm/dd/yyyy hh:mm:ss")
        rsAux!EMDUsuAlta = m_IDUsuario 'paCodigoDeUsuario
        'rsAux!EMDCodigo = retIDDireccion
    Else
        retIDDireccion = mIDDireccion
        rsAux.Edit
    End If

    With arrCorreo(idxArray)
        rsAux!EMDNombre = Trim(.DireccionNombre)
        rsAux!EMDDireccion = Trim(.Direccion)
        rsAux!EMDServidor = .ServidorID
        If prmIDCliente <> 0 Then rsAux!EMDIdCliente = prmIDCliente 'Else rsAux!EMDIdCliente = Null
        
        'If bNoEnviar.Value = vbChecked Then rsAux!EMDNoEnviarInf = 1 Else rsAux!EMDNoEnviarInf = 0
        rsAux!EMDNoEnviarInf = 0
    End With
    
    rsAux.Update
    rsAux.Close
    
    If retIDDireccion = 0 Then
        mSQL = "Select * from EMailDireccion " & _
                    " Where EMDNombre = '" & Trim(arrCorreo(idxArray).DireccionNombre) & "'" & _
                    " And EMDDireccion = '" & Trim(arrCorreo(idxArray).Direccion) & "'" & _
                    " And EMDServidor = " & arrCorreo(idxArray).ServidorID
                    
        Set rsAux = rdoCBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then retIDDireccion = rsAux!EMDCodigo
        rsAux.Close
        
    End If
    
    fnc_GrabarDireccion = retIDDireccion
    Exit Function
    
errSave:
    clsGeneral.OcurrioError "Error al grabar la direccion de correo.", Err.Description
End Function

 Private Function zfn_EMNombre(mTexto As String, QueHacer) As String
 
    mTexto = Trim(mTexto)
    zfn_EMNombre = mTexto
    Select Case QueHacer
        Case 1: zfn_EMNombre = "(" & mTexto & ")"
        Case 2: zfn_EMNombre = Mid(mTexto, 1, Len(mTexto) - 2)
    End Select
    
 End Function

Private Function zfn_PartoEmail(mDirCompleta As String, retEMail As String, retNombre As String)
Const SEP1 = " ("
Const SEP2 = ")"

    mDirCompleta = Trim(mDirCompleta)
    If InStr(mDirCompleta, SEP1) <> 0 And Right(mDirCompleta, Len(SEP2)) = SEP2 Then
        Dim mPosSep1 As Integer
        mPosSep1 = InStr(mDirCompleta, SEP1)
        retEMail = Mid(mDirCompleta, 1, mPosSep1 - 1)
        retNombre = Mid(mDirCompleta, mPosSep1 + Len(SEP1), Len(mDirCompleta) - mPosSep1 - Len(SEP1) - Len(SEP2) + 1)
        
    Else
        retEMail = mDirCompleta
        retNombre = ""
    End If
    
End Function

Public Property Get IdsPorDefecto() As String
    IdsPorDefecto = m_IdParaDirs
End Property

Public Property Let IdsPorDefecto(ByVal New_Value As String)
    m_IdParaDirs = New_Value
    PropertyChanged "IdsPorDefecto"
End Property

Public Property Get IDUsuario() As Long
    IDUsuario = m_IDUsuario
End Property

Public Property Let IDUsuario(ByVal New_Value As Long)
    m_IDUsuario = New_Value
    PropertyChanged "IDUsuario"
End Property

Public Property Get Modalidad() As Integer
    Modalidad = m_Modalidad
End Property

Public Property Let Modalidad(ByVal New_Value As Integer)
    m_Modalidad = New_Value
    PropertyChanged "Modalidad"
    
    apply_PropModalidad
    
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Value As Boolean)
    m_Enabled = New_Value
    
    If m_Enabled Then
        cDirs.BackColor = vbWindowBackground
        cDirs.ForeColor = vbWindowText
    Else
        cDirs.BackColor = &HE0E0E0      'vbButtonFace
        cDirs.ForeColor = vbWindowText
    End If
    cDirs.Enabled = m_Enabled
    
    PropertyChanged "Enabled"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    lQDirs.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Function mnu_BDerecho()

    On Error GoTo errMnuEmail
    
    Dim mDirEmail As String
    zfn_PartoEmail cDirs.Text, mDirEmail, ""
    idX = arrIndex(mDirEmail)
    If idX = -1 Then Exit Function
    If Not (arrCorreo(idX).DireccionID > 0) Then Exit Function
    
    MnuBDTitulo.Caption = cDirs.Text
    MnuBDTitulo.Tag = arrCorreo(idX).DireccionID
    MnuBDTitulo.Visible = (m_Modalidad = 2)
    MnuBDL1.Visible = MnuBDTitulo.Visible
    
    Dim I As Integer
            
    MnuBDLista(0).Visible = True
    MnuBDL3.Visible = False
    For I = 1 To MnuBDLista.UBound
        Unload MnuBDLista(I)
    Next
        
    mSQL = "Select * from ListaDistribucion left Outer Join EMailLista On LiDCodigo = EMLLista And EMLMail = " & Val(MnuBDTitulo.Tag) & _
                " Where LiDHabilitado = 1" & _
                " Order by LiDNombre"
    Set rsAux = rdoCBase.OpenResultset(mSQL, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsAux.EOF
        I = MnuBDLista.UBound + 1
        Load MnuBDLista(I)
        With MnuBDLista(I)
            .Caption = Trim(rsAux!LiDNombre)
            If Not IsNull(rsAux!LiDExcluye) Then If rsAux!LiDExcluye = 1 Then .Caption = "NO " & Trim(.Caption)
            .Tag = rsAux!LiDCodigo
            If Not IsNull(rsAux!EMLMail) Then .Checked = True Else .Checked = False
            .Visible = True
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close

    If MnuBDLista.UBound > 0 Then
        MnuBDLista(0).Visible = False
        MnuBDL3.Visible = True
    End If
    
    If Xbd <> 0 Or Ybd <> 0 Then
        If MnuBDTitulo.Visible Then
            PopupMenu MnuBDerecho, X:=Xbd, Y:=Ybd, defaultmenu:=MnuBDTitulo
        Else
            PopupMenu MnuBDerecho, X:=Xbd, Y:=Ybd
        End If
    Else
        If MnuBDTitulo.Visible Then
            PopupMenu MnuBDerecho, X:=cDirs.Left + 50, Y:=UserControl.ScaleTop + UserControl.ScaleHeight + 30, defaultmenu:=MnuBDTitulo
        Else
            PopupMenu MnuBDerecho, X:=cDirs.Left + 50, Y:=UserControl.ScaleTop + UserControl.ScaleHeight + 30
        End If
    End If
    Exit Function
errMnuEmail:
End Function

Public Function AbrirMenuDirecciones()
Dim iT As Integer
    
    If Modalidad = 1 Then Exit Function
    If cDirs.ListCount = 0 Then Exit Function
    
    If cDirs.ListCount = 1 Then
        cDirs.Text = cDirs.List(0)
        cDirs.ListIndex = 0
        mnu_BDerecho
    End If
    
    If cDirs.ListCount > 1 Then
        For iT = 1 To MnuDiItem.UBound
            Unload MnuDiItem(iT)
        Next
    
        For iT = LBound(arrCorreo) To UBound(arrCorreo)
            If iT > 0 Then Load MnuDiItem(iT)
            With MnuDiItem(iT)
                .Caption = arrCorreo(iT).ClaveCompleta
            End With
        Next
        If Xbd <> 0 Or Ybd <> 0 Then
            PopupMenu MnuDirs, X:=Xbd, Y:=Ybd, defaultmenu:=MnuDiTitulo
        Else
            PopupMenu MnuDirs, X:=cDirs.Left + 50, Y:=UserControl.ScaleTop + UserControl.ScaleHeight + 30, defaultmenu:=MnuDiTitulo
        End If
    End If

End Function

Private Function mnu_Direcciones()
'Para mantener el mismo código hago 1 método público.
    mnu_Direcciones = AbrirMenuDirecciones
End Function

Private Function zfn_SelIni()
On Error Resume Next
    cDirs.SelStart = Len(cDirs.Text)
    SendKeys "+{HOME}", True
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    
    Call PropBag.WriteProperty("Modalidad", m_Modalidad, m_def_Modalidad)
End Sub

Private Function apply_PropModalidad()
On Error GoTo errApply

    cDirs.Visible = (m_Modalidad = 1)
    lEMail.Visible = (m_Modalidad = 2)

    If m_Modalidad = 2 Then
        UserControl.Width = lEMail.Left + lEMail.Width + 20
    Else
        UserControl.Width = cDirs.Left + cDirs.Width + 20
    End If
    
errApply:
End Function

Private Function fnc_CaptionQDirs()
On Error Resume Next
    
    Select Case cDirs.ListCount
        Case 0: lQDirs.Caption = ""
        Case 1
            If m_Modalidad = 1 Then lQDirs.Caption = "(" & cDirs.ListCount & ")" Else lQDirs.Caption = ""
        
        Case Else
            lQDirs.Caption = "(" & cDirs.ListCount & ")"
    End Select
    
End Function
