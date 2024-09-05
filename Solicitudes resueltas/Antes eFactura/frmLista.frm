VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLista 
   Caption         =   "Solicitudes Resueltas"
   ClientHeight    =   4005
   ClientLeft      =   1800
   ClientTop       =   7020
   ClientWidth     =   8670
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmLista.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4005
   ScaleWidth      =   8670
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   300
      Top             =   0
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   2415
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   4260
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
      BackColor       =   12640511
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   12640511
      BackColorAlternate=   12640511
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
   Begin MSWinsockLib.Winsock wsSocket 
      Left            =   7200
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   5880
      TabIndex        =   1
      Top             =   3540
      Width           =   1635
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLista.frx":0442
            Key             =   "Si"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLista.frx":075C
            Key             =   "No"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLista.frx":0A76
            Key             =   "Condicional"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLista.frx":0D90
            Key             =   "llamara"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuBD 
      Caption         =   "MnuBD"
      Visible         =   0   'False
      Begin VB.Menu MnuDevolver 
         Caption         =   "Devolver Solicitud"
         Index           =   0
      End
      Begin VB.Menu MnuDevolver 
         Caption         =   "Dejar sin efecto"
         Index           =   1
      End
      Begin VB.Menu MnuL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrintCnfgConformes 
         Caption         =   "¿Dónde imprimo conformes?"
      End
      Begin VB.Menu MnuPrintConfig 
         Caption         =   "Configurar Impresoras"
      End
      Begin VB.Menu MnuPrintLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrintOpt 
         Caption         =   "MnuPrintOpt"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSol As rdoResultset
Dim mIDSel As Long                  'Codigo del Item Seleccionado

Dim mValor As Long
Dim globalData As String

Dim arrTramas() As String

Private Function ValidarVersionEFactura() As Boolean
On Error GoTo errEC
    With New clsCGSAEFactura
        ValidarVersionEFactura = .ValidarVersion()
    End With
    Exit Function
errEC:

End Function

Private Sub Form_Activate()
   On Error Resume Next
   vsLista.SetFocus
End Sub

Private Sub Form_Load()

    On Error Resume Next
    ReDim arrTramas(0)
    
    
    If Not ValidarVersionEFactura() Then
        MsgBox "La versión del componente CGSAEFactura está desactualizado, debe distribuir software." _
                    & vbCrLf & vbCrLf & "Se cancelará la ejecución.", vbCritical, "EFactura"
        End
    End If
    
    ObtengoSeteoForm Me
    InicializoGrilla
    FechaDelServidor
        
    QuerySolicitudes
    
    Timer1.Enabled = Not ws_IniciarConexion
    Me.BackColor = vsLista.BackColor
    
    zfn_LoadMenuOpcionPrint
    oCnfgPrint.CargarConfiguracion cnfgAppNombreConformes, cnfgKeyTicketConformes
    
    Screen.MousePointer = 0
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    vsLista.Top = Me.ScaleTop: vsLista.Left = Me.ScaleLeft
    vsLista.Width = Me.ScaleWidth
    
    If lStatus.Visible Then
        vsLista.Height = Me.ScaleHeight - (vsLista.Top * 2) - lStatus.Height
    Else
        vsLista.Height = Me.ScaleHeight - (vsLista.Top * 2)
    End If
    
    lStatus.Top = vsLista.Height
    lStatus.Left = Me.ScaleWidth - lStatus.Width - 60
    
    Refresh
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    GuardoSeteoForm Me
    
    wsSocket.Close
    
    EndMain
    
End Sub


Private Sub MnuDevolver_Click(Index As Integer)

    If Val(MnuBD.Tag) <> 0 Then
        Dim idEstado As Integer
        If Index = 0 Then idEstado = EstadoSolicitud.ParaRetomar
        If Index = 1 Then idEstado = EstadoSolicitud.SinEfecto
        
        If fnc_DevolverSolicitud(MnuBD.Tag, idEstado) Then
            On Error Resume Next
            vsLista.RemoveItem itm_FindItem(Val(MnuBD.Tag))
        End If
        
    End If
End Sub

Private Sub MnuPrintCnfgConformes_Click()

    frmDondeImprimo.Show vbModal
    oCnfgPrint.CargarConfiguracion App.title, cnfgKeyTicketConformes

End Sub

Private Sub vsLista_DblClick()
On Error GoTo errFnc

    If vsLista.Rows > 1 Then
        If vsLista.Cell(flexcpData, vsLista.Row, 0) = -1 Then Exit Sub        'Es un llamar a
        
        If FormActivo("EsCondicional") Then frmCredito.SetFocus: Exit Sub
        
        mIDSel = vsLista.Cell(flexcpText, vsLista.Row, 0)
        
        If vsLista.Cell(flexcpData, vsLista.Row, 0) <> EstadoSolicitud.Rechazada Then
            Select Case fnc_BloqueoSolicitud(mIDSel) 'Si es 1 esta todo OK--------------------------
                Case 0  'OTRO USUARIO
                    Screen.MousePointer = 0
                    MsgBox "La solicitud se está facturando por otro usuario. No podrá visualizarla.", vbExclamation, "Datos Modificados"
                    Exit Sub
               
                Case -1 'ERROR o FUE RESUELTA
                    Screen.MousePointer = 0
                    MsgBox "Posiblemente la solicitud ya fue facturada.", vbExclamation, "Datos Modificados"
                    QuerySolicitudes
                    Exit Sub
            End Select  '----------------------------------------------------------------------------------------
        End If
        
        Screen.MousePointer = 11
        
        frmCredito.prmIDSolicitud = mIDSel
        frmCredito.Show vbModal, Me
    End If
    
errFnc:
End Sub

Private Sub QuerySolicitudes()

    On Error GoTo ErrLoad
    
    Dim mIdRow As Long
    mIdRow = 1
    
    If vsLista.Rows > vsLista.FixedRows Then mIdRow = vsLista.Row
    vsLista.Rows = vsLista.FixedRows
    
    Dim mSolUsuarioR As String, mSolComentarioR As String, mSolUsuario As String
    Dim mSolNombre As String, mSolTipo As String, mSolFResolucion As String
    
    Screen.MousePointer = 11
    'Solicitudes Poceso de Resolucion = Manual o Automática
    '                Estado <> Pendiente
    '                Sucursal = a la de la maquina
    '                Que estén Visibles
                     
                     
    Cons = "Select Solicitud.*, ResComentario, UsuIdentificacion, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
           & " From Solicitud, SolicitudResolucion, CPersona, Usuario " _
           & " Where SolFecha Between '" & Format(gFechaServidor, "mm/dd/yyyy 00:00") & "' And  '" & Format(gFechaServidor, "mm/dd/yyyy 23:59") & "'" _
           & " And SolSucursal = " & paCodigoDeSucursal _
           & " And SolProceso Not In ( " & TipoResolucionSolicitud.Facturada & "," & TipoResolucionSolicitud.Facturando & ")" _
           & " And SolEstado IN ( " & EstadoSolicitud.Aprovada & ", " & EstadoSolicitud.Rechazada & ", " & EstadoSolicitud.Condicional & ")" _
           & " And CPeCliente = SolCliente" _
           & " And SolUsuarioS = UsuCodigo" _
           & " And SolVisible Is NULL " _
           & " And SolCodigo = ResSolicitud " _
           & " And ResNumero = (Select MAX(ResNumero) From SolicitudResolucion Where SolCodigo = ResSolicitud)" _
                                    & " UNION " _
           & "Select Solicitud.*, ResComentario, UsuIdentificacion, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
           & " From Solicitud, SolicitudResolucion, CEmpresa, Usuario " _
           & " Where SolFecha Between '" & Format(gFechaServidor, "mm/dd/yyyy 00:00") & "' And  '" & Format(gFechaServidor, "mm/dd/yyyy 23:59") & "'" _
           & " And SolSucursal = " & paCodigoDeSucursal _
           & " And SolProceso Not In ( " & TipoResolucionSolicitud.Facturada & "," & TipoResolucionSolicitud.Facturando & ")" _
           & " And SolEstado IN ( " & EstadoSolicitud.Aprovada & ", " & EstadoSolicitud.Rechazada & ", " & EstadoSolicitud.Condicional & ")" _
           & " And CEmCliente = SolCliente" _
           & " And SolUsuarioS = UsuCodigo" _
           & " And SolVisible Is NULL " _
           & " And SolCodigo = ResSolicitud " _
           & " And ResNumero = (Select MAX(ResNumero) From SolicitudResolucion Where SolCodigo = ResSolicitud)" _
           & " Order by SolFResolucion"

    Set rsSol = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsSol.EOF
    
        mSolUsuarioR = "": mSolComentarioR = "": mSolUsuario = "": mSolNombre = "": mSolTipo = "": mSolFResolucion = ""
        
        If Not IsNull(rsSol!SolUsuarioR) Then mSolUsuarioR = z_BuscoUsuario(rsSol!SolUsuarioR, True)
        If Not IsNull(rsSol!ResComentario) Then mSolComentarioR = Trim(rsSol!ResComentario)
        If Not IsNull(rsSol!UsuIdentificacion) Then mSolUsuario = Trim(rsSol!UsuIdentificacion)
        
        If Not IsNull(rsSol!Nombre) Then mSolNombre = Trim(rsSol!Nombre)
        If Not IsNull(rsSol!SolTipo) Then mSolTipo = Trim(rsSol!SolTipo)
        If Not IsNull(rsSol!SolFResolucion) Then mSolFResolucion = Trim(rsSol!SolFResolucion)
            
        itm_AddSolicitud rsSol!SolCodigo, rsSol!SolEstado, rsSol!SolProceso, mSolUsuarioR, mSolNombre, mSolUsuario, _
                                mSolFResolucion, mSolTipo, mSolComentarioR
            
        rsSol.MoveNext
    Loop
    rsSol.Close
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al cargar los datos en la lista. Solicitud=" & rsSol!SolCodigo, Err.Description
    Screen.MousePointer = 0
End Sub


Private Function itm_AddSolicitud(mCodigo As Long, mIDEstado As String, mIDProceso As String, mUsuarioR As String, mCliente As String, _
                                                mUsuario As String, mFResolucion As String, mIDTipoS As String, mComentarioR As String) As Boolean

    If Val(mIDEstado) = EstadoSolicitud.ParaRetomar Or Val(mIDEstado) = EstadoSolicitud.Pendiente Then
        'Si esta la borro ---> pasa por las solicitudes pendientes
         On Error Resume Next
        vsLista.RemoveItem itm_FindItem(mCodigo)
        Exit Function
    End If
    
Dim mIdRow As Integer

    itm_AddSolicitud = True
    mIdRow = itm_FindItem(mCodigo)
    
    With vsLista
        If mIdRow = -1 Then
            .AddItem ""
            mIdRow = .Rows - 1
        End If
            
        .Cell(flexcpText, mIdRow, 0) = mCodigo
                
        Select Case Val(mIDEstado)
            Case EstadoSolicitud.Aprovada: .Cell(flexcpPicture, mIdRow, 0) = ImageList1.ListImages("Si").ExtractIcon
            Case EstadoSolicitud.Rechazada: .Cell(flexcpPicture, mIdRow, 0) = ImageList1.ListImages("No").ExtractIcon
            Case EstadoSolicitud.Condicional: .Cell(flexcpPicture, mIdRow, 0) = ImageList1.ListImages("Condicional").ExtractIcon
        End Select
        
        mValor = Val(mIDEstado): .Cell(flexcpData, mIdRow, 0) = mValor
             
        If Val(mIDProceso) = TipoResolucionSolicitud.LlamarA Then
            .Cell(flexcpPicture, mIdRow, 0) = ImageList1.ListImages("llamara").ExtractIcon
            .Cell(flexcpText, mIdRow, 5) = "HABLAR CON "
            
            If Trim(mUsuarioR) = "" Then .Cell(flexcpText, mIdRow, 5) = .Cell(flexcpText, mIdRow, 5) & UCase(mUsuarioR)
            .Cell(flexcpData, mIdRow, 0) = -1
        End If
    
        .Cell(flexcpText, mIdRow, 1) = Format(mFResolucion, "hh:mm")
        .Cell(flexcpText, mIdRow, 2) = Trim(mCliente)
        
        Select Case Val(mIDTipoS)
            Case TipoSolicitud.AlMostrador: .Cell(flexcpText, mIdRow, 3) = "Mos."
            Case TipoSolicitud.Reserva: .Cell(flexcpText, mIdRow, 3) = "Tel."
            Case TipoSolicitud.Servicio: .Cell(flexcpText, mIdRow, 3) = "Ser."
        End Select
        
        If Trim(.Cell(flexcpText, mIdRow, 5)) = "" Then If Trim(mComentarioR) <> "" Then .Cell(flexcpText, mIdRow, 5) = Trim(mComentarioR)
        If Trim(mUsuario) <> "" Then .Cell(flexcpText, mIdRow, 4) = Trim(mUsuario)
    End With
            
End Function

Private Function itm_FindItem(mFID As Long) As Integer

    Dim iFnd As Integer
    itm_FindItem = -1
    
    With vsLista
        For iFnd = .FixedRows To .Rows - 1
            If .Cell(flexcpValue, iFnd, 0) = mFID Then
                itm_FindItem = iFnd
                Exit For
            End If
        Next
    End With
    
End Function

Private Sub vsLista_KeyDown(KeyCode As Integer, Shift As Integer)

    If vsLista.Rows = 1 Then Exit Sub

Dim CodigoS As Long

    If KeyCode = vbKeyDelete Then
        On Error GoTo errEliminar
        CodigoS = vsLista.Cell(flexcpText, vsLista.Row, 0)
        If MsgBox("Confirma marcar la solicitud como oculta.", vbQuestion + vbYesNo + vbDefaultButton2, "Ocultar Solicitud") = vbNo Then Exit Sub
        
        Screen.MousePointer = 11
        Cons = "Select * from Solicitud Where SolCodigo = " & CodigoS
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If RsAux!SolProceso <> TipoResolucionSolicitud.Facturada _
                And RsAux!SolProceso <> TipoResolucionSolicitud.Facturando Then
                RsAux.Edit
                RsAux!SolVisible = "N"
                RsAux.Update
                
                vsLista.RemoveItem vsLista.Row
            Else
                Screen.MousePointer = 0
                MsgBox "La solicitud se está facturando o fue facturada por otro usuario.", vbExclamation, "Datos Modificados"
            End If
        Else
            Screen.MousePointer = 0
            MsgBox "La solicitud seleccionada no existe o ha sido eliminada.", vbCritical, "Datos Modificados"
        End If
        RsAux.Close
        Screen.MousePointer = 0
    End If
    Exit Sub

errEliminar:
    clsGeneral.OcurrioError "Error al marcar la solicitud como oculta.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub vsLista_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me: Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn And vsLista.Rows > 1 Then Call vsLista_DblClick
    
End Sub

Private Sub Timer1_Timer()

    Me.Caption = Trim(Me.Caption) & " (Actualizando...)"
    QuerySolicitudes
    Me.Caption = "Solicitudes Resueltas"
    
    Timer1.Enabled = Not ws_Reconectar
    
End Sub

Private Sub InicializoGrilla()

    With vsLista
        .Rows = 1: .Cols = 1
        .FormatString = "<Código|Resuelta|<Cliente|<Tipo|Usuario|Comentarios"
        .ColWidth(0) = 900: .ColWidth(2) = 3600: .ColWidth(3) = 800: .ColWidth(4) = 1200
        .WordWrap = False
        .MergeCells = flexMergeSpill: .ExtendLastCol = True
    End With
    
End Sub

Private Sub vsLista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
    
        Dim mIDSel As Long: mIDSel = 0
        If vsLista.Rows > vsLista.FixedRows Then
            mIDSel = vsLista.Cell(flexcpText, vsLista.Row, 0)
        End If
    
        MnuBD.Tag = mIDSel
        MnuDevolver(0).Caption = "Devolver Sol: " & IIf(mIDSel <> 0, CStr(mIDSel), "...")
        MnuDevolver(0).Enabled = (mIDSel <> 0)
        
        MnuDevolver(1).Caption = "Dejar sin efecto ... Sol: " & IIf(mIDSel <> 0, CStr(mIDSel), "...")
        MnuDevolver(1).Enabled = (mIDSel <> 0)
        
        PopupMenu MnuBD
    End If
    
End Sub

Private Sub wsSocket_Close()
    On Error Resume Next
    lStatus.Visible = True
    lStatus.Caption = "Desconectado": lStatus.Refresh
    Timer1.Enabled = True
    Call Form_Resize
'    prmCurrentUser = -1
End Sub

Private Sub wsSocket_DataArrival(ByVal bytesTotal As Long)
    
    Dim strDato As String
    Dim xPos As Long
    
    If wsSocket.BytesReceived > 0 Then
        
        wsSocket.GetData strDato
        globalData = globalData & strDato
    
        Do While InStr(globalData, sc_FIN) <> 0
            xPos = InStr(globalData, sc_FIN)
            
            strDato = Mid(globalData, 1, xPos - 1)
            globalData = Mid(globalData, xPos + Len(sc_FIN))
            
            Select Case UCase(strDato)
                Case "START":
                Case "END": ws_ProcesoDataArraival
                Case Else
'                    lTrama.Caption = strDato
                    Dim mIDT As Integer
                    If Trim(arrTramas(0)) = "" Then mIDT = 0 Else mIDT = UBound(arrTramas) + 1
                    ReDim Preserve arrTramas(mIDT)
                    arrTramas(mIDT) = strDato
            End Select
            
        Loop
        
    End If
    
End Sub

Private Function ws_IniciarConexion() As Boolean

    ws_IniciarConexion = True
    lStatus.Caption = "Conectando ...": lStatus.Refresh
    
    wsSocket.Connect prmIPServer, prmPortServer
    
    Dim aQIntentos As Integer
    aQIntentos = 1
    
    Do While aQIntentos <= 4
        DoEvents
        If wsSocket.state = 7 Then Exit Do
        Sleep 500
        aQIntentos = aQIntentos + 1
        
        lStatus.Caption = "Conectando ... (" & aQIntentos & ")"
        lStatus.Refresh
    Loop
    
    If wsSocket.state <> 7 Then
        lStatus.Caption = "Sin Conexión"
        ws_IniciarConexion = False
'        tmCheckLog.Enabled = False
        lStatus.Visible = True
    Else
        lStatus.Caption = "Conectado"
'        prmCurrentUser = -1
        ws_SendWhoIam
'        tmCheckLog.Enabled = True
        lStatus.Visible = False
    End If
    
    Call Form_Resize
    lStatus.Refresh
    
End Function

Private Function ws_SendWhoIam()
        
    'IDUsr|NombreUsr|0000 (Info Sol Ser Gas Sus SRe)
    'ATENCION: En la facturacion mando en ID de Sucursal en vez del id de Usuario !!
    Dim mDatas As String
    
    mDatas = Trim(miConexion.NombreTerminal)
    mDatas = paCodigoDeSucursal & "|" & mDatas & "|00001"
    ws_SendData mDatas
    
End Function

Public Function ws_Reconectar() As Boolean
    On Error Resume Next
    wsSocket.Close
    
    ws_Reconectar = ws_IniciarConexion
    
End Function

Public Sub ws_SendData(Trama As String)

    On Error Resume Next
    If wsSocket.state = 7 Then
        wsSocket.SendData Trama
        DoEvents
    End If

End Sub

Private Function ws_ProcesoDataArraival()
    
    If Trim(arrTramas(0)) = "" Then Exit Function
    
Dim arrValue() As String
Dim mRow As Integer, mRowSel As Integer

    On Error GoTo errCargar
    Screen.MousePointer = 11
        
    mRowSel = 1
    If vsLista.Rows > vsLista.FixedRows Then mRowSel = vsLista.Row
    
    '.IDUserPara = UsuarioPara
    '.IdTipo = TipoT
    '.IDEstadoTrama = Trim(EstadoT)
    '.DatosTrama = DatosT
    
    Dim iT As Integer
    
    For iT = LBound(arrTramas) To UBound(arrTramas)
'        Debug.Print arrTramas(iT)
        arrValue = Split(arrTramas(iT), "|")
    
        Select Case Trim(arrValue(2))
            Case "E"        'Eliminado          -------------------------------------------------------
                    mRow = itm_FindItem(Val(arrValue(3)))
                    If mRow <> -1 Then vsLista.RemoveItem mRow
                    
            Case "N"        'Nuevo o Modificada     ------------------------------------------------
                    Select Case Val(arrValue(1))
                        Case Asuntos.SolicitudesResueltas
                            'Codigo|Estado|Proceso||NameUsrR|Cliente|NameUsrS|FResolucion|Tipo|ComentarioR
                            '0|5|N|3297|1|2|Adrian|OCCHIUZZI MARTINEZ, WALTER|Monica|30/08/2002 17:51:32|1|Si0|5|E|3297
                            itm_AddSolicitud Val(arrValue(3)), arrValue(4), arrValue(5), arrValue(6), arrValue(7), arrValue(8), arrValue(9), arrValue(10), arrValue(11)
                            
                    End Select
        End Select
        
        If vsLista.Rows > vsLista.FixedRows Then
            vsLista.Select vsLista.FixedRows, 0 'vsLista.Cols - 1
            vsLista.Sort = flexSortGenericAscending
        End If
        
    Next
    
    ReDim arrTramas(0)
    
    If vsLista.Rows > vsLista.FixedRows Then
        If Not (vsLista.Rows > mRowSel) Then mRowSel = vsLista.Rows - 1
        vsLista.Select mRowSel, 2
    End If
    
    Screen.MousePointer = 0
    Exit Function
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos en la lista.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub MnuPrintConfig_Click()
On Error Resume Next
    
    prj_LoadConfigPrint True
    
    Dim iQ As Integer
    For iQ = MnuPrintOpt.LBound To MnuPrintOpt.UBound
        MnuPrintOpt(iQ).Checked = False
        MnuPrintOpt(iQ).Checked = (MnuPrintOpt(iQ).Caption = paOptPrintSel)
    Next
    
End Sub

Private Sub MnuPrintOpt_Click(Index As Integer)
On Error GoTo errLCP
Dim objPrint As New clsCnfgPrintDocument
Dim sPrint As String
Dim vPrint() As String
Dim iQ As Integer
    
    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        
        If .ChangeConfigPorOpcion(MnuPrintOpt(Index).Caption) Then
            For iQ = MnuPrintOpt.LBound To MnuPrintOpt.UBound
                MnuPrintOpt(iQ).Checked = False
            Next
            MnuPrintOpt(Index).Checked = True
        End If

    End With
    Set objPrint = Nothing
    
    On Error Resume Next
    prj_LoadConfigPrint False
    
    Exit Sub
errLCP:
    MsgBox "Error al setear los datos de configuración: " & Err.Description, vbExclamation, "ATENCIÓN"
End Sub

Private Sub zfn_LoadMenuOpcionPrint()
On Error Resume Next

Dim vOpt() As String
Dim iQ As Integer
    
    MnuPrintLine1.Visible = (paOptPrintList <> "")
    MnuPrintOpt(0).Visible = (paOptPrintList <> "")
    
    If paOptPrintList = "" Then
        Exit Sub
    ElseIf InStr(1, paOptPrintList, "|", vbTextCompare) = 0 Then
        MnuPrintOpt(0).Caption = paOptPrintList
    Else
        vOpt = Split(paOptPrintList, "|")
        For iQ = 0 To UBound(vOpt)
            If iQ > 0 Then Load MnuPrintOpt(iQ)
            With MnuPrintOpt(iQ)
                .Caption = Trim(vOpt(iQ))
                .Checked = (LCase(.Caption) = LCase(paOptPrintSel))
                .Visible = True
            End With
        Next
    End If
    
End Sub


