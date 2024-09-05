VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form DeSolicitud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Solicitudes"
   ClientHeight    =   4785
   ClientLeft      =   2775
   ClientTop       =   2625
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DeSolicitud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8985
   Begin MSMask.MaskEdBox tCi 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      ForeColor       =   12582912
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
   Begin MSMask.MaskEdBox tRuc 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   503
      _Version        =   393216
      ForeColor       =   12582912
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99 999 999 9999"
      PromptChar      =   "_"
   End
   Begin ComctlLib.ListView lOperacion 
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(Filtro de Solicitudes resueltas y no facturadas. No        visibles en formulario solicitudes resueltas)."
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&R.U.C.:"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S/D"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lTitular 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S/D"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&C.I.:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   135
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   120
      Top             =   40
      Width           =   8775
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DeSolicitud.frx":1272
            Key             =   "Si"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DeSolicitud.frx":158C
            Key             =   "No"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DeSolicitud.frx":18A6
            Key             =   "Alerta"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "DeSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gCliente As Long
Dim aSeleccionado As Long
Dim iJobCon As Integer, iJobCre As Integer

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then Unload Me
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = 5190

    InicializoCrystalEngine
    SetearLView lvValores.Grilla Or lvValores.FullRow, lOperacion
    EncabezadoLista
    Screen.MousePointer = 11
    
End Sub

Private Sub InicializoCrystalEngine()
    
    'Inicializa el Engine del Crystal y setea la impresora para el JOB
    On Error GoTo ErrCrystal
        
    'Inicializo el Reporte Para el Credito-----------------------------------------------------------------------------------
    iJobCre = crAbroReporte(prmPathListados & "Credito.RPT")
    If iJobCre = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paICreditoN) Then SeteoImpresoraPorDefecto paICreditoN
    If Not crSeteoImpresora(iJobCre, Printer, paICreditoB) Then GoTo ErrCrystal
    '----------------------------------------------------------------------------------------------------------------------------
    
    'Inicializo el Reporte Para el Conforme---------------------------------------------------------------------------------
    iJobCon = crAbroReporte(prmPathListados & "Conforme.RPT")
    If iJobCon = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paIConformeN) Then SeteoImpresoraPorDefecto paIConformeN
    If Not crSeteoImpresora(iJobCon, Printer, paIConformeB) Then GoTo ErrCrystal
    '----------------------------------------------------------------------------------------------------------------------------
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError Trim(crMsgErr) & " No se podrán imprimir facturas."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    crCierroTrabajo (iJobCre)        'Cierro los reportes de credito y conforme
    crCierroTrabajo (iJobCon)
    
    Forms(Forms.Count - 2).SetFocus
    Exit Sub
    
End Sub

Private Sub lOperacion_DblClick()
    AccionFacturar
End Sub

Private Sub lOperacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionFacturar
End Sub

Private Sub tCi_GotFocus()
    tCi.SelStart = 0
    tCi.SelLength = 11
End Sub

Private Sub tCi_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF4: If Shift = 0 Then BuscarClientes TipoCliente.Persona
    End Select
    
End Sub

Private Sub TCI_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tCi.Tag) = Trim(tCi.Text) Then tRuc.SetFocus: Exit Sub
        
        Dim aCi As String
        Screen.MousePointer = 11
        
        If Len(clsGeneral.QuitoFormatoCedula(tCi.Text)) = 7 Then tCi.Text = clsGeneral.AgregoDigitoControlCI(tCi.Text)
                
        'Valido la Cédula ingresada----------
        If Trim(tCi.Text) <> FormatoCedula Then
            If Len(clsGeneral.QuitoFormatoCedula(tCi.Text)) <> 8 Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            If Not clsGeneral.CedulaValida(clsGeneral.QuitoFormatoCedula(tCi.Text)) Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        
        'Busco el Cliente -----------------------
        If Trim(tCi.Text) <> FormatoCedula Then
            gCliente = BuscoClienteCIRUC(clsGeneral.QuitoFormatoCedula(tCi.Text))
            If gCliente = 0 Then
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para la cédula ingresada.", vbExclamation, "ATENCIÓN"
            Else
                 CargoDatosCliente gCliente
            End If
        Else
            tRuc.SetFocus
        End If
        Screen.MousePointer = 0
    End If

End Sub

Private Function BuscoClienteCIRUC(CiRuc As String) As Long

    On Error GoTo errBuscar
    BuscoClienteCIRUC = 0
    Cons = "Select * from Cliente Where CliCiRuc = '" & Trim(CiRuc) & "'"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then BuscoClienteCIRUC = rsAux!CliCodigo
    rsAux.Close
    Exit Function

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el cliente."
End Function

Private Sub tRuc_GotFocus()
    tRuc.SelStart = 0
    tRuc.SelLength = 15
End Sub

Private Sub tRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF4: If Shift = 0 Then BuscarClientes TipoCliente.Empresa
    End Select
    
End Sub

Private Sub tRuc_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tRuc.Tag) = Trim(tRuc.Text) Then tCi.SetFocus: Exit Sub
        
        If Trim(tRuc.Text) <> "" Then
            Screen.MousePointer = 11
            gCliente = BuscoClienteCIRUC(Trim(tRuc.Text))
            If gCliente = 0 Then
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para el número de RUC ingresado.", vbExclamation, "ATENCIÓN"
            Else
                'Cargo Datos del Cliente Seleccionado------------------------------------------------
                 CargoDatosCliente gCliente
            End If
        Else
            tCi.SetFocus
        End If
        Screen.MousePointer = 0
    End If
    
End Sub

Private Sub BuscarClientes(aTipoCliente As Integer)
    
    Screen.MousePointer = 11
    On Error GoTo errCargar
    
    Dim aIdCliente As Long, aTipo As Integer
    Dim objBuscar As New clsBuscarCliente
    
    If aTipoCliente = TipoCliente.Persona Then objBuscar.ActivoFormularioBuscarClientes cBase, Persona:=True
    If aTipoCliente = TipoCliente.Empresa Then objBuscar.ActivoFormularioBuscarClientes cBase, Empresa:=True
        
    aIdCliente = objBuscar.BCClienteSeleccionado
    aTipo = objBuscar.BCTipoClienteSeleccionado
    Set objBuscar = Nothing
    Me.Refresh
    DoEvents
    
    If aIdCliente <> 0 Then
        gCliente = aIdCliente
        CargoDatosCliente gCliente
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub EncabezadoLista()
        
    lOperacion.ColumnHeaders.Add , , "Código", 600
    lOperacion.ColumnHeaders.Add , , "Fecha", 600
    lOperacion.ColumnHeaders.Add , , "Estado", 400
    lOperacion.ColumnHeaders.Add , , "Monto ", 1000, 1
    lOperacion.ColumnHeaders.Add , , "Cuotas", 500
    lOperacion.ColumnHeaders.Add , , "Artículos", 2500
    lOperacion.ColumnHeaders.Add , , "Comentarios", 1100
    
End Sub

Private Sub CargoSolicitud(Cliente As Long)

    lOperacion.ListItems.Clear
    lOperacion.Refresh
    If Cliente = 0 Then Exit Sub

Dim aMonto As Currency
Dim aCuota As Long
Dim aCodSolicitud As Long
Dim aFecha As Date
Dim aMoneda As Integer
Dim aArticulo As String
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    Cons = "Select * From Solicitud, RenglonSolicitud, Articulo, TipoCuota" _
           & " Where SolCliente = " & Cliente _
           & " And SolProceso Not In ( " & TipoResolucionSolicitud.Facturada & "," & TipoResolucionSolicitud.Facturando & ")" _
           & " And SolEstado <> " & EstadoSolicitud.Pendiente _
           & " And SolCodigo = RSoSolicitud" _
           & " And RSoTipoCuota = TCuCodigo" _
           & " And RSoArticulo = ArtID" _
           & " Order by SolCodigo DESC"

    Set rsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsAux.EOF
        Set itmX = lOperacion.ListItems.Add(, , rsAux!SolCodigo)
        itmX.Tag = rsAux!SolCodigo
        
        Select Case rsAux!SolEstado
            Case EstadoSolicitud.Aprovada: itmX.SmallIcon = "Si"
            Case EstadoSolicitud.Rechazada: itmX.SmallIcon = "No"
            Case EstadoSolicitud.Condicional: itmX.SmallIcon = "Alerta"
        End Select
        
        itmX.SubItems(1) = Format(rsAux!SolFecha, "dd/mm/yy")
        If IsNull(rsAux!SolVisible) Then itmX.SubItems(2) = "Normal" Else itmX.SubItems(2) = "Oculta"
        itmX.SubItems(4) = Trim(rsAux!TCuAbreviacion)
        
        If Not IsNull(rsAux!SolComentarioR) Then itmX.SubItems(6) = Trim(rsAux!SolComentarioR)
        
        aCuota = rsAux!TCuCodigo
        aCodSolicitud = rsAux!SolCodigo
        aFecha = rsAux!SolFecha
        aMoneda = rsAux!SolMoneda
        aMonto = 0
        aArticulo = ""
        Do While aCuota = rsAux!TCuCodigo And aCodSolicitud = rsAux!SolCodigo
            aArticulo = aArticulo & Trim(rsAux!ArtNombre) & "; "
            If IsNull(rsAux!RSoValorEntrega) Then   '------------------------------------------------------
                aMonto = aMonto + rsAux!RSoValorCuota * rsAux!TCuCantidad
            Else
                aMonto = aMonto + rsAux!RSoValorEntrega + (rsAux!RSoValorCuota * rsAux!TCuCantidad)
            End If
            rsAux.MoveNext
            
            If rsAux.EOF Then
                itmX.SubItems(3) = BuscoSignoMoneda(aMoneda) & " " & Format(aMonto, "#,##0.00")
                itmX.SubItems(5) = Mid(aArticulo, 1, Len(aArticulo) - 2)
                Exit Do
            End If
            
        Loop
        If rsAux.EOF Then Exit Do
        itmX.SubItems(3) = BuscoSignoMoneda(aMoneda) & " " & Format(aMonto, "#,##0.00")
        itmX.SubItems(5) = Mid(aArticulo, 1, Len(aArticulo) - 2)
    Loop
    
    rsAux.Close
    If lOperacion.ListItems.Count > 0 Then lOperacion.SetFocus
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar las solicitudes realizadas."
End Sub


Private Sub CargoDatosCliente(Cliente As Long)

    On Error GoTo errCliente
    'Cargo Datos Tabla Cliente----------------------------------------------------------------------
    
    Cons = "Select CliCiRuc, CliTipo, CliDireccion, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
           & " From Cliente, CPersona " _
           & " Where CliCodigo = " & Cliente _
           & " And CliCodigo = CPeCliente " _
                                                & " UNION " _
           & " Select CliCiRuc, CliTipo, CliDireccion, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
           & " From Cliente, CEmpresa " _
           & " Where CliCodigo = " & Cliente _
           & " And CliCodigo = CEmCliente"

    Set rsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not IsNull(rsAux!CliCIRuc) Then      'CI o RUC
        Select Case rsAux!CliTipo
            Case TipoCliente.Persona
                tCi.Text = clsGeneral.RetornoFormatoCedula(rsAux!CliCIRuc)
                tCi.Tag = Trim(tCi.Text)
                tRuc.Text = "": tRuc.Tag = ""
            Case TipoCliente.Empresa
                tRuc.Text = Trim(rsAux!CliCIRuc)
                tRuc.Tag = Trim(tRuc.Text)
                tCi.Text = FormatoCedula: tCi.Tag = FormatoCedula
        End Select
    End If
    
    lTitular.Caption = Trim(rsAux!Nombre)
    'Direccion
    If Not IsNull(rsAux!CliDireccion) Then lDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, rsAux!CliDireccion) Else: lDireccion.Caption = "S/D"
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------
    
    CargoSolicitud Cliente
    Exit Sub
    
errCliente:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente."
End Sub


Private Sub AccionFacturar()

    If lOperacion.ListItems.Count = 0 Then Exit Sub
    
    Set itmX = lOperacion.SelectedItem
    
    If FormActivo("EsCondicional") Then EsCondicional.SetFocus: Exit Sub
        
    aSeleccionado = CLng(itmX.Tag)
    
    Dim vbDefButton As Long: vbDefButton = vbDefaultButton1
    
    If itmX.SmallIcon <> "No" Then
        If Not ValidoVigenciaPrecios(aSeleccionado) Then
            MsgBox "Los precios de los artículos han cambiado." & vbCrLf & _
                        "La fecha de solicitud es anterior a la de los precios vigentes de los artículos.", vbExclamation, "Los Precios han Cambiado !!!"
            vbDefButton = vbDefaultButton2
        End If
    End If
    
    If MsgBox("Confirma acceder a la pantalla de facturación crédito.", vbQuestion + vbYesNo + vbDefButton, "ATENCIÓN") = vbNo Then Exit Sub
    
    If itmX.SmallIcon <> "No" Then
        Select Case BloqueoSolicitud(aSeleccionado) 'Si es 1 esta todo OK--------------------------
            Case 0  'OTRO USUARIO
                Screen.MousePointer = 0
                MsgBox "La solicitud se está facturando por otro usuario. No podrá visualizarla.", vbExclamation, "ATENCIÓN"
                Exit Sub
           
            Case -1 'ERROR o FUE RESUELTA
                Screen.MousePointer = 0
                MsgBox "Posiblemente la solicitud ya fue facturada.", vbExclamation, "ATENCIÓN"
                Screen.MousePointer = 11
                CargoSolicitud gCliente
                Screen.MousePointer = 0
                Exit Sub
        End Select  '----------------------------------------------------------------------------------------
    End If
    
    Screen.MousePointer = 11
    
    EsCondicional.pSolicitud = aSeleccionado
    EsCondicional.prmPreciosViejos = (vbDefButton = vbDefaultButton2)

    Select Case itmX.SmallIcon
        Case "Alerta": EsCondicional.pSolicitudEstado = EstadoSolicitud.Condicional
        Case "No": EsCondicional.pSolicitudEstado = EstadoSolicitud.Rechazada
        Case "Si": EsCondicional.pSolicitudEstado = EstadoSolicitud.Aprovada
    End Select
    
    'Antes de acceder si  no es rechazada la borro
    If itmX.SmallIcon <> "No" Then
        I = 1
        Do While I <= lOperacion.ListItems.Count
            If lOperacion.ListItems(I).Tag = aSeleccionado Then lOperacion.ListItems.Remove I Else I = I + 1
        Loop
        lOperacion.Refresh
    End If      '----------------------------------------------
    
    EsCondicional.pJobConforme = iJobCon
    EsCondicional.pJobCredito = iJobCre
    EsCondicional.Show vbModal, Me
    
End Sub

'---------------------------------------------------------------------------------------------------------------
'   Valores que Retorna:    -1: Error o No Existe
'                                       0: Facturando o Facturada
'                                       1: Bloqueada OK
Private Function BloqueoSolicitud(Codigo As Long)

    BloqueoSolicitud = 0
    Screen.MousePointer = 11
    On Error GoTo errorBT
    
    'Bloqueo la solicitud y Actulizo el SolTipoResolucion a Facturando
    Cons = "Select * from Solicitud Where SolCodigo = " & Codigo
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsAux.EOF Then
        
        If rsAux!SolProceso <> TipoResolucionSolicitud.Facturada _
            And rsAux!SolProceso <> TipoResolucionSolicitud.Facturando _
            And Not IsNull(rsAux!SolUsuarioR) Then

            cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
            On Error GoTo errorET
            
            rsAux.Requery
            
            If rsAux!SolProceso = TipoResolucionSolicitud.Facturada Or rsAux!SolProceso = TipoResolucionSolicitud.Facturando Then
                cBase.RollbackTrans
                Exit Function
            End If
            
            rsAux.Edit
            rsAux!SolProceso = TipoResolucionSolicitud.Facturando
            rsAux.Update
            
            cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
            rsAux.Requery
            BloqueoSolicitud = 1    'OK
        
        Else
            BloqueoSolicitud = -1    'OK
        End If
    End If
    
    rsAux.Close
    Screen.MousePointer = 0
    Exit Function

errorBT:
    BloqueoSolicitud = -1
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Function
errorET:
    Resume ErrorRoll
ErrorRoll:
    BloqueoSolicitud = -1
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
End Function

Private Function ValidoVigenciaPrecios(mIDSolicitud As Long) As Boolean
On Error GoTo errVig

    ValidoVigenciaPrecios = True

    Cons = "Select SolFecha, Max(PViVigencia) as SolVigencia" & _
            " From Solicitud, RenglonSolicitud, PrecioVigente " & _
            " Where RSoSolicitud = " & mIDSolicitud & _
            " And SolCodigo = RSoSolicitud" & _
            " And RSoArticulo = PViArticulo " & _
            " And PViTipoCuota = " & paTipoCuotaContado & _
            " And PViMoneda = SolMoneda " & _
            " Group by SolFecha"
    
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If rsAux!SolVigencia > rsAux!SolFecha Then ValidoVigenciaPrecios = False
    End If
    rsAux.Close
    
    Exit Function
errVig:
End Function
