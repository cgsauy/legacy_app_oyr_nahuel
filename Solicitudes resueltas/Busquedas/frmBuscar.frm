VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Solicitudes"
   ClientHeight    =   4785
   ClientLeft      =   2310
   ClientTop       =   4635
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBuscar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8895
   Begin VB.Frame Frame1 
      Caption         =   "Selección del Cliente"
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   8775
      Begin MSMask.MaskEdBox tCi 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&C.I.:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   255
         Width           =   855
      End
      Begin VB.Label lTitular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   7455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&R.U.C.:"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(*) Filtro de Solicitudes resueltas y no facturadas.             No visibles en formulario solicitudes resueltas."
         Height          =   495
         Left            =   4800
         TabIndex        =   6
         Top             =   240
         Width           =   3855
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid lOperacion 
      Height          =   3555
      Left            =   60
      TabIndex        =   4
      Top             =   1140
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6271
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
      BackColorBkg    =   -2147483643
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
            Picture         =   "frmBuscar.frx":1272
            Key             =   "Si"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBuscar.frx":158C
            Key             =   "No"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBuscar.frx":18A6
            Key             =   "Alerta"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim prmCliente As Long

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then Unload Me
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    ObtengoSeteoForm Me
    
    EncabezadoLista
    Screen.MousePointer = 11
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub lOperacion_DblClick()
    AccionFacturar
End Sub

Private Sub lOperacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionFacturar
End Sub

Private Sub tCi_GotFocus()
    tCi.SelStart = 0: tCi.SelLength = 11
End Sub

Private Sub tCi_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF4: If Shift = 0 Then BuscarClientes TipoCliente.Cliente
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
            prmCliente = BuscoClienteCIRUC(clsGeneral.QuitoFormatoCedula(tCi.Text))
            If prmCliente = 0 Then
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para la cédula ingresada.", vbExclamation, "ATENCIÓN"
            Else
                 CargoDatosCliente prmCliente
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
    cons = "Select * from Cliente Where CliCiRuc = '" & Trim(CiRuc) & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then BuscoClienteCIRUC = rsAux!CliCodigo
    rsAux.Close
    Exit Function

errBuscar:
    clsGeneral.OcurrioError "Error al buscar el cliente.", Err.Description
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
            prmCliente = BuscoClienteCIRUC(Trim(tRuc.Text))
            If prmCliente = 0 Then
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para el número de RUC ingresado.", vbExclamation, "ATENCIÓN"
            Else
                'Cargo Datos del Cliente Seleccionado------------------------------------------------
                 CargoDatosCliente prmCliente
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
    
    If aTipoCliente = TipoCliente.Cliente Then objBuscar.ActivoFormularioBuscarClientes cBase, Persona:=True
    If aTipoCliente = TipoCliente.Empresa Then objBuscar.ActivoFormularioBuscarClientes cBase, Empresa:=True
    Me.Refresh
    
    aIdCliente = objBuscar.BCClienteSeleccionado
    aTipo = objBuscar.BCTipoClienteSeleccionado
    Set objBuscar = Nothing
    
    
    If aIdCliente <> 0 Then
        prmCliente = aIdCliente
        CargoDatosCliente prmCliente
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub EncabezadoLista()
        
    With lOperacion
        .Rows = 1: .Cols = 1
        .FormatString = "<Código|Fecha|<Estado|>Monto|Cuotas|<Artículos|<Comentarios"
        .ColWidth(0) = 900: .ColWidth(1) = 750: .ColWidth(2) = 600: .ColWidth(3) = 1200: .ColWidth(4) = 700: .ColWidth(5) = 2500
        .WordWrap = False
        .MergeCells = flexMergeSpill: .ExtendLastCol = True
    End With
    
End Sub

Private Sub CargoSolicitud(mCliente As Long)

    lOperacion.Rows = 1
    lOperacion.Refresh
    If mCliente = 0 Then Exit Sub

Dim aMonto As Currency
Dim aCuota As Long
Dim aCodSolicitud As Long
Dim aFecha As Date
Dim aMoneda As Integer
Dim aArticulo As String
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    cons = "Select * From Solicitud Left Outer Join SolicitudResolucion ON SolCodigo = ResSolicitud," & _
                        " RenglonSolicitud, Articulo, TipoCuota" _
           & " Where SolCliente = " & mCliente _
           & " And SolProceso Not In ( " & TipoResolucionSolicitud.Facturada & "," & TipoResolucionSolicitud.Facturando & ")" _
           & " And SolEstado <> " & EstadoSolicitud.Pendiente _
           & " And SolCodigo = RSoSolicitud" _
           & " And RSoTipoCuota = TCuCodigo" _
           & " And RSoArticulo = ArtID" _
           & " And ResNumero = (Select Max(ResNumero) From SolicitudResolucion Where SolCodigo = ResSolicitud) " _
           & " Order by SolCodigo DESC"

    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        Do While Not rsAux.EOF
            With lOperacion
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = rsAux!SolCodigo
            
                Select Case rsAux!SolEstado
                    Case EstadoSolicitud.Aprovada:  .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("Si").ExtractIcon: .Cell(flexcpData, .Rows - 1, 0) = EstadoSolicitud.Aprovada
                    Case EstadoSolicitud.Rechazada: .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("No").ExtractIcon: .Cell(flexcpData, .Rows - 1, 0) = EstadoSolicitud.Rechazada
                    Case EstadoSolicitud.Condicional: .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages("Alerta").ExtractIcon: .Cell(flexcpData, .Rows - 1, 0) = EstadoSolicitud.Condicional
                End Select
            
                .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!SolFecha, "dd/mm/yy")
                If IsNull(rsAux!SolVisible) Then .Cell(flexcpText, .Rows - 1, 2) = "Normal" Else .Cell(flexcpText, .Rows - 1, 2) = "Oculta"
                .Cell(flexcpText, .Rows - 1, 4) = Trim(rsAux!TCuAbreviacion)
                
                If Not IsNull(rsAux!ResComentario) Then .Cell(flexcpText, .Rows - 1, 6) = Trim(rsAux!ResComentario)
                
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
                        .Cell(flexcpText, .Rows - 1, 3) = BuscoSignoMoneda(aMoneda) & " " & Format(aMonto, "#,##0.00")
                        .Cell(flexcpText, .Rows - 1, 5) = Mid(aArticulo, 1, Len(aArticulo) - 2)
                        Exit Do
                    End If
                    
                Loop
                If rsAux.EOF Then Exit Do
                .Cell(flexcpText, .Rows - 1, 3) = BuscoSignoMoneda(aMoneda) & " " & Format(aMonto, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 5) = Mid(aArticulo, 1, Len(aArticulo) - 2)
            End With
        Loop
        
    Else
        MsgBox "No hay solicitudes pendientes para facturar, que pertenezcan al cliente seleccionado.", vbInformation, "No hay datos."
    End If
    rsAux.Close
    
    If lOperacion.Rows > 1 Then lOperacion.SetFocus
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar las solicitudes realizadas.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub CargoDatosCliente(Cliente As Long)

    On Error GoTo errCliente
    'Cargo Datos Tabla Cliente----------------------------------------------------------------------
    
    cons = "Select CliCiRuc, CliTipo, CliDireccion, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
           & " From Cliente, CPersona " _
           & " Where CliCodigo = " & Cliente _
           & " And CliCodigo = CPeCliente " _
                                                & " UNION " _
           & " Select CliCiRuc, CliTipo, CliDireccion, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
           & " From Cliente, CEmpresa " _
           & " Where CliCodigo = " & Cliente _
           & " And CliCodigo = CEmCliente"

    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not IsNull(rsAux!CliCIRuc) Then      'CI o RUC
        Select Case rsAux!CliTipo
            Case TipoCliente.Cliente
                tCi.Text = clsGeneral.RetornoFormatoCedula(rsAux!CliCIRuc)
                tCi.Tag = Trim(tCi.Text)
                tRuc.Text = "": tRuc.Tag = ""
                
            Case TipoCliente.Empresa
                tRuc.Text = Trim(rsAux!CliCIRuc): tRuc.Tag = Trim(tRuc.Text)
                tCi.Text = FormatoCedula: tCi.Tag = FormatoCedula
        End Select
    End If
    
    lTitular.Caption = Trim(rsAux!Nombre)
    
    lDireccion.Caption = "S/D"
    If Not IsNull(rsAux!CliDireccion) Then lDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, rsAux!CliDireccion, True, True)
    
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------
    
    CargoSolicitud Cliente
    Exit Sub
    
errCliente:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los datos del cliente.", Err.Description
End Sub

Private Sub AccionFacturar()

    If lOperacion.Rows = 1 Then Exit Sub
    
    Dim mValor As Long
            
    mValor = lOperacion.Cell(flexcpText, lOperacion.Row, 0)
        
    Dim vbDefButton As Long: vbDefButton = vbDefaultButton1
    
    If lOperacion.Cell(flexcpData, lOperacion.Row, 0) <> EstadoSolicitud.Rechazada Then 'If itmX.SmallIcon <> "No" Then
        If Not ValidoVigenciaPrecios(mValor) Then
            MsgBox "Los precios de los artículos han cambiado." & vbCrLf & _
                        "La fecha de solicitud es anterior a la de los precios vigentes de los artículos.", vbExclamation, "Los Precios han Cambiado !!!"
            vbDefButton = vbDefaultButton2
        End If
    End If
    
    If MsgBox("Confirma acceder a la pantalla de facturación crédito ?.", vbQuestion + vbYesNo, "Facturar Solicitud") = vbNo Then Exit Sub
        
    EjecutarApp prmPathApp & "\Solicitudes_Resueltas.exe", "/ID=" & CStr(mValor)
        
    Exit Sub
    'Antes de acceder si es no es rechazada la borro
    If lOperacion.Cell(flexcpData, lOperacion.Row, 0) <> EstadoSolicitud.Rechazada Then
        With lOperacion
            I = 1
            Do While I <= .Rows - 1
                If .Cell(flexcpText, I, 0) = mValor Then .RemoveItem I Else I = I + 1
            Loop
            .Refresh
        End With
    End If      '----------------------------------------------
    
End Sub

Private Function BuscoSignoMoneda(idMoneda As Integer) As String

    On Error Resume Next
    Dim rsM As rdoResultset
    BuscoSignoMoneda = ""
    
    cons = "Select * from Moneda Where MonCodigo = " & idMoneda
    Set rsM = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsM.EOF Then BuscoSignoMoneda = Trim(rsM!MonSigno)
    rsM.Close
    
End Function

Private Function ValidoVigenciaPrecios(mIDSolicitud As Long) As Boolean
On Error GoTo errVig

    ValidoVigenciaPrecios = True

    cons = "Select SolFecha, Max(PViVigencia) as SolVigencia" & _
            " From Solicitud, RenglonSolicitud, PrecioVigente " & _
            " Where RSoSolicitud = " & mIDSolicitud & _
            " And SolCodigo = RSoSolicitud" & _
            " And RSoArticulo = PViArticulo " & _
            " And PViTipoCuota = " & paTipoCuotaContado & _
            " And PViMoneda = SolMoneda " & _
            " Group by SolFecha"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If rsAux!SolVigencia > rsAux!SolFecha Then ValidoVigenciaPrecios = False
    End If
    rsAux.Close
    
    Exit Function
errVig:
End Function


