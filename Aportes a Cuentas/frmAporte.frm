VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{162F4D73-979C-4E83-84D4-C9D8E6AB1FE3}#1.8#0"; "ORGCTR~1.OCX"
Object = "{5EA2D00A-68AC-4888-98E6-53F6035BBEE3}#1.3#0"; "CGSABuscarCliente.ocx"
Begin VB.Form frmAporte 
   Caption         =   "Aportes a Cuentas"
   ClientHeight    =   5145
   ClientLeft      =   2460
   ClientTop       =   2505
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAporte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8415
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "A Cuenta"
      ForeColor       =   &H00000080&
      Height          =   2835
      Left            =   120
      TabIndex        =   16
      Top             =   1980
      Width           =   8175
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   45
         TabIndex        =   9
         Top             =   900
         Width           =   6615
      End
      Begin VB.CommandButton bEmitir 
         Caption         =   "&Emitir Recibo"
         Height          =   285
         Left            =   6840
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox tTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         MaxLength       =   45
         TabIndex        =   8
         Top             =   600
         Width           =   6615
      End
      Begin VB.ComboBox cCuenta 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   13
         Top             =   2040
         Width           =   5295
      End
      Begin prjBuscarCliente.ucBuscarCliente txtCliente2 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         Text            =   "_.___.___-_"
         DocumentoCliente=   1
         QueryFind       =   "EXEC [dbo].[prg_BuscarCliente] 0, '', '', '', '', '', '[KeyQuery]', 0, 0, '', '', 7"
         KeyQuery        =   "[KeyQuery]"
         NeedCheckDigit  =   0   'False
      End
      Begin VB.Label Label13 
         Caption         =   "Comentarios.:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lDiferencia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6720
         TabIndex        =   28
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label lDifEx 
         Caption         =   "Diferencia:"
         Height          =   255
         Left            =   5880
         TabIndex        =   27
         Top             =   2460
         Width           =   855
      End
      Begin VB.Label lAportado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4200
         TabIndex        =   26
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Aportado:"
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   2460
         Width           =   855
      End
      Begin VB.Label lContado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   2460
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Precio Ctdo:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label lblInfoPersona2 
         Caption         =   "&Persona:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lCCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Top             =   1620
         Width           =   6615
      End
      Begin VB.Label Label4 
         Caption         =   "&Se destina a:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "&Nombre de Cta.:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label8 
         Caption         =   "&Tipo de Cuenta:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Aporte"
      ForeColor       =   &H00000080&
      Height          =   1755
      Left            =   120
      TabIndex        =   15
      Top             =   180
      Width           =   8175
      Begin prjBuscarCliente.ucBuscarCliente txtCliente 
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Text            =   "_.___.___-_"
         DocumentoCliente=   1
         QueryFind       =   "EXEC [dbo].[prg_BuscarCliente] 0, '', '', '', '', '', '[KeyQuery]', 0, 0, '', '', 7"
         KeyQuery        =   "[KeyQuery]"
         NeedCheckDigit  =   0   'False
      End
      Begin VB.TextBox tImporte 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         MaxLength       =   14
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
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
      Begin OrgCtrlFlat.orgHiperLink hlConQueSeCobra 
         Height          =   240
         Left            =   4080
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   423
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "¿Con que se cobra?  $15.255.555"
         MouseIcon       =   "frmAporte.frx":0442
         MousePointer    =   99
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   960
         Width           =   6735
      End
      Begin VB.Label lCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   600
         Width           =   6735
      End
      Begin VB.Label Label3 
         Caption         =   "&Aporte:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblInfoCliente 
         Caption         =   "&Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   4890
      Width           =   8415
      _ExtentX        =   14843
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
            Object.Width           =   9181
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lPrinter 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Label14"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   540
      TabIndex        =   30
      Top             =   0
      Width           =   7755
   End
End
Attribute VB_Name = "frmAporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCnfgPrint As New clsImpresoraTicketsCnfg
Dim oConQueSeCobra As clsDesicionConQuePaga
Dim oCnfgPrintSalidaCaja As New clsImpresoraTicketsCnfg

Public prmTipoCta As Integer
Public prmIdCta As Long
Public prmIDAporta As Long

Private jobnum As Integer       'Nro. de Trabajo para el recibo
Private CantForm As Integer    'Cantidad de formulas del reporte

Dim gPathListados As String
Dim paBD As String

Private Sub bEmitir_Click()
    AccionGrabar
End Sub

Private Sub cCuenta_Click()
     
    On Error Resume Next
    If cCuenta.ItemData(cCuenta.ListIndex) = Cuenta.Colectivo Then
        txtCliente2.Text = vbNullString
        lCliente.Caption = ""
        txtCliente2.Enabled = False
        lCCliente.Enabled = False: lCCliente.BackColor = Colores.Gris
        tTitulo.Enabled = True: tTitulo.BackColor = Colores.Blanco
    End If
    
    If cCuenta.ItemData(cCuenta.ListIndex) = Cuenta.Personal Then
        txtCliente2.Enabled = True
        lCCliente.Enabled = True: lCCliente.BackColor = Colores.Azul
        tTitulo.Text = "": lCCliente.Caption = ""
        tTitulo.Enabled = False: tTitulo.BackColor = Colores.Gris
    End If
    tComentario.Text = ""
End Sub

Private Sub cCuenta_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And IsNumeric(tImporte.Text) Then
        If tTitulo.Enabled Then Foco tTitulo
        If txtCliente2.Enabled Then txtCliente2.SetFocus
    End If
    
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cMoneda.ListIndex <> -1 Then Foco tImporte
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    
    oCnfgPrintSalidaCaja.CargarConfiguracion "MovimientosDeCaja", "TickeadoraMovimientosDeCaja"
    oCnfgPrint.CargarConfiguracion "ImpresionDocumentos", "TicketCuota"
    
'    oCnfgPrint.CargarConfiguracion "ImpresionDocumentos", cnfgKeyTicketConformes
 '   If oCnfgPrint.Opcion = 0 Then
  '      lPrinter.Caption = "Imprimir Recibos en " & paIReciboN & " (Bja:" & paIReciboB & ")"
  '  Else
   '     lPrinter.Caption = "Imprimir Recibos en tickeadora " & oCnfgPrint.ImpresoraTickets
    'End If
    
    If oCnfgPrint.ImpresoraTickets = 0 Then
        MsgBox "No tiene asignada la tickeadora de recibos.", vbCritical, "ATENCIÓN"
        End
    End If
    lPrinter.Caption = "Imprimir Recibos en tickeadora " & oCnfgPrint.ImpresoraTickets
    
    paBD = PropiedadesConnect(miConexion.TextoConexion(logComercio), Database:=True)
    ChDir App.Path
    ChDir "..": gPathListados = CurDir & "\Reportes\"
    '------------------------------------------------------------------------------------------
    Cons = "Select MonCodigo, MonSigno from Moneda Where MonFactura = 1"
    CargoCombo Cons, cMoneda
    BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
    '------------------------------------------------------------------------------------------
    
    'Cargo datos en combo Cuenta-------------------------------------------------------
    cCuenta.AddItem "Colectivos": cCuenta.ItemData(cCuenta.NewIndex) = Cuenta.Colectivo
    cCuenta.AddItem "Personal": cCuenta.ItemData(cCuenta.NewIndex) = Cuenta.Personal
    BuscoCodigoEnCombo cCuenta, Cuenta.Colectivo
    '------------------------------------------------------------------------------------------
    
    InicializoCrystalEngine
    
    If prmTipoCta <> 0 And prmIdCta <> 0 Then
        If prmTipoCta = Cuenta.Colectivo Or prmTipoCta = Cuenta.Personal Then
            BuscoCodigoEnCombo cCuenta, CLng(prmTipoCta)
            CargoDatosCuenta prmTipoCta, prmIdCta
        End If
    End If
    
    If prmIDAporta > 0 Then CargoDatosAporta prmIDAporta
    
    Set oConQueSeCobra = Nothing
    hlConQueSeCobra.Caption = "¿Con que se cobra?"

    Set txtCliente.Connect = cBase
    txtCliente.NeedCheckDigit = True

    Set txtCliente2.Connect = cBase
    txtCliente2.NeedCheckDigit = True

    Screen.MousePointer = 0
    Exit Sub

errLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub InicializoCrystalEngine()
    
    'Inicializa el Engine del Crystal y setea la impresora para el JOB
    On Error GoTo ErrCrystal
    If crAbroEngine = 0 Then GoTo ErrCrystal: Exit Sub
    
    'Inicializo el Reporte y SubReportes
    jobnum = crAbroReporte(gPathListados & "Aporte.RPT")
    If jobnum = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paIReciboN) Then SeteoImpresoraPorDefecto paIReciboN
    'paperSize:=11 --> A5
    If Not crSeteoImpresora(jobnum, Printer, paIReciboB, paperSize:=13, mOrientation:=2) Then GoTo ErrCrystal

    'Obtengo la cantidad de formulas que tiene el reporte.
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError Trim(crMsgErr) & " No se podrán imprimir Recibos de Pago.", Err.Description & " Path: " & gPathListados
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    crCierroTrabajo jobnum
    crCierroEngine
    'GuardoSeteoForm Me
    
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    
End Sub

Private Sub hlConQueSeCobra_Click()
On Error GoTo errConQue
Dim oTransacciones As New clsCobrarConQuePaga
    
    hlConQueSeCobra.Caption = "¿Con que se cobra?"
    
    If Val(lCliente.Tag) = 0 Or Not IsNumeric(tImporte.Text) Then Exit Sub
    
    Set oTransacciones.Conexion = cBase
    Set oConQueSeCobra = New clsDesicionConQuePaga
    oTransacciones.AdmitirEfectivo = True
    oTransacciones.QuienInvoca = AporteACuenta
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    Set oConQueSeCobra = oTransacciones.AsignarPagosASaldo(Val(lCliente.Tag), CCur(tImporte.Text), paCodigoDeUsuario, Nothing)
    
    'Si no asigno lo pelo.
    If Not oConQueSeCobra Is Nothing Then
        
        If oConQueSeCobra.ConQuePaga.Count > 0 Then
        
            If oConQueSeCobra.SaldoAFavor > 0 Then
                MsgBox "NO PUEDE DEJAR SALDO A FAVOR.", vbExclamation, "ATENCIÓN"
                Set oConQueSeCobra = Nothing
                Exit Sub
            End If
        
            hlConQueSeCobra.Caption = "¿Con que se cobra? " & Format(oConQueSeCobra.TotalImporteAsignado, "#,##0.00")
        Else
            Set oConQueSeCobra = Nothing
        End If
        
    Else
        Set oConQueSeCobra = Nothing
    End If
    Exit Sub
    
errConQue:
    clsGeneral.OcurrioError "Error al asignar con que paga.", Err.Description, "Error"
End Sub

Private Sub Label1_Click()
    txtCliente.SetFocus
End Sub

Private Sub Label4_Click()
    Foco tArticulo
End Sub

Private Sub Label5_Click()
    Foco tTitulo
End Sub

Private Sub Label8_Click()
    cCuenta.SetFocus
End Sub

Private Sub lblInfoPersona2_Click()
    Foco txtCliente2
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = 0
    lContado.Caption = "": lAportado.Caption = "": lDiferencia.Caption = ""
End Sub

Private Sub tArticulo_GotFocus()
    tArticulo.SelStart = 0: tArticulo.SelLength = Len(tArticulo.Text)
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mSelectID As Long

    If KeyCode = vbKeyReturn Then
    
        If Val(tArticulo.Tag) <> 0 Then bEmitir.SetFocus:: Exit Sub
        If Trim(tArticulo.Text) = "" Then bEmitir.SetFocus: Exit Sub
            
        If Not IsNumeric(tArticulo.Text) Then   'Busqueda de articulos por lista de ayuda-------------------
            On Error GoTo errLista
            Screen.MousePointer = 11
            Dim aLista As New clsListadeAyuda
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Descripción' = ArtNombre  From Articulo " _
                    & " Where ArtNombre Like '" & Trim(tArticulo.Text) & "%'" _
                    & " And ArtEnUso = 1" _
                    & " Order by ArtNombre"
            
            mSelectID = aLista.ActivarAyuda(cBase, Cons, 5200, 1)
            If mSelectID > 0 Then
                tArticulo.Text = aLista.RetornoDatoSeleccionado(2)
                tArticulo.Tag = aLista.RetornoDatoSeleccionado(0)
                CargoDatosArticulo Val(tArticulo.Tag)
                bEmitir.SetFocus
            End If
            
            Set aLista = Nothing
            Screen.MousePointer = 0
        
        Else    'Busqueda de Articulos por codigo--------------
            Screen.MousePointer = 11
            Cons = "Select ArtID, ArtNombre from Articulo Where ArtCodigo = " & Trim(tArticulo.Text) & " And ArtEnUso = 1"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                tArticulo.Text = Trim(RsAux!ArtNombre)
                tArticulo.Tag = RsAux!ArtID
                CargoDatosArticulo Val(tArticulo.Tag)
                bEmitir.SetFocus
            Else
                MsgBox "No existe un artículo para el código ingresado.", vbExclamation, "ATENCIÓN"
            End If
            RsAux.Close
            Screen.MousePointer = 0
        End If
    End If
    Exit Sub
    
errLista:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al activar la lista de ayuda. ", Err.Description
End Sub

Private Sub CargoDatosArticulo(idArticulo As Long)
    On Error GoTo errArticulo
    Screen.MousePointer = 11
    Dim rs1 As rdoResultset
    Dim aSignoM As String
    
    Cons = "Select * from Moneda Where MonCodigo = " & paMonedaPesos
    Set rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rs1.EOF Then aSignoM = Trim(rs1!MonSigno) Else aSignoM = ""
    rs1.Close
    
    'Saco el precio contado del articulo para la moneda pesos------------------------------------------------
    Cons = "Select * from PrecioVigente, Moneda " _
            & " Where PViArticulo = " & idArticulo _
            & " And PViMoneda = " & paMonedaPesos _
            & " And PViTipoCuota = " & paTipoCuotaContado _
            & " And PViMoneda = MonCodigo"
    Set rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rs1.EOF Then lContado.Caption = Format(rs1!PViPrecio, FormatoMonedaP) Else lContado.Caption = ""
    rs1.Close
    '------------------------------------------------------------------------------------------------------------------
    
    'Cargo lo aportado hasta el momento em pesos-------------------------------------------------------------
    Cons = "Select DocMoneda, Total = Sum(DocTotal) From Documento, CuentaDocumento " _
            & " Where CDoTipo = " & cCuenta.ItemData(cCuenta.ListIndex) _
            & " And CDoIDArticulo = " & idArticulo
            
    Select Case cCuenta.ItemData(cCuenta.ListIndex)
        Case Cuenta.Colectivo: Cons = Cons & " And CDoIDTipo = " & Val(tTitulo.Tag)
        Case Cuenta.Personal: Cons = Cons & " And CDoIDTipo = " & Val(lCCliente.Tag)
    End Select
    
    Cons = Cons & " And CDoIDDocumento = DocCodigo " _
                        & " And DocAnulado = 0 " _
                        & " And DocTipo = " & TipoDocumento.ReciboDePago _
                        & " Group by DocMoneda"
                        
    Set rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Dim aAportes As Currency: aAportes = 0
    Dim aTC As Currency
    
    Do While Not rs1.EOF
        If rs1!DocMoneda = paMonedaPesos Then
            aAportes = aAportes + rs1!Total
        Else
            aTC = TasadeCambio(rs1!DocMoneda, paMonedaPesos, gFechaServidor)
            aAportes = aAportes + (rs1!Total * aTC)
        End If
        rs1.MoveNext
    Loop
    rs1.Close
    
    lAportado.Caption = Format(aAportes, FormatoMonedaP)
    '------------------------------------------------------------------------------------------------------------------
    
    If Val(lContado.Caption) <> 0 Then
        lDiferencia.Caption = Format(CCur(lContado.Caption) - CCur(lAportado.Caption), FormatoMonedaP)
        lContado.Caption = aSignoM & " " & lContado.Caption
        lAportado.Caption = aSignoM & " " & lAportado.Caption
        
        If CCur(lDiferencia.Caption) > 0 Then
            lDifEx.Caption = "Diferencia:": lDiferencia.ForeColor = vbBlack
        Else
            lDifEx.Caption = "Excedido:": lDiferencia.ForeColor = Colores.RojoClaro
        End If
        lDiferencia.Caption = aSignoM & " " & Format(Abs(CCur(lDiferencia.Caption)), FormatoMonedaP)
        
    End If
    Screen.MousePointer = 0
    Exit Sub
errArticulo:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

'Private Sub tCCi_Change()
'    lCCliente.Tag = 0
'    tArticulo.Text = ""
'End Sub
'
'Private Sub tCCi_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    On Error GoTo errorB
'    If KeyCode = vbKeyF4 Then
'        Dim objCliente As New clsBuscarCliente
'        objCliente.ActivoFormularioBuscarClientes cBase, True
'        Me.Refresh
'        If objCliente.BCTipoClienteSeleccionado = TipoCliente.Cliente Then
'            BuscoClienteID idPersona:=objCliente.BCClienteSeleccionado, Cuenta:=True
'        Else
'            BuscoClienteID idEmpresa:=objCliente.BCClienteSeleccionado, Cuenta:=True
'        End If
'        Set objCliente = Nothing
'        If Val(lCCliente.Tag) <> 0 Then Foco tArticulo
'    End If
'    Exit Sub
'
'errorB:
'    clsGeneral.OcurrioError "Ocurrió un error al procesar la información del cliente", Err.Description
'    Screen.MousePointer = 0
'End Sub

'Private Sub tCCi_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyReturn Then
'        If Val(lCCliente.Tag) <> 0 Then Foco tArticulo: Exit Sub
'        If Trim(tCCi.Text) = "" Then tCRuc.SetFocus: Exit Sub
'        If Len(tCCi.Text) = 7 Then tCCi.Text = clsGeneral.AgregoDigitoControlCI(tCCi.Text)
'
'        'Valido la Cédula ingresada-------------------------------------------------------------------------------------------
'        If Trim(tCCi.Text) <> "" Then
'            If Len(tCCi.Text) <> 8 Then
'                MsgBox "La cédula de identidad ingresada no es válida. Verifique", vbExclamation, "ATENCIÓN"
'                Exit Sub
'            End If
'            If Not clsGeneral.CedulaValida(tCCi.Text) Then
'                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
'                Exit Sub
'            End If
'        End If
'
'        BuscoCliente Cedula:=tCCi.Text, Cliente:=False, Cuenta:=True
'        If Val(lCCliente.Tag) <> 0 Then Foco tArticulo
'    End If
'
'End Sub
'
'
'Private Sub tCi_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    On Error GoTo errorB
'    If KeyCode = vbKeyF4 Then
'        Dim objCliente As New clsBuscarCliente
'        objCliente.ActivoFormularioBuscarClientes cBase, True
'        Me.Refresh
'        If objCliente.BCTipoClienteSeleccionado = TipoCliente.Cliente Then
'            BuscoClienteID idPersona:=objCliente.BCClienteSeleccionado, Cliente:=True
'        Else
'            BuscoClienteID idEmpresa:=objCliente.BCClienteSeleccionado, Cliente:=True
'        End If
'        Set objCliente = Nothing
'        If Val(lCliente.Tag) <> 0 Then Foco tImporte
'    End If
'    Exit Sub
'
'errorB:
'    clsGeneral.OcurrioError "Ocurrió un error al procesar la información del cliente", Err.Description
'    Screen.MousePointer = 0
'End Sub

'Private Sub tCi_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyReturn Then
'        If Val(lCliente.Tag) <> 0 Then Foco tImporte: Exit Sub
'        If Trim(tCi.Text) = "" Then tRuc.SetFocus: Exit Sub
'        If Len(tCi.Text) = 7 Then tCi.Text = clsGeneral.AgregoDigitoControlCI(tCi.Text)
'
'        'Valido la Cédula ingresada-------------------------------------------------------------------------------------------
'        If Trim(tCi.Text) <> "" Then
'            If Len(tCi.Text) <> 8 Then
'                MsgBox "La cédula de identidad ingresada no es válida. Verifique", vbExclamation, "ATENCIÓN"
'                Exit Sub
'            End If
'            If Not clsGeneral.CedulaValida(tCi.Text) Then
'                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
'                Exit Sub
'            End If
'        End If
'
'        BuscoCliente Cedula:=tCi.Text
'        If Val(lCliente.Tag) <> 0 Then Foco tImporte
'    End If
'
'End Sub
'
'Private Sub BuscoCliente(Optional Cedula As String = "", Optional Ruc As String = "", _
'                                    Optional Cliente As Boolean = True, Optional Cuenta As Boolean = False)
'
'    On Error GoTo errBuscar
'    Screen.MousePointer = 11
'
'    If Cedula <> "" Then
'        Cons = "Select Cliente.*, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2)" _
'                & " From Cliente, CPersona " _
'                & " Where CliCiRuc = '" & Cedula & "'" _
'                & " And CliCodigo = CPeCliente"
'    End If
'
'    If Ruc <> "" Then
'        Cons = "Select Cliente.*, Nombre = (RTrim(CEmNombre) + RTrim(' (' + CEmFantasia) + ')')" _
'                & " From Cliente, CEmpresa " _
'                & " Where CliCiRuc = '" & Ruc & "'" _
'                & " And CliCodigo = CEmCliente"
'    End If
'
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'
'    If Not RsAux.EOF Then
'        If Cliente Then
'            If Cedula <> "" Then tRuc.Text = "" Else tCi.Text = ""
'            lCliente.Tag = RsAux!CliCodigo
'            lCliente.Caption = " " & Trim(RsAux!Nombre)
'            lDireccion.Caption = ""
'            If Not IsNull(RsAux!CliDireccion) Then lDireccion.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, , Localidad:=True)
'        Else
'            If Cedula <> "" Then tCRuc.Text = "" Else tCCi.Text = ""
'            lCCliente.Tag = RsAux!CliCodigo
'            lCCliente.Caption = " " & Trim(RsAux!Nombre)
'        End If
'
'    Else
'        MsgBox "No existe un cliente para la CI/Ruc, ingresado.", vbExclamation, "Cliente Inexistente"
'        If Cliente Then
'            lCliente.Caption = "": lCliente.Tag = ""
'            lDireccion.Caption = ""
'        Else
'            lCCliente.Caption = "": lCCliente.Tag = ""
'        End If
'    End If
'
'    RsAux.Close
'
'    If Val(lCliente.Tag) <> 0 Then ValidoDireccionEMail Val(lCliente.Tag)
'    Screen.MousePointer = 0
'    Exit Sub
'
'errBuscar:
'    Screen.MousePointer = 0
'    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
'End Sub

'Private Sub BuscoClienteID(Optional idPersona As Long = 0, Optional idEmpresa As Long = 0, _
'                                        Optional Cliente As Boolean = False, Optional Cuenta As Boolean = False)
'
'    On Error GoTo errBuscar
'    If idPersona = 0 And idEmpresa = 0 Then Exit Sub
'    Screen.MousePointer = 11
'
'    If idPersona <> 0 Then
'        Cons = "Select Cliente.*, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2)" _
'                & " From Cliente, CPersona " _
'                & " Where CliCodigo = " & idPersona _
'                & " And CliCodigo = CPeCliente"
'    End If
'
'    If idEmpresa <> 0 Then
'        Cons = "Select Cliente.*, Nombre = (RTrim(CEmNombre) + RTrim(' (' + CEmFantasia) + ')')" _
'                & " From Cliente, CEmpresa " _
'                & " Where CliCodigo = " & idEmpresa _
'                & " And CliCodigo = CEmCliente"
'    End If
'
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'
'    If Not RsAux.EOF Then
'        If Cliente Then
'            tCi.Text = "": tRuc.Text = ""
'            If Not IsNull(RsAux!CliCiRuc) Then
'                If idPersona <> 0 Then tCi.Text = Trim(RsAux!CliCiRuc) Else tRuc.Text = Trim(RsAux!CliCiRuc)
'            End If
'
'            lCliente.Tag = RsAux!CliCodigo
'            lCliente.Caption = " " & Trim(RsAux!Nombre)
'            lDireccion.Caption = ""
'            If Not IsNull(RsAux!CliDireccion) Then lDireccion.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, , Localidad:=True)
'        Else
'            tCCi.Text = "": tCRuc.Text = ""
'            If Not IsNull(RsAux!CliCiRuc) Then
'                If idPersona <> 0 Then tCCi.Text = Trim(RsAux!CliCiRuc) Else tCRuc.Text = Trim(RsAux!CliCiRuc)
'            End If
'
'            lCCliente.Tag = RsAux!CliCodigo
'            lCCliente.Caption = " " & Trim(RsAux!Nombre)
'        End If
'    Else
'        MsgBox "No existe un cliente para el código seleccionado.", vbExclamation, "Cliente Inexistente"
'        lCCliente.Caption = "": lCCliente.Tag = ""
'    End If
'
'    RsAux.Close
'
'    If Val(lCliente.Tag) <> 0 Then ValidoDireccionEMail Val(lCliente.Tag)
'
'    Screen.MousePointer = 0
'
'    Exit Sub
'errBuscar:
'    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
'    Screen.MousePointer = 0
'End Sub

Private Sub ValidoDireccionEMail(idCliente As Long)
On Error GoTo errVDE
    
    Dim bAlta As Boolean
    Cons = "Select Top 1 * from EMailDireccion Where EMDIDCliente = " & Val(lCliente.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    bAlta = RsAux.EOF
    RsAux.Close
    
    If bAlta Then
        If MsgBox("El cliente no tiene ingresada la dirección de e-mail." & vbCrLf & _
                        "Quiere ingresarla ahora ? ", vbYesNo + vbQuestion, "Falta Dirección de e-mail") = vbYes Then
                EjecutarApp prmPathApp & "\Emails.exe " & Val(lCliente.Tag)
        End If
    End If
    Exit Sub

errVDE:
    clsGeneral.OcurrioError "Error validar la dirección de correo.", Err.Description
    Screen.MousePointer = 0
End Sub


'Private Sub tCRuc_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    On Error GoTo errorB
'    If KeyCode = vbKeyF4 Then
'        Dim objCliente As New clsBuscarCliente
'        objCliente.ActivoFormularioBuscarClientes cBase, Empresa:=True
'        Me.Refresh
'        If objCliente.BCTipoClienteSeleccionado = TipoCliente.Cliente Then
'            BuscoClienteID idPersona:=objCliente.BCClienteSeleccionado, Cuenta:=True
'        Else
'            BuscoClienteID idEmpresa:=objCliente.BCClienteSeleccionado, Cuenta:=True
'        End If
'        Set objCliente = Nothing
'        If Val(lCCliente.Tag) <> 0 Then Foco tArticulo
'    End If
'    Exit Sub
'
'errorB:
'    clsGeneral.OcurrioError "Ocurrió un error al procesar la información del cliente", Err.Description
'    Screen.MousePointer = 0
'End Sub
'
'Private Sub tCRuc_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyReturn Then
'        If Val(lCCliente.Tag) <> 0 Then Foco tArticulo: Exit Sub
'        If Trim(tCRuc.Text) = "" Then tCCi.SetFocus: Exit Sub
'
'        If Len(tCRuc.Text) <> 12 Then
'            MsgBox "La número de RUC ingresado no es correcto. Verifique", vbExclamation, "ATENCIÓN"
'            Exit Sub
'        End If
'
'        BuscoCliente Ruc:=tCRuc.Text, Cliente:=False, Cuenta:=True
'        Foco tArticulo
'    End If
'
'End Sub

Private Sub tImporte_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And IsNumeric(tImporte.Text) Then cCuenta.SetFocus: Exit Sub
      
End Sub

Private Sub tImporte_LostFocus()
    If IsNumeric(tImporte.Text) Then tImporte.Text = Format(tImporte.Text, FormatoMonedaP)
End Sub

'Private Sub tRuc_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    On Error GoTo errorB
'    If KeyCode = vbKeyF4 Then
'        Dim objCliente As New clsBuscarCliente
'        objCliente.ActivoFormularioBuscarClientes cBase, Empresa:=True
'        Me.Refresh
'        If objCliente.BCTipoClienteSeleccionado = TipoCliente.Cliente Then
'            BuscoClienteID idPersona:=objCliente.BCClienteSeleccionado, Cliente:=True
'        Else
'            BuscoClienteID idEmpresa:=objCliente.BCClienteSeleccionado, Cliente:=True
'        End If
'        Set objCliente = Nothing
'        If Val(lCliente.Tag) <> 0 Then Foco tImporte
'    End If
'    Exit Sub
'
'errorB:
'    clsGeneral.OcurrioError "Ocurrió un error al procesar la información del cliente", Err.Description
'    Screen.MousePointer = 0
'End Sub

'Private Sub tRuc_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyReturn Then
'        If Val(lCliente.Tag) <> 0 Then Foco tImporte: Exit Sub
'        If Trim(tRuc.Text) = "" Then tCi.SetFocus: Exit Sub
'
'        If Len(tRuc.Text) <> 12 Then
'            MsgBox "La número de RUC ingresado no es correcto. Verifique", vbExclamation, "ATENCIÓN"
'            Exit Sub
'        End If
'
'        BuscoCliente Ruc:=tRuc.Text
'        Foco tImporte
'    End If
'
'End Sub

Private Sub tTitulo_Change()
    tTitulo.Tag = 0
    tArticulo.Text = ""
End Sub

Private Sub tTitulo_KeyPress(KeyAscii As Integer)

Dim mSelectID As Long

    If KeyAscii = vbKeyReturn Then
    
        If Val(tTitulo.Tag) <> 0 Then Foco tArticulo: Exit Sub
        If Trim(tTitulo.Text) = "" Then Exit Sub
            
            On Error GoTo errLista
            Screen.MousePointer = 11
            Dim aLista As New clsListadeAyuda
            Cons = "Select ColCodigo, 'Título' = ColNombre, 'Fecha Civil' = ColFechaCivil, 'Fecha Iglesia' = ColFechaIglesia, " _
                            & " 'Cliente1' = (RTrim(P1.CPeNombre1) + ' ' + RTrim(P1.CPeApellido1)),  " _
                            & " 'Cliente2' = (RTrim(P2.CPeNombre1) + ' ' + RTrim(P2.CPeApellido1))  " _
                    & " From Colectivo left Outer Join CPersona P2 On ColCliente2 = P2.CPeCliente, CPersona P1" _
                    & " Where ColNombre Like '" & Trim(tTitulo.Text) & "%'" _
                    & " And ColCerrado = 0" _
                    & " And ColCliente1 = P1.CPeCliente" _
                    & " Order by ColNombre"
                    
            mSelectID = aLista.ActivarAyuda(cBase, Cons, 8500, 1)
            If mSelectID <> 0 Then
                tTitulo.Text = aLista.RetornoDatoSeleccionado(1)
                tTitulo.Tag = aLista.RetornoDatoSeleccionado(0)
                Foco tArticulo
            End If
            
            Set aLista = Nothing
            
            If mSelectID <> 0 Then
                Cons = "Select  * from Colectivo Where ColCodigo = " & mSelectID
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    If Not IsNull(RsAux!ColComentario) Then tComentario.Text = Trim(RsAux!ColComentario) Else tComentario.Text = ""
                End If
                RsAux.Close
            End If
            
            Screen.MousePointer = 0
    End If
    Exit Sub
    
errLista:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al activar la lista de ayuda. ", Err.Description

End Sub

Private Sub AccionGrabar()
Dim aTexto As String, Serie As String, Numero As Long
Dim aDocumentoRecibo As Long

    If Not ValidoCampos Then Exit Sub
    
    'Valido que lo asignado no supere al monto de la cuota.
    If Not oConQueSeCobra Is Nothing Then
        If CCur(tImporte.Text) < oConQueSeCobra.SaldoAFavor Then
            MsgBox "El importe del aporte no puede ser menor al importe asignado por redpagos.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
    End If
    
    'Valido si tengo importe asignado.
    'Si el importe es mayor al valor final del recibo le digo que no puede.
    If MsgBox("Confirma emitir el recibo para el aporte ingresado.", vbQuestion + vbYesNo, "Emitir Recibo") = vbNo Then Exit Sub
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    FechaDelServidor
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    Dim sDocNom As String
    sDocNom = "Tkt" & IIf(oCnfgPrint.ImpresoraTickets > 1, oCnfgPrint.ImpresoraTickets, ".") & paDRecibo
    
    'Pido el Numero de Documento para hacer el RECIBO-------------
    aTexto = NumeroDocumento(sDocNom)
    Serie = Mid(aTexto, 1, 1): Numero = CLng(Trim(Mid(aTexto, 2, Len(aTexto))))

    'Inserto campos en la tabla documento---------------------------------------------------------------------
    Cons = "Select * from Documento Where DocCodigo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!DocFecha = Format(gFechaServidor, sqlFormatoFH)
    RsAux!DocTipo = TipoDocumento.ReciboDePago
    RsAux!DocSerie = Serie
    RsAux!DocNumero = Numero
    RsAux!DocCliente = Val(lCliente.Tag)
    RsAux!DocMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    RsAux!DocTotal = CCur(tImporte.Text)
    RsAux!DocIva = 0
    RsAux!DocSucursal = paCodigoDeSucursal
    RsAux!DocUsuario = paCodigoDeUsuario
    RsAux!DocFModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsAux!DocAnulado = 0
    
    Dim sMemo As String
    If cCuenta.ListIndex > -1 Then
        If cCuenta.ItemData(cCuenta.ListIndex) = Cuenta.Colectivo Then
            sMemo = "Colectivo (Nº " & Val(tTitulo.Tag) & "): " & Trim(tTitulo.Text)
        ElseIf cCuenta.ItemData(cCuenta.ListIndex) = Cuenta.Personal Then
            sMemo = "Cuenta Personal: "
            If txtCliente2.Cliente.Codigo > 0 And txtCliente2.Text <> "" Then sMemo = sMemo & txtCliente2.Text & " "
            sMemo = sMemo & Trim(lCCliente.Caption)
        End If
        RsAux("DocComentario") = sMemo
    End If
    
    RsAux.Update: RsAux.Close
    
    Cons = "SELECT MAX(DocCodigo) From Documento" _
            & " WHERE DocTipo = " & TipoDocumento.ReciboDePago _
            & " AND DocSerie = '" & Serie & "'" & " AND DocNumero = " & Numero
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aDocumentoRecibo = RsAux(0)
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------
    
    cBase.Execute "EXEC prg_PosInsertoDocumentosATickets '" & aDocumentoRecibo & "', " & oCnfgPrint.ImpresoraTickets
    
    'Cargo campos CuentaDocumento-------------------------------------------------------------------------------
'    Cons = "Select * from CuentaDocumento Where CDoTipo = 0 And CDoIdTipo = 0 And CDoIdDocumento = 0"
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    RsAux.AddNew
'    RsAux!CDoTipo = cCuenta.ItemData(cCuenta.ListIndex)
'
'    If cCuenta.ItemData(cCuenta.ListIndex) = Cuenta.Colectivo Then
'        RsAux!CDoIdTipo = Val(tTitulo.Tag)
'    Else
'        RsAux!CDoIdTipo = Val(lCCliente.Tag)
'    End If
'    RsAux!CDoIdDocumento = aDocumentoRecibo
'    If Val(tArticulo.Tag) <> 0 Then RsAux!CDoIdArticulo = Val(tArticulo.Tag)
'    RsAux.Update: RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------
    
    Dim oApoCta As New clsAporteACuenta
    oApoCta.Documento = aDocumentoRecibo
    oApoCta.Fecha = Now
    If cCuenta.ItemData(cCuenta.ListIndex) = Cuenta.Colectivo Then
        oApoCta.idCuenta = Val(tTitulo.Tag)
    Else
        oApoCta.idCuenta = Val(lCCliente.Tag)
    End If
    oApoCta.Importe = 0 'CCur(tImporte.Text)
    oApoCta.MovimientoAporte = Aporte
    oApoCta.TipoCuenta = cCuenta.ItemData(cCuenta.ListIndex)
    oApoCta.InsertarAsignacion cBase
    
    Dim oTransacciones As New clsCobrarConQuePaga
    Dim colDocsAImprimir As New Collection
    If Not oConQueSeCobra Is Nothing Then
        Set oTransacciones.Conexion = cBase
        oConQueSeCobra.DocumentoQueSalda = aDocumentoRecibo
        oConQueSeCobra.TipoDocQueSalda = ContadooCuota
        oConQueSeCobra.Cliente = Val(lCliente.Tag)
        oConQueSeCobra.Sucursal = paCodigoDeSucursal
        oConQueSeCobra.Terminal = paCodigoDeTerminal
        Set colDocsAImprimir = oTransacciones.GrabarAsignaciones(oConQueSeCobra)
    End If
    
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    On Error GoTo errPrint
    'AccionImprimir aDocumentoRecibo
    If Not oConQueSeCobra Is Nothing Then
        If Not colDocsAImprimir Is Nothing Then
            If colDocsAImprimir.Count > 0 Then
                ImprimoDocumentosDeAportes colDocsAImprimir
            End If
        End If
    End If
    AccionLimpiar dejarCuenta:=True
    Screen.MousePointer = 0
    Exit Sub
    
errPrint:
    clsGeneral.OcurrioError "Error al imprimir", Err.Description, "Imprimir aportes"
    AccionLimpiar dejarCuenta:=True
    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub ImprimoDocumentosDeAportes(ByVal colsDocs As Collection)
Dim oDocs As clsDocAImprimir
Dim oPrint As New clsImpresionDeDocumentos
    
    Set oPrint.Conexion = cBase
    oPrint.PathReportes = gPathListados
    oPrint.NombreBaseDatos = miConexion.RetornoPropiedad(False, False, False, True)
    For Each oDocs In colsDocs
        Dim oCnfg As New clsConfigImpresora
        If oDocs.TipoDocumento = Imp_MovCaja_Aporte Then
'            oCnfg.Impresora = paIReciboN
'            oCnfg.Bandeja = paIReciboB
'            Set oPrint.DondeImprimo = oCnfg
'            oPrint.ImprimoAporteACuenta oDocs.IDDocumento, paDRecibo, 1, Val(lCliente.Tag), "Saldo a favor, pago de aporte."
            cBase.Execute "EXEC prg_PosInsertoDocumentosATickets '" & oDocs.IDDocumento & "', " & oCnfgPrint.ImpresoraTickets
        ElseIf oDocs.TipoDocumento = Imp_MovimientoCaja Then
            If oCnfgPrintSalidaCaja.Opcion = 0 Then
                oCnfg.Impresora = paIRemitoN
                oCnfg.Bandeja = paIRemitoB
            Else
                oCnfg.Impresora = oCnfgPrintSalidaCaja.ImpresoraTickets
            End If
            Set oPrint.DondeImprimo = oCnfg
            If oCnfgPrintSalidaCaja.Opcion = 0 Then
                oPrint.ImprimoSalidaCaja_Crystal oDocs.IDDocumento, "Señas Recibas", "$", paCodigoDeUsuario, miConexion.NombreTerminal
            Else
                oPrint.ImprimoSalidaCajaTicket oDocs.IDDocumento, paNombreSucursal, miConexion.UsuarioLogueado(False, True), "Señas recibidas"
            End If
        End If
    Next
    
End Sub

Private Function ValidoCampos() As Boolean

    On Error GoTo errValidar
    ValidoCampos = False
    
    If oCnfgPrint.ImpresoraTickets = 0 Then
        MsgBox "La tickeadora no está configurada.", vbCritical, "ATENCIÓN"
        Exit Function
    End If
    
    If Val(lCliente.Tag) = 0 Then
        MsgBox "Debe ingresar el cliente que realiza el aporte a la cuenta.", vbExclamation, "ATENCIÓN"
        txtCliente.SetFocus: Exit Function
    End If
        
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda del importe a depositar en la cuenta.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    If Not IsNumeric(tImporte.Text) Then
        MsgBox "Debe ingresar el importe a depositar en la cuenta.", vbExclamation, "ATENCIÓN"
        Foco tImporte: Exit Function
    End If
    
    If Val(lCliente.Tag) = 0 Then
        MsgBox "Debe ingresar el cliente que realiza el aporte a la cuenta.", vbExclamation, "ATENCIÓN"
        txtCliente.SetFocus: Exit Function
    End If
    
    If cCuenta.ItemData(cCuenta.ListIndex) = Cuenta.Colectivo Then
        If Val(tTitulo.Tag) = 0 Then
            MsgBox "Debe seleccionar el colectivo al que se realiza el aporte.", vbExclamation, "ATENCIÓN"
            Foco tTitulo: Exit Function
        End If
    End If
    If cCuenta.ItemData(cCuenta.ListIndex) = Cuenta.Personal Then
        If Val(lCCliente.Tag) = 0 Then
            MsgBox "Debe ingresar el cliente propietario de la cuenta.", vbExclamation, "ATENCIÓN"
            txtCliente2.SetFocus: Exit Function
        End If
    End If
    
    ValidoCampos = True
    Exit Function

errValidar:
    clsGeneral.OcurrioError "Ocurrió un error al validar los datos ingresados.", Err.Description
End Function

Private Sub AccionLimpiar(Optional dejarCuenta As Boolean = False)

    Set oConQueSeCobra = Nothing
    hlConQueSeCobra.Caption = "¿Con que se cobra?"

    txtCliente.Text = ""
    lCliente.Caption = "": lDireccion.Caption = ""
    tImporte.Text = ""
    
    If Not dejarCuenta Then
        tTitulo.Text = ""
        txtCliente2.Text = ""
        lCCliente.Caption = ""
    End If
    
    tArticulo.Text = ""
    
    txtCliente.SetFocus
    lContado.Caption = "": lAportado.Caption = "": lDiferencia.Caption = ""
    tComentario.Text = ""
End Sub

Private Sub CargoDatosCuenta(idTipo As Integer, idCuenta As Long)
On Error GoTo errCargar

Dim bColectivo As Boolean
    
    If idTipo = Cuenta.Colectivo Then bColectivo = True Else bColectivo = False
    
    If bColectivo Then
        Cons = "Select * from Colectivo Where ColCodigo = " & idCuenta
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            tTitulo.Text = Trim(RsAux!ColNombre)
            tTitulo.Tag = RsAux!ColCodigo
            
            If Not IsNull(RsAux!ColComentario) Then tComentario.Text = Trim(RsAux!ColComentario) Else tComentario.Text = ""
        End If
        RsAux.Close
    
    Else
        Dim mTipo As Integer
        
        Cons = "Select * from Cliente Where CliCodigo = " & idCuenta
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then mTipo = RsAux!CliTipo
        RsAux.Close
        
        txtCliente2.CargarControl idCuenta
        
'        If mTipo = TipoCliente.Cliente Then
'
'            BuscoClienteID idCuenta, 0, False, True
'        Else
'            BuscoClienteID 0, idCuenta, False, True
'        End If
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Error al cargar los parámetros de la cuenta.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosAporta(idCliente As Long)
On Error GoTo errCargar

    Dim mTipo As Integer
    Me.Show
    
    Cons = "Select * from Cliente Where CliCodigo = " & idCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then mTipo = RsAux!CliTipo
    RsAux.Close
    txtCliente.CargarControl idCliente
'    If mTipo = TipoCliente.Cliente Then
'        BuscoClienteID idCliente, 0, True, False
'    Else
'        BuscoClienteID 0, idCliente, True, False
'    End If
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Error al cargar los parámetros del aporte.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub txtCliente_BorroCliente()
    lCliente.Tag = 0
    Set oConQueSeCobra = Nothing
    hlConQueSeCobra.Caption = "¿Con que se cobra?"
End Sub

Private Sub txtCliente_CambioTipoDocumento()
    SeteoInfoDocumentoCliente
End Sub

Private Sub SeteoInfoDocumentoCliente()
    lblInfoCliente.ForeColor = vbBlack
    Select Case txtCliente.DocumentoCliente
        Case DC_CI
            lblInfoCliente.Caption = "C.I.:"
        Case DC_RUT
            lblInfoCliente.Caption = "R.U.T.:"
        Case Else
            If txtCliente.Cliente.TipoDocumento.Nombre = "" Then
                lblInfoCliente.Caption = "Otro:"
            Else
                lblInfoCliente.Caption = txtCliente.Cliente.TipoDocumento.Abreviacion
            End If
            lblInfoCliente.ForeColor = &HFF&
    End Select
End Sub

Private Sub txtCliente_PresionoEnter()
    If tImporte.Enabled Then tImporte.SetFocus
End Sub

Private Sub txtCliente_SeleccionoCliente()

    lCliente.Tag = txtCliente.Cliente.Codigo
    lCliente.Caption = " " & txtCliente.Cliente.Nombre
    lDireccion.Caption = ""
    lDireccion.Caption = txtCliente.Cliente.Direccion
    ValidoDireccionEMail txtCliente.Cliente.Codigo
End Sub

Private Sub SeteoInfoDocumentoCliente2()
    lblInfoPersona2.ForeColor = vbBlack
    Select Case txtCliente2.DocumentoCliente
        Case DC_CI
            lblInfoPersona2.Caption = "C.I.:"
        Case DC_RUT
            lblInfoPersona2.Caption = "R.U.T.:"
        Case Else
            If txtCliente.Cliente.TipoDocumento.Nombre = "" Then
                lblInfoPersona2.Caption = "Otro:"
            Else
                lblInfoPersona2.Caption = txtCliente.Cliente.TipoDocumento.Abreviacion
            End If
            lblInfoPersona2.ForeColor = &HFF&
    End Select
End Sub

Private Sub txtCliente2_CambioTipoDocumento()
    SeteoInfoDocumentoCliente2
End Sub

Private Sub txtCliente2_PresionoEnter()
    If txtCliente2.Cliente.Codigo > 0 Then
        Foco tArticulo
    End If
End Sub

Private Sub txtCliente2_SeleccionoCliente()
    lCCliente.Tag = txtCliente2.Cliente.Codigo
    lCCliente.Caption = " " & txtCliente2.Cliente.Nombre
    ValidoDireccionEMail txtCliente2.Cliente.Codigo
End Sub
