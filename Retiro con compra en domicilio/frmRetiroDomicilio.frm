VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{190700F0-8894-461B-B9F5-5E731283F4E1}#1.1#0"; "orHiperlink.ocx"
Begin VB.Form frmRetiroDomicilio 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retiro en domicilio"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRetiroDomicilio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicPasoNota 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   0
      ScaleHeight     =   3630
      ScaleWidth      =   8490
      TabIndex        =   11
      Top             =   600
      Width           =   8520
      Begin VB.TextBox txtDocumento 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   12
         Top             =   120
         Width           =   1815
      End
      Begin prjHiperLink.orHiperLink hliCliente 
         Height          =   255
         Left            =   960
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         BackColor       =   14737632
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorOver   =   16711680
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
      Begin prjHiperLink.orHiperLink hliDocumento 
         Height          =   285
         Left            =   4080
         Top             =   120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   503
         BackColor       =   14737632
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorOver   =   16711680
         MouseIcon       =   "frmRetiroDomicilio.frx":2B05
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
      Begin VSFlex8LCtl.VSFlexGrid lstArticulos 
         Height          =   2415
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   8175
         _cx             =   14420
         _cy             =   4260
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   285
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
         AutoSearchDelay =   2
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
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Artículos A RETIRAR"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblDocumento 
         BackStyle       =   0  'Transparent
         Caption         =   "&Documento, C.I./R.U.C."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   150
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtMemo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   80
      TabIndex        =   6
      Top             =   4920
      Width           =   5535
   End
   Begin VB.PictureBox picTitulo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   10815
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   10845
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   120
         Picture         =   "frmRetiroDomicilio.frx":2E1F
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   10
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Retiro en domicilio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   10815
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5355
      Width           =   10845
      Begin VB.CommandButton butCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7320
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton butAceptar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   6240
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblAyuda 
         BackStyle       =   0  'Transparent
         Caption         =   "Devolución de mercadería ayuda rápida de lo que tienen que realizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   450
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5340
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   7
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentario:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
End
Attribute VB_Name = "frmRetiroDomicilio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'book vox-076spk1205000323 el de mi nootb.
Option Explicit

Public Enum Accion
    Informacion = 1         'No toma accion es un comentario +
    Alerta = 2             'Activa la pantalla de comentarios Todas
    Cuota = 3              'Activa en Cobranza, Decision, Visualizacion
    Decision = 4            'Activa en Decision
End Enum

Public Enum TipoAccionEntrada
    TAE_Devolucion = 1
    TAE_Cambio = 2
End Enum

Private EmpresaEmisora As clsClienteCFE
Private TasaBasica As Currency, TasaMinima As Currency

Private Sub CargoValoresIVA()
Dim RsIva As rdoResultset
Dim sQy As String
    sQy = "SELECT IvaCodigo, IvaPorcentaje FROM TipoIva WHERE IvaCodigo IN (1,2)"
    Set RsIva = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsIva.EOF
        Select Case RsIva("IvaCodigo")
            Case 1: TasaBasica = RsIva("IvaPorcentaje")
            Case 2: TasaMinima = RsIva("IvaPorcentaje")
        End Select
        RsIva.MoveNext
    Loop
    RsIva.Close
End Sub

Private Sub EmitirCFERepetitivo(ByVal Documento As Long)
Dim resM As VbMsgBoxResult
Dim sPaso As String

    resM = vbYes
    sPaso = FirmarUnDocumento(Documento)
    Do While sPaso <> ""
        resM = MsgBox("ATENCIÓN no se firmó el documento" & vbCrLf & vbCrLf & "Presione SI para reintentar" & vbCrLf & " Presione NO para abandonar ", vbExclamation + vbYesNo, "ATENCIÓN")
        If resM = vbNo Then Exit Do
        sPaso = FirmarUnDocumento(Documento)
    Loop
    
End Sub

Private Function FirmarUnDocumento(ByVal Documento As Long) As String
On Error GoTo errEC
    
    If (TasaBasica = 0) Then CargoValoresIVA
    
    FirmarUnDocumento = vbNullString
    With New clsCGSAEFactura
        .URLAFirmar = ParametrosSist.ObtenerValorParametro(URLFirmaEFactura).Texto
        .ImporteConInfoDeCliente = ParametrosSist.ObtenerValorParametro(efactImporteDatosCliente).Valor
        .TasaBasica = TasaBasica
        .TasaMinima = TasaMinima
        Set .Connect = cBase
        Dim sResult As String
        sResult = .FirmarUnDocumento(Documento)
        If UCase(sResult) <> "TRUE" Then FirmarUnDocumento = sResult
    End With
    Exit Function
    
errEC:
    FirmarUnDocumento = "Error en firma: " & Err.Description
End Function

Private Function GrabarRemitosEnBD(ByVal cliente As clsClienteCFE) As Boolean
   
   GrabarRemitosEnBD = False
   Dim sIDsEnvios As String
   sIDsEnvios = NuevoEnvio
   If sIDsEnvios = vbNullString Or sIDsEnvios = "0" Then
        MsgBox "Sin envío no se puede realizar el cambio.", vbExclamation, "ATENCIÓN"
        Exit Function
   End If
   
   On Error GoTo errBT
   cBase.BeginTrans
    
    Dim caeG As New clsCAEGenerador
    Dim CAE As clsCAEDocumento
    Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI)
    
    Dim docRemitoRet As clsDocumentoCGSA
    Set docRemitoRet = New clsDocumentoCGSA
        
    Dim oRenDoc As clsDocumentoRenglon
    Dim i As Integer
    For i = 1 To lstArticulos.Rows - 1
        Set oRenDoc = New clsDocumentoRenglon
        
        oRenDoc.Cantidad = Val(lstArticulos.Cell(flexcpText, i, 0))
        oRenDoc.CantidadARetirar = oRenDoc.Cantidad
        oRenDoc.EstadoMercaderia = paEstadoArticuloEntrega
        oRenDoc.IVA = 0
        oRenDoc.Precio = 0
        oRenDoc.Articulo.ID = lstArticulos.Cell(flexcpData, i, 0)
        
        docRemitoRet.AddRenglon oRenDoc
    Next
        
    With docRemitoRet
        Set .cliente = cliente
        .Emision = gFechaServidor
        .Tipo = TD_RemitoRetiro
        .Numero = CAE.Numero
        .Serie = CAE.Serie
        .Moneda.Codigo = 1
        .Total = 0
        .IVA = 0
        .Sucursal = paCodigoDeSucursal
        .Digitador = CInt(txtUser.Tag)
        .Comentario = "Retiro en domicilio en envío: " & sIDsEnvios & ". " & txtMemo.Text
        .Vendedor = Val(txtUser.Tag)
    End With
    Set docRemitoRet.Conexion = cBase
    docRemitoRet.Codigo = docRemitoRet.InsertoDocumentoBD(0)
    
    Dim rsEnv As rdoResultset
    Dim sQy As String
    sQy = "SELECT * FROM Envio WHERE EnvCodigo IN (" & sIDsEnvios & ")"
    Set rsEnv = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    Do While Not rsEnv.EOF
    
        cBase.Execute "INSERT INTO EnviosRemitos (EReEnvio, EReRemito, EReFactura) VALUES (" & rsEnv("EnvCodigo") & ", " & docRemitoRet.Codigo & ", " & Val(hliDocumento.Tag) & ")"
    
        rsEnv.Edit
        rsEnv("EnvFormaPago") = 2
        rsEnv("EnvDocumento") = docRemitoRet.Codigo
        rsEnv("EnvCliente") = cliente.Codigo
        rsEnv("EnvUsuario") = CInt(txtUser.Tag)
        rsEnv.Update
        
        rsEnv.MoveNext
        
    Loop
    rsEnv.Close
    
    cBase.CommitTrans
    GrabarRemitosEnBD = True
    
    On Error GoTo errYaGrabe
    EmitirCFERepetitivo docRemitoRet.Codigo
    Exit Function
    
errBT:
    clsGeneral.OcurrioError "Error inesperado al inicializar la transacción.", Err.Description, "Grabar"
    Screen.MousePointer = 0
    Exit Function
    
    
errYaGrabe:
    clsGeneral.OcurrioError "Error inesperado al finalizar el evento grabar.", Err.Description, "Restauración de formulario"
    Screen.MousePointer = 0
    Exit Function
    
ErrResumo:
    Resume ErrRelajo
    
ErrRelajo:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al emitir los remitos de cambio.", Err.Description, "Grabar"
    Exit Function

End Function

Private Function PedirSuceso(ByRef Usuario As Integer, ByRef autoriza As Integer, ByRef defensa As String) As Boolean
    
    Usuario = 0
    
    Dim objSuceso As New clsSuceso
    With objSuceso
        .TipoSuceso = 11 ' TipoSuceso.DiferenciaDeArticulos
        .ActivoFormulario Val(txtUser.Tag), "Cliente con Cuotas Atrasadas", cBase
        Usuario = .RetornoValor(Usuario:=True)
        If Usuario > 0 Then
            defensa = .RetornoValor(defensa:=True)
            If .autoriza > 0 Then autoriza = .autoriza
        End If
    End With
    Set objSuceso = Nothing
    Me.Refresh
    PedirSuceso = (Usuario > 0)

End Function

Private Function BuscarCuotasVencidasCliente(ByVal lCliente As Long, ByVal sCliente As String, Optional bShowMsg As Boolean) As Boolean
'---------------------------------------------------
'Retorno True si lleva suceso
'---------------------------------------------------
On Error GoTo errCV
Dim rsC As rdoResultset
Dim iMaxAtraso As Integer

    BuscarCuotasVencidasCliente = False
    
    'Condición para no consultar que el cliente sea de la esta lista.
    If InStr(1, "," & paClienteNoVtoCta & ",", "," & lCliente & ",") > 0 Then
        Exit Function
    End If
    '.......................................................................................
    
    iMaxAtraso = 0
    Cons = "Select Min(CreProximoVto) " & _
                " From Documento (index = iClienteTipo), Credito" & _
                " Where DocCliente = " & lCliente & _
                " And DocCodigo = CreFactura " & _
                " And DocTipo = 2" & _
                " And DocAnulado = 0  And CreSaldoFactura > 0 "
    
    Set rsC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsC.EOF Then
        If Not IsNull(rsC(0)) Then iMaxAtraso = DateDiff("d", rsC(0), Now)
    End If
    rsC.Close
    
    Select Case iMaxAtraso
        Case Is > 20
                If bShowMsg Then MsgBox "El cliente '" & sCliente & "' no está al día." & vbCrLf & _
                            "Tiene coutas vencidas con más de 20 días." & vbCrLf & vbCrLf & _
                            "Consulte antes de realizar el ingreso del artículo.", vbExclamation, "Cliente con Ctas. Vencidas"
                BuscarCuotasVencidasCliente = True
                
        Case Is > 5
                If bShowMsg Then MsgBox "El cliente '" & sCliente & "' no está al día. Tiene coutas vencidas." & vbCrLf & _
                            "Consulte antes de realizar el ingreso del artículo.", vbExclamation, "Cliente con Ctas. Vencidas"
    End Select
    Exit Function
    
errCV:
    clsGeneral.OcurrioError "Error al buscar las cuotas vencidas.", Err.Description
End Function

Private Function BuscarArticulosParaDevolverDelDocumento() As Boolean
On Error GoTo errBADD
Dim rsArt As rdoResultset
    
    Dim oRenglon As clsArticuloRenglones
    Cons = "EXEC prg_IngresoMercaderiaCliente_ArticulosPosibles " & Val(hliDocumento.Tag)
    Set rsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsArt.EOF
        
'        Set oRenglon = New clsArticuloRenglones
'        oRenglon.Articulo = rsArt(0)
'        oRenglon.Cantidad = rsArt(1)
'        colArtsDoc.Add oRenglon
        
        With lstArticulos
            .AddItem 0
            .Cell(flexcpText, .Rows - 1, 1) = rsArt(1)
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsArt("ArtCodigo"), "#,##0,000") & " " & Trim(rsArt("ArtNombre"))
            
            .Cell(flexcpData, .Rows - 1, 0) = CStr(rsArt(0))
        End With
        
        rsArt.MoveNext
    Loop
    rsArt.Close
    BuscarArticulosParaDevolverDelDocumento = True
    Exit Function
    
errBADD:
    clsGeneral.OcurrioError "Error al buscar los artículos del documento.", Err.Description, "Artículos del documento"
End Function

Public Sub MostrarAyuda(Optional msg As String = "")
    If msg = "" Then
        Select Case Me.ActiveControl.Name
            Case txtDocumento.Name
                lblAyuda.Caption = "Buscar por C.I./R.U.C. o por código de barras/serie-número del Documento (F12 Vis. Ope.)"
            Case txtMemo.Name
                lblAyuda.Caption = "Ingrese un comentario."
            Case txtUser.Name
                lblAyuda.Caption = "Ingrese su dígito de usuario y presione Enter para poder grabar"
        End Select
    Else
        lblAyuda.Caption = msg
    End If
End Sub

Private Sub EstadoControlesIngreso(ByVal habilitados As Boolean)
Dim lcolor As Long

    lcolor = IIf(habilitados, vbWindowBackground, vbButtonFace)
    
    With txtMemo
        .Enabled = habilitados
        .BackColor = lcolor
    End With
    
    With txtUser
        .Enabled = habilitados
        .BackColor = lcolor
    End With
    
    lstArticulos.Enabled = habilitados
    butAceptar.Enabled = habilitados
    
End Sub

Private Sub LimpiarControlesArticulos()
    
    txtMemo.Text = ""
    txtUser.Text = "": txtUser.Tag = 0
    lstArticulos.Rows = 1
    
End Sub

Private Sub LimpiarControlesDocumento()
    hliCliente.Caption = ""
    hliDocumento.Caption = ""
    hliCliente.Tag = ""
    hliDocumento.Tag = ""
'    Set colArtsDoc = Nothing
'    Set colArtsDoc = New Collection
End Sub

Private Sub butAceptar_Click()
    
    On Error GoTo errValidar
    If Val(txtUser.Tag) = 0 Then
        MsgBox "Ingrese su dígito de usuario.", vbExclamation, "Validación"
        txtUser.SetFocus
        Exit Sub
    End If
    
    Dim bHay As Boolean
    bHay = False
    Dim i As Long
    'Veo que hayan artículos para retirar.
    For i = 1 To lstArticulos.Rows - 1
        If Val(lstArticulos.Cell(flexcpText, i, 0)) > 0 Then
            bHay = True
            Exit For
        End If
    Next
    
    If Not bHay Then
        MsgBox "No hay artículos para devolver.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    Dim oInfoCAE As New clsCAEGenerador
    If Not oInfoCAE.SucursalTieneCae(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI) Then
        MsgBox "No hay un CAE disponible para emitir el eRemito, por favor comuníquese con administración." & vbCrLf & vbCrLf & "No podrá recepcionar", vbCritical, "eFactura"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    
    
    'Si es un remito recepción tengo que obligar a ingresar todos los artículos.
    If MsgBox("¿Confirma almacenar los datos ingresados?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        
        If (EmpresaEmisora Is Nothing) Then
            Set EmpresaEmisora = New clsClienteCFE
            EmpresaEmisora.CargoClienteCarlosGutierrez paCodigoDeSucursal
        End If

        FechaDelServidor
        Dim oCli As New clsClienteCFE
        oCli.CargoInformacionCliente2 cBase, hliCliente.Tag, True, True
        If GrabarRemitosEnBD(oCli) Then
            txtDocumento.Text = vbNullString
        End If
    End If
    Exit Sub
    
errValidar:
    clsGeneral.OcurrioError "Error al validar para grabar", Err.Description
    Exit Sub

End Sub

Private Function NuevoEnvio() As String
Dim oEnvAux As New clsEnvioAuxiliar

    NuevoEnvio = vbNullString
    Dim i As Integer
    Dim iQTotal As Integer
    iQTotal = 0
    Dim oArt As clsArtEnvAux
    For i = 1 To lstArticulos.Rows - 1
        If Val(lstArticulos.Cell(flexcpText, i, 0)) > 0 Then
            Set oArt = New clsArtEnvAux
            oArt.Articulo = lstArticulos.Cell(flexcpData, i, 0)
            oArt.Cantidad = Val(lstArticulos.Cell(flexcpText, i, 0))
            iQTotal = iQTotal + oArt.Cantidad
            oEnvAux.Articulos.Add oArt
        End If
    Next
    
    If oEnvAux.GrabarNuevoAuxiliar(cBase) Then
        Dim objEnvio As New clsEnvio
        objEnvio.NuevoEnvio cBase, "", oEnvAux.ID, Val(hliCliente.Tag), 1, 1
        Me.Refresh
        NuevoEnvio = objEnvio.RetornoEnvios
        If Len(Trim(NuevoEnvio)) > 1 Then
            
            If InStr(1, NuevoEnvio, ",", vbTextCompare) > 0 Then
                MsgBox "Sólo puede cambiar la mercadería en un único envío.", vbExclamation, "ATENCIÓN"
                BorroEnvios NuevoEnvio
                NuevoEnvio = ""
                Exit Function
            End If
            
            Cons = "SELECT SUM(REvAEntregar) FROM RenglonEnvio WHERE REvEnvio IN (" & NuevoEnvio & ")"
            Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then
                If rsAux(0) <> iQTotal Then
                    rsAux.Close
                    MsgBox "Se debe enviar toda la mercadería.", vbExclamation, "ERROR"
                    BorroEnvios NuevoEnvio
                    NuevoEnvio = ""
                    Exit Function
                End If
            End If
            rsAux.Close
        End If
    End If
    Set objEnvio = Nothing
    Exit Function
    
errEnvio:
    NuevoEnvio = vbNullString
    clsGeneral.OcurrioError "Error al invocar el envío", Err.Description
    
End Function


Private Sub BorroEnvios(ByVal strCodigoEnvio As String)

On Error GoTo ErrBE
Dim CodEnvios As String
Dim lngCodEnvio As Long

    If strCodigoEnvio = "0" Then strCodigoEnvio = ""
    Do While strCodigoEnvio <> ""
    
        If InStr(1, strCodigoEnvio, ",") > 0 Then
            CodEnvios = Left(strCodigoEnvio, InStr(1, strCodigoEnvio, ","))
            lngCodEnvio = CLng(Left(CodEnvios, InStr(1, CodEnvios, ",") - 1))
            strCodigoEnvio = Right(strCodigoEnvio, Len(strCodigoEnvio) - InStr(1, strCodigoEnvio, ","))
        Else
            lngCodEnvio = CLng(strCodigoEnvio)
            strCodigoEnvio = ""
        End If
        
        cBase.BeginTrans
        On Error GoTo ErrResumo
        
        Dim idEVC As Long
        Cons = "SELECT EVCID, IsNull(count(*), 0) From EnvioVaCon WHERE EVCEnvio = " & lngCodEnvio & " GROUP BY EVCID"
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            If rsAux(1) > 2 Then
                cBase.Execute ("DELETE EnvioVaCon WHERE EVCID = " & rsAux(0) & " AND EVCEnvio = " & lngCodEnvio)
            Else
                cBase.Execute ("DELETE EnvioVaCon WHERE EVCID = " & rsAux(0))
            End If
        End If
        rsAux.Close
        
        'Borro los renglones del envío.
        Cons = "DELETE RenglonEnvio Where REvEnvio = " & lngCodEnvio
        cBase.Execute (Cons)
        
        'Borro el envío
        Cons = "Select EnvDireccion From Envio Where EnvCodigo = " & lngCodEnvio
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not IsNull(rsAux("EnvDireccion")) Then lngCodEnvio = rsAux!EnvDireccion Else lngCodEnvio = 0
        rsAux.Delete
        rsAux.Close
        
        'Borro la dirección.
        If lngCodEnvio > 0 Then
            Cons = "DELETE Direccion Where DirCodigo = " & lngCodEnvio
            cBase.Execute (Cons)
        End If
        
        cBase.CommitTrans
    Loop
    Exit Sub
    
ErrBE:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error inesperado al intentar la transacción."
    
ErrResumo:
    Resume Relajo
    
Relajo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "No se pudo eliminar algún envío."

End Sub



Private Sub butCancelar_Click()
    
    'cancelo el ingreso.
    LimpiarControlesDocumento
    LimpiarControlesArticulos
    EstadoControlesIngreso False
    txtDocumento.Text = ""
    txtDocumento.SetFocus
    
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
Dim sPaso As String

    sPaso = "Controles"
    EstadoControlesIngreso False
    LimpiarControlesArticulos
    LimpiarControlesDocumento
    
    With lstArticulos
        .Rows = 1: .Cols = 1
        .RowHeight(0) = 315
        .RowHeightMin = 285
        .FixedCols = 0
        .FormatString = ">Devuelve|>Compró|<Artículo"
        .ColWidth(0) = 1000
        .ColWidth(1) = 1000
        .ColWidth(2) = 2000
        .ExtendLastCol = True
    End With
    Exit Sub
    
errLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario", sPaso & vbCrLf & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    cBase.Close
    Set cBase = Nothing
    Set clsGeneral = Nothing
    
    Dim objFnc As New clsFncGlobales
    objFnc.SetPositionForm Me
    Set objFnc = Nothing
    
End Sub

Private Sub hliDocumento_Click()
On Error Resume Next
    If Val(hliDocumento.Tag) > 0 Then
        Shell App.Path & "\detalle de factura.exe " & hliDocumento.Tag, vbNormalFocus
    End If
End Sub

Private Sub lstArticulos_GotFocus()
    lblAyuda.Caption = "Artículos que puede devolver el cliente."
End Sub

Private Sub lstArticulos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If lstArticulos.Row < 1 Then Exit Sub
    
    Dim iQ As Byte
    Select Case KeyCode
        Case vbKeyDelete
            'Elimino de la grilla y de la colección.
            Dim oArt As clsRenglonIngreso
            
            With lstArticulos
                .Cell(flexcpText, lstArticulos.Row, 0) = "0"
            End With
        Case vbKeyAdd
            If Val(lstArticulos.Cell(flexcpText, lstArticulos.Row, 0)) < Val(lstArticulos.Cell(flexcpText, lstArticulos.Row, 1)) Then
                lstArticulos.Cell(flexcpText, lstArticulos.Row, 0) = Val(lstArticulos.Cell(flexcpText, lstArticulos.Row, 0)) + 1
            End If
        
        Case vbKeySubtract
            If Val(lstArticulos.Cell(flexcpText, lstArticulos.Row, 0)) > 0 Then
                lstArticulos.Cell(flexcpText, lstArticulos.Row, 0) = Val(lstArticulos.Cell(flexcpText, lstArticulos.Row, 0)) - 1
            End If
    End Select
End Sub

Private Sub lstArticulos_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then txtMemo.SetFocus
End Sub

Private Sub lstArticulos_LostFocus()
    lblAyuda.Caption = ""
End Sub

Private Sub PicPaso1_Click()

End Sub

Private Sub txtDocumento_Change()
    If Val(hliDocumento.Tag) > 0 Or Val(hliCliente.Tag) > 0 Then
        LimpiarControlesDocumento
        LimpiarControlesArticulos
        EstadoControlesIngreso False
    End If
End Sub

Private Sub txtDocumento_GotFocus()
    With txtDocumento
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    MostrarAyuda
End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF12
            Shell App.Path & "\voperaciones.exe " & Val(hliCliente.Tag)
    End Select
End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And Trim(txtDocumento.Text) <> "" Then
        
        If Val(hliDocumento.Tag) > 0 Or Val(hliCliente.Tag) > 0 Then
            If (lstArticulos.Enabled) Then lstArticulos.SetFocus
            Exit Sub
        End If
        
        Dim objHelp As clsListadeAyuda
        On Error GoTo errBD
        Screen.MousePointer = 11
        
        
        Cons = "EXEC prg_BuscarCliente 1, '', '', '', '', '', '" & txtDocumento.Text & "', 0, 0, '', '', 16"
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Not rsAux.EOF Then
            
            If Not IsNull(rsAux("CliCodigo")) Then
                
                If rsAux("CliCodigo") > 0 Then
                    
                    If rsAux("DocCodigo") > 0 Then
                        hliDocumento.Tag = rsAux("DocCodigo")
                        hliDocumento.Caption = rsAux("Documento")
                        lblDocumento.Tag = rsAux("DocFModificacion")
                    End If
                    
                    hliCliente.Caption = "(" & RTrim(rsAux("C.I./R.U.C.")) & ") " & rsAux("Cliente")
                    hliCliente.Tag = rsAux("CliCodigo")
                    rsAux.MoveNext
                    If Not rsAux.EOF Then
                        rsAux.Close
                        
                        hliDocumento.Tag = ""
                        hliDocumento.Caption = ""
                        hliCliente.Caption = ""
                        hliCliente.Tag = ""
                        lblDocumento.Tag = ""
                                                
                        'Abro lista de ayuda.
                        Set objHelp = New clsListadeAyuda
                        If objHelp.ActivarAyuda(cBase, Cons, 5000, 3, "Búsqueda") > 0 Then
                            
                            hliDocumento.Tag = objHelp.RetornoDatoSeleccionado(0)
                            hliCliente.Tag = objHelp.RetornoDatoSeleccionado(1)
                                                        
                        End If
                    Else
                        rsAux.Close
                    End If
                    
                End If
            Else
                rsAux.Close
            End If
            
        Else
        
            MsgBox "No hay resultados para el dato ingresado.", vbInformation, "Búsqueda"
            rsAux.Close
            
        End If
        
        If Val(hliDocumento.Tag) > 0 Then
            
            Dim TipoDoc As Byte
            Cons = "SELECT DocTipo FROM Documento WHERE DocCodigo = " & Val(hliDocumento.Tag)
            Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            TipoDoc = rsAux("DocTipo")
            rsAux.Close
            
            If TipoDoc = 6 Then
            
                Cons = "SELECT RDoDocumento FROM RemitoDocumento WHERE RDoRemito = " & Val(hliDocumento.Tag)
                Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAux.EOF Then hliDocumento.Tag = rsAux(0) Else hliDocumento.Tag = 0
                rsAux.Close
                
                If Val(hliDocumento.Tag) = 0 Then
                    MsgBox "Ud. escaneo un remito que no tiene asociado ningún documento de compra.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
                
            End If
            
            'Cargo a partir de un documento
            If hliDocumento.Caption = "" Then
                
                'es seleccionado por lista de ayuda.
                Cons = "SELECT DocCodigo, CliCodigo, DocFModificacion, dbo.NombreTipoDocumento(100+DocTipo) + ' ' + rTrim(DocSerie)+Convert(VarChar(10), DocNumero) Documento," & _
                    " RTrim(IsNull(CEmFantasia, rTrim(CPeApellido1) + ', ' + RTrim(CPeNombre1))) Cliente, ISNULL(CliCiRUC, '') [C.I./R.U.C.]" & _
                    " FROM Documento INNER JOIN CLiente ON DocCliente = CliCodigo" & _
                    " LEFT OUTER JOIN CPersona ON CPeCliente = CliCodigo" & _
                    " LEFT OUTER JOIN CEmpresa ON CEmCliente = CliCodigo" & _
                    " WHERE DocCodigo = " & hliDocumento.Tag
                    
                Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAux.EOF Then
                    hliDocumento.Tag = rsAux("DocCodigo")
                    hliDocumento.Caption = rsAux("Documento")
                    hliCliente.Caption = "(" & RTrim(rsAux("C.I./R.U.C.")) & ") " & rsAux("Cliente")
                    hliCliente.Tag = rsAux("CliCodigo")
                    lblDocumento.Tag = rsAux("DocFModificacion")
                End If
                rsAux.Close
                
            End If
            
        ElseIf Val(hliCliente.Tag) > 0 Then
            
            'Cargo a partir de un cliente.
            'Si estoy buscando para cambio de productos entonces le pido que seleccione un documento.
            
            'Busco los documentos que tengan artículos entregados.
            Cons = "SELECT DocCodigo, DocFModificacion, DocFecha Fecha, dbo.NombreTipoDocumento(100+DocTipo) + ' ' + rTrim(DocSerie)+'-'+Convert(VarChar(10), DocNumero) Documento, dbo.ListaArticulosDelDocumento(DocCodigo) Artículos" & _
                " FROM ((Documento INNER JOIN Renglon ON RenDocumento = DocCodigo)" & _
                " INNER JOIN Articulo ON ArtID = RenArticulo AND ArtTipo <> 151)" & _
                " WHERE DocTipo IN (1,2,6) AND RenCantidad <> RenARetirar AND DocCliente = " & hliCliente.Tag & " ORDER BY DocFecha DESC"
                
            Set objHelp = New clsListadeAyuda
            If objHelp.ActivarAyuda(cBase, Cons, 5500, 2, "Búsqueda") > 0 Then
                hliDocumento.Tag = objHelp.RetornoDatoSeleccionado(0)
                hliDocumento.Caption = objHelp.RetornoDatoSeleccionado(3)
                lblDocumento.Tag = objHelp.RetornoDatoSeleccionado(1)
            Else
                'no permito seguir con el ingreso.
                hliCliente.Tag = ""
                Set objHelp = Nothing
                Screen.MousePointer = 0
                Exit Sub
            End If
            Set objHelp = Nothing
            
            If hliDocumento.Tag = "" Then
                MsgBox "No hay un documento para poder realizar cambio de productos.", vbInformation, "Cambio de producto"
            End If
            
        End If
    
        'Si tengo cliente o documento asignado
        If Val(hliDocumento.Tag) > 0 Or Val(hliCliente.Tag) > 0 Then
            
            If Val(hliCliente.Tag) > 0 Then
                BuscoComentariosAlerta Val(hliCliente.Tag), True
            End If
            
            If Val(hliDocumento.Tag) > 0 Then
                
                'Busco si el documento posee artículos disponibles para devolver.
                BuscarArticulosParaDevolverDelDocumento
                If lstArticulos.Rows = 1 Then
                    MsgBox "Atención el documento no posee artículos entregados o del mismo no se pueden devolver más artículos.", vbInformation, "ATENCIÓN"
                Else
                    EstadoControlesIngreso True
                    lstArticulos.SetFocus
                End If
                
            End If
            
            If lstArticulos.Enabled Then
                BuscarCuotasVencidasCliente hliCliente.Tag, hliCliente.Caption, True
            End If
            
        End If
        
    End If
    Screen.MousePointer = 0
    Exit Sub
errBD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar.", Err.Description, "Búsqueda"
    
End Sub

Private Sub txtDocumento_LostFocus()
    lblAyuda.Caption = ""
End Sub

Private Sub txtMemo_GotFocus()
    MostrarAyuda
    With txtMemo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtUser.SetFocus
End Sub

Private Sub txtMemo_LostFocus()
lblAyuda.Caption = ""
End Sub

Private Sub txtUser_Change()
    txtUser.Tag = ""
End Sub

Private Sub txtUser_GotFocus()
    MostrarAyuda
    With txtUser
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsNumeric(txtUser.Text) Then
        If Val(txtUser.Tag) = 0 Then
            Dim objFnc As New clsFncGlobales
            txtUser.Tag = objFnc.BuscarUsuario(CInt(txtUser.Text))
            Set objFnc = Nothing
        End If
        
        If Val(txtUser.Tag) > 0 Then
            butAceptar_Click
        End If
        
    End If
End Sub

Private Sub txtUser_LostFocus()
    lblAyuda.Caption = ""
End Sub

Private Function CargoPosibleFactura(ByVal IDArticulo As Long) As Long

    Cons = "SELECT DocCodigo, DocFModificacion, dbo.NombreTipoDocumento(100+DocTipo) + ' ' + rTrim(DocSerie)+Convert(VarChar(10), DocNumero) Documento, DocFecha Fecha, dbo.ListaArticulosDelDocumento(DocCodigo) Artículos" & _
        " FROM ((Documento INNER JOIN Renglon ON RenDocumento = DocCodigo And (RenArticulo = " & IDArticulo & " OR " & IDArticulo & " = 0))" & _
        " INNER JOIN Articulo ON ArtID = RenArticulo AND ArtTipo <> 151)" & _
        " WHERE DocTipo IN (1,2,6) AND RenCantidad <> RenARetirar AND DocCliente = " & hliCliente.Tag
        
    Dim objHelp As New clsListadeAyuda
    objHelp.CerrarSiEsUnico = True
    If objHelp.ActivarAyuda(cBase, Cons, 5000, 2, "Búsqueda") > 0 Then
        hliDocumento.Tag = objHelp.RetornoDatoSeleccionado(0)
        hliDocumento.Caption = objHelp.RetornoDatoSeleccionado(2)
        lblDocumento.Tag = objHelp.RetornoDatoSeleccionado(1)
        If Format(objHelp.RetornoDatoSeleccionado(3), "dd/MM/yyyy") = Date Then
            
        End If
    End If
    Set objHelp = Nothing
    CargoPosibleFactura = Val(hliDocumento.Tag)
    
End Function

Public Sub BuscoComentariosAlerta(idCliente As Long, _
                                                   Optional Alerta As Boolean = False, Optional Cuota As Boolean = False, _
                                                   Optional Decision As Boolean = False, Optional Informacion As Boolean = False)
                                                   
Dim rsCom As rdoResultset
Dim aCom As String
Dim sHay As Boolean

    On Error GoTo errMenu
    Screen.MousePointer = 11
    sHay = False
    'Armo el str con los comentarios a consultar-------------------------------------------------
    If Not Alerta And Not Cuota And Not Decision And Not Informacion Then Exit Sub
    aCom = ""
    If Alerta Then aCom = aCom & Accion.Alerta & ", "
    If Cuota Then aCom = aCom & Accion.Cuota & ", "
    If Decision Then aCom = aCom & Accion.Decision & ", "
    If Informacion Then aCom = aCom & Accion.Informacion & ", "
    aCom = Mid(aCom, 1, Len(aCom) - 2)
    '---------------------------------------------------------------------------------------------------
    
    Cons = "Select * From Comentario, TipoComentario " _
            & " Where ComCliente = " & idCliente _
            & " And ComTipo = TCoCodigo " _
            & " And TCoAccion IN (" & aCom & ")"
    Set rsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not rsCom.EOF Then sHay = True
    rsCom.Close
    
    If sHay Then
        Dim aObj As New clsCliente
        aObj.Comentarios idCliente:=idCliente
        DoEvents
        Set aObj = Nothing
    End If
    MsgClienteNoVender idCliente, True
    Screen.MousePointer = 0
    Exit Sub
    
errMenu:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al fomulario de comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub

Public Function MsgClienteNoVender(ByVal iCliente As Long, ByVal bShowMsg As Boolean) As Boolean
Dim rsCom As rdoResultset
    MsgClienteNoVender = False
    Set rsCom = cBase.OpenResultset("exec gennovender " & iCliente, rdOpenDynamic, rdConcurValues)
    If Not rsCom.EOF Then
        If Not IsNull(rsCom(0)) Then
            If rsCom(0) = 1 Then
                MsgClienteNoVender = True
                If bShowMsg Then
                    Screen.MousePointer = 0
                    MsgBox "Atención: el cliente tiene la categoría de no vender. Consultar con gerencia!", vbCritical, "ATENCIÓN"
                End If
            End If
        End If
    End If
    rsCom.Close
End Function


