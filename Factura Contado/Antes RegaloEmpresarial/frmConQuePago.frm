VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmConQuePago 
   Appearance      =   0  'Flat
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cobrar con?"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPWD 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton butCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton butAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConQuePaga 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3201
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
      BackColorBkg    =   16777215
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
      SelectionMode   =   0
      GridLines       =   1
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
   Begin VB.TextBox txtConQue 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox cboConQue 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   840
      Width           =   2295
   End
   Begin VB.PictureBox picTitulo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Informe con que cobra"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblAyuda 
      BackColor       =   &H00F0FFFF&
      Caption         =   "Label7"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   7080
      X2              =   120
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblImpSaldo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A cobrar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblImpACobrar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A cobrar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "A cobrar:"
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblImpAsignado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8,888,888.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total asignado:"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione con que paga"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Paga con:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmConQuePago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public UsuarioGraba As Long
Public Usuario As Long

Public ConQuePaga As clsDesicionConQuePaga
'parámetros
Public ImporteACobrar As Currency       'es lo que tengo que controlar que se cubra.
Public TransaccionEntrada As clsConQuePaga
Public idCliente As Long

Public Function ValidoCobroSaldado() As Boolean
On Error GoTo errL
ValidoCobroSaldado = False  'Indica que no se salda el cobro.

    lblAyuda.BackColor = Me.BackColor
    lblImpSaldo.ForeColor = lblImpACobrar.ForeColor
    txtPWD.Text = ""
    InicializoFormulario
    IntentoLevantarDatosCliente
    SumarImporteAsignado
    
    If (CCur(lblImpSaldo.Caption) = 0) Then
        ValidoCobroSaldado = AccionGrabar(False)
    End If
    Exit Function
errL:
    clsGeneral.OcurrioError "Error al cargar el formulario.", Err.Description, "Con que cobra"
End Function

Function CreoConQuePaga(ByVal ID As Long, ByVal Tipo As eDocConQuePaga, ByVal importe As Currency) As clsConQuePaga
    Set CreoConQuePaga = New clsConQuePaga
    With CreoConQuePaga
        .IDDocumentoPaga = ID
        .importe = importe
        .TipoConQuePaga = Tipo
    End With
End Function

Sub CargoAportesParaCedulaRUT(ByVal sCI As String)
On Error GoTo errCA
'TODO:
'PEDIR Sucesos si no es de él.
'Si no es del cliente hay que registrar el cliente dueño.
Dim sQY As String
Dim rsA As rdoResultset
Dim bOK As Boolean
    sQY = "SELECT dbo.FormatCIRuc(CliCiRuc) IDDoc, IsNull(dbo.SaldoCtaPersonal(CliCodigo), 0) Importe, CliCodigo" & _
        " FROM Cliente WHERE CliCIRUC = '" & sCI & "'"
    Set rsA = cBase.OpenResultset(sQY, rdOpenDynamic, rdConcurValues)
    If Not rsA.EOF Then
        If rsA("Importe") > 0 Then
            bOK = True
            If Not YaEstaEnGrilla(AsignoAporteCta, rsA("CliCodigo")) Then
                InsertoItemEnGrilla "Aporte a cuenta", rsA("IDDoc"), CreoConQuePaga(rsA("CliCodigo"), AsignoAporteCta, rsA("Importe")), flexChecked
                SumarImporteAsignado
            Else
                MsgBox "Ya está en la grilla.", vbExclamation, "Duplicación"
            End If
        End If
    End If
    rsA.Close
    
    If Not bOK Then MsgBox "No hay datos para la cédula / RUT ingresado.", vbExclamation, "ATENCIÓN"
    Exit Sub
errCA:
    clsGeneral.OcurrioError "Error al buscar el aporte.", Err.Description, "Buscar"
End Sub

Sub BuscoCargoTransaccionesRedPagos(ByVal idTra As Long)
On Error GoTo errCA
Dim sQY As String
Dim rsA As rdoResultset
Dim bOK As Boolean
    
    sQY = "SELECT TraID ID, CASE TItTipoItem WHEN 1 THEN 'Cuota' WHEN 2 THEN 'Web' WHEN 9 THEN 'Aporte' ELSE 'Op'+ CONVERT(VarChar(3), TItTipoItem) END Tipo" & _
        ", TItImporte Importe" & _
        " FROM comTransacciones INNER JOIN comTransaccionItems ON TraID = TItTransaccion" & _
        " WHERE TraEstado In (1,16) AND TItTipoItem IN (2,9) AND TraID = " & idTra
    Set rsA = cBase.OpenResultset(sQY, rdOpenDynamic, rdConcurValues)
    If Not rsA.EOF Then
        bOK = True
        If Not YaEstaEnGrilla(GiroRedPagos, rsA("ID")) Then
            InsertoItemEnGrilla "Transacción " & rsA("Tipo"), rsA("ID"), CreoConQuePaga(rsA("ID"), GiroRedPagos, rsA("Importe")), flexChecked
        Else
            MsgBox "Ya está en la grilla.", vbExclamation, "Duplicación"
        End If
    End If
    rsA.Close
    If Not bOK Then MsgBox "No hay datos para la transacción ingresada.", vbExclamation, "ATENCIÓN"
    Exit Sub
errCA:
    clsGeneral.OcurrioError "Error al buscar el aporte.", Err.Description, "Buscar"
End Sub

Sub BuscoPendientesDeCaja(ByVal idP As Long)
On Error GoTo errCA
Dim sQY As String
Dim rsA As rdoResultset
Dim bOK As Boolean

    sQY = "SELECT PCaID ID, PCaImporte * -1 Importe" & _
        " FROM PendientesCaja INNER JOIN Documento ON PCaDocumento = DocCodigo AND DocCliente = " & idCliente & _
        " WHERE PCaImporte < 0 AND PCaFLiquidacion IS NULL AND PCaID = " & idP
    Set rsA = cBase.OpenResultset(sQY, rdOpenDynamic, rdConcurValues)
    If Not rsA.EOF Then
        bOK = True
        If Not YaEstaEnGrilla(eDocConQuePaga.PendienteCajaNegativoNuevo, rsA("ID")) Then
            InsertoItemEnGrilla "Pendiente negativo", rsA("ID"), CreoConQuePaga(rsA("ID"), PendienteCajaNegativoNuevo, rsA("Importe"))
        Else
            MsgBox "Ya está en la grilla.", vbExclamation, "Duplicación"
        End If
    Else
        MsgBox "El ID ingresado no retorna datos", vbExclamation, "ATENCIÓN"
    End If
    rsA.Close
    Exit Sub
errCA:
    clsGeneral.OcurrioError "Error al buscar el aporte.", Err.Description, "Buscar"
End Sub

Function YaEstaEnGrilla(ByVal TipoAporte As eDocConQuePaga, ByVal ID As Long) As Boolean
Dim iR As Integer
    With vsConQuePaga
        For iR = 1 To .Rows - 1
            If .Cell(flexcpData, iR, 0) = ID And .Cell(flexcpData, iR, 1) = TipoAporte Then
                YaEstaEnGrilla = True
                Exit Function
            End If
        Next
    End With
End Function

Sub InicializoFormulario()
    
    lblImpACobrar.Caption = Format(ImporteACobrar, "#,##0.00")
    lblAyuda.Caption = ""
    
    'Cargo el combo de opciones.
    With cboConQue
        .Clear
        .AddItem "Aporte a cuenta"
        .ItemData(.NewIndex) = eDocConQuePaga.AsignoAporteCta
        
        .AddItem "Pendientes"
        .ItemData(.NewIndex) = eDocConQuePaga.PendienteCajaNegativoNuevo
        
        .AddItem "Transacción redpagos"
        .ItemData(.NewIndex) = eDocConQuePaga.GiroRedPagos
    End With
        
    With vsConQuePaga
        .Rows = 1
        .Cols = 1
        .FixedCols = 0
        .FormatString = "Asignar|Tipo aporte|Código|Importe"
        .ExtendLastCol = True
        .ColDataType(0) = flexDTBoolean
        .ColWidth(1) = 2700
        .ColWidth(2) = 1600
        .ColWidth(3) = 1500
        .Editable = True
    End With
    
End Sub

Sub InsertoItemEnGrilla(ByVal Texto As String, ByVal Codigo As String, ByVal PagaCon As clsConQuePaga, Optional ByVal Checked As Byte = flexUnchecked)
    
    With vsConQuePaga
        .AddItem ""
        .Cell(flexcpChecked, .Rows - 1, 0) = Checked
        .Cell(flexcpText, .Rows - 1, 1) = Texto
        .Cell(flexcpText, .Rows - 1, 2) = Codigo
        .Cell(flexcpText, .Rows - 1, 3) = Format(PagaCon.importe, "#,##0.00")
        
        .Cell(flexcpData, .Rows - 1, 0) = PagaCon.IDDocumentoPaga
        .Cell(flexcpData, .Rows - 1, 1) = PagaCon.TipoConQuePaga
    End With
    
    If TransaccionEntrada Is Nothing Then Exit Sub
    If TransaccionEntrada.TipoConQuePaga = PagaCon.TipoConQuePaga And TransaccionEntrada.IDDocumentoPaga = PagaCon.IDDocumentoPaga Then
        vsConQuePaga.Cell(flexcpChecked, vsConQuePaga.Rows - 1, 0) = flexChecked
        Set TransaccionEntrada = Nothing
    End If
    
End Sub

Sub IntentoLevantarDatosCliente()
    'Obtengo la información de posibles cobros del cliente.
    'Una vez cargada la grilla asigno la transacción recibida.
Dim sQY As String
Dim rsA As rdoResultset

    'Cargo los aportes a cuenta.
    sQY = "SELECT dbo.FormatCIRuc(CliCiRuc) IDDoc, IsNull(dbo.SaldoCtaPersonal(CliCodigo), 0) Importe" & _
        " FROM Cliente WHERE CliCodigo = " & idCliente
    Set rsA = cBase.OpenResultset(sQY, rdOpenDynamic, rdConcurValues)
    If Not rsA.EOF Then
        If rsA("Importe") > 0 Then
            InsertoItemEnGrilla "Aporte a cuenta", rsA("IDDoc"), CreoConQuePaga(idCliente, eDocConQuePaga.AsignoAporteCta, rsA("Importe"))
        End If
    End If
    rsA.Close
    
    'Cargo todas las transacciones redpago.
    sQY = "SELECT TraID ID, CASE TItTipoItem WHEN 1 THEN 'Cuota' WHEN 2 THEN 'Web' WHEN 9 THEN 'Aporte' ELSE 'Op'+ CONVERT(VarChar(3), TItTipoItem) END Tipo" & _
        ", TItImporte Importe" & _
        " FROM comTransacciones INNER JOIN comTransaccionItems ON TraID = TItTransaccion" & _
        " WHERE TraEstado In (1,16) AND TItTipoItem IN (2,9) AND TraClienteCGSA = " & idCliente & _
        " ORDER BY TraFecha DESC"
    Set rsA = cBase.OpenResultset(sQY, rdOpenDynamic, rdConcurValues)
    Do While Not rsA.EOF
        InsertoItemEnGrilla "Transacción " & rsA("Tipo"), rsA("ID"), CreoConQuePaga(rsA("ID"), GiroRedPagos, rsA("Importe"))
        rsA.MoveNext
    Loop
    rsA.Close

    sQY = "SELECT PCaID ID, PCaImporte * -1 Importe" & _
        " FROM PendientesCaja INNER JOIN Documento ON PCaDocumento = DocCodigo AND DocCliente = " & idCliente & _
        " WHERE PCaImporte < 0 AND PCaFLiquidacion IS NULL"
    Set rsA = cBase.OpenResultset(sQY, rdOpenDynamic, rdConcurValues)
    Do While Not rsA.EOF
        InsertoItemEnGrilla "Pendiente negativo", rsA("ID"), CreoConQuePaga(rsA("ID"), PendienteCajaNegativoNuevo, rsA("Importe"))
        rsA.MoveNext
    Loop
    rsA.Close
    
End Sub

Sub AsignoAportesSeleccionados()
Dim iR As Integer
    Set ConQuePaga.ConQuePaga = Nothing
    Set ConQuePaga.ConQuePaga = New Collection
    With vsConQuePaga
        For iR = 1 To .Rows - 1
            If .Cell(flexcpChecked, iR, 0) = flexChecked Then
                ConQuePaga.ConQuePaga.Add CreoConQuePaga(.Cell(flexcpData, iR, 0), .Cell(flexcpData, iR, 1), .Cell(flexcpText, iR, 3))
            End If
        Next
    End With
End Sub

Sub SumarImporteAsignado()
Dim iR As Integer
Dim cAsignado As Currency
    With vsConQuePaga
        For iR = 1 To .Rows - 1
            If .Cell(flexcpChecked, iR, 0) = flexChecked Then
                cAsignado = cAsignado + CCur(.Cell(flexcpValue, iR, 3))
            End If
        Next
    End With
    lblImpAsignado.Caption = Format(cAsignado, "#,##0.00")
    lblImpSaldo.Caption = Format(cAsignado - CCur(lblImpACobrar.Caption), "#,##0.00")
    lblImpSaldo.ForeColor = IIf(cAsignado - CCur(lblImpACobrar.Caption) < 0, &H80&, lblImpACobrar.ForeColor)
End Sub

Function UsuarioClave() As Boolean
On Error GoTo errUC
Dim usr As Long
    Dim oCnxt As New clsConexion
    If Usuario = 0 Then
        usr = oCnxt.UsuarioLogueado(True)
    Else
        usr = Usuario
    End If
    UsuarioClave = oCnxt.ValidoClave(usr, txtPWD.Text)
    Set oCnxt = Nothing
    If UsuarioClave Then UsuarioGraba = usr Else UsuarioGraba = 0
    Exit Function
errUC:
    UsuarioClave = False
End Function

Function AccionGrabar(ByVal bPedirClave As Boolean) As Boolean
On Error GoTo errBA
    
    Set ConQuePaga = Nothing
    UsuarioGraba = 0
    
    If CCur(lblImpSaldo.Caption) < 0 Then
        MsgBox "No se está cubriendo el importe a cobrar, debe ingresar nuevos aportes y asignarlos.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If bPedirClave Then
        If Not UsuarioClave() Then
            MsgBox "La contraseña ingresada no corresponde al usuario.", vbExclamation, "ATENCIÓN"
            txtPWD.SetFocus
            Exit Function
        End If
    Else
        UsuarioGraba = Usuario
    End If
    
    Dim SaldoA As Byte
    If CCur(lblImpSaldo.Caption) > 0 Then
        'TENGO QUE VOLCAR EL SALDO A UN PENDIENTE O UN NUEVO APORTE.
        Dim frmSA As New frmSaldoAFavor
        frmSA.Show vbModal, Me
        If frmSA.TipoAporte > 0 Then
            SaldoA = frmSA.TipoAporte
        Else
            Exit Function
        End If
    End If
    
    Set ConQuePaga = New clsDesicionConQuePaga
    ConQuePaga.SaldoAFavor = CCur(lblImpSaldo.Caption)
    ConQuePaga.VolcarSaldoAFavor = SaldoA
    'Cargo colección de aportes.
    AsignoAportesSeleccionados
    
    If ConQuePaga.ConQuePaga.Count = 0 Then
        MsgBox "No hay aportes asignados.", vbExclamation, "ATENCIÓN"
        Set ConQuePaga = Nothing
        Exit Function
    End If

    AccionGrabar = True
    Exit Function
errBA:
    clsGeneral.OcurrioError "Error al confirmar los datos ingresados.", Err.Description, "ATENCIÓN"
End Function

Private Sub butAceptar_Click()
    If AccionGrabar(True) Then Unload Me
End Sub

Private Sub butCancel_Click()
    Unload Me
End Sub

Private Sub cboConQue_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then txtConQue.SetFocus
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
End Sub

Private Sub txtConQue_GotFocus()
Dim msgH As String

    If cboConQue.ListIndex = -1 Then cboConQue.SetFocus: Exit Sub
    
    Select Case cboConQue.ItemData(cboConQue.ListIndex)
        Case eDocConQuePaga.AsignoAporteCta
            msgH = "Ingrese la C.I./R.U.T. del cliente que posee el aporte y de enter para buscar."
        Case eDocConQuePaga.PendienteCajaNegativoNuevo
            msgH = "Ingrese el código del pendiente y presione enter para buscar."
        Case eDocConQuePaga.GiroRedPagos
            msgH = "Ingrese el código de la transacción de redpagos y de enter."
        Case Else
            msgH = ""
    End Select
    lblAyuda.Caption = msgH
    If lblAyuda.Caption <> "" Then lblAyuda.BackColor = &HF0FFFF
End Sub

Private Sub txtConQue_KeyPress(KeyAscii As Integer)
On Error GoTo errKP
    If KeyAscii = vbKeyReturn Then
        If Trim(txtConQue.Text) <> "" Then
            Screen.MousePointer = 11

            If cboConQue.ListIndex = -1 Then cboConQue.SetFocus: Exit Sub
            
            Select Case cboConQue.ItemData(cboConQue.ListIndex)
                Case eDocConQuePaga.AsignoAporteCta
                    CargoAportesParaCedulaRUT txtConQue.Text
                Case eDocConQuePaga.PendienteCajaNegativoNuevo
                    BuscoPendientesDeCaja txtConQue.Text
                Case eDocConQuePaga.GiroRedPagos
                    BuscoCargoTransaccionesRedPagos txtConQue.Text
            End Select
            SumarImporteAsignado
            txtConQue.SelStart = 0
            txtConQue.SelLength = Len(txtConQue.Text)
            
            Screen.MousePointer = 0
        End If
    End If
    Exit Sub
errKP:
    clsGeneral.OcurrioError "Error al buscar la información.", Err.Description, "ATENCIÓN"
    Screen.MousePointer = 0
End Sub

Private Sub txtConQue_LostFocus()
    lblAyuda.Caption = ""
    lblAyuda.BackColor = Me.BackColor
End Sub

Private Sub txtPWD_GotFocus()
On Error Resume Next
txtPWD.SelStart = 0
txtPWD.SelLength = Len(txtPWD.Text)
End Sub

Private Sub txtPWD_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then butAceptar_Click
End Sub

Private Sub vsConQuePaga_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    SumarImporteAsignado
End Sub

Private Sub vsConQuePaga_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col <> 0)
End Sub

