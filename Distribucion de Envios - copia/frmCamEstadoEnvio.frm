VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form CamEstadoEnvio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retornar envío a pendiente de entrega"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCamEstadoEnvio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2600
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox tCodigo 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   960
      MaxLength       =   2
      TabIndex        =   5
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox tComentario 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   3855
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4605
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   0
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
            Picture         =   "frmCamEstadoEnvio.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoEnvio.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoEnvio.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoEnvio.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoEnvio.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoEnvio.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoEnvio.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoEnvio.frx":0DC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Al grabar el envío el estado del mismo volverá a ser 'por entregar' y la mercadería se retornará nuevamente al camión."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   3855
   End
   Begin VB.Label labLiquidacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Liquidación:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label labFEntregado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label labCamion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentario del envío:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Entregado:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Camión:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Código de Envío:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
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
         Caption         =   "&Volver al Formulario Anterior"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "CamEstadoEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const cte_NOMBREFORM As String = "Cambio estado de envío"

Private Function BuscoUsuario(Digito As Integer) As Integer
On Error GoTo ErrBU

    Cons = "SELECT * FROM USUARIO WHERE UsuDigito = " & Digito
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAux.EOF Then
        BuscoUsuario = 0
        MsgBox "No existe un usuario para el dígito ingresado.", vbExclamation, "ATENCIÓN"
    Else
        BuscoUsuario = RsAux!UsuCodigo
    End If
    RsAux.Close
    Exit Function
    
ErrBU:
    objGral.OcurrioError "Ocurrió un error inesperado."
    BuscoUsuario = 0
End Function


Private Sub Form_Activate()
    Me.Refresh
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    LimpioCampos
    Exit Sub
ErrLoad:
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Error al iniciar el formulario.", Err.Description, cte_NOMBREFORM
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Forms(Forms.Count - 2).SetFocus
End Sub

Private Sub Label1_Click()
    tCodigo.SetFocus
End Sub

Private Sub Label4_Click()
    tComentario.SetFocus
End Sub

Private Sub Label5_Click()
    tUsuario.SetFocus
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub
Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub
Private Sub MnuVolver_Click()
    Unload Me
End Sub
Private Sub AccionGrabar()
Dim iEstadoArticulo As Integer
Dim rsE As rdoResultset, rs As rdoResultset
Dim iCodImp As Long, lDoc As Long

    If Trim(tUsuario.Tag) <> vbNullString Then
        If Not CInt(tUsuario.Tag) > 0 Then
            MsgBox "Ingrese su dígito de usuario.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
    Else
        MsgBox "Ingrese su dígito de usuario.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("¿Confirma volver el envío a estado 'POR ENTREGAR'?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
        
        'Veo si es Venta Telefónica
        Cons = "Select * From VentaTelefonica Where VTeDocumento = " & Val(tCodigo.Text) & " AND VTeTipo <> 44"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)
            
        If Not RsAux.EOF Then
            If labCamion.Tag = "1" Then    ' Not IsNull(rsEnvio!EnvReclamoCobro)
                RsAux.Close
                
                Cons = "Select * From Envio" _
                    & " Where EnvCodigo <> " & Val(tCodigo.Text) _
                    & " And EnvDocumento = " & Val(tComentario.Tag) _
                    & " And EnvEstado IN(" & EstadoEnvio.Entregado & ", " & EstadoEnvio.Impreso & ")"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux.EOF Then
                    RsAux.Close
                    MsgBox "Este envío es de COBRANZA y existen otros que ya fueron entregados o están impresos, no podrá volver este envío atras sin tener los otros envíos en estado de espera de impresión, verifique.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
                RsAux.Close
            End If
        Else
            RsAux.Close
        End If
        
        Dim objSuceso As New clsSuceso
        Dim aUsuario As Long, strDefensa As String, sDesc As String
        aUsuario = 0
        objSuceso.TipoSuceso = 98
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "Cambio de estado de envío", cBase
        Me.Refresh
        aUsuario = objSuceso.Usuario
        strDefensa = objSuceso.Defensa
        Set objSuceso = Nothing
        If aUsuario = 0 Then Screen.MousePointer = 0: Exit Sub
        sDesc = "Se retornó a estado impreso envío: " & Val(tCodigo.Text)
        
        On Error GoTo ErrBT
        Screen.MousePointer = vbHourglass
        FechaDelServidor
        cBase.BeginTrans
        On Error GoTo errResumo
        
        Cons = "Select * From Envio Where EnvCodigo = " & Val(tCodigo.Text)
        Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If rsE.EOF Then
            Screen.MousePointer = vbDefault
            rsE.Close
            cBase.RollbackTrans
            MsgBox "El envío fue eliminado por otra terminal, verifique.", vbExclamation, "ATENCIÓN"
            AccionCancelar
            Exit Sub
        Else
            If CDate(labFEntregado.Tag) <> rsE!EnvFModificacion Then
                Screen.MousePointer = vbDefault
                rsE.Close
                cBase.RollbackTrans
                MsgBox "El envío fue modificado por otra terminal, refresque la información y verifique.", vbExclamation, "ATENCIÓN"
                BuscoEnvio Val(tCodigo.Text)
                Exit Sub
            Else
                'Modifico el Stock.-----------------------------
                
                '1) _Devuelvo la cantidad como a Entregar.
                Cons = "Update RenglonEnvio Set REvAEntregar = REvCantidad " _
                    & " Where REvEnvio = " & Val(tCodigo.Text)
                cBase.Execute (Cons)
                
                '2) _Tengo que ver si los renglones entrega para cada artículo existe, sino lo inserto.
                Cons = "Select * From RenglonEnvio Where REvEnvio = " & Val(tCodigo.Text) _
                    & " And REvArticulo NOT IN" _
                        & "(Select ArtID From Articulo Where ArtTipo IN(SELECT TipID from  dbo.InTipos(" & paTipoArticuloServicio & ")))"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                'Me quedo con el código de impresión.
                If Not RsAux.EOF Then
                    iCodImp = rsE("EnvCodImpresion")
                    sDesc = sDesc & ", Código de impresión: " & iCodImp
                    Do While Not RsAux.EOF
                        Cons = "Select * From RenglonEntrega Where ReECodImpresion = " & iCodImp _
                                & " And ReEArticulo = " & RsAux!REvArticulo
                        Set rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        If rs.EOF Then
                            rs.AddNew
                            rs!ReECodImpresion = RsAux!REvCodImpresion
                            rs!ReEArticulo = RsAux!REvArticulo
                            rs!ReECantidadTotal = RsAux!REvCantidad
                            rs!ReECantidadEntregada = RsAux!REvCantidad
                            rs!ReEFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:nn:ss")
                            rs!ReEEstado = paEstadoArticuloEntrega
                            rs!ReECamion = rsE!EnvCamion
                            rs!ReEUsuario = tUsuario.Tag
                            rs.Update
                            iEstadoArticulo = paEstadoArticuloEntrega
                        Else
                            iEstadoArticulo = rs!ReEEstado
                            rs.Edit
                            rs!ReECantidadTotal = rs!ReECantidadTotal + RsAux!REvCantidad
                            rs!ReECantidadEntregada = rs!ReECantidadEntregada + RsAux!REvCantidad
                            rs!ReEUsuario = tUsuario.Tag
                            rs.Update
                        End If
                        rs.Close
                        
                        '3)_ Tengo que ajustar el stock del camión y el stock total.
                        MarcoMovimientoStockTotal RsAux("REvArticulo"), TipoEstadoMercaderia.Virtual, TipoMovimientoEstado.AEntregar, RsAux("REvCantidad"), 1
'                        Cons = "Update StockTotal Set StTCantidad = StTCantidad + " & RsAux!REvAEntregar _
                            & " Where StTArticulo = " & RsAux!REvArticulo _
                            & " And StTEstado = " & TipoMovimientoEstado.AEntregar _
                            & " And StTTipoEstado = " & TipoEstadoMercaderia.Virtual
                        
                        'Le doy la mercadería al camión nuevamente.
                        MarcoMovimientoStockFisicoEnLocal TipoLocal.Camion, rsE("EnvCamion"), RsAux("REvArticulo"), RsAux("REvCantidad"), iEstadoArticulo, 1
                        
                        lDoc = rsE!EnvDocumento
                        Cons = "Select * From Documento Where DocCodigo= " & rsE!EnvDocumento
                        Set rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                        MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Camion, rsE!EnvCamion, RsAux!REvArticulo, RsAux!REvAEntregar, iEstadoArticulo, 1, rs!docTipo, rs!DocCodigo
                        MarcoMovimientoStockEstado CLng(tUsuario.Tag), RsAux!REvArticulo, RsAux!REvAEntregar, TipoMovimientoEstado.AEntregar, 1, rs!docTipo, rs!DocCodigo, paCodigoDeSucursal
                        rs.Close
                        RsAux.MoveNext
                    Loop
                    Cons = "Update RenglonEntrega Set ReEFModificacion = '" & Format(gFechaServidor, "mm/dd/yyyy hh:nn:ss") & "' Where ReECodImpresion = " & iCodImp
                    cBase.Execute (Cons)
                End If
                RsAux.Close
            
                rsE.Edit
                rsE!EnvFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:nn:ss")
                rsE!EnvEstado = EstadoEnvio.Impreso
                rsE!EnvFechaEntregado = Null
                If Trim(tComentario.Text) <> vbNullString Then rsE!EnvComentario = Trim(Replace(tComentario.Text, vbCrLf, ""))
                'Si fue liquidado lo saco (ya que se pueden generar nuevos documentos de flete a liquidar).
                rsE("EnvLiquidacion") = Null
                rsE.Update
            End If
        End If
        rsE.Close
        objGral.RegistroSucesoAutorizado cBase, gFechaServidor, 98, paCodigoDeTerminal, aUsuario, lDoc, Descripcion:=sDesc, Defensa:=Trim(strDefensa)
        cBase.CommitTrans
        AccionCancelar
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrBT:
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Ocurrió un error al iniciar la transacción."
    
    
errResumo:
    Resume ErrErrorGrabar
    Exit Sub
    
ErrErrorGrabar:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    objGral.OcurrioError "Ocurrió un error al intentar grabar los datos."

End Sub

Private Sub AccionCancelar()
    LimpioCampos
End Sub

Private Sub optEstado_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tUsuario.SetFocus
End Sub

Private Sub optEstado_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.SimpleText = " Seleccione el estado a asignarle al envío."
End Sub

Private Sub tCodigo_Change()
    If Val(tCodigo.Tag) > 0 Then LimpioCampos
End Sub

Private Sub tCodigo_GotFocus()
    Status.SimpleText = " Ingrese el código de envío."
    tCodigo.SelStart = 0
    tCodigo.SelLength = Len(tCodigo.Text)
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tCodigo.Text) <> vbNullString Then
        If Not IsNumeric(tCodigo.Text) Then
            MsgBox "No se ingreso un formato válido.", vbExclamation, "ATENCIÓN"
            tCodigo.SetFocus
        Else
            BuscoEnvio Val(tCodigo.Text)
            If MnuGrabar.Enabled Then tComentario.SetFocus
        End If
    End If
End Sub

Private Sub tComentario_GotFocus()
    Status.SimpleText = " Ingrese o modifique el comentario del envío."
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tUsuario.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select

End Sub

Private Sub tUsuario_GotFocus()

    tUsuario.SelStart = 0
    tUsuario.SelLength = Len(tUsuario.Text)
    Status.SimpleText = " Ingrese el dígito de usuario."

End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And IsNumeric(tUsuario.Text) Then
        With tUsuario
            .Tag = BuscoUsuario(Val(.Text))
            If .Tag = 0 Then
                .Text = vbNullString
                .Tag = vbNullString
                Exit Sub
            End If
            If .Tag <> vbNullString Then AccionGrabar
        End With
    End If
    
End Sub

Private Sub LimpioCampos()

    tCodigo.Text = vbNullString
    tCodigo.Tag = ""
    labFEntregado.Caption = vbNullString: labFEntregado.Tag = ""
    labCamion.Caption = vbNullString: labCamion.Tag = ""
    labLiquidacion.Caption = vbNullString
    tComentario.Text = vbNullString: tComentario.Tag = ""
    tUsuario.Text = vbNullString
    tUsuario.Tag = vbNullString
    
    'Inhabilito los botones
    Toolbar1.Buttons("grabar").Enabled = False
    MnuGrabar.Enabled = False
    Toolbar1.Buttons("cancelar").Enabled = False
    MnuCancelar.Enabled = False
    
End Sub

Private Sub BuscoEnvio(ByVal lCod As Long)
On Error GoTo ErrBE
Dim rsEnvio As rdoResultset
    
    Screen.MousePointer = 11
    LimpioCampos
    Cons = "Select * From Envio Where EnvCodigo = " & lCod
    Set rsEnvio = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsEnvio.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "No existe un envío con ese código, verifique.", vbInformation, "ATENCIÓN"
    Else
        If rsEnvio!EnvEstado = EstadoEnvio.Entregado Then
            tCodigo.Text = rsEnvio!EnvCodigo
            tCodigo.Tag = rsEnvio!EnvCodigo
            If Not IsNull(rsEnvio!EnvFechaEntregado) Then
                labFEntregado.Caption = Format(rsEnvio!EnvFechaEntregado, "d-Mmm-yyyy")
                labFEntregado.Tag = rsEnvio("EnvFModificacion")
            End If
            If Not IsNull(rsEnvio!EnvCamion) Then labCamion.Caption = BuscoNombreCamionero(rsEnvio!EnvCamion)
            If Not IsNull(rsEnvio!EnvReclamoCobro) Then labCamion.Tag = "1"
            If Not IsNull(rsEnvio!EnvComentario) Then
                tComentario.Text = Trim(rsEnvio!EnvComentario)
                tComentario.Tag = rsEnvio("EnvDocumento")
            End If
            If Not IsNull(rsEnvio!Envliquidacion) Then
                If Not IsNull(rsEnvio!EnvLiquidar) Then labLiquidacion.Caption = Format(rsEnvio!EnvLiquidar, "#,##0.00")
                MsgBox "Atención este envío ya fue liquidado al camión, verifique.", vbInformation, "Atención"
            End If
            
            Toolbar1.Buttons("grabar").Enabled = True
            MnuGrabar.Enabled = True
            Toolbar1.Buttons("cancelar").Enabled = True
            MnuCancelar.Enabled = True
        Else
            Screen.MousePointer = vbDefault
            MsgBox "El envio seleccionado no ha sido entregado, verifique.", vbExclamation, "ATENCIÓN"
        End If
    End If
    rsEnvio.Close
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrBE:
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Ocurrió un error al buscar el envío."
    
End Sub

Private Function BuscoNombreCamionero(iCamionero As Integer) As String

    Cons = "Select CamNombre From Camion Where CamCodigo = " & iCamionero
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        BuscoNombreCamionero = vbNullString
    Else
        BuscoNombreCamionero = Trim(RsAux!CamNombre)
    End If
    RsAux.Close
    
End Function
