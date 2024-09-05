VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConteo 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conteo de Artículos"
   ClientHeight    =   3120
   ClientLeft      =   4125
   ClientTop       =   2925
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConteo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4845
   Begin MSComctlLib.StatusBar sbHelp 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   2850
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8043
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmUsr 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   4260
      Top             =   0
   End
   Begin VB.TextBox tQ 
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   1035
   End
   Begin AACombo99.AACombo cLocal 
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Top             =   60
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VB.TextBox tArticulo 
      Height          =   315
      Left            =   780
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
   Begin AACombo99.AACombo cEstado 
      Height          =   315
      Left            =   780
      TabIndex        =   5
      Top             =   840
      Width           =   1395
      _ExtentX        =   2461
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
   Begin VB.Label lUsr 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   15
      Left            =   0
      TabIndex        =   11
      Top             =   2400
      Width           =   4755
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   15
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   4755
   End
   Begin VB.Label lMsg 
      Appearance      =   0  'Flat
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   15
      TabIndex        =   9
      Top             =   1380
      Width           =   4755
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Q:"
      Height          =   255
      Left            =   2700
      TabIndex        =   6
      Top             =   900
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Estado:"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Local:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Artículo:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   1035
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "?"
      Begin VB.Menu MnuHlp 
         Caption         =   "Ayuda"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmConteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmIDLocal As Long
Public prmIDArticulo As Long

Dim dConteo As typConteo
Dim prmVerificador As Boolean

Private Sub cEstado_Change()
    If dConteo.flagData Then
        dConteo.flagData = False
        LimpioMensaje
    End If
End Sub

Private Sub cEstado_Click()
    If dConteo.flagData Then
        dConteo.flagData = False
        LimpioMensaje
    End If
End Sub

Private Sub cEstado_GotFocus()
    ReestartUsr
End Sub

Private Sub cEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tQ
End Sub

Private Sub cLocal_Change()
    If dConteo.flagData Then
        dConteo.flagData = False
        LimpioMensaje
    End If
    
End Sub

Private Sub cLocal_GotFocus()
    ReestartUsr
End Sub

Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tArticulo
End Sub

Private Sub Form_Load()
    InicializoForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndMain
End Sub

Private Sub InicializoForm()
    
    On Error Resume Next
    LimpioFicha
    
    cons = "Select LocCodigo, LocNombre from Local order by LocNombre"
    CargoCombo cons, cLocal
    BuscoCodigoEnCombo cLocal, paCodigoDeSucursal
    
    cons = "Select EsMCodigo, EsMNombre from EstadoMercaderia Order by ESMNombre"
    CargoCombo cons, cEstado
    BuscoCodigoEnCombo cEstado, paEstadoSano
    
    If prmIDLocal <> 0 Then BuscoCodigoEnCombo cLocal, prmIDLocal
    
    If prmIDArticulo <> 0 Then
        cons = "Select * from Articulo Where ArtId = " & prmIDArticulo
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            tArticulo.Text = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
            tArticulo.Tag = rsAux!ArtID
            Me.Show
            Foco cEstado
        End If
        rsAux.Close
    End If
    
End Sub

Private Sub LimpioFicha()
    
    cLocal.Text = ""
    cEstado.Text = ""
    tArticulo.Text = ""
    tQ.Text = ""
    
    lMsg.Caption = ""
    lUsr.Caption = "Ingresando Datos: ....."
    
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = 0
    If dConteo.flagData Then
        dConteo.flagData = False
        LimpioMensaje
    End If
End Sub

Private Function ReestartUsr()
    
    If Val(lUsr.Tag) <> 0 Then
        tmUsr.Enabled = False: tmUsr.Tag = 0
        tmUsr.Enabled = True
    End If
    
End Function

Private Sub tArticulo_GotFocus()
    ReestartUsr
    sbHelp.Panels("help").Text = ""
    On Error Resume Next
    tArticulo.SelStart = 0: tArticulo.SelLength = Len(tArticulo.Text)
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tArticulo.Text) = "" Then Exit Sub
        If Val(tArticulo.Tag) <> 0 Then Foco cEstado: Exit Sub
        
        On Error GoTo errBuscaG
        
        cons = "Select ArtID, ArtCodigo as Codigo, ArtNombre  as Nombre from Articulo "
        If IsNumeric(tArticulo.Text) Then
            cons = cons & " Where ArtCodigo = " & Trim(tArticulo.Text)
        Else
            cons = cons & "Where ArtNombre like '" & Replace(Trim(tArticulo.Text), " ", "%") & "%'"
        End If
        cons = cons & " Order by ArtNombre"
        
        Dim aQ As Integer, aIdArticulo As Long, aTexto As String
        aQ = 0: aIdArticulo = 0
        
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            aQ = 1
            aIdArticulo = rsAux!ArtID: aTexto = Format(rsAux!Codigo, "(#,000,000)") & " " & Trim(rsAux!Nombre)
            rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
        End If
        rsAux.Close
        
        Select Case aQ
            Case 0: MsgBox "No hay datos que coincidan con el texto ingersado.", vbExclamation, "No hay datos"
            
            Case 2:
                        Dim miLista As New clsListadeAyuda
                        aIdArticulo = miLista.ActivarAyuda(cBase, cons, 4000, 1, "Lista de Articulos")
                        Me.Refresh
                        If aIdArticulo > 0 Then
                            aIdArticulo = miLista.RetornoDatoSeleccionado(0)
                            
                            aTexto = Format(miLista.RetornoDatoSeleccionado(1), "(#,000,000)") & " "
                            aTexto = aTexto & miLista.RetornoDatoSeleccionado(2)
                        End If
                        Set miLista = Nothing
        End Select
        
        If aIdArticulo > 0 Then
            tArticulo.Text = aTexto
            tArticulo.Tag = aIdArticulo
        End If
        Screen.MousePointer = 0
    End If
   
    Exit Sub
errBuscaG:
    clsGeneral.OcurrioError "Error al buscar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tArticulo_LostFocus()
    sbHelp.Panels("help").Text = ""
End Sub


Private Function ValidoUsuario() As Boolean
    On Error GoTo errVUsr
    
    If Val(lUsr.Tag) = 0 Then
        Dim bAccesoOK As Boolean
        bAccesoOK = False
        
            
        frmLogin.Show vbModal
        Me.Refresh
        
        prmVerificador = ValidoAccesoMnu("ContadorVerificador", frmLogin.prmIDUsuario)
        
        If frmLogin.prmIDUsuario <> 0 Then
        
            lUsr.Tag = frmLogin.prmIDUsuario
            
            If Val(lUsr.Tag) <> 0 Then
                lUsr.Caption = "Ingresando Datos:  " & frmLogin.prmNombre
                tmUsr.Enabled = True
            End If
        End If
    End If
    
    ValidoUsuario = Not (Val(lUsr.Tag) = 0)
    Exit Function
    
errVUsr:
    clsGeneral.OcurrioError "Error al validar usuario logueado.", Err.Description
End Function

Private Function ValidoGrabar() As Boolean

    ValidoGrabar = False
    If Not ValidoUsuario Then Exit Function
    
    If Not dConteo.flagData Then
        MsgBox "Falta validar los datos del último conteo." & vbCrLf & _
                    "Verifique haber ingresado todos los datos.", vbExclamation, "Faltan Datos"
        Foco cLocal: Exit Function
    End If
    
    If cLocal.ListIndex = -1 Then
        Foco cLocal: Exit Function
    End If
    
    If Val(tArticulo.Tag) = 0 Then
        Foco tArticulo: Exit Function
    End If
    
    If cEstado.ListIndex = -1 Then
        Foco cEstado: Exit Function
    End If
    
    If Trim(tQ.Text) = "" Then
        Foco tQ: Exit Function
    Else
        If Not IsNumeric(tQ.Text) Then Foco tQ: Exit Function
    End If
    
    ValidoGrabar = True
    
End Function

Private Sub tmUsr_Timer()
    On Error Resume Next
    
    If Val(tmUsr.Tag) < 2 Then
        tmUsr.Tag = Val(tmUsr.Tag) + 1
        tmUsr.Enabled = False
        tmUsr.Enabled = True
        Exit Sub
    End If
    
    tmUsr.Enabled = False
    tmUsr.Tag = 0
    lUsr.Tag = 0
    lUsr.Caption = "Ingresando Datos: ....."
    dConteo.flagData = False
End Sub

Private Sub tQ_GotFocus()
    
    If Not dConteo.flagData Then InicializoData
    ReestartUsr
    tQ.SelStart = 0: tQ.SelLength = Len(tQ.Text)
End Sub

Private Sub tQ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub AccionGrabar()

On Error GoTo errValidar
Dim mQMovs As Long
Dim mMsg As String, mMsgDG As String

    'Si es Contador Verificador Y Puso AG en la Q ---> Agendo el Conteo     ----------------------------------------------------
    If prmVerificador And UCase(tQ.Text) = "AG" Then
        GrabarAgendarConteo
        Exit Sub
    End If
    
    With dConteo
        .QStockAnteriorBueno = 0
    End With
    
    If Not ValidoGrabar Then Exit Sub
    
    FechaDelServidor
    
    CargoUltimoStock
    
    If dConteo.EstadoAnteriorConteo = enuEConteo.paraRecuento Then
       mQMovs = ValidoMovimientosStock(dConteo.FechaInicialRecuento)
    Else
        mQMovs = ValidoMovimientosStock(gFechaServidor)
    End If
    
    'Valido Diferencias de stock con el útlimo conteo BUENO ----------------------------------------------------
    dConteo.EstadoConteo = enuEConteo.OK
    Dim mDifConAnterior As Long
    mDifConAnterior = CCur(tQ.Text) - dConteo.QStockLocalHoy
    
    If CCur(tQ.Text) <> dConteo.QStockLocalHoy Then
        If mDifConAnterior <> dConteo.QDifAnteriorBueno Then
        
            Select Case dConteo.EstadoAnteriorConteo
                Case enuEConteo.Agendado, enuEConteo.OK
                        
                        dConteo.EstadoConteo = enuEConteo.paraRecuento      'Hay que predir un recuento
                        
                        mMsgDG = "El conteo no coincide con el último stock válido. "
                        If mQMovs <> 0 Then mMsgDG = mMsgDG & vbCrLf & "En los últimos minutos se realizaron movimientos, los consideró ?."
                                                
                        mMsgDG = mMsgDG & vbCrLf & vbCrLf & _
                                        "Como las cantidades no coinciden, es necesario que ud. esté seguro de lo que contó. " & vbCrLf & _
                                        "Por favor recuente la mercadería, esto es MUY IMPORTANTE." & vbCrLf & vbCrLf & _
                                        "Verifique si se han realizado movimientos de stock durante el conteo." & vbCrLf & _
                                        "Ud. va a recontar la mercadería ahora ?"
                                        
                                    
                Case enuEConteo.paraRecuento
                        dConteo.EstadoConteo = enuEConteo.ParaVerificar
            End Select
            
        End If
    End If
    '---------------------------------------------------------------------------------------------------------------------
    
    If prmVerificador And dConteo.EstadoConteo <> enuEConteo.OK Then
        If MsgBox("Ud. es un Contador Verificador, y las cantidades no coinciden con las válidas." & vbCrLf & _
                       "Quiere asumir éstas cantidades como definitivas ?.", vbQuestion + vbYesNo + vbDefaultButton2, "Asumir Cantidades como Definitivas ?") = vbYes Then
            dConteo.EstadoConteo = enuEConteo.OK
        End If
    End If
    
    mMsg = "Confirma actualizar el conteo con los datos ingresados."
    If MsgBox(mMsg, vbQuestion + vbYesNo, "Grabar Conteo") = vbNo Then Exit Sub
    
    On Error GoTo errGrabar
    Screen.MousePointer = 11
    
    cons = "Select * from ConteoArticulo Where CArID = " & dConteo.IDNuevoConteo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then rsAux.AddNew Else rsAux.Edit
    
    rsAux!CArLocal = dConteo.cLocal
    rsAux!CArArticulo = dConteo.cArticulo
    rsAux!CArEstado = dConteo.cEstado
    
    If dConteo.IDNuevoConteo = 0 Then
        rsAux!CArFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm")
    Else
        If dConteo.EstadoAnteriorConteo = enuEConteo.paraRecuento Then
            Dim mDemora As Currency
            mDemora = DateDiff("s", rsAux!CArFecha, gFechaServidor)
            If mDemora > 32000 Then mDemora = 32000
            rsAux!CArDemoraRecuento = mDemora
        End If
        
        If dConteo.EstadoConteo = enuEConteo.OK Then
            rsAux!CArFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm")
        End If
    End If
    
    rsAux!CArQ = CCur(tQ.Text)
    rsAux!CArUsuario = Val(lUsr.Tag)
    
    rsAux!CarDifConLocal = CCur(tQ.Text) - dConteo.QStockLocalHoy
    
    rsAux!CArEstadoConteo = dConteo.EstadoConteo
    
    rsAux.Update
    rsAux.Close
    Screen.MousePointer = 0
    
    If dConteo.EstadoConteo = enuEConteo.paraRecuento Then
        Dim bPRec As Boolean: bPRec = False
        
        If mMsgDG <> "" Then
            bPRec = (MsgBox(mMsgDG, vbQuestion + vbYesNo, "Conteo Para Recontar") = vbYes)
        End If
        
        NuevoIngreso bPRec
        
    Else
        NuevoIngreso
        tmUsr.Tag = 0: tmUsr.Enabled = False: tmUsr.Enabled = True
    End If
    
    Exit Sub

errGrabar:
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Sub

errValidar:
    clsGeneral.OcurrioError "Error al procesar los datos para grabar.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub

Private Function GrabarAgendarConteo()

    '1) Valido datos y saco usuario contador de la sucursal
    If Not ValidoUsuario Then Exit Function
    If Not prmVerificador Then Exit Function
        
    If cLocal.ListIndex = -1 Then Foco cLocal: Exit Function
    If Val(tArticulo.Tag) = 0 Then Foco tArticulo: Exit Function
    If cEstado.ListIndex = -1 Then Foco cEstado: Exit Function
    If UCase(Trim(tQ.Text)) <> "AG" Then Foco tQ: Exit Function
    
    Dim mUsrContador As Long, mUsrName As String
    mUsrName = "(Sin Datos)"
    cons = "Select * from Local, Usuario " & _
               " Where LocCodigo = " & cLocal.ItemData(cLocal.ListIndex) & _
               " And LocUsuarioContador = UsuCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!LocUsuarioContador) Then
            mUsrContador = rsAux!LocUsuarioContador
            mUsrName = Trim(rsAux!UsuIdentificacion)
        End If
    End If
    rsAux.Close
    
    '2) Agendo el Conteo
    Dim mMsg As String
    mMsg = " Agendar Conteo de Mercadería a """ & mUsrName & """" & vbCrLf & _
                vbTab & "Local: " & Trim(cLocal.Text) & vbCrLf & _
                vbTab & "Artículo: " & Trim(tArticulo.Text) & vbCrLf & _
                vbTab & "Estado: " & Trim(cEstado.Text) & vbCrLf & vbCrLf & _
                "¿Confirma Agendar el Nuevo Conteo de Mercadería?"
    
    If MsgBox(mMsg, vbQuestion + vbYesNo, "Agendar Conteo") = vbNo Then Exit Function
    
    On Error GoTo errGrabar
    Screen.MousePointer = 11
    FechaDelServidor
    
    cons = "Select * from ConteoArticulo Where CArID = -1"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then rsAux.AddNew Else rsAux.Edit
    
    rsAux!CArLocal = dConteo.cLocal
    rsAux!CArArticulo = dConteo.cArticulo
    rsAux!CArEstado = dConteo.cEstado
    
    rsAux!CArFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm")
    rsAux!CArQ = 0
    rsAux!CArUsuario = mUsrContador
    rsAux!CarDifConLocal = 0
    rsAux!CArEstadoConteo = enuEConteo.Agendado
    
    rsAux.Update
    rsAux.Close
    Screen.MousePointer = 0
    
    
    NuevoIngreso
    tmUsr.Tag = 0: tmUsr.Enabled = False: tmUsr.Enabled = True
    Exit Function

errGrabar:
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Function
errValidar:
    clsGeneral.OcurrioError "Error al procesar los datos para grabar.", Err.Description
    Screen.MousePointer = 0: Exit Function
End Function

Private Function CargoUltimoStock() As Boolean
    
    On Error GoTo errCargoStock
    Screen.MousePointer = 11
    CargoUltimoStock = False
    
    Dim bOk As Boolean
    '1) Veo si hay registro de Conteo con Estado = 1 para Artículo y Local     ------------------------------------------------------
    '   Si no hay comparo con el Stock Local
    
     bOk = False
    cons = "Select Top 1 * from ConteoArticulo " & _
               " Where CArLocal = " & cLocal.ItemData(cLocal.ListIndex) & _
               " And CArArticulo = " & Val(tArticulo.Tag) & _
               " And CArEstado = " & cEstado.ItemData(cEstado.ListIndex) & _
               " And CArEstadoConteo = " & enuEConteo.OK & _
               " Order by CArFecha DESC "
               
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        dConteo.QStockAnteriorBueno = rsAux!CArQ
        dConteo.QDifAnteriorBueno = rsAux!CarDifConLocal
        bOk = True
    End If
    rsAux.Close
    
    dConteo.QStockLocalHoy = 0
    
    cons = "Select * from StockLocal " & _
           " Where StLLocal = " & cLocal.ItemData(cLocal.ListIndex) & _
           " And StLArticulo = " & Val(tArticulo.Tag) & _
           " And StLEstado = " & cEstado.ItemData(cEstado.ListIndex)

    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not bOk Then dConteo.QStockAnteriorBueno = rsAux!StLCantidad
        dConteo.QStockLocalHoy = rsAux!StLCantidad
    End If
    rsAux.Close
    
    CargoUltimoStock = True
    Screen.MousePointer = 0
    Exit Function
    
errCargoStock:
    clsGeneral.OcurrioError "Error al buscar datos del último stock.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function InicializoData()
    On Error GoTo errData
    
    If cLocal.ListIndex = -1 Then Exit Function
    If cEstado.ListIndex = -1 Then Exit Function
    If Val(tArticulo.Tag) = 0 Then Exit Function
    
    If Not ValidoUsuario Then Exit Function
    
    Screen.MousePointer = 11
    With dConteo
        .IDConteoAnterior = 0
        .IDNuevoConteo = 0
        
        .cArticulo = Val(tArticulo.Tag)
        .cLocal = cLocal.ItemData(cLocal.ListIndex)
        .cEstado = cEstado.ItemData(cEstado.ListIndex)
        
        .EstadoAnteriorConteo = 0
        .EstadoConteo = 0
        
    End With
    
    'Busco Si Hay datos de Conteo p/ Artículo ----- (Local, Articulo, Estado y Usuario) -------------------------------------------
    '1) Debe existir un conteo para el usuario
    '2) Uno Agendado (no importa para Quien)
    
    cons = "Select Top 1 * from ConteoArticulo " & _
               " Where CArLocal = " & dConteo.cLocal & _
               " And CArArticulo = " & dConteo.cArticulo & _
               " And CArEstado = " & dConteo.cEstado & _
               " And (   ( CArUsuario = " & Val(lUsr.Tag) & " And CArEstadoConteo <> " & enuEConteo.ParaVerificar & ")" & _
                        " OR  ( CArUsuario <> " & Val(lUsr.Tag) & " And CArEstadoConteo = " & enuEConteo.Agendado & ")  ) " & _
               " Order by CArFecha DESC"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        dConteo.IDConteoAnterior = rsAux!CArID
        
        'Veo el estado del conteo
        dConteo.EstadoAnteriorConteo = rsAux!CArEstadoConteo
        Select Case rsAux!CArEstadoConteo
            Case enuEConteo.Agendado
                    dConteo.IDNuevoConteo = rsAux!CArID
                    lMsg.Caption = "Conteo de Mercadería Agendado."
                    lMsg.BackColor = Colores.clVerde: lMsg.ForeColor = vbButtonText
                    
            Case enuEConteo.OK
                    lMsg.Caption = "Nuevo Conteo de Mercadería."
                    lMsg.BackColor = vbButtonFace: lMsg.ForeColor = vbButtonText
                    
            Case enuEConteo.paraRecuento:
                    dConteo.IDNuevoConteo = rsAux!CArID
                    dConteo.FechaInicialRecuento = rsAux!CArFecha
                    lMsg.Caption = "Recuento de Mercadería." & vbCrLf & _
                                          "El conteo no coincide con el stock almacenado. " & vbCrLf & vbCrLf & _
                                          "Es importante que ud. realice un recuento de la mercadería para validar la cantidad ingresada."
                    
                    lMsg.BackColor = Colores.Rojo: lMsg.ForeColor = Colores.Blanco
                    
        End Select
    
    Else
        lMsg.Caption = "Nuevo Conteo de Mercadería."
    End If
    rsAux.Close
    
    dConteo.flagData = True
    Screen.MousePointer = 0
    Exit Function

errData:
    clsGeneral.OcurrioError "Error al buscar los datos del conteo anterior.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub NuevoIngreso(Optional paraRecuento As Boolean = False)
    
    On Error Resume Next
    
    lMsg.Caption = ""
    lMsg.BackColor = vbButtonFace: lMsg.ForeColor = vbButtonText
    tQ.Text = ""
    If paraRecuento Then
        dConteo.flagData = False
        InicializoData
        Foco tQ
    Else
        cEstado.Text = ""
        Foco cEstado
    End If
        
End Sub

Private Function ValidoMovimientosStock(dDesde As Date) As Long
    
    On Error GoTo errVMS
    Screen.MousePointer = 11
    
    ValidoMovimientosStock = 0
    Dim mFecha As Date
    If dDesde = gFechaServidor Then
        mFecha = DateAdd("n", -5, dDesde)
    Else
        mFecha = dDesde
    End If
    
    cons = "Select Count(*) as Q from MovimientoStockFisico " & _
                " Where MSFLocal = " & dConteo.cLocal & _
                " And MSFArticulo = " & dConteo.cArticulo & _
                " And MSFEstado = " & dConteo.cEstado & _
                " And MSFFecha > " & Format(mFecha, "'mm/dd/yyyy hh:mm:ss'")

    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!Q) Then ValidoMovimientosStock = rsAux!Q
    End If
    rsAux.Close
    
    Screen.MousePointer = 0
    Exit Function

errVMS:
    clsGeneral.OcurrioError "Error al validar los movimientos de stock.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub LimpioMensaje()
    lMsg.Caption = ""
    lMsg.BackColor = vbButtonFace: lMsg.ForeColor = vbButtonText
End Sub

Private Sub MnuHlp_Click()
On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    cons = "Select * from Aplicacion Where AplNombre = '" & "ContadorComun" & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux!AplHelp) Then aFile = Trim(rsAux!AplHelp)
    rsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub
